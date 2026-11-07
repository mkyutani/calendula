# Writes Outlook (classic) calendar events as JSON. Run by calendula (outlook.py) through run_hidden.vbs.
#
# Input (next to this script):  start.txt  first day, yyyyMMdd
#                                days.txt   number of days
# Output:                        events.json  {"days": [{"date": "yyyy-MM-dd", "events": [...]}]}  (UTF-8, no BOM)
#                                error.txt    "ERROR: <message>" when something fails (exit code 1)
#
# When dot-sourced (". .\get_calendar.ps1"), only the functions are defined; the tests use that.
# The contract is in docs/spec.md ("PowerShell script").

$inv = [Globalization.CultureInfo]::InvariantCulture

# Body lines: trimmed, blank lines dropped, and the Teams meeting block dropped (B-EV-3).
# The block starts at a line of 10+ underscores followed by "Microsoft Teams", or at the heading
# "Microsoft Teams 会議" or "Microsoft Teams meeting" followed by a Teams URL. It runs through the
# next line of 40+ underscores, or to the end: the shorter lines seen inside a block do not end it.
# An underscore line used for anything else (a quoted mail, say) is kept with what follows it.
function Get-BodyLines([string]$body) {
    $raw  = $body -split "`r?`n"
    $out  = [System.Collections.Generic.List[string]]::new()
    $skip = $false
    for ($j = 0; $j -lt $raw.Count; $j++) {
        $t = $raw[$j].Trim()
        if ($skip) {
            if ($t -match '^_{40,}$') { $skip = $false }
            continue
        }
        if ($t -eq "") { continue }
        $k = $j + 1
        while ($k -lt $raw.Count -and $raw[$k].Trim() -eq "") { $k++ }
        $next = if ($k -lt $raw.Count) { $raw[$k] } else { "" }
        $opening = $t -match '^_{10,}$' -and $next -match 'Microsoft Teams'
        $heading = $t -match '^Microsoft Teams (会議|meeting)$' -and $next -match 'teams\.microsoft\.'
        if ($opening -or $heading) { $skip = $true; continue }
        $out.Add($t)
    }
    return ,$out
}

# One appointment as an ordered dictionary with the fields of events.json.
function ConvertTo-EventRecord($item) {
    [ordered]@{
        start      = $item.Start.ToString("yyyy-MM-ddTHH:mm:ss", $inv)
        end        = $item.End.ToString("yyyy-MM-ddTHH:mm:ss", $inv)
        all_day    = [bool]$item.AllDayEvent
        subject    = [string]$item.Subject
        location   = [string]$item.Location
        organizer  = [string]$item.Organizer
        required   = [string]$item.RequiredAttendees
        optional   = [string]$item.OptionalAttendees
        categories = [string]$item.Categories
        busy       = [int]$item.BusyStatus
        body       = (Get-BodyLines ($item.Body + "")).ToArray()
    }
}

# The Restrict filter: a coarse one, a day wider than $day on each side. Restrict does not compare
# every appointment in local time (an all-day event in another time zone can match the day after
# its last day), so it only narrows the search. Get-DayEvents decides.
function Get-DayFilter([datetime]$day) {
    "[Start] < '" + $day.AddDays(2).ToString("g") + "' AND [End] > '" + $day.AddDays(-1).ToString("g") + "'"
}

# Records of the appointments in $folder (a calendar MAPIFolder) that overlap $day, sorted by start.
# They overlap when, in local time, they start before the next day and end after this day began
# (an all-day event ends at 00:00 of the day after its last day).
# Always an array, even with zero or one event.
function Get-DayEvents($folder, [datetime]$day) {
    $items = $folder.Items
    $items.IncludeRecurrences = $true
    $items.Sort("[Start]")
    $next = $day.AddDays(1)
    $events = [System.Collections.Generic.List[object]]::new()
    foreach ($item in ($items.Restrict((Get-DayFilter $day)) | Where-Object { $_.Start -lt $next -and $_.End -gt $day } | Sort-Object Start)) {
        $events.Add((ConvertTo-EventRecord $item))
    }
    return ,$events.ToArray()
}

# events.json text for a list of @{date; events}.
function ConvertTo-CalendarJson($days) {
    ConvertTo-Json -InputObject ([ordered]@{ days = @($days) }) -Depth 6 -Compress
}

function Invoke-Main {
    $outFile   = "$PSScriptRoot\events.json"
    $errorFile = "$PSScriptRoot\error.txt"
    $utf8      = [System.Text.UTF8Encoding]::new($false)
    Remove-Item $outFile -ErrorAction SilentlyContinue
    Remove-Item $errorFile -ErrorAction SilentlyContinue
    try {
        $start = [DateTime]::ParseExact((Get-Content "$PSScriptRoot\start.txt" -Raw).Trim(), "yyyyMMdd", $inv)
        $count = [int](Get-Content "$PSScriptRoot\days.txt" -Raw).Trim()

        $ol     = New-Object -ComObject Outlook.Application
        $folder = $ol.GetNamespace("MAPI").GetDefaultFolder(9)   # olFolderCalendar

        $days = [System.Collections.Generic.List[object]]::new()
        for ($i = 0; $i -lt $count; $i++) {
            $day = $start.AddDays($i)
            $days.Add([ordered]@{ date = $day.ToString("yyyy-MM-dd", $inv); events = (Get-DayEvents $folder $day) })
        }
        [System.IO.File]::WriteAllText($outFile, (ConvertTo-CalendarJson $days), $utf8)
    } catch {
        [System.IO.File]::WriteAllText($errorFile, "ERROR: " + $_.Exception.Message, $utf8)
        exit 1
    }
}

if ($MyInvocation.InvocationName -ne '.') { Invoke-Main }
