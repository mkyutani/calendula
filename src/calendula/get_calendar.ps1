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

# Body lines: trimmed, blank lines dropped, and the Teams meeting block dropped
# (a line of 10+ underscores followed by "Microsoft Teams", up to the next such line or the end).
# An underscore line used for anything else (a quoted mail, say) is kept with what follows it.
function Get-BodyLines([string]$body) {
    $raw  = $body -split "`r?`n"
    $out  = [System.Collections.Generic.List[string]]::new()
    $skip = $false
    for ($j = 0; $j -lt $raw.Count; $j++) {
        $t = $raw[$j].Trim()
        if ($t -match '^_{10,}$') {
            if ($skip) { $skip = $false; continue }
            $k = $j + 1
            while ($k -lt $raw.Count -and $raw[$k].Trim() -eq "") { $k++ }
            if ($k -lt $raw.Count -and $raw[$k] -match 'Microsoft Teams') { $skip = $true; continue }
        }
        if ($skip -or $t -eq "") { continue }
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

# The Restrict filter for the appointments that overlap $day: they start before the next day
# and end after this day began (an all-day event ends at 00:00 of the day after its last day).
function Get-DayFilter([datetime]$day) {
    "[Start] < '" + $day.AddDays(1).ToString("g") + "' AND [End] > '" + $day.ToString("g") + "'"
}

# Records of the appointments in $folder (a calendar MAPIFolder) that overlap $day, sorted by start.
# Always an array, even with zero or one event.
function Get-DayEvents($folder, [datetime]$day) {
    $items = $folder.Items
    $items.IncludeRecurrences = $true
    $items.Sort("[Start]")
    $events = [System.Collections.Generic.List[object]]::new()
    foreach ($item in ($items.Restrict((Get-DayFilter $day)) | Sort-Object Start)) {
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
