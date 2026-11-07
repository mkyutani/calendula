# Test harness for get_calendar.ps1 (docs/test-design.md, 4.9). Never touches Outlook:
# it dot-sources the script (which then only defines its functions) and runs them on fake items.
# Writes everything it observed to the JSON file given as -Out; tests/test_windows.py checks it.
param([Parameter(Mandatory)][string]$Out)

. "$PSScriptRoot\get_calendar.ps1"

$r = [ordered]@{}

# W1: syntax
$errs = $null
[void][System.Management.Automation.Language.Parser]::ParseFile("$PSScriptRoot\get_calendar.ps1", [ref]$null, [ref]$errs)
$r.parse_errors = @($errs | ForEach-Object { $_.Message })

# W2: Get-BodyLines
function Join-Body { $args -join "`r`n" }
$u   = "_" * 80
$s32 = "_" * 32
$url = "https://teams.microsoft.invalid/meet/000?p=x"
$bodies = [ordered]@{
    teams_middle = "Agenda`r`n`r`n1. Budget`r`n$u`r`nMicrosoft Teams meeting`r`nJoin: https://example.invalid/x`r`n$u`r`nThanks"
    teams_end    = "Agenda`r`n$u`r`n`r`nMicrosoft Teams meeting`r`nMeeting ID: 123"
    quoted       = "Hi`r`n$u`r`nFrom: Someone`r`nOld text"
    short_line   = "_____`r`nMicrosoft Teams`r`nkept"
    spaces       = "  a  `r`n`r`n`tb`t`n`n"
    empty        = ""
    no_opening   = Join-Body "Text of the invitation" "" "Microsoft Teams 会議 " "参加する: $url" "会議 ID: 000 000" "パスコード: x" $s32 "携帯電話でダイヤルイン" "+81 0-0000-0000,,000# 日本" $s32 "開催者向け" "会議のオプション: $url"
    no_opening_en = Join-Body "Text" "Microsoft Teams meeting" "" "Join: $url" "Meeting ID: 000 000"
    inner_lines  = Join-Body "Before" $u "Microsoft Teams 会議" "参加する: $url" $s32 "携帯電話でダイヤルイン" $s32 "開催者向け" $u "After"
    need_help    = Join-Body "Before" $u "Microsoft Teams Need help?" "Join: $url" $s32 "Dial in" $s32 $u "After"
    plain_heading = Join-Body "開催形式" "Microsoft Teams 会議" "議題"
    sentence     = Join-Body "Microsoft Teams で開催します" $url "議題"
    quoted_32    = Join-Body "Hi" $s32 "From: Someone" "Old text"
    limit_url    = Join-Body "Microsoft Teams 会議" $url "議題"
    limit_end    = Join-Body "Before" $s32 "Microsoft Teams meeting" "Join: $url" $s32 "After"
    open_9       = Join-Body "Before" ("_" * 9) "Microsoft Teams meeting" "After"
    open_10      = Join-Body "Before" ("_" * 10) "Microsoft Teams meeting" "After"
    close_39     = Join-Body "Before" $u "Microsoft Teams meeting" ("_" * 39) "After"
    close_40     = Join-Body "Before" $u "Microsoft Teams meeting" ("_" * 40) "After"
}
$r.bodies = [ordered]@{}
foreach ($k in $bodies.Keys) { $r.bodies[$k] = (Get-BodyLines $bodies[$k]).ToArray() }

# W3: ConvertTo-EventRecord on an item with null text fields and a one-line body
function New-Item($start, $end, $subject, $body = "", $allDay = $false) {
    [pscustomobject]@{
        Start = [datetime]$start; End = [datetime]$end; AllDayEvent = $allDay; Subject = $subject
        Location = $null; Organizer = $null; RequiredAttendees = $null; OptionalAttendees = $null
        Categories = $null; BusyStatus = 1; Body = $body
    }
}
$r.record = ConvertTo-EventRecord (New-Item "2026-11-05 09:05:00" "2026-11-05 10:00:00" $null "one line")

# W4, W5: Get-DayEvents on fake folders, and Get-DayFilter
function New-Folder($items) {
    $coll = [pscustomobject]@{ IncludeRecurrences = $false; Filter = $null; Data = @($items) }
    $coll | Add-Member ScriptMethod Sort { param($key) }
    $coll | Add-Member ScriptMethod Restrict { param($f) $this.Filter = $f; $this.Data }
    [pscustomobject]@{ Items = $coll }
}
$day = [datetime]"2026-11-05"
$folders = [ordered]@{
    zero = New-Folder @()
    one  = New-Folder @(New-Item "2026-11-05 09:00" "2026-11-05 10:00" "solo")
    two  = New-Folder @(
        (New-Item "2026-11-05 15:00" "2026-11-05 16:00" "late"),
        (New-Item "2026-11-05 09:00" "2026-11-05 10:00" "early")
    )
}
$r.day_events = [ordered]@{}
foreach ($k in $folders.Keys) {
    $events = Get-DayEvents $folders[$k] $day
    $r.day_events[$k] = [ordered]@{
        is_array            = $events -is [array]
        count               = @($events).Count
        subjects            = @($events | ForEach-Object { $_.subject })
        include_recurrences = $folders[$k].Items.IncludeRecurrences
        filter              = $folders[$k].Items.Filter
    }
}
$filter = Get-DayFilter $day
$r.filter = $filter
$m = [regex]::Match($filter, "^\[Start\] < '(.+)' AND \[End\] > '(.+)'$")
$r.filter_times = @()
if ($m.Success) {
    $r.filter_times = @(
        ([datetime]::Parse($m.Groups[1].Value)).ToString("yyyy-MM-ddTHH:mm:ss", $inv),
        ([datetime]::Parse($m.Groups[2].Value)).ToString("yyyy-MM-ddTHH:mm:ss", $inv)
    )
}

# W6: ConvertTo-CalendarJson keeps arrays as arrays and text unchanged
$jp = New-Item "2026-11-05 09:00" "2026-11-05 10:00" '日本語の件名 "Q4" <draft>' "一行目`r`n2. 二行目"
$days = @(
    [ordered]@{ date = "2026-11-05"; events = @() },
    [ordered]@{ date = "2026-11-06"; events = @((ConvertTo-EventRecord $jp)) },
    [ordered]@{ date = "2026-11-07"; events = @((ConvertTo-EventRecord $jp), (ConvertTo-EventRecord $jp)) }
)
$r.json = ConvertTo-CalendarJson $days

[System.IO.File]::WriteAllText($Out, (ConvertTo-Json -InputObject $r -Depth 8), [System.Text.UTF8Encoding]::new($false))
