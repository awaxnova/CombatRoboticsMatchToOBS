Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:App = [ordered]@{
    Paths     = [ordered]@{}
    Divisions = @()
    AppState  = [ordered]@{
        SelectedDivisionId = $null
        SelectedMatchId    = $null
        LiveMatchId        = $null
    }
    ObsPreview = [ordered]@{
        Division = ''
        Match    = ''
        Status   = ''
        Vs       = ''
        Bot1     = ''
        Bot2     = ''
        Bot3     = ''
        Bot4     = ''
        Bot5     = ''
        Bot6     = ''
    }
    Controls = [ordered]@{}
}

function New-DeterministicGuid {
    param([Parameter(Mandatory = $true)][string]$Text)

    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
    $hash = [System.Security.Cryptography.MD5]::Create().ComputeHash($bytes)
    return ([System.Guid]::new($hash)).Guid
}

function Initialize-AppPaths {
    $root = Split-Path -Path $PSCommandPath -Parent

    $script:App.Paths = [ordered]@{
        Root          = $root
        DivisionsDir  = Join-Path $root 'divisions'
        DataDir       = Join-Path $root 'data'
        OutputDir     = Join-Path $root 'output'
        LogsDir       = Join-Path $root 'logs'
        MatchPlans    = Join-Path $root 'data\matchPlans.json'
        BotOverrides  = Join-Path $root 'data\botOverrides.json'
        AppState      = Join-Path $root 'data\appState.json'
        LogFile       = Join-Path $root 'logs\app.log'
    }

    foreach ($folder in @($script:App.Paths.DivisionsDir, $script:App.Paths.DataDir, $script:App.Paths.OutputDir, $script:App.Paths.LogsDir)) {
        if (-not (Test-Path -LiteralPath $folder)) {
            New-Item -ItemType Directory -Path $folder | Out-Null
        }
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )

    try {
        $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        $line = "[$stamp] [$Level] $Message"
        Add-Content -LiteralPath $script:App.Paths.LogFile -Value $line -Encoding UTF8
    }
    catch {
        # Ignore logging failures to avoid crashing the UI.
    }
}

function Set-StatusText {
    param([string]$Text)

    if ($script:App.Controls.Contains('StatusLabel') -and $script:App.Controls.StatusLabel) {
        $script:App.Controls.StatusLabel.Text = $Text
    }
}

function Show-NonFatalWarning {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [string]$Title = 'Warning'
    )

    Write-Log -Level 'WARN' -Message $Message
    Set-StatusText -Text $Message
    [System.Windows.Forms.MessageBox]::Show($Message, $Title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
}

function ConvertTo-SafeArray {
    param($Value)

    if ($null -eq $Value) { return @() }
    if ($Value -is [System.Array]) { return @($Value) }
    return @($Value)
}

function Normalize-BotName {
    param([string]$Name)

    if ($null -eq $Name) { return '' }
    return (($Name -replace '\s+', ' ').Trim())
}

function Get-BotNameKey {
    param([string]$Name)
    return (Normalize-BotName -Name $Name).ToLowerInvariant()
}

function New-BotRecord {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [string]$Team = '',
        [string]$Seed = '',
        [string]$Driver = '',
        [string]$Notes = ''
    )

    $cleanName = Normalize-BotName -Name $Name
    return [PSCustomObject]@{
        Id      = ([guid]::NewGuid()).Guid
        BotName = $cleanName
        Team    = $Team
        Seed    = $Seed
        Driver  = $Driver
        Notes   = $Notes
        RawData = @{}
    }
}

function Deduplicate-Bots {
    param([Parameter(Mandatory = $true)]$Division)

    $seen = @{}
    $unique = @()
    foreach ($bot in (ConvertTo-SafeArray -Value $Division.Bots)) {
        $name = Normalize-BotName -Name ([string]$bot.BotName)
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        $key = Get-BotNameKey -Name $name
        if ($seen.ContainsKey($key)) {
            continue
        }

        $seen[$key] = $true
        $bot.BotName = $name
        $unique += $bot
    }

    $Division.Bots = $unique
}

function Import-CsvFlexible {
    param([Parameter(Mandatory = $true)][string]$Path)

    try {
        return Import-Csv -LiteralPath $Path -Encoding UTF8
    }
    catch {
        try {
            return Import-Csv -LiteralPath $Path -Encoding Default
        }
        catch {
            throw
        }
    }
}

function Import-DivisionRows {
    param([Parameter(Mandatory = $true)][string]$Path)

    $rows = Import-CsvFlexible -Path $Path
    if (-not $rows) {
        return @()
    }

    $firstRow = $rows | Select-Object -First 1
    $propertyNames = @($firstRow.PSObject.Properties | ForEach-Object { [string]$_.Name })
    $normalizedNames = @($propertyNames | ForEach-Object { $_.Trim().TrimStart([char]0xFEFF).ToLowerInvariant() })

    if ($normalizedNames -contains 'botname') {
        return $rows
    }

    # Fallback for headerless roster files where each line is bot/team data.
    try {
        return Import-Csv -LiteralPath $Path -Header @('BotName', 'Team', 'Seed', 'Driver', 'Notes', 'Arena', 'Division') -Encoding UTF8
    }
    catch {
        return Import-Csv -LiteralPath $Path -Header @('BotName', 'Team', 'Seed', 'Driver', 'Notes', 'Arena', 'Division') -Encoding Default
    }
}

function Load-DivisionCsvFiles {
    $divisions = @()
    $csvFiles = @(Get-ChildItem -LiteralPath $script:App.Paths.DivisionsDir -Filter '*.csv' -File | Sort-Object Name)

    foreach ($file in $csvFiles) {
        try {
            $rows = Import-DivisionRows -Path $file.FullName
            if (-not $rows) { $rows = @() }

            $divisionName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
            $divisionId = New-DeterministicGuid -Text ("division|" + $divisionName.ToLowerInvariant())

            $bots = @()
            $rowIndex = 0
            foreach ($row in $rows) {
                $rowIndex++
                $props = $row.PSObject.Properties

                $botNameProp = $props | Where-Object { ([string]$_.Name).Trim().TrimStart([char]0xFEFF).ToLowerInvariant() -eq 'botname' } | Select-Object -First 1
                if (-not $botNameProp) {
                    continue
                }

                $botName = Normalize-BotName -Name ([string]$botNameProp.Value)
                if ([string]::IsNullOrWhiteSpace($botName)) {
                    continue
                }

                $rawData = @{}
                foreach ($p in $props) {
                    $normalizedName = ([string]$p.Name).Trim().TrimStart([char]0xFEFF).ToLowerInvariant()
                    if ($normalizedName -notin @('botname', 'team', 'seed', 'driver', 'notes', 'arena', 'division')) {
                        $rawData[$p.Name] = [string]$p.Value
                    }
                }

                $botId = New-DeterministicGuid -Text ("bot|" + $divisionName.ToLowerInvariant() + '|' + $botName.ToLowerInvariant() + '|' + $rowIndex)
                $bots += [PSCustomObject]@{
                    Id      = $botId
                    BotName = $botName
                    Team    = ([string]$row.Team).Trim()
                    Seed    = ([string]$row.Seed).Trim()
                    Driver  = ([string]$row.Driver).Trim()
                    Notes   = ([string]$row.Notes).Trim()
                    RawData = $rawData
                }
            }

            $division = [PSCustomObject]@{
                Id            = $divisionId
                Name          = $divisionName
                SourceCsvPath = $file.FullName
                Bots          = $bots
                Matches       = @()
            }
            Deduplicate-Bots -Division $division
            $divisions += $division

            Write-Log -Message "Loaded division '$divisionName' with $($division.Bots.Count) bots from $($file.Name)"
        }
        catch {
            $msg = "Failed to parse CSV '$($file.Name)': $($_.Exception.Message)"
            Write-Log -Level 'ERROR' -Message $msg
            if ($script:App.Controls.Form) {
                Show-NonFatalWarning -Message $msg
            }
        }
    }

    $script:App.Divisions = $divisions
}

function Load-MatchPlans {
    if (-not (Test-Path -LiteralPath $script:App.Paths.MatchPlans)) {
        return @{}
    }

    try {
        $json = Get-Content -LiteralPath $script:App.Paths.MatchPlans -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($json)) {
            return @{}
        }

        $parsed = ConvertFrom-Json -InputObject $json -ErrorAction Stop
        $map = @{}

        $divisionEntries = ConvertTo-SafeArray -Value $parsed.Divisions
        foreach ($entry in $divisionEntries) {
            if (-not $entry.DivisionName) { continue }
            $map[[string]$entry.DivisionName] = ConvertTo-SafeArray -Value $entry.Matches
        }

        Write-Log -Message 'Loaded match plans from data\matchPlans.json'
        return $map
    }
    catch {
        $msg = "Failed to load match plans JSON: $($_.Exception.Message)"
        Write-Log -Level 'ERROR' -Message $msg
        if ($script:App.Controls.Form) {
            Show-NonFatalWarning -Message $msg
        }
        return @{}
    }
}

function Save-MatchPlans {
    try {
        $payload = [ordered]@{
            Version   = 1
            SavedAt   = (Get-Date).ToString('o')
            Divisions = @()
        }

        foreach ($division in $script:App.Divisions) {
            $matches = @()
            foreach ($m in $division.Matches) {
                $matches += [ordered]@{
                    Id         = $m.Id
                    DivisionId = $m.DivisionId
                    MatchNumber = [int]$m.MatchNumber
                    Bots       = @(ConvertTo-SafeArray -Value $m.Bots)
                    Status     = [string]$m.Status
                    ArenaLabel = [string]$m.ArenaLabel
                    Notes      = [string]$m.Notes
                    CreatedAt  = [string]$m.CreatedAt
                    UpdatedAt  = [string]$m.UpdatedAt
                }
            }

            $payload.Divisions += [ordered]@{
                DivisionName = $division.Name
                DivisionId   = $division.Id
                Matches      = $matches
            }
        }

        ($payload | ConvertTo-Json -Depth 8) | Set-Content -LiteralPath $script:App.Paths.MatchPlans -Encoding UTF8
        Write-Log -Message 'Saved match plans to data\matchPlans.json'
    }
    catch {
        $msg = "Failed to save match plans JSON: $($_.Exception.Message)"
        Write-Log -Level 'ERROR' -Message $msg
        Show-NonFatalWarning -Message $msg
    }
}

function Load-BotOverrides {
    if (-not (Test-Path -LiteralPath $script:App.Paths.BotOverrides)) {
        return @{}
    }

    try {
        $json = Get-Content -LiteralPath $script:App.Paths.BotOverrides -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($json)) {
            return @{}
        }

        $parsed = ConvertFrom-Json -InputObject $json -ErrorAction Stop
        $map = @{}
        foreach ($entry in (ConvertTo-SafeArray -Value $parsed.Divisions)) {
            $divisionName = [string]$entry.DivisionName
            if ([string]::IsNullOrWhiteSpace($divisionName)) {
                continue
            }

            $names = @()
            foreach ($rawName in (ConvertTo-SafeArray -Value $entry.Bots)) {
                $name = Normalize-BotName -Name ([string]$rawName)
                if (-not [string]::IsNullOrWhiteSpace($name)) {
                    $names += $name
                }
            }
            $map[$divisionName] = $names
        }

        Write-Log -Message 'Loaded bot overrides from data\botOverrides.json'
        return $map
    }
    catch {
        $msg = "Failed to load bot overrides JSON: $($_.Exception.Message)"
        Write-Log -Level 'ERROR' -Message $msg
        if ($script:App.Controls.Form) {
            Show-NonFatalWarning -Message $msg
        }
        return @{}
    }
}

function Apply-BotOverrides {
    param([hashtable]$Overrides)

    if (-not $Overrides) { return }

    foreach ($division in $script:App.Divisions) {
        if (-not $Overrides.ContainsKey($division.Name)) {
            continue
        }

        $sourceByKey = @{}
        foreach ($bot in (ConvertTo-SafeArray -Value $division.Bots)) {
            $key = Get-BotNameKey -Name ([string]$bot.BotName)
            if ([string]::IsNullOrWhiteSpace($key)) { continue }
            if (-not $sourceByKey.ContainsKey($key)) {
                $sourceByKey[$key] = $bot
            }
        }

        $newBots = @()
        $seen = @{}
        foreach ($rawName in (ConvertTo-SafeArray -Value $Overrides[$division.Name])) {
            $name = Normalize-BotName -Name ([string]$rawName)
            $key = Get-BotNameKey -Name $name
            if ([string]::IsNullOrWhiteSpace($key) -or $seen.ContainsKey($key)) {
                continue
            }
            $seen[$key] = $true

            if ($sourceByKey.ContainsKey($key)) {
                $bot = $sourceByKey[$key]
                $bot.BotName = $name
                $newBots += $bot
            }
            else {
                $newBots += (New-BotRecord -Name $name)
            }
        }

        $division.Bots = $newBots
        Deduplicate-Bots -Division $division
    }
}

function Save-BotOverrides {
    try {
        $payload = [ordered]@{
            Version   = 1
            SavedAt   = (Get-Date).ToString('o')
            Divisions = @()
        }

        foreach ($division in $script:App.Divisions) {
            Deduplicate-Bots -Division $division
            $names = @()
            foreach ($bot in (ConvertTo-SafeArray -Value $division.Bots)) {
                $name = Normalize-BotName -Name ([string]$bot.BotName)
                if (-not [string]::IsNullOrWhiteSpace($name)) {
                    $names += $name
                }
            }

            $payload.Divisions += [ordered]@{
                DivisionName = $division.Name
                DivisionId   = $division.Id
                Bots         = $names
            }
        }

        ($payload | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $script:App.Paths.BotOverrides -Encoding UTF8
        Write-Log -Message 'Saved bot overrides to data\botOverrides.json'
    }
    catch {
        $msg = "Failed to save bot overrides JSON: $($_.Exception.Message)"
        Write-Log -Level 'ERROR' -Message $msg
        Show-NonFatalWarning -Message $msg
    }
}

function Load-AppState {
    $defaults = [ordered]@{
        SelectedDivisionId = $null
        SelectedMatchId    = $null
        LiveMatchId        = $null
    }

    if (-not (Test-Path -LiteralPath $script:App.Paths.AppState)) {
        return $defaults
    }

    try {
        $json = Get-Content -LiteralPath $script:App.Paths.AppState -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($json)) {
            return $defaults
        }

        $obj = ConvertFrom-Json -InputObject $json -ErrorAction Stop
        return [ordered]@{
            SelectedDivisionId = [string]$obj.SelectedDivisionId
            SelectedMatchId    = [string]$obj.SelectedMatchId
            LiveMatchId        = [string]$obj.LiveMatchId
        }
    }
    catch {
        $msg = "Failed to load app state JSON: $($_.Exception.Message)"
        Write-Log -Level 'ERROR' -Message $msg
        if ($script:App.Controls.Form) {
            Show-NonFatalWarning -Message $msg
        }
        return $defaults
    }
}

function Save-AppState {
    try {
        $payload = [ordered]@{
            Version           = 1
            SavedAt           = (Get-Date).ToString('o')
            SelectedDivisionId = $script:App.AppState.SelectedDivisionId
            SelectedMatchId    = $script:App.AppState.SelectedMatchId
            LiveMatchId        = $script:App.AppState.LiveMatchId
        }

        ($payload | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath $script:App.Paths.AppState -Encoding UTF8
    }
    catch {
        $msg = "Failed to save app state JSON: $($_.Exception.Message)"
        Write-Log -Level 'ERROR' -Message $msg
        Show-NonFatalWarning -Message $msg
    }
}

function Renumber-Matches {
    param([Parameter(Mandatory = $true)]$Division)

    $i = 1
    foreach ($m in $Division.Matches) {
        $m.MatchNumber = $i
        $m.UpdatedAt = (Get-Date).ToString('o')
        $i++
    }
}

function Reconcile-DivisionsAndMatches {
    param([hashtable]$SavedPlans)

    foreach ($division in $script:App.Divisions) {
        $savedMatches = @()
        if ($SavedPlans.ContainsKey($division.Name)) {
            $savedMatches = ConvertTo-SafeArray -Value $SavedPlans[$division.Name]
        }

        $newMatches = @()
        foreach ($sm in $savedMatches) {
            $status = [string]$sm.Status
            if ($status -notin @('Queued', 'Live', 'Done')) {
                $status = 'Queued'
            }

            $bots = @()
            $botSeen = @{}
            foreach ($b in (ConvertTo-SafeArray -Value $sm.Bots)) {
                $name = Normalize-BotName -Name ([string]$b)
                $key = Get-BotNameKey -Name $name
                if (-not [string]::IsNullOrWhiteSpace($name) -and -not $botSeen.ContainsKey($key)) {
                    $bots += $name
                    $botSeen[$key] = $true
                }
            }

            if ($bots.Count -lt 2) {
                continue
            }

            $createdAt = [string]$sm.CreatedAt
            if ([string]::IsNullOrWhiteSpace($createdAt)) { $createdAt = (Get-Date).ToString('o') }
            $updatedAt = [string]$sm.UpdatedAt
            if ([string]::IsNullOrWhiteSpace($updatedAt)) { $updatedAt = $createdAt }

            $matchId = [string]$sm.Id
            if ([string]::IsNullOrWhiteSpace($matchId)) {
                $matchId = ([guid]::NewGuid()).Guid
            }

            $newMatches += [PSCustomObject]@{
                Id          = $matchId
                DivisionId  = $division.Id
                MatchNumber = 0
                Bots        = $bots
                Status      = $status
                ArenaLabel  = if ([string]::IsNullOrWhiteSpace([string]$sm.ArenaLabel)) { $division.Name } else { [string]$sm.ArenaLabel }
                Notes       = [string]$sm.Notes
                CreatedAt   = $createdAt
                UpdatedAt   = $updatedAt
            }
        }

        $division.Matches = $newMatches
        Renumber-Matches -Division $division
    }
}

function Write-OutputFile {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Value
    )

    $path = Join-Path $script:App.Paths.OutputDir $Name
    Set-Content -LiteralPath $path -Value $Value -Encoding UTF8
}

function Get-ObsBotOutputMaxIndex {
    $maxIndex = 0
    $files = @(Get-ChildItem -LiteralPath $script:App.Paths.OutputDir -Filter 'current_bot_*.txt' -File -ErrorAction SilentlyContinue)
    foreach ($file in $files) {
        if ($file.BaseName -match '^current_bot_(\d+)$') {
            $idx = [int]$Matches[1]
            if ($idx -gt $maxIndex) {
                $maxIndex = $idx
            }
        }
    }
    return $maxIndex
}

function Set-ObsPreview {
    param([hashtable]$Values)

    foreach ($k in @($script:App.ObsPreview.Keys)) {
        if ($Values.ContainsKey($k)) {
            $script:App.ObsPreview[$k] = [string]$Values[$k]
        }
        else {
            $script:App.ObsPreview[$k] = ''
        }
    }

    Refresh-LivePreview
}

function Format-RumbleBotList {
    param([string[]]$Bots)

    $normalizedBots = @($Bots |
        ForEach-Object { Normalize-BotName -Name ([string]$_) } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($normalizedBots.Count -eq 0) { return '' }

    $lines = @()
    for ($i = 0; $i -lt $normalizedBots.Count; $i += 4) {
        $count = [Math]::Min(4, $normalizedBots.Count - $i)
        $lines += (($normalizedBots[$i..($i + $count - 1)]) -join ', ')
    }

    return ($lines -join [Environment]::NewLine)
}

function Write-ObsOutputFiles {
    param(
        [Parameter(Mandatory = $true)]$Division,
        [Parameter(Mandatory = $true)]$Match
    )

    try {
        $divisionText = if ([string]::IsNullOrWhiteSpace([string]$Match.ArenaLabel)) { $Division.Name } else { [string]$Match.ArenaLabel }
        $matchText = "Match $($Match.MatchNumber)"
        $bots = @(ConvertTo-SafeArray -Value $Match.Bots)
        $rumbleList = Format-RumbleBotList -Bots $bots
        $slotCount = [Math]::Max(6, [Math]::Max($bots.Count, (Get-ObsBotOutputMaxIndex)))

        $values = [ordered]@{
            Division = $divisionText
            Match    = $matchText
            Status   = 'LIVE'
            Vs       = if ($bots.Count -eq 2) { 'vs' } elseif ($bots.Count -ge 3) { 'Battle Royale' } else { '' }
            Bot1     = ''
            Bot2     = ''
            Bot3     = ''
            Bot4     = ''
            Bot5     = ''
            Bot6     = ''
        }

        for ($i = 0; $i -lt [Math]::Min(6, $bots.Count); $i++) {
            $values["Bot$($i + 1)"] = [string]$bots[$i]
        }

        Write-OutputFile -Name 'current_division.txt' -Value $values.Division
        Write-OutputFile -Name 'current_match.txt' -Value $values.Match
        Write-OutputFile -Name 'current_status.txt' -Value $values.Status
        Write-OutputFile -Name 'current_vs.txt' -Value $values.Vs
        Write-OutputFile -Name 'current_bot_rumble.txt' -Value $rumbleList
        for ($i = 1; $i -le $slotCount; $i++) {
            $botValue = if ($i -le $bots.Count) { [string]$bots[$i - 1] } else { '' }
            Write-OutputFile -Name "current_bot_$i.txt" -Value $botValue
        }

        Set-ObsPreview -Values $values
        Write-Log -Message "Wrote OBS output files for live match '$matchText' in '$divisionText'"
    }
    catch {
        $msg = "Failed writing OBS output files: $($_.Exception.Message)"
        Write-Log -Level 'ERROR' -Message $msg
        Show-NonFatalWarning -Message $msg
    }
}

function Clear-ObsOutputFiles {
    try {
        $slotCount = [Math]::Max(6, (Get-ObsBotOutputMaxIndex))

        Write-OutputFile -Name 'current_division.txt' -Value ''
        Write-OutputFile -Name 'current_match.txt' -Value ''
        Write-OutputFile -Name 'current_status.txt' -Value ''
        Write-OutputFile -Name 'current_vs.txt' -Value ''
        Write-OutputFile -Name 'current_bot_rumble.txt' -Value ''
        for ($i = 1; $i -le $slotCount; $i++) {
            Write-OutputFile -Name "current_bot_$i.txt" -Value ''
        }

        Set-ObsPreview -Values @{}
        Write-Log -Message 'Cleared OBS output files'
    }
    catch {
        $msg = "Failed clearing OBS output files: $($_.Exception.Message)"
        Write-Log -Level 'ERROR' -Message $msg
        Show-NonFatalWarning -Message $msg
    }
}

function Get-DivisionById {
    param([string]$Id)
    if ([string]::IsNullOrWhiteSpace($Id)) { return $null }
    return $script:App.Divisions | Where-Object { $_.Id -eq $Id } | Select-Object -First 1
}

function Get-CurrentDivision {
    $list = $script:App.Controls.DivisionList
    if (-not $list -or -not $list.SelectedItem) { return $null }
    return $list.SelectedItem
}

function Get-CurrentMatch {
    $list = $script:App.Controls.MatchList
    if (-not $list -or -not $list.SelectedItem) { return $null }
    return $list.SelectedItem.Match
}

function Get-SelectedMatches {
    $list = $script:App.Controls.MatchList
    $selected = @()
    if (-not $list) { return $selected }
    foreach ($item in $list.SelectedItems) {
        if ($item -and $item.Match) {
            $selected += $item.Match
        }
    }
    return @($selected)
}

function Find-MatchById {
    param([string]$MatchId)

    if ([string]::IsNullOrWhiteSpace($MatchId)) { return $null }

    foreach ($division in $script:App.Divisions) {
        foreach ($m in $division.Matches) {
            if ($m.Id -eq $MatchId) {
                return [PSCustomObject]@{
                    Division = $division
                    Match    = $m
                }
            }
        }
    }

    return $null
}

function Clear-AllLiveStatuses {
    foreach ($division in $script:App.Divisions) {
        foreach ($m in $division.Matches) {
            if ($m.Status -eq 'Live') {
                $m.Status = 'Queued'
                $m.UpdatedAt = (Get-Date).ToString('o')
            }
        }
    }
}

function Save-AllState {
    Save-MatchPlans
    Save-BotOverrides
    Save-AppState
}

function Set-LiveMatch {
    param(
        [Parameter(Mandatory = $true)]$Division,
        [Parameter(Mandatory = $true)]$Match
    )

    Clear-AllLiveStatuses

    $Match.Status = 'Live'
    $Match.UpdatedAt = (Get-Date).ToString('o')
    $script:App.AppState.LiveMatchId = $Match.Id
    $script:App.AppState.SelectedDivisionId = $Division.Id
    $script:App.AppState.SelectedMatchId = $Match.Id

    Write-ObsOutputFiles -Division $Division -Match $Match
    Save-AllState
    Write-Log -Message "Set live match '$($Match.Id)' in division '$($Division.Name)'"

    Refresh-MatchList
}

function Clear-LiveMatch {
    Clear-AllLiveStatuses
    $script:App.AppState.LiveMatchId = $null
    Clear-ObsOutputFiles
    Save-AllState
    Write-Log -Message 'Cleared live match'
    Refresh-MatchList
}

function Write-RumbleListForMatch {
    param(
        [Parameter(Mandatory = $true)]$Division,
        [Parameter(Mandatory = $true)]$Match
    )

    try {
        $bots = @(ConvertTo-SafeArray -Value $Match.Bots) |
            ForEach-Object { Normalize-BotName -Name ([string]$_) } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        $rumbleList = Format-RumbleBotList -Bots $bots

        Write-OutputFile -Name 'current_bot_rumble.txt' -Value $rumbleList
        Set-StatusText -Text "RUMBLE wrote $($bots.Count) bot(s) to current_bot_rumble.txt."
        Write-Log -Message "RUMBLE wrote selected match bots to current_bot_rumble.txt for division '$($Division.Name)' match '$($Match.Id)'"
    }
    catch {
        $msg = "Failed writing RUMBLE bot list: $($_.Exception.Message)"
        Write-Log -Level 'ERROR' -Message $msg
        Show-NonFatalWarning -Message $msg
    }
}

function Get-MatchDisplayText {
    param([Parameter(Mandatory = $true)]$Match)

    $botText = ((ConvertTo-SafeArray -Value $Match.Bots) -join ' vs ')
    return "Match $($Match.MatchNumber) | $botText | $($Match.Status)"
}

function Refresh-DivisionList {
    $list = $script:App.Controls.DivisionList
    $selectedDivisionId = $script:App.AppState.SelectedDivisionId

    $list.BeginUpdate()
    $list.Items.Clear()
    foreach ($division in $script:App.Divisions) {
        [void]$list.Items.Add($division)
    }
    $list.EndUpdate()

    $list.DisplayMember = 'Name'

    if ($selectedDivisionId) {
        for ($i = 0; $i -lt $list.Items.Count; $i++) {
            if ($list.Items[$i].Id -eq $selectedDivisionId) {
                $list.SelectedIndex = $i
                return
            }
        }
    }

    if ($list.Items.Count -gt 0) {
        $list.SelectedIndex = 0
    }
}

function Refresh-BotList {
    $division = Get-CurrentDivision
    $list = $script:App.Controls.BotList
    $filterText = ([string]$script:App.Controls.BotFilter.Text).Trim().ToLowerInvariant()

    $list.BeginUpdate()
    $list.Items.Clear()

    if ($division) {
        $bots = $division.Bots
        if (-not [string]::IsNullOrWhiteSpace($filterText)) {
            $bots = $bots | Where-Object { ([string]$_.BotName).ToLowerInvariant().Contains($filterText) }
        }

        foreach ($bot in $bots) {
            [void]$list.Items.Add($bot)
        }
    }

    $list.EndUpdate()
    $list.DisplayMember = 'BotName'
}

function Refresh-MatchList {
    $division = Get-CurrentDivision
    $list = $script:App.Controls.MatchList
    $selectedMatchId = $script:App.AppState.SelectedMatchId

    $list.BeginUpdate()
    $list.Items.Clear()

    if ($division) {
        foreach ($m in $division.Matches) {
            $item = [PSCustomObject]@{
                Match   = $m
                Display = (Get-MatchDisplayText -Match $m)
            }
            [void]$list.Items.Add($item)
        }
    }

    $list.EndUpdate()
    $list.DisplayMember = 'Display'

    if ($selectedMatchId) {
        for ($i = 0; $i -lt $list.Items.Count; $i++) {
            if ($list.Items[$i].Match.Id -eq $selectedMatchId) {
                $list.SelectedIndex = $i
                return
            }
        }
    }

    if ($list.Items.Count -gt 0) {
        $list.SelectedIndex = 0
    }
}

function Refresh-LivePreview {
    $script:App.Controls.PreviewDivision.Text = $script:App.ObsPreview.Division
    $script:App.Controls.PreviewMatch.Text = $script:App.ObsPreview.Match
    $script:App.Controls.PreviewStatus.Text = $script:App.ObsPreview.Status
    $script:App.Controls.PreviewVs.Text = $script:App.ObsPreview.Vs
    $script:App.Controls.PreviewBot1.Text = $script:App.ObsPreview.Bot1
    $script:App.Controls.PreviewBot2.Text = $script:App.ObsPreview.Bot2
    $script:App.Controls.PreviewBot3.Text = $script:App.ObsPreview.Bot3
    $script:App.Controls.PreviewBot4.Text = $script:App.ObsPreview.Bot4
    $script:App.Controls.PreviewBot5.Text = $script:App.ObsPreview.Bot5
    $script:App.Controls.PreviewBot6.Text = $script:App.ObsPreview.Bot6
}

function Show-EditMatchDialog {
    param(
        [Parameter(Mandatory = $true)]$Division,
        [Parameter(Mandatory = $true)]$Match
    )

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = 'Edit Match Bots'
    $dialog.Size = New-Object System.Drawing.Size(420, 520)
    $dialog.StartPosition = 'CenterParent'
    $dialog.FormBorderStyle = 'FixedDialog'
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = 'Select 2 or more bots:'
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(12, 12)

    $checked = New-Object System.Windows.Forms.CheckedListBox
    $checked.Location = New-Object System.Drawing.Point(12, 36)
    $checked.Size = New-Object System.Drawing.Size(380, 380)
    $checked.CheckOnClick = $true

    $currentSet = @{}
    foreach ($name in (ConvertTo-SafeArray -Value $Match.Bots)) {
        $currentSet[[string]$name] = $true
    }

    foreach ($bot in $Division.Bots) {
        $idx = $checked.Items.Add($bot.BotName)
        if ($currentSet.ContainsKey($bot.BotName)) {
            $checked.SetItemChecked($idx, $true)
        }
    }

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Save'
    $ok.Location = New-Object System.Drawing.Point(220, 430)
    $ok.Size = New-Object System.Drawing.Size(80, 30)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Cancel'
    $cancel.Location = New-Object System.Drawing.Point(312, 430)
    $cancel.Size = New-Object System.Drawing.Size(80, 30)

    $dialog.AcceptButton = $ok
    $dialog.CancelButton = $cancel

    $selectedBots = $null

    $ok.Add_Click({
        $picked = @()
        foreach ($item in $checked.CheckedItems) {
            $name = ([string]$item).Trim()
            if (-not [string]::IsNullOrWhiteSpace($name)) {
                $picked += $name
            }
        }

        if ($picked.Count -lt 2) {
            [System.Windows.Forms.MessageBox]::Show('A match must include at least 2 bots.', 'Validation', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }

        $script:__editBotsResult = $picked
        $dialog.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dialog.Close()
    })

    $cancel.Add_Click({
        $dialog.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $dialog.Close()
    })

    $dialog.Controls.Add($label)
    $dialog.Controls.Add($checked)
    $dialog.Controls.Add($ok)
    $dialog.Controls.Add($cancel)

    $script:__editBotsResult = $null
    $res = $dialog.ShowDialog($script:App.Controls.Form)
    if ($res -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedBots = $script:__editBotsResult
    }
    $script:__editBotsResult = $null

    return $selectedBots
}

function Get-SelectedBotNames {
    $bots = @()
    foreach ($item in $script:App.Controls.BotList.SelectedItems) {
        $name = ([string]$item.BotName).Trim()
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            $bots += $name
        }
    }
    return $bots
}

function Show-TextInputDialog {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$Prompt,
        [string]$InitialValue = '',
        [string]$OkText = 'OK'
    )

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = $Title
    $dialog.Size = New-Object System.Drawing.Size(760, 170)
    $dialog.StartPosition = 'CenterParent'
    $dialog.FormBorderStyle = 'FixedDialog'
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Prompt
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(12, 14)

    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Location = New-Object System.Drawing.Point(12, 40)
    $textbox.Size = New-Object System.Drawing.Size(720, 24)
    $textbox.Anchor = 'Top, Left, Right'
    $textbox.Text = $InitialValue

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = $OkText
    $ok.Location = New-Object System.Drawing.Point(572, 78)
    $ok.Size = New-Object System.Drawing.Size(76, 28)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Cancel'
    $cancel.Location = New-Object System.Drawing.Point(656, 78)
    $cancel.Size = New-Object System.Drawing.Size(76, 28)

    $dialog.AcceptButton = $ok
    $dialog.CancelButton = $cancel

    $ok.Add_Click({
        $dialog.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dialog.Close()
    })

    $cancel.Add_Click({
        $dialog.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $dialog.Close()
    })

    $dialog.Controls.Add($label)
    $dialog.Controls.Add($textbox)
    $dialog.Controls.Add($ok)
    $dialog.Controls.Add($cancel)

    [void]$dialog.ShowDialog($script:App.Controls.Form)
    if ($dialog.DialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
        return $null
    }

    return ([string]$textbox.Text).Trim()
}

function Ensure-BotsInDivision {
    param(
        [Parameter(Mandatory = $true)]$Division,
        [Parameter(Mandatory = $true)][string[]]$BotNames
    )

    $existing = @{}
    foreach ($bot in (ConvertTo-SafeArray -Value $Division.Bots)) {
        $key = Get-BotNameKey -Name ([string]$bot.BotName)
        if ([string]::IsNullOrWhiteSpace($key)) { continue }
        $existing[$key] = $true
    }

    $addedCount = 0
    foreach ($rawName in (ConvertTo-SafeArray -Value $BotNames)) {
        $name = Normalize-BotName -Name ([string]$rawName)
        $key = Get-BotNameKey -Name $name
        if ([string]::IsNullOrWhiteSpace($key) -or $existing.ContainsKey($key)) {
            continue
        }

        $Division.Bots += (New-BotRecord -Name $name)
        $existing[$key] = $true
        $addedCount++
    }

    Deduplicate-Bots -Division $Division
    return $addedCount
}

function Add-BotToDivision {
    $division = Get-CurrentDivision
    if (-not $division) {
        Show-NonFatalWarning -Message 'Select a division first.'
        return
    }

    $name = Show-TextInputDialog `
        -Title 'Add Bot' `
        -Prompt 'Enter bot name:'
    if ([string]::IsNullOrWhiteSpace($name)) {
        return
    }

    $name = Normalize-BotName -Name $name
    $key = Get-BotNameKey -Name $name
    foreach ($bot in (ConvertTo-SafeArray -Value $division.Bots)) {
        if ((Get-BotNameKey -Name ([string]$bot.BotName)) -eq $key) {
            Show-NonFatalWarning -Message "Bot '$name' already exists in this division."
            return
        }
    }

    $division.Bots += (New-BotRecord -Name $name)
    Deduplicate-Bots -Division $division
    Save-AllState
    Refresh-BotList
    Set-StatusText -Text "Added bot '$name' to '$($division.Name)'."
    Write-Log -Message "Added bot '$name' to division '$($division.Name)'"
}

function Edit-SelectedBot {
    $division = Get-CurrentDivision
    if (-not $division) {
        Show-NonFatalWarning -Message 'Select a division first.'
        return
    }

    $selected = @($script:App.Controls.BotList.SelectedItems)
    if ($selected.Count -ne 1) {
        Show-NonFatalWarning -Message 'Select exactly one bot to edit.'
        return
    }

    $bot = $selected[0]
    $oldName = Normalize-BotName -Name ([string]$bot.BotName)
    if ([string]::IsNullOrWhiteSpace($oldName)) {
        Show-NonFatalWarning -Message 'Selected bot has no valid name.'
        return
    }

    $newName = Show-TextInputDialog `
        -Title 'Edit Bot Name' `
        -Prompt 'Update bot name:' `
        -InitialValue $oldName
    if ([string]::IsNullOrWhiteSpace($newName)) {
        return
    }

    $newName = Normalize-BotName -Name $newName
    $oldKey = Get-BotNameKey -Name $oldName
    $newKey = Get-BotNameKey -Name $newName
    if ($oldKey -eq $newKey) {
        return
    }

    foreach ($other in (ConvertTo-SafeArray -Value $division.Bots)) {
        if ($other.Id -eq $bot.Id) { continue }
        if ((Get-BotNameKey -Name ([string]$other.BotName)) -eq $newKey) {
            Show-NonFatalWarning -Message "Bot '$newName' already exists in this division."
            return
        }
    }

    $bot.BotName = $newName
    $now = (Get-Date).ToString('o')
    foreach ($match in (ConvertTo-SafeArray -Value $division.Matches)) {
        $updated = $false
        $newBots = @()
        foreach ($matchBot in (ConvertTo-SafeArray -Value $match.Bots)) {
            $candidate = [string]$matchBot
            if ((Get-BotNameKey -Name $candidate) -eq $oldKey) {
                $candidate = $newName
                $updated = $true
            }
            $newBots += $candidate
        }

        if ($updated) {
            $match.Bots = $newBots
            $match.UpdatedAt = $now
            if ($match.Status -eq 'Live') {
                Write-ObsOutputFiles -Division $division -Match $match
            }
        }
    }

    Deduplicate-Bots -Division $division
    Save-AllState
    Refresh-BotList
    Refresh-MatchList
    Set-StatusText -Text "Renamed bot '$oldName' to '$newName'."
    Write-Log -Message "Renamed bot '$oldName' to '$newName' in division '$($division.Name)'"
}

function Remove-SelectedBots {
    $division = Get-CurrentDivision
    if (-not $division) {
        Show-NonFatalWarning -Message 'Select a division first.'
        return
    }

    $selected = @($script:App.Controls.BotList.SelectedItems)
    if ($selected.Count -eq 0) {
        Show-NonFatalWarning -Message 'Select one or more bots to remove.'
        return
    }

    $removeMap = @{}
    foreach ($bot in $selected) {
        $name = Normalize-BotName -Name ([string]$bot.BotName)
        $key = Get-BotNameKey -Name $name
        if ([string]::IsNullOrWhiteSpace($key) -or $removeMap.ContainsKey($key)) { continue }
        $removeMap[$key] = $true
    }

    if ($removeMap.Count -eq 0) {
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Remove $($removeMap.Count) bot(s) from '$($division.Name)'? This also removes them from matches.",
        'Confirm Remove Bots',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    $remainingBots = @()
    foreach ($bot in (ConvertTo-SafeArray -Value $division.Bots)) {
        $key = Get-BotNameKey -Name ([string]$bot.BotName)
        if (-not $removeMap.ContainsKey($key)) {
            $remainingBots += $bot
        }
    }
    $division.Bots = $remainingBots
    Deduplicate-Bots -Division $division

    $deletedMatches = @()
    $liveDeleted = $false
    $keptMatches = @()
    $now = (Get-Date).ToString('o')
    foreach ($match in (ConvertTo-SafeArray -Value $division.Matches)) {
        $newMatchBots = @()
        $updated = $false
        foreach ($name in (ConvertTo-SafeArray -Value $match.Bots)) {
            $key = Get-BotNameKey -Name ([string]$name)
            if ($removeMap.ContainsKey($key)) {
                $updated = $true
                continue
            }
            $newMatchBots += [string]$name
        }

        if ($newMatchBots.Count -lt 2) {
            $deletedMatches += $match
            if ($match.Status -eq 'Live') {
                $liveDeleted = $true
            }
            continue
        }

        if ($updated) {
            $match.Bots = $newMatchBots
            $match.UpdatedAt = $now
            if ($match.Status -eq 'Live') {
                Write-ObsOutputFiles -Division $division -Match $match
            }
        }
        $keptMatches += $match
    }
    $division.Matches = $keptMatches

    if ($liveDeleted) {
        $script:App.AppState.LiveMatchId = $null
        Clear-ObsOutputFiles
    }
    elseif (-not (Find-MatchById -MatchId $script:App.AppState.LiveMatchId)) {
        $script:App.AppState.LiveMatchId = $null
    }

    Renumber-Matches -Division $division
    $script:App.AppState.SelectedMatchId = $null
    Save-AllState
    Refresh-BotList
    Refresh-MatchList

    $status = "Removed $($removeMap.Count) bot(s). Deleted $($deletedMatches.Count) invalid match(es)."
    Set-StatusText -Text $status
    Write-Log -Message "$status Division '$($division.Name)'."
}

function Convert-ChallongeSvgToMatches {
    param([Parameter(Mandatory = $true)][string]$SvgText)

    $xml = New-Object System.Xml.XmlDocument
    $xml.XmlResolver = $null
    $xml.LoadXml($SvgText)

    $matchNodes = $xml.SelectNodes("//*[local-name()='g' and contains(concat(' ', normalize-space(@class), ' '), ' match ')]")
    $parsed = @()
    $ordinal = 0

    foreach ($node in $matchNodes) {
        $ordinal++

        $identifier = 0
        $identifierAttr = $node.Attributes['data-identifier']
        if ($identifierAttr -and -not [string]::IsNullOrWhiteSpace([string]$identifierAttr.Value)) {
            [void][int]::TryParse([string]$identifierAttr.Value, [ref]$identifier)
        }

        $x = 0.0
        $y = 0.0
        $transformAttr = $node.Attributes['transform']
        if ($transformAttr -and ([string]$transformAttr.Value -match 'translate\(\s*(-?\d+(\.\d+)?)\s+(-?\d+(\.\d+)?)\s*\)')) {
            [void][double]::TryParse([string]$Matches[1], [ref]$x)
            [void][double]::TryParse([string]$Matches[3], [ref]$y)
        }

        $playerNodes = $node.SelectNodes(".//*[local-name()='svg' and contains(concat(' ', normalize-space(@class), ' '), ' match--player ')]")
        $bots = @()
        foreach ($playerNode in $playerNodes) {
            $name = ''
            $titleNode = $playerNode.SelectSingleNode("./*[local-name()='title']")
            if ($titleNode) {
                $name = ([string]$titleNode.InnerText).Trim()
            }

            if ([string]::IsNullOrWhiteSpace($name)) {
                $nameNode = $playerNode.SelectSingleNode(".//*[local-name()='text' and contains(concat(' ', normalize-space(@class), ' '), ' match--player-name ')]")
                if ($nameNode) {
                    $name = ([string]$nameNode.InnerText).Trim()
                }
            }

            $ignore = $false
            if ([string]::IsNullOrWhiteSpace($name)) { $ignore = $true }
            if ($name -match '^(TBD|BYE)$') { $ignore = $true }
            if ($name -match '^(Winner|Loser)\s+of\b') { $ignore = $true }

            if (-not $ignore -and -not ($bots -contains $name)) {
                $bots += $name
            }
        }

        if ($bots.Count -ge 2) {
            $parsed += [PSCustomObject]@{
                Identifier = $identifier
                X          = $x
                Y          = $y
                Ordinal    = $ordinal
                Bots       = @($bots)
            }
        }
    }

    $ordered = $parsed | Sort-Object `
        @{ Expression = { if ([int]$_.Identifier -gt 0) { 0 } else { 1 } } }, `
        @{ Expression = { [int]$_.Identifier } }, `
        @{ Expression = { [double]$_.X } }, `
        @{ Expression = { [double]$_.Y } }, `
        @{ Expression = { [int]$_.Ordinal } }
    return @($ordered)
}

function Create-MatchFromSelection {
    $division = Get-CurrentDivision
    if (-not $division) {
        Show-NonFatalWarning -Message 'Select a division first.'
        return
    }

    $botNames = Get-SelectedBotNames
    if ($botNames.Count -lt 2) {
        Show-NonFatalWarning -Message 'Select at least 2 bots to create a match.'
        return
    }

    $duplicates = @()
    foreach ($name in $botNames) {
        if ($division.Matches | Where-Object { $_.Bots -contains $name }) {
            $duplicates += $name
        }
    }
    if ($duplicates.Count -gt 0) {
        $warn = "Selected bot(s) already appear in other matches: $($duplicates -join ', ')."
        Set-StatusText -Text $warn
        Write-Log -Level 'WARN' -Message $warn
    }

    $now = (Get-Date).ToString('o')
    $match = [PSCustomObject]@{
        Id          = ([guid]::NewGuid()).Guid
        DivisionId  = $division.Id
        MatchNumber = 0
        Bots        = @($botNames)
        Status      = 'Queued'
        ArenaLabel  = $division.Name
        Notes       = ''
        CreatedAt   = $now
        UpdatedAt   = $now
    }

    $division.Matches += $match
    Renumber-Matches -Division $division

    $script:App.AppState.SelectedMatchId = $match.Id
    Save-AllState
    Write-Log -Message "Created match '$($match.Id)' in division '$($division.Name)'"

    Refresh-MatchList
}

function Import-ChallongeSvgMatches {
    $division = Get-CurrentDivision
    if (-not $division) {
        Show-NonFatalWarning -Message 'Select a division first.'
        return
    }

    $source = Show-TextInputDialog `
        -Title 'Import Challonge SVG' `
        -Prompt 'Paste a Challonge printer-friendly SVG URL or a local .svg path:' `
        -InitialValue 'https://challonge.com/your_tournament.svg' `
        -OkText 'Import'

    if ([string]::IsNullOrWhiteSpace($source)) {
        return
    }

    $svgText = $null
    try {
        if ($source -match '^(?i)https?://') {
            $response = Invoke-WebRequest -Uri $source -UseBasicParsing -TimeoutSec 30
            $svgText = [string]$response.Content
        }
        else {
            if (-not (Test-Path -LiteralPath $source -PathType Leaf)) {
                Show-NonFatalWarning -Message "SVG source not found: $source"
                return
            }
            $svgText = Get-Content -LiteralPath $source -Raw -Encoding UTF8
        }
    }
    catch {
        $msg = "Failed to load Challonge SVG: $($_.Exception.Message)"
        Write-Log -Level 'ERROR' -Message $msg
        Show-NonFatalWarning -Message $msg
        return
    }

    $imported = @()
    try {
        $imported = Convert-ChallongeSvgToMatches -SvgText $svgText
    }
    catch {
        $msg = "Failed to parse Challonge SVG: $($_.Exception.Message)"
        Write-Log -Level 'ERROR' -Message $msg
        Show-NonFatalWarning -Message $msg
        return
    }

    if ($imported.Count -eq 0) {
        Show-NonFatalWarning -Message 'No valid 2+ bot matches were found in the SVG.'
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Import $($imported.Count) match(es) into division '$($division.Name)'? This appends to existing matches.",
        'Confirm SVG Import',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    $allImportedNames = @()
    foreach ($entry in $imported) {
        foreach ($name in (ConvertTo-SafeArray -Value $entry.Bots)) {
            $clean = Normalize-BotName -Name ([string]$name)
            if (-not [string]::IsNullOrWhiteSpace($clean)) {
                $allImportedNames += $clean
            }
        }
    }
    $addedBots = Ensure-BotsInDivision -Division $division -BotNames $allImportedNames

    $now = (Get-Date).ToString('o')
    $addedCount = 0
    foreach ($entry in $imported) {
        $botNames = @()
        $matchSeen = @{}
        foreach ($name in (ConvertTo-SafeArray -Value $entry.Bots)) {
            $trimmed = Normalize-BotName -Name ([string]$name)
            $key = Get-BotNameKey -Name $trimmed
            if (-not [string]::IsNullOrWhiteSpace($trimmed) -and -not $matchSeen.ContainsKey($key)) {
                $botNames += $trimmed
                $matchSeen[$key] = $true
            }
        }

        if ($botNames.Count -lt 2) {
            continue
        }

        $division.Matches += [PSCustomObject]@{
            Id          = ([guid]::NewGuid()).Guid
            DivisionId  = $division.Id
            MatchNumber = 0
            Bots        = @($botNames)
            Status      = 'Queued'
            ArenaLabel  = $division.Name
            Notes       = "Imported from Challonge SVG: $source"
            CreatedAt   = $now
            UpdatedAt   = $now
        }
        $addedCount++
    }

    if ($addedCount -eq 0) {
        Show-NonFatalWarning -Message 'No matches were added from the SVG.'
        return
    }

    Renumber-Matches -Division $division
    $script:App.AppState.SelectedMatchId = $null
    Save-AllState
    Refresh-MatchList

    $status = "Imported $addedCount match(es) and added $addedBots new bot(s) from Challonge SVG into '$($division.Name)'."
    Set-StatusText -Text $status
    Write-Log -Message $status
}

function Move-SelectedMatch {
    param([Parameter(Mandatory = $true)][int]$Delta)

    $division = Get-CurrentDivision
    $matchList = $script:App.Controls.MatchList
    if (-not $division -or $matchList.SelectedIndex -lt 0) {
        return
    }

    $from = $matchList.SelectedIndex
    $to = $from + $Delta
    if ($to -lt 0 -or $to -ge $division.Matches.Count) {
        return
    }

    $temp = $division.Matches[$from]
    $division.Matches[$from] = $division.Matches[$to]
    $division.Matches[$to] = $temp

    Renumber-Matches -Division $division
    $script:App.AppState.SelectedMatchId = $temp.Id
    Save-AllState
    Write-Log -Message "Reordered matches in division '$($division.Name)'"

    Refresh-MatchList
    if ($to -ge 0 -and $to -lt $matchList.Items.Count) {
        $matchList.SelectedIndex = $to
    }
}

function Delete-SelectedMatch {
    $division = Get-CurrentDivision
    $matchList = $script:App.Controls.MatchList
    $selectedMatches = @(Get-SelectedMatches)

    if (-not $division -or $selectedMatches.Count -eq 0) {
        Show-NonFatalWarning -Message 'Select one or more matches to delete.'
        return
    }

    $count = $selectedMatches.Count
    $label = if ($count -eq 1) { 'this match' } else { "these $count matches" }
    $confirmText = ('Delete {0}?' -f $label)
    $confirm = [System.Windows.Forms.MessageBox]::Show($confirmText, 'Confirm Delete', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    $selectedIds = @{}
    foreach ($m in $selectedMatches) {
        $selectedIds[$m.Id] = $true
    }

    $remaining = @()
    foreach ($m in $division.Matches) {
        if (-not $selectedIds.ContainsKey($m.Id)) {
            $remaining += $m
        }
    }
    $division.Matches = $remaining

    if ($selectedIds.ContainsKey([string]$script:App.AppState.LiveMatchId)) {
        $script:App.AppState.LiveMatchId = $null
        Clear-ObsOutputFiles
    }

    Renumber-Matches -Division $division
    $script:App.AppState.SelectedMatchId = $null

    Save-AllState
    Write-Log -Message "Deleted $count match(es) from division '$($division.Name)'"

    Refresh-MatchList
    if ($matchList.Items.Count -gt 0) {
        $matchList.SelectedIndex = [Math]::Min($matchList.Items.Count - 1, 0)
    }
}

function Edit-SelectedMatch {
    $division = Get-CurrentDivision
    $match = Get-CurrentMatch

    if (-not $division -or -not $match) {
        Show-NonFatalWarning -Message 'Select a match to edit.'
        return
    }

    $newBots = Show-EditMatchDialog -Division $division -Match $match
    if (-not $newBots) {
        return
    }

    $match.Bots = @($newBots)
    $match.UpdatedAt = (Get-Date).ToString('o')

    if ($match.Status -eq 'Live') {
        Write-ObsOutputFiles -Division $division -Match $match
    }

    Save-AllState
    Write-Log -Message "Edited match '$($match.Id)' in division '$($division.Name)'"
    Refresh-MatchList
}

function Set-SelectedMatchStatus {
    param([Parameter(Mandatory = $true)][ValidateSet('Queued', 'Done')][string]$Status)

    $division = Get-CurrentDivision
    $match = Get-CurrentMatch

    if (-not $division -or -not $match) {
        Show-NonFatalWarning -Message 'Select a match first.'
        return
    }

    $wasLive = ($match.Status -eq 'Live')
    $match.Status = $Status
    $match.UpdatedAt = (Get-Date).ToString('o')

    if ($wasLive -and $Status -eq 'Queued') {
        $script:App.AppState.LiveMatchId = $null
        Clear-ObsOutputFiles
    }
    elseif ($wasLive -and $Status -eq 'Done') {
        # Match completed: clear live state and blank OBS output files.
        $script:App.AppState.LiveMatchId = $null
        Clear-ObsOutputFiles
    }

    Save-AllState
    Write-Log -Message "Set match '$($match.Id)' status to $Status"
    Refresh-MatchList
}

function Reload-Divisions {
    $priorState = [ordered]@{
        SelectedDivisionId = $script:App.AppState.SelectedDivisionId
        SelectedMatchId    = $script:App.AppState.SelectedMatchId
        LiveMatchId        = $script:App.AppState.LiveMatchId
    }

    Load-DivisionCsvFiles
    $botOverrides = Load-BotOverrides
    Apply-BotOverrides -Overrides $botOverrides
    $plans = Load-MatchPlans
    Reconcile-DivisionsAndMatches -SavedPlans $plans

    $script:App.AppState.SelectedDivisionId = $priorState.SelectedDivisionId
    $script:App.AppState.SelectedMatchId = $priorState.SelectedMatchId
    $script:App.AppState.LiveMatchId = $priorState.LiveMatchId

    $liveRef = Find-MatchById -MatchId $script:App.AppState.LiveMatchId
    Clear-AllLiveStatuses
    if ($liveRef) {
        $liveRef.Match.Status = 'Live'
        Write-ObsOutputFiles -Division $liveRef.Division -Match $liveRef.Match
    }
    else {
        $script:App.AppState.LiveMatchId = $null
        Clear-ObsOutputFiles
    }

    Refresh-DivisionList
    Refresh-BotList
    Refresh-MatchList
    Save-AllState
    Set-StatusText -Text 'Divisions reloaded.'
    Write-Log -Message 'Reloaded divisions from CSV files'
}

function Build-MainForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Combat Match Manager'
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size(1280, 800)
    $form.MinimumSize = New-Object System.Drawing.Size(1100, 680)
    $form.KeyPreview = $true

    $root = New-Object System.Windows.Forms.TableLayoutPanel
    $root.Dock = 'Fill'
    $root.RowCount = 2
    $root.ColumnCount = 1
    [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 72)))
    [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 28)))

    $main = New-Object System.Windows.Forms.TableLayoutPanel
    $main.Dock = 'Fill'
    $main.RowCount = 1
    $main.ColumnCount = 3
    [void]$main.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 24)))
    [void]$main.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
    [void]$main.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 43)))

    $leftPanel = New-Object System.Windows.Forms.Panel
    $leftPanel.Dock = 'Fill'

    $btnReload = New-Object System.Windows.Forms.Button
    $btnReload.Text = 'Reload Divisions'
    $btnReload.Dock = 'Top'
    $btnReload.Height = 32

    $divList = New-Object System.Windows.Forms.ListBox
    $divList.Dock = 'Fill'
    $divList.IntegralHeight = $false

    $leftPanel.Controls.Add($divList)
    $leftPanel.Controls.Add($btnReload)

    $midPanel = New-Object System.Windows.Forms.Panel
    $midPanel.Dock = 'Fill'

    $botHeader = New-Object System.Windows.Forms.Label
    $botHeader.Text = 'Bots (Multi-select)'
    $botHeader.Dock = 'Top'
    $botHeader.Height = 18

    $botFilter = New-Object System.Windows.Forms.TextBox
    $botFilter.Dock = 'Top'

    $botFilterLabel = New-Object System.Windows.Forms.Label
    $botFilterLabel.Text = 'Search:'
    $botFilterLabel.Dock = 'Top'
    $botFilterLabel.Height = 18

    $btnCreate = New-Object System.Windows.Forms.Button
    $btnCreate.Text = 'Create Match From Selected Bots'
    $btnCreate.Dock = 'Bottom'
    $btnCreate.Height = 32

    $botActions = New-Object System.Windows.Forms.FlowLayoutPanel
    $botActions.Dock = 'Bottom'
    $botActions.Height = 34
    $botActions.WrapContents = $false

    $btnAddBot = New-Object System.Windows.Forms.Button
    $btnAddBot.Text = 'Add Bot'
    $btnAddBot.Width = 90

    $btnEditBot = New-Object System.Windows.Forms.Button
    $btnEditBot.Text = 'Edit Bot'
    $btnEditBot.Width = 90

    $btnRemoveBot = New-Object System.Windows.Forms.Button
    $btnRemoveBot.Text = 'Remove Bot(s)'
    $btnRemoveBot.Width = 110

    foreach ($btn in @($btnAddBot, $btnEditBot, $btnRemoveBot)) {
        [void]$botActions.Controls.Add($btn)
    }

    $botList = New-Object System.Windows.Forms.ListBox
    $botList.Dock = 'Fill'
    $botList.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
    $botList.IntegralHeight = $false

    $midPanel.Controls.Add($botList)
    $midPanel.Controls.Add($btnCreate)
    $midPanel.Controls.Add($botActions)
    $midPanel.Controls.Add($botFilter)
    $midPanel.Controls.Add($botFilterLabel)
    $midPanel.Controls.Add($botHeader)

    $rightPanel = New-Object System.Windows.Forms.Panel
    $rightPanel.Dock = 'Fill'

    $matchList = New-Object System.Windows.Forms.ListBox
    $matchList.Dock = 'Fill'
    $matchList.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
    $matchList.IntegralHeight = $false

    $actions = New-Object System.Windows.Forms.FlowLayoutPanel
    $actions.Dock = 'Bottom'
    $actions.Height = 108
    $actions.WrapContents = $true

    $btnSetLive = New-Object System.Windows.Forms.Button
    $btnSetLive.Text = 'Set Live'
    $btnSetLive.Width = 95

    $btnClearLive = New-Object System.Windows.Forms.Button
    $btnClearLive.Text = 'Clear Live'
    $btnClearLive.Width = 95

    $btnUp = New-Object System.Windows.Forms.Button
    $btnUp.Text = 'Move Up'
    $btnUp.Width = 95

    $btnDown = New-Object System.Windows.Forms.Button
    $btnDown.Text = 'Move Down'
    $btnDown.Width = 95

    $btnEdit = New-Object System.Windows.Forms.Button
    $btnEdit.Text = 'Edit Match'
    $btnEdit.Width = 95

    $btnDelete = New-Object System.Windows.Forms.Button
    $btnDelete.Text = 'Delete Match'
    $btnDelete.Width = 95

    $btnDone = New-Object System.Windows.Forms.Button
    $btnDone.Text = 'Mark Done'
    $btnDone.Width = 95

    $btnQueued = New-Object System.Windows.Forms.Button
    $btnQueued.Text = 'Mark Queued'
    $btnQueued.Width = 95

    $btnRumble = New-Object System.Windows.Forms.Button
    $btnRumble.Text = 'RUMBLE'
    $btnRumble.Width = 95

    $btnImportSvg = New-Object System.Windows.Forms.Button
    $btnImportSvg.Text = 'Import Challonge SVG'
    $btnImportSvg.Width = 145

    foreach ($btn in @($btnSetLive, $btnClearLive, $btnUp, $btnDown, $btnEdit, $btnDelete, $btnDone, $btnQueued, $btnRumble, $btnImportSvg)) {
        [void]$actions.Controls.Add($btn)
    }

    $rightPanel.Controls.Add($matchList)
    $rightPanel.Controls.Add($actions)

    $previewPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $previewPanel.Dock = 'Fill'
    $previewPanel.ColumnCount = 4
    $previewPanel.RowCount = 5

    for ($c = 0; $c -lt 4; $c++) {
        [void]$previewPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
    }
    for ($r = 0; $r -lt 5; $r++) {
        [void]$previewPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 20)))
    }

    function New-PreviewTextBox {
        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Dock = 'Fill'
        $tb.ReadOnly = $true
        $tb.TabStop = $false
        return $tb
    }

    function Add-PreviewField {
        param(
            [System.Windows.Forms.TableLayoutPanel]$Table,
            [string]$Label,
            [System.Windows.Forms.Control]$Control,
            [int]$Col,
            [int]$Row
        )

        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $Label
        $lbl.Dock = 'Fill'
        $lbl.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft

        $cell = New-Object System.Windows.Forms.Panel
        $cell.Dock = 'Fill'
        $cell.Padding = New-Object System.Windows.Forms.Padding(4)

        $inner = New-Object System.Windows.Forms.TableLayoutPanel
        $inner.Dock = 'Fill'
        $inner.RowCount = 2
        $inner.ColumnCount = 1
        [void]$inner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 18)))
        [void]$inner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
        $inner.Controls.Add($lbl, 0, 0)
        $inner.Controls.Add($Control, 0, 1)

        $cell.Controls.Add($inner)
        $Table.Controls.Add($cell, $Col, $Row)
    }

    $tbDivision = New-PreviewTextBox
    $tbMatch = New-PreviewTextBox
    $tbStatus = New-PreviewTextBox
    $tbVs = New-PreviewTextBox
    $tbBot1 = New-PreviewTextBox
    $tbBot2 = New-PreviewTextBox
    $tbBot3 = New-PreviewTextBox
    $tbBot4 = New-PreviewTextBox
    $tbBot5 = New-PreviewTextBox
    $tbBot6 = New-PreviewTextBox

    Add-PreviewField -Table $previewPanel -Label 'Division' -Control $tbDivision -Col 0 -Row 0
    Add-PreviewField -Table $previewPanel -Label 'Match' -Control $tbMatch -Col 1 -Row 0
    Add-PreviewField -Table $previewPanel -Label 'Status' -Control $tbStatus -Col 2 -Row 0
    Add-PreviewField -Table $previewPanel -Label 'VS' -Control $tbVs -Col 3 -Row 0
    Add-PreviewField -Table $previewPanel -Label 'Bot 1' -Control $tbBot1 -Col 0 -Row 1
    Add-PreviewField -Table $previewPanel -Label 'Bot 2' -Control $tbBot2 -Col 1 -Row 1
    Add-PreviewField -Table $previewPanel -Label 'Bot 3' -Control $tbBot3 -Col 2 -Row 1
    Add-PreviewField -Table $previewPanel -Label 'Bot 4' -Control $tbBot4 -Col 3 -Row 1
    Add-PreviewField -Table $previewPanel -Label 'Bot 5' -Control $tbBot5 -Col 0 -Row 2
    Add-PreviewField -Table $previewPanel -Label 'Bot 6' -Control $tbBot6 -Col 1 -Row 2

    $statusStrip = New-Object System.Windows.Forms.StatusStrip
    $statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $statusLabel.Text = 'Ready'
    [void]$statusStrip.Items.Add($statusLabel)

    $main.Controls.Add($leftPanel, 0, 0)
    $main.Controls.Add($midPanel, 1, 0)
    $main.Controls.Add($rightPanel, 2, 0)

    $root.Controls.Add($main, 0, 0)
    $root.Controls.Add($previewPanel, 0, 1)

    $form.Controls.Add($root)
    $form.Controls.Add($statusStrip)

    $script:App.Controls = [ordered]@{
        Form            = $form
        DivisionList    = $divList
        BotList         = $botList
        MatchList       = $matchList
        BotFilter       = $botFilter
        StatusLabel     = $statusLabel
        PreviewDivision = $tbDivision
        PreviewMatch    = $tbMatch
        PreviewStatus   = $tbStatus
        PreviewVs       = $tbVs
        PreviewBot1     = $tbBot1
        PreviewBot2     = $tbBot2
        PreviewBot3     = $tbBot3
        PreviewBot4     = $tbBot4
        PreviewBot5     = $tbBot5
        PreviewBot6     = $tbBot6
    }

    $divList.Add_SelectedIndexChanged({
        $division = Get-CurrentDivision
        if ($division) {
            $script:App.AppState.SelectedDivisionId = $division.Id
            $script:App.AppState.SelectedMatchId = $null
        }
        Refresh-BotList
        Refresh-MatchList
        Save-AppState
    })

    $matchList.Add_SelectedIndexChanged({
        $match = Get-CurrentMatch
        if ($match) {
            $script:App.AppState.SelectedMatchId = $match.Id
            Save-AppState
        }
    })

    $botFilter.Add_TextChanged({
        Refresh-BotList
    })

    $btnReload.Add_Click({ Reload-Divisions })
    $btnCreate.Add_Click({ Create-MatchFromSelection })
    $btnAddBot.Add_Click({ Add-BotToDivision })
    $btnEditBot.Add_Click({ Edit-SelectedBot })
    $btnRemoveBot.Add_Click({ Remove-SelectedBots })
    $btnUp.Add_Click({ Move-SelectedMatch -Delta -1 })
    $btnDown.Add_Click({ Move-SelectedMatch -Delta 1 })
    $btnDelete.Add_Click({ Delete-SelectedMatch })
    $btnEdit.Add_Click({ Edit-SelectedMatch })
    $btnDone.Add_Click({ Set-SelectedMatchStatus -Status 'Done' })
    $btnQueued.Add_Click({ Set-SelectedMatchStatus -Status 'Queued' })
    $btnRumble.Add_Click({
        $division = Get-CurrentDivision
        $match = Get-CurrentMatch
        if (-not $match) {
            Show-NonFatalWarning -Message 'Select a match to RUMBLE.'
            return
        }
        Write-RumbleListForMatch -Division $division -Match $match
    })
    $btnImportSvg.Add_Click({ Import-ChallongeSvgMatches })

    $btnSetLive.Add_Click({
        $division = Get-CurrentDivision
        $match = Get-CurrentMatch
        if (-not $match) {
            Show-NonFatalWarning -Message 'Select a match to set live.'
            return
        }
        Set-LiveMatch -Division $division -Match $match
    })

    $btnClearLive.Add_Click({ Clear-LiveMatch })

    $form.Add_KeyDown({
        param($sender, $e)

        $activeHost = if ($sender -is [System.Windows.Forms.Form]) { $sender } else { $script:App.Controls.Form }
        $active = $activeHost.ActiveControl
        if ($active -is [System.Windows.Forms.TextBox] -and -not $active.ReadOnly) {
            return
        }

        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::R) {
            Reload-Divisions
            $e.SuppressKeyPress = $true
            return
        }

        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::N) {
            Create-MatchFromSelection
            $e.SuppressKeyPress = $true
            return
        }

        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::B) {
            Add-BotToDivision
            $e.SuppressKeyPress = $true
            return
        }

        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::E) {
            Edit-SelectedBot
            $e.SuppressKeyPress = $true
            return
        }

        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::L) {
            Clear-LiveMatch
            $e.SuppressKeyPress = $true
            return
        }

        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::I) {
            Import-ChallongeSvgMatches
            $e.SuppressKeyPress = $true
            return
        }

        if ($script:App.Controls.MatchList.Focused) {
            if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                $division = Get-CurrentDivision
                $match = Get-CurrentMatch
                if ($match) {
                    Set-LiveMatch -Division $division -Match $match
                }
                else {
                    Show-NonFatalWarning -Message 'Select a match to set live.'
                }
                $e.SuppressKeyPress = $true
                return
            }

            if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::Up) {
                Move-SelectedMatch -Delta -1
                $e.SuppressKeyPress = $true
                return
            }

            if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::Down) {
                Move-SelectedMatch -Delta 1
                $e.SuppressKeyPress = $true
                return
            }

            if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Delete) {
                Delete-SelectedMatch
                $e.SuppressKeyPress = $true
                return
            }
        }

        if ($script:App.Controls.BotList.Focused) {
            if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Delete) {
                Remove-SelectedBots
                $e.SuppressKeyPress = $true
                return
            }
        }
    })
}

function Initialize-Data {
    Load-DivisionCsvFiles
    $botOverrides = Load-BotOverrides
    Apply-BotOverrides -Overrides $botOverrides
    $savedPlans = Load-MatchPlans
    Reconcile-DivisionsAndMatches -SavedPlans $savedPlans
    $script:App.AppState = Load-AppState

    Clear-AllLiveStatuses
    $liveRef = Find-MatchById -MatchId $script:App.AppState.LiveMatchId
    if ($liveRef) {
        $liveRef.Match.Status = 'Live'
        Write-ObsOutputFiles -Division $liveRef.Division -Match $liveRef.Match
    }
    else {
        $script:App.AppState.LiveMatchId = $null
        Clear-ObsOutputFiles
    }
}

function Start-App {
    Initialize-AppPaths
    Write-Log -Message 'Application startup'

    Build-MainForm
    Initialize-Data

    Refresh-DivisionList
    Refresh-BotList
    Refresh-MatchList
    Refresh-LivePreview

    $script:App.Controls.Form.Add_Shown({
        Set-StatusText -Text 'Combat Match Manager ready.'
    })

    [void]$script:App.Controls.Form.ShowDialog()

    Save-AllState
    Write-Log -Message 'Application shutdown'
}

Start-App
