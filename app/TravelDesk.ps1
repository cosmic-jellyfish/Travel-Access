# Travel Desk. Day-to-day window for the travel schedule.
# The scheduled sync is app\Sync.ps1. It does not open this window.

if ($PSVersionTable.PSEdition -ne "Desktop" -or [Threading.Thread]::CurrentThread.GetApartmentState() -ne "STA") {
    $windowsPowerShell = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
    Start-Process -FilePath $windowsPowerShell -ArgumentList "-NoProfile -STA -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    exit 0
}

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$script:InstallRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$script:BasePath = Join-Path $script:InstallRoot "data"
$script:ConfigPath = Join-Path $script:BasePath "Config.json"
$script:SchedulePath = Join-Path $script:BasePath "TravelSchedule.csv"
$script:CommonPath = Join-Path $PSScriptRoot "Common.ps1"
if (!(Test-Path -LiteralPath $script:CommonPath)) {
    Write-Host "Common.ps1 is missing from $PSScriptRoot"
    exit 1
}

. $script:CommonPath
$script:Version = $script:TravelAccessVersion
$deskPaths = Initialize-TravelAccessData -AppDirectory $PSScriptRoot
$script:InstallRoot = $deskPaths.InstallRoot
$script:BasePath = $deskPaths.DataPath
$script:ConfigPath = Join-Path $script:BasePath "Config.json"
$script:SchedulePath = Join-Path $script:BasePath "TravelSchedule.csv"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

try {
    [System.Windows.Forms.Application]::EnableVisualStyles()
    [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
}
catch {
}

try {
    Add-Type -TypeDefinition @"
using System.Runtime.InteropServices;
public static class TravelDeskDpi {
    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();
}
"@ -ErrorAction Stop
    [TravelDeskDpi]::SetProcessDPIAware() | Out-Null
}
catch {
}

$script:ColorHeader = [System.Drawing.Color]::FromArgb(15, 56, 92)
$script:ColorAccent = [System.Drawing.Color]::FromArgb(15, 108, 189)
$script:ColorDanger = [System.Drawing.Color]::FromArgb(156, 52, 52)
$script:ColorOk = [System.Drawing.Color]::FromArgb(20, 110, 48)
$script:ColorMuted = [System.Drawing.Color]::FromArgb(90, 99, 108)
$script:ColorCanvas = [System.Drawing.Color]::FromArgb(244, 246, 248)
$script:ColorLine = [System.Drawing.Color]::FromArgb(214, 220, 226)
$script:ColorActive = [System.Drawing.Color]::FromArgb(226, 244, 230)
$script:ColorInvalid = [System.Drawing.Color]::FromArgb(252, 235, 233)
$script:FontUi = New-Object System.Drawing.Font("Segoe UI", 9)
$script:FontBold = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$script:FontTitle = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$script:FontMono = New-Object System.Drawing.Font("Consolas", 9)

$script:Rows = New-Object System.Collections.Generic.List[object]
$script:Dirty = $false
$script:ExtraColumns = @()
$script:IsAdmin = Test-TravelAccessAdministrator

function Get-DeskConfig {
    if (!(Test-Path -LiteralPath $script:ConfigPath)) { return $null }
    try {
        return Get-Content -LiteralPath $script:ConfigPath -Raw | ConvertFrom-Json
    }
    catch {
        return $null
    }
}

function Get-DeskTimeZoneId {
    $config = Get-DeskConfig
    if ($config -and $config.TimeZone) { return [string]$config.TimeZone }
    return Get-TravelAccessLocalTimeZoneId
}

function Show-DeskMessage {
    param(
        [string]$Text,
        [string]$Caption = "Travel Access Scheduler",
        [System.Windows.Forms.MessageBoxButtons]$Buttons = "OK",
        [System.Windows.Forms.MessageBoxIcon]$Icon = "Information"
    )

    $owner = $null
    if ($script:ActiveDialog -and -not $script:ActiveDialog.IsDisposed) { $owner = $script:ActiveDialog }
    elseif ($script:Form -and -not $script:Form.IsDisposed) { $owner = $script:Form }
    if ($owner) {
        return [System.Windows.Forms.MessageBox]::Show($owner, $Text, $Caption, $Buttons, $Icon)
    }
    return [System.Windows.Forms.MessageBox]::Show($Text, $Caption, $Buttons, $Icon)
}

function New-DeskButton {
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [ValidateSet("Primary", "Secondary", "Danger")][string]$Kind = "Secondary",
        [int]$Width = 132
    )

    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Width = $Width
    $button.Height = 32
    $button.Font = $script:FontUi
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $button.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
    $button.UseVisualStyleBackColor = $false

    switch ($Kind) {
        "Primary" {
            $button.BackColor = $script:ColorAccent
            $button.ForeColor = [System.Drawing.Color]::White
            $button.FlatAppearance.BorderSize = 0
        }
        "Danger" {
            $button.BackColor = $script:ColorDanger
            $button.ForeColor = [System.Drawing.Color]::White
            $button.FlatAppearance.BorderSize = 0
        }
        default {
            $button.BackColor = [System.Drawing.Color]::White
            $button.ForeColor = [System.Drawing.Color]::FromArgb(32, 37, 41)
            $button.FlatAppearance.BorderColor = $script:ColorLine
            $button.FlatAppearance.BorderSize = 1
        }
    }

    return $button
}

function New-DeskLabel {
    param(
        [string]$Text,
        [switch]$Bold,
        [switch]$Muted,
        [switch]$Wrap
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.AutoSize = $true
    $label.Font = $(if ($Bold) { $script:FontBold } else { $script:FontUi })
    $label.ForeColor = $(if ($Muted) { $script:ColorMuted } else { [System.Drawing.Color]::FromArgb(32, 37, 41) })
    $label.UseMnemonic = $false
    if ($Wrap) {
        $label.MaximumSize = New-Object System.Drawing.Size(920, 0)
    }
    return $label
}

function Get-DeskZoneLabel {
    $zoneId = Get-DeskTimeZoneId
    try {
        $zone = [System.TimeZoneInfo]::FindSystemTimeZoneById($zoneId)
        $now = [System.TimeZoneInfo]::ConvertTimeFromUtc([datetime]::UtcNow, $zone)
        $totalMinutes = [int]$zone.GetUtcOffset($now).TotalMinutes
        $sign = if ($totalMinutes -lt 0) { "-" } else { "+" }
        $absMinutes = [Math]::Abs($totalMinutes)
        $hours = [Math]::Floor($absMinutes / 60)
        $minutes = $absMinutes % 60
        $offset = "UTC{0}{1:00}:{2:00}" -f $sign, $hours, $minutes
        return "$offset  $zoneId"
    }
    catch {
        return $zoneId
    }
}

function Update-Footer {
    $state = if ($script:Dirty) { "Unsaved changes" } else { "Schedule saved" }
    $rights = if ($script:IsAdmin) { "Administrator" } else { "Limited rights" }
    $script:FooterLabel.Text = "Version $($script:Version)    $state    $rights"
    if ($script:FooterZone) {
        $script:FooterZone.Text = Get-DeskZoneLabel
        if ($script:FooterZone.Parent) {
            $script:FooterZone.Left = [Math]::Max(180, $script:FooterZone.Parent.ClientSize.Width - $script:FooterZone.Width - 16)
        }
    }
}

function Save-DeskTimeZone {
    param([string]$ZoneId)

    if ([string]::IsNullOrWhiteSpace($ZoneId)) { return }
    if (!(Test-Path -LiteralPath $script:ConfigPath)) {
        Show-DeskMessage "Install the scheduler before changing the time zone. Config.json is not in the data folder yet." "Travel Access Scheduler" "OK" "Warning"
        return
    }

    try {
        $null = [System.TimeZoneInfo]::FindSystemTimeZoneById($ZoneId)
    }
    catch {
        Show-DeskMessage "Windows does not have a time zone named '$ZoneId'." "Travel Access Scheduler" "OK" "Warning"
        return
    }

    try {
        $config = Get-Content -LiteralPath $script:ConfigPath -Raw | ConvertFrom-Json
        $config | Add-Member -NotePropertyName TimeZone -NotePropertyValue $ZoneId -Force
        $json = $config | ConvertTo-Json -Depth 6
        [System.IO.File]::WriteAllText($script:ConfigPath, $json)
    }
    catch {
        Show-DeskMessage "The time zone could not be saved. $($_.Exception.Message)" "Travel Access Scheduler" "OK" "Warning"
        return
    }

    Update-Footer
    Update-ScheduleGrid
    Show-DeskMessage "Travel dates now use $ZoneId. The next sync uses this time zone."
}

function Get-FilteredRows {
    $today = Get-TravelAccessToday -TimeZoneId (Get-DeskTimeZoneId)
    $query = ""
    if ($script:SearchBox) { $query = $script:SearchBox.Text.Trim() }
    $like = "*"
    if ($query) {
        $like = "*" + ($query.Replace("[", "[[]").Replace("*", "[*]").Replace("?", "[?]")) + "*"
    }

    $decorated = foreach ($row in @($script:Rows.ToArray())) {
        $status = Get-TravelRowStatus -StartDate $row.StartDate -EndDate $row.EndDate -Today $today
        $days = Get-TravelDayCount -StartDate $row.StartDate -EndDate $row.EndDate
        [pscustomobject]@{
            Row    = $row
            Status = $status
            Days   = $days
            Rank   = $(switch ($status) { "Active" { 0 } "Upcoming" { 1 } "Invalid" { 2 } "Ended" { 3 } default { 4 } })
        }
    }

    if ($query) {
        $decorated = @($decorated | Where-Object { $_.Row.UserPrincipalName -like $like })
    }

    return @($decorated | Sort-Object Rank, @{ Expression = { $_.Row.StartDate }; Descending = $false }, @{ Expression = { $_.Row.UserPrincipalName } })
}

function Update-ScheduleGrid {
    if (-not $script:ScheduleGrid) { return }

    $today = Get-TravelAccessToday -TimeZoneId (Get-DeskTimeZoneId)
    $visible = @(Get-FilteredRows)
    $rowSnapshot = @($script:Rows.ToArray())
    $activeCount = @($rowSnapshot | Where-Object { (Get-TravelRowStatus -StartDate $_.StartDate -EndDate $_.EndDate -Today $today) -eq "Active" }).Count
    $people = @($rowSnapshot | ForEach-Object { $_.UserPrincipalName.Trim().ToLowerInvariant() } | Where-Object { $_ } | Sort-Object -Unique).Count

    $script:SummaryLabel.Text = "$activeCount active today   ·   $($script:Rows.Count) trips   ·   $people people"
    $script:ScheduleGrid.Rows.Clear()

    foreach ($item in $visible) {
        $dayText = if ($null -eq $item.Days) { "" } else { [string]$item.Days }
        $index = $script:ScheduleGrid.Rows.Add(
            $item.Row.UserPrincipalName,
            $item.Row.StartDate,
            $item.Row.EndDate,
            $dayText,
            $item.Status
        )
        $gridRow = $script:ScheduleGrid.Rows[$index]
        $gridRow.Tag = $item.Row

        switch ($item.Status) {
            "Active" {
                $gridRow.DefaultCellStyle.BackColor = $script:ColorActive
                $gridRow.DefaultCellStyle.ForeColor = $script:ColorOk
                $gridRow.DefaultCellStyle.SelectionForeColor = $script:ColorOk
            }
            "Invalid" {
                $gridRow.DefaultCellStyle.BackColor = $script:ColorInvalid
                $gridRow.DefaultCellStyle.ForeColor = $script:ColorDanger
                $gridRow.DefaultCellStyle.SelectionForeColor = $script:ColorDanger
            }
            "Ended" {
                $gridRow.DefaultCellStyle.ForeColor = $script:ColorMuted
                $gridRow.DefaultCellStyle.SelectionForeColor = $script:ColorMuted
            }
        }
    }

    Update-Footer
}

function Import-ScheduleIntoDesk {
    param([switch]$Quiet)

    try {
        $loaded = Read-TravelScheduleFile -Path $script:SchedulePath
    }
    catch {
        Show-DeskMessage "The schedule file could not be read.`r`n`r`n$($_.Exception.Message)" -Icon Error | Out-Null
        return
    }
    $script:Rows = New-Object System.Collections.Generic.List[object]
    foreach ($row in @($loaded.Rows)) {
        [void]$script:Rows.Add($row)
    }
    $script:ExtraColumns = @($loaded.ExtraColumns)
    $script:Dirty = $false
    Update-ScheduleGrid

    if (-not $Quiet -and $script:ExtraColumns.Count -gt 0) {
        Show-DeskMessage "This file has extra columns ($($script:ExtraColumns -join ', ')). Saving from Travel Desk keeps only UserPrincipalName, StartDate and EndDate." -Icon Warning | Out-Null
    }
}

function Save-ScheduleFromDesk {
    $invalid = @($script:Rows.ToArray() | Where-Object {
        (Get-TravelRowStatus -StartDate $_.StartDate -EndDate $_.EndDate -Today (Get-Date).Date) -eq "Invalid"
    })

    if ($invalid.Count -gt 0) {
        $answer = Show-DeskMessage "There are $($invalid.Count) invalid row(s). Invalid rows grant no access. Save anyway?" -Buttons YesNo -Icon Warning
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return $false }
    }

    try {
        Save-TravelScheduleFile -Path $script:SchedulePath -Rows @($script:Rows.ToArray())
        $script:Dirty = $false
        Update-Footer
        return $true
    }
    catch {
        Show-DeskMessage $_.Exception.Message -Icon Error | Out-Null
        return $false
    }
}

function Test-TripOverlap {
    param(
        $Candidate,
        $Ignore
    )

    $candidateStatus = Get-TravelRowStatus -StartDate $Candidate.StartDate -EndDate $Candidate.EndDate -Today (Get-Date).Date
    if ($candidateStatus -eq "Invalid") { return $false }

    $start = [datetime]::ParseExact($Candidate.StartDate, "dd/MM/yyyy", [System.Globalization.CultureInfo]::InvariantCulture).Date
    $end = [datetime]::ParseExact($Candidate.EndDate, "dd/MM/yyyy", [System.Globalization.CultureInfo]::InvariantCulture).Date
    $upn = $Candidate.UserPrincipalName.Trim()

    foreach ($row in @($script:Rows.ToArray())) {
        if ($Ignore -and [object]::ReferenceEquals($row, $Ignore)) { continue }
        if ($row.UserPrincipalName.Trim() -ne $upn) { continue }
        if ((Get-TravelRowStatus -StartDate $row.StartDate -EndDate $row.EndDate -Today (Get-Date).Date) -eq "Invalid") { continue }

        $otherStart = [datetime]::ParseExact($row.StartDate, "dd/MM/yyyy", [System.Globalization.CultureInfo]::InvariantCulture).Date
        $otherEnd = [datetime]::ParseExact($row.EndDate, "dd/MM/yyyy", [System.Globalization.CultureInfo]::InvariantCulture).Date
        if ($start -le $otherEnd -and $otherStart -le $end) { return $true }
    }

    return $false
}

function Show-TripDialog {
    param($Existing)

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = $(if ($Existing) { "Edit trip" } else { "Add traveller" })
    $dialog.FormBorderStyle = "FixedDialog"
    $dialog.StartPosition = "CenterParent"
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false
    $dialog.ShowInTaskbar = $false
    $dialog.ClientSize = New-Object System.Drawing.Size(460, 248)
    $dialog.Font = $script:FontUi
    $dialog.BackColor = [System.Drawing.Color]::White

    $intro = New-DeskLabel "Use the person's Entra sign-in address. Dates are inclusive and follow the scheduler time zone." -Muted -Wrap
    $intro.Location = New-Object System.Drawing.Point(20, 16)
    $intro.MaximumSize = New-Object System.Drawing.Size(420, 40)
    $dialog.Controls.Add($intro)

    $upnLabel = New-DeskLabel "User principal name"
    $upnLabel.Location = New-Object System.Drawing.Point(20, 68)
    $dialog.Controls.Add($upnLabel)

    $upnBox = New-Object System.Windows.Forms.TextBox
    $upnBox.Location = New-Object System.Drawing.Point(20, 90)
    $upnBox.Size = New-Object System.Drawing.Size(420, 24)
    $upnBox.Font = $script:FontUi
    if ($Existing) { $upnBox.Text = $Existing.UserPrincipalName }
    $dialog.Controls.Add($upnBox)

    $startLabel = New-DeskLabel "Start"
    $startLabel.Location = New-Object System.Drawing.Point(20, 124)
    $dialog.Controls.Add($startLabel)

    $endLabel = New-DeskLabel "End"
    $endLabel.Location = New-Object System.Drawing.Point(240, 124)
    $dialog.Controls.Add($endLabel)

    $startPicker = New-Object System.Windows.Forms.DateTimePicker
    $endPicker = New-Object System.Windows.Forms.DateTimePicker
    foreach ($picker in @($startPicker, $endPicker)) {
        $picker.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
        $picker.CustomFormat = "dd/MM/yyyy"
        $picker.Width = 180
    }
    $startPicker.Location = New-Object System.Drawing.Point(20, 146)
    $endPicker.Location = New-Object System.Drawing.Point(240, 146)

    $today = Get-TravelAccessToday -TimeZoneId (Get-DeskTimeZoneId)
    $startPicker.Value = $today
    $endPicker.Value = $today.AddDays(7)
    if ($Existing) {
        try {
            $startPicker.Value = [datetime]::ParseExact($Existing.StartDate, "dd/MM/yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
            $endPicker.Value = [datetime]::ParseExact($Existing.EndDate, "dd/MM/yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
        }
        catch {
        }
    }
    $dialog.Controls.Add($startPicker)
    $dialog.Controls.Add($endPicker)

    $ok = New-DeskButton -Text $(if ($Existing) { "Save trip" } else { "Add trip" }) -Kind Primary -Width 110
    $ok.Location = New-Object System.Drawing.Point(230, 198)
    $cancel = New-DeskButton -Text "Cancel" -Width 100
    $cancel.Location = New-Object System.Drawing.Point(348, 198)
    $dialog.AcceptButton = $ok
    $dialog.CancelButton = $cancel
    $dialog.Controls.Add($ok)
    $dialog.Controls.Add($cancel)

    $script:TripResult = $null

    $ok.Add_Click({
        $upn = $upnBox.Text.Trim()
        if ($upn -notmatch '^[^@\s,]+@[^@\s,]+\.[^@\s,]+$') {
            Show-DeskMessage "Enter the Entra sign-in address, for example alex.chen@contoso.com." -Icon Warning | Out-Null
            return
        }

        $start = $startPicker.Value.Date
        $end = $endPicker.Value.Date
        if ($end -lt $start) {
            Show-DeskMessage "The end date is before the start date." -Icon Warning | Out-Null
            return
        }

        $candidate = [pscustomobject]@{
            UserPrincipalName = $upn
            StartDate         = $start.ToString("dd/MM/yyyy")
            EndDate           = $end.ToString("dd/MM/yyyy")
        }

        if (Test-TripOverlap -Candidate $candidate -Ignore $Existing) {
            $answer = Show-DeskMessage "This person already has a trip that overlaps these dates. Add it as well?" -Buttons YesNo -Icon Warning
            if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }

        $script:TripResult = $candidate
        $dialog.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dialog.Close()
    }.GetNewClosure())

    $cancel.Add_Click({
        $dialog.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $dialog.Close()
    })

    $script:ActiveDialog = $dialog
    try {
        $dialog.ShowDialog($script:Form) | Out-Null
    }
    finally {
        $script:ActiveDialog = $null
    }
    $dialog.Dispose()
    return $script:TripResult
}

function Get-SelectedTrip {
    if ($script:ScheduleGrid.SelectedRows.Count -lt 1) { return $null }
    return $script:ScheduleGrid.SelectedRows[0].Tag
}

function Add-Trip {
    $created = Show-TripDialog
    if (-not $created) { return }
    [void]$script:Rows.Add($created)
    $script:Dirty = $true
    Update-ScheduleGrid
}

function Edit-SelectedTrip {
    $selected = Get-SelectedTrip
    if (-not $selected) {
        Show-DeskMessage "Select a trip first." | Out-Null
        return
    }

    $edited = Show-TripDialog -Existing $selected
    if (-not $edited) { return }
    $selected.UserPrincipalName = $edited.UserPrincipalName
    $selected.StartDate = $edited.StartDate
    $selected.EndDate = $edited.EndDate
    $script:Dirty = $true
    Update-ScheduleGrid
}

function Remove-SelectedTrip {
    $selected = Get-SelectedTrip
    if (-not $selected) {
        Show-DeskMessage "Select a trip first." | Out-Null
        return
    }

    $today = Get-TravelAccessToday -TimeZoneId (Get-DeskTimeZoneId)
    $status = Get-TravelRowStatus -StartDate $selected.StartDate -EndDate $selected.EndDate -Today $today
    $message = "Remove $($selected.UserPrincipalName) ($($selected.StartDate) - $($selected.EndDate)) from the schedule?"
    if ($status -eq "Active") {
        $another = @($script:Rows.ToArray() | Where-Object {
            -not [object]::ReferenceEquals($_, $selected) -and
            $_.UserPrincipalName.Trim().Equals($selected.UserPrincipalName.Trim(), [StringComparison]::OrdinalIgnoreCase) -and
            (Get-TravelRowStatus -StartDate $_.StartDate -EndDate $_.EndDate -Today $today) -eq "Active"
        })
        if ($another.Count -gt 0) {
            $message += "`r`n`r`nThis trip is active today, and another trip still covers today. Access stays until that other trip ends."
        }
        else {
            $message += "`r`n`r`nThis trip is active today. After you save, the next sync removes this person from the travel group."
        }
    }
    else {
        $message += "`r`n`r`nNothing changes in Entra until you save and the scheduler runs."
    }

    $answer = Show-DeskMessage $message -Buttons YesNo -Icon Warning
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    for ($i = $script:Rows.Count - 1; $i -ge 0; $i--) {
        if ([object]::ReferenceEquals($script:Rows[$i], $selected)) {
            $script:Rows.RemoveAt($i)
            break
        }
    }
    $script:Dirty = $true
    Update-ScheduleGrid
}

function Start-DeskSync {
    if (-not $script:IsAdmin) {
        Show-DeskMessage "Run sync now needs an elevated Travel Desk window. Double-click TravelDesk.cmd and approve the prompt. Until then, the saved schedule is applied on the next scheduled run." -Icon Information | Out-Null
        return
    }

    $config = Get-DeskConfig
    if (-not $config -or -not $config.InstallCompleted) {
        Show-DeskMessage "The scheduler is not installed yet. Use the Setup tab to run the installer." -Icon Warning | Out-Null
        return
    }

    if ($script:Dirty) {
        $answer = Show-DeskMessage "Save the schedule before syncing?" -Buttons YesNoCancel -Icon Question
        if ($answer -eq [System.Windows.Forms.DialogResult]::Cancel) { return }
        if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
            if (-not (Save-ScheduleFromDesk)) { return }
        }
    }

    $taskName = Get-TravelAccessTaskName -ConfigPath $script:ConfigPath
    try {
        Start-ScheduledTask -TaskName $taskName
        Show-DeskMessage "Sync started. It usually finishes within a minute. Status and the runtime log update when it completes." | Out-Null
    }
    catch {
        Show-DeskMessage "The scheduled task could not be started.`r`n`r`n$($_.Exception.Message)" -Icon Error | Out-Null
    }
}

function Set-DeskRunNote {
    param([string]$NextRun)

    if (-not $script:RunNote -or $script:RunNote.IsDisposed) { return }
    $when = "checking..."
    if (-not [string]::IsNullOrWhiteSpace($NextRun)) { $when = $NextRun }
    $script:RunNote.Text = "You can close this window. A scheduled task updates access for you. Next update: $when."
}

function Set-DeskStatusText {
    param([string]$Report, [string]$Log, [string]$NextRun)

    if ($script:StatusReport -and -not $script:StatusReport.IsDisposed) {
        $script:StatusReport.Text = $Report
    }
    if ($script:LogBox -and -not $script:LogBox.IsDisposed) {
        $script:LogBox.Text = $Log
    }
    if ($PSBoundParameters.ContainsKey("NextRun")) {
        Set-DeskRunNote $NextRun
    }
}

function Stop-DeskStatusLoad {
    if ($script:StatusTimer) {
        $script:StatusTimer.Stop()
        $script:StatusTimer.Dispose()
        $script:StatusTimer = $null
    }
    if ($script:StatusPowerShell) {
        $script:StatusPowerShell.Dispose()
        $script:StatusPowerShell = $null
    }
    if ($script:StatusRunspace) {
        $script:StatusRunspace.Close()
        $script:StatusRunspace.Dispose()
        $script:StatusRunspace = $null
    }
    $script:StatusBusy = $false
}

function Update-StatusView {
    if ($script:StatusBusy) { return }
    if (-not $script:StatusSnapshot -and $script:StatusReport -and [string]::IsNullOrWhiteSpace($script:StatusReport.Text)) {
        Set-DeskStatusText "Loading status..." "Loading the latest log..."
    }

    $script:StatusBusy = $true
    try {
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()
    $powerShell = [powershell]::Create()
    $powerShell.Runspace = $runspace
    [void]$powerShell.AddScript({
        param($BasePath, $DefaultTaskName)
        $ErrorActionPreference = "Stop"
        $lines = New-Object System.Collections.Generic.List[string]
        $config = $null
        $configPath = Join-Path $BasePath "Config.json"
        if (Test-Path -LiteralPath $configPath) {
            try { $config = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json } catch { $config = $null }
        }

        if (-not $config) {
            [void]$lines.Add("Not installed. Config.json is not in the data folder.")
            [void]$lines.Add("Use the Setup tab, then run Install.cmd.")
        }
        else {
            [void]$lines.Add(("Install complete").PadRight(22) + $(if ($config.InstallCompleted) { "Yes" } else { "No" }))
            [void]$lines.Add(("Organisation").PadRight(22) + $(if ($config.OrganisationName) { [string]$config.OrganisationName } else { "Not recorded" }))
            [void]$lines.Add(("Travel group").PadRight(22) + [string]$config.TravelGroupName)
            [void]$lines.Add(("Group object ID").PadRight(22) + [string]$config.TravelGroupId)
            [void]$lines.Add(("Tenant ID").PadRight(22) + [string]$config.TenantId)
            [void]$lines.Add(("Application ID").PadRight(22) + [string]$config.ClientId)
            [void]$lines.Add(("Time zone").PadRight(22) + [string]$config.TimeZone)
            [void]$lines.Add(("Sync").PadRight(22) + "At startup and every $([string]$config.SyncIntervalHours) hours")
        }

        [void]$lines.Add("")
        $taskName = $DefaultTaskName
        if ($config -and $config.TaskName -and -not [string]::IsNullOrWhiteSpace([string]$config.TaskName)) {
            $taskName = [string]$config.TaskName
        }
        $nextLabel = "not scheduled yet"
        try {
            Import-Module ScheduledTasks -ErrorAction SilentlyContinue
            $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
            if ($task) {
                $info = Get-ScheduledTaskInfo -TaskName $taskName
                [void]$lines.Add(("Scheduled task").PadRight(22) + $taskName)
                [void]$lines.Add(("Task state").PadRight(22) + [string]$task.State)
                [void]$lines.Add(("Last run").PadRight(22) + [string]$info.LastRunTime)
                [void]$lines.Add(("Last result").PadRight(22) + [string]$info.LastTaskResult)
                if ($info.NextRunTime -is [datetime] -and $info.NextRunTime.Year -gt 2000) {
                    $when = $info.NextRunTime
                    $clock = $when.ToString("h:mm tt")
                    $day = (Get-Date).Date
                    if ($when.Date -eq $day) { $nextLabel = "today at $clock" }
                    elseif ($when.Date -eq $day.AddDays(1)) { $nextLabel = "tomorrow at $clock" }
                    else { $nextLabel = ($when.ToString("dddd d MMMM")) + " at " + $clock }
                }
                [void]$lines.Add(("Next update").PadRight(22) + $nextLabel)
            }
            else {
                [void]$lines.Add("Scheduled task          Not installed")
                $nextLabel = "the task is not installed yet"
            }
        }
        catch {
            [void]$lines.Add("Scheduled task          " + $_.Exception.Message)
            $nextLabel = "could not be checked"
        }

        [void]$lines.Add("")
        try {
            $thumb = ""
            if ($config -and $config.CertificateThumbprint) { $thumb = [string]$config.CertificateThumbprint }
            $cert = $null
            if ($thumb) { $cert = Get-Item "Cert:\LocalMachine\My\$thumb" -ErrorAction SilentlyContinue }
            if ($cert) {
                $days = [math]::Floor(($cert.NotAfter - (Get-Date)).TotalDays)
                [void]$lines.Add(("Certificate").PadRight(22) + $cert.Thumbprint)
                [void]$lines.Add(("Expires").PadRight(22) + $cert.NotAfter.ToString("dd/MM/yyyy") + " - " + [string]$days + " days")
            }
            elseif ($thumb) {
                [void]$lines.Add("Certificate            Missing from this computer. Run the installer again.")
            }
            else {
                [void]$lines.Add("Certificate            Not configured. Run the installer.")
            }
        }
        catch {
            [void]$lines.Add("Certificate            " + $_.Exception.Message)
        }

        $logText = "No runtime log yet. It is created the first time the scheduler runs."
        $logPath = Join-Path $BasePath "Logs\TravelAccess.log"
        if (Test-Path -LiteralPath $logPath) {
            try {
                $tail = @(Get-Content -LiteralPath $logPath -Tail 40 -ErrorAction Stop)
                if ($tail.Count -gt 0) { $logText = ($tail -join "`r`n") }
            }
            catch {
                $logText = "The runtime log could not be read. $($_.Exception.Message)"
            }
        }

        [pscustomobject]@{
            Report  = ($lines -join "`r`n")
            Log     = $logText
            NextRun = $nextLabel
        }
    }).AddArgument($script:BasePath).AddArgument($script:TravelAccessDefaultTaskName)

    $script:StatusRunspace = $runspace
    $script:StatusPowerShell = $powerShell
    $script:StatusHandle = $powerShell.BeginInvoke()
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 80
    $script:StatusTimer = $timer
    $timer.Add_Tick({
        if (-not $script:StatusHandle.IsCompleted) { return }
        $timer = $script:StatusTimer
        if ($timer) { $timer.Stop() }
        try {
            if ($script:Form -and -not $script:Form.IsDisposed) {
                $output = @($script:StatusPowerShell.EndInvoke($script:StatusHandle))
                $snap = $output | Select-Object -Last 1
                if ($snap) {
                    $script:StatusSnapshot = $snap
                    Set-DeskStatusText -Report ([string]$snap.Report) -Log ([string]$snap.Log) -NextRun ([string]$snap.NextRun)
                }
            }
        }
        catch {
            if ($script:Form -and -not $script:Form.IsDisposed) {
                Set-DeskStatusText -Report $_.Exception.Message -Log ""
            }
        }
        finally {
            Stop-DeskStatusLoad
        }
    })
    $timer.Start()
    }
    catch {
        Stop-DeskStatusLoad
        Set-DeskStatusText -Report $_.Exception.Message -Log ""
    }
}

function Open-External {
    param([string]$Target)
    try {
        Start-Process $Target
    }
    catch {
        Show-DeskMessage "Could not open $Target`r`n`r`n$($_.Exception.Message)" -Icon Error | Out-Null
    }
}

function Copy-DeskText {
    param([string]$Value, [string]$EmptyMessage)
    if ([string]::IsNullOrWhiteSpace($Value)) {
        Show-DeskMessage $EmptyMessage -Icon Information | Out-Null
        return
    }
    [System.Windows.Forms.Clipboard]::SetText($Value)
    Show-DeskMessage "Copied to the clipboard:`r`n$Value" | Out-Null
}

function New-DeskActionBar {
    param([System.Windows.Forms.Control]$Page)

    $bar = New-Object System.Windows.Forms.Panel
    $bar.Dock = "Bottom"
    $bar.Height = 64
    $bar.BackColor = [System.Drawing.Color]::White
    $Page.Controls.Add($bar)

    $line = New-Object System.Windows.Forms.Panel
    $line.Dock = "Top"
    $line.Height = 1
    $line.BackColor = $script:ColorLine
    $bar.Controls.Add($line)

    $flow = New-Object System.Windows.Forms.FlowLayoutPanel
    $flow.Dock = "Fill"
    $flow.FlowDirection = "LeftToRight"
    $flow.WrapContents = $false
    $flow.BackColor = [System.Drawing.Color]::White
    $flow.Padding = New-Object System.Windows.Forms.Padding(16, 15, 16, 0)
    $bar.Controls.Add($flow)
    $bar.Controls.SetChildIndex($flow, 0)
    return $flow
}

function Build-ScheduleTab {
    param($Page)

    $page.BackColor = $script:ColorCanvas
    $buttons = New-DeskActionBar $Page

    $root = New-Object System.Windows.Forms.Panel
    $root.Dock = "Fill"
    $root.Padding = New-Object System.Windows.Forms.Padding(16, 12, 16, 10)
    $root.BackColor = $script:ColorCanvas
    $page.Controls.Add($root)
    $page.Controls.SetChildIndex($root, 0)

    $script:RunNote = New-DeskLabel "You can close this window. A scheduled task updates access for you. Next update: checking..." -Bold
    $script:RunNote.AutoSize = $false
    $script:RunNote.Dock = "Top"
    $script:RunNote.Height = 36
    $script:RunNote.ForeColor = $script:ColorHeader
    $root.Controls.Add($script:RunNote)

    $help = New-DeskLabel "Add a person for the days they are overseas. Access starts on the start date and ends on the end date. Save, then you can close this window. Someone added by hand in Entra, who has never been in this list, is left alone." -Muted
    $help.AutoSize = $false
    $help.Dock = "Top"
    $help.Height = 48
    $help.Padding = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
    $root.Controls.Add($help)

    $add = New-DeskButton "Add traveller" Primary 130
    $edit = New-DeskButton "Edit" Secondary 80
    $remove = New-DeskButton "Remove" Danger 90
    $reload = New-DeskButton "Reload" Secondary 90
    $save = New-DeskButton "Save" Primary 90
    $sync = New-DeskButton "Run sync now" Secondary 130
    $add.Add_Click({ Add-Trip })
    $edit.Add_Click({ Edit-SelectedTrip })
    $remove.Add_Click({ Remove-SelectedTrip })
    $reload.Add_Click({
        if ($script:Dirty) {
            $answer = Show-DeskMessage "Discard unsaved schedule changes and reload the file?" -Buttons YesNo -Icon Warning
            if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }
        Import-ScheduleIntoDesk
    })
    $save.Add_Click({
        if (Save-ScheduleFromDesk) {
            Show-DeskMessage "Schedule saved. It is applied on the next sync, or when you click Run sync now." | Out-Null
        }
    })
    $sync.Add_Click({ Start-DeskSync })
    foreach ($button in @($add, $edit, $remove, $reload, $save, $sync)) {
        [void]$buttons.Controls.Add($button)
    }

    $script:SummaryLabel = New-DeskLabel " " -Bold
    $script:SummaryLabel.AutoSize = $false
    $script:SummaryLabel.Dock = "Top"
    $script:SummaryLabel.Height = 24
    $root.Controls.Add($script:SummaryLabel)

    $searchRow = New-Object System.Windows.Forms.Panel
    $searchRow.Dock = "Top"
    $searchRow.Height = 36
    $searchRow.BackColor = $script:ColorCanvas
    $root.Controls.Add($searchRow)

    $searchLabel = New-DeskLabel "Find" -Muted
    $searchLabel.Location = New-Object System.Drawing.Point(0, 6)
    $searchRow.Controls.Add($searchLabel)

    $script:SearchBox = New-Object System.Windows.Forms.TextBox
    $script:SearchBox.Location = New-Object System.Drawing.Point(40, 4)
    $script:SearchBox.Size = New-Object System.Drawing.Size(280, 24)
    $script:SearchBox.Font = $script:FontUi
    $script:SearchBox.Add_TextChanged({ Update-ScheduleGrid })
    $searchRow.Controls.Add($script:SearchBox)

    $script:ScheduleGrid = New-Object System.Windows.Forms.DataGridView
    $script:ScheduleGrid.Dock = "Fill"
    $script:ScheduleGrid.BackgroundColor = $script:ColorCanvas
    $script:ScheduleGrid.BorderStyle = "None"
    $script:ScheduleGrid.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
    $script:ScheduleGrid.RowHeadersVisible = $false
    $script:ScheduleGrid.AllowUserToAddRows = $false
    $script:ScheduleGrid.AllowUserToDeleteRows = $false
    $script:ScheduleGrid.AllowUserToResizeRows = $false
    $script:ScheduleGrid.MultiSelect = $false
    $script:ScheduleGrid.SelectionMode = "FullRowSelect"
    $script:ScheduleGrid.ReadOnly = $true
    $script:ScheduleGrid.AutoSizeColumnsMode = "Fill"
    $script:ScheduleGrid.Font = $script:FontUi
    $script:ScheduleGrid.EnableHeadersVisualStyles = $false
    $script:ScheduleGrid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(241, 244, 247)
    $script:ScheduleGrid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(32, 37, 41)
    $script:ScheduleGrid.ColumnHeadersDefaultCellStyle.Font = $script:FontBold
    $script:ScheduleGrid.ColumnHeadersHeight = 34
    $script:ScheduleGrid.RowTemplate.Height = 30
    $script:ScheduleGrid.ScrollBars = "Vertical"
    $script:ScheduleGrid.CellBorderStyle = "SingleHorizontal"
    $script:ScheduleGrid.GridColor = $script:ColorLine
    $script:ScheduleGrid.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(214, 232, 248)
    $script:ScheduleGrid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(32, 37, 41)
    [void]$script:ScheduleGrid.Columns.Add("User", "User principal name")
    [void]$script:ScheduleGrid.Columns.Add("Start", "Start")
    [void]$script:ScheduleGrid.Columns.Add("End", "End")
    [void]$script:ScheduleGrid.Columns.Add("Days", "Days")
    [void]$script:ScheduleGrid.Columns.Add("Status", "Status")
    $script:ScheduleGrid.Columns["Start"].FillWeight = 55
    $script:ScheduleGrid.Columns["End"].FillWeight = 55
    $script:ScheduleGrid.Columns["Days"].FillWeight = 35
    $script:ScheduleGrid.Columns["Status"].FillWeight = 50
    $script:ScheduleGrid.Columns["User"].FillWeight = 160
    $script:ScheduleGrid.Add_CellDoubleClick({ Edit-SelectedTrip })
    $root.Controls.Add($script:ScheduleGrid)

    # Highest index is docked first. Search, summary, then help sit above the grid.
    $root.Controls.SetChildIndex($searchRow, 0)
    $root.Controls.SetChildIndex($script:SummaryLabel, 0)
    $root.Controls.SetChildIndex($script:RunNote, 0)
    $root.Controls.SetChildIndex($help, 0)
    $root.Controls.SetChildIndex($script:ScheduleGrid, 0)
}

function Build-StatusTab {
    param($Page)

    $page.BackColor = $script:ColorCanvas
    $buttons = New-DeskActionBar $Page
    $refresh = New-DeskButton "Refresh" Primary 110
    $refresh.Add_Click({ Update-StatusView })
    [void]$buttons.Controls.Add($refresh)

    $root = New-Object System.Windows.Forms.Panel
    $root.Dock = "Fill"
    $root.Padding = New-Object System.Windows.Forms.Padding(16, 12, 16, 10)
    $root.BackColor = $script:ColorCanvas
    $page.Controls.Add($root)
    $page.Controls.SetChildIndex($root, 0)

    $script:StatusReport = New-Object System.Windows.Forms.TextBox
    $script:StatusReport.Multiline = $true
    $script:StatusReport.ReadOnly = $true
    $script:StatusReport.BorderStyle = "FixedSingle"
    $script:StatusReport.BackColor = [System.Drawing.Color]::White
    $script:StatusReport.Font = $script:FontMono
    $script:StatusReport.Dock = "Top"
    $script:StatusReport.Height = 230
    $script:StatusReport.ScrollBars = "Vertical"
    $script:StatusReport.Text = "Loading status..."
    $root.Controls.Add($script:StatusReport)

    $logLabel = New-DeskLabel "Latest runtime log" -Bold
    $logLabel.Dock = "Top"
    $logLabel.Padding = New-Object System.Windows.Forms.Padding(0, 12, 0, 6)
    $logLabel.Height = 36
    $root.Controls.Add($logLabel)

    $script:LogBox = New-Object System.Windows.Forms.TextBox
    $script:LogBox.Multiline = $true
    $script:LogBox.ReadOnly = $true
    $script:LogBox.BorderStyle = "FixedSingle"
    $script:LogBox.BackColor = [System.Drawing.Color]::White
    $script:LogBox.Font = $script:FontMono
    $script:LogBox.Dock = "Fill"
    $script:LogBox.ScrollBars = "Both"
    $script:LogBox.WordWrap = $false
    $script:LogBox.Text = "Loading the latest log..."
    $root.Controls.Add($script:LogBox)
    $root.Controls.SetChildIndex($script:StatusReport, 0)
    $root.Controls.SetChildIndex($logLabel, 0)
    $root.Controls.SetChildIndex($script:LogBox, 0)
}

function Build-SetupTab {
    param($Page)

    $page.BackColor = $script:ColorCanvas
    $root = New-Object System.Windows.Forms.Panel
    $root.Dock = "Fill"
    $root.Padding = New-Object System.Windows.Forms.Padding(16, 12, 16, 12)
    $root.BackColor = $script:ColorCanvas
    $page.Controls.Add($root)

    $bar = New-Object System.Windows.Forms.FlowLayoutPanel
    $bar.Dock = "Top"
    $bar.Height = 84
    $bar.WrapContents = $true
    $root.Controls.Add($bar)

    $install = New-DeskButton "Run installer" Primary 130
    $apps = New-DeskButton "Open app registrations" Secondary 180
    $ca = New-DeskButton "Open Conditional Access" Secondary 190
    $copyGroup = New-DeskButton "Copy group name" Secondary 150
    $copyId = New-DeskButton "Copy group ID" Secondary 130
    $install.Add_Click({
        Open-External (Join-Path $script:InstallRoot "Install.cmd")
    })
    $apps.Add_Click({ Open-External "https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade" })
    $ca.Add_Click({ Open-External "https://entra.microsoft.com/#view/Microsoft_AAD_ConditionalAccess/ConditionalAccessBlade" })
    $copyGroup.Add_Click({
        $config = Get-DeskConfig
        $name = if ($config -and $config.TravelGroupName) { [string]$config.TravelGroupName } else { "CA - Approved Overseas Travellers" }
        Copy-DeskText $name "The group name is not available yet."
    })
    $copyId.Add_Click({
        $config = Get-DeskConfig
        Copy-DeskText $(if ($config) { [string]$config.TravelGroupId } else { "" }) "The group has not been created yet. Run the installer first."
    })
    foreach ($button in @($install, $apps, $ca, $copyGroup, $copyId)) {
        [void]$bar.Controls.Add($button)
    }

    $zoneRow = New-Object System.Windows.Forms.Panel
    $zoneRow.Dock = "Top"
    $zoneRow.Height = 72
    $zoneRow.BackColor = [System.Drawing.Color]::White
    $root.Controls.Add($zoneRow)

    $zoneLabel = New-Object System.Windows.Forms.Label
    $zoneLabel.Text = "Time zone for trip dates"
    $zoneLabel.AutoSize = $true
    $zoneLabel.Font = $script:FontBold
    $zoneLabel.Location = New-Object System.Drawing.Point(0, 6)
    $zoneRow.Controls.Add($zoneLabel)

    $script:ZoneBox = New-Object System.Windows.Forms.ComboBox
    $script:ZoneBox.DropDownStyle = "DropDownList"
    $script:ZoneBox.Width = 420
    $script:ZoneBox.Font = $script:FontUi
    $script:ZoneBox.Location = New-Object System.Drawing.Point(0, 30)
    foreach ($zone in @([System.TimeZoneInfo]::GetSystemTimeZones() | Sort-Object BaseUtcOffset, Id)) {
        [void]$script:ZoneBox.Items.Add($zone.Id)
    }
    $currentZone = Get-DeskTimeZoneId
    $zoneIndex = $script:ZoneBox.Items.IndexOf($currentZone)
    if ($zoneIndex -lt 0) { $zoneIndex = 0 }
    if ($script:ZoneBox.Items.Count -gt 0) { $script:ZoneBox.SelectedIndex = $zoneIndex }
    $zoneRow.Controls.Add($script:ZoneBox)

    $saveZone = New-DeskButton "Save time zone" Primary 150
    $saveZone.Location = New-Object System.Drawing.Point(432, 28)
    $saveZone.Add_Click({
        if ($script:ZoneBox.SelectedItem) {
            Save-DeskTimeZone ([string]$script:ZoneBox.SelectedItem)
        }
    })
    $zoneRow.Controls.Add($saveZone)

    $guide = New-Object System.Windows.Forms.TextBox
    $guide.Multiline = $true
    $guide.ReadOnly = $true
    $guide.BorderStyle = "None"
    $guide.BackColor = [System.Drawing.Color]::White
    $guide.Font = $script:FontUi
    $guide.Dock = "Fill"
    $guide.WordWrap = $true
    $guide.ScrollBars = "Vertical"
    $guideText = @"
WHAT IT DOES
This tool makes one group. The group is the list of people who are travelling.

You still do one thing. Add that group to the policy that lets people use the company from overseas.

Already have that policy? Add this group to it.
No policy yet? Create one. Call it Travel access. Then add this group to it.

That policy is what lets those people reach the company while they are abroad.
If the policy blocks sign-in from other countries, add the group under Exclude. Exclude means these people are let through.

This window does not change the policy. A scheduled task puts people in the group for their trip, and takes them out when the trip ends. You can close this window.

GROUPS
You do not make the group yourself, unless you already have one to reuse.
Default name: CA - Approved Overseas Travellers

The group must be:
  - security enabled
  - assigned membership, not dynamic
  - not mail-enabled
  - not a Microsoft 365 group

Do not create a second group for removing access. When a trip ends, the scheduler removes the person from this same group.
A person added only in Entra, who has never appeared in the schedule, is not removed.

APP REGISTRATION
On a new install, choose "Create a dedicated app registration" and sign in as a Global Administrator.
The installer creates a single-tenant app and grants these Microsoft Graph application permissions:

  GroupMember.ReadWrite.All
  User.Read.All

If you create the app yourself, add those two permissions, grant admin consent, and choose "I already created one" in the installer.
Use an app that belongs only to this scheduler. The installer replaces the certificate on that app.

CERTIFICATE
You do not buy or upload a certificate. The installer creates one self-signed certificate in Local Computer, Personal.
Issued to and issued by are the same name. That is expected.
The private key stays on this server and is not exported.
Running the installer again reuses that certificate while it has more than 30 days left.
The certificate on the app registration is replaced, so Entra keeps one.
Any older certificate with the same name is removed from this computer.
A new certificate is created only when the current one is missing or due to expire within 30 days.

CONDITIONAL ACCESS
Open the policy that lets people use the company from overseas.
No policy yet? Create one and name it Travel access.

Add the travel group to that policy.
If the policy blocks other countries, the place to add the group is:

  Users  >  Exclude  >  Users and groups  >  the travel group

Leave MFA and the other sign-in rules as they are.
This tool never edits the policy.

WHERE IT RUNS
Extract the folder to a permanent path, for example C:\TravelAccess.
Do not run it from inside the ZIP.
A scheduled task runs at startup and every 2 hours, even when this window is closed.
Day to day, add and remove people here, then close the window. You only need the installer again to repair setup or rotate the certificate.

TIME ZONE
A trip date is that calendar day in the time zone above, including daylight saving for that zone.
A new install suggests this server's Windows time zone. An existing Config.json keeps the zone already saved there.
Change it here if the server clock is UTC but the travel day should follow another country. The next sync uses the saved zone.
"@
    # A Windows text box only starts a new line on CR+LF. The script file uses LF.
    $guide.Lines = @($guideText -split '\r?\n')
    $root.Controls.Add($guide)
    $root.Controls.SetChildIndex($bar, 0)
    $root.Controls.SetChildIndex($zoneRow, 0)
    $root.Controls.SetChildIndex($guide, 0)
}

$script:Form = New-Object System.Windows.Forms.Form
$script:Form.Text = "Travel Access Scheduler"
$script:Form.StartPosition = "CenterScreen"
$script:Form.ClientSize = New-Object System.Drawing.Size(1080, 700)
$script:Form.MinimumSize = New-Object System.Drawing.Size(920, 620)
$script:Form.Font = $script:FontUi
$script:Form.BackColor = $script:ColorHeader
$script:Form.KeyPreview = $true
try {
    $script:Form.GetType().InvokeMember(
        "DoubleBuffered",
        [System.Reflection.BindingFlags]::NonPublic -bor [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::SetProperty,
        $null,
        $script:Form,
        @($true)
    ) | Out-Null
}
catch {
}

$header = New-Object System.Windows.Forms.Panel
$header.Dock = "Top"
$header.Height = 72
$header.BackColor = $script:ColorHeader
$script:Form.Controls.Add($header)

$title = New-Object System.Windows.Forms.Label
$title.Text = "Travel Access Scheduler"
$title.Font = $script:FontTitle
$title.ForeColor = [System.Drawing.Color]::White
$title.AutoSize = $true
$title.Location = New-Object System.Drawing.Point(20, 12)
$header.Controls.Add($title)

$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = "Temporary overseas access for people who are travelling"
$subtitle.Font = $script:FontUi
$subtitle.ForeColor = [System.Drawing.Color]::FromArgb(214, 226, 236)
$subtitle.AutoSize = $true
$subtitle.Location = New-Object System.Drawing.Point(22, 42)
$header.Controls.Add($subtitle)

$footer = New-Object System.Windows.Forms.Panel
$footer.Dock = "Bottom"
$footer.Height = 32
$footer.BackColor = [System.Drawing.Color]::White
$script:Form.Controls.Add($footer)

$script:FooterLabel = New-Object System.Windows.Forms.Label
$script:FooterLabel.AutoSize = $false
$script:FooterLabel.Dock = "Fill"
$script:FooterLabel.TextAlign = "MiddleLeft"
$script:FooterLabel.Padding = New-Object System.Windows.Forms.Padding(16, 0, 480, 0)
$script:FooterLabel.ForeColor = $script:ColorMuted
$script:FooterLabel.Font = $script:FontUi
$footer.Controls.Add($script:FooterLabel)

$script:FooterZone = New-Object System.Windows.Forms.Label
$script:FooterZone.AutoSize = $true
$script:FooterZone.ForeColor = $script:ColorMuted
$script:FooterZone.Font = $script:FontBold
$script:FooterZone.Text = "Time zone"
$script:FooterZone.Top = 8
$footer.Controls.Add($script:FooterZone)
$footer.Add_Resize({ Update-Footer })
$script:ZoneTip = New-Object System.Windows.Forms.ToolTip
$script:ZoneTip.SetToolTip($script:FooterZone, "Trip dates are calendar days in this Windows time zone, including its daylight saving. Change it on the Setup tab.")

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = "Fill"
$tabs.Font = $script:FontUi
$tabs.SizeMode = [System.Windows.Forms.TabSizeMode]::Fixed
$tabs.ItemSize = New-Object System.Drawing.Size(150, 32)
$tabs.Padding = New-Object System.Drawing.Point(8, 4)
$script:Form.Controls.Add($tabs)

# Docking runs in reverse z-order. The header has to be laid out first,
# or it sits on top of the tab labels and hides Status and Setup.
$script:Form.Controls.SetChildIndex($header, 0)
$script:Form.Controls.SetChildIndex($footer, 0)
$script:Form.Controls.SetChildIndex($tabs, 0)

$schedulePage = New-Object System.Windows.Forms.TabPage
$schedulePage.Text = "Schedule"
$statusPage = New-Object System.Windows.Forms.TabPage
$statusPage.Text = "Status"
$setupPage = New-Object System.Windows.Forms.TabPage
$setupPage.Text = "Setup"
$tabs.TabPages.AddRange(@($schedulePage, $statusPage, $setupPage))

Build-ScheduleTab $schedulePage
Build-StatusTab $statusPage
Build-SetupTab $setupPage

$script:Tabs = $tabs
$tabs.Add_SelectedIndexChanged({
    if ($script:Tabs.SelectedIndex -eq 1) { Update-StatusView }
})

$script:Form.Add_KeyDown({
    if ($_.Control -and $_.KeyCode -eq [System.Windows.Forms.Keys]::S) {
        $_.SuppressKeyPress = $true
        if (Save-ScheduleFromDesk) {
            Update-Footer
        }
    }
    elseif ($_.KeyCode -eq [System.Windows.Forms.Keys]::Delete -and $script:Tabs.SelectedIndex -eq 0 -and $script:Form.ActiveControl -ne $script:SearchBox) {
        Remove-SelectedTrip
    }
})

$script:Form.Add_FormClosing({
    if (-not $script:Dirty) { return }
    $answer = Show-DeskMessage "Save schedule changes before closing?" -Buttons YesNoCancel -Icon Question
    if ($answer -eq [System.Windows.Forms.DialogResult]::Cancel) {
        $_.Cancel = $true
        return
    }
    if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
        if (-not (Save-ScheduleFromDesk)) {
            $_.Cancel = $true
        }
    }
})

Import-ScheduleIntoDesk -Quiet
Update-Footer
$script:Form.Add_Shown({ Update-StatusView })
$script:Form.Add_FormClosed({ Stop-DeskStatusLoad })
[void]$script:Form.ShowDialog()
$script:Form.Dispose()
