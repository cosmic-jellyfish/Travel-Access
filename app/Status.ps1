[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$commonPath = Join-Path $PSScriptRoot "Common.ps1"
if (!(Test-Path -LiteralPath $commonPath)) {
    throw "Common.ps1 is missing from $PSScriptRoot."
}
. $commonPath
$paths = Initialize-TravelAccessData -AppDirectory $PSScriptRoot
$BasePath = $paths.DataPath
$ConfigPath   = Join-Path $BasePath "Config.json"
$SchedulePath = Join-Path $BasePath "TravelSchedule.csv"
$TaskName     = $script:TravelAccessDefaultTaskName
if (Get-Command Get-TravelAccessTaskName -ErrorAction SilentlyContinue) {
    $TaskName = Get-TravelAccessTaskName -ConfigPath $ConfigPath
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " $($script:TravelAccessProductName) - Status" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

if (!(Test-Path $ConfigPath)) {
    Write-Host "NOT INSTALLED: Config.json is missing." -ForegroundColor Red
    exit 1
}

$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

Write-Host "Version       : $($config.Version)"
Write-Host "Install done  : $($config.InstallCompleted)"
Write-Host "Travel group  : $($config.TravelGroupName)"
Write-Host "Group ID      : $($config.TravelGroupId)"
Write-Host "Time zone     : $($config.TimeZone)"
Write-Host "Interval      : Every $($config.SyncIntervalHours) hours + startup"
Write-Host ""

if ($config.CertificateThumbprint) {
    $cert = Get-Item "Cert:\LocalMachine\My\$($config.CertificateThumbprint)" -ErrorAction SilentlyContinue

    if ($cert) {
        $remaining = [math]::Floor(($cert.NotAfter - (Get-Date)).TotalDays)
        Write-Host "Certificate   : $($cert.Thumbprint)"
        Write-Host "Cert expires  : $($cert.NotAfter) ($remaining days remaining)"
    }
    else {
        Write-Host "Certificate   : MISSING" -ForegroundColor Red
    }
}
else {
    Write-Host "Certificate   : not configured" -ForegroundColor Yellow
}

Write-Host ""

$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($task) {
    $info = Get-ScheduledTaskInfo -TaskName $TaskName
    Write-Host "Task state    : $($task.State)"
    Write-Host "Last run      : $($info.LastRunTime)"
    Write-Host "Last result   : $($info.LastTaskResult)"
    Write-Host "Next run      : $($info.NextRunTime)"
}
else {
    Write-Host "Scheduled task: not installed" -ForegroundColor Yellow
}

Write-Host ""

if (Test-Path $SchedulePath) {
    $rows = @(Import-Csv $SchedulePath)
    $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById([string]$config.TimeZone)
    $today = [System.TimeZoneInfo]::ConvertTimeFromUtc([datetime]::UtcNow, $tz).Date

    $active = @()

    foreach ($r in $rows) {
        try {
            $s = [datetime]::ParseExact(
                $r.StartDate.Trim(),
                "dd/MM/yyyy",
                [Globalization.CultureInfo]::InvariantCulture
            ).Date

            $e = [datetime]::ParseExact(
                $r.EndDate.Trim(),
                "dd/MM/yyyy",
                [Globalization.CultureInfo]::InvariantCulture
            ).Date

            if ($today -ge $s -and $today -le $e) {
                $active += $r
            }
        }
        catch {}
    }

    Write-Host "Schedule rows  : $($rows.Count)"
    Write-Host "Active today   : $($active.Count)"

    foreach ($r in $active) {
        Write-Host "  ACTIVE -> $($r.UserPrincipalName) | $($r.StartDate)-$($r.EndDate)" -ForegroundColor Green
    }
}

$installLog = Join-Path $BasePath "Logs\Install.log"
if (Test-Path $installLog) {
    Write-Host ""
    Write-Host "Last 10 install-log lines:" -ForegroundColor Cyan
    Get-Content $installLog -Tail 10 | ForEach-Object {
        Write-Host "  $_"
    }
}

if ($config.LogFile) {
    $runtimeLog = Join-Path $BasePath $config.LogFile
    if (Test-Path $runtimeLog) {
        Write-Host ""
        Write-Host "Last 12 runtime-log lines:" -ForegroundColor Cyan
        Get-Content $runtimeLog -Tail 12 | ForEach-Object {
            Write-Host "  $_"
        }
    }
}

Write-Host ""
