#Requires -RunAsAdministrator
[CmdletBinding()]
param(
    [switch]$RemoveCertificate
)

$ErrorActionPreference = "Stop"

$commonPath = Join-Path $PSScriptRoot "Common.ps1"
if (!(Test-Path -LiteralPath $commonPath)) {
    throw "Common.ps1 is missing from $PSScriptRoot."
}
. $commonPath
$paths = Initialize-TravelAccessData -AppDirectory $PSScriptRoot
$BasePath = $paths.DataPath
$ConfigPath = Join-Path $BasePath "Config.json"

$TaskName = Get-TravelAccessTaskName -ConfigPath $ConfigPath

Write-Host ""
Write-Host "$($script:TravelAccessProductName) - Uninstall" -ForegroundColor Cyan
Write-Host ""

$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

if ($task) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "[OK] Scheduled task removed." -ForegroundColor Green
}
else {
    Write-Host "[INFO] Scheduled task was not present."
}

if ($RemoveCertificate) {
    if (!(Test-Path -LiteralPath $ConfigPath)) {
        Write-Host "[WARN] Config.json was not found, so no certificates were removed." -ForegroundColor Yellow
    }
    else {
        $config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
        $subject = $script:TravelAccessDefaultSubject
        if ($config.CertificateSubject -and -not [string]::IsNullOrWhiteSpace([string]$config.CertificateSubject)) {
            $subject = [string]$config.CertificateSubject
        }
        $removedCerts = @(Remove-TravelAccessCertificates -Subject $subject)
        Write-Host "[OK] Removed $($removedCerts.Count) local certificate(s) for this scheduler." -ForegroundColor Green
        foreach ($thumb in $removedCerts) {
            Write-Host "     $thumb"
        }
        Write-Host "[NOTE] The certificate on the Entra app registration was NOT removed." -ForegroundColor Yellow
        Write-Host "       Delete it under Certificates & secrets if you are retiring that app." -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "The data folder was left in place:"
Write-Host "  $BasePath"
Write-Host "Delete the install folder manually if you no longer need the schedule or settings."
Write-Host ""
