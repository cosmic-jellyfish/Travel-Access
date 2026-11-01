# Shared helpers. Sync.ps1 loads this file for the data folder only.
# The membership rules stay in Sync.ps1.

function Initialize-TravelAccessData {
    param([Parameter(Mandatory = $true)][string]$AppDirectory)

    $installRoot = [System.IO.Path]::GetFullPath((Join-Path $AppDirectory ".."))
    $dataPath = Join-Path $installRoot "data"
    $logsPath = Join-Path $dataPath "Logs"
    New-Item -ItemType Directory -Path $logsPath -Force | Out-Null

    $legacyConfig = Join-Path $installRoot "Config.json"
    $dataConfig = Join-Path $dataPath "Config.json"
    if ((Test-Path -LiteralPath $legacyConfig) -and -not (Test-Path -LiteralPath $dataConfig)) {
        foreach ($name in @("Config.json", "State.json", "TravelSchedule.csv", "TravelAccess.cer")) {
            $from = Join-Path $installRoot $name
            $to = Join-Path $dataPath $name
            if ((Test-Path -LiteralPath $from) -and -not (Test-Path -LiteralPath $to)) {
                Move-Item -LiteralPath $from -Destination $to
            }
        }

        $legacyLogs = Join-Path $installRoot "Logs"
        if (Test-Path -LiteralPath $legacyLogs) {
            foreach ($file in @(Get-ChildItem -LiteralPath $legacyLogs -File -ErrorAction SilentlyContinue)) {
                $destination = Join-Path $logsPath $file.Name
                if (-not (Test-Path -LiteralPath $destination)) {
                    Move-Item -LiteralPath $file.FullName -Destination $destination
                }
            }
            if (-not (Get-ChildItem -LiteralPath $legacyLogs -Force -ErrorAction SilentlyContinue)) {
                Remove-Item -LiteralPath $legacyLogs -Force -ErrorAction SilentlyContinue
            }
        }
    }

    return [pscustomobject]@{
        InstallRoot = $installRoot
        DataPath    = $dataPath
    }
}

$script:TravelAccessVersion = "1.0.0"
$script:TravelAccessProductName = "Travel Access Scheduler"
$script:TravelAccessDefaultTaskName = "Travel Access Scheduler"
$script:TravelAccessDefaultSubject  = "CN=Travel Access Scheduler"

function Get-TravelAccessLocalTimeZoneId {
    try {
        $local = [System.TimeZoneInfo]::Local
        if ($local -and $local.Id) { return [string]$local.Id }
    }
    catch {
    }
    return "UTC"
}

function Test-TravelAccessAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-TravelAccessTaskName {
    param([Parameter(Mandatory = $true)][string]$ConfigPath)

    if (Test-Path -LiteralPath $ConfigPath) {
        try {
            $config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
            if ($config.TaskName -and -not [string]::IsNullOrWhiteSpace([string]$config.TaskName)) {
                return [string]$config.TaskName
            }
        }
        catch {
        }
    }

    return $script:TravelAccessDefaultTaskName
}

function Get-TravelAccessToday {
    param([string]$TimeZoneId)

    $resolved = $TimeZoneId
    if ([string]::IsNullOrWhiteSpace($resolved)) {
        $resolved = Get-TravelAccessLocalTimeZoneId
    }

    try {
        $timeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById($resolved)
    }
    catch {
        $timeZone = [System.TimeZoneInfo]::Local
    }

    return [System.TimeZoneInfo]::ConvertTimeFromUtc([datetime]::UtcNow, $timeZone).Date
}

function Get-TravelRowStatus {
    param(
        [string]$StartDate,
        [string]$EndDate,
        [datetime]$Today
    )

    try {
        $start = [datetime]::ParseExact(
            $StartDate.Trim(),
            "dd/MM/yyyy",
            [System.Globalization.CultureInfo]::InvariantCulture
        ).Date

        $end = [datetime]::ParseExact(
            $EndDate.Trim(),
            "dd/MM/yyyy",
            [System.Globalization.CultureInfo]::InvariantCulture
        ).Date

        if ($end -lt $start) { return "Invalid" }
        if ($Today -ge $start -and $Today -le $end) { return "Active" }
        if ($Today -lt $start) { return "Upcoming" }
        return "Ended"
    }
    catch {
        return "Invalid"
    }
}

function Get-TravelDayCount {
    param(
        [string]$StartDate,
        [string]$EndDate
    )

    try {
        $start = [datetime]::ParseExact(
            $StartDate.Trim(),
            "dd/MM/yyyy",
            [System.Globalization.CultureInfo]::InvariantCulture
        ).Date

        $end = [datetime]::ParseExact(
            $EndDate.Trim(),
            "dd/MM/yyyy",
            [System.Globalization.CultureInfo]::InvariantCulture
        ).Date

        if ($end -lt $start) { return $null }
        return [int]($end - $start).TotalDays + 1
    }
    catch {
        return $null
    }
}

function Read-TravelScheduleFile {
    param([Parameter(Mandatory = $true)][string]$Path)

    $extra = New-Object System.Collections.Generic.List[string]
    $clean = New-Object System.Collections.Generic.List[object]

    if (!(Test-Path -LiteralPath $Path)) {
        return [pscustomobject]@{
            Rows        = @()
            ExtraColumns = @()
            MissingFile = $true
        }
    }

    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -gt 0) {
        foreach ($name in @($rows[0].PSObject.Properties.Name)) {
            if ($name -notin @("UserPrincipalName", "StartDate", "EndDate")) {
                [void]$extra.Add([string]$name)
            }
        }
    }

    foreach ($row in $rows) {
        [void]$clean.Add([pscustomobject]@{
            UserPrincipalName = [string]$row.UserPrincipalName
            StartDate         = [string]$row.StartDate
            EndDate           = [string]$row.EndDate
        })
    }

    return [pscustomobject]@{
        Rows         = @($clean.ToArray())
        ExtraColumns = @($extra.ToArray())
        MissingFile  = $false
    }
}

function Save-TravelScheduleFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [AllowEmptyCollection()][object[]]$Rows = @()
    )

    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add("UserPrincipalName,StartDate,EndDate")

    $rowsToSave = @()
    if ($null -ne $Rows) {
        $rowsToSave = @($Rows | Where-Object { $null -ne $_ })
    }

    foreach ($row in $rowsToSave) {
        $upn = ([string]$row.UserPrincipalName).Trim()
        $start = ([string]$row.StartDate).Trim()
        $end = ([string]$row.EndDate).Trim()

        if ([string]::IsNullOrWhiteSpace($upn)) {
            throw "A row is missing the user principal name."
        }

        foreach ($value in @($upn, $start, $end)) {
            if ($value -match '[,"\r\n]') {
                throw "This value cannot be saved because it contains a comma or a quote: $value"
            }
        }

        if ($start -notmatch '^\d{2}/\d{2}/\d{4}$' -or $end -notmatch '^\d{2}/\d{2}/\d{4}$') {
            throw "Dates must use dd/MM/yyyy. Check the row for $upn ($start - $end)."
        }

        [void]$lines.Add("$upn,$start,$end")
    }

    $directory = Split-Path -Parent $Path
    if ($directory -and !(Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    $temp = "$Path.tmp"
    $utf8 = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllLines($temp, $lines.ToArray(), $utf8)

    try {
        if (Test-Path -LiteralPath $Path) {
            [System.IO.File]::Copy($temp, $Path, $true)
            Remove-Item -LiteralPath $temp -Force
        }
        else {
            [System.IO.File]::Move($temp, $Path)
        }
    }
    catch {
        if (Test-Path -LiteralPath $temp) {
            Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue
        }
        throw
    }
}

function Set-TravelAccessCertificateFriendlyName {
    param(
        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [Parameter(Mandatory = $true)][string]$FriendlyName
    )

    if (-not (Test-TravelAccessAdministrator)) {
        throw "Naming the certificate requires an elevated window."
    }

    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store(
        [System.Security.Cryptography.X509Certificates.StoreName]::My,
        [System.Security.Cryptography.X509Certificates.StoreLocation]::LocalMachine
    )
    $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    try {
        $found = $store.Certificates.Find(
            [System.Security.Cryptography.X509Certificates.X509FindType]::FindByThumbprint,
            $Certificate.Thumbprint,
            $false
        )

        if ($found.Count -lt 1) {
            throw "Certificate $($Certificate.Thumbprint) was not found while naming it."
        }

        $found[0].FriendlyName = $FriendlyName
    }
    finally {
        $store.Close()
    }
}

function Remove-TravelAccessCertificates {
    param(
        [Parameter(Mandatory = $true)][string]$Subject,
        [string]$KeepThumbprint
    )

    if (-not (Test-TravelAccessAdministrator)) {
        throw "Removing certificates requires an elevated window."
    }
    if ([string]::IsNullOrWhiteSpace($Subject)) {
        throw "A certificate name is required before any certificate is removed."
    }

    $removed = New-Object System.Collections.Generic.List[string]
    foreach ($cert in @(Get-ChildItem -Path "Cert:\LocalMachine\My" -ErrorAction Stop)) {
        if ([string]$cert.Subject -ne $Subject) { continue }
        if ($KeepThumbprint -and $cert.Thumbprint -eq $KeepThumbprint) { continue }
        Remove-Item -LiteralPath "Cert:\LocalMachine\My\$($cert.Thumbprint)" -Force
        [void]$removed.Add([string]$cert.Thumbprint)
    }

    return @($removed.ToArray())
}
