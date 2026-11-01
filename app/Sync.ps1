[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

. (Join-Path $PSScriptRoot "Common.ps1")
$paths = Initialize-TravelAccessData -AppDirectory $PSScriptRoot
$BasePath   = $paths.DataPath
$ConfigPath = Join-Path $BasePath "Config.json"

if (!(Test-Path $ConfigPath)) {
    throw "Config.json not found in data. Run Install.cmd first."
}

$Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

if (!$Config.InstallCompleted) {
    throw "InstallCompleted is false in Config.json. Run Install.cmd again and finish setup."
}

$SchedulePath = Join-Path $BasePath $Config.ScheduleFile
$StatePath    = Join-Path $BasePath $Config.StateFile
$LogPath      = Join-Path $BasePath $Config.LogFile

$LogDirectory = Split-Path $LogPath -Parent
New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null

function Rotate-LogIfNeeded {
    if ((Test-Path $LogPath) -and ((Get-Item $LogPath).Length -gt 5MB)) {
        $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $archive = Join-Path $LogDirectory "TravelAccess-$stamp.log"
        Move-Item -Path $LogPath -Destination $archive -Force
    }
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","OK","ACTIVE","ADD","REMOVE","WARN","ERROR")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$timestamp [$Level] $Message"
    Write-Output $line
    Add-Content -Path $LogPath -Value $line -Encoding UTF8
}


function Convert-ToTravelDate {
    param([string]$DateText)

    return [datetime]::ParseExact(
        $DateText.Trim(),
        "dd/MM/yyyy",
        [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Globalization.DateTimeStyles]::None
    ).Date
}

function Normalize-Upn {
    param([string]$Upn)
    if ([string]::IsNullOrWhiteSpace($Upn)) { return $null }
    return $Upn.Trim().ToLowerInvariant()
}

function Get-GraphPagedValues {
    param([Parameter(Mandatory=$true)][string]$Uri)

    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Uri

    while ($next) {
        $response = Invoke-MgGraphRequest `
            -Method GET `
            -Uri $next `
            -OutputType PSObject

        foreach ($item in @($response.value)) {
            if ($null -ne $item) {
                $items.Add($item)
            }
        }

        $next = $response.'@odata.nextLink'
    }

    return $items.ToArray()
}

function Get-GraphUser {
    param(
        [Parameter(Mandatory=$true)]
        [string]$IdOrUpn
    )

    $value = $IdOrUpn.Trim()

    # Object ID lookup: use the direct /users/{id} endpoint.
    $parsedGuid = [Guid]::Empty
    if ([Guid]::TryParse($value, [ref]$parsedGuid)) {
        $encodedId = [System.Uri]::EscapeDataString($value)
        $uri = "https://graph.microsoft.com/v1.0/users/${encodedId}?`$select=id,displayName,userPrincipalName,accountEnabled"

        return Invoke-MgGraphRequest `
            -Method GET `
            -Uri $uri `
            -OutputType PSObject
    }

    # UPN lookup: use an OData filter rather than putting the UPN in the URL path.
    # This avoids PowerShell 5.1 / URI interpolation and encoding edge cases.
    $safeUpn = $value.Replace("'", "''")
    $filter = "userPrincipalName eq '$safeUpn'"
    $encodedFilter = [System.Uri]::EscapeDataString($filter)

    $uri = "https://graph.microsoft.com/v1.0/users?`$filter=$encodedFilter&`$select=id,displayName,userPrincipalName,accountEnabled&`$top=2"

    $response = Invoke-MgGraphRequest `
        -Method GET `
        -Uri $uri `
        -OutputType PSObject

    $matches = @($response.value)

    if ($matches.Count -eq 0) {
        throw "No Entra user matched UPN '$value'."
    }

    if ($matches.Count -gt 1) {
        throw "More than one Entra user unexpectedly matched UPN '$value'."
    }

    return $matches[0]
}

Rotate-LogIfNeeded

$Mutex = New-Object System.Threading.Mutex($false, "Global\TravelAccessScheduler")
$HasMutex = $false

try {
    try {
        $HasMutex = $Mutex.WaitOne(0)
    }
    catch [System.Threading.AbandonedMutexException] {
        $HasMutex = $true
    }

    if (!$HasMutex) {
        Write-Log "Another scheduler instance is already running. Exiting." "WARN"
        exit 0
    }

    Write-Log "============================================================"
    Write-Log "Travel Access Scheduler starting."

    if (!(Test-Path $SchedulePath)) {
        throw "Schedule file not found: $SchedulePath"
    }

    $TimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById([string]$Config.TimeZone)
    $NowEastern = [System.TimeZoneInfo]::ConvertTimeFromUtc([datetime]::UtcNow, $TimeZone)
    $Today = $NowEastern.Date

    Write-Log "Travel date ($($Config.TimeZone)): $($Today.ToString('dd/MM/yyyy')); scheduler time: $($NowEastern.ToString('yyyy-MM-dd HH:mm:ss'))."

    $cert = Get-Item "Cert:\LocalMachine\My\$($Config.CertificateThumbprint)" -ErrorAction SilentlyContinue
    if (!$cert) {
        throw "Configured certificate '$($Config.CertificateThumbprint)' was not found in LocalMachine\My."
    }
    if (!$cert.HasPrivateKey) {
        throw "Configured certificate exists but has no private key."
    }

    $daysRemaining = [math]::Floor(($cert.NotAfter - (Get-Date)).TotalDays)
    if ($daysRemaining -lt 30) {
        Write-Log "Authentication certificate expires in $daysRemaining day(s) on $($cert.NotAfter.ToString('yyyy-MM-dd')). Run Install.cmd again to rotate it." "WARN"
    }

    $Rows = @(Import-Csv $SchedulePath)

    $RequiredColumns = @("UserPrincipalName","StartDate","EndDate")
    if ($Rows.Count -gt 0) {
        $headers = @($Rows[0].PSObject.Properties.Name)
        foreach ($required in $RequiredColumns) {
            if ($headers -notcontains $required) {
                throw "TravelSchedule.csv is missing required column '$required'."
            }
        }
    }

    $PreviousStateUsers = @()
    if (Test-Path $StatePath) {
        try {
            $state = Get-Content $StatePath -Raw | ConvertFrom-Json
            $PreviousStateUsers = @($state.Users)
        }
        catch {
            Write-Log "State.json could not be parsed. Current CSV will still be processed." "WARN"
        }
    }

    $CurrentCsvUpns = @(
        $Rows |
        ForEach-Object { Normalize-Upn $_.UserPrincipalName } |
        Where-Object { $_ } |
        Sort-Object -Unique
    )

    $ActiveUpns = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $InvalidRows = 0

    foreach ($row in $Rows) {
        $upn = Normalize-Upn $row.UserPrincipalName
        if (!$upn) { continue }

        try {
            $start = Convert-ToTravelDate $row.StartDate
            $end   = Convert-ToTravelDate $row.EndDate

            if ($end -lt $start) {
                throw "End date is before start date."
            }

            if ($Today -ge $start -and $Today -le $end) {
                [void]$ActiveUpns.Add($upn)

                Write-Log "$upn | $($start.ToString('dd/MM/yyyy'))-$($end.ToString('dd/MM/yyyy'))" "ACTIVE"
            }
        }
        catch {
            $InvalidRows++
            Write-Log "Invalid schedule row for '$upn': $($_.Exception.Message). This row grants NO access." "ERROR"
        }
    }

    $PreviousByUpn = @{}
    foreach ($item in $PreviousStateUsers) {
        $pUpn = Normalize-Upn $item.UserPrincipalName
        if ($pUpn) {
            $PreviousByUpn[$pUpn] = $item
        }
    }

    $KnownUpns = @(
        @($CurrentCsvUpns) +
        @($PreviousByUpn.Keys)
    ) | Sort-Object -Unique

    Write-Log "CSV rows: $($Rows.Count); scheduled users: $($CurrentCsvUpns.Count); active today: $($ActiveUpns.Count)."

    Import-Module Microsoft.Graph.Authentication -Force

    Connect-MgGraph `
        -TenantId ([string]$Config.TenantId) `
        -ClientId ([string]$Config.ClientId) `
        -CertificateThumbprint ([string]$Config.CertificateThumbprint) `
        -ContextScope Process `
        -NoWelcome

    $HadErrors = ($InvalidRows -gt 0)
    $NextStateUsers = [System.Collections.Generic.List[object]]::new()

    try {
        $GroupId = [string]$Config.TravelGroupId
        $GroupDisplayName = [string]$Config.TravelGroupName

        if ([string]::IsNullOrWhiteSpace($GroupId)) {
            throw "TravelGroupId is blank in Config.json."
        }

        # REST-only membership read; handles pagination if the group grows.
        $CurrentMemberObjects = Get-GraphPagedValues `
            -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/members?`$select=id&`$top=999"

        $CurrentMemberIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($member in $CurrentMemberObjects) {
            if ($member.id) {
                [void]$CurrentMemberIds.Add([string]$member.id)
            }
        }

        foreach ($upn in $KnownUpns) {
            $InCurrentCsv = $CurrentCsvUpns -contains $upn
            $ShouldHaveAccess = $ActiveUpns.Contains($upn)

            $Previous = $PreviousByUpn[$upn]
            $User = $null
            $UserId = $null
            $DisplayName = $upn

            $LookupError = $null

            try {
                $User = Get-GraphUser -IdOrUpn $upn
            }
            catch {
                $LookupError = $_.Exception.Message

                if ($Previous -and $Previous.UserId) {
                    try {
                        $User = Get-GraphUser -IdOrUpn ([string]$Previous.UserId)
                    }
                    catch {
                        $LookupError = "$LookupError Previous Object ID lookup also failed: $($_.Exception.Message)"
                        $User = $null
                    }
                }
            }

            if ($User) {
                $UserId = [string]$User.id
                if ($User.displayName) {
                    $DisplayName = [string]$User.displayName
                }
            }
            elseif ($Previous -and $Previous.UserId) {
                # We still know the historical Object ID and can use it to
                # clean up stale membership even if the user was renamed/deleted.
                $UserId = [string]$Previous.UserId
            }

            if (!$UserId) {
                if (!$InCurrentCsv) {
                    # This is an old state-only entry (for example a corrected typo)
                    # with no known Object ID. It cannot correspond to membership
                    # created by this scheduler, so retire it instead of erroring forever.
                    Write-Log "Dropping stale state entry '$upn' because it is no longer in the CSV and has no known Object ID. Lookup detail: $LookupError" "WARN"
                    continue
                }

                Write-Log "Unable to resolve current CSV user '$upn'. Graph detail: $LookupError" "ERROR"
                $HadErrors = $true

                $NextStateUsers.Add([pscustomobject]@{
                    UserPrincipalName = $upn
                    UserId            = $null
                })

                continue
            }

            $HasAccess = $CurrentMemberIds.Contains($UserId)

            try {
                if ($ShouldHaveAccess -and !$HasAccess) {
                    $memberBody = @{
                        "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$UserId"
                    }

                    Invoke-MgGraphRequest `
                        -Method POST `
                        -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/members/`$ref" `
                        -Body ($memberBody | ConvertTo-Json -Depth 5) `
                        -ContentType "application/json" | Out-Null

                    [void]$CurrentMemberIds.Add($UserId)
                    Write-Log "$DisplayName <$upn> added to '$GroupDisplayName'." "ADD"
                }
                elseif (!$ShouldHaveAccess -and $HasAccess) {
                    Invoke-MgGraphRequest `
                        -Method DELETE `
                        -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/members/$UserId/`$ref" | Out-Null

                    [void]$CurrentMemberIds.Remove($UserId)
                    Write-Log "$DisplayName <$upn> removed from '$GroupDisplayName'." "REMOVE"
                }
                elseif ($ShouldHaveAccess) {
                    Write-Log "$DisplayName <$upn> already has the required overseas access." "OK"
                }
                else {
                    Write-Log "$DisplayName <$upn> does not require overseas access today." "OK"
                }

                if ($InCurrentCsv) {
                    $stateUpn = if ($User -and $User.userPrincipalName) {
                        Normalize-Upn $User.userPrincipalName
                    }
                    else {
                        $upn
                    }

                    $NextStateUsers.Add([pscustomobject]@{
                        UserPrincipalName = $stateUpn
                        UserId            = $UserId
                    })
                }
            }
            catch {
                Write-Log "Membership operation failed for '$upn': $($_.Exception.Message)" "ERROR"
                $HadErrors = $true

                $NextStateUsers.Add([pscustomobject]@{
                    UserPrincipalName = $upn
                    UserId            = $UserId
                })
            }
        }

        $newState = [ordered]@{
            LastRunUtc = [datetime]::UtcNow.ToString("o")
            Users      = @($NextStateUsers.ToArray() | Sort-Object UserPrincipalName -Unique)
        }

        $newState |
            ConvertTo-Json -Depth 5 |
            Set-Content -Path $StatePath -Encoding UTF8

        Write-Log "State saved. Managed users retained: $($newState.Users.Count)."
    }
    finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    }

    if ($HadErrors) {
        Write-Log "Scheduler completed with one or more errors. Review the log." "ERROR"
        exit 2
    }

    Write-Log "Scheduler completed successfully." "OK"
    Write-Log "============================================================"
    exit 0
}
catch {
    try {
        Write-Log "FATAL: $($_.Exception.Message)" "ERROR"
    }
    catch {
        Write-Error $_
    }

    exit 1
}
finally {
    if ($HasMutex) {
        try { $Mutex.ReleaseMutex() } catch {}
    }

    if ($Mutex) {
        $Mutex.Dispose()
    }
}
