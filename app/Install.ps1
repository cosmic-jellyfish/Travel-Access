#Requires -RunAsAdministrator
[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

# The launchers stay in the parent folder. Settings and the schedule live in data\.
$InstallRoot  = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$BasePath     = Join-Path $InstallRoot "data"
$ConfigPath   = Join-Path $BasePath "Config.json"
$SchedulePath = Join-Path $BasePath "TravelSchedule.csv"
$LogsPath     = Join-Path $BasePath "Logs"
$InstallLog   = Join-Path $LogsPath "Install.log"
$CerPath      = Join-Path $BasePath "TravelAccess.cer"

$TaskName         = "Travel Access Scheduler"
$CertSubject      = "CN=Travel Access Scheduler"
$DefaultGroupName = "CA - Approved Overseas Travellers"

function Write-Step {
    param([string]$Text)
    Write-Host ""
    Write-Host "==> $Text" -ForegroundColor Cyan
}

function Write-Ok {
    param([string]$Text)
    Write-Host "[OK] $Text" -ForegroundColor Green
}

function Write-Warn {
    param([string]$Text)
    Write-Host "[WARN] $Text" -ForegroundColor Yellow
}

function Read-HostDefault {
    param(
        [string]$Prompt,
        [string]$Default
    )
    if ([string]::IsNullOrWhiteSpace($Default)) {
        return (Read-Host $Prompt).Trim()
    }

    $value = Read-Host "$Prompt [$Default]"
    if ([string]::IsNullOrWhiteSpace($value)) {
        return $Default
    }
    return $value.Trim()
}

function Escape-ODataString {
    param([string]$Value)
    return $Value.Replace("'", "''")
}

function Save-Config {
    param([hashtable]$Config)

    $Config |
        ConvertTo-Json -Depth 10 |
        Set-Content -Path $ConfigPath -Encoding UTF8

    Write-Ok "Config saved: $ConfigPath"
}

function Grant-SystemPrivateKeyRead {
    param([System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate)

    Write-Step "Ensuring SYSTEM can read the certificate private key"

    $rsa = $null
    try {
        $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)
        $keyPath = $null

        if ($rsa -is [System.Security.Cryptography.RSACng]) {
            $keyName = $rsa.Key.UniqueName
            $keyPath = Join-Path $env:ProgramData "Microsoft\Crypto\Keys\$keyName"
        }
        elseif ($rsa -is [System.Security.Cryptography.RSACryptoServiceProvider]) {
            $keyName = $rsa.CspKeyContainerInfo.UniqueKeyContainerName
            $keyPath = Join-Path $env:ProgramData "Microsoft\Crypto\RSA\MachineKeys\$keyName"
        }

        if (!$keyPath -or !(Test-Path $keyPath)) {
            Write-Warn "Could not resolve the private-key file path. LocalMachine certificates normally allow SYSTEM access; app-only validation later will still catch a problem."
            return
        }

        $acl = Get-Acl $keyPath
        $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            "SYSTEM",
            [System.Security.AccessControl.FileSystemRights]::Read,
            [System.Security.AccessControl.AccessControlType]::Allow
        )
        $acl.SetAccessRule($rule)
        Set-Acl -Path $keyPath -AclObject $acl

        Write-Ok "SYSTEM has read access to the private key."
    }
    finally {
        if ($rsa) { $rsa.Dispose() }
    }
}

function ConvertTo-GraphJson {
    param($Value)

    if ($null -eq $Value) { return "null" }

    if ($Value -is [string]) {
        $escaped = $Value.Replace('\', '\\').Replace('"', '\"').Replace("`r", '\r').Replace("`n", '\n').Replace("`t", '\t')
        return '"' + $escaped + '"'
    }

    if ($Value -is [bool]) {
        if ($Value) { return "true" }
        return "false"
    }

    if ($Value -is [datetime]) {
        return '"' + $Value.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + '"'
    }

    if ($Value -is [guid]) {
        return '"' + ([string]$Value) + '"'
    }

    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal] -or $Value -is [float]) {
        return ([string]$Value)
    }

    if ($Value -is [System.Collections.IDictionary]) {
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($key in @($Value.Keys)) {
            [void]$parts.Add((ConvertTo-GraphJson ([string]$key)) + ":" + (ConvertTo-GraphJson $Value[$key]))
        }
        return "{" + ($parts -join ",") + "}"
    }

    if ($Value -is [System.Collections.IEnumerable]) {
        # Windows PowerShell 5.1 throws "Argument types do not match" when
        # foreach enumerates a generic List[object]. Index the list instead.
        $parts = New-Object System.Collections.Generic.List[string]
        if ($Value -is [System.Collections.IList]) {
            for ($index = 0; $index -lt $Value.Count; $index++) {
                $item = $Value[$index]
                if ($null -eq $item) { continue }
                [void]$parts.Add((ConvertTo-GraphJson $item))
            }
        }
        else {
            $enumerator = $Value.GetEnumerator()
            while ($enumerator.MoveNext()) {
                if ($null -eq $enumerator.Current) { continue }
                [void]$parts.Add((ConvertTo-GraphJson $enumerator.Current))
            }
        }
        return "[" + ($parts -join ",") + "]"
    }

    throw "Cannot convert a value of type $($Value.GetType().FullName) for Microsoft Graph."
}

function Get-ScheduledTaskScriptPaths {
    param([string]$Name)

    $task = Get-ScheduledTask -TaskName $Name -ErrorAction SilentlyContinue
    if (-not $task) { return $null }

    $paths = New-Object System.Collections.Generic.List[string]
    foreach ($action in @($task.Actions)) {
        $arguments = [string]$action.Arguments
        $path = $null
        if ($arguments -match '-File\s+"([^"]+)"') { $path = $Matches[1] }
        elseif ($arguments -match "-File\s+'([^']+)'") { $path = $Matches[1] }
        elseif ($arguments -match '-File\s+(\S+)') { $path = $Matches[1] }
        if ($path) { [void]$paths.Add($path) }
    }

    return [pscustomobject]@{
        Paths = @($paths.ToArray())
    }
}

function Test-ThisInstallTaskPath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    try { $full = [System.IO.Path]::GetFullPath($Path) } catch { return $false }
    $sync = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "Sync.ps1"))
    $legacy = [System.IO.Path]::GetFullPath((Join-Path $InstallRoot "TravelAccess.ps1"))
    if ($full.Equals($sync, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    if ($full.Equals($legacy, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    return $false
}

function Assert-TaskTargetsThisFolder {
    param([string]$Name)

    $info = Get-ScheduledTaskScriptPaths -Name $Name
    if (-not $info) { return }

    $paths = @($info.Paths | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    if ($paths.Count -eq 0) {
        throw "Scheduled task '$Name' already exists, and its program could not be confirmed. Nothing was changed."
    }

    foreach ($path in $paths) {
        if (Test-ThisInstallTaskPath $path) { return }
    }

    $syncPath = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "Sync.ps1"))
    throw @"
Scheduled task '$Name' already runs a different copy:
  $($paths[0])

This folder runs:
  $syncPath

Nothing was changed. Copy these files into the folder the task already uses, then run Install.cmd there.
"@
}

function Resolve-InstallNames {
    param($ExistingConfig)

    $defaultTask = "Travel Access Scheduler"
    $defaultSubject = "CN=Travel Access Scheduler"
    $defaultApp = "Travel Access Scheduler"
    if ($script:TravelAccessDefaultTaskName) { $defaultTask = [string]$script:TravelAccessDefaultTaskName }
    if ($script:TravelAccessDefaultSubject) { $defaultSubject = [string]$script:TravelAccessDefaultSubject }
    if ($script:TravelAccessProductName) { $defaultApp = [string]$script:TravelAccessProductName }

    $taskToCheck = $defaultTask
    if ($ExistingConfig -and $ExistingConfig.TaskName -and -not [string]::IsNullOrWhiteSpace([string]$ExistingConfig.TaskName)) {
        $taskToCheck = [string]$ExistingConfig.TaskName
    }

    $pointsHere = $false
    $info = Get-ScheduledTaskScriptPaths -Name $taskToCheck
    if ($info) {
        foreach ($path in @($info.Paths)) {
            if (Test-ThisInstallTaskPath $path) { $pointsHere = $true }
        }
    }

    $hasConfigIdentity = $ExistingConfig -and -not [string]::IsNullOrWhiteSpace([string]$ExistingConfig.ClientId)
    if ($hasConfigIdentity -or $pointsHere) {
        $org = ""
        $task = $defaultTask
        $subject = $defaultSubject
        $app = $defaultApp
        if ($ExistingConfig) {
            if ($ExistingConfig.OrganisationName) { $org = [string]$ExistingConfig.OrganisationName }
            if ($ExistingConfig.TaskName) { $task = [string]$ExistingConfig.TaskName }
            if ($ExistingConfig.CertificateSubject) { $subject = [string]$ExistingConfig.CertificateSubject }
            if ($ExistingConfig.AppDisplayName) { $app = [string]$ExistingConfig.AppDisplayName }
        }

        return [pscustomobject]@{
            OrganisationName   = $org
            TaskName           = $task
            CertificateSubject = $subject
            AppDisplayName     = $app
            ExistingDeployment = $true
        }
    }

    Write-Host ""
    Write-Host "New installation" -ForegroundColor Cyan
    Write-Host "  The organisation name is used for the scheduled task, the certificate"
    Write-Host "  and the app registration. Letters, numbers, spaces and hyphens only."
    Write-Host "  Example: Contoso"
    Write-Host ""

    $org = ""
    while ($org -notmatch '^[A-Za-z0-9]([A-Za-z0-9 -]{0,38}[A-Za-z0-9])?$') {
        $org = (Read-Host "Organisation short name").Trim()
        if ($org -notmatch '^[A-Za-z0-9]([A-Za-z0-9 -]{0,38}[A-Za-z0-9])?$') {
            Write-Warn "Use 1-40 letters, numbers, spaces or hyphens, and do not start or end with a space."
        }
    }

    return [pscustomobject]@{
        OrganisationName   = $org
        TaskName           = "$org - Travel Access Scheduler"
        CertificateSubject = "CN=$org Travel Access Scheduler"
        AppDisplayName     = "$org Travel Access Scheduler"
        ExistingDeployment = $false
    }
}

function Get-GraphApplicationRoleIds {
    $uri = "https://graph.microsoft.com/v1.0/servicePrincipals(appId='00000003-0000-0000-c000-000000000000')?`$select=id,appRoles"
    $graph = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject

    $map = @{}
    foreach ($role in @($graph.appRoles)) {
        if (-not $role) { continue }
        $isApplication = $false
        foreach ($member in @($role.allowedMemberTypes)) {
            if ([string]$member -eq "Application") { $isApplication = $true }
        }
        if ($isApplication -and $role.value -and $role.id) {
            $map[[string]$role.value] = [string]$role.id
        }
    }

    foreach ($name in @("User.Read.All", "GroupMember.ReadWrite.All")) {
        if (-not $map.ContainsKey($name) -or [string]::IsNullOrWhiteSpace([string]$map[$name])) {
            throw "Microsoft Graph did not return the application permission '$name', so admin consent could not be granted."
        }
    }

    return [pscustomobject]@{
        GraphServicePrincipalId = [string]$graph.id
        UserReadAll             = [string]$map["User.Read.All"]
        GroupMemberReadWriteAll = [string]$map["GroupMember.ReadWrite.All"]
    }
}

function Find-TravelAccessApplication {
    param([string]$DisplayName)

    $filter = "displayName eq '$(Escape-ODataString $DisplayName)'"
    $encoded = [System.Uri]::EscapeDataString($filter)
    $uri = "https://graph.microsoft.com/v1.0/applications?`$filter=$encoded&`$select=id,appId,displayName,requiredResourceAccess"
    $response = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject

    $found = New-Object System.Collections.Generic.List[object]
    foreach ($item in @($response.value)) {
        if ($item -and $item.appId) { [void]$found.Add($item) }
    }
    return @($found.ToArray())
}

function New-TravelAccessApplication {
    param(
        [string]$DisplayName,
        $Roles
    )

    $resourceAccess = New-Object System.Collections.Generic.List[object]
    [void]$resourceAccess.Add(@{ id = $Roles.UserReadAll; type = "Role" })
    [void]$resourceAccess.Add(@{ id = $Roles.GroupMemberReadWriteAll; type = "Role" })

    $resources = New-Object System.Collections.Generic.List[object]
    [void]$resources.Add(@{
        resourceAppId  = "00000003-0000-0000-c000-000000000000"
        resourceAccess = $resourceAccess
    })

    $body = @{
        displayName            = $DisplayName
        signInAudience         = "AzureADMyOrg"
        requiredResourceAccess = $resources
    }

    return Invoke-MgGraphRequest `
        -Method POST `
        -Uri "https://graph.microsoft.com/v1.0/applications" `
        -Body (ConvertTo-GraphJson $body) `
        -ContentType "application/json" `
        -OutputType PSObject
}

function Get-ServicePrincipalObject {
    param([string]$AppId)

    $uri = "https://graph.microsoft.com/v1.0/servicePrincipals(appId='$AppId')?`$select=id,appId,displayName"
    try {
        return Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject
    }
    catch {
        if ($_.Exception.Message -match "Request_ResourceNotFound|ResourceNotFound|404") {
            return $null
        }
        throw
    }
}

function Grant-TravelAccessApplicationPermissions {
    param(
        $App,
        $Roles
    )

    $graphAppId = "00000003-0000-0000-c000-000000000000"
    $needed = @(
        @{ Id = $Roles.UserReadAll; Name = "User.Read.All" }
        @{ Id = $Roles.GroupMemberReadWriteAll; Name = "GroupMember.ReadWrite.All" }
    )

    $resources = New-Object System.Collections.Generic.List[object]
    $sawGraph = $false

    foreach ($resource in @($App.requiredResourceAccess)) {
        if (-not $resource -or -not $resource.resourceAppId) { continue }

        $access = New-Object System.Collections.Generic.List[object]
        foreach ($entry in @($resource.resourceAccess)) {
            if (-not $entry -or -not $entry.id) { continue }
            [void]$access.Add(@{
                id   = [string]$entry.id
                type = [string]$entry.type
            })
        }

        if ([string]$resource.resourceAppId -eq $graphAppId) {
            $sawGraph = $true
            foreach ($role in $needed) {
                $already = $false
                for ($accessIndex = 0; $accessIndex -lt $access.Count; $accessIndex++) {
                    $entry = $access[$accessIndex]
                    if ($entry.id -eq $role.Id -and $entry.type -eq "Role") { $already = $true }
                }
                if (-not $already) {
                    [void]$access.Add(@{ id = $role.Id; type = "Role" })
                }
            }
        }

        [void]$resources.Add(@{
            resourceAppId  = [string]$resource.resourceAppId
            resourceAccess = $access
        })
    }

    if (-not $sawGraph) {
        $access = New-Object System.Collections.Generic.List[object]
        foreach ($role in $needed) {
            [void]$access.Add(@{ id = $role.Id; type = "Role" })
        }
        [void]$resources.Add(@{
            resourceAppId  = $graphAppId
            resourceAccess = $access
        })
    }

    $patch = @{ requiredResourceAccess = $resources }
    Invoke-MgGraphRequest `
        -Method PATCH `
        -Uri "https://graph.microsoft.com/v1.0/applications/$($App.id)" `
        -Body (ConvertTo-GraphJson $patch) `
        -ContentType "application/json" | Out-Null

    Write-Ok "Application permissions are listed on the app registration."

    $clientId = [string]$App.appId
    $servicePrincipal = Get-ServicePrincipalObject -AppId $clientId
    if (-not $servicePrincipal) {
        $servicePrincipal = Invoke-MgGraphRequest `
            -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/servicePrincipals" `
            -Body (ConvertTo-GraphJson @{ appId = $clientId }) `
            -ContentType "application/json" `
            -OutputType PSObject

        if (-not $servicePrincipal.id) {
            Start-Sleep -Seconds 3
            $servicePrincipal = Get-ServicePrincipalObject -AppId $clientId
        }
    }

    if (-not $servicePrincipal -or -not $servicePrincipal.id) {
        throw "The app registration exists, but its enterprise application was not created."
    }

    foreach ($role in $needed) {
        $assignment = @{
            principalId = [string]$servicePrincipal.id
            resourceId  = [string]$Roles.GraphServicePrincipalId
            appRoleId   = [string]$role.Id
        }

        try {
            Invoke-MgGraphRequest `
                -Method POST `
                -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($Roles.GraphServicePrincipalId)/appRoleAssignedTo" `
                -Body (ConvertTo-GraphJson $assignment) `
                -ContentType "application/json" | Out-Null
            Write-Ok "Granted $($role.Name)."
        }
        catch {
            if ($_.Exception.Message -match "already exists") {
                Write-Ok "$($role.Name) was already granted."
            }
            else {
                throw "Could not grant admin consent for $($role.Name). Sign in as a Global Administrator, or grant User.Read.All and GroupMember.ReadWrite.All on the app registration and run Install.ps1 again. $($_.Exception.Message)"
            }
        }
    }
}

# Do not run from Explorer's compressed-folder temp area.
$tempRoot = [System.IO.Path]::GetFullPath($env:TEMP).TrimEnd('\')
$baseFull = [System.IO.Path]::GetFullPath($InstallRoot).TrimEnd('\')
if ($baseFull.StartsWith($tempRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
    Write-Host ""
    Write-Host "ERROR: This installer is running from a temporary folder." -ForegroundColor Red
    Write-Host ""
    Write-Host "Please EXTRACT the ZIP to a permanent folder first, for example:"
    Write-Host "  C:\TravelAccess"
    Write-Host ""
    Read-Host "Press ENTER to close" | Out-Null
    exit 1
}

$TranscriptStarted = $false
$ExitCode = 0

try {
    $commonPath = Join-Path $PSScriptRoot "Common.ps1"
    if (!(Test-Path -LiteralPath $commonPath)) {
        throw "Common.ps1 is missing from $PSScriptRoot."
    }
    . $commonPath
    $paths = Initialize-TravelAccessData -AppDirectory $PSScriptRoot
    $InstallRoot = $paths.InstallRoot
    $BasePath = $paths.DataPath
    $ConfigPath = Join-Path $BasePath "Config.json"
    $SchedulePath = Join-Path $BasePath "TravelSchedule.csv"
    $LogsPath = Join-Path $BasePath "Logs"
    $InstallLog = Join-Path $LogsPath "Install.log"
    $CerPath = Join-Path $BasePath "TravelAccess.cer"

    Start-Transcript -Path $InstallLog -Append -Force | Out-Null
    $TranscriptStarted = $true

    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host " $($script:TravelAccessProductName) - Installer $($script:TravelAccessVersion)" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Install folder:"
    Write-Host "  $InstallRoot" -ForegroundColor White
    Write-Host ""
    Write-Host "Schedule and settings:"
    Write-Host "  $BasePath" -ForegroundColor White
    Write-Host ""
    Write-Host "Install log:"
    Write-Host "  $InstallLog" -ForegroundColor White
    Write-Host ""

    Write-Host "This installs a scheduled task that adds and removes people in one"
    Write-Host "group while they are travelling. You can close Travel Desk afterwards."
    Write-Host "The task keeps running. It does not edit Conditional Access."
    Write-Host "You add that group to the policy that lets people in from overseas."
    Write-Host ""

    $ExistingConfig = $null
    if (Test-Path $ConfigPath) {
        try {
            $ExistingConfig = Get-Content $ConfigPath -Raw | ConvertFrom-Json
            Write-Ok "Existing Config.json found; values will be reused as defaults."
        }
        catch {
            Write-Warn "Existing Config.json exists but could not be parsed. It will be replaced."
        }
    }

    $names = Resolve-InstallNames -ExistingConfig $ExistingConfig
    $OrgName = [string]$names.OrganisationName
    $TaskName = [string]$names.TaskName
    $CertSubject = [string]$names.CertificateSubject
    $AppDisplayName = [string]$names.AppDisplayName

    if ($names.ExistingDeployment) {
        Write-Ok "Existing installation detected. The task, certificate and app names will stay as they are."
    }

    Assert-TaskTargetsThisFolder -Name $TaskName

    $CreateApp = $false
    $EnsurePermissions = $false

    $installWasCompleted = $false
    if ($ExistingConfig -and $ExistingConfig.InstallCompleted) {
        $installWasCompleted = [bool]$ExistingConfig.InstallCompleted
    }

    if ($names.ExistingDeployment -and $ExistingConfig.TenantId -and $ExistingConfig.ClientId -and $installWasCompleted) {
        $TenantId = Read-HostDefault "Tenant ID" ([string]$ExistingConfig.TenantId)
        $ClientId = Read-HostDefault "Dedicated App Registration Client ID" ([string]$ExistingConfig.ClientId)
    }
    elseif ($names.ExistingDeployment -and $ExistingConfig.TenantId -and $ExistingConfig.ClientId) {
        Write-Host ""
        Write-Host "A previous installation did not finish." -ForegroundColor Yellow
        Write-Host "  [1] Keep the saved app registration and grant its Graph permissions"
        Write-Host "  [2] Keep the saved app registration and do not change its permissions"
        $resumeChoice = Read-HostDefault "Choice" "1"
        if ($resumeChoice -eq "1") { $EnsurePermissions = $true }
        $TenantId = Read-HostDefault "Tenant ID" ([string]$ExistingConfig.TenantId)
        $ClientId = Read-HostDefault "Dedicated App Registration Client ID" ([string]$ExistingConfig.ClientId)
    }
    else {
        Write-Host ""
        Write-Host "App registration" -ForegroundColor Cyan
        Write-Host "  [1] Create a dedicated app registration now (recommended)"
        Write-Host "  [2] I already created one and have the Application (client) ID"
        Write-Host ""
        Write-Host "  Option 1 signs you in as a Global Administrator, creates a single-tenant"
        Write-Host "  app, and grants these Microsoft Graph application permissions:"
        Write-Host "    GroupMember.ReadWrite.All"
        Write-Host "    User.Read.All"
        $appChoice = Read-HostDefault "Choice" "1"
        $TenantId = Read-HostDefault "Tenant ID (Entra admin center > Overview)" ([string]$ExistingConfig.TenantId)
        if ($appChoice -eq "1") {
            $CreateApp = $true
            $EnsurePermissions = $true
            $ClientId = ""
        }
        else {
            $ClientId = Read-HostDefault "Dedicated App Registration Client ID" ([string]$ExistingConfig.ClientId)
        }
    }

    $defaultGroup = if ($ExistingConfig.TravelGroupName) {
        [string]$ExistingConfig.TravelGroupName
    } else {
        $DefaultGroupName
    }

    Write-Host ""
    Write-Host "Travel group" -ForegroundColor Cyan
    Write-Host "  You do not need to create this group first."
    Write-Host "  The installer creates one static security group, or reuses this name"
    Write-Host "  if that group already exists. Do not use a dynamic group or a"
    Write-Host "  Microsoft 365 group."
    Write-Host "  Afterwards, add this group to the policy that lets people use the"
    Write-Host "  company from overseas. If you do not have that policy, create one"
    Write-Host "  named Travel access. Leave MFA alone."
    Write-Host ""
    $GroupName = Read-HostDefault "Travel security group name" $defaultGroup

    $defaultZone = Get-TravelAccessLocalTimeZoneId
    if ($ExistingConfig -and $ExistingConfig.TimeZone) {
        $defaultZone = [string]$ExistingConfig.TimeZone
    }

    Write-Host ""
    Write-Host "Travel time zone" -ForegroundColor Cyan
    Write-Host "  A trip date is that calendar day in this time zone, including daylight saving."
    Write-Host "  The suggestion is this server's time zone. Examples:"
    Write-Host "    AUS Eastern Standard Time"
    Write-Host "    GMT Standard Time"
    Write-Host "    Pacific Standard Time"
    Write-Host "    Central European Standard Time"
    Write-Host ""

    $TimeZoneId = ""
    $zoneAccepted = $false
    while (-not $zoneAccepted) {
        $TimeZoneId = Read-HostDefault "Windows time zone id" $defaultZone
        try {
            $null = [System.TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneId)
            $zoneAccepted = $true
        }
        catch {
            Write-Warn "'$TimeZoneId' is not a Windows time zone id."
        }
    }

    if ([string]::IsNullOrWhiteSpace($TenantId)) { throw "Tenant ID is required." }
    if (-not $CreateApp -and [string]::IsNullOrWhiteSpace($ClientId)) { throw "Client ID is required." }
    if ([string]::IsNullOrWhiteSpace($GroupName)) { throw "Travel group name is required." }

    # IMPORTANT: write Config.json NOW, before Graph/certificate setup.
    # This makes partial installs recoverable and preserves user input.
    $Config = [ordered]@{
        Version               = $script:TravelAccessVersion
        InstallCompleted      = $false
        OrganisationName      = $OrgName
        TenantId              = $TenantId
        ClientId              = $(if ([string]::IsNullOrWhiteSpace($ClientId)) { $null } else { $ClientId })
        CertificateThumbprint = if ($ExistingConfig.CertificateThumbprint) { [string]$ExistingConfig.CertificateThumbprint } else { $null }
        CertificateExpires    = if ($ExistingConfig.CertificateExpires) { [string]$ExistingConfig.CertificateExpires } else { $null }
        CertificateSubject    = $CertSubject
        TravelGroupId         = if ($ExistingConfig.TravelGroupId) { [string]$ExistingConfig.TravelGroupId } else { $null }
        TravelGroupName       = $GroupName
        AppDisplayName        = $AppDisplayName
        TaskName              = $TaskName
        TimeZone              = $TimeZoneId
        ScheduleFile          = "TravelSchedule.csv"
        StateFile             = "State.json"
        LogFile               = "Logs\TravelAccess.log"
        SyncIntervalHours     = 2
        LastInstallAttempt    = (Get-Date).ToString("o")
    }

    Save-Config -Config $Config

    if (!(Test-Path $SchedulePath)) {
        "UserPrincipalName,StartDate,EndDate" |
            Set-Content -Path $SchedulePath -Encoding UTF8
        Write-Ok "Blank TravelSchedule.csv created."
    }

    Write-Step "Installing/loading required Microsoft Graph modules"

    # Directory operations use Invoke-MgGraphRequest. Only the authentication
    # module is loaded, which avoids Microsoft.Graph assembly conflicts on
    # Windows PowerShell 5.1.
    $Modules = @(
        "Microsoft.Graph.Authentication"
    )

    foreach ($Module in $Modules) {
        if (!(Get-Module -ListAvailable -Name $Module)) {
            Write-Host "Installing $Module..."
            Install-Module `
                -Name $Module `
                -Scope AllUsers `
                -Force `
                -AllowClobber `
                -Repository PSGallery
        }

        Import-Module $Module -Force
        Write-Ok "$Module loaded."
    }

    Write-Step "Creating or reusing the scheduler certificate"

    $Certificate = $null

    if ($Config.CertificateThumbprint) {
        $candidate = Get-Item "Cert:\LocalMachine\My\$($Config.CertificateThumbprint)" -ErrorAction SilentlyContinue

        if ($candidate -and $candidate.HasPrivateKey -and $candidate.NotAfter -gt (Get-Date).AddDays(30)) {
            $Certificate = $candidate
            Write-Ok "Reusing certificate $($Certificate.Thumbprint) (expires $($Certificate.NotAfter.ToString('yyyy-MM-dd')))."
        }
    }

    if (!$Certificate) {
        $Certificate = New-SelfSignedCertificate `
            -Subject $CertSubject `
            -CertStoreLocation "Cert:\LocalMachine\My" `
            -KeyAlgorithm RSA `
            -KeyLength 2048 `
            -HashAlgorithm SHA256 `
            -KeySpec Signature `
            -KeyExportPolicy NonExportable `
            -KeyUsage DigitalSignature `
            -NotAfter (Get-Date).AddYears(2)

        Write-Ok "Created certificate $($Certificate.Thumbprint) (expires $($Certificate.NotAfter.ToString('yyyy-MM-dd')))."
    }

    Grant-SystemPrivateKeyRead -Certificate $Certificate

    Export-Certificate -Cert $Certificate -FilePath $CerPath -Force | Out-Null
    Write-Ok "Public certificate exported to $CerPath"

    try {
        $friendlyName = ($CertSubject -replace '^CN=', '')
        Set-TravelAccessCertificateFriendlyName -Certificate $Certificate -FriendlyName $friendlyName
        Write-Ok "Certificate friendly name set to '$friendlyName'."
    }
    catch {
        Write-Warn "Could not set the certificate friendly name. $($_.Exception.Message)"
    }

    $Config.CertificateThumbprint = $Certificate.Thumbprint
    $Config.CertificateExpires = $Certificate.NotAfter.ToString("o")
    Save-Config -Config $Config

    $removedCertificates = @(Remove-TravelAccessCertificates -Subject ([string]$Certificate.Subject) -KeepThumbprint $Certificate.Thumbprint)
    if ($removedCertificates.Count -gt 0) {
        Write-Ok "Removed $($removedCertificates.Count) older certificate(s) with the same name."
    }

    Write-Step "One-time Microsoft Graph admin sign-in"

    Write-Host "A Microsoft sign-in window will open."
    if ($CreateApp -or $EnsurePermissions) {
        Write-Host "Sign in as a Global Administrator. This grants the app permission to"
        Write-Host "read users and update membership of the travel group."
    }
    else {
        Write-Host "Use an account that can manage App Registrations and security groups."
    }
    Write-Host ""

    $scopes = @("Application.ReadWrite.All", "Group.ReadWrite.All")
    if ($CreateApp -or $EnsurePermissions) {
        $scopes += @("AppRoleAssignment.ReadWrite.All", "Directory.Read.All")
    }

    Connect-MgGraph `
        -TenantId $TenantId `
        -Scopes $scopes `
        -ContextScope Process `
        -NoWelcome

    try {
        $signedInTenant = [string](Get-MgContext).TenantId
        if ($signedInTenant -and -not $signedInTenant.Equals($TenantId, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Signed in to tenant $signedInTenant, but this setup is for tenant $TenantId. The app registration and travel group were not changed."
        }

        $App = $null

        if ($CreateApp) {
            Write-Step "Locating or creating the dedicated App Registration"

            $foundApps = New-Object System.Collections.Generic.List[object]
            foreach ($candidate in (Find-TravelAccessApplication -DisplayName $AppDisplayName)) {
                if ($candidate -and $candidate.appId) {
                    [void]$foundApps.Add($candidate)
                }
            }

            if ($foundApps.Count -gt 1) {
                throw "More than one app registration is named '$AppDisplayName'. Rename the extras, or run the installer again and enter one Application (client) ID."
            }

            if ($foundApps.Count -eq 1) {
                $App = $foundApps[0]
                Write-Ok "Found existing app registration '$($App.displayName)'."
            }
            else {
                Write-Host "No app registration named '$AppDisplayName' exists yet. It will be created after you confirm."
            }
        }
        else {
            Write-Step "Locating the dedicated App Registration"

            # Direct Graph REST avoids intermittent Kiota/SDK failures on Windows PowerShell 5.1.
            $encodedClientId = [System.Uri]::EscapeDataString($ClientId)
            $appUri = "https://graph.microsoft.com/v1.0/applications(appId='$encodedClientId')?`$select=id,appId,displayName,keyCredentials,requiredResourceAccess"

            try {
                $App = Invoke-MgGraphRequest `
                    -Method GET `
                    -Uri $appUri `
                    -OutputType PSObject
            }
            catch {
                throw "Unable to retrieve App Registration with Client ID '$ClientId' via Microsoft Graph REST. $($_.Exception.Message)"
            }

            if (!$App -or !$App.id) {
                throw "No App Registration was found with Client ID '$ClientId'. Check that you entered the Application (client) ID from Entra App registrations."
            }
        }

        if ($App) {
            Write-Host "App:       $($App.displayName)"
            Write-Host "Client ID: $($App.appId)"
            Write-Host "Object ID: $($App.id)"
        }
        else {
            Write-Host "App:       $AppDisplayName"
            Write-Host "Client ID: will be created"
        }
        Write-Host ""
        Write-Warn "This installer assumes this App Registration is DEDICATED to the Travel Access Scheduler."
        Write-Warn "Its certificate credentials will be replaced with the certificate on this server."

        $continue = Read-Host "Type YES to continue"
        if ($continue -ne "YES") {
            throw "Installation cancelled by user."
        }

        $permissionRoles = $null
        if ($CreateApp -and -not $App) {
            Write-Step "Creating the dedicated App Registration"
            $permissionRoles = Get-GraphApplicationRoleIds
            $App = New-TravelAccessApplication -DisplayName $AppDisplayName -Roles $permissionRoles
            if (-not $App -or -not $App.appId) {
                throw "Microsoft Graph did not return the new app registration."
            }

            $ClientId = [string]$App.appId
            $Config.ClientId = $ClientId
            Save-Config -Config $Config
            Write-Ok "Created app registration '$AppDisplayName' ($ClientId)."
        }

        if ($EnsurePermissions) {
            Write-Step "Granting Microsoft Graph application permissions"
            if (-not $App -or -not $App.appId) {
                throw "The app registration does not have an Application (client) ID."
            }
            if (-not $permissionRoles) {
                $permissionRoles = Get-GraphApplicationRoleIds
            }
            Grant-TravelAccessApplicationPermissions -App $App -Roles $permissionRoles
            $ClientId = [string]$App.appId
            $Config.ClientId = $ClientId
            Save-Config -Config $Config
        }

        Write-Step "Uploading the public certificate to the App Registration"

        # One-element arrays collapse in Windows PowerShell, so the credential
        # is stored in a list and written with ConvertTo-GraphJson. PATCH
        # replaces the app's certificate list with this single public certificate.
        $keyCredentials = New-Object System.Collections.Generic.List[object]
        [void]$keyCredentials.Add(@{
            type          = "AsymmetricX509Cert"
            usage         = "Verify"
            key           = [Convert]::ToBase64String($Certificate.RawData)
            displayName   = $AppDisplayName
            startDateTime = $Certificate.NotBefore.ToUniversalTime().ToString("o")
            endDateTime   = $Certificate.NotAfter.ToUniversalTime().ToString("o")
        })

        $patchUri = "https://graph.microsoft.com/v1.0/applications/$($App.id)"

        try {
            Invoke-MgGraphRequest `
                -Method PATCH `
                -Uri $patchUri `
                -Body (ConvertTo-GraphJson @{ keyCredentials = $keyCredentials }) `
                -ContentType "application/json" | Out-Null
        }
        catch {
            throw "Failed to upload the public certificate to App Registration '$($App.displayName)'. $($_.Exception.Message)"
        }

        Write-Ok "Public certificate uploaded to the App Registration."

        Write-Step "Creating/finding the travel security group"

        # Direct Graph REST is used here instead of Get-MgGroup/New-MgGroup.
        # This avoids Microsoft.Graph.Authentication assembly mismatches.
        $escapedGroup = Escape-ODataString $GroupName
        $groupFilter = "displayName eq '$escapedGroup'"
        $encodedFilter = [System.Uri]::EscapeDataString($groupFilter)
        $groupLookupUri = "https://graph.microsoft.com/v1.0/groups?`$filter=$encodedFilter&`$select=id,displayName,securityEnabled,groupTypes"

        try {
            $groupLookup = Invoke-MgGraphRequest `
                -Method GET `
                -Uri $groupLookupUri `
                -OutputType PSObject
        }
        catch {
            throw "Failed to search for travel group '$GroupName' via Graph REST. $($_.Exception.Message)"
        }

        $Groups = @($groupLookup.value)

        if ($Groups.Count -gt 1) {
            throw "Multiple groups named '$GroupName' exist. Use a unique group name."
        }

        if ($Groups.Count -eq 1) {
            $TravelGroup = $Groups[0]

            if (!$TravelGroup.id) {
                throw "Graph returned the existing group '$GroupName' without an Object ID."
            }

            Write-Ok "Found existing group '$($TravelGroup.displayName)' ($($TravelGroup.id))."
        }
        else {
            $nicknameBase = ($GroupName -replace '[^A-Za-z0-9]', '')
            if ([string]::IsNullOrWhiteSpace($nicknameBase)) {
                $nicknameBase = "TravelAccess"
            }

            $mailNickname = "$nicknameBase-$((Get-Random -Minimum 1000 -Maximum 9999))"

            $groupBody = @{
                displayName     = $GroupName
                description     = "Members are temporarily excluded from the geographic sign-in block while the Travel Access Scheduler has an active trip for them."
                groupTypes      = @()
                mailEnabled     = $false
                mailNickname    = $mailNickname
                securityEnabled = $true
            }

            try {
                $TravelGroup = Invoke-MgGraphRequest `
                    -Method POST `
                    -Uri "https://graph.microsoft.com/v1.0/groups" `
                    -Body ($groupBody | ConvertTo-Json -Depth 10) `
                    -ContentType "application/json" `
                    -OutputType PSObject
            }
            catch {
                throw "Failed to create security group '$GroupName' via Graph REST. $($_.Exception.Message)"
            }

            if (!$TravelGroup -or !$TravelGroup.id) {
                throw "Graph did not return an Object ID after creating '$GroupName'. Installation cannot continue."
            }

            Write-Ok "Created security group '$GroupName' ($($TravelGroup.id))."
        }

        if (@($TravelGroup.groupTypes) -contains "DynamicMembership") {
            throw "The travel group is dynamic. This scheduler requires an assigned/static security group."
        }

        if (![bool]$TravelGroup.securityEnabled) {
            throw "The group '$GroupName' is not security-enabled."
        }

        $Config.TravelGroupId = [string]$TravelGroup.id
        Save-Config -Config $Config

        Write-Host ""
        Write-Host "============================================================" -ForegroundColor Yellow
        Write-Host " CONDITIONAL ACCESS ACTION REQUIRED" -ForegroundColor Yellow
        Write-Host "============================================================" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Add this group to the policy that lets people use the company"
        Write-Host "from overseas. If you do not have that policy, create one and"
        Write-Host "name it Travel access."
        Write-Host ""
        Write-Host "If the policy blocks sign-in from other countries, add the group"
        Write-Host "under Exclude. Exclude means these people are let through:"
        Write-Host ""
        Write-Host "  Users -> Exclude -> Users and groups -> $GroupName"
        Write-Host ""
        Write-Host "Group Object ID:"
        Write-Host "  $($TravelGroup.id)" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Leave MFA, device compliance, and sign-in frequency alone."
        Write-Host "This installer does not change Conditional Access policies."
        Write-Host ""

        $caDone = Read-Host "After this group has been added to that policy, type YES"
        if ($caDone -ne "YES") {
            throw "The travel group was not confirmed in Conditional Access. The scheduled task was not installed."
        }
    }
    finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    }

    Write-Step "Validating unattended app-only Graph access"

    $validated = $false
    $validationError = $null

    for ($attempt = 1; $attempt -le 4 -and -not $validated; $attempt++) {
        try {
            Connect-MgGraph `
                -TenantId $TenantId `
                -ClientId $ClientId `
                -CertificateThumbprint $Certificate.Thumbprint `
                -ContextScope Process `
                -NoWelcome

            $Context = Get-MgContext
            if ($Context.AuthType -ne "AppOnly") {
                throw "Authentication succeeded but AuthType was '$($Context.AuthType)', not AppOnly."
            }

            # Validate the exact REST operations the runtime needs.
            $validateGroupId = [string]$Config.TravelGroupId
            if ([string]::IsNullOrWhiteSpace($validateGroupId)) {
                throw "TravelGroupId is blank in Config.json. Group discovery/creation did not complete."
            }

            Invoke-MgGraphRequest `
                -Method GET `
                -Uri "https://graph.microsoft.com/v1.0/groups/$validateGroupId/members?`$top=1&`$select=id" `
                -OutputType PSObject | Out-Null

            Invoke-MgGraphRequest `
                -Method GET `
                -Uri "https://graph.microsoft.com/v1.0/users?`$top=1&`$select=id" `
                -OutputType PSObject | Out-Null

            Write-Ok "App-only authentication and runtime REST permission checks succeeded."
            $validated = $true
        }
        catch {
            $validationError = $_
            $retryable = $_.Exception.Message -match "Authorization_RequestDenied|Insufficient privileges|Forbidden|403"
            if ($attempt -lt 4 -and $retryable) {
                Write-Warn "App-only access is not ready yet (attempt $attempt of 4). Waiting 10 seconds for Entra to finish applying permissions."
                Start-Sleep -Seconds 10
            }
            else {
                break
            }
        }
        finally {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
        }
    }

    if (-not $validated) {
        Write-Host ""
        Write-Host "APP-ONLY VALIDATION FAILED" -ForegroundColor Red
        Write-Host ""
        Write-Host $validationError.Exception.Message -ForegroundColor Red
        Write-Host ""
        Write-Host "The App Registration requires these MICROSOFT GRAPH APPLICATION permissions:"
        Write-Host ""
        Write-Host "  GroupMember.ReadWrite.All"
        Write-Host "  User.Read.All"
        Write-Host ""
        Write-Host "Then click 'Grant admin consent'."
        Write-Host ""
        throw $validationError
    }

    Write-Step "Creating Windows Scheduled Task"

    $ScriptPath = Join-Path $PSScriptRoot "Sync.ps1"
    if (!(Test-Path -LiteralPath $ScriptPath)) {
        throw "Sync.ps1 is missing from '$PSScriptRoot'."
    }

    $PowerShellExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"

    $Action = New-ScheduledTaskAction `
        -Execute $PowerShellExe `
        -Argument "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$ScriptPath`""

    # Every 2 hours at :05, plus at startup.
    $Triggers = @()
    foreach ($hour in 0,2,4,6,8,10,12,14,16,18,20,22) {
        $Triggers += New-ScheduledTaskTrigger `
            -Daily `
            -At ([datetime]::Today.AddHours($hour).AddMinutes(5))
    }
    $Triggers += New-ScheduledTaskTrigger -AtStartup

    $Settings = New-ScheduledTaskSettingsSet `
        -StartWhenAvailable `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -MultipleInstances IgnoreNew `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 15)

    $Principal = New-ScheduledTaskPrincipal `
        -UserId "SYSTEM" `
        -LogonType ServiceAccount `
        -RunLevel Highest

    Register-ScheduledTask `
        -TaskName $TaskName `
        -Action $Action `
        -Trigger $Triggers `
        -Settings $Settings `
        -Principal $Principal `
        -Description "Synchronises approved overseas travellers with the Conditional Access travel exclusion group." `
        -Force | Out-Null

    Write-Ok "Scheduled task '$TaskName' created."
    Write-Ok "Runs at startup and every 2 hours at :05."
    Write-Ok "StartWhenAvailable is enabled."

    $Config.InstallCompleted = $true
    $Config.LastSuccessfulInstall = (Get-Date).ToString("o")
    Save-Config -Config $Config

    Write-Step "Running the scheduler immediately"

    Start-ScheduledTask -TaskName $TaskName
    Start-Sleep -Seconds 4

    $Task = Get-ScheduledTask -TaskName $TaskName
    $TaskInfo = Get-ScheduledTaskInfo -TaskName $TaskName

    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host " INSTALLATION COMPLETE" -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Config:"
    Write-Host "  $ConfigPath"
    Write-Host ""
    Write-Host "Schedule:"
    Write-Host "  $SchedulePath"
    Write-Host ""
    Write-Host "Task:"
    Write-Host "  $TaskName"
    Write-Host ""
    Write-Host "Task state:"
    Write-Host "  $($Task.State)"
    Write-Host ""
    Write-Host "Next run:"
    Write-Host "  $($TaskInfo.NextRunTime)"
    Write-Host ""
    Write-Host "Install log:"
    Write-Host "  $InstallLog"
    Write-Host ""
    Write-Host "Runtime log:"
    Write-Host "  $(Join-Path $BasePath 'Logs\TravelAccess.log')"
    Write-Host ""
    Write-Host "Day to day, double-click:"
    Write-Host "  TravelDesk.cmd"
    Write-Host ""
}
catch {
    $ExitCode = 1

    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Red
    Write-Host " INSTALLATION FAILED" -ForegroundColor Red
    Write-Host "============================================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "ERROR:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "Script line:"
    Write-Host $_.InvocationInfo.PositionMessage
    Write-Host ""
    Write-Host "Full error details have been saved to:"
    Write-Host "  $InstallLog" -ForegroundColor Yellow
    Write-Host ""

    try {
        Add-Content -Path $InstallLog -Value ""
        Add-Content -Path $InstallLog -Value "================ FULL ERROR ================"
        Add-Content -Path $InstallLog -Value ($_ | Format-List * -Force | Out-String)
        Add-Content -Path $InstallLog -Value "============================================"
    }
    catch {}
}
finally {
    if ($TranscriptStarted) {
        try { Stop-Transcript | Out-Null } catch {}
    }

    Write-Host ""
    if ($ExitCode -eq 0) {
        Write-Host "Installer finished successfully." -ForegroundColor Green
    }
    else {
        Write-Host "Installer stopped with an error. Nothing will silently close." -ForegroundColor Red
    }

    Write-Host ""
    Read-Host "Press ENTER to close" | Out-Null
}

exit $ExitCode
