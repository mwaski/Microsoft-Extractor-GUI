param(
    [switch]$NewCert,
    [switch]$Help,
    [switch]$ShowConsole,
    [string]$ConfigRoot = "C:\Microsoft_Extractor_GUI"
)

# ----------------------------------------
# Help
# ----------------------------------------
if ($Help) {
    Write-Host "============================================================"
    Write-Host "M365 App Install Script - Help"
    Write-Host "============================================================"
    Write-Host ""
    Write-Host "USAGE:"
    Write-Host "  .\M365_App_Install.ps1                         # Normal operation (Commercial cloud)"
    Write-Host "  .\M365_App_Install.ps1 -NewCert                # Replace existing certificates with new one"
    Write-Host "  .\M365_App_Install.ps1 -Help                   # Show this help message"
    Write-Host ""
    Write-Host "PARAMETERS:"
    Write-Host "  -NewCert           Replaces any existing certificates with a new certificate"
    Write-Host "                     Use this option when you want to rotate the certificate on an"
    Write-Host "                     existing application. Admin consent is not re-requested if the"
    Write-Host "                     application already existed."
    Write-Host ""
    Write-Host "  -ShowConsole       Keeps the console window visible behind the GUI"
    Write-Host ""
    Write-Host "  -Help              Shows this help message and exits"
    Write-Host ""
    exit
}

# ----------------------------------------
# Hide our console window (GUI app)
# ----------------------------------------
# Only hides when this process owns the console (launched via double-click /
# "Run with PowerShell"). When started from an existing terminal, more than one
# process is attached to the console, so it stays visible for development.
# -ShowConsole forces it to stay visible either way.
if (-not $ShowConsole) {
    try {
        Add-Type -Namespace Win32 -Name ConsoleUtil -MemberDefinition @'
[DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]  public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
[DllImport("kernel32.dll")] public static extern uint GetConsoleProcessList(uint[] processList, uint processCount);
'@
        $consoleWindow = [Win32.ConsoleUtil]::GetConsoleWindow()
        $attachedProcs = [Win32.ConsoleUtil]::GetConsoleProcessList((New-Object uint32[] 2), 2)
        if ($consoleWindow -ne [IntPtr]::Zero -and $attachedProcs -le 1) {
            [void][Win32.ConsoleUtil]::ShowWindow($consoleWindow, 0)   # 0 = SW_HIDE
        }
    }
    catch { }
}

# ----------------------------------------
# Embedded standalone installer (single source of truth).
# When the Install / Setup Tenant button is clicked, this text is written
# to a temp .ps1 and launched in a new powershell.exe window with
# -ConfigRoot (and -NewCert when applicable) passed through.
# ----------------------------------------
$installerScriptText = @'
param(
    [switch]$NewCert,
    [switch]$Help,
    [string]$ConfigRoot = "C:\M365_App"
)

# ----------------------------------------
# Help
# ----------------------------------------
if ($Help) {
    Write-Host "============================================================"
    Write-Host "M365 App Install Script - Help"
    Write-Host "============================================================"
    Write-Host ""
    Write-Host "USAGE:"
    Write-Host "  .\M365_App_Install.ps1                         # Normal operation (Commercial cloud)"
    Write-Host "  .\M365_App_Install.ps1 -NewCert                # Replace existing certificates with new one"
    Write-Host "  .\M365_App_Install.ps1 -Help                   # Show this help message"
    Write-Host ""
    Write-Host "PARAMETERS:"
    Write-Host "  -NewCert           Replaces any existing certificates with a new certificate"
    Write-Host "                     Use this option when you want to rotate the certificate on an"
    Write-Host "                     existing application. Admin consent is not re-requested if the"
    Write-Host "                     application already existed."
    Write-Host ""
    Write-Host "  -Help              Shows this help message and exits"
    Write-Host ""
    exit
}

# ----------------------------------------
# Script Version / Header
# ----------------------------------------
$ScriptVersion = "2.0"

Write-Host "============================================================"
Write-Host "M365 App Install Script - Version $ScriptVersion"
if ($NewCert) {
    Write-Host "Running with -NewCert: Will replace existing certificates on existing apps"
}
Write-Host "Cloud Environment: Commercial (graph.microsoft.com)"
Write-Host "============================================================"

Write-Host -ForegroundColor DarkMagenta "Script Version: $ScriptVersion"

# ----------------------------------------
# Cloud Endpoints (Commercial Only)
# ----------------------------------------
$GraphEndpoint = "https://graph.microsoft.com"
$LoginEndpoint = "https://login.microsoftonline.com"
$ExchangeEnvironment = "O365Default"
$GraphEnvironment = "Global"

Write-Host "Graph Endpoint: $GraphEndpoint"
Write-Host "Login Endpoint: $LoginEndpoint"
Write-Host "Exchange Environment: $ExchangeEnvironment"


Write-Host ""  # Add blank line for readability

# ----------------------------------------
# Module Import (Microsoft Graph)
# ----------------------------------------
try {
    Get-Module Microsoft.Graph* | Remove-Module -Force -ErrorAction SilentlyContinue

    Import-Module Microsoft.Graph.Authentication -Force -ErrorAction Stop
    Import-Module Microsoft.Graph.Applications -Force -ErrorAction Stop
    Import-Module Microsoft.Graph.Users -Force -ErrorAction Stop
    Import-Module Microsoft.Graph.DirectoryObjects -Force -ErrorAction Stop
}
catch {
    Write-Host "Failed to import Microsoft Graph modules: $($_.Exception.Message)"
    Write-Host "Attempting to uninstall and reinstall Microsoft Graph modules..."

    try {
        Write-Host "Uninstalling existing Microsoft Graph modules..."
        Uninstall-Module Microsoft.Graph -AllVersions -Force -ErrorAction SilentlyContinue

        Write-Host "Installing latest Microsoft Graph module..."
        Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber

        Write-Host "Re-importing modules..."
        Import-Module Microsoft.Graph.Authentication -Force -ErrorAction Stop
        Import-Module Microsoft.Graph.Applications -Force -ErrorAction Stop
        Import-Module Microsoft.Graph.Users -Force -ErrorAction Stop
        Import-Module Microsoft.Graph.DirectoryObjects -Force -ErrorAction Stop

        Write-Host "Microsoft Graph modules reinstalled and imported successfully."
    }
    catch {
        Write-Host "Failed to reinstall Microsoft Graph modules: $($_.Exception.Message)"
        Write-Host "Please manually run the following commands as Administrator:"
        Write-Host "Uninstall-Module Microsoft.Graph -AllVersions -Force"
        Write-Host "Install-Module Microsoft.Graph -Scope CurrentUser -Force"
        exit
    }
}

# ----------------------------------------
# Logging Helpers
# ----------------------------------------
$baseOutputDir = $ConfigRoot

Function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )

    $timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    $logEntry = "$timestamp [$Level] $Message"

    switch ($Level) {
        "INFO" { Write-Host $logEntry -ForegroundColor Green }
        "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
        "ERROR" { Write-Host $logEntry -ForegroundColor Red }
        default { Write-Host $logEntry -ForegroundColor White }
    }

    if ($Level -eq "ERROR") {
        if (-not [string]::IsNullOrEmpty($errorLogPath)) {
            try {
                $errorLogDir = Split-Path $errorLogPath -Parent
                if (-not (Test-Path $errorLogDir)) {
                    New-Item -ItemType Directory -Path $errorLogDir -Force | Out-Null
                }
                if (-not (Test-Path $errorLogPath)) {
                    New-Item -ItemType File -Path $errorLogPath -Force | Out-Null
                }
                Add-Content -Path $errorLogPath -Value $logEntry
            }
            catch {
                Write-Host "[WARNING] Could not write to error log file: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }
    else {
        if (-not [string]::IsNullOrEmpty($logFilePath)) {
            try {
                $logDir = Split-Path $logFilePath -Parent
                if (-not (Test-Path $logDir)) {
                    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
                }
                if (-not (Test-Path $logFilePath)) {
                    New-Item -ItemType File -Path $logFilePath -Force | Out-Null
                }
                Add-Content -Path $logFilePath -Value $logEntry
            }
            catch {
                Write-Host "[WARNING] Could not write to log file: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }
}

Function Test-PathExists {
    param ([string]$Path)
    if (-not (Test-Path -Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

Function Invoke-RetryOperation {
    param (
        [scriptblock]$Operation,
        [int]$MaxRetries = 3,
        [int]$RetryDelaySeconds = 5
    )

    for ($i = 1; $i -le $MaxRetries; $i++) {
        try {
            & $Operation
            return $true
        }
        catch {
            Write-Log "Attempt $i failed: $($_.Exception.Message)" -Level "WARNING"
            if ($i -lt $MaxRetries) {
                Write-Log "Retrying in $RetryDelaySeconds seconds..."
                Start-Sleep -Seconds $RetryDelaySeconds
            }
            else {
                Write-Log "All retry attempts failed." -Level "ERROR"
                throw
            }
        }
    }
}

# ----------------------------------------
# Connect to Microsoft Graph (Interactive)
# ----------------------------------------
Write-Host "Connecting to Microsoft Graph ($GraphEndpoint)..."
try {
    Connect-MgGraph -Scopes "RoleManagement.ReadWrite.Directory", "Application.ReadWrite.All", "Directory.ReadWrite.All", "RoleManagement.ReadWrite.Exchange", "Organization.Read.All", "AppRoleAssignment.ReadWrite.All" -NoWelcome -ErrorAction Stop
    Write-Host "Successfully connected to Microsoft Graph."
}
catch {
    Write-Host "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    Write-Host "Please ensure you have the necessary permissions and try again."
    exit
}

# ----------------------------------------
# Tenant / Domain Discovery
# ----------------------------------------
$appName = "Microsoft Extractor GUI"
$redirectUri = "https://www.google.com"

Write-Host "Retrieving the Tenant ID and Primary Domain..."
try {
    $context = Get-MgContext
    if (-not $context) {
        throw "Not authenticated to Microsoft Graph. Please run Connect-MgGraph first."
    }

    Write-Host "Connected to tenant: $($context.TenantId)"
    Write-Host "Account: $($context.Account)"

    $tenantId = $null
    $PrimaryDomain = $null

    try {
        Write-Log "Trying Method 1: Get-MgOrganization..." -Level "INFO"
        $tenant = Get-MgOrganization -ErrorAction Stop
        if ($tenant -and $tenant.Count -gt 0) {
            $tenantId = $tenant[0].Id
            Write-Log "Successfully retrieved tenant ID via Get-MgOrganization: $tenantId" -Level "INFO"
        }
    }
    catch {
        Write-Log "Method 1 failed: $($_.Exception.Message)" -Level "WARNING"
    }

    if (-not $tenantId) {
        Write-Log "Trying Method 2: Using context TenantId..." -Level "INFO"
        $tenantId = $context.TenantId
        if ($tenantId) {
            Write-Log "Using tenant ID from context: $tenantId" -Level "INFO"
        }
    }

    try {
        Write-Host "Trying to get domain information using Graph API..."
        $domainsResult = Invoke-MgGraphRequest -Method GET -Uri "$GraphEndpoint/v1.0/domains"

        if ($domainsResult.value -and $domainsResult.value.Count -gt 0) {
            $defaultDomain = $domainsResult.value | Where-Object { $_.isDefault -eq $true }
            if ($defaultDomain) {
                $PrimaryDomain = $defaultDomain.id
                Write-Host "Primary Domain (default): $PrimaryDomain"
            }
            else {
                $verifiedDomain = $domainsResult.value | Where-Object { $_.isVerified -eq $true } | Select-Object -First 1
                if ($verifiedDomain) {
                    $PrimaryDomain = $verifiedDomain.id
                    Write-Host "Using first verified domain: $PrimaryDomain"
                }
                else {
                    $PrimaryDomain = $domainsResult.value[0].id
                    Write-Host "Using first available domain: $PrimaryDomain"
                }
            }
        }
        else {
            Write-Host "No domains found in tenant"
            $PrimaryDomain = $null
        }
    }
    catch {
        Write-Host "Failed to retrieve domain information via Graph API: $($_.Exception.Message)"
        $PrimaryDomain = $null
    }

    if (-not $tenantId) {
        throw "Could not retrieve tenant ID using any method"
    }

    if (-not $PrimaryDomain) {
        Write-Host "Could not retrieve domain information. Using tenant ID as domain name..."
        $PrimaryDomain = $tenantId
        Write-Host "Using Tenant ID as domain: $PrimaryDomain"
    }

    Write-Host "Final Tenant ID: $tenantId"
    Write-Host "Final Primary Domain: $PrimaryDomain"
}
catch {
    Write-Host "Failed to retrieve tenant information: $($_.Exception.Message)"
    Write-Host "Error details: $($_.Exception.InnerException.Message)"
    Write-Host ""
    Write-Host "Troubleshooting suggestions:"
    Write-Host "1. Ensure you have the required permissions (Organization.Read.All)"
    Write-Host "2. Check if Conditional Access policies are blocking the connection"
    Write-Host "3. Verify MFA requirements are satisfied"
    Write-Host "4. Try running: Connect-MgGraph -Scopes 'Organization.Read.All' -Force"
    exit
}

# ----------------------------------------
# Initialize Output / Logging Paths
# ----------------------------------------
$orgFolderName = $PrimaryDomain -replace '[\\/:*?"<>|]', '_'
$outputDir = Join-Path -Path $baseOutputDir -ChildPath $orgFolderName
$configDir = Join-Path -Path $outputDir -ChildPath "Config"

@($outputDir, $configDir) | ForEach-Object {
    Test-PathExists -Path $_
}

$logFilePath = Join-Path -Path $outputDir -ChildPath "log.txt"
$errorLogPath = Join-Path -Path $outputDir -ChildPath "error_log.txt"

foreach ($file in @($logFilePath, $errorLogPath)) {
    if (-not (Test-Path -Path $file)) {
        New-Item -Path $file -ItemType File -Force | Out-Null
    }
}

Write-Log "Logging initialized. Output directory: $configDir" -Level "INFO"
Write-Log "All output files will be stored in: $configDir" -Level "INFO"

$ErrorActionPreference = "Continue"
$global:Error.Clear()

# ----------------------------------------
# API Permissions Helper
# ----------------------------------------
# Application permission IDs (Graph + Exchange Online).
# Used both for the requiredResourceAccess manifest and for programmatic admin consent.
$script:graphResourceAppId    = "00000003-0000-0000-c000-000000000000"
$script:exchangeResourceAppId = "00000002-0000-0ff1-ce00-000000000000"

$script:graphPermissions = @(
    @{ id = "1bfefb4e-e0b5-418b-a88f-73c46d2cc8e9"; type = "Role" }, # Application.ReadWrite.All
    @{ id = "9a5d68dd-52b0-4cc2-bd40-abcf44ac3a30"; type = "Role" }, # Application.Read.All
    @{ id = "b0afded3-3588-46d8-8b3d-9842eff778da"; type = "Role" }, # AuditLog.Read.All
    @{ id = "5e1e9171-754d-478c-812c-f1755a9a4c2d"; type = "Role" }, # AuditLogsQuery.Read.All
    @{ id = "7ab1d382-f21e-4acd-a863-ba3e13f7da61"; type = "Role" }, # Directory.Read.All
    @{ id = "6e472fd1-ad78-48da-a0f0-97ab2c6b769e"; type = "Role" }, # IdentityRiskEvent.Read.All
    @{ id = "dc5007c0-2d7d-4c42-879c-2dab87571379"; type = "Role" }, # IdentityRiskyUser.Read.All
    @{ id = "693c5e45-0940-467d-9b8a-1022fb9d42ef"; type = "Role" }, # Mail.ReadBasic.All
    @{ id = "246dd0d5-5bd0-4def-940b-0421030a5b68"; type = "Role" }, # Policy.Read.All
    @{ id = "38d9df27-64da-44fd-b7c5-a6fbac20248f"; type = "Role" }, # UserAuthenticationMethod.Read.All
    @{ id = "df021288-bdef-4463-88db-98f22de89214"; type = "Role" }, # User.Read.All
    @{ id = "5b567255-7703-4780-807c-7be8301ae99b"; type = "Role" }, # Group.Read.All
    @{ id = "7438b122-aefc-4978-80ed-43db9fcc7715"; type = "Role" }, # Device.Read.All
    @{ id = "e2a3a72e-5f79-4c64-b1b1-878b674786c9"; type = "Role" }, # Mail.ReadWrite
    @{ id = "483bed4a-2ad3-4361-a73b-c83ccdbdc53c"; type = "Role" }, # RoleManagement.Read.Directory
    @{ id = "bf394140-e372-4bf9-a898-299cfc7564e5"; type = "Role" }, # SecurityEvents.Read.All
    @{ id = "ff278e11-4a33-4d0c-83d2-d01dc58929a5"; type = "Role" }, # RoleEligibilitySchedule.Read.Directory
    @{ id = "d5fe8ce8-684c-4c83-a52c-46e882ce4be1"; type = "Role" }, # RoleAssignmentSchedule.Read.Directory
    @{ id = "40f97065-369a-49f4-947c-6a255697ae91"; type = "Role" }  # MailboxSettings.Read
)

$script:exchangePermissions = @(
    @{ id = "dc50a0fb-09a3-484d-be87-e023b12c6440"; type = "Role" }, # Exchange.ManageAsApp
    @{ id = "e2a3a72e-5f79-4c64-b1b1-878b674786c9"; type = "Role" }  # Mail.ReadWrite
)

Function Set-ApiPermissions {
    param (
        [string]$AppId
    )

    Write-Log "Assigning API permissions to the application..."

    $apiPermissionsBody = @{
        requiredResourceAccess = @(
            @{
                resourceAppId  = $script:graphResourceAppId;
                resourceAccess = $script:graphPermissions
            },
            @{
                resourceAppId  = $script:exchangeResourceAppId;
                resourceAccess = $script:exchangePermissions
            }
        )
    } | ConvertTo-Json -Depth 10

    try {
        Invoke-MgGraphRequest -Method PATCH `
            -Uri "$GraphEndpoint/v1.0/applications/$AppId" `
            -Body $apiPermissionsBody `
            -ContentType "application/json"

        Write-Log "API permissions assigned successfully."
    }
    catch {
        Write-Log "Failed to assign API permissions or update notes: $($_.Exception.Message)" -Level "ERROR"
    }
}

Function Grant-AdminConsent {
    param (
        [string]$ServicePrincipalId
    )

    Write-Log "Attempting to grant admin consent programmatically (no browser handoff)..."

    try {
        $graphLookup = Invoke-MgGraphRequest -Method GET `
            -Uri "$GraphEndpoint/v1.0/servicePrincipals?`$filter=appId eq '$($script:graphResourceAppId)'"
        if (-not $graphLookup.value -or $graphLookup.value.Count -eq 0) {
            Write-Log "Microsoft Graph service principal not found in tenant." -Level "ERROR"
            return $false
        }
        $graphResourceSpId = $graphLookup.value[0].id

        $exoLookup = Invoke-MgGraphRequest -Method GET `
            -Uri "$GraphEndpoint/v1.0/servicePrincipals?`$filter=appId eq '$($script:exchangeResourceAppId)'"
        $exoResourceSpId = if ($exoLookup.value -and $exoLookup.value.Count -gt 0) {
            $exoLookup.value[0].id
        } else {
            Write-Log "Exchange Online service principal not found in tenant - skipping Exchange grants." -Level "WARNING"
            $null
        }
    }
    catch {
        Write-Log "Resource service principal lookup failed: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }

    $granted = 0
    $skipped = 0
    $failed  = 0

    $assignmentTargets = @()
    foreach ($perm in $script:graphPermissions) {
        $assignmentTargets += [pscustomobject]@{ ResourceSpId = $graphResourceSpId; AppRoleId = $perm.id; ResourceLabel = 'Graph' }
    }
    if ($exoResourceSpId) {
        foreach ($perm in $script:exchangePermissions) {
            $assignmentTargets += [pscustomobject]@{ ResourceSpId = $exoResourceSpId; AppRoleId = $perm.id; ResourceLabel = 'Exchange' }
        }
    }

    foreach ($target in $assignmentTargets) {
        $body = @{
            principalId = $ServicePrincipalId
            resourceId  = $target.ResourceSpId
            appRoleId   = $target.AppRoleId
        } | ConvertTo-Json

        try {
            Invoke-MgGraphRequest -Method POST `
                -Uri "$GraphEndpoint/v1.0/servicePrincipals/$ServicePrincipalId/appRoleAssignments" `
                -Body $body -ContentType "application/json" | Out-Null
            $granted++
        }
        catch {
            $msg = $_.Exception.Message
            $detail = if ($_.ErrorDetails) { $_.ErrorDetails.Message } else { '' }
            if ($msg -like '*already exists*' -or $detail -like '*Permission being assigned already exists*' -or $detail -like '*already exists*') {
                $skipped++
            }
            else {
                $failed++
                $errCode = ''
                if ($detail) {
                    try { $errCode = ($detail | ConvertFrom-Json).error.code } catch { }
                }
                $codeSuffix = if ($errCode) { " [$errCode]" } else { '' }
                Write-Log "Failed to grant $($target.ResourceLabel) appRole $($target.AppRoleId)$codeSuffix : $msg" -Level "WARNING"
                if ($detail) { Write-Log "  Detail: $detail" -Level "WARNING" }
            }
        }
    }

    Write-Log "Programmatic consent result: granted=$granted, already-present=$skipped, failed=$failed" -Level "INFO"
    return ($failed -eq 0)
}

Write-Log "Starting the application setup process..." -Level "INFO"

# ----------------------------------------
# Application Lookup / Creation
# ----------------------------------------
Function Get-ApplicationObjectId {
    param (
        [string]$ApplicationId
    )

    try {
        $filter = "`$filter=appId eq '$ApplicationId'"
        $appResult = Invoke-MgGraphRequest -Method GET -Uri "$GraphEndpoint/v1.0/applications?$filter"

        if ($appResult.value -and $appResult.value.Count -gt 0) {
            return $appResult.value[0].id
        }
        else {
            Write-Host "[ERROR] Application with AppId $ApplicationId not found!" -ForegroundColor Red
            return $null
        }
    }
    catch {
        Write-Host "[ERROR] Failed to retrieve application: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

Write-Log "Checking if the application already exists..." -Level "INFO"
$appCreatedInThisSession = $false

Function New-M365Application {
    try {
        $appBody = @{
            displayName    = $appName
            signInAudience = "AzureADMyOrg"
            web            = @{
                redirectUris = @($redirectUri)
            }
        } | ConvertTo-Json -Depth 10

        $newApp = Invoke-MgGraphRequest -Method POST `
            -Uri "$GraphEndpoint/v1.0/applications" `
            -Body $appBody `
            -ContentType "application/json"

        if ($null -ne $newApp -and -not [string]::IsNullOrEmpty($newApp.id)) {
            Write-Log "Application registered successfully!"
            Write-Log "App Name: $($newApp.displayName)"
            Write-Log "App ID: $($newApp.appId)"
            return $newApp
        }
        else {
            Write-Log "Application creation failed! No App ID returned." -Level "ERROR"
            exit
        }
    }
    catch {
        Write-Log "Failed to register application! Error: $($_.Exception.Message)" -Level "ERROR"

        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            Write-Log "Raw ErrorDetails.Message:" -Level "ERROR"
            Write-Log "$($_.ErrorDetails.Message)" -Level "ERROR"

            try {
                $parsedError = $_.ErrorDetails.Message | ConvertFrom-Json
                Write-Log "Parsed Graph API Error Message: $($parsedError.error.message)" -Level "ERROR"
            }
            catch {
                Write-Log "Failed to parse ErrorDetails.Message as JSON." -Level "ERROR"
            }
        }
        else {
            Write-Log "No ErrorDetails available in exception." -Level "ERROR"
        }

        Write-Log "Script Stack Trace: $($_.ScriptStackTrace)" -Level "ERROR"
        exit
    }
}

try {
    Write-Log "Searching for existing application using Graph API..."
    $filter = "`$filter=displayName eq '$appName'"
    $existingApps = Invoke-MgGraphRequest -Method GET -Uri "$GraphEndpoint/v1.0/applications?$filter"

    if ($existingApps.value -and $existingApps.value.Count -gt 0) {
        $appDetails = $existingApps.value[0]
        Write-Log "Application already exists: $($appDetails.displayName) (App ID: $($appDetails.appId))"
        $appCreatedInThisSession = $false
    }
    else {
        Write-Log "Application not found. Registering a new one..." -Level "INFO"
        $appCreatedInThisSession = $true
        $appDetails = New-M365Application
    }
}
catch {
    Write-Log "Failed to search for existing applications: $($_.Exception.Message)" -Level "ERROR"
    Write-Log "Proceeding with application creation..." -Level "INFO"
    $appCreatedInThisSession = $true
    $appDetails = New-M365Application
}

$ApplicationId = $appDetails.appId
$appObjectId = Get-ApplicationObjectId -ApplicationId $ApplicationId

if (-not $appObjectId) {
    Write-Log "Error: Application ObjectId is missing! Cannot proceed." -Level "ERROR"
    exit
}

# ----------------------------------------
# Service Principal Lookup / Creation
# ----------------------------------------
try {
    Write-Log "Searching for existing Service Principal using Graph API..."
    $filter = "`$filter=appId eq '$ApplicationId'"
    $spResult = Invoke-MgGraphRequest -Method GET -Uri "$GraphEndpoint/v1.0/servicePrincipals?$filter"

    if ($spResult.value -and $spResult.value.Count -gt 0) {
        $servicePrincipal = $spResult.value[0]
        Write-Log "Service Principal already exists. ID: $($servicePrincipal.id)"
    }
    else {
        Write-Log "Service Principal not found. Attempting to create it..."
        try {
            $spBody = @{
                appId = $ApplicationId
            } | ConvertTo-Json

            $servicePrincipal = Invoke-MgGraphRequest -Method POST `
                -Uri "$GraphEndpoint/v1.0/servicePrincipals" `
                -Body $spBody `
                -ContentType "application/json"

            Write-Log "Service Principal created successfully. ID: $($servicePrincipal.id)"
        }
        catch {
            Write-Log "Failed to create Service Principal: $($_.Exception.Message)" -Level "ERROR"
            exit
        }
    }
}
catch {
    Write-Log "Failed to retrieve Service Principal: $($_.Exception.Message)" -Level "ERROR"
    exit
}

Write-Log "Service Principal successfully retrieved: $($servicePrincipal.id)"

# ----------------------------------------
# Certificate Creation / Upload
# ----------------------------------------
Write-Log "Generating a self-signed certificate valid for 30 days..."
try {
    $cert = New-SelfSignedCertificate -CertStoreLocation Cert:\CurrentUser\My `
        -DnsName "$appName" `
        -Subject "CN=$appName" `
        -KeyAlgorithm RSA `
        -KeyLength 2048 `
        -KeyExportPolicy Exportable `
        -HashAlgorithm SHA256 `
        -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" `
        -NotAfter (Get-Date).AddDays(30)
    Write-Log "Self-signed certificate created successfully. Thumbprint: $($cert.Thumbprint)"
}
catch {
    Write-Log "Failed to generate self-signed certificate: $($_.Exception.Message)" -Level "ERROR"
    return
}

$certFilename = "$appName-$PrimaryDomain.cer"
$certPath = Join-Path -Path $configDir -ChildPath $certFilename

Write-Log "Exporting the certificate to: $certPath"
try {
    Export-Certificate -Cert $cert -FilePath $certPath
    if (-not (Test-Path -Path $certPath)) {
        Write-Log "Failed to export the .cer file. Exiting..." -Level "ERROR"
        return
    }
    Write-Log "Certificate exported successfully to: $certPath"
}
catch {
    Write-Log "Failed to export the certificate: $($_.Exception.Message)" -Level "ERROR"
    return
}

# Shared helpers for adding the new certificate to the app registration.
# Both rely on script-scope variables set below: $newCertificate, $appDetails, $certPath.
Function Add-CertificateViaAddKey {
    param([string]$ProofToken)

    $addKeyBody = @{
        keyCredential      = $newCertificate
        passwordCredential = $null
    }
    if ($ProofToken) { $addKeyBody.proof = $ProofToken }
    $addKeyBodyJson = $addKeyBody | ConvertTo-Json -Depth 4

    $addKeyResult = Invoke-MgGraphRequest -Method POST -Uri "$GraphEndpoint/v1.0/applications/$($appDetails.Id)/addKey" `
        -Body $addKeyBodyJson `
        -ContentType "application/json"

    Write-Log "New certificate added successfully alongside existing certificates!" -Level "INFO"
    Write-Log "Added certificate key ID: $($addKeyResult.keyId)" -Level "INFO"
}

Function Add-CertificateViaPatchReconstruction {
    try {
        $currentApp = Invoke-MgGraphRequest -Method GET -Uri "$GraphEndpoint/v1.0/applications/$($appDetails.Id)"

        $reconstructedCerts = @()
        $skippedCerts = 0

        foreach ($existingCert in $currentApp.keyCredentials) {
            $localCert = $null
            try {
                $localCerts = Get-ChildItem -Path "Cert:\CurrentUser\My" -ErrorAction SilentlyContinue
                foreach ($c in $localCerts) {
                    $certData = [System.Convert]::ToBase64String($c.RawData)
                    if ($existingCert.key -and $existingCert.key -eq $certData) {
                        $localCert = $c
                        break
                    }

                    if ($existingCert.displayName -and $c.Thumbprint -and $existingCert.displayName.Contains($c.Thumbprint.Substring(0, 8))) {
                        $localCert = $c
                        break
                    }
                }
            }
            catch {
                Write-Log "Error searching local certificates: $($_.Exception.Message)" -Level "INFO"
            }

            if ($localCert) {
                $reconstructedCerts += @{
                    type          = "AsymmetricX509Cert"
                    usage         = "Verify"
                    key           = [System.Convert]::ToBase64String($localCert.RawData)
                    displayName   = $existingCert.displayName
                    startDateTime = $localCert.NotBefore.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
                    endDateTime   = $localCert.NotAfter.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
                }
            }
            elseif ($existingCert.key -and -not [string]::IsNullOrEmpty($existingCert.key)) {
                $preservedCert = @{
                    type        = $existingCert.type
                    usage       = $existingCert.usage
                    key         = $existingCert.key
                    displayName = $existingCert.displayName
                }

                if ($existingCert.startDateTime) {
                    try {
                        $preservedCert.startDateTime = ([DateTime]::Parse($existingCert.startDateTime)).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
                    }
                    catch {
                        $preservedCert.startDateTime = $existingCert.startDateTime
                    }
                }

                if ($existingCert.endDateTime) {
                    try {
                        $preservedCert.endDateTime = ([DateTime]::Parse($existingCert.endDateTime)).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
                    }
                    catch {
                        $preservedCert.endDateTime = $existingCert.endDateTime
                    }
                }

                $reconstructedCerts += $preservedCert
            }
            else {
                Write-Log "Cannot reconstruct certificate '$($existingCert.displayName)' - no key data and not found locally" -Level "WARNING"
                $skippedCerts++
            }
        }

        $reconstructedCerts += $newCertificate

        $successRate = if ($currentApp.keyCredentials.Count -gt 0) {
            ($reconstructedCerts.Count - 1) / $currentApp.keyCredentials.Count
        }
        else {
            1.0
        }

        if ($skippedCerts -eq 0 -or $successRate -ge 0.8) {
            $keyCredentialBody = @{
                keyCredentials = $reconstructedCerts
            } | ConvertTo-Json -Depth 4

            Invoke-MgGraphRequest -Method PATCH -Uri "$GraphEndpoint/v1.0/applications/$($appDetails.Id)" `
                -Body $keyCredentialBody `
                -ContentType "application/json"

            Write-Log "SUCCESS: Certificate added using enhanced PATCH method!" -Level "INFO"
            if ($skippedCerts -gt 0) {
                Write-Log "Note: $skippedCerts certificate(s) could not be preserved and were removed." -Level "WARNING"
            }
        }
        else {
            Write-Log "Reconstruction success rate too low ($([int]($successRate * 100))%). Cannot safely preserve existing certificates." -Level "WARNING"
            Write-Log "IMPORTANT: Automatic certificate addition skipped to preserve existing certificates." -Level "WARNING"
            Write-Log "Certificate file location: $certPath" -Level "INFO"
            Write-Log "" -Level "INFO"
            Write-Log "MANUAL CERTIFICATE ADDITION REQUIRED:" -Level "INFO"
            Write-Log "- Open Azure Portal > Azure Active Directory > App registrations" -Level "INFO"
            Write-Log "- Find '$appName' application" -Level "INFO"
            Write-Log "- Go to 'Certificates & secrets'" -Level "INFO"
            Write-Log "- Upload the certificate file: $certPath" -Level "INFO"
            Write-Log "" -Level "INFO"
            Write-Log "ALTERNATIVE: Use -NewCert parameter to replace all existing certificates" -Level "INFO"
            Write-Log "The application can continue to use existing certificates for now." -Level "INFO"
        }
    }
    catch {
        Write-Log "Enhanced PATCH method failed: $($_.Exception.Message)" -Level "ERROR"
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            Write-Log "PATCH error details: $($_.ErrorDetails.Message)" -Level "ERROR"
        }
        Write-Log "Certificate addition failed. Manual addition required." -Level "WARNING"
        Write-Log "Certificate file location: $certPath" -Level "INFO"
    }
}

Write-Log "Uploading certificate to Azure AD application..."
try {
    $certBytes = [System.IO.File]::ReadAllBytes($certPath)
    $certificateValue = [System.Convert]::ToBase64String($certBytes)

    $certStart = $cert.NotBefore.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    $certEnd = $cert.NotAfter.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

    $newCertificate = @{
        type          = "AsymmetricX509Cert"
        usage         = "Verify"
        key           = $certificateValue
        displayName   = "$appName-Cert"
        startDateTime = $certStart
        endDateTime   = $certEnd
    }

    if ($NewCert) {
        Write-Log "NEW CERT MODE: Replacing existing certificates with new certificate..." -Level "INFO"
        Write-Log "This will overwrite any existing certificates in the Azure AD application." -Level "WARNING"

        $keyCredentialBody = @{
            keyCredentials = @($newCertificate)
        } | ConvertTo-Json -Depth 4

        try {
            Invoke-MgGraphRequest -Method PATCH -Uri "$GraphEndpoint/v1.0/applications/$($appDetails.Id)" `
                -Body $keyCredentialBody `
                -ContentType "application/json"

            Write-Log "Certificate uploaded successfully! Existing certificates have been replaced." -Level "INFO"
            Write-Log "Certificate display name: $($newCertificate.displayName)" -Level "INFO"
        }
        catch {
            Write-Log "Failed to upload certificate: $($_.Exception.Message)" -Level "ERROR"
            if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                Write-Log "Error details: $($_.ErrorDetails.Message)" -Level "ERROR"
            }
            Write-Log "Certificate upload failed. You may need to manually add the certificate." -Level "WARNING"
        }
    }
    else {
        Write-Log "Retrieving existing certificates from application..."
        $existingApp = Invoke-MgGraphRequest -Method GET -Uri "$GraphEndpoint/v1.0/applications/$($appDetails.Id)"

        if ($existingApp.keyCredentials -and $existingApp.keyCredentials.Count -gt 0) {
            Write-Log "Found $($existingApp.keyCredentials.Count) existing certificate(s) in Azure AD application."

            $duplicateFound = $false
            foreach ($existingCert in $existingApp.keyCredentials) {
                if ($existingCert.key -eq $certificateValue) {
                    Write-Log "Certificate with same content already exists. Skipping upload..." -Level "WARNING"
                    $duplicateFound = $true
                    break
                }
            }

            if (-not $duplicateFound) {
                Write-Log "Automatically adding new certificate alongside existing certificates..." -Level "INFO"

                $validExistingCert = $null
                foreach ($existingCert in $existingApp.keyCredentials) {
                    try {
                        $endDate = [DateTime]::Parse($existingCert.endDateTime)
                        if ($endDate -gt (Get-Date)) {
                            $localCert = Get-ChildItem -Path "Cert:\CurrentUser\My" | Where-Object {
                                $_.RawData -and [System.Convert]::ToBase64String($_.RawData) -eq $existingCert.key
                            }
                            if ($localCert -and $localCert.HasPrivateKey) {
                                $validExistingCert = @{
                                    LocalCert = $localCert
                                    KeyId     = $existingCert.keyId
                                }
                                Write-Log "Found valid existing certificate with private key access for proof-of-possession" -Level "INFO"
                                break
                            }
                        }
                    }
                    catch {
                        continue
                    }
                }

                if ($validExistingCert) {
                    try {
                        Write-Log "Generating proof-of-possession token..." -Level "INFO"

                        $jwtHeader = @{
                            alg = "RS256"
                            typ = "JWT"
                            x5t = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($validExistingCert.LocalCert.Thumbprint)) -replace '\+', '-' -replace '/', '_' -replace '='
                        } | ConvertTo-Json -Compress

                        $now = [Math]::Floor((Get-Date -UFormat %s))
                        $jwtPayload = @{
                            aud = "00000002-0000-0000-c000-000000000000"
                            iss = $appDetails.appId
                            nbf = $now
                            exp = $now + 600
                        } | ConvertTo-Json -Compress

                        $jwtHeaderEncoded = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($jwtHeader)) -replace '\+', '-' -replace '/', '_' -replace '='
                        $jwtPayloadEncoded = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($jwtPayload)) -replace '\+', '-' -replace '/', '_' -replace '='

                        $signatureData = "$jwtHeaderEncoded.$jwtPayloadEncoded"
                        $signatureBytes = [System.Text.Encoding]::UTF8.GetBytes($signatureData)

                        $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($validExistingCert.LocalCert)
                        $signature = $rsa.SignData($signatureBytes, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
                        $signatureEncoded = [System.Convert]::ToBase64String($signature) -replace '\+', '-' -replace '/', '_' -replace '='

                        $proofToken = "$jwtHeaderEncoded.$jwtPayloadEncoded.$signatureEncoded"

                        Write-Log "Adding new certificate using addKey method with proof-of-possession..." -Level "INFO"
                        Add-CertificateViaAddKey -ProofToken $proofToken
                    }
                    catch {
                        Write-Log "AddKey method with proof-of-possession failed: $($_.Exception.Message)" -Level "WARNING"
                        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                            Write-Log "AddKey error details: $($_.ErrorDetails.Message)" -Level "WARNING"
                        }

                        try {
                            Write-Log "Attempting addKey method without proof-of-possession..." -Level "INFO"
                            Add-CertificateViaAddKey
                        }
                        catch {
                            Write-Log "AddKey without proof also failed: $($_.Exception.Message)" -Level "WARNING"
                            Write-Log "Using enhanced PATCH method as final attempt..." -Level "INFO"
                            Add-CertificateViaPatchReconstruction
                        }
                    }
                }
                else {
                    Write-Log "No valid existing certificates with private key access found for proof-of-possession." -Level "INFO"
                    Write-Log "Attempting addKey method without proof-of-possession (may work for some tenants)..." -Level "INFO"

                    try {
                        Add-CertificateViaAddKey
                    }
                    catch {
                        Write-Log "AddKey without proof-of-possession failed: $($_.Exception.Message)" -Level "WARNING"
                        Write-Log "Attempting enhanced PATCH method with certificate reconstruction..." -Level "INFO"
                        Add-CertificateViaPatchReconstruction
                    }
                }
            }
        }
        else {
            Write-Log "No existing certificates found. Uploading new certificate..."

            $keyCredentialBody = @{
                keyCredentials = @($newCertificate)
            } | ConvertTo-Json -Depth 4

            Invoke-MgGraphRequest -Method PATCH -Uri "$GraphEndpoint/v1.0/applications/$($appDetails.Id)" `
                -Body $keyCredentialBody `
                -ContentType "application/json"

            Write-Log "Certificate uploaded successfully!" -Level "INFO"
        }
    }

    Write-Log "Local certificate file created at: $certPath" -Level "INFO"
}
catch {
    Write-Log "Failed to upload certificate to Azure AD: $($_.Exception.Message)" -Level "ERROR"
    if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
        Write-Log "Error details: $($_.ErrorDetails.Message)" -Level "ERROR"
    }
    Write-Log "You may need to manually add the certificate to the application." -Level "WARNING"
}

# ----------------------------------------
# Save Configuration (JSON, no client secret)
# ----------------------------------------
$configFilePath = Join-Path -Path $configDir -ChildPath "m365Config.json"
$configContent = @{
    AppId                 = $ApplicationId
    TenantId              = $tenantId
    PrimaryDomain         = $PrimaryDomain
    CertificateThumbprint = $cert.Thumbprint
}

$configJson = $configContent | ConvertTo-Json -Depth 3
Set-Content -Path $configFilePath -Value $configJson -Encoding UTF8
Write-Log "Configuration saved to $configFilePath"

# ----------------------------------------
# Admin consent / API permissions flags
# ----------------------------------------
$needsAdminConsent = $appCreatedInThisSession
$needsApiPermissions = $appCreatedInThisSession

if ($needsAdminConsent) {
    Write-Log "Application was created in this session - admin consent and API permissions required." -Level "INFO"
}
else {
    if ($NewCert) {
        Write-Log "Existing application - running in -NewCert mode. Only rotating certificate; no new admin consent requested." -Level "INFO"
    }
    else {
        Write-Log "Using existing application - admin consent and API permissions not required." -Level "INFO"
    }
}

if ($needsApiPermissions) {
    Set-ApiPermissions -AppId $appDetails.Id
}
else {
    Write-Log "Skipping API permission assignment - using existing application with existing permissions." -Level "INFO"
}

if ($needsAdminConsent) {
    Write-Log "Admin consent is required for API permissions."

    $autoConsentOk = $false
    try {
        $autoConsentOk = Grant-AdminConsent -ServicePrincipalId $servicePrincipal.id
    }
    catch {
        Write-Log "Programmatic consent threw: $($_.Exception.Message)" -Level "WARNING"
        $autoConsentOk = $false
    }

    if ($autoConsentOk) {
        Write-Log "Admin consent granted programmatically. No browser sign-in needed." -Level "INFO"
        Write-Log "Waiting 30 seconds for Azure propagation..."
        Start-Sleep -Seconds 30
    }
    else {
        Write-Log "Programmatic admin consent failed (likely insufficient privileges or tenant policy). Falling back to manual consent URL." -Level "WARNING"
        $adminConsentUrl = "$LoginEndpoint/$tenantId/adminconsent?client_id=$($appDetails.AppId)"

        Write-Host "============================================================"
        Write-Host "ADMIN CONSENT REQUIRED (manual fallback)"
        Write-Host "============================================================"
        Write-Host "Open the following URL in your browser and grant admin consent:"
        Write-Host "$adminConsentUrl" -ForegroundColor Yellow
        Write-Host "Once admin consent is granted, press [Enter] to continue..."
        Read-Host

        Write-Log "Waiting 30 seconds for Azure propagation..."
        Start-Sleep -Seconds 30
    }
}
else {
    Write-Log "Skipping admin consent (existing application path)." -Level "INFO"
    Write-Log "Waiting 10 seconds for permission updates to propagate..." -Level "INFO"
    Start-Sleep -Seconds 10
}

# ----------------------------------------
# Exchange Administrator Role (new apps only)
# ----------------------------------------
if ($appCreatedInThisSession) {
    Write-Log "Application was created in this session - checking Exchange Administrator role assignment..." -Level "INFO"

    Write-Log "Checking if Exchange Administrator role is activated..."
    $roleTemplateId = "29232cdf-9323-42fd-ade2-1d097af3e4de"

    try {
        $filter = "`$filter=roleTemplateId eq '$roleTemplateId'"
        $roleResult = Invoke-MgGraphRequest -Method GET -Uri "$GraphEndpoint/v1.0/directoryRoles?$filter"

        if ($roleResult.value -and $roleResult.value.Count -gt 0) {
            $exchangeAdminRole = $roleResult.value[0]
            Write-Log "Exchange Administrator role is already activated. Role ID: $($exchangeAdminRole.id)"
        }
        else {
            Write-Log "Activating Exchange Administrator role..."
            $roleBody = @{
                roleTemplateId = $roleTemplateId
            } | ConvertTo-Json

            $exchangeAdminRole = Invoke-MgGraphRequest -Method POST `
                -Uri "$GraphEndpoint/v1.0/directoryRoles" `
                -Body $roleBody `
                -ContentType "application/json"

            Write-Log "Exchange Administrator role activated. Role ID: $($exchangeAdminRole.id)"
            Start-Sleep -Seconds 30
        }
    }
    catch {
        Write-Log "Failed to activate Exchange Administrator role: $($_.Exception.Message)" -Level "ERROR"
        Write-Log "Continuing with script - Exchange role assignment may need manual setup." -Level "WARNING"
        $exchangeAdminRole = $null
    }

    if ($exchangeAdminRole) {
        Write-Log "Exchange Administrator Role ID: $($exchangeAdminRole.id)"

        Write-Log "Checking if Exchange Administrator role is already assigned to Service Principal..."
        try {
            $membersResult = Invoke-MgGraphRequest -Method GET -Uri "$GraphEndpoint/v1.0/directoryRoles/$($exchangeAdminRole.id)/members"
            $existingAssignment = $membersResult.value | Where-Object { $_.id -eq $servicePrincipal.id }

            if ($existingAssignment) {
                Write-Log "Exchange Administrator role is already assigned." -Level "INFO"
            }
            else {
                Write-Log "Assigning Exchange Administrator role..."

                $body = @{
                    "@odata.id" = "$GraphEndpoint/v1.0/directoryObjects/$($servicePrincipal.id)"
                } | ConvertTo-Json

                try {
                    Invoke-MgGraphRequest -Method POST `
                        -Uri "$GraphEndpoint/v1.0/directoryRoles/$($exchangeAdminRole.id)/members/`$ref" `
                        -Body $body `
                        -ContentType "application/json"

                    Write-Log "Exchange Administrator role successfully assigned." -Level "INFO"
                    Start-Sleep -Seconds 30
                }
                catch {
                    $originalError = $_.Exception.Message

                    if ($originalError -like "*already exists*" -or $originalError -like "*BadRequest*" -or $_.ErrorDetails.Message -like "*already exists*") {
                        Write-Log "Role assignment failed with BadRequest - likely already assigned. Verifying..." -Level "WARNING"

                        try {
                            $recheckResult = Invoke-MgGraphRequest -Method GET -Uri "$GraphEndpoint/v1.0/directoryRoles/$($exchangeAdminRole.id)/members"
                            $recheckAssignment = $recheckResult.value | Where-Object { $_.id -eq $servicePrincipal.id }

                            if ($recheckAssignment) {
                                Write-Log "Confirmed: Exchange Administrator role is already assigned to service principal." -Level "INFO"
                                Write-Log "Role assignment verification successful - continuing with script." -Level "INFO"
                            }
                            else {
                                Write-Log "Role assignment verification failed - service principal not found in role members." -Level "WARNING"
                                Write-Log "Original error: $originalError" -Level "ERROR"
                                Write-Log "Please manually verify Exchange Administrator role assignment in Azure AD." -Level "ERROR"
                            }
                        }
                        catch {
                            Write-Log "Failed to verify role assignment: $($_.Exception.Message)" -Level "WARNING"
                            Write-Log "Assuming role is assigned and continuing with script execution..." -Level "INFO"
                        }
                    }
                    else {
                        Write-Log "Failed to assign Exchange Administrator role: $originalError" -Level "ERROR"
                        Write-Log "Error Details: Please manually assign the Exchange Administrator role to the service principal in Azure AD." -Level "ERROR"
                    }
                }
            }
        }
        catch {
            Write-Log "Failed to check role assignment: $($_.Exception.Message)" -Level "ERROR"
            Write-Log "Continuing with script execution. Exchange Administrator role may need to be assigned manually." -Level "WARNING"
        }
    }
}
else {
    Write-Log "Using existing application - skipping Exchange Administrator role assignment." -Level "INFO"
}

# ----------------------------------------
# Load Config Back In (for command generation)
# ----------------------------------------
if (-not (Test-Path -Path $configFilePath)) {
    Write-Log "Error: Configuration file not found at $configFilePath. Exiting..." -Level "ERROR"
    exit
}

try {
    $configContent = Get-Content -Path $configFilePath -Raw | ConvertFrom-Json
}
catch {
    Write-Log "Error: Failed to load or parse configuration file. Ensure it's valid JSON. Error: $($_.Exception.Message)" -Level "ERROR"
    exit
}

if (-not $configContent.PSObject.Properties['AppId'] -or
    -not $configContent.PSObject.Properties['TenantId'] -or
    -not $configContent.PSObject.Properties['CertificateThumbprint'] -or
    -not $configContent.PSObject.Properties['PrimaryDomain']) {
    Write-Log "Error: Missing required properties in the configuration file. Exiting..." -Level "ERROR"
    exit
}

$ApplicationId = $configContent.AppId
$TenantId = $configContent.TenantId
$Thumbprint = $configContent.CertificateThumbprint
$PrimaryDomain = $configContent.PrimaryDomain

# ----------------------------------------
# Command Output Helpers
# ----------------------------------------
$commandLogFile = Join-Path -Path $outputDir -ChildPath "log.txt"
if (-not (Test-Path -Path $commandLogFile)) {
    Write-Host "[INFO] Creating log file: $commandLogFile"
    New-Item -ItemType File -Path $commandLogFile -Force | Out-Null
}

$commandOutputFilePath = Join-Path -Path $outputDir -ChildPath "command_output.txt"
$logOutputFilePath = Join-Path -Path $outputDir -ChildPath "log_output.txt"

Function Write-CommandOutput {
    param (
        [string]$Message,
        [string]$ForegroundColor = "White"
    )

    if (-not (Test-Path -Path $commandOutputFilePath)) {
        New-Item -ItemType File -Path $commandOutputFilePath -Force | Out-Null
    }

    # Console output: NO timestamps, just the message
    if ($Message -ne "") {
        Write-Host ""
        Write-Host $Message -ForegroundColor $ForegroundColor
        Write-Host ""
    }

    # File output: same as console, no timestamps
    try {
        Add-Content -Path $commandOutputFilePath -Value ""
        Add-Content -Path $commandOutputFilePath -Value $Message
        Add-Content -Path $commandOutputFilePath -Value ""
    }
    catch {
        Write-Host "[ERROR] Failed to write command to command output file: $($_.Exception.Message)" -ForegroundColor Red
    }

    # NOTE: intentionally NOT calling Write-Log here
}


# ----------------------------------------
# Disconnect from MgGraph
# ----------------------------------------
$StayConnected = $false

if (-not $StayConnected) {
    try {
        Disconnect-MgGraph | Out-Null
        Write-Host "[INFO] Disconnected from Microsoft Graph."
    }
    catch {
        Write-Host "[WARNING] Disconnect-MgGraph threw an error (continuing anyway): $($_.Exception.Message)" -ForegroundColor Yellow
    }
}
else {
    Write-Host "[INFO] The connection to Microsoft Graph is still active." -ForegroundColor Yellow
    Write-Host "To manually disconnect, run: Disconnect-MgGraph"
}
'@

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# ----------------------------------------
# XAML
# ----------------------------------------
[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="M.E.G. - Microsoft Extractor GUI"
        Height="940"
        Width="1120"
        MinHeight="640"
        MinWidth="900"
        WindowStartupLocation="CenterScreen">

    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Text="M.E.G. - Microsoft Extractor GUI"
                   FontSize="18"
                   FontWeight="SemiBold"
                   Margin="0,0,0,10"/>

        <Grid Grid.Row="1" Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="Tenant:" VerticalAlignment="Center" Margin="0,0,10,0"/>
            <ComboBox Name="cmbTenant" Grid.Column="1" Height="26"/>
            <StackPanel Grid.Column="2" Orientation="Horizontal" Margin="10,0,0,0">
                <Button Name="btnRefresh"   Content="Refresh"          Width="90"  Margin="0,0,6,0"/>
                <Button Name="btnTestUal"   Content="Test UAL Status"  Width="130" Margin="0,0,6,0"/>
                <Button Name="btnUninstall" Content="Uninstall Tenant" Width="140"/>
            </StackPanel>
        </Grid>

        <Grid Grid.Row="2" Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="180"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="140"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <TextBlock Text="Primary Domain:" Grid.Row="0" Grid.Column="0"/>
            <TextBox Name="txtDomain" Grid.Row="0" Grid.Column="1" IsReadOnly="True"/>

            <TextBlock Text="Output Directory:" Grid.Row="0" Grid.Column="2" Margin="10,0,0,0"/>
            <TextBox Name="txtOutputDir" Grid.Row="0" Grid.Column="3"/>

            <TextBlock Text="Tenant ID:" Grid.Row="1" Grid.Column="0" Margin="0,6,0,0"/>
            <TextBox Name="txtTenantId" Grid.Row="1" Grid.Column="1" IsReadOnly="True" Margin="0,6,0,0"/>

            <TextBlock Text="App ID:" Grid.Row="2" Grid.Column="0" Margin="0,6,0,0"/>
            <TextBox Name="txtAppId" Grid.Row="2" Grid.Column="1" IsReadOnly="True" Margin="0,6,0,0"/>

            <TextBlock Text="Certificate Thumbprint:" Grid.Row="3" Grid.Column="0" Margin="0,6,0,0"/>
            <TextBox Name="txtThumbprint" Grid.Row="3" Grid.Column="1" IsReadOnly="True" Margin="0,6,0,0"/>
        </Grid>

        <GroupBox Grid.Row="3" Header="Validation Results" Margin="0,0,0,10">
            <StackPanel Margin="6">
                <TextBlock Text="Errors:" FontWeight="SemiBold"/>
                <TextBox Name="txtErrors" Foreground="DarkRed" IsReadOnly="True" Height="50" VerticalScrollBarVisibility="Auto" TextWrapping="Wrap"/>
                <TextBlock Text="Warnings:" FontWeight="SemiBold" Margin="0,6,0,0"/>
                <TextBox Name="txtWarnings" Foreground="DarkGoldenrod" IsReadOnly="True" Height="50" VerticalScrollBarVisibility="Auto" TextWrapping="Wrap"/>
            </StackPanel>
        </GroupBox>

        <TabControl Grid.Row="4" Name="tabMain">
            <TabItem Header="M365 Command Builder">
                <Grid Margin="0,8,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="440"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <GroupBox Header="Log Collection Selection (Categorized)" Grid.Column="0" Margin="0,0,10,0">
                        <DockPanel>
                            <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="6,6,6,4">
                                <Button Name="btnSelectAll" Content="Select Standard" Width="120" Margin="0,0,6,0"
                                        ToolTip="Selects all broad-collection cmdlets. Skips targeted/per-user items (Show-MailboxRules, MailboxAuditLog, Get-UAL classic, Email/Attachment/Show-Email, Sessions/MessageIDs, Show-TransportRules, MessageTrace) - tick those individually."/>
                                <Button Name="btnSelectNone" Content="Select None" Width="100"/>
                            </StackPanel>
                            <ScrollViewer VerticalScrollBarVisibility="Auto">
                                <StackPanel Margin="6">
                                    <TextBlock Text="EXO - Mailbox" FontWeight="SemiBold" Margin="0,4,0,4"/>
                                    <CheckBox Name="cbMailboxAuditStatus" Content="MailboxAuditStatus (Get-MailboxAuditStatus)"/>
                                    <CheckBox Name="cbMailboxPermissions" Content="MailboxPermissions (Get-MailboxPermissions)"/>
                                    <CheckBox Name="cbMailboxRules"       Content="MailboxRules (Get-MailboxRules)"/>
                                    <CheckBox Name="cbShowMailboxRules"   Content="Show MailboxRules (Show-MailboxRules)"/>

                                    <TextBlock Text="EXO - Audit Logs" FontWeight="SemiBold" Margin="0,8,0,4"/>
                                    <CheckBox Name="cbAdminAuditLog"      Content="AdminAuditLog (Get-AdminAuditLog)"/>
                                    <CheckBox Name="cbMailboxAuditLog"    Content="MailboxAuditLog (Get-MailboxAuditLog)"/>
                                    <CheckBox Name="cbUALClassic"         Content="Unified Audit Log via EXO (Get-UAL)"/>

                                    <TextBlock Text="EXO - Email" FontWeight="SemiBold" Margin="0,8,0,4"/>
                                    <CheckBox Name="cbEmail"              Content="Email (Get-Email)"/>
                                    <CheckBox Name="cbAttachment"         Content="Attachment (Get-Attachment)"/>
                                    <CheckBox Name="cbShowEmail"          Content="Show Email (Show-Email)"/>

                                    <TextBlock Text="EXO - Mail Items Accessed" FontWeight="SemiBold" Margin="0,8,0,4"/>
                                    <CheckBox Name="cbSessions"           Content="Sessions (Get-Sessions)"/>
                                    <CheckBox Name="cbMessageIDs"         Content="MessageIDs (Get-MessageIDs)"/>

                                    <TextBlock Text="EXO - Transport" FontWeight="SemiBold" Margin="0,8,0,4"/>
                                    <CheckBox Name="cbTransportRules"     Content="TransportRules (Get-TransportRules)"/>
                                    <CheckBox Name="cbShowTransportRules" Content="Show TransportRules (Show-TransportRules)"/>
                                    <CheckBox Name="cbMessageTrace"       Content="90 Day MessageTraceLog (Get-MessageTraceLog)"/>

                                    <Separator Margin="0,8,0,8"/>

                                    <TextBlock Text="Graph - Identity" FontWeight="SemiBold" Margin="0,4,0,4"/>
                                    <CheckBox Name="cbUsers"              Content="Users (Get-Users)"/>
                                    <CheckBox Name="cbAdminUsers"         Content="AdminUsers (Get-AdminUsers)"/>
                                    <CheckBox Name="cbMFA"                Content="MFA (Get-MFA)"/>
                                    <CheckBox Name="cbRisky"              Content="RiskyDetections (Get-RiskyDetections)"/>
                                    <CheckBox Name="cbRiskyUsers"         Content="RiskyUsers (Get-RiskyUsers)"/>
                                    <CheckBox Name="cbAuth"               Content="OAuth Permissions (Get-OAuthPermissionsGraph)"/>

                                    <TextBlock Text="Graph - Devices &amp; Groups" FontWeight="SemiBold" Margin="0,8,0,4"/>
                                    <CheckBox Name="cbDevices"            Content="Devices (Get-Devices)"/>
                                    <CheckBox Name="cbGroups"             Content="Groups (Get-Groups)"/>
                                    <CheckBox Name="cbGroupMembers"       Content="GroupMembers (Get-GroupMembers)"/>
                                    <CheckBox Name="cbDynamicGroups"      Content="DynamicGroups (Get-DynamicGroups)"/>

                                    <TextBlock Text="Graph - Policies &amp; Posture" FontWeight="SemiBold" Margin="0,8,0,4"/>
                                    <CheckBox Name="cbConditionalAccess"  Content="ConditionalAccessPolicies (Get-ConditionalAccessPolicies)"/>
                                    <CheckBox Name="cbSecurityDefaults"   Content="EntraSecurityDefaults (Get-EntraSecurityDefaults)"/>
                                    <CheckBox Name="cbSecureScore"        Content="SecureScore (Get-SecureScore)"/>
                                    <CheckBox Name="cbSecurityAlerts"     Content="SecurityAlerts (Get-SecurityAlerts)"/>

                                    <TextBlock Text="Graph - Roles &amp; Licenses" FontWeight="SemiBold" Margin="0,8,0,4"/>
                                    <CheckBox Name="cbAllRoleActivity"    Content="AllRoleActivity (Get-AllRoleActivity)"/>
                                    <CheckBox Name="cbPIMAssignments"     Content="PIMAssignments (Get-PIMAssignments)"/>
                                    <CheckBox Name="cbLicenses"           Content="LicensesByUser (Get-LicensesByUser)"/>
                                    <CheckBox Name="cbAllLicenses"        Content="Licenses (Get-Licenses)"/>
                                    <CheckBox Name="cbProductLicenses"    Content="ProductLicenses (Get-ProductLicenses)"/>

                                    <TextBlock Text="Graph - Logs" FontWeight="SemiBold" Margin="0,8,0,4"/>
                                    <CheckBox Name="cbGraphAudit"         Content="GraphEntraAuditLogs (Get-GraphEntraAuditLogs)"/>
                                    <CheckBox Name="cbGraphSignin"        Content="GraphEntraSignInLogs (Get-GraphEntraSignInLogs)"/>
                                    <CheckBox Name="cbMailboxRulesGraph"  Content="MailboxRules via Graph (Get-MailboxRulesGraph)"/>

                                    <Separator Margin="0,8,0,8"/>

                                    <TextBlock Text="Unified Audit Log (UAL via Graph)" FontWeight="SemiBold" Margin="0,4,0,4"/>
                                    <Grid Margin="0,2,0,2">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="110"/>
                                            <ColumnDefinition Width="140"/>
                                            <ColumnDefinition Width="*"/>
                                        </Grid.ColumnDefinitions>
                                        <Grid.RowDefinitions>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                        </Grid.RowDefinitions>

                                        <TextBlock Text="Last N days:"     Grid.Row="0" Grid.Column="0" VerticalAlignment="Center" Margin="0,2,4,2"/>
                                        <TextBox   Name="txtUalDaysBack"   Grid.Row="0" Grid.Column="1" Margin="0,2,0,2"
                                                   ToolTip="Pull UAL for the last N days from now. Ignored when Start date is set."/>

                                        <TextBlock Text="Start date:"      Grid.Row="1" Grid.Column="0" VerticalAlignment="Center" Margin="0,2,4,2"/>
                                        <DatePicker Name="dpUalStart"      Grid.Row="1" Grid.Column="1" Margin="0,2,0,2"/>

                                        <TextBlock Text="End date:"        Grid.Row="2" Grid.Column="0" VerticalAlignment="Center" Margin="0,2,4,2"/>
                                        <DatePicker Name="dpUalEnd"        Grid.Row="2" Grid.Column="1" Margin="0,2,0,2"
                                                    ToolTip="Optional. If left blank but Start is set, the window runs through now."/>
                                    </Grid>
                                    <TextBlock Margin="0,2,0,4" TextWrapping="Wrap" FontStyle="Italic" Foreground="#666"
                                               Text="Precedence: Pull all (below) &gt; Start/End &gt; Last N days. Leave all blank to skip UAL."/>
                                    <CheckBox Name="cbUalAll" Content="Pull all UAL" Margin="0,4,0,0"
                                              ToolTip="Pulls the maximum window most tenants retain (365 days). Tenants on Audit Standard will get back what's actually retained (typically 180 days). For longer windows on Audit Premium with retention add-ons, use the Start/End date pickers instead."/>
                                    <CheckBox Name="cbUALTriage" Content="Triage mode (filter by operations list)" Margin="0,4,0,0"
                                              ToolTip="When checked, UAL pulls are filtered to the triage operations list. Uncheck to pull the entire UAL for the window."/>
                                </StackPanel>
                            </ScrollViewer>
                        </DockPanel>
                    </GroupBox>

                    <GroupBox Grid.Column="1" Header="Generated Commands">
                        <TextBox Name="txtCommands"
                                 Margin="6"
                                 AcceptsReturn="True"
                                 VerticalScrollBarVisibility="Auto"
                                 HorizontalScrollBarVisibility="Auto"
                                 TextWrapping="NoWrap"
                                 IsReadOnly="True"
                                 FontFamily="Consolas"/>
                    </GroupBox>
                </Grid>
            </TabItem>

            <TabItem Header="Global Options">
                <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="0,8,0,0">
                    <Grid Margin="10">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="180"/>
                            <ColumnDefinition Width="260"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>

                        <TextBlock Text="-UserIds:"           Grid.Row="0" Grid.Column="0" Margin="0,4,0,4"/>
                        <TextBox   Name="txtGlobalUserId"     Grid.Row="0" Grid.Column="1" Margin="0,2,0,2"
                                   ToolTip="One or more users (email/UPN/ID), separated by commas, semicolons, or spaces. Quoting and formatting are handled automatically per cmdlet; single-user cmdlets (Get-Users, Get-Email, etc.) get one command per user."/>

                        <TextBlock Text="-Output:"            Grid.Row="1" Grid.Column="0" Margin="0,4,0,4"/>
                        <ComboBox  Name="cmbGlobalOutput"     Grid.Row="1" Grid.Column="1" Margin="0,2,0,2">
                            <ComboBoxItem Content=""/>
                            <ComboBoxItem Content="CSV"/>
                            <ComboBoxItem Content="JSON"/>
                            <ComboBoxItem Content="JSONL"/>
                            <ComboBoxItem Content="SOF-ELK"/>
                        </ComboBox>

                        <TextBlock Text="-Encoding:"          Grid.Row="2" Grid.Column="0" Margin="0,4,0,4"/>
                        <ComboBox  Name="cmbGlobalEncoding"   Grid.Row="2" Grid.Column="1" Margin="0,2,0,2" IsEditable="True">
                            <ComboBoxItem Content=""/>
                            <ComboBoxItem Content="ASCII"/>
                            <ComboBoxItem Content="BigEndianUnicode"/>
                            <ComboBoxItem Content="Default"/>
                            <ComboBoxItem Content="OEM"/>
                            <ComboBoxItem Content="Unicode"/>
                            <ComboBoxItem Content="UTF7"/>
                            <ComboBoxItem Content="UTF8"/>
                            <ComboBoxItem Content="UTF32"/>
                        </ComboBox>

                        <TextBlock Text="-LogLevel:"          Grid.Row="3" Grid.Column="0" Margin="0,4,0,4"/>
                        <ComboBox  Name="cmbGlobalLogLevel"   Grid.Row="3" Grid.Column="1" Margin="0,2,0,2">
                            <ComboBoxItem Content=""/>
                            <ComboBoxItem Content="None"/>
                            <ComboBoxItem Content="Minimal"/>
                            <ComboBoxItem Content="Standard"/>
                            <ComboBoxItem Content="Debug"/>
                        </ComboBox>

                        <CheckBox  Name="cbGlobalMergeOutput" Grid.Row="4" Grid.Column="0" Grid.ColumnSpan="2" Margin="0,8,0,4" Content="-MergeOutput"/>

                        <TextBlock Grid.Row="6" Grid.Column="0" Grid.ColumnSpan="3" Margin="0,12,0,0" TextWrapping="Wrap"
                                   Text="Global switches are appended to every emitted cmdlet that supports them. Blank ComboBoxes are omitted. -UserIds, -Output, and -MergeOutput only apply to cmdlets that document them; the rest use their defaults."/>
                    </Grid>
                </ScrollViewer>
            </TabItem>

            <TabItem Header="Per-Cmdlet Options">
                <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="0,8,0,0">
                    <StackPanel Margin="10">

                        <Expander Header="Get-UALGraph" FontWeight="SemiBold" Margin="0,0,0,6">
                            <Grid Margin="10,8,10,10">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="180"/>
                                    <ColumnDefinition Width="260"/>
                                    <ColumnDefinition Width="20"/>
                                    <ColumnDefinition Width="180"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>

                                <TextBlock Text="-IpAddress:"        Grid.Row="0" Grid.Column="0" Margin="0,4,0,4" FontWeight="Normal"/>
                                <TextBox   Name="txtUalIp"           Grid.Row="0" Grid.Column="1" Margin="0,2,0,2"
                                           ToolTip="One or more IPs, separated by commas or spaces - emitted as a PowerShell array."/>

                                <TextBlock Text="-Service:"          Grid.Row="1" Grid.Column="0" Margin="0,4,0,4" FontWeight="Normal"/>
                                <TextBox   Name="txtUalService"      Grid.Row="1" Grid.Column="1" Margin="0,2,0,2"/>

                                <TextBlock Text="-Keyword:"          Grid.Row="2" Grid.Column="0" Margin="0,4,0,4" FontWeight="Normal"/>
                                <TextBox   Name="txtUalKeyword"      Grid.Row="2" Grid.Column="1" Margin="0,2,0,2"/>

                                <TextBlock Text="-RecordType:"       Grid.Row="3" Grid.Column="0" Margin="0,4,0,4" FontWeight="Normal"/>
                                <TextBox   Name="txtUalRecordType"   Grid.Row="3" Grid.Column="1" Margin="0,2,0,2"/>

                                <TextBlock Text="-ObjectID:"         Grid.Row="4" Grid.Column="0" Margin="0,4,0,4" FontWeight="Normal"/>
                                <TextBox   Name="txtUalObjectId"     Grid.Row="4" Grid.Column="1" Margin="0,2,0,2"/>

                                <TextBlock Text="-Operations:"       Grid.Row="5" Grid.Column="0" Margin="0,4,0,4" FontWeight="Normal"/>
                                <TextBox   Name="txtUalOperations"   Grid.Row="5" Grid.Column="1" Margin="0,2,0,2"
                                           ToolTip="Comma-separated operation names (spaces inside names are fine). Emitted as a PowerShell array. Overrides Triage mode when set."/>

                                <TextBlock Text="-MaxEventsPerFile:" Grid.Row="0" Grid.Column="3" Margin="10,4,0,4" FontWeight="Normal"/>
                                <TextBox   Name="txtUalMaxEvents"    Grid.Row="0" Grid.Column="4" Margin="0,2,0,2" Text="50000"/>

                                <CheckBox  Name="cbUalSplitFiles"    Grid.Row="1" Grid.Column="3" Grid.ColumnSpan="2" Margin="10,8,0,4" Content="-SplitFiles" IsChecked="True" FontWeight="Normal"/>
                                <CheckBox  Name="cbUalUseV1"         Grid.Row="2" Grid.Column="3" Grid.ColumnSpan="2" Margin="10,4,0,4" Content="-UseV1" FontWeight="Normal"/>

                                <TextBlock Grid.Row="6" Grid.Column="0" Grid.ColumnSpan="5" Margin="0,12,0,0" TextWrapping="Wrap" FontWeight="Normal"
                                           Text="-Output, -Encoding, and -LogLevel live on the Global Options tab. If -Operations is set here it overrides Triage mode; otherwise Triage mode (Command Builder tab) controls whether the built-in operations list is sent."/>
                            </Grid>
                        </Expander>

                        <Expander Header="Get-UAL (EXO classic)" FontWeight="SemiBold" Margin="0,0,0,6">
                            <Grid Name="gridUalClassicOpts" Margin="10,8,10,10">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="200"/>
                                    <ColumnDefinition Width="280"/>
                                    <ColumnDefinition Width="20"/>
                                    <ColumnDefinition Width="200"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>

                                <TextBlock Text="-Group:"                Grid.Row="0" Grid.Column="0" Margin="0,4,0,4" FontWeight="Normal"/>
                                <ComboBox  Name="cmbUalClassicGroup"     Grid.Row="0" Grid.Column="1" Margin="0,2,0,2">
                                    <ComboBoxItem Content=""/>
                                    <ComboBoxItem Content="Exchange"/>
                                    <ComboBoxItem Content="Azure"/>
                                    <ComboBoxItem Content="SharePoint"/>
                                    <ComboBoxItem Content="Skype"/>
                                    <ComboBoxItem Content="Defender"/>
                                </ComboBox>

                                <TextBlock Text="-RecordType:"           Grid.Row="1" Grid.Column="0" Margin="0,4,0,4" FontWeight="Normal"/>
                                <TextBox   Name="txtUalClassicRecordType" Grid.Row="1" Grid.Column="1" Margin="0,2,0,2"/>

                                <TextBlock Text="-Operation:"            Grid.Row="2" Grid.Column="0" Margin="0,4,0,4" FontWeight="Normal"/>
                                <TextBox   Name="txtUalClassicOperation" Grid.Row="2" Grid.Column="1" Margin="0,2,0,2"/>

                                <TextBlock Text="-IPAddresses:"          Grid.Row="3" Grid.Column="0" Margin="0,4,0,4" FontWeight="Normal"/>
                                <TextBox   Name="txtUalClassicIp"        Grid.Row="3" Grid.Column="1" Margin="0,2,0,2"
                                           ToolTip="Comma-separated client IPs"/>

                                <TextBlock Text="-ObjectIDs:"            Grid.Row="4" Grid.Column="0" Margin="0,4,0,4" FontWeight="Normal"/>
                                <TextBox   Name="txtUalClassicObjectIds" Grid.Row="4" Grid.Column="1" Margin="0,2,0,2"/>

                                <TextBlock Text="-Interval:"             Grid.Row="0" Grid.Column="3" Margin="10,4,0,4" FontWeight="Normal"/>
                                <TextBox   Name="txtUalClassicInterval"  Grid.Row="0" Grid.Column="4" Margin="0,2,0,2"
                                           ToolTip="TimeSpan, e.g. 1.00:00:00"/>

                                <TextBlock Text="-TargetEventsPerWindow:" Grid.Row="1" Grid.Column="3" Margin="10,4,0,4" FontWeight="Normal"/>
                                <TextBox   Name="txtUalClassicTargetEvents" Grid.Row="1" Grid.Column="4" Margin="0,2,0,2"
                                           ToolTip="Default 3000, range 1-5000"/>

                                <CheckBox  Name="cbUalClassicAuditDataOnly" Grid.Row="2" Grid.Column="3" Grid.ColumnSpan="2" Margin="10,8,0,4" Content="-AuditDataOnly" FontWeight="Normal"/>

                                <TextBlock Grid.Row="7" Grid.Column="0" Grid.ColumnSpan="5" Margin="0,12,0,0" TextWrapping="Wrap" FontWeight="Normal"
                                           Text="Applies only when the Get-UAL checkbox is selected. -UserIds, -StartDate, -EndDate, -Output, -Encoding, -LogLevel come from elsewhere."/>
                            </Grid>
                        </Expander>

                        <Expander Header="Email  -  Get-Email / Get-Attachment / Show-Email" FontWeight="SemiBold" Margin="0,0,0,6">
                            <Grid Name="gridEmailOpts" Margin="10,8,10,10">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="200"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>

                                <TextBlock Text="-InternetMessageId:"     Grid.Row="0" Grid.Column="0" Margin="0,4,0,4" FontWeight="Normal"/>
                                <TextBox   Name="txtEmailInternetMessageId" Grid.Row="0" Grid.Column="1" Margin="0,2,0,2"
                                           ToolTip="Required by Get-Email, Get-Attachment, Show-Email"/>

                                <TextBlock Text="-inputFile (Get-Email):" Grid.Row="1" Grid.Column="0" Margin="0,4,0,4" FontWeight="Normal"/>
                                <TextBox   Name="txtEmailInputFile"       Grid.Row="1" Grid.Column="1" Margin="0,2,0,2"
                                           ToolTip="Path to .txt with multiple InternetMessageIds"/>

                                <CheckBox  Name="cbEmailAttachment"        Grid.Row="2" Grid.Column="0" Grid.ColumnSpan="2" Margin="0,8,0,4" Content="-Attachment (Get-Email)" FontWeight="Normal"/>
                                <CheckBox  Name="cbEmailDownloadDuplicates" Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="2" Margin="0,4,0,4" Content="-DownloadDuplicates (Get-Email)" FontWeight="Normal"/>

                                <TextBlock Grid.Row="4" Grid.Column="0" Grid.ColumnSpan="2" Margin="0,12,0,0" TextWrapping="Wrap" FontWeight="Normal"
                                           Text="-InternetMessageId is mandatory for Get-Email, Get-Attachment, and Show-Email. If blank, those cmdlets are skipped and a warning shows."/>
                            </Grid>
                        </Expander>

                        <Expander Header="Mail Items Accessed  -  Get-Sessions / Get-MessageIDs" FontWeight="SemiBold" Margin="0,0,0,6">
                            <Grid Name="gridMailItemsOpts" Margin="10,8,10,10">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="200"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>

                                <TextBlock Text="-IP:"                Grid.Row="0" Grid.Column="0" Margin="0,4,0,4" FontWeight="Normal"/>
                                <TextBox   Name="txtMailItemsIp"       Grid.Row="0" Grid.Column="1" Margin="0,2,0,2"
                                           ToolTip="Get-Sessions and Get-MessageIDs. SINGLE IP only - the module compares it exactly against ClientIPAddress; if multiple are entered only the first is used."/>

                                <TextBlock Text="-Sessions:"          Grid.Row="1" Grid.Column="0" Margin="0,4,0,4" FontWeight="Normal"/>
                                <TextBox   Name="txtMailItemsSessions" Grid.Row="1" Grid.Column="1" Margin="0,2,0,2"
                                           ToolTip="Get-MessageIDs only"/>

                                <TextBlock Text="-Download:"          Grid.Row="2" Grid.Column="0" Margin="0,4,0,4" FontWeight="Normal"/>
                                <ComboBox  Name="cmbMailItemsDownload" Grid.Row="2" Grid.Column="1" Margin="0,2,0,2">
                                    <ComboBoxItem Content=""/>
                                    <ComboBoxItem Content="Yes"/>
                                    <ComboBoxItem Content="No"/>
                                </ComboBox>

                                <TextBlock Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="2" Margin="0,12,0,0" TextWrapping="Wrap" FontWeight="Normal"
                                           Text="Switches for Get-Sessions and Get-MessageIDs."/>
                            </Grid>
                        </Expander>

                        <Expander Header="Entra Logs  -  Get-GraphEntraSignInLogs / Get-GraphEntraAuditLogs" FontWeight="SemiBold" Margin="0,0,0,6">
                            <Grid Name="gridEntraOpts" Margin="10,8,10,10">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="240"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>

                                <TextBlock Text="-EventTypes (SignIn):"  Grid.Row="0" Grid.Column="0" Margin="0,4,0,4" FontWeight="Normal"/>
                                <TextBox   Name="txtEntraEventTypes"      Grid.Row="0" Grid.Column="1" Margin="0,2,0,2"
                                           ToolTip="Comma-separated. Values: All, interactiveUser, nonInteractiveUser, servicePrincipal, managedIdentity. Default: All"/>

                                <CheckBox  Name="cbEntraAuditAll" Grid.Row="1" Grid.Column="0" Grid.ColumnSpan="2" Margin="0,8,0,4" Content="-All (GraphEntraAuditLogs)" FontWeight="Normal"
                                           ToolTip="Match against userPrincipalNames AND targetResources"/>

                                <TextBlock Grid.Row="2" Grid.Column="0" Grid.ColumnSpan="2" Margin="0,12,0,0" TextWrapping="Wrap" FontWeight="Normal"
                                           Text="-EventTypes applies to Get-GraphEntraSignInLogs. -All applies to Get-GraphEntraAuditLogs."/>
                            </Grid>
                        </Expander>

                        <Expander Header="Identity  -  Get-MFA / Get-AllRoleActivity" FontWeight="SemiBold" Margin="0,0,0,6">
                            <Grid Name="gridIdentityOpts" Margin="10,8,10,10">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>

                                <CheckBox  Name="cbMfaIncludePhone"      Grid.Row="0" Margin="0,4,0,4" Content="-IncludePhoneNumbers (Get-MFA)" FontWeight="Normal"/>
                                <CheckBox  Name="cbRolesIncludeEmpty"    Grid.Row="1" Margin="0,4,0,4" Content="-IncludeEmptyRoles (Get-AllRoleActivity)" FontWeight="Normal"/>

                                <TextBlock Grid.Row="2" Margin="0,12,0,0" TextWrapping="Wrap" FontWeight="Normal"
                                           Text="Optional flags for Get-MFA and Get-AllRoleActivity."/>
                            </Grid>
                        </Expander>

                        <Expander Header="Security  -  Get-SecureScore / Get-SecurityAlerts" FontWeight="SemiBold" Margin="0,0,0,6">
                            <Grid Name="gridSecurityOpts" Margin="10,8,10,10">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="200"/>
                                    <ColumnDefinition Width="280"/>
                                    <ColumnDefinition Width="20"/>
                                    <ColumnDefinition Width="200"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>

                                <TextBlock Text="SecureScore -Category:"  Grid.Row="0" Grid.Column="0" Margin="0,4,0,4" FontWeight="Normal"/>
                                <ComboBox  Name="cmbSecScoreCategory"      Grid.Row="0" Grid.Column="1" Margin="0,2,0,2">
                                    <ComboBoxItem Content=""/>
                                    <ComboBoxItem Content="All"/>
                                    <ComboBoxItem Content="Identity"/>
                                    <ComboBoxItem Content="Data"/>
                                    <ComboBoxItem Content="Device"/>
                                    <ComboBoxItem Content="Apps"/>
                                    <ComboBoxItem Content="Infrastructure"/>
                                </ComboBox>

                                <TextBlock Text="SecureScore -Service:"   Grid.Row="1" Grid.Column="0" Margin="0,4,0,4" FontWeight="Normal"/>
                                <ComboBox  Name="cmbSecScoreService"       Grid.Row="1" Grid.Column="1" Margin="0,2,0,2">
                                    <ComboBoxItem Content=""/>
                                    <ComboBoxItem Content="All"/>
                                    <ComboBoxItem Content="Exchange"/>
                                    <ComboBoxItem Content="SharePoint"/>
                                    <ComboBoxItem Content="AAD"/>
                                </ComboBox>

                                <TextBlock Text="SecureScore -StatusFilter:" Grid.Row="2" Grid.Column="0" Margin="0,4,0,4" FontWeight="Normal"
                                           ToolTip="Leave blank to include all statuses."/>
                                <ComboBox  Name="cmbSecScoreStatus"        Grid.Row="2" Grid.Column="1" Margin="0,2,0,2"
                                           ToolTip="Leave blank to include all statuses.">
                                    <ComboBoxItem Content=""/>
                                    <ComboBoxItem Content="AtRisk"/>
                                    <ComboBoxItem Content="Partial"/>
                                    <ComboBoxItem Content="MeetsStandard"/>
                                    <ComboBoxItem Content="NotApplicable"/>
                                </ComboBox>

                                <TextBlock Text="SecurityAlerts -AlertId:" Grid.Row="0" Grid.Column="3" Margin="10,4,0,4" FontWeight="Normal"/>
                                <TextBox   Name="txtSecAlertId"            Grid.Row="0" Grid.Column="4" Margin="0,2,0,2"/>

                                <TextBlock Text="SecurityAlerts -DaysBack:" Grid.Row="1" Grid.Column="3" Margin="10,4,0,4" FontWeight="Normal"/>
                                <TextBox   Name="txtSecAlertDaysBack"      Grid.Row="1" Grid.Column="4" Margin="0,2,0,2"
                                           ToolTip="Default 90"/>

                                <TextBlock Text="SecurityAlerts -Filter:"  Grid.Row="2" Grid.Column="3" Margin="10,4,0,4" FontWeight="Normal"/>
                                <TextBox   Name="txtSecAlertFilter"        Grid.Row="2" Grid.Column="4" Margin="0,2,0,2"/>

                                <TextBlock Grid.Row="4" Grid.Column="0" Grid.ColumnSpan="5" Margin="0,12,0,0" TextWrapping="Wrap" FontWeight="Normal"
                                           Text="Options for Get-SecureScore and Get-SecurityAlerts."/>
                            </Grid>
                        </Expander>

                    </StackPanel>
                </ScrollViewer>
            </TabItem>

            <TabItem Header="M.E.G Setup">
                <Grid Margin="0,8,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="440"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <GroupBox Header="Modules &amp; Scripts" Grid.Column="0" Margin="0,0,10,0">
                        <DockPanel>
                            <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="6,6,6,4">
                                <Button Name="btnSetupSelectAll" Content="Select All" Width="100" Margin="0,0,6,0"/>
                                <Button Name="btnSetupSelectNone" Content="Select None" Width="100"/>
                            </StackPanel>
                            <ScrollViewer VerticalScrollBarVisibility="Auto">
                                <StackPanel Margin="6">
                                    <TextBlock Text="PowerShell Modules" FontWeight="SemiBold" Margin="0,4,0,4"/>
                                    <CheckBox Name="cbModExchangeOnline" Content="ExchangeOnlineManagement"/>
                                    <CheckBox Name="cbModGraph"          Content="Microsoft.Graph"/>
                                    <CheckBox Name="cbModGraphBeta"      Content="Microsoft.Graph.Beta"/>
                                    <CheckBox Name="cbModExtractor"      Content="Microsoft-Extractor-Suite"/>

                                </StackPanel>
                            </ScrollViewer>
                        </DockPanel>
                    </GroupBox>

                    <DockPanel Grid.Column="1">
                        <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="0,6,6,6">
                            <Button Name="btnSetupInstall"   Content="Install Selected"   Width="140" Margin="0,0,6,0"/>
                            <Button Name="btnSetupUpdate"    Content="Update Selected"    Width="140" Margin="0,0,6,0"/>
                            <Button Name="btnSetupUninstall" Content="Uninstall Selected" Width="140" Margin="0,0,6,0"/>
                        </StackPanel>
                        <GroupBox Header="Setup Commands">
                            <TextBox Name="txtSetupCommands"
                                     Margin="6"
                                     AcceptsReturn="True"
                                     VerticalScrollBarVisibility="Auto"
                                     HorizontalScrollBarVisibility="Auto"
                                     TextWrapping="NoWrap"
                                     IsReadOnly="True"
                                     FontFamily="Consolas"/>
                        </GroupBox>
                    </DockPanel>
                </Grid>
            </TabItem>
        </TabControl>

        <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button Name="btnInstall" Content="Install / Setup Tenant" Width="170" Margin="0,0,10,0"/>
            <Button Name="btnBuild" Content="Build Commands" Width="200" Margin="0,0,10,0"/>
            <Button Name="btnExecute" Content="Execute" Width="100" Margin="0,0,10,0"/>
            <Button Name="btnExit" Content="Exit" Width="80"/>
        </StackPanel>

        <StatusBar Grid.Row="6" Margin="0,8,0,0">
            <StatusBarItem>
                <TextBlock Name="txtStatus" Text="Ready"/>
            </StatusBarItem>
        </StatusBar>
    </Grid>
</Window>
'@

# ----------------------------------------
# Load XAML
# ----------------------------------------
$reader = New-Object System.Xml.XmlNodeReader $xaml
try {
    $window = [Windows.Markup.XamlReader]::Load($reader)
}
catch {
    Write-Host "Failed to load XAML: $($_.Exception.Message)" -ForegroundColor Red
    throw
}

# Map every Name= node into the $controls hashtable
$controls = @{}
$xaml.SelectNodes("//*[@*[local-name()='Name']]") | ForEach-Object {
    $name = $_.GetAttribute('Name')
    $controls[$name] = $window.FindName($name)
}

# Clamp startup size to the work area so the window fits smaller displays
$workArea = [System.Windows.SystemParameters]::WorkArea
if ($window.Height -gt $workArea.Height) { $window.Height = [Math]::Max($window.MinHeight, $workArea.Height - 20) }
if ($window.Width -gt $workArea.Width)   { $window.Width  = [Math]::Max($window.MinWidth,  $workArea.Width  - 20) }

# UAL operations list used by Triage mode when building Get-UALGraph commands
$ualOperationsListText = '"Change user password.","Consent to application.","Disable-InboxRule","New-InboxRule","Reset user password.","Set-inboxrule","SignInEvent","UpdateInboxRules","UserLoggedIn","Add service principal.","Enable-InboxRule","Remove-InboxRule","UserLoginFailed","Add app role assignment grant to user.","Add app role assignment to service principal.","Add application.","Add delegated permission grant.","Add device.","Add group.","Add member to group.","Add member to role.","Add owner to group.","Add registered owner to device.","Add registered users to device.","Add user.","AddedToGroup","AppDeleted","AppInstalled","ApplicationInstallationCompleted","ApplicationInstallationStarted","AppUninstalled","AppUpgraded","AuditSearchCompleted","AuditSearchCreated","AuditSearchExportJobCompleted","AuditSearchExportJobCreated","BulkUpdate","Change user license.","Create","Delete application.","Delete device.","Delete group.","Delete user.","Disable account.","Enable-TransportRule","FileAccessed","FileCopied","FileDeleted","FileDeletedFirstStageRecycleBin","FileDownloaded","FileModified","FileMoved","FilePreviewed","FileRecycled","FileRenamed","FileRestored","FileSyncDownloadedFull","FileSyncUploadedFull","FolderCreated","FolderModified","FolderMoved","FolderRecycled","FolderRenamed","GATFRTokenIssue","HardDelete","MailItemsAccessed","MessageSent","MessageUpdated","MoveToDeletedItems","New-ComplianceSearch","New-Mailbox","New-TransportRule","PageViewed","Remove app role assignment from service principal.","Remove delegated permission grant.","Remove member from group.","Remove member from role.","Remove service principal.","Send","SendAs","SendOnBehalf","Set-Mailbox","SharingLinkUpdated","SharingLinkUsed","SoftDelete","Update","Update application - Certificates and secrets management","Update application.","Update device.","Update service principal.","Update StsRefreshTokenValidFrom Timestamp.","Update user.","UserSubmission","isThrottled"'

# Maps every EXO/Graph checkbox to the Get-* line it should emit.
# Order in each hashtable is preserved by [ordered].
$exoCommands = [ordered]@{
    cbMailboxAuditStatus = 'Get-MailboxAuditStatus'
    cbMailboxPermissions = 'Get-MailboxPermissions'
    cbMailboxRules       = 'Get-MailboxRules'
    cbShowMailboxRules   = 'Show-MailboxRules'
    cbAdminAuditLog      = 'Get-AdminAuditLog'
    cbMailboxAuditLog    = 'Get-MailboxAuditLog'
    cbUALClassic         = 'Get-UAL'
    cbEmail              = 'Get-Email'
    cbAttachment         = 'Get-Attachment'
    cbShowEmail          = 'Show-Email'
    cbSessions           = 'Get-Sessions'
    cbMessageIDs         = 'Get-MessageIDs'
    cbTransportRules     = 'Get-TransportRules'
    cbShowTransportRules = 'Show-TransportRules'
    cbMessageTrace       = 'Get-MessageTraceLog'
}

$graphCommands = [ordered]@{
    cbUsers             = 'Get-Users'
    cbAdminUsers        = 'Get-AdminUsers'
    cbMFA               = 'Get-MFA'
    cbRisky             = 'Get-RiskyDetections'
    cbRiskyUsers        = 'Get-RiskyUsers'
    cbAuth              = 'Get-OAuthPermissionsGraph'
    cbDevices           = 'Get-Devices'
    cbGroups            = 'Get-Groups'
    cbGroupMembers      = 'Get-GroupMembers'
    cbDynamicGroups     = 'Get-DynamicGroups'
    cbConditionalAccess = 'Get-ConditionalAccessPolicies'
    cbSecurityDefaults  = 'Get-EntraSecurityDefaults'
    cbSecureScore       = 'Get-SecureScore'
    cbSecurityAlerts    = 'Get-SecurityAlerts'
    cbAllRoleActivity   = 'Get-AllRoleActivity'
    cbPIMAssignments    = 'Get-PIMAssignments'
    cbLicenses          = 'Get-LicensesByUser'
    cbAllLicenses       = 'Get-Licenses'
    cbProductLicenses   = 'Get-ProductLicenses'
    cbGraphAudit        = 'Get-GraphEntraAuditLogs'
    cbGraphSignin       = 'Get-GraphEntraSignInLogs'
    cbMailboxRulesGraph = 'Get-MailboxRulesGraph'
}

$setupItems = [ordered]@{
    cbModExchangeOnline    = @{ Type = 'Module'; Name = 'ExchangeOnlineManagement' }
    cbModGraph             = @{ Type = 'Module'; Name = 'Microsoft.Graph' }
    cbModGraphBeta         = @{ Type = 'Module'; Name = 'Microsoft.Graph.Beta' }
    cbModExtractor         = @{ Type = 'Module'; Name = 'Microsoft-Extractor-Suite' }
}

# ----------------------------------------
# Tenant discovery
# ----------------------------------------
function Get-TenantConfigs {
    param([string]$Root)

    $result = @()
    if (-not (Test-Path -Path $Root)) { return $result }

    Get-ChildItem -Path $Root -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $cfg = Join-Path $_.FullName 'Config\m365Config.json'
        if (Test-Path -Path $cfg) {
            try {
                $json = Get-Content -Path $cfg -Raw | ConvertFrom-Json
                $required = 'AppId','TenantId','CertificateThumbprint','PrimaryDomain'
                $missing  = $required | Where-Object { -not $json.PSObject.Properties[$_] }
                $result += [pscustomobject]@{
                    FolderName            = $_.Name
                    ConfigPath            = $cfg
                    AppId                 = $json.AppId
                    TenantId              = $json.TenantId
                    CertificateThumbprint = $json.CertificateThumbprint
                    PrimaryDomain         = $json.PrimaryDomain
                    MissingFields         = $missing
                }
            } catch {
                $result += [pscustomobject]@{
                    FolderName    = $_.Name
                    ConfigPath    = $cfg
                    ParseError    = $_.Exception.Message
                    MissingFields = @('*** PARSE ERROR ***')
                }
            }
        }
    }
    $result
}

function Set-Status {
    param([string]$Text)
    $controls.txtStatus.Text = if ([string]::IsNullOrEmpty($Text)) { 'Ready' } else { $Text }
}

function Update-TenantButtonStates {
    if ($script:AsyncJob) { return }   # async completion re-calls this once the job finishes
    $hasTenant = ($controls.cmbTenant.SelectedIndex -ge 0)
    foreach ($name in @('btnTestUal', 'btnUninstall', 'btnBuild', 'btnExecute')) {
        if ($controls[$name]) { $controls[$name].IsEnabled = $hasTenant }
    }
}

# Per-Cmdlet Options panels stay greyed out until at least one of their cmdlets is selected
$script:expanderDeps = @{
    gridUalClassicOpts = @('cbUALClassic')
    gridEmailOpts      = @('cbEmail', 'cbAttachment', 'cbShowEmail')
    gridMailItemsOpts  = @('cbSessions', 'cbMessageIDs')
    gridEntraOpts      = @('cbGraphSignin', 'cbGraphAudit')
    gridIdentityOpts   = @('cbMFA', 'cbAllRoleActivity')
    gridSecurityOpts   = @('cbSecureScore', 'cbSecurityAlerts')
}

function Update-ExpanderStates {
    foreach ($gridName in $script:expanderDeps.Keys) {
        $enabled = $false
        foreach ($cbName in $script:expanderDeps[$gridName]) {
            if ($controls[$cbName] -and $controls[$cbName].IsChecked) { $enabled = $true; break }
        }
        if ($controls[$gridName]) { $controls[$gridName].IsEnabled = $enabled }
    }
}

# "Pull all UAL" overrides Last N days and Start/End dates - grey them out so the
# precedence is visible instead of surprising
function Update-UalRangeFieldStates {
    $pullAll = [bool]$controls.cbUalAll.IsChecked
    foreach ($name in @('txtUalDaysBack', 'dpUalStart', 'dpUalEnd')) {
        if ($controls[$name]) { $controls[$name].IsEnabled = -not $pullAll }
    }
}

function Update-TenantList {
    $controls.cmbTenant.Items.Clear()
    $script:tenantConfigs = @(Get-TenantConfigs -Root $ConfigRoot)

    if ($script:tenantConfigs.Count -eq 0) {
        $controls.txtErrors.Text = 'No existing tenants found. Click Install / Setup Tenant to create a new tenant.'
        $controls.txtWarnings.Text = ''
        $controls.txtDomain.Text = ''
        $controls.txtTenantId.Text = ''
        $controls.txtAppId.Text = ''
        $controls.txtThumbprint.Text = ''
        $controls.txtOutputDir.Text = ''
        $controls.cmbTenant.SelectedIndex = -1
        Update-TenantButtonStates
        return
    }

    foreach ($t in $script:tenantConfigs) {
        $label = if ($t.PrimaryDomain) { $t.PrimaryDomain } else { $t.FolderName }
        [void]$controls.cmbTenant.Items.Add($label)
    }
    $controls.cmbTenant.SelectedIndex = 0
    $controls.txtErrors.Text = ''
    Update-TenantButtonStates
}

function Set-TenantFields {
    $idx = $controls.cmbTenant.SelectedIndex
    if ($idx -lt 0 -or $idx -ge $script:tenantConfigs.Count) {
        $controls.txtDomain.Text     = ''
        $controls.txtTenantId.Text   = ''
        $controls.txtAppId.Text      = ''
        $controls.txtThumbprint.Text = ''
        $controls.txtOutputDir.Text  = ''
        $controls.txtErrors.Text     = ''
        $controls.txtWarnings.Text   = ''
        Update-TenantButtonStates
        return
    }
    $t = $script:tenantConfigs[$idx]

    $controls.txtDomain.Text     = [string]$t.PrimaryDomain
    $controls.txtTenantId.Text   = [string]$t.TenantId
    $controls.txtAppId.Text      = [string]$t.AppId
    $controls.txtThumbprint.Text = [string]$t.CertificateThumbprint
    $controls.txtOutputDir.Text  = Join-Path $ConfigRoot (Join-Path $t.FolderName 'Logs')

    $errs = @()
    $warns = @()
    if ($t.ParseError)               { $errs += "Failed to parse $($t.ConfigPath): $($t.ParseError)" }
    if ($t.MissingFields)             { $errs += "Missing fields in config: $($t.MissingFields -join ', ')" }
    if (-not $t.CertificateThumbprint){ $warns += 'No certificate thumbprint in config.' }

    $controls.txtErrors.Text   = ($errs -join [Environment]::NewLine)
    $controls.txtWarnings.Text = ($warns -join [Environment]::NewLine)
    Update-TenantButtonStates
}

# ----------------------------------------
# Command generation
# ----------------------------------------
function Get-SelectedCommands {
    param([System.Collections.Specialized.OrderedDictionary]$Map)
    $out = @()
    foreach ($key in $Map.Keys) {
        if ($controls[$key] -and $controls[$key].IsChecked) {
            $out += $Map[$key]
        }
    }
    $out
}

# Splits free-form multi-value input (commas, semicolons, or whitespace) into clean
# values. -CommaOnly keeps whitespace inside values (operation names contain spaces).
# Quote characters are stripped: users paste PowerShell-style "a","b" syntax into the
# boxes, and the leftover quotes break parameter binding in the generated commands.
function Split-MultiValue {
    param(
        [string]$Raw,
        [switch]$CommaOnly
    )
    $pattern = if ($CommaOnly) { '[,;]+' } else { '[\s,;]+' }
    @($Raw -split $pattern | ForEach-Object { ($_ -replace '["'']', '').Trim() } | Where-Object { $_ })
}

# One quoted comma-separated string: "a@x.com,b@x.com"
function Format-CommaString {
    param([string[]]$Values)
    '"' + ($Values -join ',') + '"'
}

# PowerShell array literal: "a@x.com","b@x.com"
function Format-PsArray {
    param([string[]]$Values)
    ($Values | ForEach-Object { '"{0}"' -f $_ }) -join ','
}

# How each cmdlet's -UserIds parameter accepts multiple values (verified against
# Microsoft-Extractor-Suite 4.1.0 source):
#   CommaString - [string] param; the module or EXO splits commas internally
#   Array       - [string[]] param with no internal split; each element lands in a
#                 Graph 'eq' filter, so a comma-joined string silently matches nothing
#   Single      - one value only; Build-CommandText fans out one command per user
$script:cmdletUserIdsFormat = @{
    'Get-UAL'                  = 'CommaString'
    'Get-AdminAuditLog'        = 'CommaString'
    'Get-MailboxAuditLog'      = 'CommaString'
    'Get-Devices'              = 'CommaString'
    'Get-Sessions'             = 'CommaString'
    'Get-MailboxRules'         = 'CommaString'
    'Get-MailboxPermissions'   = 'CommaString'
    'Get-MessageTraceLog'      = 'CommaString'
    'Get-MailboxAuditStatus'   = 'Array'
    'Get-UALGraph'             = 'Array'
    'Get-MFA'                  = 'Array'
    'Get-RiskyUsers'           = 'Array'
    'Get-RiskyDetections'      = 'Array'
    'Get-GraphEntraSignInLogs' = 'Array'
    'Get-GraphEntraAuditLogs'  = 'Array'
    'Get-Users'                = 'Single'
    'Get-Email'                = 'Single'
    'Get-Attachment'           = 'Single'
    'Show-Email'               = 'Single'
    'Get-MailboxRulesGraph'    = 'Single'
}

$script:cmdletCapabilities = @{
    'Get-MailboxAuditStatus'        = @('Encoding','LogLevel','UserIds')
    'Get-MailboxPermissions'        = @('Encoding','LogLevel','UserIds')
    'Get-MailboxRules'              = @('Encoding','LogLevel','UserIds')
    'Show-MailboxRules'             = @()
    'Get-TransportRules'            = @('Encoding','LogLevel')
    'Show-TransportRules'           = @()
    'Get-MessageTraceLog'           = @('Encoding','LogLevel','UserIds')
    'Get-AdminAuditLog'             = @('Encoding','LogLevel','Output','MergeOutput','UserIds')
    'Get-MailboxAuditLog'           = @('Encoding','LogLevel','Output','MergeOutput','UserIds')
    'Get-UAL'                       = @('Encoding','LogLevel','Output','MergeOutput','UserIds')
    'Get-Email'                     = @('LogLevel','UserIds')
    'Get-Attachment'                = @('LogLevel','UserIds')
    'Show-Email'                    = @('LogLevel','UserIds')
    'Get-Sessions'                  = @('Encoding','LogLevel','UserIds')
    'Get-MessageIDs'                = @('Encoding','LogLevel')
    'Get-ConditionalAccessPolicies' = @('Encoding','LogLevel')
    'Get-RiskyDetections'           = @('Encoding','LogLevel','UserIds')
    'Get-RiskyUsers'                = @('Encoding','LogLevel','UserIds')
    'Get-Users'                     = @('Encoding','LogLevel','UserIds')
    'Get-AdminUsers'                = @('Encoding','LogLevel')
    'Get-MFA'                       = @('Encoding','LogLevel','UserIds')
    'Get-Devices'                   = @('Encoding','LogLevel','Output','UserIds')
    'Get-Groups'                    = @('Encoding','LogLevel')
    'Get-GroupMembers'              = @('Encoding','LogLevel')
    'Get-DynamicGroups'             = @('Encoding','LogLevel')
    'Get-EntraSecurityDefaults'     = @('LogLevel')
    'Get-SecureScore'               = @('Encoding','LogLevel')
    'Get-SecurityAlerts'            = @('Encoding','LogLevel')
    'Get-AllRoleActivity'           = @('Encoding','LogLevel')
    'Get-PIMAssignments'            = @('Encoding','LogLevel')
    'Get-Licenses'                  = @('LogLevel')
    'Get-LicensesByUser'            = @('LogLevel')
    'Get-ProductLicenses'           = @('Encoding','LogLevel')
    'Get-OAuthPermissionsGraph'     = @('Encoding','LogLevel')
    'Get-GraphEntraAuditLogs'       = @('Encoding','LogLevel','MergeOutput','UserIds')
    'Get-GraphEntraSignInLogs'      = @('Encoding','LogLevel','Output','MergeOutput','UserIds')
    'Get-MailboxRulesGraph'         = @('Encoding','LogLevel','UserIds')
    'Get-UALGraph'                  = @('Encoding','LogLevel','Output','UserIds')
}

$script:cmdletUnsupportedOutputs = @{
    'Get-GraphEntraSignInLogs' = @('CSV')
}

function Get-GlobalFlags {
    param(
        [string]$Cmdlet,
        [string[]]$SkipCaps = @()
    )

    $caps = $script:cmdletCapabilities[$Cmdlet]
    if (-not $caps) { return '' }

    $sb = New-Object System.Text.StringBuilder

    if ($caps -contains 'UserIds' -and $SkipCaps -notcontains 'UserIds' -and $controls.txtGlobalUserId.Text) {
        $users = Split-MultiValue $controls.txtGlobalUserId.Text
        if ($users.Count -gt 0) {
            switch ($script:cmdletUserIdsFormat[$Cmdlet]) {
                'Array'       { [void]$sb.Append(" -UserIds $(Format-PsArray $users)") }
                'Single'      {
                    # Multiple users are fanned out per-command by Build-CommandText
                    if ($users.Count -eq 1) { [void]$sb.Append(" -UserIds `"$($users[0])`"") }
                }
                default       { [void]$sb.Append(" -UserIds $(Format-CommaString $users)") }
            }
        }
    }
    if ($caps -contains 'Output') {
        $output = if ($controls.cmbGlobalOutput.SelectedItem) { $controls.cmbGlobalOutput.SelectedItem.Content } else { '' }
        $unsupported = $script:cmdletUnsupportedOutputs[$Cmdlet]
        if ($output -and -not ($unsupported -and ($unsupported -contains $output))) {
            [void]$sb.Append(" -Output $output")
        }
    }
    if ($caps -contains 'Encoding') {
        $encodingRaw = if ($controls.cmbGlobalEncoding.SelectedItem) { $controls.cmbGlobalEncoding.SelectedItem.Content } else { $controls.cmbGlobalEncoding.Text }
        $encoding = @(Split-MultiValue $encodingRaw) | Select-Object -First 1
        if ($encoding) { [void]$sb.Append(" -Encoding $encoding") }
    }
    if ($caps -contains 'LogLevel') {
        $logLevel = if ($controls.cmbGlobalLogLevel.SelectedItem) { $controls.cmbGlobalLogLevel.SelectedItem.Content } else { '' }
        if ($logLevel) { [void]$sb.Append(" -LogLevel $logLevel") }
    }
    if ($caps -contains 'MergeOutput' -and $controls.cbGlobalMergeOutput.IsChecked) {
        [void]$sb.Append(' -MergeOutput')
    }

    $sb.ToString()
}

function Get-CmdletOwnFlags {
    param([string]$Cmdlet)

    $sb = New-Object System.Text.StringBuilder

    switch ($Cmdlet) {
        'Get-UAL' {
            $g = $controls.cmbUalClassicGroup.SelectedItem
            if ($g -and $g.Content) { [void]$sb.Append(" -Group $($g.Content)") }
            $recordTypes = Split-MultiValue $controls.txtUalClassicRecordType.Text
            if ($recordTypes.Count -gt 0)                 { [void]$sb.Append(" -RecordType $(Format-PsArray $recordTypes)") }
            $operations = Split-MultiValue $controls.txtUalClassicOperation.Text -CommaOnly
            if ($operations.Count -gt 0)                  { [void]$sb.Append(" -Operations $(Format-PsArray $operations)") }
            if ($controls.txtUalClassicIp.Text)           { [void]$sb.Append(" -IPAddresses `"$($controls.txtUalClassicIp.Text)`"") }
            if ($controls.txtUalClassicObjectIds.Text)    { [void]$sb.Append(" -ObjectIDs `"$($controls.txtUalClassicObjectIds.Text)`"") }
            if ($controls.txtUalClassicInterval.Text)     { [void]$sb.Append(" -Interval $($controls.txtUalClassicInterval.Text)") }
            if ($controls.txtUalClassicTargetEvents.Text) { [void]$sb.Append(" -TargetEventsPerWindow $($controls.txtUalClassicTargetEvents.Text)") }
            if ($controls.cbUalClassicAuditDataOnly.IsChecked) { [void]$sb.Append(' -AuditDataOnly') }
        }
        'Get-Email' {
            if ($controls.txtEmailInternetMessageId.Text) { [void]$sb.Append(" -InternetMessageId `"$($controls.txtEmailInternetMessageId.Text)`"") }
            if ($controls.txtEmailInputFile.Text)         { [void]$sb.Append(" -inputFile `"$($controls.txtEmailInputFile.Text)`"") }
            if ($controls.cbEmailAttachment.IsChecked)        { [void]$sb.Append(' -Attachment $true') }
            if ($controls.cbEmailDownloadDuplicates.IsChecked){ [void]$sb.Append(' -DownloadDuplicates $true') }
        }
        'Get-Attachment' {
            if ($controls.txtEmailInternetMessageId.Text) { [void]$sb.Append(" -InternetMessageId `"$($controls.txtEmailInternetMessageId.Text)`"") }
        }
        'Show-Email' {
            if ($controls.txtEmailInternetMessageId.Text) { [void]$sb.Append(" -InternetMessageId `"$($controls.txtEmailInternetMessageId.Text)`"") }
        }
        'Get-Sessions' {
            # Module compares -IP with -eq against ClientIPAddress: single IP only
            $ips = Split-MultiValue $controls.txtMailItemsIp.Text
            if ($ips.Count -gt 0) { [void]$sb.Append(" -IP `"$($ips[0])`"") }
        }
        'Get-MessageIDs' {
            $ips = Split-MultiValue $controls.txtMailItemsIp.Text
            if ($ips.Count -gt 0)                    { [void]$sb.Append(" -IP `"$($ips[0])`"") }
            if ($controls.txtMailItemsSessions.Text) { [void]$sb.Append(" -Sessions `"$($controls.txtMailItemsSessions.Text)`"") }
            $d = $controls.cmbMailItemsDownload.SelectedItem
            if ($d -and $d.Content) { [void]$sb.Append(" -Download $($d.Content)") }
        }
        'Get-GraphEntraSignInLogs' {
            $eventTypes = Split-MultiValue $controls.txtEntraEventTypes.Text
            if ($eventTypes.Count -gt 0) { [void]$sb.Append(" -EventTypes $(Format-PsArray $eventTypes)") }
        }
        'Get-GraphEntraAuditLogs' {
            if ($controls.cbEntraAuditAll.IsChecked) { [void]$sb.Append(' -All') }
        }
        'Get-MFA' {
            if ($controls.cbMfaIncludePhone.IsChecked) { [void]$sb.Append(' -IncludePhoneNumbers') }
        }
        'Get-AllRoleActivity' {
            if ($controls.cbRolesIncludeEmpty.IsChecked) { [void]$sb.Append(' -IncludeEmptyRoles') }
        }
        'Get-SecureScore' {
            $c = $controls.cmbSecScoreCategory.SelectedItem; if ($c -and $c.Content) { [void]$sb.Append(" -Category $($c.Content)") }
            $s = $controls.cmbSecScoreService.SelectedItem;  if ($s -and $s.Content) { [void]$sb.Append(" -Service $($s.Content)") }
            $f = $controls.cmbSecScoreStatus.SelectedItem;   if ($f -and $f.Content) { [void]$sb.Append(" -StatusFilter $($f.Content)") }
        }
        'Get-SecurityAlerts' {
            if ($controls.txtSecAlertId.Text)       { [void]$sb.Append(" -AlertId `"$($controls.txtSecAlertId.Text)`"") }
            if ($controls.txtSecAlertDaysBack.Text) { [void]$sb.Append(" -DaysBack $($controls.txtSecAlertDaysBack.Text)") }
            if ($controls.txtSecAlertFilter.Text)   { [void]$sb.Append(" -Filter `"$($controls.txtSecAlertFilter.Text)`"") }
        }
    }

    $sb.ToString()
}

function Test-CmdletPrerequisites {
    param([string]$Cmdlet)

    switch ($Cmdlet) {
        { $_ -in 'Get-Email', 'Get-Attachment', 'Show-Email' } {
            if (-not $controls.txtEmailInternetMessageId.Text) { return "$Cmdlet skipped - set -InternetMessageId on the Email Options tab." }
            if ((Split-MultiValue $controls.txtGlobalUserId.Text).Count -eq 0) { return "$Cmdlet skipped - the module requires -UserIds; set the Global Options -UserIds field." }
        }
    }
    return $null
}

function Get-UalOptionFlags {
    $sb = New-Object System.Text.StringBuilder

    # [string[]] parameters on Get-UALGraph - must be PS arrays, not comma-joined strings
    $arrayInputs = @(
        @{ Param = 'IpAddress';  Ctl = 'txtUalIp' }
        @{ Param = 'RecordType'; Ctl = 'txtUalRecordType' }
        @{ Param = 'ObjectIDs';  Ctl = 'txtUalObjectId' }
    )
    foreach ($i in $arrayInputs) {
        $vals = Split-MultiValue $controls[$i.Ctl].Text
        if ($vals.Count -gt 0) { [void]$sb.Append(" -$($i.Param) $(Format-PsArray $vals)") }
    }

    # [string] parameters - single value
    $stringInputs = @(
        @{ Param = 'Service';     Ctl = 'txtUalService' }
        @{ Param = 'Keyword';     Ctl = 'txtUalKeyword' }
    )
    foreach ($i in $stringInputs) {
        $val = $controls[$i.Ctl].Text
        if ($val) { [void]$sb.Append(" -$($i.Param) `"$val`"") }
    }

    $maxEvents = $controls.txtUalMaxEvents.Text
    if ($maxEvents) { [void]$sb.Append(" -MaxEventsPerFile $maxEvents") }

    if ($controls.cbUalSplitFiles.IsChecked) { [void]$sb.Append(' -SplitFiles') }
    if ($controls.cbUalUseV1.IsChecked)      { [void]$sb.Append(' -UseV1') }

    $sb.ToString()
}

function Add-WrappedCmdlet {
    param(
        [System.Text.StringBuilder]$Sb,
        [string]$Label,
        [string]$Line,
        [System.Text.StringBuilder]$PreviewSb
    )
    $labelEsc = $Label -replace "'", "''"
    [void]$Sb.AppendLine("try {")
    [void]$Sb.AppendLine("    Write-Host `"[RUN] $Label`" -ForegroundColor Cyan")
    [void]$Sb.AppendLine("    $Line")
    [void]$Sb.AppendLine("} catch {")
    [void]$Sb.AppendLine("    `$script:m365Failures += '$labelEsc'")
    [void]$Sb.AppendLine("    Write-Host `"[ERROR] $Label : `$(`$_.Exception.Message)`" -ForegroundColor Red")
    [void]$Sb.AppendLine("}")
    [void]$Sb.AppendLine("")

    if ($PreviewSb) {
        [void]$PreviewSb.AppendLine($Line)
        [void]$PreviewSb.AppendLine("")
    }
}

function Build-CommandText {
    $appId      = $controls.txtAppId.Text
    $tenantId   = $controls.txtTenantId.Text
    $thumbprint = $controls.txtThumbprint.Text
    $domain     = $controls.txtDomain.Text
    $userId     = $controls.txtGlobalUserId.Text
    $userList   = Split-MultiValue $userId
    $outputRoot = $controls.txtOutputDir.Text

    $errs = @()
    $warns = @()
    if (-not $appId)      { $errs += 'No App ID - pick a tenant first.' }
    if (-not $tenantId)   { $errs += 'No Tenant ID - pick a tenant first.' }
    if (-not $thumbprint) { $errs += 'No certificate thumbprint - pick a tenant first.' }
    if (-not $domain)     { $errs += 'No primary domain - pick a tenant first.' }
    if (-not $outputRoot) { $errs += 'Output directory is required.' }

    $exoSelected   = Get-SelectedCommands -Map $exoCommands
    $graphSelected = Get-SelectedCommands -Map $graphCommands

    $ualRange = $null
    $startDt  = $controls.dpUalStart.SelectedDate
    $endDt    = $controls.dpUalEnd.SelectedDate
    $daysRaw  = if ($controls.txtUalDaysBack.Text) { $controls.txtUalDaysBack.Text.Trim() } else { '' }
    $pullAll  = [bool]$controls.cbUalAll.IsChecked

    if ($pullAll) {
        $ualRange = [pscustomobject]@{
            Start  = '(Get-Date).AddDays(-365).ToString("yyyy-MM-ddTHH:mm:ssZ")'
            End    = $null
            Label  = 'PullAll365Days'
            IsExpr = $true
        }
    }
    elseif ($startDt -or $endDt) {
        if (-not $startDt) {
            $errs += 'UAL: Start date required when End date is set.'
        }
        else {
            $effectiveEnd = if ($endDt) { $endDt } else { Get-Date }
            if ($effectiveEnd -lt $startDt) {
                $errs += 'UAL: End date must be >= Start date.'
            }
            else {
                $ualRange = [pscustomobject]@{
                    Start  = $startDt.ToString("yyyy-MM-dd")
                    End    = if ($endDt) { $endDt.ToString("yyyy-MM-dd") } else { $null }
                    Label  = "$($startDt.ToString('yyyyMMdd'))-$(if ($endDt) { $endDt.ToString('yyyyMMdd') } else { 'now' })"
                    IsExpr = $false
                }
            }
        }
    }
    elseif ($daysRaw) {
        $n = 0
        if (-not [int]::TryParse($daysRaw, [ref]$n) -or $n -le 0) {
            $errs += "UAL: 'Last N days' must be a positive integer (got '$daysRaw')."
        }
        else {
            $ualRange = [pscustomobject]@{
                Start  = "(Get-Date).AddDays(-$n).ToString(`"yyyy-MM-ddTHH:mm:ssZ`")"
                End    = $null
                Label  = "Last${n}Days"
                IsExpr = $true
            }
        }
    }

    if (-not $exoSelected -and -not $graphSelected -and -not $ualRange) {
        $warns += 'No log collections selected.'
    }

    $mailItemIps = Split-MultiValue $controls.txtMailItemsIp.Text
    if ($mailItemIps.Count -gt 1 -and (@('Get-Sessions','Get-MessageIDs') | Where-Object { $exoSelected -contains $_ })) {
        $warns += "Get-Sessions/Get-MessageIDs -IP accepts a single IP only; using the first ($($mailItemIps[0]))."
    }

    $maxEventsRaw = $controls.txtUalMaxEvents.Text.Trim()
    if ($ualRange -and $maxEventsRaw) {
        $maxEventsVal = 0
        if (-not [int]::TryParse($maxEventsRaw, [ref]$maxEventsVal) -or $maxEventsVal -le 0) {
            $errs += "UAL: MaxEventsPerFile must be a single positive integer (got '$maxEventsRaw')."
        }
    }

    $controls.txtErrors.Text = ($errs -join [Environment]::NewLine)
    if ($errs.Count -gt 0) {
        $controls.txtWarnings.Text = ($warns -join [Environment]::NewLine)
        return $null
    }

    $graphSb        = New-Object System.Text.StringBuilder
    $exoSb          = New-Object System.Text.StringBuilder
    $graphPreviewSb = New-Object System.Text.StringBuilder
    $exoPreviewSb   = New-Object System.Text.StringBuilder

    if ($graphSelected.Count -gt 0 -or $ualRange) {
        $connectLine = "Connect-MgGraph -ClientId `"$appId`" -TenantId `"$tenantId`" -CertificateThumbprint `"$thumbprint`" -NoWelcome"
        [void]$graphSb.AppendLine($connectLine)
        [void]$graphSb.AppendLine('')
        [void]$graphPreviewSb.AppendLine($connectLine)
        [void]$graphPreviewSb.AppendLine('')
        foreach ($cmd in $graphSelected) {
            $skip = Test-CmdletPrerequisites -Cmdlet $cmd
            if ($skip) { $warns += $skip; continue }
            $own = Get-CmdletOwnFlags -Cmdlet $cmd
            $caps = $script:cmdletCapabilities[$cmd]
            if ($script:cmdletUserIdsFormat[$cmd] -eq 'Single' -and $caps -contains 'UserIds' -and $userList.Count -gt 1) {
                # Module only accepts one user for this cmdlet - emit one command per user
                $gNoUser = Get-GlobalFlags -Cmdlet $cmd -SkipCaps 'UserIds'
                foreach ($u in $userList) {
                    Add-WrappedCmdlet -Sb $graphSb -PreviewSb $graphPreviewSb -Label "$cmd ($u)" -Line "$cmd -UserIds `"$u`"$gNoUser$own"
                }
            }
            else {
                $g = Get-GlobalFlags -Cmdlet $cmd
                Add-WrappedCmdlet -Sb $graphSb -PreviewSb $graphPreviewSb -Label $cmd -Line "$cmd$g$own"
            }
        }
        if ($ualRange) {
            $triage      = [bool]$controls.cbUALTriage.IsChecked
            $customOps   = $controls.txtUalOperations.Text
            $globalFlag  = Get-GlobalFlags -Cmdlet 'Get-UALGraph'
            $ualFlags    = Get-UalOptionFlags
            $customOpsList = Split-MultiValue $customOps -CommaOnly
            $opsTail     =
                if     ($customOpsList.Count -gt 0) { " -Operations $(Format-PsArray $customOpsList)" }
                elseif ($triage)                    { ' -Operations $ualOperationsListText' }
                else                                { '' }
            [void]$graphSb.AppendLine('')
            [void]$graphPreviewSb.AppendLine('')
            if ($customOpsList.Count -eq 0 -and $triage) {
                $preludeLine = "`$ualOperationsListText = $ualOperationsListText"
                [void]$graphSb.AppendLine($preludeLine)
                [void]$graphSb.AppendLine('')
                [void]$graphPreviewSb.AppendLine($preludeLine)
                [void]$graphPreviewSb.AppendLine('')
            }

            $startArg = if ($ualRange.IsExpr) { $ualRange.Start } else { "`"$($ualRange.Start)`"" }
            $ualLine  = "Get-UALGraph -SearchName SearchName-UAL-$($ualRange.Label) -StartDate $startArg"
            if ($ualRange.End) { $ualLine += " -EndDate `"$($ualRange.End)`"" }
            $ualLine += $globalFlag + $ualFlags + $opsTail
            Add-WrappedCmdlet -Sb $graphSb -PreviewSb $graphPreviewSb -Label "Get-UALGraph $($ualRange.Label)" -Line $ualLine
        }
        [void]$graphSb.AppendLine('')
        [void]$graphSb.AppendLine('Disconnect-MgGraph')
        [void]$graphPreviewSb.AppendLine('')
        [void]$graphPreviewSb.AppendLine('Disconnect-MgGraph')
    }

    if ($exoSelected.Count -gt 0) {
        $exoConnect = "Connect-ExchangeOnline -CertificateThumbprint `"$thumbprint`" -AppId `"$appId`" -Organization `"$domain`""
        [void]$exoSb.AppendLine($exoConnect)
        [void]$exoSb.AppendLine('')
        [void]$exoPreviewSb.AppendLine($exoConnect)
        [void]$exoPreviewSb.AppendLine('')
        foreach ($cmd in $exoSelected) {
            $skip = Test-CmdletPrerequisites -Cmdlet $cmd
            if ($skip) { $warns += $skip; continue }
            $own = Get-CmdletOwnFlags -Cmdlet $cmd
            $caps = $script:cmdletCapabilities[$cmd]
            $dateTail = if ($cmd -eq 'Get-MessageTraceLog' -and $userList.Count -gt 0) {
                ' -StartDate (Get-Date).AddDays(-90) -EndDate (Get-Date)'
            } else { '' }
            if ($script:cmdletUserIdsFormat[$cmd] -eq 'Single' -and $caps -contains 'UserIds' -and $userList.Count -gt 1) {
                # Module only accepts one user for this cmdlet - emit one command per user
                $gNoUser = Get-GlobalFlags -Cmdlet $cmd -SkipCaps 'UserIds'
                foreach ($u in $userList) {
                    Add-WrappedCmdlet -Sb $exoSb -PreviewSb $exoPreviewSb -Label "$cmd ($u)" -Line "$cmd -UserIds `"$u`"$gNoUser$own$dateTail"
                }
            }
            else {
                $g = Get-GlobalFlags -Cmdlet $cmd
                Add-WrappedCmdlet -Sb $exoSb -PreviewSb $exoPreviewSb -Label $cmd -Line "$cmd$g$own$dateTail"
            }
        }
        [void]$exoSb.AppendLine('')
        [void]$exoSb.AppendLine('Disconnect-ExchangeOnline -Confirm:$false')
        [void]$exoPreviewSb.AppendLine('')
        [void]$exoPreviewSb.AppendLine('Disconnect-ExchangeOnline -Confirm:$false')
    }

    $previewSb = New-Object System.Text.StringBuilder
    if ($graphPreviewSb.Length -gt 0) {
        [void]$previewSb.AppendLine('# --- Graph window ---')
        [void]$previewSb.Append($graphPreviewSb.ToString())
    }
    if ($graphPreviewSb.Length -gt 0 -and $exoPreviewSb.Length -gt 0) {
        [void]$previewSb.AppendLine('')
    }
    if ($exoPreviewSb.Length -gt 0) {
        [void]$previewSb.AppendLine('# --- EXO window ---')
        [void]$previewSb.Append($exoPreviewSb.ToString())
    }

    $controls.txtWarnings.Text = ($warns -join [Environment]::NewLine)
    [pscustomobject]@{
        Preview = $previewSb.ToString()
        Graph   = $graphSb.ToString()
        Exo     = $exoSb.ToString()
    }
}

function Wrap-WithStaging {
    param(
        [string]$Body,
        [ValidateSet('Graph','Exo')][string]$Phase = 'Graph',
        [string]$RunStamp
    )

    $idx = $controls.cmbTenant.SelectedIndex
    if ($idx -lt 0) { return $Body }
    $tenantDir     = Join-Path $ConfigRoot $script:tenantConfigs[$idx].FolderName
    $tenantDirLit  = $tenantDir.Replace("'", "''")
    $outputRootLit = $controls.txtOutputDir.Text.Replace("'", "''")
    if (-not $RunStamp) { $RunStamp = Get-Date -Format 'yyyyMMdd_HHmmss' }

    $modules = if ($Phase -eq 'Graph') {
        @('Microsoft.Graph.Authentication','Microsoft-Extractor-Suite')
    } else {
        @('ExchangeOnlineManagement','Microsoft-Extractor-Suite')
    }
    $moduleList = ($modules | ForEach-Object { "'$_'" }) -join ','

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("`$Host.UI.RawUI.WindowTitle = 'M365 GUI - $Phase'")
    [void]$sb.AppendLine("`$ErrorActionPreference = 'Continue'")
    [void]$sb.AppendLine("try { Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction Stop } catch { }")
    [void]$sb.AppendLine("if (`$PSVersionTable.PSEdition -eq 'Core') {")
    [void]$sb.AppendLine("    # Make Windows PowerShell CurrentUser modules visible when running under pwsh")
    [void]$sb.AppendLine("    `$winPsUserModules = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'WindowsPowerShell\Modules'")
    [void]$sb.AppendLine("    if ((Test-Path `$winPsUserModules) -and `$env:PSModulePath -notlike `"*`$winPsUserModules*`") { `$env:PSModulePath += ';' + `$winPsUserModules }")
    [void]$sb.AppendLine("}")
    [void]$sb.AppendLine("foreach (`$m in $moduleList) {")
    [void]$sb.AppendLine("    try { Import-Module `$m -ErrorAction Stop } catch { Write-Host `"[WARN] Could not import `$m : `$(`$_.Exception.Message)`" -ForegroundColor Yellow }")
    [void]$sb.AppendLine("}")
    [void]$sb.AppendLine("`$runStamp   = '$RunStamp'")
    [void]$sb.AppendLine("`$tenantDir  = '$tenantDirLit'")
    [void]$sb.AppendLine("`$outputRoot = '$outputRootLit'")
    [void]$sb.AppendLine("`$stagingDir = Join-Path `$tenantDir ('_staging\' + `$runStamp + '_$Phase')")
    [void]$sb.AppendLine("`$finalDir   = Join-Path `$outputRoot (`$runStamp + '\$Phase')")
    [void]$sb.AppendLine("New-Item -ItemType Directory -Path `$stagingDir -Force | Out-Null")
    [void]$sb.AppendLine("New-Item -ItemType Directory -Path `$finalDir   -Force | Out-Null")
    [void]$sb.AppendLine("Set-Location `$stagingDir")
    [void]$sb.AppendLine("`$script:m365Failures = @()")
    [void]$sb.AppendLine("try {")
    [void]$sb.Append($Body)
    [void]$sb.AppendLine("} finally {")
    [void]$sb.AppendLine("    Set-Location `$tenantDir")
    [void]$sb.AppendLine("    Get-ChildItem -Path `$stagingDir -Force | Move-Item -Destination `$finalDir -Force -ErrorAction Continue")
    [void]$sb.AppendLine("    if (Get-ChildItem -Path `$stagingDir -Recurse -Force -ErrorAction SilentlyContinue) {")
    [void]$sb.AppendLine("        Write-Host `"[WARN] Some files could not be moved and remain in: `$stagingDir`" -ForegroundColor Yellow")
    [void]$sb.AppendLine("    } else {")
    [void]$sb.AppendLine("        Remove-Item -Path `$stagingDir -Recurse -Force -ErrorAction SilentlyContinue")
    [void]$sb.AppendLine("    }")
    [void]$sb.AppendLine("    Write-Host `"$Phase logs moved to: `$finalDir`" -ForegroundColor Green")
    [void]$sb.AppendLine("    if (`$script:m365Failures.Count -gt 0) {")
    [void]$sb.AppendLine("        Write-Host ''")
    [void]$sb.AppendLine("        Write-Host `"[SUMMARY] `$(`$script:m365Failures.Count) cmdlet(s) failed in $Phase phase:`" -ForegroundColor Yellow")
    [void]$sb.AppendLine("        `$script:m365Failures | ForEach-Object { Write-Host `"  - `$_`" -ForegroundColor Yellow }")
    [void]$sb.AppendLine("    } else {")
    [void]$sb.AppendLine("        Write-Host `"[SUMMARY] All $Phase cmdlets completed without terminating errors.`" -ForegroundColor Green")
    [void]$sb.AppendLine("    }")
    [void]$sb.AppendLine("}")
    $sb.ToString()
}

function Invoke-SetupAction {
    param([ValidateSet('Install','Update','Uninstall')][string]$Action)

    $selected = @()
    foreach ($key in $setupItems.Keys) {
        if ($controls[$key] -and $controls[$key].IsChecked) {
            $selected += $setupItems[$key]
        }
    }

    if ($selected.Count -eq 0) {
        $controls.txtErrors.Text = 'Nothing selected — tick one or more items in the M.E.G Setup tab.'
        return
    }
    $controls.txtErrors.Text = ''

    $confirm = [System.Windows.MessageBox]::Show(
        "Run $Action for $($selected.Count) item(s) in this session (CurrentUser scope)?",
        "M.E.G Setup - $Action",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question)
    if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }

    # Flatten selection to plain data for the worker runspace
    $items = @(foreach ($i in $selected) { @{ Name = [string]$i.Name; Type = [string]$i.Type } })

    $controls.txtSetupCommands.Text = "# M.E.G Setup - $Action starting..."
    Set-Status "M.E.G Setup: $Action running..."

    Start-AsyncTask `
        -TaskArgs @{ Action = $Action; Items = $items } `
        -DisableButtons @('btnSetupInstall', 'btnSetupUpdate', 'btnSetupUninstall') `
        -Work {
            param($Sync, $TaskArgs)

            $Action = $TaskArgs.Action
            $log = New-Object System.Text.StringBuilder
            [void]$log.AppendLine("# M.E.G Setup - $Action (CurrentUser scope)")
            [void]$log.AppendLine('')
            $Sync.Log = $log.ToString()

            $hasPsItem = $TaskArgs.Items | Where-Object { $_.Type -eq 'Module' -or $_.Type -eq 'Script' }
            if ($hasPsItem) {
                try {
                    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction Stop
                    [void]$log.AppendLine('[INFO] Process execution policy set to Bypass.')
                }
                catch {
                    [void]$log.AppendLine("[WARN] Could not set execution policy: $($_.Exception.Message)")
                }
                try {
                    Import-Module PackageManagement -ErrorAction Stop
                    Import-Module PowerShellGet     -ErrorAction Stop
                    [void]$log.AppendLine('[INFO] PackageManagement and PowerShellGet imported.')
                }
                catch {
                    [void]$log.AppendLine("[WARN] Could not import PowerShellGet: $($_.Exception.Message)")
                }
            }
            if ($Action -eq 'Install' -and $hasPsItem) {
                try {
                    if (-not (Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue)) {
                        Register-PSRepository -Default
                    }
                    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
                    [void]$log.AppendLine('[INFO] PSGallery registered and trusted.')
                }
                catch {
                    [void]$log.AppendLine("[WARN] PSGallery setup: $($_.Exception.Message)")
                }
            }
            $Sync.Log = $log.ToString()

            foreach ($item in $TaskArgs.Items) {
                $name = $item.Name
                $Sync.Status = "M.E.G Setup: $Action $name..."
                try {
                    switch ($item.Type) {
                        'Module' {
                            switch ($Action) {
                                'Install' {
                                    [void]$log.AppendLine("[RUN]  Install-Module -Name $name -Scope CurrentUser -Force -AllowClobber")
                                    $Sync.Log = $log.ToString()
                                    Install-Module -Name $name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                                    [void]$log.AppendLine("[OK]   $name installed.")
                                }
                                'Update' {
                                    $installed = Get-InstalledModule -Name $name -AllVersions -ErrorAction SilentlyContinue |
                                                 Sort-Object Version -Descending | Select-Object -First 1
                                    if (-not $installed) {
                                        [void]$log.AppendLine("[SKIP] $name not installed - use Install Selected first.")
                                    }
                                    else {
                                        $latest = $null
                                        try { $latest = Find-Module -Name $name -Repository PSGallery -ErrorAction Stop }
                                        catch { [void]$log.AppendLine("[WARN] $name : could not query PSGallery - $($_.Exception.Message)") }

                                        if ($latest) {
                                            if ([version]$latest.Version -le [version]$installed.Version) {
                                                [void]$log.AppendLine("[SKIP] $name already current at $($installed.Version).")
                                            }
                                            else {
                                                [void]$log.AppendLine("[RUN]  Update-Module -Name $name -Force  ($($installed.Version) -> $($latest.Version))")
                                                $Sync.Log = $log.ToString()
                                                Update-Module -Name $name -Force -ErrorAction Stop
                                                [void]$log.AppendLine("[OK]   $name updated $($installed.Version) -> $($latest.Version).")
                                            }
                                        }
                                    }
                                }
                                'Uninstall' {
                                    [void]$log.AppendLine("[RUN]  Uninstall-Module -Name $name -AllVersions -Force")
                                    $Sync.Log = $log.ToString()
                                    Uninstall-Module -Name $name -AllVersions -Force -ErrorAction Stop
                                    [void]$log.AppendLine("[OK]   $name uninstalled.")
                                }
                            }
                        }
                        'Script' {
                            switch ($Action) {
                                'Install' {
                                    [void]$log.AppendLine("[RUN]  Install-Script -Name $name -Scope CurrentUser -Force")
                                    $Sync.Log = $log.ToString()
                                    Install-Script -Name $name -Scope CurrentUser -Force -ErrorAction Stop
                                    [void]$log.AppendLine("[OK]   $name script installed.")
                                }
                                'Update' {
                                    [void]$log.AppendLine("[RUN]  Install-Script -Name $name -Scope CurrentUser -Force")
                                    $Sync.Log = $log.ToString()
                                    Install-Script -Name $name -Scope CurrentUser -Force -ErrorAction Stop
                                    [void]$log.AppendLine("[OK]   $name script updated.")
                                }
                                'Uninstall' {
                                    [void]$log.AppendLine("[RUN]  Uninstall-Script -Name $name -Force")
                                    $Sync.Log = $log.ToString()
                                    Uninstall-Script -Name $name -Force -ErrorAction Stop
                                    [void]$log.AppendLine("[OK]   $name script uninstalled.")
                                }
                            }
                        }
                        'Git' {
                            $wingetArgs = switch ($Action) {
                                'Install'   { @('install', $name, '--silent', '--accept-package-agreements', '--accept-source-agreements') }
                                'Update'    { @('upgrade', $name, '--silent', '--accept-package-agreements', '--accept-source-agreements') }
                                'Uninstall' { @('uninstall', $name, '--silent') }
                            }
                            [void]$log.AppendLine("[RUN]  winget $($wingetArgs -join ' ')")
                            $Sync.Log = $log.ToString()
                            $wingetOut = & winget @wingetArgs 2>&1 | Out-String
                            if ($wingetOut.Trim()) { [void]$log.AppendLine($wingetOut.TrimEnd()) }
                            if ($LASTEXITCODE -ne 0) { throw "winget exited with code $LASTEXITCODE" }
                            [void]$log.AppendLine("[OK]   $name ($Action via winget) complete.")
                        }
                    }
                }
                catch {
                    [void]$log.AppendLine("[FAIL] $name : $($_.Exception.Message)")
                }
                $Sync.Log = $log.ToString()
            }

            [void]$log.AppendLine('')
            [void]$log.AppendLine('# Done.')
            $Sync.Log = $log.ToString()
            $Sync.Result = @{ Action = $Action }
            $Sync.Done = $true
        } `
        -OnProgress {
            param($Sync)
            if ($Sync.Status) { Set-Status $Sync.Status }
            if ($Sync.Log -and $controls.txtSetupCommands.Text -ne $Sync.Log) {
                $controls.txtSetupCommands.Text = $Sync.Log
                $controls.txtSetupCommands.ScrollToEnd()
            }
        } `
        -OnComplete {
            param($Sync)
            if ($Sync.Log) {
                $controls.txtSetupCommands.Text = $Sync.Log
                $controls.txtSetupCommands.ScrollToEnd()
            }
            if ($Sync.Error) {
                $controls.txtErrors.Text = "Setup task error: $($Sync.Error)"
                Set-Status ''
            }
            else {
                $actionLabel = if ($Sync.Result -and $Sync.Result.Action) { $Sync.Result.Action } else { 'action' }
                Set-Status "M.E.G Setup: $actionLabel complete."
            }
        }
}

# ----------------------------------------
# Background task plumbing (one task at a time)
# ----------------------------------------
# Pattern: the worker runspace writes plain values (strings/bools/hashtables of
# strings) into a synchronized hashtable; a DispatcherTimer on the UI thread
# polls it. Worker scriptblocks must only reference $Sync and $TaskArgs -
# $controls/$window/$script: variables do not exist in the worker runspace.
$script:AsyncJob   = $null
$script:AsyncTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:AsyncTimer.Interval = [TimeSpan]::FromMilliseconds(250)

function Start-AsyncTask {
    param(
        [Parameter(Mandatory)][scriptblock]$Work,   # param($Sync,$TaskArgs) - runs in background runspace
        [hashtable]$TaskArgs = @{},                 # plain data only
        [string[]]$DisableButtons = @(),
        [scriptblock]$OnProgress = $null,           # param($Sync) - UI thread, every tick
        [scriptblock]$OnComplete = $null            # param($Sync) - UI thread, once
    )

    if ($script:AsyncJob) {
        $controls.txtErrors.Text = 'Another background task is still running.'
        return
    }

    $sync = [hashtable]::Synchronized(@{
        Status = ''      # short transient text for the status bar
        Log    = ''      # full log text (setup actions)
        Done   = $false
        Result = $null   # plain hashtable describing the outcome
        Error  = $null   # terminal error message string
    })

    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.ThreadOptions  = 'ReuseThread'
    $rs.Open()
    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript($Work).AddArgument($sync).AddArgument($TaskArgs)

    foreach ($name in $DisableButtons) {
        if ($controls[$name]) { $controls[$name].IsEnabled = $false }
    }

    $script:AsyncJob = @{
        PowerShell = $ps
        Runspace   = $rs
        Handle     = $ps.BeginInvoke()
        Sync       = $sync
        OnProgress = $OnProgress
        OnComplete = $OnComplete
        Buttons    = $DisableButtons
    }
    $script:AsyncTimer.Start()
}

$script:AsyncTimer.Add_Tick({
    $job = $script:AsyncJob
    if (-not $job) { $script:AsyncTimer.Stop(); return }
    $sync = $job.Sync

    if ($job.OnProgress) { & $job.OnProgress $sync }

    if ($sync.Done -or $job.Handle.IsCompleted) {
        $script:AsyncTimer.Stop()
        $script:AsyncJob = $null
        try { [void]$job.PowerShell.EndInvoke($job.Handle) }
        catch { if (-not $sync.Error) { $sync.Error = $_.Exception.Message } }
        $job.PowerShell.Dispose()
        $job.Runspace.Close()
        $job.Runspace.Dispose()

        foreach ($name in $job.Buttons) {
            if ($controls[$name]) { $controls[$name].IsEnabled = $true }
        }
        Update-TenantButtonStates
        if ($job.OnComplete) { & $job.OnComplete $sync }
    }
})

# ----------------------------------------
# Event wiring
# ----------------------------------------
$controls.btnRefresh.Add_Click({ Update-TenantList })
$controls.cmbTenant.Add_SelectionChanged({ Set-TenantFields })

$controls.btnSelectAll.Add_Click({
    $excludeFromSelectAll = @(
        'cbShowMailboxRules',
        'cbMailboxAuditLog',
        'cbUALClassic',
        'cbEmail','cbAttachment','cbShowEmail',
        'cbSessions','cbMessageIDs',
        'cbShowTransportRules',
        'cbMessageTrace'
    )
    foreach ($key in @($exoCommands.Keys) + @($graphCommands.Keys)) {
        if ($excludeFromSelectAll -contains $key) { continue }
        if ($controls[$key]) { $controls[$key].IsChecked = $true }
    }
})

$controls.btnSelectNone.Add_Click({
    foreach ($key in @($exoCommands.Keys) + @($graphCommands.Keys)) {
        if ($controls[$key]) { $controls[$key].IsChecked = $false }
    }
})

$controls.btnBuild.Add_Click({
    $built = Build-CommandText
    if ($null -ne $built) {
        $script:lastBuildResult = $built
        $controls.txtCommands.Text = $built.Preview
    }
})

$controls.btnExecute.Add_Click({
    if (-not $script:lastBuildResult -or
        (-not $script:lastBuildResult.Graph -and -not $script:lastBuildResult.Exo)) {
        $controls.txtErrors.Text = 'Nothing to execute - click Build Commands first.'
        return
    }

    $windowCount = 0
    if ($script:lastBuildResult.Graph) { $windowCount++ }
    if ($script:lastBuildResult.Exo)   { $windowCount++ }

    $prompt = if ($windowCount -eq 2) {
        'Run Graph and Exchange Online in two separate PowerShell windows in parallel?'
    } else {
        'Run the generated commands in a new PowerShell window?'
    }
    $result = [System.Windows.MessageBox]::Show($prompt, 'Execute',
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question)
    if ($result -ne [System.Windows.MessageBoxResult]::Yes) { return }

    $runStamp = Get-Date -Format 'yyyyMMdd_HHmmss'

    if ($script:lastBuildResult.Graph) {
        # Get-UALGraph 400s under Windows PowerShell 5.1: ConvertTo-Json serializes the
        # query dates as "\/Date(ms)\/" which the Graph API rejects. PS7 emits ISO 8601,
        # so prefer pwsh for the Graph window when it is installed.
        $graphPsExe = if (Get-Command pwsh.exe -ErrorAction SilentlyContinue) { 'pwsh.exe' } else { 'powershell.exe' }
        if ($graphPsExe -eq 'powershell.exe' -and $script:lastBuildResult.Graph -match 'Get-UALGraph') {
            $controls.txtWarnings.Text = 'PowerShell 7 (pwsh) not found - Get-UALGraph is known to fail with BadRequest under Windows PowerShell 5.1. Install PowerShell 7 to run UAL Graph searches.'
        }
        $tempGraph = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(),
                     "M365_GUI_Graph_$([guid]::NewGuid().ToString('N')).ps1")
        Set-Content -Path $tempGraph -Value (Wrap-WithStaging -Body $script:lastBuildResult.Graph -Phase 'Graph' -RunStamp $runStamp) -Encoding UTF8
        Start-Process $graphPsExe -ArgumentList @('-NoExit','-ExecutionPolicy','Bypass','-File',$tempGraph)
    }
    if ($script:lastBuildResult.Exo) {
        $tempExo = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(),
                   "M365_GUI_EXO_$([guid]::NewGuid().ToString('N')).ps1")
        Set-Content -Path $tempExo -Value (Wrap-WithStaging -Body $script:lastBuildResult.Exo -Phase 'Exo' -RunStamp $runStamp) -Encoding UTF8
        Start-Process powershell.exe -ArgumentList @('-NoExit','-ExecutionPolicy','Bypass','-File',$tempExo)
    }
})

$controls.btnInstall.Add_Click({
    try {
        $tempScript = Join-Path $env:TEMP ("M365_App_Install_{0}.ps1" -f ([guid]::NewGuid().ToString('N')))
        Set-Content -Path $tempScript -Value $installerScriptText -Encoding UTF8

        $argList = @('-NoExit','-ExecutionPolicy','Bypass','-File',$tempScript,
                     '-ConfigRoot',('"{0}"' -f $ConfigRoot.TrimEnd('\')))
        if ($NewCert) { $argList += '-NewCert' }

        Start-Process powershell.exe -ArgumentList $argList
        $controls.txtErrors.Text = ''
        Set-Status 'Installer launched in a new PowerShell window. Click Refresh when it finishes.'
    }
    catch {
        $controls.txtErrors.Text = "Failed to launch installer window: $($_.Exception.Message)"
    }
})

$controls.btnUninstall.Add_Click({
    $idx = $controls.cmbTenant.SelectedIndex
    if ($idx -lt 0 -or $idx -ge $script:tenantConfigs.Count) {
        $controls.txtErrors.Text = 'Pick a tenant first.'
        return
    }
    $t = $script:tenantConfigs[$idx]
    if (-not $t.AppId -or -not $t.TenantId -or -not $t.CertificateThumbprint) {
        $controls.txtErrors.Text = 'Selected tenant is missing AppId / TenantId / CertificateThumbprint in its config.'
        return
    }
    $tenantFolder = Join-Path $ConfigRoot $t.FolderName
    $domainLabel  = if ($t.PrimaryDomain) { $t.PrimaryDomain } else { $t.FolderName }

    $confirm = [System.Windows.MessageBox]::Show(
        ("Uninstall the M365 App from tenant '{0}'?`r`n`r`nThis will:`r`n  - Delete the Azure AD application (AppId {1})`r`n  - Delete the local folder '{2}'`r`n`r`nThe service principal is left in place. This cannot be undone." -f $domainLabel, $t.AppId, $tenantFolder),
        'Confirm Uninstall',
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning)
    if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }

    $tenantFolderLit = $tenantFolder.Replace("'", "''")
    $appIdLit        = $t.AppId.Replace("'", "''")
    $tenantIdLit     = $t.TenantId.Replace("'", "''")
    $thumbLit        = $t.CertificateThumbprint.Replace("'", "''")
    $domainLit       = $domainLabel.Replace("'", "''")

    $uninstallScript = @"
`$ErrorActionPreference = 'Continue'
Write-Host '============================================================'
Write-Host "Uninstall M365 App - tenant '$domainLit'"
Write-Host '============================================================'

try {
    Get-Module Microsoft.Graph* | Remove-Module -Force -ErrorAction SilentlyContinue
    Import-Module Microsoft.Graph.Authentication -Force -ErrorAction Stop
    Import-Module Microsoft.Graph.Applications  -Force -ErrorAction Stop
}
catch {
    Write-Host "[ERROR] Failed to import Microsoft.Graph modules: `$(`$_.Exception.Message)" -ForegroundColor Red
    Read-Host 'Press Enter to close'
    return
}

Write-Host 'Connecting to Microsoft Graph (cert auth, app-only)...'
try {
    Connect-MgGraph -ClientId '$appIdLit' -TenantId '$tenantIdLit' -CertificateThumbprint '$thumbLit' -NoWelcome -ErrorAction Stop
    Write-Host '  [OK] Connected as the application (no user sign-in).' -ForegroundColor Green
}
catch {
    Write-Host "[ERROR] Connect-MgGraph (cert) failed: `$(`$_.Exception.Message)" -ForegroundColor Red
    Write-Host '  Cert auth requires Application.ReadWrite.All consent on the app.' -ForegroundColor Yellow
    Read-Host 'Press Enter to close'
    return
}

Write-Host "Looking up application by AppId $appIdLit..."
try {
    `$appResp = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/applications?`$filter=appId eq ''$appIdLit''' -ErrorAction Stop
    `$targetApps = @(`$appResp.value | Where-Object { `$_.appId -eq '$appIdLit' })
    if (`$targetApps.Count -eq 0) {
        Write-Host '[WARN] No application found with that AppId.' -ForegroundColor Yellow
    }
    elseif (`$targetApps.Count -gt 1) {
        Write-Host "[ABORT] Filter returned `$(`$targetApps.Count) applications for one AppId - refusing to delete (safety guard)." -ForegroundColor Red
    }
    else {
        `$app = `$targetApps[0]
        if (`$app.appId -ne '$appIdLit') {
            Write-Host "[ABORT] App `$(`$app.id) appId mismatch (`$(`$app.appId)) - refusing to delete." -ForegroundColor Red
        }
        else {
            Write-Host "Deleting application object `$(`$app.id)..."
            try {
                Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/applications/`$(`$app.id)" -ErrorAction Stop
                Write-Host "  [OK] Application deleted." -ForegroundColor Green
            }
            catch {
                Write-Host "[ERROR] Application delete failed: `$(`$_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
}
catch {
    Write-Host "[ERROR] Application lookup failed: `$(`$_.Exception.Message)" -ForegroundColor Red
}

try { Disconnect-MgGraph | Out-Null } catch { }

Write-Host ''
Write-Host "Removing local folder '$tenantFolderLit'..."
try {
    if (Test-Path -LiteralPath '$tenantFolderLit') {
        Remove-Item -LiteralPath '$tenantFolderLit' -Recurse -Force -ErrorAction Stop
        Write-Host '  [OK] Folder removed.' -ForegroundColor Green
    } else {
        Write-Host '  [INFO] Folder does not exist - nothing to remove.'
    }
}
catch {
    Write-Host "[ERROR] Folder removal failed: `$(`$_.Exception.Message)" -ForegroundColor Red
}

Write-Host ''
Write-Host 'Done. Return to the GUI and click Refresh.'
Read-Host 'Press Enter to close'
"@

    try {
        $tempScript = Join-Path $env:TEMP ("M365_App_Uninstall_{0}.ps1" -f ([guid]::NewGuid().ToString('N')))
        Set-Content -Path $tempScript -Value $uninstallScript -Encoding UTF8
        Start-Process powershell.exe -ArgumentList @('-NoExit','-ExecutionPolicy','Bypass','-File',$tempScript)
        $controls.txtErrors.Text = ''
        Set-Status "Uninstall launched for '$domainLabel'. Click Refresh when it finishes."
    }
    catch {
        $controls.txtErrors.Text = "Failed to launch uninstall window: $($_.Exception.Message)"
    }
})

$controls.btnSetupSelectAll.Add_Click({
    foreach ($key in $setupItems.Keys) {
        if ($controls[$key]) { $controls[$key].IsChecked = $true }
    }
})

$controls.btnSetupSelectNone.Add_Click({
    foreach ($key in $setupItems.Keys) {
        if ($controls[$key]) { $controls[$key].IsChecked = $false }
    }
})

$controls.btnSetupInstall.Add_Click({ Invoke-SetupAction -Action 'Install' })
$controls.btnSetupUpdate.Add_Click({ Invoke-SetupAction -Action 'Update' })
$controls.btnSetupUninstall.Add_Click({ Invoke-SetupAction -Action 'Uninstall' })


$controls.btnTestUal.Add_Click({
    $idx = $controls.cmbTenant.SelectedIndex
    if ($idx -lt 0 -or $idx -ge $script:tenantConfigs.Count) {
        $controls.txtErrors.Text = 'Pick a tenant first.'
        return
    }
    $t = $script:tenantConfigs[$idx]
    if (-not $t.AppId -or -not $t.CertificateThumbprint -or -not $t.PrimaryDomain) {
        $controls.txtErrors.Text = 'Selected tenant is missing AppId / Thumbprint / PrimaryDomain in its config.'
        return
    }

    $controls.txtErrors.Text = ''
    Set-Status "Connecting to Exchange Online for $($t.PrimaryDomain)..."

    Start-AsyncTask `
        -TaskArgs @{
            AppId      = [string]$t.AppId
            Thumbprint = [string]$t.CertificateThumbprint
            Domain     = [string]$t.PrimaryDomain
        } `
        -DisableButtons @('btnTestUal', 'btnUninstall', 'btnRefresh') `
        -Work {
            param($Sync, $TaskArgs)

            # EXO v3 lazy-loads cmdlets on first use; that step writes a progress bar
            # that can block a background runspace. Silence progress so the first
            # cmdlet call (Get-AdminAuditLogConfig) cannot hang on it.
            $ProgressPreference = 'SilentlyContinue'

            try { Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction Stop } catch { }

            try {
                $Sync.Status = 'Importing ExchangeOnlineManagement...'
                Import-Module ExchangeOnlineManagement -ErrorAction Stop
            }
            catch {
                $Sync.Error = "ExchangeOnlineManagement module not available: $($_.Exception.Message)"
                $Sync.Done = $true
                return
            }

            $connected = $false
            try {
                $Sync.Status = "Connecting to Exchange Online for $($TaskArgs.Domain)..."
                # -CommandName limits cmdlet generation to just what we need, so the
                #   first call doesn't trigger a large auto-module load (the hang point).
                # -DisableWAM avoids the WAM/window-handle path that newer modules
                #   (>=3.9.0) default to, which has no window handle in this runspace.
                Connect-ExchangeOnline -CertificateThumbprint $TaskArgs.Thumbprint -AppId $TaskArgs.AppId -Organization $TaskArgs.Domain -CommandName 'Get-AdminAuditLogConfig' -DisableWAM -ShowBanner:$false -ErrorAction Stop
                $connected = $true
            }
            catch {
                $Sync.Error = "Connect-ExchangeOnline failed: $($_.Exception.Message)"
                $Sync.Done = $true
                return
            }

            try {
                $Sync.Status = 'Querying audit log configuration...'
                $cfg = Get-AdminAuditLogConfig -ErrorAction Stop
                $Sync.Result = @{
                    Domain   = [string]$TaskArgs.Domain
                    Enabled  = [bool]$cfg.UnifiedAuditLogIngestionEnabled
                    OptInRaw = [string]$cfg.UnifiedAuditLogFirstOptInDate
                }
            }
            catch {
                $Sync.Error = "Get-AdminAuditLogConfig failed: $($_.Exception.Message)"
            }
            finally {
                if ($connected) {
                    $Sync.Status = 'Disconnecting...'
                    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
                }
                $Sync.Done = $true
            }
        } `
        -OnProgress {
            param($Sync)
            if ($Sync.Status) { Set-Status $Sync.Status }
        } `
        -OnComplete {
            param($Sync)
            if ($Sync.Error) {
                $controls.txtErrors.Text = $Sync.Error
                Set-Status ''
                return
            }
            if (-not $Sync.Result) { Set-Status ''; return }

            $domain  = [string]$Sync.Result.Domain
            $enabled = [bool]$Sync.Result.Enabled
            $statusText = if ($enabled) { 'ENABLED' } else { 'DISABLED' }

            $optInRaw = $Sync.Result.OptInRaw
            $optInDisplay = $null
            if (-not [string]::IsNullOrWhiteSpace([string]$optInRaw)) {
                try {
                    $optInDate = [DateTime]$optInRaw
                    $optInDisplay = $optInDate.ToString('yyyy-MM-dd HH:mm:ss')
                }
                catch {
                    $optInDisplay = [string]$optInRaw
                }
            }

            $msg = if ($enabled) {
                if ($optInDisplay) {
                    "Unified Audit Log is ENABLED for $domain.`r`nFirst opt-in date: $optInDisplay"
                } else {
                    "Unified Audit Log is ENABLED for $domain.`r`nFirst opt-in date: (not recorded)"
                }
            } else {
                $disabledMsg = "Unified Audit Log is DISABLED for $domain."
                if ($optInDisplay) {
                    $disabledMsg += "`r`nPrevious first opt-in date: $optInDisplay"
                }
                $disabledMsg += "`r`n`r`nTo enable, an Exchange admin can run:`r`n  Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled `$true"
                $disabledMsg
            }
            $icon = if ($enabled) { [System.Windows.MessageBoxImage]::Information } else { [System.Windows.MessageBoxImage]::Warning }
            [System.Windows.MessageBox]::Show($msg, 'UAL Status', [System.Windows.MessageBoxButton]::OK, $icon)

            $statusLine = "UAL for ${domain}: $statusText"
            if ($optInDisplay) { $statusLine += " (opt-in: $optInDisplay)" }
            Set-Status $statusLine
        }
})

$controls.btnExit.Add_Click({ $window.Close() })

# Grey/un-grey per-cmdlet option panels as their cmdlet checkboxes change.
# Checked/Unchecked also fire on programmatic changes, so Select Standard /
# Select None need no extra handling.
foreach ($depList in $script:expanderDeps.Values) {
    foreach ($cbName in $depList) {
        if ($controls[$cbName]) {
            $controls[$cbName].Add_Checked({ Update-ExpanderStates })
            $controls[$cbName].Add_Unchecked({ Update-ExpanderStates })
        }
    }
}

$controls.cbUalAll.Add_Checked({ Update-UalRangeFieldStates })
$controls.cbUalAll.Add_Unchecked({ Update-UalRangeFieldStates })

$window.Add_Closing({
    if ($script:AsyncJob) {
        $script:AsyncTimer.Stop()
        try { $script:AsyncJob.PowerShell.Stop() } catch { }
        try { $script:AsyncJob.Runspace.Close() } catch { }
        $script:AsyncJob = $null
    }
})

# ----------------------------------------
# Initial state
# ----------------------------------------
$script:tenantConfigs = @()
Update-TenantList
Update-ExpanderStates
Update-UalRangeFieldStates

[void]$window.ShowDialog()
