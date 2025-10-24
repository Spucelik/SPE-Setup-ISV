<#
.SYNOPSIS
    Registers a SharePoint Embedded owning application in a consuming tenant.

.DESCRIPTION
    This script registers an ISV's SharePoint Embedded owning application in a customer's 
    (consuming) tenant using certificate-based authentication. It creates the necessary 
    container type registration and configures the required settings.

.PARAMETER OwningAppId
    The Application (Client) ID of the ISV's owning application from Azure AD.

.PARAMETER CertificatePath
    Path to the PFX certificate file used for authentication.

.PARAMETER CertificatePassword
    Password for the PFX certificate file.

.PARAMETER CertificateThumbprint
    Alternative to CertificatePath - use a certificate already in the cert store.

.PARAMETER ConsumingTenantId
    The Tenant ID or tenant domain name (e.g., customer.onmicrosoft.com) of the consuming tenant.

.PARAMETER ConsumingTenantAdminUrl
    The SharePoint admin URL of the consuming tenant (e.g., https://customer-admin.sharepoint.com).

.PARAMETER ContainerTypeDisplayName
    Display name for the container type. If not specified, uses the app name.

.PARAMETER ContainerTypeDescription
    Description for the container type.

.PARAMETER ApplicationPermissions
    Comma-separated list of application permissions to request.
    Default: "Sites.FullControl.All,Files.ReadWrite.All"

.EXAMPLE
    .\Register-SPEOwningApp.ps1 `
        -OwningAppId "12345678-1234-1234-1234-123456789abc" `
        -CertificatePath ".\certificate.pfx" `
        -CertificatePassword "MyPassword123!" `
        -ConsumingTenantId "customer.onmicrosoft.com" `
        -ConsumingTenantAdminUrl "https://customer-admin.sharepoint.com"

.EXAMPLE
    .\Register-SPEOwningApp.ps1 `
        -OwningAppId "12345678-1234-1234-1234-123456789abc" `
        -CertificateThumbprint "ABCDEF1234567890ABCDEF1234567890ABCDEF12" `
        -ConsumingTenantId "abcd1234-5678-90ab-cdef-1234567890ab" `
        -ConsumingTenantAdminUrl "https://customer-admin.sharepoint.com" `
        -ContainerTypeDisplayName "My Custom Container Type"

.NOTES
    Author: SharePoint Embedded ISV Setup Guide
    Version: 1.0
    Prerequisites:
        - PnP.PowerShell module installed
        - Microsoft.Graph module installed
        - SharePoint Administrator or Global Administrator role in consuming tenant
        - Admin consent already granted in the consuming tenant
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Application (Client) ID of the owning app")]
    [ValidateNotNullOrEmpty()]
    [string]$OwningAppId,

    [Parameter(Mandatory = $false, HelpMessage = "Path to PFX certificate file")]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CertificatePath,

    [Parameter(Mandatory = $false, HelpMessage = "Password for PFX certificate")]
    [string]$CertificatePassword,

    [Parameter(Mandatory = $false, HelpMessage = "Certificate thumbprint from cert store")]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory = $true, HelpMessage = "Consuming tenant ID or domain")]
    [ValidateNotNullOrEmpty()]
    [string]$ConsumingTenantId,

    [Parameter(Mandatory = $true, HelpMessage = "SharePoint admin URL of consuming tenant")]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern("^https://.*-admin\.sharepoint\.com$")]
    [string]$ConsumingTenantAdminUrl,

    [Parameter(Mandatory = $false, HelpMessage = "Display name for container type")]
    [string]$ContainerTypeDisplayName,

    [Parameter(Mandatory = $false, HelpMessage = "Description for container type")]
    [string]$ContainerTypeDescription = "SharePoint Embedded container type",

    [Parameter(Mandatory = $false, HelpMessage = "Application permissions (comma-separated)")]
    [string]$ApplicationPermissions = "Sites.FullControl.All,Files.ReadWrite.All"
)

# Ensure we have certificate authentication configured
if (-not $CertificatePath -and -not $CertificateThumbprint) {
    Write-Error "You must provide either -CertificatePath or -CertificateThumbprint"
    exit 1
}

# Function to check if required modules are installed
function Test-RequiredModules {
    Write-Host "Checking required PowerShell modules..." -ForegroundColor Cyan
    
    $requiredModules = @('PnP.PowerShell', 'Microsoft.Graph.Authentication')
    $missingModules = @()
    
    foreach ($module in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            $missingModules += $module
            Write-Warning "Module '$module' is not installed"
        } else {
            Write-Host "✓ Module '$module' is installed" -ForegroundColor Green
        }
    }
    
    if ($missingModules.Count -gt 0) {
        Write-Host "`nTo install missing modules, run:" -ForegroundColor Yellow
        foreach ($module in $missingModules) {
            Write-Host "  Install-Module -Name $module -Force -AllowClobber" -ForegroundColor Yellow
        }
        return $false
    }
    
    return $true
}

# Function to connect to SharePoint using certificate authentication
function Connect-SharePointWithCert {
    param(
        [string]$Url,
        [string]$ClientId,
        [string]$TenantId,
        [string]$CertPath,
        [string]$CertPassword,
        [string]$Thumbprint
    )
    
    Write-Host "Connecting to SharePoint: $Url" -ForegroundColor Cyan
    
    try {
        if ($Thumbprint) {
            # Connect using certificate from store
            Connect-PnPOnline -Url $Url `
                -ClientId $ClientId `
                -Tenant $TenantId `
                -Thumbprint $Thumbprint `
                -ErrorAction Stop
        } else {
            # Connect using PFX file
            $securePassword = ConvertTo-SecureString -String $CertPassword -AsPlainText -Force
            Connect-PnPOnline -Url $Url `
                -ClientId $ClientId `
                -Tenant $TenantId `
                -CertificatePath $CertPath `
                -CertificatePassword $securePassword `
                -ErrorAction Stop
        }
        
        Write-Host "✓ Successfully connected to SharePoint" -ForegroundColor Green
        return $true
    } catch {
        Write-Error "Failed to connect to SharePoint: $($_.Exception.Message)"
        Write-Host "`nTroubleshooting tips:" -ForegroundColor Yellow
        Write-Host "  1. Verify admin consent has been granted in the consuming tenant"
        Write-Host "  2. Check that the certificate is valid and not expired"
        Write-Host "  3. Ensure the certificate thumbprint matches what's in Azure AD"
        Write-Host "  4. Verify you have SharePoint Administrator role"
        return $false
    }
}

# Function to check if the owning app exists in the tenant
function Test-OwningAppExists {
    param([string]$AppId)
    
    Write-Host "Verifying owning app exists in consuming tenant..." -ForegroundColor Cyan
    
    try {
        # Try to get the service principal
        $sp = Get-MgServicePrincipal -Filter "appId eq '$AppId'" -ErrorAction SilentlyContinue
        
        if ($sp) {
            Write-Host "✓ Owning app found in tenant: $($sp.DisplayName)" -ForegroundColor Green
            Write-Host "  App ID: $($sp.AppId)" -ForegroundColor Gray
            Write-Host "  Object ID: $($sp.Id)" -ForegroundColor Gray
            return $true
        } else {
            Write-Warning "Owning app not found in consuming tenant"
            Write-Host "  This usually means admin consent has not been granted yet." -ForegroundColor Yellow
            Write-Host "  Please ensure the tenant administrator has granted consent first." -ForegroundColor Yellow
            return $false
        }
    } catch {
        Write-Warning "Unable to verify owning app: $($_.Exception.Message)"
        return $false
    }
}

# Function to register container type
function Register-ContainerType {
    param(
        [string]$AppId,
        [string]$DisplayName,
        [string]$Description
    )
    
    Write-Host "`nRegistering container type..." -ForegroundColor Cyan
    
    try {
        # Check if container type already exists
        $existingContainerType = Get-PnPContainerType -ErrorAction SilentlyContinue | 
            Where-Object { $_.OwningAppId -eq $AppId }
        
        if ($existingContainerType) {
            Write-Host "✓ Container type already registered" -ForegroundColor Green
            Write-Host "  Container Type ID: $($existingContainerType.ContainerTypeId)" -ForegroundColor Gray
            Write-Host "  Display Name: $($existingContainerType.DisplayName)" -ForegroundColor Gray
            Write-Host "  Owning App ID: $($existingContainerType.OwningAppId)" -ForegroundColor Gray
            return $existingContainerType
        }
        
        # Register new container type
        Write-Host "Creating new container type..." -ForegroundColor Cyan
        
        if (-not $DisplayName) {
            $DisplayName = "Container Type for App $AppId"
        }
        
        $containerType = Add-PnPContainerType `
            -OwningApplicationId $AppId `
            -DisplayName $DisplayName `
            -ErrorAction Stop
        
        Write-Host "✓ Container type registered successfully!" -ForegroundColor Green
        Write-Host "  Container Type ID: $($containerType.ContainerTypeId)" -ForegroundColor Gray
        Write-Host "  Display Name: $($containerType.DisplayName)" -ForegroundColor Gray
        Write-Host "  Owning App ID: $($containerType.OwningAppId)" -ForegroundColor Gray
        
        return $containerType
        
    } catch {
        Write-Error "Failed to register container type: $($_.Exception.Message)"
        
        if ($_.Exception.Message -like "*already exists*") {
            Write-Host "`nThe container type may already exist. Attempting to retrieve it..." -ForegroundColor Yellow
            $containerType = Get-PnPContainerType | Where-Object { $_.OwningAppId -eq $AppId }
            if ($containerType) {
                Write-Host "✓ Found existing container type" -ForegroundColor Green
                return $containerType
            }
        }
        
        Write-Host "`nTroubleshooting tips:" -ForegroundColor Yellow
        Write-Host "  1. Verify the owning app has proper permissions"
        Write-Host "  2. Check that SharePoint Embedded is enabled in this tenant"
        Write-Host "  3. Ensure you have SharePoint Administrator privileges"
        
        return $null
    }
}

# Function to display summary
function Show-Summary {
    param(
        [object]$ContainerType,
        [string]$TenantId,
        [string]$AppId
    )
    
    Write-Host "`n" + ("=" * 70) -ForegroundColor Cyan
    Write-Host "  REGISTRATION SUMMARY" -ForegroundColor Cyan
    Write-Host ("=" * 70) -ForegroundColor Cyan
    
    Write-Host "`nOwning App Information:" -ForegroundColor Green
    Write-Host "  Application ID:    $AppId"
    Write-Host "  Consuming Tenant:  $TenantId"
    
    if ($ContainerType) {
        Write-Host "`nContainer Type Information:" -ForegroundColor Green
        Write-Host "  Container Type ID: $($ContainerType.ContainerTypeId)"
        Write-Host "  Display Name:      $($ContainerType.DisplayName)"
        Write-Host "  Description:       $($ContainerType.Description)"
        Write-Host "  Owning App ID:     $($ContainerType.OwningAppId)"
    }
    
    Write-Host "`nNext Steps:" -ForegroundColor Yellow
    Write-Host "  1. Save the Container Type ID for creating containers"
    Write-Host "  2. Use the New-SPEContainer.ps1 script to create containers"
    Write-Host "  3. Test container creation and access"
    
    Write-Host "`n" + ("=" * 70) -ForegroundColor Cyan
}

# Main execution
try {
    Write-Host "`n" + ("=" * 70) -ForegroundColor Cyan
    Write-Host "  SharePoint Embedded Owning App Registration" -ForegroundColor Cyan
    Write-Host ("=" * 70) -ForegroundColor Cyan
    
    # Step 1: Check required modules
    Write-Host "`n[Step 1/4] Checking prerequisites..." -ForegroundColor Cyan
    if (-not (Test-RequiredModules)) {
        Write-Error "Missing required modules. Please install them and try again."
        exit 1
    }
    
    # Import modules
    Import-Module PnP.PowerShell -ErrorAction Stop
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    
    # Step 2: Connect to Microsoft Graph to verify app
    Write-Host "`n[Step 2/4] Verifying owning app..." -ForegroundColor Cyan
    
    $graphConnectParams = @{
        ClientId = $OwningAppId
        TenantId = $ConsumingTenantId
    }
    
    if ($CertificateThumbprint) {
        $graphConnectParams.CertificateThumbprint = $CertificateThumbprint
    } else {
        $securePassword = ConvertTo-SecureString -String $CertificatePassword -AsPlainText -Force
        $graphConnectParams.CertificatePath = $CertificatePath
        $graphConnectParams.CertificatePassword = $securePassword
    }
    
    try {
        Connect-MgGraph @graphConnectParams -NoWelcome -ErrorAction Stop
        
        # Verify app exists
        if (Test-OwningAppExists -AppId $OwningAppId) {
            Write-Host "✓ Owning app verification complete" -ForegroundColor Green
        } else {
            Write-Error "Owning app not found. Please ensure admin consent is granted first."
            Disconnect-MgGraph
            exit 1
        }
        
        Disconnect-MgGraph
    } catch {
        Write-Error "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
        Write-Host "`nPlease verify:" -ForegroundColor Yellow
        Write-Host "  - Certificate is valid and not expired"
        Write-Host "  - Certificate thumbprint is correct"
        Write-Host "  - App has been granted admin consent"
        exit 1
    }
    
    # Step 3: Connect to SharePoint
    Write-Host "`n[Step 3/4] Connecting to SharePoint..." -ForegroundColor Cyan
    
    $connectParams = @{
        Url = $ConsumingTenantAdminUrl
        ClientId = $OwningAppId
        TenantId = $ConsumingTenantId
    }
    
    if ($CertificateThumbprint) {
        $connectParams.Thumbprint = $CertificateThumbprint
    } else {
        $connectParams.CertPath = $CertificatePath
        $connectParams.CertPassword = $CertificatePassword
    }
    
    if (-not (Connect-SharePointWithCert @connectParams)) {
        Write-Error "Failed to connect to SharePoint. Registration cannot continue."
        exit 1
    }
    
    # Step 4: Register container type
    Write-Host "`n[Step 4/4] Registering container type..." -ForegroundColor Cyan
    
    $containerType = Register-ContainerType `
        -AppId $OwningAppId `
        -DisplayName $ContainerTypeDisplayName `
        -Description $ContainerTypeDescription
    
    if (-not $containerType) {
        Write-Error "Container type registration failed"
        Disconnect-PnPOnline
        exit 1
    }
    
    # Disconnect
    Disconnect-PnPOnline
    
    # Show summary
    Show-Summary -ContainerType $containerType -TenantId $ConsumingTenantId -AppId $OwningAppId
    
    Write-Host "`n✓ Registration completed successfully!" -ForegroundColor Green
    
} catch {
    Write-Error "An unexpected error occurred: $($_.Exception.Message)"
    Write-Host "`nStack Trace:" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace
    
    # Clean up connections
    try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch { }
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch { }
    
    exit 1
}
