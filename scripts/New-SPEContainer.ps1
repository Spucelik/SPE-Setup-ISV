<#
.SYNOPSIS
    Creates a new SharePoint Embedded container in a consuming tenant.

.DESCRIPTION
    This script creates a SharePoint Embedded container in a customer's tenant using
    certificate-based authentication. The owning application must already be registered
    in the consuming tenant before running this script.

.PARAMETER ContainerTypeId
    The Container Type ID that was created when registering the owning app.

.PARAMETER DisplayName
    Display name for the new container.

.PARAMETER Description
    Description for the new container.

.PARAMETER OwningAppId
    The Application (Client) ID of the owning application.

.PARAMETER CertificatePath
    Path to the PFX certificate file used for authentication.

.PARAMETER CertificatePassword
    Password for the PFX certificate file.

.PARAMETER CertificateThumbprint
    Alternative to CertificatePath - use a certificate already in the cert store.

.PARAMETER ConsumingTenantId
    The Tenant ID or tenant domain name of the consuming tenant.

.PARAMETER SiteUrl
    The SharePoint site URL to use for connection. Can be any site in the tenant.
    If not specified, uses the root site collection.

.PARAMETER SetPermissions
    If specified, sets initial permissions on the container.

.PARAMETER Owners
    Comma-separated list of user principal names to set as container owners.

.PARAMETER Members
    Comma-separated list of user principal names to set as container members.

.PARAMETER Readers
    Comma-separated list of user principal names to set as container readers.

.EXAMPLE
    .\New-SPEContainer.ps1 `
        -ContainerTypeId "b!ISV.default|12345678-1234-1234-1234-123456789abc" `
        -DisplayName "Customer Project Container" `
        -OwningAppId "12345678-1234-1234-1234-123456789abc" `
        -CertificatePath ".\certificate.pfx" `
        -CertificatePassword "MyPassword123!" `
        -ConsumingTenantId "customer.onmicrosoft.com"

.EXAMPLE
    .\New-SPEContainer.ps1 `
        -ContainerTypeId "b!ISV.default|12345678-1234-1234-1234-123456789abc" `
        -DisplayName "Customer Project Container" `
        -Description "Container for customer project files" `
        -OwningAppId "12345678-1234-1234-1234-123456789abc" `
        -CertificateThumbprint "ABCDEF1234567890ABCDEF1234567890ABCDEF12" `
        -ConsumingTenantId "customer.onmicrosoft.com" `
        -SetPermissions `
        -Owners "user1@customer.com,user2@customer.com" `
        -Members "user3@customer.com"

.NOTES
    Author: SharePoint Embedded ISV Setup Guide
    Version: 1.0
    Prerequisites:
        - PnP.PowerShell module installed
        - Owning app already registered in the consuming tenant
        - SharePoint Administrator or Global Administrator role
        - Container Type must already exist
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Container Type ID from registration")]
    [ValidateNotNullOrEmpty()]
    [string]$ContainerTypeId,

    [Parameter(Mandatory = $true, HelpMessage = "Display name for the container")]
    [ValidateNotNullOrEmpty()]
    [string]$DisplayName,

    [Parameter(Mandatory = $false, HelpMessage = "Description for the container")]
    [string]$Description,

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

    [Parameter(Mandatory = $false, HelpMessage = "SharePoint site URL for connection")]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false, HelpMessage = "Set initial permissions on container")]
    [switch]$SetPermissions,

    [Parameter(Mandatory = $false, HelpMessage = "Comma-separated list of owner UPNs")]
    [string]$Owners,

    [Parameter(Mandatory = $false, HelpMessage = "Comma-separated list of member UPNs")]
    [string]$Members,

    [Parameter(Mandatory = $false, HelpMessage = "Comma-separated list of reader UPNs")]
    [string]$Readers
)

# Ensure we have certificate authentication configured
if (-not $CertificatePath -and -not $CertificateThumbprint) {
    Write-Error "You must provide either -CertificatePath or -CertificateThumbprint"
    exit 1
}

# Function to check if required modules are installed
function Test-RequiredModules {
    Write-Host "Checking required PowerShell modules..." -ForegroundColor Cyan
    
    $requiredModules = @('PnP.PowerShell')
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
        Write-Host "  1. Verify the owning app is registered in the consuming tenant"
        Write-Host "  2. Check that the certificate is valid and not expired"
        Write-Host "  3. Ensure the certificate thumbprint matches what's in Azure AD"
        Write-Host "  4. Verify you have SharePoint Administrator role"
        return $false
    }
}

# Function to verify container type exists
function Test-ContainerTypeExists {
    param([string]$TypeId)
    
    Write-Host "Verifying container type exists..." -ForegroundColor Cyan
    
    try {
        $containerType = Get-PnPContainerType -ContainerTypeId $TypeId -ErrorAction SilentlyContinue
        
        if ($containerType) {
            Write-Host "✓ Container type found" -ForegroundColor Green
            Write-Host "  Container Type ID: $($containerType.ContainerTypeId)" -ForegroundColor Gray
            Write-Host "  Display Name: $($containerType.DisplayName)" -ForegroundColor Gray
            Write-Host "  Owning App ID: $($containerType.OwningAppId)" -ForegroundColor Gray
            return $true
        } else {
            Write-Warning "Container type not found: $TypeId"
            Write-Host "  Please ensure the owning app is registered first using Register-SPEOwningApp.ps1" -ForegroundColor Yellow
            return $false
        }
    } catch {
        Write-Warning "Unable to verify container type: $($_.Exception.Message)"
        return $false
    }
}

# Function to create container
function New-Container {
    param(
        [string]$TypeId,
        [string]$Name,
        [string]$Desc
    )
    
    Write-Host "`nCreating SharePoint Embedded container..." -ForegroundColor Cyan
    Write-Host "  Display Name: $Name" -ForegroundColor Gray
    
    try {
        # Create the container
        $containerParams = @{
            ContainerTypeId = $TypeId
            DisplayName = $Name
        }
        
        if ($Desc) {
            $containerParams.Description = $Desc
        }
        
        $container = New-PnPContainer @containerParams -ErrorAction Stop
        
        Write-Host "✓ Container created successfully!" -ForegroundColor Green
        Write-Host "  Container ID: $($container.ContainerId)" -ForegroundColor Gray
        Write-Host "  Display Name: $($container.DisplayName)" -ForegroundColor Gray
        Write-Host "  Container Type ID: $($container.ContainerTypeId)" -ForegroundColor Gray
        Write-Host "  Created: $($container.CreatedDateTime)" -ForegroundColor Gray
        
        return $container
        
    } catch {
        Write-Error "Failed to create container: $($_.Exception.Message)"
        
        if ($_.Exception.Message -like "*not found*") {
            Write-Host "`nThe container type may not be registered." -ForegroundColor Yellow
            Write-Host "Please run Register-SPEOwningApp.ps1 first." -ForegroundColor Yellow
        }
        
        if ($_.Exception.Message -like "*unauthorized*" -or $_.Exception.Message -like "*forbidden*") {
            Write-Host "`nPermission denied. Please ensure:" -ForegroundColor Yellow
            Write-Host "  1. The owning app is registered in the tenant"
            Write-Host "  2. You have SharePoint Administrator privileges"
            Write-Host "  3. Admin consent has been granted"
        }
        
        return $null
    }
}

# Function to set container permissions
function Set-ContainerPermissions {
    param(
        [string]$ContainerId,
        [string]$OwnerList,
        [string]$MemberList,
        [string]$ReaderList
    )
    
    Write-Host "`nSetting container permissions..." -ForegroundColor Cyan
    
    try {
        # Set owners
        if ($OwnerList) {
            $ownerArray = $OwnerList -split ',' | ForEach-Object { $_.Trim() }
            foreach ($owner in $ownerArray) {
                try {
                    Set-PnPContainerPermission -ContainerId $ContainerId `
                        -UserPrincipalName $owner `
                        -Role "Owner" `
                        -ErrorAction Stop
                    Write-Host "  ✓ Added owner: $owner" -ForegroundColor Green
                } catch {
                    Write-Warning "  Failed to add owner '$owner': $($_.Exception.Message)"
                }
            }
        }
        
        # Set members
        if ($MemberList) {
            $memberArray = $MemberList -split ',' | ForEach-Object { $_.Trim() }
            foreach ($member in $memberArray) {
                try {
                    Set-PnPContainerPermission -ContainerId $ContainerId `
                        -UserPrincipalName $member `
                        -Role "Member" `
                        -ErrorAction Stop
                    Write-Host "  ✓ Added member: $member" -ForegroundColor Green
                } catch {
                    Write-Warning "  Failed to add member '$member': $($_.Exception.Message)"
                }
            }
        }
        
        # Set readers
        if ($ReaderList) {
            $readerArray = $ReaderList -split ',' | ForEach-Object { $_.Trim() }
            foreach ($reader in $readerArray) {
                try {
                    Set-PnPContainerPermission -ContainerId $ContainerId `
                        -UserPrincipalName $reader `
                        -Role "Reader" `
                        -ErrorAction Stop
                    Write-Host "  ✓ Added reader: $reader" -ForegroundColor Green
                } catch {
                    Write-Warning "  Failed to add reader '$reader': $($_.Exception.Message)"
                }
            }
        }
        
        Write-Host "✓ Permissions configured" -ForegroundColor Green
        
    } catch {
        Write-Warning "Error setting permissions: $($_.Exception.Message)"
        Write-Host "You can set permissions later using Set-PnPContainerPermission cmdlet" -ForegroundColor Yellow
    }
}

# Function to display container details
function Show-ContainerDetails {
    param(
        [object]$Container
    )
    
    Write-Host "`n" + ("=" * 70) -ForegroundColor Cyan
    Write-Host "  CONTAINER CREATED SUCCESSFULLY" -ForegroundColor Cyan
    Write-Host ("=" * 70) -ForegroundColor Cyan
    
    Write-Host "`nContainer Information:" -ForegroundColor Green
    Write-Host "  Container ID:       $($Container.ContainerId)"
    Write-Host "  Display Name:       $($Container.DisplayName)"
    
    if ($Container.Description) {
        Write-Host "  Description:        $($Container.Description)"
    }
    
    Write-Host "  Container Type ID:  $($Container.ContainerTypeId)"
    Write-Host "  Created:            $($Container.CreatedDateTime)"
    
    if ($Container.WebUrl) {
        Write-Host "  Web URL:            $($Container.WebUrl)"
    }
    
    Write-Host "`nUsing the Container:" -ForegroundColor Yellow
    Write-Host "  1. Use the Container ID in your application"
    Write-Host "  2. Access via Microsoft Graph API:"
    Write-Host "     GET https://graph.microsoft.com/v1.0/storage/fileStorage/containers/$($Container.ContainerId)"
    Write-Host "  3. Manage permissions using Set-PnPContainerPermission cmdlet"
    
    Write-Host "`nNext Steps:" -ForegroundColor Yellow
    Write-Host "  - Upload files to the container using your application"
    Write-Host "  - Set additional permissions as needed"
    Write-Host "  - Test accessing the container from your application"
    
    Write-Host "`n" + ("=" * 70) -ForegroundColor Cyan
}

# Main execution
try {
    Write-Host "`n" + ("=" * 70) -ForegroundColor Cyan
    Write-Host "  SharePoint Embedded Container Creation" -ForegroundColor Cyan
    Write-Host ("=" * 70) -ForegroundColor Cyan
    
    # Step 1: Check required modules
    Write-Host "`n[Step 1/4] Checking prerequisites..." -ForegroundColor Cyan
    if (-not (Test-RequiredModules)) {
        Write-Error "Missing required modules. Please install them and try again."
        exit 1
    }
    
    # Import modules
    Import-Module PnP.PowerShell -ErrorAction Stop
    
    # Determine site URL if not provided
    if (-not $SiteUrl) {
        # Extract tenant name from tenant ID
        if ($ConsumingTenantId -like "*.onmicrosoft.com") {
            $tenantName = $ConsumingTenantId -replace '\.onmicrosoft\.com.*$', ''
        } else {
            # If tenant ID is a GUID, we need the tenant name
            Write-Host "Note: Using root site collection. Specify -SiteUrl if you need a different site." -ForegroundColor Yellow
            # For now, we'll try to connect to the admin center which is required anyway
            $SiteUrl = "https://$ConsumingTenantId.sharepoint.com"
        }
        
        if (-not ($SiteUrl -like "https://*")) {
            $SiteUrl = "https://$tenantName.sharepoint.com"
        }
    }
    
    # Step 2: Connect to SharePoint
    Write-Host "`n[Step 2/4] Connecting to SharePoint..." -ForegroundColor Cyan
    
    $connectParams = @{
        Url = $SiteUrl
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
        Write-Error "Failed to connect to SharePoint. Cannot create container."
        exit 1
    }
    
    # Step 3: Verify container type exists
    Write-Host "`n[Step 3/4] Verifying container type..." -ForegroundColor Cyan
    
    if (-not (Test-ContainerTypeExists -TypeId $ContainerTypeId)) {
        Write-Error "Container type not found. Please register the owning app first."
        Disconnect-PnPOnline
        exit 1
    }
    
    # Step 4: Create container
    Write-Host "`n[Step 4/4] Creating container..." -ForegroundColor Cyan
    
    $container = New-Container -TypeId $ContainerTypeId -Name $DisplayName -Desc $Description
    
    if (-not $container) {
        Write-Error "Container creation failed"
        Disconnect-PnPOnline
        exit 1
    }
    
    # Set permissions if requested
    if ($SetPermissions -and ($Owners -or $Members -or $Readers)) {
        Set-ContainerPermissions -ContainerId $container.ContainerId `
            -OwnerList $Owners `
            -MemberList $MemberList `
            -ReaderList $Readers
    }
    
    # Disconnect
    Disconnect-PnPOnline
    
    # Show container details
    Show-ContainerDetails -Container $container
    
    Write-Host "`n✓ Container creation completed successfully!" -ForegroundColor Green
    
} catch {
    Write-Error "An unexpected error occurred: $($_.Exception.Message)"
    Write-Host "`nStack Trace:" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace
    
    # Clean up connections
    try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch { }
    
    exit 1
}
