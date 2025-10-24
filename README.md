# SharePoint Embedded ISV Setup Guide

This repository provides comprehensive step-by-step guidance for Independent Software Vendors (ISVs) to configure their SharePoint Embedded applications in customer tenants.

## Overview

SharePoint Embedded enables ISVs to build multi-tenant applications with dedicated file and document management capabilities. This guide walks through the complete process of setting up your application in a consuming tenant, including:

- Creating Azure App Registrations
- Configuring necessary permissions and API access
- Obtaining admin consent from consuming tenant administrators
- Registering the owning application using certificate-based authentication
- Creating SharePoint Embedded containers

## Repository Structure

```
├── docs/
│   ├── azure-app-registration.md    # Detailed Azure App Registration steps
│   ├── admin-consent.md             # Admin consent process
│   └── troubleshooting.md           # Common issues and solutions
├── scripts/
│   ├── Register-SPEOwningApp.ps1    # Register owning app in consuming tenant
│   └── New-SPEContainer.ps1         # Create SharePoint Embedded container
└── README.md                        # This file
```

## Quick Start

### Prerequisites

Before you begin, ensure you have:

1. **For ISV (Owning Tenant)**:
   - Azure subscription and access to create App Registrations
   - Global Administrator or Application Administrator role
   - PowerShell 7.x or later
   - Required PowerShell modules (see [Installation](#installation))

2. **For Customer (Consuming Tenant)**:
   - SharePoint Administrator or Global Administrator role
   - Ability to grant admin consent to applications
   - PowerShell 7.x or later

### Installation

Install required PowerShell modules:

```powershell
# Install PnP PowerShell for SharePoint operations
Install-Module -Name PnP.PowerShell -Force -AllowClobber

# Install Microsoft Graph PowerShell for Azure AD operations
Install-Module -Name Microsoft.Graph -Force -AllowClobber
```

## Step-by-Step Guide

### Step 1: Create Azure App Registration (ISV)

Follow the detailed instructions in [docs/azure-app-registration.md](docs/azure-app-registration.md) to:
- Create a new App Registration in Azure Portal
- Configure API permissions
- Create and upload a certificate for authentication
- Configure redirect URIs and application settings

### Step 2: Obtain Admin Consent (Customer)

Follow the instructions in [docs/admin-consent.md](docs/admin-consent.md) to:
- Generate an admin consent URL
- Have the customer tenant admin grant consent
- Verify consent was granted successfully

### Step 3: Register Owning App in Consuming Tenant

Use the provided PowerShell script to register your owning application:

```powershell
.\scripts\Register-SPEOwningApp.ps1 `
    -OwningAppId "your-app-id" `
    -CertificatePath ".\certificate.pfx" `
    -CertificatePassword "your-password" `
    -ConsumingTenantId "customer-tenant-id"
```

### Step 4: Create SharePoint Embedded Container

Once the owning app is registered, create containers:

```powershell
.\scripts\New-SPEContainer.ps1 `
    -ContainerTypeId "your-container-type-id" `
    -DisplayName "My Container" `
    -OwningAppId "your-app-id"
```

## Documentation

- **[Azure App Registration Guide](docs/azure-app-registration.md)** - Complete walkthrough of creating and configuring the Azure App Registration
- **[Admin Consent Process](docs/admin-consent.md)** - How to obtain and verify admin consent from customer tenants
- **[Troubleshooting Guide](docs/troubleshooting.md)** - Common issues and their solutions

## Scripts

All PowerShell scripts are located in the `scripts/` directory:

- **Register-SPEOwningApp.ps1** - Automates the registration of your owning application in a consuming tenant using certificate-based authentication
- **New-SPEContainer.ps1** - Creates a new SharePoint Embedded container in the consuming tenant

## Support and Resources

- [SharePoint Embedded Documentation](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/overview)
- [Azure App Registration Documentation](https://learn.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
- [Microsoft Graph API Reference](https://learn.microsoft.com/en-us/graph/api/overview)

## Contributing

If you find issues or have suggestions for improvements, please open an issue in this repository.

## License

This project is provided as-is for guidance purposes.
