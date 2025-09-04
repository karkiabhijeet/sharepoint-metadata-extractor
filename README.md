# SharePoint Metadata Extractor Tool

A PowerShell solution for extracting Microsoft Purview Sensitivity Labels from SharePoint Online files. Perfect for compliance reporting and understanding your organization's data protection coverage.

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://github.com/PowerShell/PowerShell)
[![Graph API](https://img.shields.io/badge/Microsoft%20Graph-v1.0%2Bbeta-green.svg)](https://docs.microsoft.com/en-us/graph/)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

## What It Does

- Scans your SharePoint sites and finds files with sensitivity labels
- Creates easy-to-read CSV reports showing which files are protected
- Works with thousands of sites efficiently
- Handles authentication automatically
- Provides detailed progress tracking

## Prerequisites

### PowerShell Modules
- `Microsoft.Graph` (automatically installed)
- `ExchangeOnlineManagement` (automatically installed)

### API Permissions Required

**Microsoft Graph Application Permissions:**
- `Sites.Read.All` - Read SharePoint sites and document libraries
- `Files.Read.All` - Read files and extract sensitivity labels
- `InformationProtectionPolicy.Read.All` - Read sensitivity label catalog

## Quick Start

### 1. Clone the Repository
```bash
git clone https://github.com/karkiabhijeet/sharepoint-metadata-extractor.git
cd sharepoint-metadata-extractor
```

### 2. Configure Authentication
Copy the template and add your credentials:
```powershell
Copy-Item config.template.json config.json
```

Edit `config.json` and add:
- Your **Tenant ID**
- Your **Client ID** 
- Your **Client Secret**

### 3. Run the Tool
```powershell
.\Master-SharePoint-Extractor.ps1
```

The tool will guide you through:
1. Automatically discover all your SharePoint sites
2. Choose to scan ALL sites or select specific ones
3. Extract sensitivity labels with progress tracking
4. Generate consolidated CSV reports

## What You Get

The tool creates timestamped folders with:
- **Sites Inventory**: Complete list of all your SharePoint sites
- **Individual Site Reports**: Detailed results for each site
- **Consolidated Report**: Combined results from all processed sites
- **Labels Cache**: List of all sensitivity labels in your organization

## Troubleshooting

**Authentication Issues:** Make sure your app registration has the required permissions

**No Data Found:** Check that your sites actually contain files with sensitivity labels

**Performance:** For large organizations, the tool automatically optimizes for bulk operations

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

- **Issues**: [GitHub Issues](https://github.com/karkiabhijeet/sharepoint-metadata-extractor/issues)
- **Discussions**: [GitHub Discussions](https://github.com/karkiabhijeet/sharepoint-metadata-extractor/discussions)

---

**Built for Microsoft 365 Administrators and Compliance Teams**