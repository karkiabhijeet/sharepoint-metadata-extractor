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

**For App Registration (Recommended):**
- `Sites.Read.All` - Read SharePoint sites and document libraries
- `Files.Read.All` - Read files and extract sensitivity labels
- `InformationProtectionPolicy.Read.All` - Read sensitivity label catalog

**For Interactive Login:**
If using interactive authentication, ensure your user account has:
- **SharePoint Administrator** role (to access all sites)
- **Compliance Administrator** role (to read sensitivity labels)
- **Global Reader** role (minimum read access)
- Or site-specific permissions if scanning only certain sites

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
1. **Automatically discover all your SharePoint sites** - Gets complete inventory
2. **Choose extraction mode:**
   - **Option 1 - Extract ALL sites**: Processes every site automatically (best for complete audits)
   - **Option 2 - Extract SPECIFIC sites**: Lets you review and select which sites to process (best for targeted scans)
3. **Extract sensitivity labels with progress tracking** - Shows real-time progress
4. **Generate consolidated CSV reports** - Creates easy-to-analyze reports

## What You Get

The tool creates timestamped folders with these files:

### 1. Sites Inventory (`SharePoint-Sites-Inventory_TIMESTAMP.csv`)
Complete list of all SharePoint sites with columns:
- `SiteName` - Display name of the SharePoint site
- `SiteUrl` - Full URL to the site
- `Extract` - Whether this site will be processed (True/False)
- `Status` - Processing status (Pending/Completed/Error)
- `LastScanned` - When the site was last processed

### 2. Individual Site Reports (`Site_[SiteName]_TIMESTAMP.csv`)
Detailed results for each processed site with columns:
- `TenantId` - Your Azure AD tenant identifier
- `SiteName` - SharePoint site name
- `SiteUrl` - Full site URL
- `LibraryName` - Document library name (e.g., "Documents")
- `FileName` - Name of the file
- `FileExtension` - File type (docx, xlsx, pdf, etc.)
- `FileSizeBytes` - File size in bytes
- `CreatedBy` - Who created the file
- `CreatedDateTime` - When file was created
- `LastModifiedDateTime` - When file was last modified
- `LabelIds` - Sensitivity label unique identifiers
- `LabelNames` - Human-readable label names (e.g., "Confidential")
- `AssignmentMethods` - How labels were applied (manual, automatic, etc.)
- `ScanDateTime` - When this scan was performed

### 3. Consolidated Report (`Consolidated_SharePoint_SensitivityLabels_TIMESTAMP.csv`)
Combined results from all processed sites - same columns as individual reports but includes data from all sites

### 4. Labels Cache (`Sensitivity-Labels-Cache.json`)
List of all sensitivity labels in your organization with their IDs and display names

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