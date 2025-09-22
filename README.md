# SharePoint Metadata Extractor Tool #



A comprehensive, enterprise-grade PowerShell solution for extracting Microsoft Purview Sensitivity Labels from SharePoint Online files at scale. Built for organizations with thousands of sites and millions of files.



[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://github.com/PowerShell/PowerShell)A comprehensive, enterprise-grade PowerShell solution for extracting Microsoft Purview Sensitivity Labels from SharePoint Online files at scale. Built for organizations with thousands of sites and millions of files.A comprehensive PowerShell solution for enumerating SharePoint Online files and reporting their applied Microsoft Purview Sensitivity Labels using the Microsoft Graph API.

[![Graph API](https://img.shields.io/badge/Microsoft%20Graph-v1.0%2Bbeta-green.svg)](https://docs.microsoft.com/en-us/graph/)

[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)



## Enterprise Features[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://github.com/PowerShell/PowerShell)[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://github.com/PowerShell/PowerShell)



- **Enterprise Scale**: Optimized for organizations with 1000+ SharePoint sites[![Graph API](https://img.shields.io/badge/Microsoft%20Graph-v1.0%2Bbeta-green.svg)](https://docs.microsoft.com/en-us/graph/)[![Graph API](https://img.shields.io/badge/Microsoft%20Graph-v1.0-green.svg)](https://docs.microsoft.com/en-us/graph/)

- **High Performance**: Bulk Graph API operations with silent processing mode

- **Interactive Workflow**: User-friendly site selection with automated inventory[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

- **Comprehensive Reporting**: Detailed extraction with consolidated CSV outputs

- **Robust Processing**: Advanced error handling and site validation

- **Progress Tracking**: Real-time progress indicators for large operations

- **Production Ready**: App-only authentication for automated enterprise deployment## 🚀 Enterprise Features## 🚀 Features



## Prerequisites



### PowerShell Modules- **🏢 Enterprise Scale**: Optimized for organizations with 1000+ SharePoint sites- **Comprehensive Enumeration**: Scan all SharePoint sites or target specific sites/departments

- `Microsoft.Graph` (automatically installed)

- `ExchangeOnlineManagement` (for Purview label resolution - automatically installed)- **⚡ High Performance**: Bulk Graph API operations with silent processing mode- **Multiple Authentication Methods**: Interactive (delegated) or App-only authentication



### API Permissions Required- **🎯 Interactive Workflow**: User-friendly site selection with automated inventory- **Label Resolution**: Automatic mapping of label GUIDs to display names via Purview or Graph API



**Microsoft Graph Application Permissions:**- **📊 Comprehensive Reporting**: Detailed extraction with consolidated CSV outputs- **Large Tenant Optimization**: Streaming mode, checkpointing, and memory optimization for millions of files

- `Sites.Read.All` - Read SharePoint sites and document libraries

- `Files.Read.All` - Read files and extract sensitivity labels- **🔄 Robust Processing**: Advanced error handling and site validation- **Flexible Export Options**: CSV, JSON, split files, or database export

- `InformationProtectionPolicy.Read.All` - Read sensitivity label catalog

- **📈 Progress Tracking**: Real-time progress indicators for large operations- **Robust Error Handling**: Retry logic, graceful failures, progress reporting

## Quick Start

- **🛡️ Production Ready**: App-only authentication for automated enterprise deployment- **Configurable Filtering**: Include/exclude sites, libraries, and file types

### 1. Clone the Repository

```bash- **Resume Capability**: Resume processing from checkpoints for large operations

git clone https://github.com/yourusername/sharepoint-sensitivity-labels-extractor.git

cd sharepoint-sensitivity-labels-extractor## 📋 Prerequisites

```

## 📋 Prerequisites

### 2. Configure Authentication

Copy the template and add your credentials:### PowerShell Modules

```powershell

Copy-Item config.template.json config.json- `Microsoft.Graph` (automatically installed)### PowerShell Modules

# Edit config.json with your tenant details

```- `ExchangeOnlineManagement` (for Purview label resolution - automatically installed)- `Microsoft.Graph` (automatically installed)



**App-Only Authentication (Recommended for Enterprise):**- `ExchangeOnlineManagement` (for Purview label resolution - automatically installed)

```json

{### API Permissions Required

  "authentication": {

    "preferredMethod": "appOnly",### API Permissions Required

    "appOnly": {

      "tenantId": "YOUR_TENANT_ID_HERE",**Microsoft Graph Application Permissions:**

      "clientId": "YOUR_CLIENT_ID_HERE", 

      "clientSecret": "YOUR_CLIENT_SECRET_HERE"- `Sites.Read.All` - Read SharePoint sites and document libraries**Microsoft Graph Application Permissions:**

    }

  }- `Files.Read.All` - Read files and extract sensitivity labels- `Sites.Read.All` - Read SharePoint sites and document libraries

}

```- `InformationProtectionPolicy.Read.All` - Read sensitivity label catalog- `Files.Read.All` - Read files and extract sensitivity labels



### 3. Run the Master Extractor- `InformationProtectionPolicy.Read.All` - Read sensitivity label catalog

```powershell

.\Master-SharePoint-Extractor.ps1## 🔧 Quick Start

```

## Quick Start

The tool will guide you through an interactive workflow:

1. **Inventory Generation**: Automatically discovers all SharePoint sites### 1. Clone the Repository

2. **Site Selection**: Choose between ALL sites or SPECIFIC sites

3. **Bulk Processing**: High-performance extraction with progress tracking```bash### 1. Clone the Repository

4. **Consolidated Reporting**: Unified CSV output with detailed statistics

git clone https://github.com/yourusername/sharepoint-sensitivity-labels-extractor.git```powershell

## Architecture

cd sharepoint-sensitivity-labels-extractorgit clone https://github.com/yourusername/sharepoint-sensitivity-labels.git

### Core Components

```cd sharepoint-sensitivity-labels

| Component | Purpose | Enterprise Features |

|-----------|---------|-------------------|```

| `Master-SharePoint-Extractor.ps1` | Main orchestrator | Bulk site enumeration, user workflow, consolidated reporting |

| `Comprehensive-Scanner.ps1` | Individual site processor | Silent mode, progress indicators, robust error handling |### 2. Configure Authentication

| `ConfigHelpers.ps1` | Configuration utilities | Template validation, secure credential handling |

| `config.json` | Authentication & settings | App-only auth, performance tuning, filter configuration |Copy the template and add your credentials:### 2. Configure Authentication



## Configuration```powershell



### Enterprise Performance SettingsCopy-Item config.template.json config.json**Option A: Interactive Authentication (Testing)**

```json

{# Edit config.json with your tenant details```json

  "performance": {

    "graphPageSize": 999,```// config.json

    "maxRetries": 5,

    "batchSize": 1000,{

    "enableProgressReporting": true,

    "verboseLogging": false**App-Only Authentication (Recommended for Enterprise):**  "authentication": {

  }

}```json    "preferredMethod": "interactive"

```

{  }

### Site Filtering for Large Organizations

```json  "authentication": {}

{

  "siteFilters": {    "preferredMethod": "appOnly",```

    "includeSiteSearch": ["*"],

    "excludeSiteUrls": [    "appOnly": {

      "https://tenant.sharepoint.com/sites/appcatalog",

      "https://tenant.sharepoint.com/sites/contentcenter"      "tenantId": "YOUR_TENANT_ID_HERE",**Option B: App-Only Authentication (Production)**

    ],

    "sampleSiteSearches": {      "clientId": "YOUR_CLIENT_ID_HERE", ```json

      "specificDepartments": ["Finance*", "HR*", "Legal*", "Compliance*"],

      "projectSites": ["Project*", "Team*", "Initiative*"],      "clientSecret": "YOUR_CLIENT_SECRET_HERE"// config.json

      "businessUnits": ["North America*", "EMEA*", "APAC*"]

    }    }{

  }

}  }  "authentication": {

```

}    "preferredMethod": "appOnly",

## Sample Enterprise Output

```    "appOnly": {

### Extraction Summary

```      "tenantId": "your-tenant-id",

EXTRACTION SUMMARY:

Sites Processed: 1,247### 3. Run the Master Extractor      "clientId": "your-app-registration-id", 

Total Records: 2,847,392

Files with Labels: 2,156,744 (76%)```powershell      "certThumbprint": "your-certificate-thumbprint"

Unique Sites: 1,201

Unique File Types: 12.\Master-SharePoint-Extractor.ps1    }

Processing Time: 4h 23m

``````  }



### Consolidated Report Schema}

| Column | Description | Example |

|--------|-------------|---------|The tool will guide you through an interactive workflow:```

| `TenantId` | Azure AD Tenant ID | `12345678-1234-5678-9abc-def012345678` |

| `SiteName` | SharePoint site display name | `Finance Department - North America` |1. **📊 Inventory Generation**: Automatically discovers all SharePoint sites

| `SiteUrl` | Full SharePoint site URL | `https://contoso.sharepoint.com/sites/finance-na` |

| `LibraryName` | Document library name | `Financial Reports` |2. **🎯 Site Selection**: Choose between ALL sites or SPECIFIC sites### 3. Run the Script

| `ItemName` | File name with extension | `Q4_2024_Budget_Analysis.xlsx` |

| `ItemPath` | Full file path | `/sites/finance-na/Shared Documents/Reports/Q4_Budget.xlsx` |3. **⚡ Bulk Processing**: High-performance extraction with progress tracking

| `FileType` | File extension | `xlsx` |

| `SizeInBytes` | File size | `2,147,483` |4. **📈 Consolidated Reporting**: Unified CSV output with detailed statistics**Basic scan (all sites, default file types):**

| `CreatedBy` | File creator | `john.smith@contoso.com` |

| `CreatedDate` | Creation timestamp | `2024-10-15 09:30:00` |```powershell

| `ModifiedBy` | Last modifier | `jane.doe@contoso.com` |

| `ModifiedDate` | Last modified | `2024-12-01 14:22:00` |## 🏗️ Architecture.\Enumerate-SharePointSensitivityLabels.ps1

| `LabelIds` | Sensitivity label GUIDs | `defa4170-0d19-0005-0005-bc88714345d2` |

| `LabelNames` | Human-readable label names | `Confidential - Finance` |```

| `AssignmentMethod` | How label was applied | `Standard` |

| `Justification` | Label justification | `Contains financial data requiring protection` |### Core Components



## Enterprise Performance**Target specific departments:**



### Scalability Benchmarks| Component | Purpose | Enterprise Features |```powershell



| Organization Size | Sites | Files | Processing Time | Memory Usage ||-----------|---------|-------------------|.\Enumerate-SharePointSensitivityLabels.ps1 -IncludeSiteSearch @('Finance*', 'HR*', 'Legal*')

|------------------|-------|--------|----------------|--------------|

| **Small** | 10-100 | 1K-100K | 5-30 minutes | < 1GB || `Master-SharePoint-Extractor.ps1` | Main orchestrator | Bulk site enumeration, user workflow, consolidated reporting |```

| **Medium** | 100-500 | 100K-1M | 1-3 hours | 1-2GB |

| **Large** | 500-2000 | 1M-10M | 3-8 hours | 2-4GB || `Comprehensive-Scanner.ps1` | Individual site processor | Silent mode, progress indicators, robust error handling |

| **Enterprise** | 2000+ | 10M+ | 8-24 hours | 4-8GB |

| `ConfigHelpers.ps1` | Configuration utilities | Template validation, secure credential handling |**Large tenant mode:**

### Performance Optimizations

| `config.json` | Authentication & settings | App-only auth, performance tuning, filter configuration |```powershell

#### Implemented

- **Bulk Graph API Calls**: Single requests for maximum page sizes (999 items).\Enumerate-SharePointSensitivityLabels.ps1 -EnableLargeTenantMode

- **Silent Processing Mode**: Eliminates console verbosity that degrades performance

- **Conservative File Estimation**: Avoids expensive API calls for file counting## ⚙️ Configuration```

- **Efficient Site Enumeration**: High-speed bulk operations vs individual requests

- **Memory Management**: Streaming output and garbage collection optimization

- **Progress Indicators**: Minimal UI updates every 10 files vs per-file logging

### Enterprise Performance Settings## ⚙️ Configuration

## Enterprise Security

```json

### Authentication Best Practices

- **App-Only Authentication**: Recommended for production automation{The script uses a flexible JSON configuration system. See [CONFIG.md](CONFIG.md) for detailed configuration options.

- **Certificate-Based Auth**: Enhanced security over client secrets

- **Least Privilege**: Minimal required permissions (Read-only)  "performance": {

- **Audit Trail**: All operations logged with timestamps and user context

    "graphPageSize": 999,### Key Configuration Sections:

### Data Protection

- **In-Transit Encryption**: All Graph API calls use HTTPS/TLS    "maxRetries": 5,

- **Local Data Handling**: Temporary files cleared after processing  

- **Sensitive Data**: No credentials stored in output files    "batchSize": 1000,| Section | Purpose | Example |

- **Compliance**: Supports data residency and retention policies

    "enableProgressReporting": true,|---------|---------|---------|

## Advanced Usage

    "verboseLogging": false| `authentication` | Auth method and credentials | Interactive vs App-only |

### Department-Specific Extraction

```powershell  }| `siteFilters` | Which sites to include/exclude | Department-specific scans |

# Process only Finance and HR departments

# Edit the sites inventory CSV to mark specific departments for extraction}| `fileProcessing` | File types and processing options | Office docs, PDFs, etc. |

.\Master-SharePoint-Extractor.ps1

# Select Option 2: Extract from SPECIFIC sites```| `performance` | API throttling and batch sizes | Optimize for tenant size |

# Edit CSV to set Extract=True only for desired departments

```| `output` | Export format and file splitting | CSV, JSON, database |



### Automated Enterprise Deployment### Site Filtering for Large Organizations| `largeTenant` | Optimizations for millions of files | Streaming, checkpoints |

```powershell

# Run in silent mode for automation```json

$sites = Import-Csv "enterprise-sites-list.csv"

$sites | ForEach-Object { $_.Extract = "True" }{## 📊 Large Organization Support

$sites | Export-Csv "sites-inventory.csv" -NoTypeInformation

  "siteFilters": {

# Process all sites automatically

.\Master-SharePoint-Extractor.ps1    "includeSiteSearch": ["*"],For organizations with millions of files, this tool provides several optimization strategies:

# Select Option 1: Extract from ALL sites

```    "excludeSiteUrls": [



## Troubleshooting      "https://tenant.sharepoint.com/sites/appcatalog",### 🔄 Streaming Mode



### Common Enterprise Issues      "https://tenant.sharepoint.com/sites/contentcenter"```json



#### Graph API Throttling    ],{

```json

{    "sampleSiteSearches": {  "largeTenant": {

  "performance": {

    "graphPageSize": 200,      "specificDepartments": ["Finance*", "HR*", "Legal*", "Compliance*"],    "enabled": true,

    "baseDelaySeconds": 5,

    "maxRetries": 10      "projectSites": ["Project*", "Team*", "Initiative*"],    "streamingMode": true,

  }

}      "businessUnits": ["North America*", "EMEA*", "APAC*"]    "checkpointInterval": 10000

```

    }  }

#### Empty SiteName Errors

The tool automatically filters sites with empty/invalid names:  }}

```

WARNING: Filtered out 3 sites with empty/invalid names}```

   - Row with SiteUrl: 'https://contoso.sharepoint.com/sites/archived-site' (empty SiteName)

``````



## Output Structure### 📁 File Splitting Options



### Run Folder Organization## 📊 Sample Enterprise Output```json

```

Output/{

└── Run_20240904_170637/

    ├── Sensitivity-Labels-Cache.json### Extraction Summary  "output": {

    ├── SharePoint-Sites-Inventory_20240904_170637.csv

    ├── Site_Finance_Department_20240904_170637.csv```    "splitting": {

    ├── Site_HR_Command_Centre_20240904_170637.csv

    └── Consolidated_SharePoint_SensitivityLabels_20240904_170637.csv📈 EXTRACTION SUMMARY:      "enabled": true,

```

Sites Processed: 1,247      "maxRecordsPerFile": 100000,

## Contributing

Total Records: 2,847,392      "splitBySite": true

1. Fork the repository

2. Create a feature branch (`git checkout -b feature/enterprise-enhancement`)Files with Labels: 2,156,744 (76%)    }

3. Commit changes with clear messages (`git commit -m 'Add bulk processing optimization'`)

4. Push to the branch (`git push origin feature/enterprise-enhancement`)Unique Sites: 1,201  }

5. Open a Pull Request with detailed description

Unique File Types: 12}

### Development Guidelines

- Follow PowerShell best practices and style guidelinesProcessing Time: 4h 23m```

- Include comprehensive error handling for enterprise scenarios

- Add performance benchmarks for new features```

- Update documentation for configuration changes

- Test with large data sets when possible### 💾 Export Formats for Scale



## License### Consolidated Report Schema



This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.| Column | Description | Example || Format | Best For | Max Recommended Size |



## Acknowledgments|--------|-------------|---------||--------|----------|---------------------|



- **Microsoft Graph Team**: For the powerful extractSensitivityLabels API| `TenantId` | Azure AD Tenant ID | `12345678-1234-5678-9abc-def012345678` || **CSV** | Excel, Power BI | < 1M records |

- **PowerShell Community**: For excellent module ecosystem

- **Microsoft 365 Community**: For real-world enterprise feedback and testing| `SiteName` | SharePoint site display name | `Finance Department - North America` || **JSON Lines** | APIs, streaming | Unlimited |

- **Compliance Professionals**: For requirements gathering and validation

| `SiteUrl` | Full SharePoint site URL | `https://contoso.sharepoint.com/sites/finance-na` || **Database** | Enterprise reporting | Unlimited |

## Enterprise Support

| `LibraryName` | Document library name | `Financial Reports` || **Parquet** | Analytics platforms | Unlimited |

- **Issues**: [GitHub Issues](https://github.com/yourusername/sharepoint-sensitivity-labels-extractor/issues)

- **Enterprise Consulting**: Available for large-scale implementations| `ItemName` | File name with extension | `Q4_2024_Budget_Analysis.xlsx` |

- **Custom Reporting**: Tailored dashboard and integration development

- **Training**: Enterprise PowerShell and Graph API workshops| `ItemPath` | Full file path | `/sites/finance-na/Shared Documents/Reports/Q4_Budget.xlsx` |## 📈 Performance Guidelines



---| `FileType` | File extension | `xlsx` |



**Built for Microsoft 365 Enterprise Administrators and Compliance Teams**| `SizeInBytes` | File size | `2,147,483` || Tenant Size | Recommended Settings | Expected Runtime |



> *Efficiently extract sensitivity labels from millions of SharePoint files across thousands of sites with enterprise-grade performance and reliability.*| `CreatedBy` | File creator | `john.smith@contoso.com` ||-------------|---------------------|------------------|

| `CreatedDate` | Creation timestamp | `2024-10-15 09:30:00` || **Small** (< 100K files) | Default settings | 1-2 hours |

| `ModifiedBy` | Last modifier | `jane.doe@contoso.com` || **Medium** (100K-1M files) | Increased batch sizes | 4-8 hours |

| `ModifiedDate` | Last modified | `2024-12-01 14:22:00` || **Large** (> 1M files) | Streaming mode + splitting | 12-24+ hours |

| `LabelIds` | Sensitivity label GUIDs | `defa4170-0d19-0005-0005-bc88714345d2` |

| `LabelNames` | Human-readable label names | `Confidential - Finance` |## 🛡️ Security Considerations

| `AssignmentMethod` | How label was applied | `Standard` |

| `Justification` | Label justification | `Contains financial data requiring protection` |- **Least Privilege**: Use only required API permissions

- **Certificate Storage**: Store certificates securely (Azure Key Vault recommended)

## 🚀 Enterprise Performance- **Audit Logging**: All operations are logged with timestamps

- **Data Handling**: Consider data classification when exporting results

### Scalability Benchmarks

## 📁 Output Schema

| Organization Size | Sites | Files | Processing Time | Memory Usage |

|------------------|-------|--------|----------------|--------------|The tool generates comprehensive reports with the following data points:

| **Small** | 10-100 | 1K-100K | 5-30 minutes | < 1GB |

| **Medium** | 100-500 | 100K-1M | 1-3 hours | 1-2GB || Column | Description | Example |

| **Large** | 500-2000 | 1M-10M | 3-8 hours | 2-4GB ||--------|-------------|---------|

| **Enterprise** | 2000+ | 10M+ | 8-24 hours | 4-8GB || `TenantId` | Azure AD Tenant ID | `12345678-1234-...` |

| `SiteName` | SharePoint site display name | `Finance Department` |

### Performance Optimizations| `SiteUrl` | Full SharePoint site URL | `https://contoso.sharepoint.com/sites/finance` |

| `LibraryName` | Document library name | `Shared Documents` |

#### ✅ Implemented| `FolderPath` | File folder path | `/Budget/2024` |

- **Bulk Graph API Calls**: Single requests for maximum page sizes (999 items)| `FileName` | File name with extension | `Q1_Budget.xlsx` |

- **Silent Processing Mode**: Eliminates console verbosity that degrades performance| `FileSizeBytes` | File size in bytes | `2048576` |

- **Conservative File Estimation**: Avoids expensive API calls for file counting| `CreatedBy` | File creator | `john.doe@contoso.com` |

- **Efficient Site Enumeration**: High-speed bulk operations vs individual requests| `CreatedUtc` | Creation timestamp | `2024-01-15T10:30:00Z` |

- **Memory Management**: Streaming output and garbage collection optimization| `ModifiedUtc` | Last modified timestamp | `2024-03-10T14:20:00Z` |

- **Progress Indicators**: Minimal UI updates every 10 files vs per-file logging| `LabelIds` | Sensitivity label GUIDs | `abcd1234-5678-...` |

| `LabelNames` | Human-readable label names | `Confidential - Finance` |

## 🛡️ Enterprise Security| `AssignmentMethod` | How label was applied | `Standard`, `Privileged`, `Auto` |



### Authentication Best Practices## 🚨 Common Issues & Troubleshooting

- **App-Only Authentication**: Recommended for production automation

- **Certificate-Based Auth**: Enhanced security over client secrets### Authentication Issues

- **Least Privilege**: Minimal required permissions (Read-only)```powershell

- **Audit Trail**: All operations logged with timestamps and user context# Verify Graph connection

Get-MgContext

### Data Protection

- **In-Transit Encryption**: All Graph API calls use HTTPS/TLS# Clear cached credentials

- **Local Data Handling**: Temporary files cleared after processing  Disconnect-MgGraph

- **Sensitive Data**: No credentials stored in output files```

- **Compliance**: Supports data residency and retention policies

### API Throttling

## 🔧 Advanced UsageThe script automatically handles Microsoft Graph throttling with exponential backoff. For heavily throttled tenants, consider:

- Reducing `graphPageSize` in configuration

### Department-Specific Extraction- Increasing `baseDelaySeconds`

```powershell- Running during off-peak hours

# Process only Finance and HR departments

# Edit the sites inventory CSV to mark specific departments for extraction### Memory Issues (Large Tenants)

.\Master-SharePoint-Extractor.ps1```json

# Select Option 2: Extract from SPECIFIC sites{

# Edit CSV to set Extract=True only for desired departments  "largeTenant": {

```    "memoryOptimization": true,

    "streamingMode": true

### Automated Enterprise Deployment  }

```powershell}

# Run in silent mode for automation```

$sites = Import-Csv "enterprise-sites-list.csv"

$sites | ForEach-Object { $_.Extract = "True" }### Resume Failed Operations

$sites | Export-Csv "sites-inventory.csv" -NoTypeInformation```powershell

.\Enumerate-SharePointSensitivityLabels.ps1 -ResumeFromCheckpoint

# Process all sites automatically```

.\Master-SharePoint-Extractor.ps1

# Select Option 1: Extract from ALL sites## 📚 Advanced Usage Examples

```

### Compliance Audit - All Sensitive Files

## 🚨 Troubleshooting```powershell

# Focus on files likely to have sensitivity labels

### Common Enterprise Issues.\Enumerate-SharePointSensitivityLabels.ps1 -Extensions @('docx','xlsx','pptx','pdf') -IncludeSiteSearch @('*')

```

#### Graph API Throttling

```json### Department-Specific Scan

{```powershell

  "performance": {# Scan only HR and Legal departments

    "graphPageSize": 200,.\Enumerate-SharePointSensitivityLabels.ps1 -IncludeSiteSearch @('HR*','Legal*','Compliance*')

    "baseDelaySeconds": 5,```

    "maxRetries": 10

  }### High-Performance Scan for Large Tenants

}```powershell

```# Optimized for millions of files

.\Enumerate-SharePointSensitivityLabels.ps1 -EnableLargeTenantMode -ConfigPath "large-tenant-config.json"

#### Empty SiteName Errors```

The tool automatically filters sites with empty/invalid names:

```## 🤝 Contributing

⚠️ Filtered out 3 sites with empty/invalid names

   - Row with SiteUrl: 'https://contoso.sharepoint.com/sites/archived-site' (empty SiteName)1. Fork the repository

```2. Create a feature branch (`git checkout -b feature/amazing-feature`)

3. Commit your changes (`git commit -m 'Add amazing feature'`)

## 📁 Output Structure4. Push to the branch (`git push origin feature/amazing-feature`)

5. Open a Pull Request

## Output Structure

### Run Folder Organization

```## 📄 License

Output/

└── Run_20240904_170637/This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

    ├── Sensitivity-Labels-Cache.json

    ├── SharePoint-Sites-Inventory_20240904_170637.csv## 🙏 Acknowledgments

    ├── Site_Finance_Department_20240904_170637.csv

    ├── Site_HR_Command_Centre_20240904_170637.csv- Based on Microsoft Graph `driveItem: extractSensitivityLabels` API

    └── Consolidated_SharePoint_SensitivityLabels_20240904_170637.csv- Inspired by [Practical 365](https://practical365.com/) guidance

```- Built for the Microsoft 365 community



## 🤝 Contributing## 📞 Support



1. Fork the repository- 🐛 **Issues**: [GitHub Issues](https://github.com/yourusername/sharepoint-sensitivity-labels/issues)

2. Create a feature branch (`git checkout -b feature/enterprise-enhancement`)- 💬 **Discussions**: [GitHub Discussions](https://github.com/yourusername/sharepoint-sensitivity-labels/discussions)

3. Commit changes with clear messages (`git commit -m 'Add bulk processing optimization'`)- 📖 **Documentation**: [Wiki](https://github.com/yourusername/sharepoint-sensitivity-labels/wiki)

4. Push to the branch (`git push origin feature/enterprise-enhancement`)

5. Open a Pull Request with detailed description---



## 📄 License**⚡ Made for Microsoft 365 Administrators and Compliance Officers**

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- **Microsoft Graph Team**: For the powerful extractSensitivityLabels API
- **PowerShell Community**: For excellent module ecosystem
- **Microsoft 365 Community**: For real-world enterprise feedback and testing
- **Compliance Professionals**: For requirements gathering and validation

## 📞 Enterprise Support

- 🐛 **Issues**: [GitHub Issues](https://github.com/yourusername/sharepoint-sensitivity-labels-extractor/issues)
- 💼 **Enterprise Consulting**: Available for large-scale implementations
- 📊 **Custom Reporting**: Tailored dashboard and integration development
- 🎓 **Training**: Enterprise PowerShell and Graph API workshops

---

**⚡ Built for Microsoft 365 Enterprise Administrators and Compliance Teams**

> *Efficiently extract sensitivity labels from millions of SharePoint files across thousands of sites with enterprise-grade performance and reliability.*