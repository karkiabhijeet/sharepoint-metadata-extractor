# SharePoint Sensitivity Labels Configuration

This configuration file allows you to customize the SharePoint Sensitivity Labels enumeration script for your environment.

## 🔧 Quick Setup

### Step 1: Authentication Configuration

**Option A: Interactive Authentication (Recommended for testing)**
```json
"authentication": {
  "preferredMethod": "interactive"
}
```

**Option B: App-Only Authentication (Recommended for production)**
```json
"authentication": {
  "preferredMethod": "appOnly",
  "appOnly": {
    "tenantId": "your-tenant-id-here",
    "clientId": "your-app-registration-client-id",
    "certThumbprint": "your-certificate-thumbprint",
    "certStore": "CurrentUser"
  }
}
```

### Step 2: Site Filtering

Choose your site discovery strategy:

**All Sites:**
```json
"includeSiteSearch": ["*"]
```

**Specific Departments:**
```json
"includeSiteSearch": ["Finance*", "HR*", "Legal*"]
```

**Specific Sites:**
```json
"includeSiteSearch": ["Contoso Intranet", "Executive Team"]
```

## 📊 Large Tenant Optimization

For organizations with millions of files, enable large tenant mode:

```json
"largeTenant": {
  "enabled": true,
  "streamingMode": true,
  "checkpointInterval": 10000,
  "memoryOptimization": true
}
```

### Output Splitting Options:

**Split by Record Count:**
```json
"splitting": {
  "enabled": true,
  "maxRecordsPerFile": 100000
}
```

**Split by Site:**
```json
"splitting": {
  "enabled": true,
  "splitBySite": true
}
```

## 🚀 App Registration Setup

### Required API Permissions:

**Microsoft Graph:**
- `Sites.Read.All` (Application)
- `Files.Read.All` (Application)
- `InformationProtectionPolicy.Read.All` (Application)

**PowerShell Commands to Create App Registration:**
```powershell
# Install required modules
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module Microsoft.Graph.Applications -Scope CurrentUser

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Application.ReadWrite.All"

# Create app registration
$appName = "SharePoint-SensitivityLabels-Scanner"
$app = New-MgApplication -DisplayName $appName

# Add required permissions
$graphServicePrincipal = Get-MgServicePrincipal -Filter "displayName eq 'Microsoft Graph'"
$permissions = @(
    @{ Id = "Sites.Read.All"; Type = "Role" },
    @{ Id = "Files.Read.All"; Type = "Role" },
    @{ Id = "InformationProtectionPolicy.Read.All"; Type = "Role" }
)

# Update app with permissions (additional setup required in Azure Portal)
```

### Certificate Setup:

**Generate Self-Signed Certificate:**
```powershell
# Generate certificate
$cert = New-SelfSignedCertificate -Subject "CN=SharePointLabelsScanner" -CertStoreLocation "Cert:\CurrentUser\My" -KeySpec KeyExchange

# Export public key for Azure AD
Export-Certificate -Cert $cert -FilePath "SharePointLabelsScanner.cer"

# Get thumbprint
$cert.Thumbprint
```

## 📈 Performance Tuning

### Small Tenants (< 100K files):
```json
"performance": {
  "graphPageSize": 200,
  "batchSize": 1000,
  "enableProgressReporting": true
}
```

### Medium Tenants (100K - 1M files):
```json
"performance": {
  "graphPageSize": 500,
  "batchSize": 5000,
  "enableProgressReporting": true
}
```

### Large Tenants (> 1M files):
```json
"performance": {
  "graphPageSize": 1000,
  "batchSize": 10000,
  "enableProgressReporting": false
},
"largeTenant": {
  "enabled": true,
  "streamingMode": true,
  "backgroundProcessing": true
}
```

## 🔍 File Type Configuration

**Office Documents Only:**
```json
"extensions": ["docx", "xlsx", "pptx"]
```

**Include Legacy Office:**
```json
"extensions": ["docx", "xlsx", "pptx", "doc", "xls", "ppt"]
```

**Include PDFs:**
```json
"extensions": ["docx", "xlsx", "pptx", "pdf"]
```

## 📤 Export Options

### CSV (Default - Best for Excel/Power BI):
```json
"exportFormat": "csv"
```

### JSON (Best for APIs/Integration):
```json
"exportFormat": "json"
```

### Database (Best for Large Datasets):
```json
"exportFormat": "database",
"database": {
  "connectionString": "Server=server;Database=db;Integrated Security=true;",
  "tableName": "SharePointSensitivityLabels"
}
```

## 🎯 Common Scenarios

### Scenario 1: Security Audit (All Files)
```json
{
  "includeSiteSearch": ["*"],
  "extensions": ["docx", "xlsx", "pptx", "pdf", "doc", "xls", "ppt"],
  "exportFormat": "csv",
  "generateSummary": true
}
```

### Scenario 2: Compliance Check (Finance Department)
```json
{
  "includeSiteSearch": ["Finance*", "Accounting*"],
  "extensions": ["docx", "xlsx", "pdf"],
  "exportFormat": "json",
  "emailResults": true
}
```

### Scenario 3: Large Organization (Streaming Mode)
```json
{
  "includeSiteSearch": ["*"],
  "largeTenant": {
    "enabled": true,
    "streamingMode": true,
    "checkpointInterval": 5000
  },
  "splitting": {
    "enabled": true,
    "maxRecordsPerFile": 50000
  }
}
```