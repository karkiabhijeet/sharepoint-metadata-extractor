# 🔧 GitHub Download Issues Fix

## Problem: Getting Parse Errors After Downloading?

If you downloaded this script from GitHub and get errors like:
```
Unexpected token '}' in expression or statement.
Unexpected token '`n' in expression or statement.  
The ampersand (&) character is not allowed.
You must provide a value expression following the '%' operator.
Unexpected token 'records' in expression or statement.
Missing closing ')' in expression.
```

**Root Cause**: GitHub ZIP downloads can corrupt PowerShell files, causing syntax errors in:
- Graph API URLs (ampersand characters)
- String interpolation (percentage symbols)
- Quote handling (missing closures)
- Backtick escaping (newline characters)

## ✅ Quick Fix (Try in Order)

### Method 1: Automatic Fix
```powershell
# First, unblock the files
Get-ChildItem -Path . -Filter "*.ps1" | Unblock-File

# Then run the fix
.\Fix-DownloadedScript.ps1

# Finally run the main script
.\Master-SharePoint-Extractor.ps1
```

### Method 2: Emergency Rebuild (if Method 1 fails)
```powershell
.\Emergency-Rebuild.ps1
# This creates a basic validation script if corruption is severe
```

### Method 3: Git Clone (Recommended - Always Works)
```bash
git clone https://github.com/karkiabhijeet/sharepoint-metadata-extractor.git
cd sharepoint-metadata-extractor
# No corruption issues with git clone!
```

## Alternative Solutions

### Option 1: Unblock Files
```powershell
# Unblock all PowerShell files in the directory
Get-ChildItem -Path . -Filter "*.ps1" | Unblock-File
```

### Option 2: Manual Fix
If the automatic fix doesn't work:

1. Open the script in **PowerShell ISE** or **VS Code**
2. Save it as **UTF-8 with BOM** encoding
3. Replace any corrupted characters manually

### Option 3: Clone Instead of Download
```bash
git clone https://github.com/karkiabhijeet/sharepoint-metadata-extractor.git
cd sharepoint-metadata-extractor
```

## 🚀 Once Fixed

The script should run without issues and provide fast startup with lazy module loading:

```
✓ Module Microsoft.Graph available (Version: 2.x.x)
✓ Module ExchangeOnlineManagement available (Version: 3.x.x)
✓ All required modules are available!
```

## 📞 Still Having Issues?

If you continue to have problems:

1. Check PowerShell version: `$PSVersionTable`
2. Ensure you have the required modules installed
3. Run as Administrator if needed
4. Check the Issues section of the GitHub repository

---
*This fix addresses common GitHub download encoding issues with PowerShell scripts.*