# 🔧 GitHub Download Issues Fix

## Problem: Getting Parse Errors After Downloading?

If you downloaded this script from GitHub and get errors like:
```
Unexpected token '}' in expression or statement.
Unexpected token '`n' in expression or statement.
The ampersand (&) character is not allowed.
```

This is a common issue with PowerShell scripts downloaded from GitHub due to encoding corruption.

## ✅ Quick Fix

1. **Download the Fix Script**: Make sure you also download `Fix-DownloadedScript.ps1`

2. **Run the Fix**:
   ```powershell
   .\Fix-DownloadedScript.ps1
   ```

3. **Run the Main Script**:
   ```powershell
   .\Master-SharePoint-Extractor.ps1
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