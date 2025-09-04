# Output Directory

This directory contains all extraction results and generated reports.

## Structure

Each extraction run creates a timestamped subfolder with the following structure:

```
Output/
├── Run_YYYYMMDD_HHMMSS/
│   ├── Sensitivity-Labels-Cache.json
│   ├── SharePoint-Sites-Inventory_YYYYMMDD_HHMMSS.csv
│   ├── Site_[SiteName]_YYYYMMDD_HHMMSS.csv (individual site results)
│   └── Consolidated_SharePoint_SensitivityLabels_YYYYMMDD_HHMMSS.csv
└── [Additional run folders...]
```

## File Descriptions

- **Sensitivity-Labels-Cache.json**: Cached sensitivity labels from Microsoft Purview
- **SharePoint-Sites-Inventory_*.csv**: Complete inventory of SharePoint sites with processing status
- **Site_*_*.csv**: Individual site extraction results
- **Consolidated_SharePoint_SensitivityLabels_*.csv**: Combined results from all processed sites

## Data Cleanup

For GitHub repository preparation, this directory is excluded via .gitignore to prevent:
- Sensitive organizational data exposure
- Large file repository bloat
- Personal/confidential information leakage

The scripts will automatically create this folder structure when first run.
