# IT Asset Collection Script

**Author**: Marvin De Los Angeles  
**Department**: CIT Automation / AI / UX  
**Version**: 1.0  
**Date**: 2025-05-24  

## Description

This PowerShell script collects comprehensive hardware, operating system, and installed software information from Windows machines and automatically emails the report to the IT Asset Management team. The script generates both a CSV file and an HTML email report for asset tracking and inventory management.

## What Information is Collected

The script gathers the following system information:

### Hardware Information
- **System**: Manufacturer, Model, Computer Name, System Type
- **BIOS**: Serial Number, Version
- **CPU**: Name, Number of Cores, Logical Processors, Max Clock Speed
- **Memory (RAM)**: Capacity (in GB), Manufacturer, Speed for each module
- **Storage**: Disk Model, Size (in GB), Serial Number, Interface Type for each drive

### Software Information
- **Operating System**: Name, Version, Architecture (32-bit/64-bit)
- **Installed Software**: All installed programs with versions from both 32-bit and 64-bit registry locations

### User Information
- Current logged-in user at the time of script execution

## Files in This Package

- **`asset_collect.ps1`** - Main PowerShell script that collects system information
- **`asset_scan.bat`** - Batch wrapper that executes the PowerShell script with proper execution policy
- **`README.md`** - This documentation file

## How to Use

### Method 1: Using the Batch File (Recommended)
1. Ensure Microsoft Outlook is installed and configured on your machine
2. Download or copy both files into the same folder:
   - `asset_collect.ps1`
   - `asset_scan.bat`
3. **Double-click `asset_scan.bat`** to run the asset collection process
   - The batch file automatically executes the PowerShell script with the correct permissions
   - No manual PowerShell execution required

### Method 2: Direct PowerShell Execution
```powershell
powershell -ExecutionPolicy Bypass -File asset_collect.ps1
```

## Output

### CSV File
- **Filename**: `asset_report.csv`
- **Format**: Single-row CSV without quotes around values
- **Content**: All collected system information in a structured format

### Email Report
- **Recipient**: CIT.Automations@robinsonsretail.com.ph
- **Subject**: "IT Asset Report [ComputerName]"
- **Format**: HTML email with embedded data table
- **Attachment**: The generated CSV file

## System Requirements

- **Operating System**: Windows (any version with WMI support)
- **PowerShell**: Version 5.1 or later
- **Email Client**: Microsoft Outlook (installed and configured)
- **Permissions**: Standard user rights (administrator rights not required unless WMI is restricted)

## Technical Details

### Data Collection Methods
- Uses **WMI/CIM** (Windows Management Instrumentation) for hardware information
- Queries **Windows Registry** for installed software information
- Accesses both 32-bit and 64-bit registry hives for complete software inventory

### Email Integration
- Uses **Outlook COM object** for email functionality
- Requires Outlook to be installed and configured with a valid profile
- Automatically attaches the CSV file and embeds data in email body

### Security Considerations
- **Read-only operations**: Script only reads system information, makes no changes
- **No credentials stored**: Uses current user's Outlook profile
- **Minimal permissions**: Does not require administrator privileges
- **Local execution**: All processing done locally on the target machine

## Data Fields Collected

| Field | Description | Example |
|-------|-------------|---------|
| Title | Unique report identifier | AR-12345678 |
| User | Current logged-in user | DOMAIN\username |
| Manufacturer | System manufacturer | LENOVO |
| Model | System model | ThinkPad T480 |
| ComputerName | Computer name | DESKTOP-ABC123 |
| SystemType | System architecture | x64-based PC |
| BIOS_Serial | BIOS serial number | PF1A2B3C |
| BIOS_Version | BIOS version | LENOVO - 1.45 |
| OS | Operating system name | Microsoft Windows 11 Pro |
| OS_Version | OS version number | 10.0.22000 |
| OS_Architecture | OS architecture | 64-bit |
| CPU | Processor name | Intel(R) Core(TM) i7-8650U |
| CPU_Cores | Physical CPU cores | 4 |
| CPU_Logical | Logical processors | 8 |
| CPU_MaxClock | Maximum clock speed (MHz) | 2112 |
| RAM | Memory details (formatted) | Capacity=16GB;Manufacturer=Samsung;Speed=2400 |
| Disk | Storage details (formatted) | Model=SSD;Size=512GB;SerialNumber=ABC123 |
| InstalledSoftware | All installed programs | Microsoft Office (365) \| Adobe Reader (2023) |

## Troubleshooting

### Common Issues
- **Email not sending**: Verify Outlook is installed and configured
- **Access denied errors**: Ensure WMI access is not restricted by group policy
- **Script execution blocked**: Use the provided batch file or set execution policy

### Error Handling
- Script includes error handling for WMI queries
- Continues execution even if individual components fail
- Provides status messages during execution

## Change Log

- **Initial version**: Basic system information collection
- **v1.0**: Added error handling for WMI queries
- **v1.0**: Improved formatting of multi-value fields  
- **v1.0**: Updated email body to include CSV content as HTML table
- **v1.0**: Fixed CSV export to avoid quotes around values
- **v1.0**: Added random title generation for asset report
- **v1.0**: Implemented human-readable sizes (GB format) for RAM and disk storage

## Support

For questions, issues, or feature requests, please contact:
**CIT Automation / AI / UX Team**  
Email: CIT.Automations@robinsonsretail.com.ph

---

*This script is designed for internal IT asset management and inventory purposes.*
