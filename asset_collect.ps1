<#
    Script Name : IT Asset Collection Script
    Description : Collects hardware, OS, and installed software information from the local machine,
                  exports the data to a CSV file, and emails the report as an HTML table.
    Author      : Marvin De Los Angeles
    Department  : CIT Automation / AI / UX
    Version     : 1.0
    Date        : 2025-05-24
    Usage       : Run with appropriate permissions on a Windows machine with Outlook installed.
    Notes       : - Ensure Outlook is configured for sending emails.
                  - Script does not require admin rights unless WMI queries are restricted.
                  - run with PowerShell 5.1 or later.
                  - Run using command: powershell -ExecutionPolicy Bypass -File asset_collect.ps1
    Change Log  : - Initial version
    - Added error handling for WMI queries.
                  - Improved formatting of multi-value fields.
                  - Updated email body to include CSV content as HTML table.
                  - Fixed CSV export to avoid quotes around values.
    - Added random title generation for asset report.
    
#>

# Collect info
$system = Get-CimInstance Win32_ComputerSystem | Select-Object -First 1 Manufacturer,Model,Name,SystemType
$bios = Get-CimInstance Win32_BIOS | Select-Object -First 1 SerialNumber,Version
$os = Get-CimInstance Win32_OperatingSystem | Select-Object -First 1 Caption,Version,OSArchitecture
$cpu = Get-CimInstance Win32_Processor | Select-Object -First 1 Name,NumberOfCores,NumberOfLogicalProcessors,MaxClockSpeed
$ram = Get-CimInstance Win32_PhysicalMemory | Select-Object Capacity,Manufacturer,Speed
$disk = Get-CimInstance Win32_DiskDrive | Select-Object Model,Size,SerialNumber,InterfaceType
$software = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*,
                            HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |
    Where-Object { $_.DisplayName } |
    Select-Object DisplayName,DisplayVersion |
    Sort-Object DisplayName

$currentUser = $env:USERNAME

#$ramText = ($ram | ForEach-Object { "Capacity=$($_.Capacity);Manufacturer=$($_.Manufacturer);Speed=$($_.Speed)" }) -join " | "
#$diskText = ($disk | ForEach-Object { "Model=$($_.Model);Size=$($_.Size);SerialNumber=$($_.SerialNumber);InterfaceType=$($_.InterfaceType)" }) -join " | "
# Format multi-value fields as single-line text with human-readable sizes
$ramText = ($ram | ForEach-Object {
    $gb = [math]::Round($_.Capacity / 1GB, 2)
    "Capacity=${gb}GB;Manufacturer=$($_.Manufacturer);Speed=$($_.Speed)"
}) -join " | "

$diskText = ($disk | ForEach-Object {
    $gb = [math]::Round($_.Size / 1GB, 2)
    "Model=$($_.Model);Size=${gb}GB;SerialNumber=$($_.SerialNumber);InterfaceType=$($_.InterfaceType)"
}) -join " | "
$softwareText = ($software | ForEach-Object { "$($_.DisplayName) ($($_.DisplayVersion))" }) -join " | "

# Combine all info into a single object
$combined = [PSCustomObject]@{
    Title            = "AR-{0}" -f (Get-Random -Minimum 10000000 -Maximum 100000000)
    User             = $currentUser
    Manufacturer     = $system.Manufacturer
    Model            = $system.Model
    ComputerName     = $system.Name
    SystemType       = $system.SystemType
    BIOS_Serial      = $bios.SerialNumber
    BIOS_Version     = $bios.Version
    OS               = $os.Caption
    OS_Version       = $os.Version
    OS_Architecture  = $os.OSArchitecture
    CPU              = $cpu.Name
    CPU_Cores        = $cpu.NumberOfCores
    CPU_Logical      = $cpu.NumberOfLogicalProcessors
    CPU_MaxClock     = $cpu.MaxClockSpeed
    RAM              = $ramText
    Disk             = $diskText
    InstalledSoftware= $softwareText
}

# Export to CSV (single row) without quotes
$fields = 'Title','User','Manufacturer','Model','ComputerName','SystemType','BIOS_Serial','BIOS_Version','OS','OS_Version','OS_Architecture','CPU','CPU_Cores','CPU_Logical','CPU_MaxClock','RAM','Disk','InstalledSoftware','ClearSpace'
$values = $fields | ForEach-Object { $combined.$_ }
$headerLine = ($fields -join ',')
$dataLine = ($values -join ',')

$dataLine | Set-Content -Path "asset_report.csv" -Encoding UTF8


# Read CSV and convert to HTML table for email body
$csvTable = Import-Csv -Path "asset_report.csv" | ConvertTo-Html -Fragment | Out-String

# Email settings for Outlook (using default profile)
$to = "CIT.Automations@robinsonsretail.com.ph"
#$to = "Marvin.DelosAngeles@robinsonsretail.com.ph"
$subject = "IT Asset Report $($system.Name)"
$body = @"
<html>
<body>
<p>Please find the attached asset report.</p>
<p><b>CSV Content:</b></p>
$csvTable
</body>
</html>
"@
$attachment = "asset_report.csv"

# Create Outlook COM object
$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.CreateItem(0)
$mail.To = $to
$mail.Subject = $subject
$mail.HTMLBody = $body
$mail.Attachments.Add((Resolve-Path $attachment).Path)
$mail.Send()

Clear-Host
write-host "Asset report sent successfully to IT Asset Management team."