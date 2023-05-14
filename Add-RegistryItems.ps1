<#
.SYNOPSIS
    Adds registry items from a CSV file.

.DESCRIPTION
    This script reads a CSV file containing registry item information and creates new registry items based on the provided data.

.PARAMETER CsvFilePath
    Specifies the path to the CSV file containing the registry item information. This parameter is mandatory. 

.PARAMETER Silent
    Indicates whether the script should run silently without displaying log messages on the console. By default, log messages are displayed on the console. This parameter is optional.

.PARAMETER LogPath
    Path to the log file. If not specified, a default log file path is used. This parameter is optional.

.EXAMPLE
    Add-RegistryItems -CsvFilePath "C:\RegistryItems.csv"
    Adds registry items based on the information provided in the "RegistryItems.csv" file.

.EXAMPLE
    Add-RegistryItems -CsvFilePath "C:\RegistryItems.csv" -Silent -LogPath "C:\Temp\Registry.log"
    Adds registry items based on the information provided in the "RegistryItems.csv" file, supresses log output to the console, and writes the log file to the specified path.

.NOTES
    - This script requires administrative privileges.
    - The specified registry path and keys will be created if they do not exist.
    - The CSV file does not require headers. The rows should contain values in the following order: registry path, registry item name, registry item type, and registry item value.
    - All values should follow PowerShell conventions, such as using the PS drive for the path (e.g., "HKLM:\SOFTWARE\Microsoft\Windows") and supported registry types (e.g., "String", "ExpandString", "Binary", "DWord", "QWord", "MultiString").

#>

param (
    [Parameter(Mandatory=$true, HelpMessage="Path to the source CSV file containing the items to add to the registry.")]
    [string]$CsvFilePath,

    [Parameter(Mandatory=$false, HelpMessage="Path to the log file written, including file name. Default is written to %SYSTEMROOT%\Temp\Add-RegistryItems.log.")]
    [ValidateNotNullOrEmpty()]
    [string]$LogPath,

    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [switch]$Silent = $false
)

function Write-LogEntry {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [bool]$IsError = $false,
        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$LogPath,
        [Parameter(Mandatory=$false)]
        [bool]$WriteToConsole = -not $Silent
    )

    If (!$LogPath) {
        $LogFileName = "Add-RegistryItems.log"
        $LogFileFolder = $env:SystemRoot
        $LogPath = Join-Path -Path $LogFileFolder -ChildPath "Temp\$($LogFileName)"
    }

    if ($IsError) { $Status = "ERROR" } else { $Status = "INFO" }
    $Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $LogEntry = "[$Timestamp] [$Status] $Message"

    try {
        Add-Content -Path $LogPath -Value $LogEntry
    } 
    catch [System.Exception] {
        Write-Error "Error writing to the log file: $($_.Exception.Message)"
    }
    if ($WriteToConsole) { Write-Host $LogEntry }
}

function New-RegistryItem {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("String", "ExpandString", "Binary", "DWord", "QWord", "MultiString")]
        [string]$Type,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Value
    )

    if (-not (Test-Path -Path $Path)) {
        try {
            New-Item -Path $Path -ItemType Registry -Force -ErrorAction Stop | Out-Null
        } catch {
            Write-LogEntry -Message "Failed to create new path: $($item.Path) Error: $($_)" -IsError $true
        }
    }
    try {
        New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $Type -Force -ErrorAction Stop | Out-Null
        Write-LogEntry -Message "Created new item: $($item)" -IsError $false
    } catch {
        Write-LogEntry -Message "Failed to create item: $($item) Error: $($_)" -IsError $true
    }
}

function Test-IsAdmin {
    $userIsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    return $userIsAdmin
}

if (-not (Test-IsAdmin)) {
    Write-Error "This script must be run as an administrator."
    Exit 1
}

if (Test-Path -Path $CsvFilePath) {
    $CsvContent = Import-Csv -Path $CsvFilePath -Header "Path", "Name", "Type", "Value"
    foreach ($item in $CsvContent) {
        New-RegistryItem -Path $item.Path -Name $item.Name -Value $item.Value -Type $item.Type
    }
} else {
    Write-LogEntry -Message "CSV file not found: $CsvFilePath" -IsError $true
}