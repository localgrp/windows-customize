<#
.SYNOPSIS
    A script to add or remove Windows features, applications, and packages based on text files.

.DESCRIPTION
    This script provides a way to add or remove Windows features, applications, and packages from a Windows operating system. It takes as input text files containing lists of items to be added or removed, and performs the requested operation on each item. The script logs its activities to a file and to the console.

.PARAMETER Remove
    Switch to remove specified features, applications, or packages.

.PARAMETER Add
    Switch to add specified features, applications, or packages.

.PARAMETER WingetFile
    Path to the text file containing a list of Winget apps to be added or removed. The text file should list each Winget app Id on a new line.

.PARAMETER CapabilitiesFile
    Path to the text file containing a list of Windows Capabilities to be added or removed. The text file should list each Windows capability Name on a new line.

.PARAMETER OptionalFeatureFile
    Path to the text file containing a list of Windows Optional Features to be added or removed. The text file should list each Optional Feature FeatureName on a new line. 

.PARAMETER AppXPackageFile
    Path to the text file containing a list of AppX Packages to be removed. The text file should list each AppxPackage PackageFullName on a new line.

.PARAMETER AppXProvisionedPackageFile
    Path to the text file containing a list of AppX Provisioned Packages to be removed. The text file should list each AppxProvisionedPackage PackageName on a new line.

.PARAMETER LogPath
    Path to the log file.

.PARAMETER Silent
    Suppress log output to the console.

.EXAMPLE
    .\Modify-WindowsComponents.ps1 -Remove -WingetFile "C:\Temp\WingetApps.txt"

    This will remove all Winget apps listed in the WingetApps.txt file.

.EXAMPLE
    .\Modify-WindowsComponents.ps1 -Add -CapabilitiesFile "C:\Temp\Capabilities.txt"

    This will add all Windows Capabilities listed in the Capabilities.txt file.

.EXAMPLE
    .\Modify-WindowsComponents.ps1 -Remove -OptionalFeature "C:\Temp\OptionalFeatures.txt" -AppXPackageFile "C:\Temp\AppXPackages.txt" -LogPath "C:\Temp\LogFile.log" -Silent

    This will remove all Optional Features listed in the OptionalFeatures.txt file as well as AppXPackages listed in the AppXPackages.txt file. It will also log activity to the LogFile.log and suppress log output to the console.

.NOTES
    - This script requires administrative privileges, the Add or Remove parameter, and at least one of the file parameters specified with a valid path to a text file.
    - This script will not attempt to add AppX packages and will ignore AppX files when specified with the Add parameter. It is recommended to use Winget to add software to the operating system.
#>

param(
    [switch]$Remove,
    [switch]$Add,

    [Parameter(Mandatory=$false, HelpMessage="Path to the winget app list text file.")]
    [ValidateNotNullOrEmpty()]
    [string]$WingetFile,

    [Parameter(Mandatory=$false, HelpMessage="Path to the Windows capabilities text file.")]
    [ValidateNotNullOrEmpty()]
    [string]$CapabilitiesFile,

    [Parameter(Mandatory=$false, HelpMessage="Path to the optional features text file.")]
    [ValidateNotNullOrEmpty()]
    [string]$OptionalFeatureFile,

    [Parameter(Mandatory=$false, HelpMessage="Path to the AppXPackage list text file.")]
    [ValidateNotNullOrEmpty()]
    [string]$AppXPackageFile,

    [Parameter(Mandatory=$false, HelpMessage="Path to the AppXProvisionedPackage list text file.")]
    [ValidateNotNullOrEmpty()]
    [string]$AppXProvisionedPackageFile,

    [Parameter(Mandatory=$false, HelpMessage="Path to the log file written, including file name. Default is written to %SYSTEMROOT%\Temp\Modify-WindowsComponents.log.")]
    [ValidateNotNullOrEmpty()]
    [string]$LogPath,

    [Parameter(Mandatory=$false, HelpMessage="Supress log output to console.")]
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
        $LogFileName = "Modify-WindowsComponents.log"
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

function Test-IsAdmin {
    $userIsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    return $userIsAdmin
}

function Test-WingetInstalled {
    $wingetPath = (Get-Command winget -ErrorAction SilentlyContinue).Source

    if (-not $wingetPath) { 
        Write-LogEntry -Message "Winget is not installed. Attempting to register..."
        try {
            Add-AppxPackage -RegisterByFamilyName -MainPackage Microsoft.DesktopAppInstaller_8wekyb3d8bbwe -ErrorAction Stop
            Write-LogEntry -Message "Winget registered successfully."
        } catch {
            Write-LogEntry -Message "Failed to register Winget: $($_.Exception.Message)" -IsError $true
            Exit 1
        }     
        
        $wingetPath = (Get-Command winget -ErrorAction SilentlyContinue).Source
        if (-not $wingetPath) {
            Write-LogEntry -Message "Winget registration did not solve the issue, attempting to update DesktopAppInstaller package from the Windows Store."
            
            # Synchronously triggers store updates for a select set of apps. You should run this in legacy powershell.exe, as some of the code has problems in pwsh on older OS releases.
            # https://github.com/microsoft/winget-cli/discussions/1738#discussioncomment-5484927
            try
            {
                # https://fleexlab.blogspot.com/2018/02/using-winrts-iasyncoperation-in.html
                Add-Type -AssemblyName System.Runtime.WindowsRuntime
                $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
                function Await($WinRtTask, $ResultType) {
                    $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
                    $netTask = $asTask.Invoke($null, @($WinRtTask))
                    $netTask.Wait(-1) | Out-Null
                    $netTask.Result
                }
        
                # https://docs.microsoft.com/uwp/api/windows.applicationmodel.store.preview.installcontrol.appinstallmanager?view=winrt-22000
                # We need to tell PowerShell about this WinRT API before we can call it...
                [Windows.ApplicationModel.Store.Preview.InstallControl.AppInstallManager,Windows.ApplicationModel.Store.Preview,ContentType=WindowsRuntime] | Out-Null
                $appManager = New-Object -TypeName Windows.ApplicationModel.Store.Preview.InstallControl.AppInstallManager
        
                $appsToUpdate = @(
                    "Microsoft.DesktopAppInstaller_8wekyb3d8bbwe"
                )
        
                foreach ($app in $appsToUpdate)
                {
                    try
                    {
                        Write-LogEntry -Message "Update DesktopAppInstaller: Requesting an update for $app..."
                        $updateOp = $appManager.UpdateAppByPackageFamilyNameAsync($app)
                        $updateResult = Await $updateOp ([Windows.ApplicationModel.Store.Preview.InstallControl.AppInstallItem])
                        while ($true)
                        {
                            if ($null -eq $updateResult)
                            {
                                Write-LogEntry -Message "Update DesktopAppInstaller: Update is null. It must already be completed (or there was no update)..."
                                break
                            }
        
                            if ($null -eq $updateResult.GetCurrentStatus())
                            {
                                Write-LogEntry -Message "Update DesktopAppInstaller: Current status is null."
                                break
                            }
        
                            if (!$Silent) { Write-Host "Update DesktopAppInstaller: Update $($updateResult.GetCurrentStatus().PercentComplete)% completed." }
                            if ($updateResult.GetCurrentStatus().PercentComplete -eq 100)
                            {
                                Write-LogEntry -Message "Update DesktopAppInstaller: Install completed ($app)"
                                break
                            }
                            Start-Sleep -Seconds 3
                        }
                    }
                    catch [System.AggregateException]
                    {
                        # If the thing is not installed, we can't update it. In this case, we get an
                        # ArgumentException with the message "Value does not fall within the expected
                        # range." I cannot figure out why *that* is the error in the case of "app is
                        # not installed"... perhaps we could be doing something different/better, but
                        # I'm happy to just let this slide for now.
                        $problem = $_.Exception.InnerException # we'll just take the first one
                        Write-LogEntry -Message "Update DesktopAppInstaller: Error updating app $app : $problem" -IsError $true
                    }
                    catch
                    {
                        Write-LogEntry -Message "Update DesktopAppInstaller: Unexpected error updating app $app : $_" -IsError $true
                    }
                }
            }
            catch
            {
                Write-LogEntry -Message "Update DesktopAppInstaller: Problem updating store apps: $_" -IsError $true
            }

            try {
                Add-AppxPackage -RegisterByFamilyName -MainPackage Microsoft.DesktopAppInstaller_8wekyb3d8bbwe -ErrorAction Stop
                Write-LogEntry -Message "Winget registered successfully."
            } catch {
                Write-LogEntry -Message "Failed to register Winget: $($_.Exception.Message)" -IsError $true
                Exit 1
            }
        }
    }
}

if (-not (Test-IsAdmin)) {
    Write-Error "This script must be run as an administrator."
    Exit 1
}

if (!$Remove -and !$Add) {
    Write-Error "You must specify either -Remove or -Add switch."
    Exit 1
}

if ($Remove -and $Add) {
    Write-Error "You cannot specify both -Remove and -Add switches."
    Exit 1
}

# Custom validation to ensure at least one file is provided
if (-not ($WingetFile -or $CapabilitiesFile -or $OptionalFeatureFile -or $AppXPackageFile -or $AppXProvisionedPackageFile)) {
    Write-Error "At least one of the following parameters must be provided: AppXPackageFile, AppXProvisionedPackageFile, WingetAppFile, CapabilitiesFile, OptionalFeatureFile"
    Exit 1
}

if ($WingetFile) { Test-WingetInstalled }

Write-LogEntry -Message "** START RUN **"

$FileParameters = @{
    WingetFile                 = 'WingetApps';
    AppXPackageFile            = 'Packages';
    AppXProvisionedPackageFile = 'ProvisionedPackages';
    CapabilitiesFile           = 'Capabilities';
    OptionalFeatureFile        = 'OptionalFeatures';
}

foreach ($FileParam in $FileParameters.Keys) {
    $filePath = Get-Variable -Name $FileParam -ValueOnly
    if ($filePath) {
        try {
            $fileContent = Get-Content -Path $filePath -ErrorAction Stop
            Set-Variable -Name $FileParameters[$FileParam] -Value $fileContent
            Write-LogEntry -Message "Populated $($FileParameters[$FileParam]) with $($fileContent.Count) items from $filePath"
        } catch [System.Exception] {
            Write-LogEntry -Message "Error reading $filePath file: $($_.Exception.Message)" -IsError $true
            Exit 1
        }
    }
}

$FeatureTypes = @(
    @{
        FeatureType = 'WingetApp';
        Items = $WingetApps;
        AddCommand = { param($Feature) winget install $Feature --silent --accept-package-agreements --accept-source-agreements | Out-Null };
        RemoveCommand = { param($Feature) winget uninstall $Feature --silent | Out-Null };
    },
    @{
        FeatureType = 'AppXPackage';
        Items = $Packages;
        AddCommand = $null;
        RemoveCommand = {
            param($Feature)
            $InstalledPackage = Get-AppxPackage | Where-Object { $_.Name -eq ($Feature -split '_')[0] }
            if ($InstalledPackage) {
                Remove-AppxPackage -Package $InstalledPackage.PackageFullName -ErrorAction Stop | Out-Null
            }
        };        
    },
    @{
        FeatureType = 'AppXProvisionedPackage';
        Items = $ProvisionedPackages;
        AddCommand = $null;
        RemoveCommand = {
            param($Feature)
            $InstalledPackage = Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -like "$($Feature)*" }
            if ($InstalledPackage) {
                Remove-AppxProvisionedPackage -PackageName $InstalledPackage.PackageName -Online -ErrorAction Stop | Out-Null
            }
        };
    },
    @{
        FeatureType = 'WindowsCapability';
        Items = $Capabilities;
        AddCommand = { param($Feature) Add-WindowsCapability -Name $Feature -Online -ErrorAction Stop | Out-Null };
        RemoveCommand = { param($Feature) Remove-WindowsCapability -Name $Feature -Online -ErrorAction Stop | Out-Null };
    },
    @{
        FeatureType = 'OptionalFeature';
        Items = $OptionalFeatures;
        AddCommand = { param($Feature) Enable-WindowsOptionalFeature -Online -FeatureName $Feature -NoRestart -ErrorAction Stop | Out-Null };
        RemoveCommand = { param($Feature) Disable-WindowsOptionalFeature -Online -FeatureName $Feature -NoRestart -ErrorAction Stop | Out-Null };
    }
)

foreach ($FeatureType in $FeatureTypes) {
    $Features = $FeatureType.Items
    $AddCommand = $FeatureType.AddCommand
    $RemoveCommand = $FeatureType.RemoveCommand
    $Type = $FeatureType.FeatureType

    if ($Features) {
        foreach ($Feature in $Features) {
            Write-LogEntry -Message "Processing $($Type): $($Feature)"

            if ($Remove -and $RemoveCommand) {
                try {
                    Write-LogEntry -Message "Removing $($Type): $($Feature)"
                    $RemoveCommand.Invoke($Feature)
                } catch [System.Exception] {
                    Write-LogEntry -Message "Removing $Type '$($Feature)' failed: $($_.Exception.Message)" -IsError $true
                }
            }
            
            if ($Add -and $AddCommand) {
                try {
                    Write-LogEntry -Message "Adding $($Type): $($Feature)"
                    $AddCommand.Invoke($Feature)
                } catch [System.Exception] {
                    Write-LogEntry -Message "Adding $Type '$($Feature)' failed: $($_.Exception.Message)" -IsError $true
                }
            }            
        }
    }
}

Write-LogEntry -Message "** END RUN **"
