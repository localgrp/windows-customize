# Windows Customize Scripts

This repository contains PowerShell scripts to automate common administrative tasks in a Windows environment.

## Compatibility & Warranty

The scripts in this repository were tested with Windows 11 22H2. While it may be possible to use them with other versions of Windows, it is recommended to exercise caution when using them on other versions.

Please ensure you have a backup of your system and review the scripts before running them on any version of Windows.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER 
DEALINGS IN THE SOFTWARE.

## Prerequisites

- PowerShell 5.1 or later.

## Modify-WindowsComponents.ps1

The `Modify-WindowsComponents.ps1` script allows you to add or remove Windows features, applications, and packages based on text files. It reads text files containing lists of items to be added or removed and performs the requested operations.

### Usage

```
.\Modify-WindowsComponents.ps1 -Remove -WingetFile "C:\Temp\WingetApps.txt"
```
- The script requires administrative privileges, the `Add` or `Remove` parameter, and at least one of the file parameters specified with a valid path to a text file.
- The respective text files (`WingetFile`, `CapabilitiesFile`, `OptionalFeatureFile`, `AppXPackageFile`, `AppXProvisionedPackageFile`) should contain a list of items to be added or removed, with each item on a new line.
- The script provides the option to specify a `-LogPath` and use the `-Silent` parameter to suppress log output to the console.
- The script will not attempt to add AppX packages and will ignore AppX files when specified with the `Add` parameter. It is recommended to use Winget to add software to the operating system.
- If a Winget file is specified, the script will check that the Winget package is installed and attempt to register it if the package is not found. The package may be missing due to a delay in the installation of the AppInstaller package the first time a user has logged in.
- Examine the `samples` folder in the repository for sample CSV files.
- Examine the PowerShell command help for more information.

## Add-RegistryItems.ps1

The `Add-RegistryItems.ps1` script allows you to add registry items from a CSV file. It reads a CSV file containing registry item information and creates new registry items based on the provided data.

### Usage

```
.\Add-RegistryItems.ps1 -CsvFilePath "C:\RegistryItems.csv"
```
- The script requires administrative privileges. 
- The `CsvFilePath` parameter specifies the path to the CSV file containing the registry item information.
- The script provides the option to specify a `-LogPath` and use the `-Silent` parameter to suppress log output to the console.
- The specified registry path and keys will be created if they do not exist.
- The CSV file does not require headers, however the rows should contain values in the following order: registry path, registry item name, registry item type, and registry item value.
- All values should follow PowerShell conventions, such as using the PS drive for the path (e.g., "HKLM:\SOFTWARE\Microsoft\Windows") and supported registry types (e.g., "String", "ExpandString", "Binary", "DWord", "QWord", "MultiString").
- Examine the `samples` folder in the repository for sample text files.
- Examine the PowerShell command help for more information.

## License

This repository is licensed under the [MIT License](LICENSE).
