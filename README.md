# M365-LicenseCheck
This script allows you to connect to Microsoft Online Services and search for users based on their license. 
It also provides an option to export the search results to a CSV file.

## Getting Started

### Prerequisites

Before running this script, make sure you have the following:

* PowerShell
* MSOnline PowerShell module

### Installing

You can install the MSOnline module by running the following command in PowerShell:
```powershell
Install-Module -Name MSOnline
```

## Usage

1. Run M365-LicenseCheck.ps1 with Powershell or download and run the executable from [releases](https://github.com/ckh-rcit/M365-LicenseCheck/releases).

2. Click "connect" button, enter your credentials and then connect to MS Online Services.

3. Account SKUs will Populate the dropdown box by clicking the `Connect` button.

4. Select an AccountSku from the dropdown box and click the `Search` button to display the search results in the DataGridView.

5. Click the `Export` button to export the search results to a CSV file.

## Built With

* PowerShell
* MSOnline PowerShell module
* WinForms

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details.
