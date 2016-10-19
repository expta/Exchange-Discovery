# Discover-ExchangeEnvironment.ps1
A collection of PowerShell scripts that gather useful discovery information about Exchange and Active Directory from a customer environment.

## SYNOPSIS
**Discover-ExchangeEnvironment.ps1** - Discovery reports generation script.

Generates several reports of useful information for the customer environment. Typically this is done before the Design and Planning phase begins so the engineer has a working understanding of the current environment. It's also useful information to refer to after changes have been made.

The script takes no parameters. It calls the other scripts and then gathers all the output files into a single ZIP file called $Org-DiscoveryFiles.zip. Each script is signed so they will run under most Exchange Management Shell environments without having to alter the PowerShell execution policy.

Discover-ExchangeEnvironment.ps1 can also be run by the customer, who can then send the resulting ZIP file to the SPS engineer.

## Usage
To run the discovery scripts, do the following:

1. Extract **Discover-ExchangeEnvironment.zip** to a folder (i.e., C:\Discovery) on the highest version Exchange server in the environment.
2. Run **Discover-ExchangeEnvironment.ps1** from an elevated Exchange Management Shell.

Discover-ExchangeEnvironment.ps1 will run the other scripts and then gather all the output files into a single ZIP file called ***$Org*-DiscoveryFiles.zip.**

## Credits
Written by: Jeff Guillet, SPS Principal Systems Architect | MVP | MCSM | CISSP

* Email:    jeff@expta.com
* Website:  http://expta.com
* Twitter:  http://twitter.com/expta

### Credits for Individual Scripts
* **Adrian Milliner** - Get-BufferHtml.ps1
* **Karsten Palmvig** - GetLogFileUsage.ps1
* **Michael Van Horenbeeck** - Get-VirDirInfo.ps1
* **Neil Johnson** - Messagestats.ps1
* **Paul Cunningham** - Get-ADInfo.ps1, Get-MailboxReport.ps1, Test-ExchangeServerHealth.ps1, Get-EASDevicesReport.ps1, Get-DailyBackupAlerts.ps1, Get-ExchangeServerCertificateReport.ps1
* **Rob Campbell** - Get-EmailStatsPerUser.ps1
* **Serkan Varoglu** - Report-MailboxPermissions.ps1
* **Steve Goodman** - Get-ExchangeEnvironmentReport.ps1

## Change Log
- V1.00, 10/17/2016 - Initial version
- V1.01, 10/19/2016 - Added Get-EASDevicesReport.ps1, Get-DailyBackupAlerts.ps1, Get-ExchangeServerCertificateReport.ps1, Report-MailboxPermissions.ps1; Force exit of discovery if not running from highest version Exchange server
