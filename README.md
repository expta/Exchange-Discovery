# Exchange-Discovery
Jeff Guillet, SPS Principal Systems Architect | MVP | MCSM | CISSP
October 17, 2016

These scripts are used to gather discovery information about Exchange and Active Directory from a client's environment. Typically this is done before the Design and Planning phase begins so the engineer has a working understanding of the current environment. It's also useful information to refer to after changes have been made.

These PowerShell scripts are codesigned so they will run under most Exchange Management Shell environments without having to alter the PowerShell execution policy. The scripts can also be run by the customer, who will send the resulting ZIP file to the SPS engineer.

To run the discovery scripts, do the following:

1. Extract Discover-ExchangeEnvironment.zip to a folder (i.e., C:\Discovery) on the highest version Exchange server in the environment.
2. Run Discover-ExchangeEnvironment.ps1 from an elevated Exchange Management Shell.
  
This script will run the other scripts and then gather all the output files into a single ZIP file called $Org-DiscoveryFiles.zip.
