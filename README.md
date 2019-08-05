# VPNb4RDP
 A powershell script to create a menu of multiple VPN and RDP servers. Select an item from the menu to initiate a VPN Connection before connecting to an RDP server

Edit the connection details in the VPN_connectionlist.csv file.

Launch the VPNb4RDP.bat batch file to run the powershell script.

Includes the VPNCredentialsHelper.psm1 file that can be installed with Import-Module VPNCredentialsHelper, however it is included here so that you do not need to elevate privaledges to launch the script.

Also includes a Registry fix for Windows 10 and L2TP connections (where both systems are behind a NAT firewall)
