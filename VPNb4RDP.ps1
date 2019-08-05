# VPNb4RDP powershell script
# Copyright (C) 2019 Robin Collins

Import-Module .\VPNCredentialsHelper.psm1

function PopupMsg ($MessageBody)
{
	$wshell = New-Object -ComObject Wscript.Shell
	$wshell.Popup($MessageBody)

}

## the below code is if we wish to store the passwords in a secure format - currently not implemented as the ability to copy / paste passwords too useful
# create securestring for password
# Read-Host -Prompt "Enter password to be encrypted " -AsSecureString | ConvertFrom-SecureString | Out-File ".\$conectionname.creds.txt"
# $securepassword = ($securestring | ConvertTo-SecureString)
# $plainpassword = (New-Object System.Management.Automation.PSCredential 'N/A', $securepassword).GetNetworkCredential().Password

# import the list of VPN hosts & RDP Servers
$connectionlist = Import-CSV ./VPN_connectionlist.csv

#a function to build a menu of all connections imported from the CSV and prompt user to make selection
function Show-Menu
{
    param (
        [string]$Title = 'Select Host to connect to'
    )
    Clear-Host
    Write-Host "================ $Title ================"

    $Menu = @{}

	foreach ($connection in $connectionlist) 
	{
		Write-Host $connection.id "`: " $connection.Name " Press '"$connection.id"' for this option."
		$Menu.add($connection.id, $connection.id)
	}
	Write-Host "Q :  Press 'Q' to quit."
    $Selection = Read-Host "Please make a selection"

    if ($Selection -eq 'Q') { Return } Else { $Menu.$Selection }

}

function Create-VPN ($VPN)
{
	#setup the VPN credentials based on the menu selection
	$name = $VPN.Name
	$address = $VPN.hostaddress
	$username = $VPN.username
	if ($VPN.domain.Length -gt 15)
	{
		$domain = $VPN.domain.Substring(0,15)
	}
	else 
	{
		$domain = $VPN.domain.Substring(0,$VPN.domain.Length)
	}
	$plainpassword = $VPN.password
	$l2tppsk = $VPN.l2tppsk
	if($VPN.vpnType -eq "l2tp")
	{
		#create the VPN connection in windows dialup manager
		Add-VpnConnection -Name $name -ServerAddress $address -TunnelType L2tp -EncryptionLevel Required -AuthenticationMethod MSChapv2 -L2tpPsk $l2tppsk -Force:$true -RememberCredential:$true -SplitTunneling:$false 

		#assign credentials to the VPN
		Set-VpnConnectionUsernamePassword -connectionname $name -username $username -password $plainpassword -domain $domain
	}
	
	if($VPN.vpnType -eq "pptp")
	{
		#create the VPN connection in windows dialup manager
		Add-VpnConnection -Name $name -ServerAddress $address -TunnelType pptp -EncryptionLevel Required -AuthenticationMethod MSChapv2 -Force:$true -RememberCredential:$true -SplitTunneling:$false 

		#assign credentials to the VPN
		Set-VpnConnectionUsernamePassword -connectionname $name -username $username -password $plainpassword -domain $domain	
	}
	
	return $true
}

function VPN-Stop
{
	$system = $env:windir+"\SYSTEM32\"
	$expr = $system+"rasdial.exe /Disconnect"
	Invoke-Expression -Command $expr
}

function VPN-Start ($VPN)
{
	$system = $env:windir+"\SYSTEM32\"
	$expr = $system+"rasdial.exe `""+$VPN.Name+"`" "+$VPN.username+" `""+$VPN.password+"`""
	Invoke-Expression -Command $expr
}

function isVPN-Connected ($VPN)
{
	$return = $false
	$status = Get-VpnConnection -Name $VPN.Name
	$vpnStatus = $status.ConnectionStatus
	if( $vpnStatus -eq "Connected" )
	{ 
	$return = $true
	}
	return $return
}

function isRDPActive { return Get-Process mstsc -ErrorAction SilentlyContinue }

function Start-RDP ($RDPInfo)
{
	$Server= $RDPInfo.rdphost
	$Port= $RDPInfo.rdpport
	if($RDPInfo.domain)
	{
		#Domain is present, user = domain\rdpuser
		$User= $RDPInfo.domain+"\"+$RDPinfo.rdpuser
	}
	else
	{
		#Domain is not present, user=rdpuser
		$User= $RDPinfo.rdpuser
	}
	$Password= $RDPInfo.rdppassword
	$rmkey = "cmdkey /delete TERMSRV/"+$Server
	Invoke-Expression -Command $rmkey | Out-Null
	$expr1 = "cmdkey /generic:TERMSRV/"+$Server+" /user:"+$User+" /pass:`""+$Password+"`""
	Invoke-Expression -Command $expr1 | Out-Null
#	$expr2 = "mstsc /v`:"+$Server+":"+$Port
#	Invoke-Expression -Command $expr2
	$expr2 = "/v`:"+$Server+":"+$Port
	$app = Start-Process mstsc -ArgumentList $expr2 -passthru
	return $app.ID
}

function LogConnection($data)
{
	$path = ".\"
	$logfile = "VPNb4RDP_connections.log"
	$wUser = $env:UserName
	$wDom = $env:UserDomain
	$wPC = $env:ComputerName
	$date = Get-Date -Format G
	$logentry = $date+" - "+$wDom+":\\"+$wPC+" User:"+$wUser+" connected to "+$data.Name+"`n"
	if (!(Test-Path $path$logfile))
	{
	   New-Item -path $path -name $logfile -type "file" -value $logentry
	   Write-Host "Created new logfile and connection record added"
	}
	else
	{
	  Add-Content -path $path$logfile -value $logentry -Encoding utf8
	  Write-Host "Logfile already exists and connection record appended"
	}
	(gc $path$logfile) | ? {$_.trim() -ne "" } | set-content $path$logfile
}


#Display the menu and set the $UserSelection
$UserSelection = Show-Menu -Title 'VPN / RDP Hosts'
if($UserSelection) 
{
	Clear-Host
	#find the index of the array $connectionlist so we can reference the correct credentials
	$index = [array]::indexof($connectionlist.id, $UserSelection)
	Write-Host "Connecting to "$connectionlist[$index].Name
	Write-Host "Creating the VPN Connection"
	$createVPN = Create-VPN($connectionlist[$index])
	Write-Host "Connecting to VPN"
	$connectVPN = VPN-Start($connectionlist[$index])
	$VPNName = $connectionlist[$index].Name
	Start-Sleep 1
	$isVPNConnected = $false
	do
	{
		Start-Sleep -Milliseconds 100
		$isVPNConnected = isVPN-Connected($connectionlist[$index])
	}
	while ($isVPNConnected -eq $false)
	Write-Host "VPN Connected"
#	$test = isVPN-Connected($connectionlist[$index])
	Write-Host "Connecting to RDP"
	$RDPProcessID = Start-RDP($connectionlist[$index])
	Wait-Process $RDPProcessID
#	Start-Sleep 5
#	do
#	{
#		Start-Sleep -Milliseconds 100
#	}
#	until ((isRDPActive) -eq $null)
	Write-Host "RDP Session finished"
	$VPNStop = VPN-Stop
	Write-Host "VPN Disconnected"
	Remove-VpnConnection -Name $VPNName -Force:$true
	Write-Host "VPN Connection removed"
	LogConnection($connectionlist[$index])
}
else
{
	exit
}

