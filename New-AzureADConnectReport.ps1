<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2016 v5.2.128
	 Created on:   	10/19/2016 10:46 AM
	 Created by:   	Michael Van Horenbeeck
	 Organization: 	VH Consulting & Training
	 Filename:     	New-AzureADConnectReport.ps1
	===========================================================================
	.DESCRIPTION
		Basic Azure AD Connect Reporting script. 
		Create an HTML file with information on the Azure AD Connect configuration.
#>


[CmdletBinding()]
[OutputType([int])]
Param
(
	#Specify the report file path
	[Parameter(Mandatory = $true,
			   ValueFromPipelineByPropertyName = $true,
			   Position = 0)]
	[Alias("ReportPath")]
	[ValidateNotNullOrEmpty()]
	$filePath
)

#Define Variables
$SQLServer = $env:computername #Current version of the script only supports LocalDB instance. External SQL to follow in future release.

#Define Functions
Function Check-Even($num)
{
	[bool]!($num % 2)
} #from http://blogs.technet.com/b/heyscriptingguy/archive/2006/10/19/how-can-i-tell-whether-a-number-is-even-or-odd.aspx

function Exit-Script($msg)
{
	Write-Warning $msg
	Exit
}

#Check Prerequisites
If (-not (Get-Module SQLPS))
{
	Write-Warning "Could not locate SQL PowerShell Module. Please visit https://msdn.microsoft.com/en-us/library/hh245198.aspx for more information."
	exit
}
Else
{
	#Import SQL Module to execute SQL Cmdlets
	Import-Module SQLPS
}


#Determine Registry path info
$regPaths = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MicrosoftAzureADConnectionTool", "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{337E88B3-6961-420C-BF5D-FA1FDF73AA7C}"
$regPaths | %{
	if ((Test-Path $_) -eq $true)
	{
		Try
		{
			$dirsyncVersion = (Get-ItemProperty $_ -ErrorAction STOP).DisplayVersion
		}
		Catch
		{
			Exit-Script "Could not fetch registry information. Aborting script."
		}
	}
	else
	{
		Exit-Script "Could not fetch registry information. Aborting script."
	}
}

if (Test-Path "HKLM:\SOFTWARE\Microsoft\MSOLCoExistence\CurrentVersion")
{
	Try
	{
		$DirsyncPath = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\MSOLCoExistence\CurrentVersion" -ErrorAction Stop).InstallationPath
	}
	Catch
	{
		Exit-Script "Could not determine Azure AD Sync Installation path. Aborting script."
	}
}
Else
{
	Exit-Script "Could not determine Azure AD Sync Installation path. Aborting script."
}

#Get SQL information
if (Test-Path "HKLM:\SYSTEM\CurrentControlSet\services\ADSync\Parameters")
{
	Try
	{
		$SQLInstance = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\services\ADSync\Parameters' -ErrorAction Stop).SQLInstance
		if (Test-Path "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server Local DB\Shared Instances\ADSync")
		{
			$ADSyncInstance = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server Local DB\Shared Instances\ADSync' -ErrorAction Stop).InstanceName
			$MSOLInstance = ("np:\\.\pipe\" + $ADSyncInstance + "\tsql\query")
			$SQLVersion = Invoke-Sqlcmd -ServerInstance $MSOLInstance -Query "SELECT SERVERPROPERTY('productversion'), SERVERPROPERTY ('productlevel'), SERVERPROPERTY ('edition')" -ErrorAction Stop
		}
		Else
		{
			Exit-Script "Could not fetch DB information. Aborting script."
		}

	}
	Catch
	{
		Exit-Script "Could not fetch DB information. Aborting script."
	}
}
else
{
	Exit-Script "Could not fetch DB information. Aborting script."
}

#Get AD Management Agent information
Try
{
	$ADMAxml = Invoke-Sqlcmd -ServerInstance $MSOLInstance -Query "SELECT [ma_id] ,[ma_name] ,[private_configuration_xml],[ma_type] FROM [ADSync].[dbo].[mms_management_agent]" | ? { $_.ma_type -eq 'AD' } -ErrorAction Stop
	$individualADMAgent = @()
	$MaName = @()
	
	foreach ($ADMAgent in $ADMAxml) #fetching MA information from SQL DB
	{
		[xml]($ADMAgent | select -Expand private_configuration_xml)
		$individualADMAgent += [xml]($ADMAgent | select -Expand private_configuration_xml)
		$maName += ($ADMAgent | select ma_name)
	}
}
Catch
{
	Exit-Script "Could not fetch Management Agent information. Aborting script."
}

#Get DirSync Database Info
$SQLDirSyncInfo = Invoke-Sqlcmd -ServerInstance $MSOLInstance -Query "SELECT DB_NAME(database_id) AS DatabaseName, Name AS Logical_Name, Physical_Name, (size*8)/1024 SizeMB FROM sys.master_files WHERE DB_NAME(database_id) = 'AdSync'"
$DirSyncDB = $SQLDirSyncInfo | ? { $_.Logical_Name -eq 'ADSync' }
$DirSyncLog = $SQLDirSyncInfo | ? { $_.Logical_Name -eq 'ADSync_log' }

#Get connector space info (optional)
$ADMA = Invoke-Sqlcmd -ServerInstance $MSOLInstance -Query "SELECT [ma_id] ,[ma_name],[ma_type] FROM [ADSync].[dbo].[mms_management_agent] WHERE ma_type = 'AD'"
$AzureMA = Invoke-Sqlcmd -ServerInstance $MSOLInstance -Query "SELECT [ma_id] ,[ma_name],[subtype],[private_configuration_xml] FROM [ADSync].[dbo].[mms_management_agent] WHERE subtype = 'Windows Azure Active Directory (Microsoft)'"
$UsersFromBothMAs = Invoke-Sqlcmd -ServerInstance $MSOLInstance -Query "SELECT [ma_id] ,[rdn] FROM [ADSync].[dbo].[mms_connectorspace] WHERE object_type = 'user'"
$AzureUsers = $UsersFromBothMAs | ? { $_.ma_id -eq $AzureMA.ma_id }
#$ADUsers = $UsersFromBothMAs | ? {$_.ma_id -eq $ADMA.ma_id}



#Get DirSync Run History
$SyncHistory = Invoke-Sqlcmd -ServerInstance $MSOLInstance -Query "SELECT [step_result] ,[end_date] ,[stage_no_change] ,[stage_add] ,[stage_update] ,[stage_rename] ,[stage_delete] ,[stage_deleteadd] ,[stage_failure] FROM [ADSync].[dbo].[mms_step_history]" | sort end_date -Descending

#GetDirSync interval (3 hours is default)
#AADSync uses Task Scheduler to get interval
Try
{
	$SyncTimeInterval = ((Get-ScheduledTaskInfo "Azure AD Sync Scheduler" -ErrorAction Stop).NextRunTime - (Get-ScheduledTaskInfo "Azure AD Sync Scheduler" -ErrorAction Stop).LastRunTime).TotalMinutes
}
Catch
{
	$SyncTimeInterval = (Get-ADSyncScheduler).CurrentlyEffectiveSyncCycleInterval.TotalMinutes
}


#Generate Output
cls

#HTML HEADERS
$html += "<html>"
$html += "<head>"
$html += "<style type='text/css'>"
$html += "body {font-family:verdana;font-size:10pt}"
$html += "table {border:0px solid #000000;font-family:verdana; font-size:10pt;cellspacing:1;cellspacing:0}"
$html += "tr.color {background-color:#00A2E8;color:#FFFFFF;font-weight:bold}"
$html += "tr.title {background-color:#E5E5E5;text-decoration:underline}"
$html += "font.value {color:#808080}"
$html += "</style>"
$html += "</head>"
$html += "<body>"

$html += "<b>Azure AADSync Report Info</b><br/>"
$html += "Date: <font class='value'>" + (Get-Date) + "</font></br>"
$html += "Server: <font class='value'>" + $env:computername + "</font></br>"
$html += "<p>&nbsp;</p>"

$html += "<b>Account info</b><br/>"

#Get Account Info for each domain:
#$ServiceAccountGuess = (((gci 'hkcu:Software\Microsoft\MSOIdentityCRL\UserExtendedProperties' | select PSChildName)[-1]).PSChildName -split ':')[-1]
#Update to use info from SQL
$ServiceAccountGuess = (([xml]$AzureMA.private_configuration_xml | select -ExpandProperty MaConfig | select -ExpandProperty parameter-values).parameter | ? Name -eq "username").'#text'

$i = 0
foreach ($agent in $individualADMAgent)
{
	$ADServiceAccountUser = $Agent.'adma-configuration'.'forest-login-user'
	$ADServiceAccountDomain = $Agent.'adma-configuration'.'forest-login-domain'
	$ADServiceAccount = $ADServiceAccountDomain + "\" + $ADServiceAccountUser
	
	$html += "Active Directory Service Account <font class='value'>" + $ADServiceAccountDomain + ": " + $ADServiceAccount + "</font>"
	$html += "<br/>"
	#Write-Host "Active Directory Service Account $ADServiceAccountDomain : " -F Cyan -NoNewline ; Write-Host $ADServiceAccount -F DarkCyan
	
	$i++
}
$html += "Azure Service Account Guess: <font class='value'>" + $ServiceAccountGuess + "</font>"
#Write-Host "Azure Service Account Guess: " -F Cyan -NoNewline ; Write-Host $ServiceAccountGuess -F DarkCyan
$html += "<p>&nbsp;</p>"

$html += "<b>Azure AD Sync Info</b><br/>"
$html += "Version: <font class='value'>" + $DirsyncVersion + "</font><br/>"
$html += "Path: <font class='value'>" + $DirsyncPath + "</font><br/>"
$html += "Sync Interval (Minutes): <font class='value'>" + $SyncTimeInterval + "</font>"
$html += "<p>&nbsp;</p>"


$html += "<b>User Info:</b><br/>"

foreach ($ad in $ADMA)
{
	$html += "User in AD " + $ad.ma_name + ": <font class='value'>" + ($UsersFromBothMAs | ? { $_.ma_id -eq $ad.ma_id }).count + "</font><br/>"
	#Write-Host "Users in AD"$ad.ma_name": " -F Cyan -NoNewLine ; Write-Host ($UsersFromBothMAs | ? {$_.ma_id -eq $ad.ma_id}).count -ForegroundColor DarkCyan
}

$html += "Users in Azure Connector Space: <font class='value'>" + $AzureUsers.Count + "</font><br/>"
$html += "Total users: <font class='value'>" + $UsersFromBothMAs.Count + "</font><br/>"
$html += "<p>&nbsp;</p>"

$html += "<b>SQL Info</b><br/>"
$html += "Version: <font class='value'>" + $SQLVersion.Column1 + " " + $SQLVersion.Column2 + " " + $SQLVersion.Column3 + "</font><br/>"
$html += "Instance: <font class='value'>" + $MSOLInstance + "</font><br/>"
$html += "Database Location: <font class='value'>" + $DirSyncDB.Physical_Name + "</font><br/>"
$html += "Database Size: <font class='value'>" + $DirSyncDB.SizeMB + "MB</font><br/>"
$html += "Database Log Size: <font class='value'>" + $DirSyncLog.SizeMB + "MB</font>"
$html += "<p>&nbsp;</p>"

$html += "<b>Most Recent Sync Activity</b><br/>"
$html += "<i>(For more detail, launch:" + $DirsyncPath + "\UIShell\miisclient.exe)<br/><br/>"
$html += "<table>"
$html += "<tr class='title'>"
$html += "<td width='250'>"
$html += "Date"
$html += "</td>"
$html += "<td>"
$html += "Result"
$html += "</td>"
$html += "</tr>"
for ($j = 0; $j -ne 9; $j++)
{
	if (check-even $j -eq $true)
	{
		$color = "#C3C3C3"
	}
	else
	{
		$color = "#E5E5E5"
	}
	$html += "<tr style='background-color:$color'>"
	$html += "<td>"
	$html += ($SyncHistory[$j].end_date).ToLocalTime()
	$html += "</td>"
	$html += "<td>"
	$html += $SyncHistory[$j].step_result + "<br/>"
	$html += "</td>"
	$html += "</tr>"
}
$html += "</table>"
$html += "<p>&nbsp;</p>"

$html += "</body>"
$html += "</html>"

$filePath.TrimEnd("\")
$html | Out-File $filePath"\AADSyncInfo.html"

Remove-Variable html
