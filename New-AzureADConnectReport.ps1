<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2016 v5.2.128
	 Created on:   	10/19/2016 10:46 AM
	 Created by:   	Michael Van Horenbeeck
	 Organization: 	VH Consulting & Training
	 Filename:     	New-AzureADConnectReport.ps1
	===========================================================================

	Before you can run the script, you must install SQL PowerShell on the AADSync machine first. 
	Older versions of DirSync had this installed by default, but it seems that AADSync/AADConnect does not. 
	
	To install the SQL PS module, you must install the following components separately:

	- Microsoft® System CLR Types for Microsoft® SQL Server® 2012
	- Microsoft® SQL Server® 2012 Shared Management Objects
	- Microsoft® Windows PowerShell Extensions for Microsoft® SQL Server® 2012
	
	The binaries can be installed from the installation instructions on the following page: http://www.microsoft.com/en-us/download/details.aspx?id=29065
	
	.DESCRIPTION
	This script will create a basic HTML report with some information on Azure AD Connect.
	For now, the script only supports the LocalDB SQL Instance; support for an external SQL DB is coming in a future version.
	Script idea and initial source code is based on the v1 script from Mike Crowley (https://mikecrowley.us/2013/10/16/dirsync-report/). 

	.PARAMETER FilePath
    Specifies location where the report should be stored. 
	Location should be specified in the following format: "C:\Temp"

	.EXAMPLE
	New-AzureADConnectReport.ps1 -FilePath "C:\SomeFolderPath"
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
	$FilePath
)

#Define Variables
$SQLServer = $env:computername #Current version of the script only supports LocalDB instance. External SQL to follow in future release.
$date = Get-Date -Format yyyy-MM-dd
$defaultFilePath = "C:\Temp"

#Define Functions
Function Check-Even($num)
{
	[bool]!($num % 2)
} #from http://blogs.technet.com/b/heyscriptingguy/archive/2006/10/19/how-can-i-tell-whether-a-number-is-even-or-odd.aspx

Function Exit-Script($msg)
{
	Write-Warning $msg
	Exit
}

#Check Prerequisites
Write-Host "Checking script prerequisites" -ForegroundColor Cyan
If (-not (Get-Module SQLPS -ListAvailable))
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
Write-Host "Fetching registry information" -ForegroundColor Cyan
$regPathSuccess = 0
#Registry path for AAD Connect changes with newer versions, so this $regpaths array will need to be updated sometimes
$regPaths = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MicrosoftAzureADConnectionTool",
                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{337E88B3-6961-420C-BF5D-FA1FDF73AA7C}",
                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{69E51737-DAAC-40E0-BBD6-816345D62A5A}",
                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{1D8D16FF-96F1-48AA-8971-C7D4BD1D66F1}"
                )
$regPaths | %{
	if ((Test-Path $_) -eq $true)
	{
		$regPathSuccess = 1
		Try
		{
			$dirsyncVersion = (Get-ItemProperty $_ -ErrorAction STOP).DisplayVersion
		}
		Catch
		{
			Exit-Script "Could not fetch registry information. Aborting script."
		}
	}
}
if ($regPathSuccess -eq 0)
{
	Exit-Script "Could not fetch registry information. Aborting script."
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

#Get SQL Location information
if (Test-Path "HKLM:\SYSTEM\CurrentControlSet\services\ADSync\Parameters")
{
	Try
	{
		Write-Host "Fetching information from SQL" -ForegroundColor Cyan
		$SQLInstance = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\services\ADSync\Parameters' -ErrorAction Stop).SQLInstance
		if (Test-Path "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server Local DB\Shared Instances\ADSync")
		{
			$ADSyncInstance = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server Local DB\Shared Instances\ADSync' -ErrorAction Stop).InstanceName
			$MSOLInstance = ("np:\\.\pipe\" + $ADSyncInstance + "\tsql\query")
			$SQLVersion = Invoke-Sqlcmd -ServerInstance $MSOLInstance -Query "SELECT SERVERPROPERTY('productversion'), SERVERPROPERTY ('productlevel'), SERVERPROPERTY ('edition')" -ErrorAction Stop
			
			#Get DirSync Database Info
			$SQLDirSyncInfo = Invoke-Sqlcmd -ServerInstance $MSOLInstance -Query "SELECT DB_NAME(database_id) AS DatabaseName, Name AS Logical_Name, Physical_Name, (size*8)/1024 SizeMB FROM sys.master_files WHERE DB_NAME(database_id) = 'AdSync'" -ErrorAction Stop
			$DirSyncDB = $SQLDirSyncInfo | ? { $_.Logical_Name -eq 'ADSync' } #get information about the DB file
			$DirSyncLog = $SQLDirSyncInfo | ? { $_.Logical_Name -eq 'ADSync_log' } #get information about the DB log file(s)
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
Write-Host "Fetching information from Management Agents" -ForegroundColor Cyan
Try
{
	$ADMAxml = Invoke-Sqlcmd -ServerInstance $MSOLInstance -Query "SELECT [ma_id] ,[ma_name] ,[private_configuration_xml],[ma_type] FROM [ADSync].[dbo].[mms_management_agent]" | ? { $_.ma_type -eq 'AD' } -ErrorAction Stop
	$individualADMAgent = @()
	$MaName = @()
	
	foreach ($ADMAgent in $ADMAxml) #fetching MA information from SQL DB
	{
		#[xml]($ADMAgent | select -Expand private_configuration_xml)
		$individualADMAgent += [xml]($ADMAgent | select -Expand private_configuration_xml)
		$maName += ($ADMAgent | select ma_name)
	}
}
Catch
{
	Exit-Script "Could not fetch Management Agent information. Aborting script."
}

#Get connector space info
Try
{
	Write-Host "Fetching information from Connector Space" -ForegroundColor Cyan
	$ADMA = Invoke-Sqlcmd -ServerInstance $MSOLInstance -Query "SELECT [ma_id] ,[ma_name],[ma_type] FROM [ADSync].[dbo].[mms_management_agent] WHERE ma_type = 'AD'" -ErrorAction Stop
	$AzureMA = Invoke-Sqlcmd -ServerInstance $MSOLInstance -Query "SELECT [ma_id] ,[ma_name],[subtype],[private_configuration_xml] FROM [ADSync].[dbo].[mms_management_agent] WHERE subtype = 'Windows Azure Active Directory (Microsoft)'" -ErrorAction Stop
	$UsersFromBothMAs = Invoke-Sqlcmd -ServerInstance $MSOLInstance -Query "SELECT [ma_id] ,[rdn] FROM [ADSync].[dbo].[mms_connectorspace] WHERE object_type = 'user'" -ErrorAction Stop
	$AzureUsers = $UsersFromBothMAs | ? { $_.ma_id -eq $AzureMA.ma_id }
	$SyncHistory = Invoke-Sqlcmd -ServerInstance $MSOLInstance -Query "SELECT [step_result] ,[end_date] ,[stage_no_change] ,[stage_add] ,[stage_update] ,[stage_rename] ,[stage_delete] ,[stage_deleteadd] ,[stage_failure] FROM [ADSync].[dbo].[mms_step_history]" -ErrorAction Stop | sort end_date -Descending
	#$ADUsers = $UsersFromBothMAs | ? {$_.ma_id -eq $ADMA.ma_id}
}
Catch
{
	Exit-Script "Could not fetch connector space / run history information information. Aborting script."
}

#GetDirSync interval
#AADSync uses Task Scheduler to get interval
Write-Host "Determining Sync Interval" -ForegroundColor Cyan
Try
{
	#if older version of DirSync is used
	$SyncTimeInterval = ((Get-ScheduledTaskInfo "Azure AD Sync Scheduler" -ErrorAction Stop).NextRunTime - (Get-ScheduledTaskInfo "Azure AD Sync Scheduler" -ErrorAction Stop).LastRunTime).TotalMinutes
}
Catch
{
	#use the new CMDLET if previous method failed
	$SyncTimeInterval = (Get-ADSyncScheduler).CurrentlyEffectiveSyncCycleInterval.TotalMinutes
}

Write-Host "Building HTML file" -ForegroundColor Cyan
#Build HTML tags
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
		
		$html += "<b>Synchronized Objects Info:</b><br/>"
		foreach ($ad in $ADMA)
		{
			$html += "Objects in AD " + $ad.ma_name + ": <font class='value'>" + ($UsersFromBothMAs | ? { $_.ma_id -eq $ad.ma_id }).count + "</font><br/>"
		}
		$html += "Objects in Azure Connector Space: <font class='value'>" + $AzureUsers.Count + "</font><br/>"
		$html += "Total objects: <font class='value'>" + $UsersFromBothMAs.Count + "</font><br/>"
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

	$html += "</body>" #close BODY tag
$html += "</html>" #close HTML tag

#Export File to selected destination.
$FilePath = $filePath.TrimEnd("\") #trim file path if backslash was provided during input
if (Test-Path $filePath)
{
	Try
	{
		$html | Out-File $filePath"\AADSyncInfo_$date.html" -ErrorAction Stop #generate HTML file
		Write-Host "Successfully generated HTML file"$filePath"\AADSyncInfo_$date.html" -ForegroundColor Green
	}
	Catch
	{
		Write-Warning "Could not generate HTML file in selected location. Reverting to default location $defaultFilePath."
		$html | Out-File $defaultFilePath"\AADSyncInfo_$date.html" -ErrorAction Stop #generate HTML file
		Write-Host "Successfully generated HTML file"$filePath"\AADSyncInfo_$date.html" -ForegroundColor Green
	}
}
else
{
	Write-Warning "Could not find path $filePath. Reverting to default location $defaultFilePath."
	Try
	{
		$html | Out-File $defaultFilePath"\AADSyncInfo_$date.html" -ErrorAction Stop #generate HTML file
		Write-Host "Successfully generated HTML file"$filePath"\AADSyncInfo_$date.html" -ForegroundColor Green
	}
	Catch
	{
		Exit-Script "Could not generate HTML file in default location. Aborting script."
	}
}

#Cleanup
Remove-Variable html
