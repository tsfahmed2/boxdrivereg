<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.150
	 Created on:   	4/1/2019 1:10 AM
	 Created by:   	tausifkhan
	 Organization: 	FICO
	 Filename:     	
	===========================================================================
	.DESCRIPTION
		
# Install Box Drive and the reg key for qualys and csia , also change box drive location by tausif
#

#>

Function Import-SMSTSENV
{
	try
	{
		$tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment
		Write-Output "$ScriptName - tsenv is $tsenv "
		$MDTIntegration = "YES"
		
		#$tsenv.GetVariables() | % { Write-Output "$ScriptName - $_ = $($tsenv.Value($_))" }
	}
	catch
	{
		Write-Output "$ScriptName - Unable to load Microsoft.SMS.TSEnvironment"
		Write-Output "$ScriptName - Running in standalonemode"
		$MDTIntegration = "NO"
	}
	Finally
	{
		if ($MDTIntegration -eq "YES")
		{
			$Logpath = $tsenv.Value("LogPath")
			$LogFile = $Logpath + "\" + "$ScriptName" + "$(get-date -format `"yyyyMMdd_hhmmsstt`").log"
			
		}
		Else
		{
			$Logpath = $env:TEMP
			$LogFile = $Logpath + "\" + "$ScriptName" + "$(get-date -format `"yyyyMMdd_hhmmsstt`").log"
		}
	}
}
Function Start-Logging
{
	start-transcript -path $LogFile -Force
}
Function Stop-Logging
{
	Stop-Transcript
}

# Set Vars
$SCRIPTDIR = split-path -parent $MyInvocation.MyCommand.Path
$SCRIPTNAME = split-path -leaf $MyInvocation.MyCommand.Path
$SOURCEROOT = "$SCRIPTDIR\Source"
$LANG = (Get-Culture).Name
$OSV = $Null
$ARCHITECTURE = $env:PROCESSOR_ARCHITECTURE

#Try to Import SMSTSEnv
. Import-SMSTSENV

#Start Transcript Logging
. Start-Logging

#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

#Output base info
Write-Output ""
Write-Output "$ScriptName - ScriptDir: $ScriptDir"
Write-Output "$ScriptName - SourceRoot: $SOURCEROOT"
Write-Output "$ScriptName - ScriptName: $ScriptName"
Write-Output "$ScriptName - ScriptVersion: 988.1"
Write-Output "$ScriptName - Log: $LogFile"
###############
function Get-UninstallRegistryKey
{
	
	[CmdletBinding()]
	param (
		[parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[string]$softwareName,
		[parameter(ValueFromRemainingArguments = $true)]
		[Object[]]$ignoredArguments
	)
	
	#Write-FunctionCallLogMessage -Invocation $MyInvocation -Parameters $PSBoundParameters
	
	if ($softwareName -eq $null -or $softwareName -eq '')
	{
		throw "$SoftwareName cannot be empty for Get-UninstallRegistryKey"
	}
	
	$ErrorActionPreference = 'Stop'
	$local_key = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
	$machine_key = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
	$machine_key6432 = 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
	
	Write-Output "Retrieving all uninstall registry keys"
	[array]$keys = Get-ChildItem -Path @($machine_key6432, $machine_key, $local_key) -ErrorAction SilentlyContinue
	Write-Debug "Registry uninstall keys on system: $($keys.Count)"
	
	#Write-Output "Error handling check: `'Get-ItemProperty`' fails if a registry key is encoded incorrectly."
	[int]$maxAttempts = $keys.Count
	for ([int]$attempt = 1; $attempt -le $maxAttempts; $attempt++)
	{
		[bool]$success = $false
		
		$keyPaths = $keys | Select-Object -ExpandProperty PSPath
		try
		{
			[array]$foundKey = Get-ItemProperty -Path $keyPaths -ErrorAction Stop | ? { $_.DisplayName -eq $softwareName }
			$success = $true
		}
		catch
		{
			Write-Debug "Found bad key."
			foreach ($key in $keys)
			{
				try
				{
					Get-ItemProperty $key.PsPath > $null
				}
				catch
				{
					$badKey = $key.PsPath
				}
			}
			Write-Output "Skipping bad key: $badKey"
			[array]$keys = $keys | ? { $badKey -NotContains $_.PsPath }
		}
		
		if ($success) { break; }
		
		if ($attempt -ge 10)
		{
			Write-Output "Found 10 or more bad registry keys. Run command again with `'--verbose --debug`' for more info."
			Write-Output "Each key searched should correspond to an installed program. It is very unlikely to have more than a few programs with incorrectly encoded keys, if any at all. This may be indicative of one or more corrupted registry branches."
		}
	}
	if ($foundKey -eq $null -or $foundkey.Count -eq 0)
	{
		Write-Output "No registry key found based on  '$softwareName'"
	}
	Write-Output "Found $($foundKey.Count) uninstall registry key(s) with SoftwareName:`'$SoftwareName`'";
	return $foundKey
}

Set-Alias Get-InstallRegistryKey Get-UninstallRegistryKey





function downloadbox
{
	$Store = ""
	Write-Output "***************Beginning function : downloadbox***********"
	$Path = "$env:windir\Temp\"
	
	$DownloadUrl = (((Invoke-WebRequest -Uri 'https://www.box.com/resources/downloads' -UseBasicParsing).Links | Where-Object { ($_.href -match "Box-x64.msi") }).href).trim()
	Write-Output "Download URL is $DownloadUrl"
	
	$filename = "Box-x64.msi"
	$completepath = $Path + $filename
	Write-Output "Installer save path is $completepath"
	# Download the latest installer from box
	
	Write-Output "Downloading $filename."
	#Invoke-WebRequest $DownloadUrl -OutFile $completepath
	$WebClient = New-Object System.Net.WebClient
	$WebClient.DownloadFile($DownloadUrl, $completepath)
	
	
	Start-Sleep -s 35
}


#Get MSI file version for downloaded file

function get-msifileinformation
{
	param (
		[parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[System.IO.FileInfo]$Path,
		[parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[ValidateSet("ProductCode", "ProductVersion", "ProductName", "Manufacturer", "ProductLanguage", "FullVersion")]
		[string]$Property
	)
	<#Process
	{#>
	try
	{
		Write-Output "***************Beginning function : get-msifileinformation***********"
		# Read property from MSI database
		$WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
		$MSIDatabase = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($Path.FullName, 0))
		$Query = "SELECT Value FROM Property WHERE Property = '$($Property)'"
		$View = $MSIDatabase.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $MSIDatabase, ($Query))
		$View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
		$Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
		$Value = $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
		
		# Commit database and close view
		$MSIDatabase.GetType().InvokeMember("Commit", "InvokeMethod", $null, $MSIDatabase, $null)
		$View.GetType().InvokeMember("Close", "InvokeMethod", $null, $View, $null)
		$MSIDatabase = $null
		$View = $null
		
		# Return the value
		return $Value
	}
	catch
	{
		Write-Warning -Message $_.Exception.Message; break
	}
	<#}
	End
	{
		# Run garbage collection and release ComObject
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-Null
		[System.GC]::Collect()
	}#>
}



function Start-ProcessInteractive
{
	param (
		$Path = "${env:ProgramFiles}\Box\Box\Box.exe",
		#$Arguments = 'www.powertheshell.com',

		[Parameter(Mandatory = $true)]
		$Computername,
		[Parameter(Mandatory = $true)]
		$Username
	)
	
	
	
	$xml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo />
  <Triggers />
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>false</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings />
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT72H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>"$Path"</Command>
      <Arguments>$Arguments</Arguments>
    </Exec>
  </Actions>
  <Principals>
    <Principal id="Author">
      <UserId>$Username</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
</Task>
"@
	
	$jobname = 'remotejob{0}' -f (Get-Random)
	$filename = [Guid]::NewGuid().ToString('d')
	$filepath = "$env:temp\$filename"
	
	$xml | Set-Content -Path $filepath -Encoding Unicode
	
	try
	{
		$ErrorActionPreference = 'Stop'
		schtasks.exe /CREATE /TN $jobname /XML $filepath /S $ComputerName  2>&1
		schtasks.exe /RUN /TN $jobname /S $ComputerName  2>&1
		schtasks.exe /DELETE /TN $jobname /s $ComputerName /F  2>&1
	}
	catch
	{
		Write-Warning ("While accessing \\$ComputerName : " + $_.Exception.Message)
	}
	Remove-Item -Path $filepath
}


function windows10rebootnotification()
{
	[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
	[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | out-null
	[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | out-null
	$TimeStart = Get-Date
	$TimeEnd = $timeStart.addminutes(360)
	Do
	{
		$TimeNow = Get-Date
		if ($TimeNow -ge $TimeEnd)
		{
			
			Unregister-Event -SourceIdentifier click_event -ErrorAction SilentlyContinue
			Remove-Event click_event -ErrorAction SilentlyContinue
			[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
			[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
			Exit
		}
		else
		{
			$Balloon = new-object System.Windows.Forms.NotifyIcon
			$Balloon.Icon = [System.Drawing.SystemIcons]::Information
			$Balloon.BalloonTipText = "Box Sync has been uninstalled. A reboot is required to complete the file cleanup. All Box Sync content will be saved to Box."
			$Balloon.BalloonTipTitle = "Reboot Required"
			$Balloon.BalloonTipIcon = "Warning"
			$Balloon.Visible = $true;
			$Balloon.ShowBalloonTip(20000);
			$Balloon_MouseOver = [System.Windows.Forms.MouseEventHandler]{ $Balloon.ShowBalloonTip(20000) }
			$Balloon.add_MouseClick($Balloon_MouseOver)
			Unregister-Event -SourceIdentifier click_event -ErrorAction SilentlyContinue
			Register-ObjectEvent $Balloon BalloonTipClicked -sourceIdentifier click_event -Action {
				Add-Type -AssemblyName Microsoft.VisualBasic
				
				If ([Microsoft.VisualBasic.Interaction]::MsgBox('Would you like to reboot your machine now?', 'YesNo,MsgBoxSetForeground,Question', 'System Maintenance') -eq "NO")
				{ }
				else
				{
					shutdown -r -f
				}
				
			} | Out-Null
			
			Wait-Event -timeout 7200 -sourceIdentifier click_event > $null
			Unregister-Event -SourceIdentifier click_event -ErrorAction SilentlyContinue
			$Balloon.Dispose()
		}
		
	}
	Until ($TimeNow -ge $TimeEnd)
}



function killbox()
{
	Write-Output "***************Beginning function : killbox***********"
	$KillProcess = @("Box", "BoxUI", "Box.Desktop.UpdateService")
	if ($KillProcess)
	{
		foreach ($process in $KillProcess)
		{
			Write-Output "Attempting to stop process $process..."
			if ($Computername -ne 'localhost')
			{
				$WmiProcess = Get-WmiObject -Class Win32_Process -Filter "name='$process`.exe'"
				if ($WmiProcess)
				{
					$WmiProcess.Terminate() | Out-Null
				}
			}
			else
			{
				Stop-Process -Name $process -Force -ErrorAction 'SilentlyContinue'
			}
		}
	}
	sleep -Seconds 9
}

function Test-RegistryValue
{
	
	param (
		
		[parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$Path,
		[parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$Value
	)
	
	try
	{
		Write-Output "***************Beginning function : Test-Registryvalue***********"
		Get-ItemProperty -Path $Path | Select-Object -ExpandProperty $Value -ErrorAction SilentlyContinue | Out-Null
		return $true
	}
	
	catch
	{
		
		return $false
		
	}
	
}

$boxsyncinstalled = Get-UninstallRegistryKey -softwareName "Box Sync"
$boxsyninstalledversion = $boxsyncinstalled.Version

$boxdriveinstalled = get-UninstallRegistryKey -SoftwareName "Box" #| Select-Object -First 1
$installedversion = $boxdriveinstalled.DisplayVersion
$boxsyncstartmenufolder = "$Env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu\Programs\Box Sync"

$boxexepath = "$env:ProgramFiles\Box\Box\box.exe"
$result = Get-InstalledApps | Where-Object { $_.DisplayName -like "Box" }
$boxmsi = "$env:windir\Temp\Box-x64.msi"
$checkforregkey = Test-RegistryValue -Path HKLM:\SOFTWARE\Box\Box -Value CustomBoxLocation
<#
if ($boxsyncinstalledversion) {
      $boxsyncinstalledstring = $boxsyncinstalled.UninstallString 
        $boxsyncinstalledstring = $boxsyncinstalledstring.Trim()
        $boxsyncinstalledstring = $boxsyncinstalledstring -Replace "msiexec.exe","" -Replace "/I","" -Replace "/X",""
        $boxsyncinstalledstring = $boxsyncinstalledstring.Trim()
        Write-Output $boxsyncinstalledstring
        Write-Output 'Start MSIExec'
        start-process "msiexec.exe" -arg "/X $boxsyncinstalledstring /qn /l*v $env:TEMP\$boxsyncinstalledstring.log" -Wait
        Write-Output 'After MSIExec'
}
#>

if ((Test-Path $boxexepath) -or ($installedversion))
{
	
	Write-Output "***************Box is installed***********"
	#remove Box using the function
	#uninstall-boxdrive
	#sleep -Seconds 20
	# Install Box Drive
	Write-Output"Downloading box msi to check version"
	downloadbox
	
	$installedboxdriveversion = $installedversion
	Write-Output "Currently Installed Box Version is $installedboxdriveversion"
	$latestboxdriveversion = get-msifileinformation -Property ProductVersion $boxmsi
	Write-Output " Latest box drive version is $latestboxdriveversion"
	if (([System.Version]"$latestboxdriveversion" -eq [System.Version]"$installedboxdriveversion") -and ($checkforregkey -eq $false))
	{
		Write-Output "Latest box drive installed, importing reg key"
		$a = @(
			"QualysAgent.exe",
			"csia.exe"
		)
		
		$b = "C:\"
		$c = '3'
		New-ItemProperty -Path "HKLM:\SOFTWARE\Box\Box" -PropertyType MultiString -Name BannedProcessNames -Value $a
		New-ItemProperty -Path "HKLM:\SOFTWARE\Box\Box" -PropertyType String -Name CustomBoxLocation -Value $b
		New-ItemProperty -Path "HKLM:\SOFTWARE\Box\Box" -PropertyType DWORD -Name SyncUninstallMode -Value $c
		killbox
		#Start-Process $boxexepath -ErrorAction SilentlyContinue
	}
	else
	{
		
		<#
		$result = Get-InstalledApps | where { $_.DisplayName -like "Box" }
		Write-Output "Installed apps like box - $result"
		Write-Output "Begin removing older versions of box drive"
		#remove current box installs
		ForEach ($u in $result)
		{
			$UnInstall = $u.UninstallString
			$UnInstall = $UnInstall.Trim()
			$UnInstall = $UnInstall -Replace "msiexec.exe", "" -Replace "/I", "" -Replace "/X", ""
			$UnInstall = $UnInstall.Trim()
			Write-Output $UnInstall
			Write-Output 'Start MSIExec'
			Start-Process "msiexec.exe" -arg "/X $UnInstall /qn /l*v $env:TEMP\$UnInstall.log" -Wait
			Write-Output 'After MSIExec'
			sleep -Seconds 8
		}
		#>
		if (Test-Path $boxmsi)
		{
			Write-Output "Installing newer Box drive Version"
			$arguments = @(
				"/i"
				"`"$boxmsi`""
				"/quiet"
				"/norestart"
			)
			
			Write-Output "Installing $boxmsi....."
			$process = Start-Process -FilePath msiexec.exe -ArgumentList $arguments -Wait -PassThru
			if ($process.ExitCode -eq 0)
			{
				Write-Output "$boxmsi has been successfully installed"
			}
			else
			{
				Write-Output "installer exit code  $($process.ExitCode) for file  $($boxmsi)"
			}
			Write-Output "Import reg values to set bannedprocesses"
			$a = @(
				"QualysAgent.exe",
				"csia.exe"
			)
			
			$b = "C:\"
			$c = '3'
			New-ItemProperty -Path "HKLM:\SOFTWARE\Box\Box" -PropertyType MultiString -Name BannedProcessNames -Value $a
			New-ItemProperty -Path "HKLM:\SOFTWARE\Box\Box" -PropertyType String -Name CustomBoxLocation -Value $b
			New-ItemProperty -Path "HKLM:\SOFTWARE\Box\Box" -PropertyType DWORD -Name SyncUninstallMode -Value $c
			killbox
		}
		else
		{
			Write-Output "Box drive MSI not found"
			exit 99
		}
		
		
	}
	# Cleanup
	Write-Output "removing $boxmsi."
	Remove-Item $boxmsi -Force
	#remove box sync startmenu item if it exists
	Remove-Item $boxsyncstartmenufolder -Force -ea SilentlyContinue
	
	$explorerprocesses = @(Get-WmiObject -Query "Select * FROM Win32_Process WHERE Name='explorer.exe'" -ErrorAction SilentlyContinue)
	if ($explorerprocesses.Count -eq 0)
	{
		"No explorer process found / Nobody interactively logged on"
	}
	else
	{
		foreach ($i in $explorerprocesses)
		{
			$Username = $i.GetOwner().User
			#$Domain = $i.GetOwner().Domain
			#$Domain + "\" + $Username + " logged on since: " + ($i.ConvertToDateTime($i.CreationDate))
		}
	}
	Start-ProcessInteractive -Computername $env:COMPUTERNAME -Username $Username
	windows10rebootnotification
}

elseif ($installedversion -eq $Null)
{
	Write-Output "Box drive is not installed, Installing now"
	downloadbox
	
	if (Test-Path $boxmsi)
	{
		Write-Output "Installing Box drive Version"
		Write-Output "Installing newer Box drive Version"
		$arguments = @(
			"/i"
			"`"$boxmsi`""
			"/quiet"
			"/norestart"
		)
		
		Write-Output "Installing $boxmsi....."
		$process = Start-Process -FilePath msiexec.exe -ArgumentList $arguments -Wait -PassThru
		if ($process.ExitCode -eq 0)
		{
			Write-Output "$boxmsi has been successfully installed"
		}
		else
		{
			Write-Output "installer exit code  $($process.ExitCode) for file  $($boxmsi)"
		}
		$a = @(
			"QualysAgent.exe",
			"csia.exe"
		)
		
		$b = "C:\"
		$c = '3'
		New-ItemProperty -Path "HKLM:\SOFTWARE\Box\Box" -PropertyType MultiString -Name BannedProcessNames -Value $a
		New-ItemProperty -Path "HKLM:\SOFTWARE\Box\Box" -PropertyType String -Name CustomBoxLocation -Value $b
		New-ItemProperty -Path "HKLM:\SOFTWARE\Box\Box" -PropertyType DWORD -Name SyncUninstallMode -Value $c
	}
	else
	{
		Write-Output "Box drive MSI not found"
		exit 99
	}
	
	Write-Output " Killing Box processes"
	killbox
	Write-Output "Removing box MSI"
	Remove-Item $boxmsi -Force
	Remove-Item $boxsyncstartmenufolder -Force -ea SilentlyContinue
	
	$explorerprocesses = @(Get-WmiObject -Query "Select * FROM Win32_Process WHERE Name='explorer.exe'" -ErrorAction SilentlyContinue)
	if ($explorerprocesses.Count -eq 0)
	{
		"No explorer process found / Nobody interactively logged on"
	}
	else
	{
		foreach ($i in $explorerprocesses)
		{
			$Username = $i.GetOwner().User
			#$Domain = $i.GetOwner().Domain
			#$Domain + "\" + $Username + " logged on since: " + ($i.ConvertToDateTime($i.CreationDate))
		}
	}
	Start-ProcessInteractive -Computername $env:COMPUTERNAME -Username $Username
	windows10rebootnotification
}
. Stop-Logging

# SIG # Begin signature block
# MIIfXAYJKoZIhvcNAQcCoIIfTTCCH0kCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUABQXiq5tS1qF30crpnLhIIye
# uSCgghn1MIIEFDCCAvygAwIBAgILBAAAAAABL07hUtcwDQYJKoZIhvcNAQEFBQAw
# VzELMAkGA1UEBhMCQkUxGTAXBgNVBAoTEEdsb2JhbFNpZ24gbnYtc2ExEDAOBgNV
# BAsTB1Jvb3QgQ0ExGzAZBgNVBAMTEkdsb2JhbFNpZ24gUm9vdCBDQTAeFw0xMTA0
# MTMxMDAwMDBaFw0yODAxMjgxMjAwMDBaMFIxCzAJBgNVBAYTAkJFMRkwFwYDVQQK
# ExBHbG9iYWxTaWduIG52LXNhMSgwJgYDVQQDEx9HbG9iYWxTaWduIFRpbWVzdGFt
# cGluZyBDQSAtIEcyMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAlO9l
# +LVXn6BTDTQG6wkft0cYasvwW+T/J6U00feJGr+esc0SQW5m1IGghYtkWkYvmaCN
# d7HivFzdItdqZ9C76Mp03otPDbBS5ZBb60cO8eefnAuQZT4XljBFcm05oRc2yrmg
# jBtPCBn2gTGtYRakYua0QJ7D/PuV9vu1LpWBmODvxevYAll4d/eq41JrUJEpxfz3
# zZNl0mBhIvIG+zLdFlH6Dv2KMPAXCae78wSuq5DnbN96qfTvxGInX2+ZbTh0qhGL
# 2t/HFEzphbLswn1KJo/nVrqm4M+SU4B09APsaLJgvIQgAIMboe60dAXBKY5i0Eex
# +vBTzBj5Ljv5cH60JQIDAQABo4HlMIHiMA4GA1UdDwEB/wQEAwIBBjASBgNVHRMB
# Af8ECDAGAQH/AgEAMB0GA1UdDgQWBBRG2D7/3OO+/4Pm9IWbsN1q1hSpwTBHBgNV
# HSAEQDA+MDwGBFUdIAAwNDAyBggrBgEFBQcCARYmaHR0cHM6Ly93d3cuZ2xvYmFs
# c2lnbi5jb20vcmVwb3NpdG9yeS8wMwYDVR0fBCwwKjAooCagJIYiaHR0cDovL2Ny
# bC5nbG9iYWxzaWduLm5ldC9yb290LmNybDAfBgNVHSMEGDAWgBRge2YaRQ2XyolQ
# L30EzTSo//z9SzANBgkqhkiG9w0BAQUFAAOCAQEATl5WkB5GtNlJMfO7FzkoG8IW
# 3f1B3AkFBJtvsqKa1pkuQJkAVbXqP6UgdtOGNNQXzFU6x4Lu76i6vNgGnxVQ380W
# e1I6AtcZGv2v8Hhc4EvFGN86JB7arLipWAQCBzDbsBJe/jG+8ARI9PBw+DpeVoPP
# PfsNvPTF7ZedudTbpSeE4zibi6c1hkQgpDttpGoLoYP9KOva7yj2zIhd+wo7AKvg
# IeviLzVsD440RZfroveZMzV+y5qKu0VN5z+fwtmK+mWybsd+Zf/okuEsMaL3sCc2
# SI8mbzvuTXYfecPlf5Y1vC0OzAGwjn//UYCAp5LUs0RGZIyHTxZjBzFLY7Df8zCC
# BJ8wggOHoAMCAQICEhEh1pmnZJc+8fhCfukZzFNBFDANBgkqhkiG9w0BAQUFADBS
# MQswCQYDVQQGEwJCRTEZMBcGA1UEChMQR2xvYmFsU2lnbiBudi1zYTEoMCYGA1UE
# AxMfR2xvYmFsU2lnbiBUaW1lc3RhbXBpbmcgQ0EgLSBHMjAeFw0xNjA1MjQwMDAw
# MDBaFw0yNzA2MjQwMDAwMDBaMGAxCzAJBgNVBAYTAlNHMR8wHQYDVQQKExZHTU8g
# R2xvYmFsU2lnbiBQdGUgTHRkMTAwLgYDVQQDEydHbG9iYWxTaWduIFRTQSBmb3Ig
# TVMgQXV0aGVudGljb2RlIC0gRzIwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEK
# AoIBAQCwF66i07YEMFYeWA+x7VWk1lTL2PZzOuxdXqsl/Tal+oTDYUDFRrVZUjtC
# oi5fE2IQqVvmc9aSJbF9I+MGs4c6DkPw1wCJU6IRMVIobl1AcjzyCXenSZKX1GyQ
# oHan/bjcs53yB2AsT1iYAGvTFVTg+t3/gCxfGKaY/9Sr7KFFWbIub2Jd4NkZrItX
# nKgmK9kXpRDSRwgacCwzi39ogCq1oV1r3Y0CAikDqnw3u7spTj1Tk7Om+o/SWJMV
# TLktq4CjoyX7r/cIZLB6RA9cENdfYTeqTmvT0lMlnYJz+iz5crCpGTkqUPqp0Dw6
# yuhb7/VfUfT5CtmXNd5qheYjBEKvAgMBAAGjggFfMIIBWzAOBgNVHQ8BAf8EBAMC
# B4AwTAYDVR0gBEUwQzBBBgkrBgEEAaAyAR4wNDAyBggrBgEFBQcCARYmaHR0cHM6
# Ly93d3cuZ2xvYmFsc2lnbi5jb20vcmVwb3NpdG9yeS8wCQYDVR0TBAIwADAWBgNV
# HSUBAf8EDDAKBggrBgEFBQcDCDBCBgNVHR8EOzA5MDegNaAzhjFodHRwOi8vY3Js
# Lmdsb2JhbHNpZ24uY29tL2dzL2dzdGltZXN0YW1waW5nZzIuY3JsMFQGCCsGAQUF
# BwEBBEgwRjBEBggrBgEFBQcwAoY4aHR0cDovL3NlY3VyZS5nbG9iYWxzaWduLmNv
# bS9jYWNlcnQvZ3N0aW1lc3RhbXBpbmdnMi5jcnQwHQYDVR0OBBYEFNSihEo4Whh/
# uk8wUL2d1XqH1gn3MB8GA1UdIwQYMBaAFEbYPv/c477/g+b0hZuw3WrWFKnBMA0G
# CSqGSIb3DQEBBQUAA4IBAQCPqRqRbQSmNyAOg5beI9Nrbh9u3WQ9aCEitfhHNmmO
# 4aVFxySiIrcpCcxUWq7GvM1jjrM9UEjltMyuzZKNniiLE0oRqr2j79OyNvy0oXK/
# bZdjeYxEvHAvfvO83YJTqxr26/ocl7y2N5ykHDC8q7wtRzbfkiAD6HHGWPZ1BZo0
# 8AtZWoJENKqA5C+E9kddlsm2ysqdt6a65FDT1De4uiAO0NOSKlvEWbuhbds8zkSd
# wTgqreONvc0JdxoQvmcKAjZkiLmzGybu555gxEaovGEzbM9OuZy5avCfN/61PU+a
# 003/3iCOTpem/Z8JvE3KGHbJsE2FUPKA0h0G9VgEB7EYMIIFdjCCBF6gAwIBAgIQ
# MJb30qQzYi3liCtAtH5uTTANBgkqhkiG9w0BAQsFADB9MQswCQYDVQQGEwJHQjEb
# MBkGA1UECBMSR3JlYXRlciBNYW5jaGVzdGVyMRAwDgYDVQQHEwdTYWxmb3JkMRow
# GAYDVQQKExFDT01PRE8gQ0EgTGltaXRlZDEjMCEGA1UEAxMaQ09NT0RPIFJTQSBD
# b2RlIFNpZ25pbmcgQ0EwHhcNMTcwNDIwMDAwMDAwWhcNMjAwNDE5MjM1OTU5WjCB
# 3zELMAkGA1UEBhMCVVMxDjAMBgNVBBEMBTU1MTEzMRIwEAYDVQQIDAlNaW5uZXNv
# dGExEjAQBgNVBAcMCVJvc2V2aWxsZTEcMBoGA1UECQwTMjY2NSBMb25nIExha2Ug
# Um9hZDEhMB8GA1UECQwYUm9zZWRhbGUgQ29ycG9yYXRlIFBsYXphMR8wHQYDVQQK
# DBZGYWlyIElzYWFjIENvcnBvcmF0aW9uMRUwEwYDVQQLDAxJVCBDb3Jwb3JhdGUx
# HzAdBgNVBAMMFkZhaXIgSXNhYWMgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEB
# AQUAA4IBDwAwggEKAoIBAQDYpUJq0w7LcYQT/Q8KUjGybaOwDX55XFKIqQH5SDju
# fVp2YD+tjJp5UWDOSNrEYLqASUjQ73dtJgxFMvUKdjW2U8Ug5BBCxc//EPfCYcgP
# y4oU3m9oL4TK5qRI089Np5+hjUwG6QzBQkNIRwBAdljtxaKFoyC0OfIQdFYzpGbx
# xtN7zVgQTlX/g+ngUMZG9i0yepQTbyJA4KHWpTI04RSEAedlWRQHLVtYU3f4qBu/
# ZOV8NXXRedT8lzG0wN+ZR48B8nlUJAjVCfuVoCI354Gsc7EXf8i1FRSQ7HC7PXNa
# hmpFEFFP3+nJoECV4qTM0BbdVsnJVWCs6wdDmxO54ceZAgMBAAGjggGNMIIBiTAf
# BgNVHSMEGDAWgBQpkWD/ik366/mmarjP+eZLvUnOEjAdBgNVHQ4EFgQUsUVitpiw
# 221RlkBJXW1SrvInAPcwDgYDVR0PAQH/BAQDAgeAMAwGA1UdEwEB/wQCMAAwEwYD
# VR0lBAwwCgYIKwYBBQUHAwMwEQYJYIZIAYb4QgEBBAQDAgQQMEYGA1UdIAQ/MD0w
# OwYMKwYBBAGyMQECAQMCMCswKQYIKwYBBQUHAgEWHWh0dHBzOi8vc2VjdXJlLmNv
# bW9kby5uZXQvQ1BTMEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwuY29tb2Rv
# Y2EuY29tL0NPTU9ET1JTQUNvZGVTaWduaW5nQ0EuY3JsMHQGCCsGAQUFBwEBBGgw
# ZjA+BggrBgEFBQcwAoYyaHR0cDovL2NydC5jb21vZG9jYS5jb20vQ09NT0RPUlNB
# Q29kZVNpZ25pbmdDQS5jcnQwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmNvbW9k
# b2NhLmNvbTANBgkqhkiG9w0BAQsFAAOCAQEAmhwEBbNGK0hzbNBweOz22CXwweUB
# 2/lfXOVJ+UwB8V4EhhtoX8PXO8uotZgRugjuaWy8JKxnX47eVkkGrirxwIeDA19A
# JCjG3VciYocDvxAJn3iwsSNYKYLFjClyc1Su7xFvuaAtbts+Pptjb6j8z4MmGdgJ
# Ot6w5jK27pgI1aQ9GPvf+4WUQjVUlHkho7rAoEJcEgTtS6+o4AqYpUS3VhFIDTF/
# uNGUMxou2YqjomFTrnBNyt93VooWPRdd3xURSsda/dczlteu502RumRK71gSU3VV
# B2jWCHgE1aqoRV/XqfWb6HEvIrFqKzIcQS5g12fG1PDyP6eW47yq7qht5DCCBdgw
# ggPAoAMCAQICEEyq+crbY2/gH/dO2FsDhp0wDQYJKoZIhvcNAQEMBQAwgYUxCzAJ
# BgNVBAYTAkdCMRswGQYDVQQIExJHcmVhdGVyIE1hbmNoZXN0ZXIxEDAOBgNVBAcT
# B1NhbGZvcmQxGjAYBgNVBAoTEUNPTU9ETyBDQSBMaW1pdGVkMSswKQYDVQQDEyJD
# T01PRE8gUlNBIENlcnRpZmljYXRpb24gQXV0aG9yaXR5MB4XDTEwMDExOTAwMDAw
# MFoXDTM4MDExODIzNTk1OVowgYUxCzAJBgNVBAYTAkdCMRswGQYDVQQIExJHcmVh
# dGVyIE1hbmNoZXN0ZXIxEDAOBgNVBAcTB1NhbGZvcmQxGjAYBgNVBAoTEUNPTU9E
# TyBDQSBMaW1pdGVkMSswKQYDVQQDEyJDT01PRE8gUlNBIENlcnRpZmljYXRpb24g
# QXV0aG9yaXR5MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAkehUktIK
# VrGsDSTdxc9EZ3SZKzejfSNwAHG8U9/E+ioSj0t/EFa9n3Byt2F/yUsPF6c947AE
# Ye7/EZfH9IY+Cvo+XPmT5jR62RRr55yzhaCCenavcZDX7P0N+pxs+t+wgvQUfvm+
# xKYvT3+Zf7X8Z0NyvQwA1onrayzT7Y+YHBSrfuXjbvzYqOSSJNpDa2K4Vf3qwbxs
# tovzDo2a5JtsaZn4eEgwRdWt4Q08RWD8MpZRJ7xnw8outmvqRsfHIKCxH2XeSAi6
# pE6p8oNGN4Tr6MyBSENnTnIqm1y9TBsoilwie7SrmNnu4FGDwwlGTm0+mfqVF9p8
# M1dBPI1R7Qu2XK8sYxrfV8g/vOldxJuvRZnio1oktLqpVj3Pb6r/SVi+8Kj/9Lit
# 6Tf7urj0Czr56ENCHonYhMsT8dm74YlguIwoVqwUHZwK53Hrzw7dPamWoUi9PPev
# tQ0iTMARgexWO/bTouJbt7IEIlKVgJNp6I5MZfGRAy1wdALqi2cVKWlSArvX31Bq
# VUa/oKMoYX9w0MOiqiwhqkfOKJwGRXa/ghgntNWutMtQ5mv0TIZxMOmm3xaG4Nj/
# QN370EKIf6MzOi5cHkERgWPOGHFrK+ymircxXDpqR+DDeVnWIBqv8mqYqnK8V0rS
# S527EPywTEHl7R09XiidnMy/s1Hap0flhFMCAwEAAaNCMEAwHQYDVR0OBBYEFLuv
# fgI9+qbxPISOre44mOzZMjLUMA4GA1UdDwEB/wQEAwIBBjAPBgNVHRMBAf8EBTAD
# AQH/MA0GCSqGSIb3DQEBDAUAA4ICAQAK8dVGhLeuUbtssk1BFACTTJzL5cBUz6Al
# jgL5/bCiDfUgmDwTLaxWorDWfhGS6S66ni6acrG9GURsYTWimrQWEmlajOHXPqQa
# 6C8D9K5hHRAbKqSLesX+BabhwNbI/p6ujyu6PZn42HMJWEZuppz01yfTldo3g3Ic
# 03PgokeZAzhd1Ul5ACkcx+ybIBwHJGlXeLI5/DqEoLWcfI2/LpNiJ7c52hcYrr08
# CWj/hJs81dYLA+NXnhT30etPyL2HI7e2SUN5hVy665ILocboaKhMFrEamQroUyyS
# u6EJGHUMZah7yyO3GsIohcMb/9ArYu+kewmRmGeMFAHNaAZqYyF1A4CIim6BxoXy
# qaQt5/SlJBBHg8rN9I15WLEGm+caKtmdAdeUfe0DSsrw2+ipAT71VpnJHo5JPbvl
# CbngT0mSPRaCQMzMWcbmOu0SLmk8bJWx/aode3+Gvh4OMkb7+xOPdX9Mi0tGY/4A
# NEBwwcO5od2mcOIEs0G86YCR6mSceuEiA6mcbm8OZU9sh4de826g+XWlm0DoU7In
# nUq5wHchjf+H8t68jO8X37dJC9HybjALGg5Odu0R/PXpVrJ9v8dtCpOMpdDAth2+
# Ok6UotdubAvCinz6IPPE5OXNDajLkZKxfIXstRRpZg6C583OyC2mUX8hwTVThQZK
# XZ+tuxtfdDCCBeAwggPIoAMCAQICEC58h8wOk0pS/pT9HLfNNK8wDQYJKoZIhvcN
# AQEMBQAwgYUxCzAJBgNVBAYTAkdCMRswGQYDVQQIExJHcmVhdGVyIE1hbmNoZXN0
# ZXIxEDAOBgNVBAcTB1NhbGZvcmQxGjAYBgNVBAoTEUNPTU9ETyBDQSBMaW1pdGVk
# MSswKQYDVQQDEyJDT01PRE8gUlNBIENlcnRpZmljYXRpb24gQXV0aG9yaXR5MB4X
# DTEzMDUwOTAwMDAwMFoXDTI4MDUwODIzNTk1OVowfTELMAkGA1UEBhMCR0IxGzAZ
# BgNVBAgTEkdyZWF0ZXIgTWFuY2hlc3RlcjEQMA4GA1UEBxMHU2FsZm9yZDEaMBgG
# A1UEChMRQ09NT0RPIENBIExpbWl0ZWQxIzAhBgNVBAMTGkNPTU9ETyBSU0EgQ29k
# ZSBTaWduaW5nIENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAppiQ
# Y3eRNH+K0d3pZzER68we/TEds7liVz+TvFvjnx4kMhEna7xRkafPnp4ls1+BqBgP
# HR4gMA77YXuGCbPj/aJonRwsnb9y4+R1oOU1I47Jiu4aDGTH2EKhe7VSA0s6sI4j
# S0tj4CKUN3vVeZAKFBhRLOb+wRLwHD9hYQqMotz2wzCqzSgYdUjBeVoIzbuMVYz3
# 1HaQOjNGUHOYXPSFSmsPgN1e1r39qS/AJfX5eNeNXxDCRFU8kDwxRstwrgepCuOv
# wQFvkBoj4l8428YIXUezg0HwLgA3FLkSqnmSUs2HD3vYYimkfjC9G7WMcrRI8uPo
# IfleTGJ5iwIGn3/VCwIDAQABo4IBUTCCAU0wHwYDVR0jBBgwFoAUu69+Aj36pvE8
# hI6t7jiY7NkyMtQwHQYDVR0OBBYEFCmRYP+KTfrr+aZquM/55ku9Sc4SMA4GA1Ud
# DwEB/wQEAwIBhjASBgNVHRMBAf8ECDAGAQH/AgEAMBMGA1UdJQQMMAoGCCsGAQUF
# BwMDMBEGA1UdIAQKMAgwBgYEVR0gADBMBgNVHR8ERTBDMEGgP6A9hjtodHRwOi8v
# Y3JsLmNvbW9kb2NhLmNvbS9DT01PRE9SU0FDZXJ0aWZpY2F0aW9uQXV0aG9yaXR5
# LmNybDBxBggrBgEFBQcBAQRlMGMwOwYIKwYBBQUHMAKGL2h0dHA6Ly9jcnQuY29t
# b2RvY2EuY29tL0NPTU9ET1JTQUFkZFRydXN0Q0EuY3J0MCQGCCsGAQUFBzABhhho
# dHRwOi8vb2NzcC5jb21vZG9jYS5jb20wDQYJKoZIhvcNAQEMBQADggIBAAI/AjnD
# 7vjKO4neDG1NsfFOkk+vwjgsBMzFYxGrCWOvq6LXAj/MbxnDPdYaCJT/JdipiKcr
# EBrgm7EHIhpRHDrU4ekJv+YkdK8eexYxbiPvVFEtUgLidQgFTPG3UeFRAMaH9mzu
# EER2V2rx31hrIapJ1Hw3Tr3/tnVUQBg2V2cRzU8C5P7z2vx1F9vst/dlCSNJH0NX
# g+p+IHdhyE3yu2VNqPeFRQevemknZZApQIvfezpROYyoH3B5rW1CIKLPDGwDjEzN
# cweU51qOOgS6oqF8H8tjOhWn1BUbp1JHMqn0v2RH0aofU04yMHPCb7d4gp1c/0a7
# ayIdiAv4G6o0pvyM9d1/ZYyMMVcx0DbsR6HPy4uo7xwYWMUGd8pLm1GvTAhKeo/i
# o1Lijo7MJuSy2OU4wqjtxoGcNWupWGFKCpe0S0K2VZ2+medwbVn4bSoMfxlgXwya
# iGwwrFIJkBYb/yud29AgyonqKH4yjhnfe0gzHtdl+K7J+IMUk3Z9ZNCOzr41ff9y
# MU2fnr0ebC+ojwwGUPuMJ7N2yfTm18M04oyHIYZh/r9VdOEhdwMKaGy75Mmp5s9Z
# Jet87EUOeWZo6CLNuO+YhU2WETwJitB/vCgoE/tqylSNklzNwmWYBp7OSFvUtTeT
# RkF8B93P+kPvumdh/31J4LswfVyA4+YWOUunMYIE0TCCBM0CAQEwgZEwfTELMAkG
# A1UEBhMCR0IxGzAZBgNVBAgTEkdyZWF0ZXIgTWFuY2hlc3RlcjEQMA4GA1UEBxMH
# U2FsZm9yZDEaMBgGA1UEChMRQ09NT0RPIENBIExpbWl0ZWQxIzAhBgNVBAMTGkNP
# TU9ETyBSU0EgQ29kZSBTaWduaW5nIENBAhAwlvfSpDNiLeWIK0C0fm5NMAkGBSsO
# AwIaBQCgcDAQBgorBgEEAYI3AgEMMQIwADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGC
# NwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQx
# FgQUrYl6QSNhKWqxXqjP067+m2Wa73UwDQYJKoZIhvcNAQEBBQAEggEAUzKTl8Bp
# eJoaW/TJlfyN0lr08juAqGaYWuPRmttO8ADh7T1EwWSyjRfILtJpXoWeTSVQQwWJ
# +IoOUtD96R/QbGwUbg95xa0jfBSmomuPr6tXdWX+UOvYAVMoTGbl2LdRSTIL6TTR
# LHh8MX4AVTFn72AZ7bAcZkeCjVzIiVanR6G5sEBZhN6ScPjxElw7g2rcRlPbO0K2
# VciXYVP9VSvdJchCfL92euK1s9MCcjmwnRzqDvy5wOzMtWPF7Ix95AAv+Eh0RG6v
# HkgDR6Y9w2U92kOCP16vy1tP9/mxdU9G00/ROiGPej0LzWdm8rHYAwyYo6ZwjYm5
# /prNlQj4GQobd6GCAqIwggKeBgkqhkiG9w0BCQYxggKPMIICiwIBATBoMFIxCzAJ
# BgNVBAYTAkJFMRkwFwYDVQQKExBHbG9iYWxTaWduIG52LXNhMSgwJgYDVQQDEx9H
# bG9iYWxTaWduIFRpbWVzdGFtcGluZyBDQSAtIEcyAhIRIdaZp2SXPvH4Qn7pGcxT
# QRQwCQYFKw4DAhoFAKCB/TAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0xOTA0MDEwNjU1NDRaMCMGCSqGSIb3DQEJBDEWBBQnAmukVn3J
# f4lPA3AZCVrJsx4eMTCBnQYLKoZIhvcNAQkQAgwxgY0wgYowgYcwgYQEFGO4L6th
# 9YOQlpUFCwAknFApM+x5MGwwVqRUMFIxCzAJBgNVBAYTAkJFMRkwFwYDVQQKExBH
# bG9iYWxTaWduIG52LXNhMSgwJgYDVQQDEx9HbG9iYWxTaWduIFRpbWVzdGFtcGlu
# ZyBDQSAtIEcyAhIRIdaZp2SXPvH4Qn7pGcxTQRQwDQYJKoZIhvcNAQEBBQAEggEA
# YooCh8xmdHof7PgIOTH8I4F63vPlyIeq4ClQw1lzlDot3yzjDXwdhKIHidJ/nfbH
# efxIS59M76GxaZ/jBTIgpos5gimiGEzJkSfXiLYQSTHO3gNkUIMYJLTdX4YvcC+p
# WqC1ooYSVYHCOKXzemrJ2eDdOjLP/MwOSWRSBS3TzBaSv/eJzDbgo/CzM1KHbTgv
# MWYvxV5a45/gNEKGBVE222hR2YZUDEHm6JNMRKm/BgBB5FipAAiLCqlEC9rT5qWU
# zDEKW0+gUkAfluMED8aS+Sa3BEUPF2ZF3HlmqxR1BwrSGUcPZvRg4ljBpgm2yvfv
# E9M2KoE3uJNwtA2GPHPz3Q==
# SIG # End signature block
