<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.150
	 Created on:   	3/6/2019 1:10 AM
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
  param(
    $Path = "${env:ProgramFiles}\Box\Box\Box.exe",
    
    #$Arguments = 'www.powertheshell.com',
    
    [Parameter(Mandatory=$true)]
    $Computername,
    
    [Parameter(Mandatory=$true)]
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
    $explorerprocesses = @(Get-WmiObject -Query "Select * FROM Win32_Process WHERE Name='explorer.exe'" -ErrorAction SilentlyContinue)
    if ($explorerprocesses.Count -eq 0)
    {
        "No explorer process found / Nobody interactively logged on"
    } else {
        foreach ($i in $explorerprocesses)
        {
            $Username = $i.GetOwner().User
            #$Domain = $i.GetOwner().Domain
            #$Domain + "\" + $Username + " logged on since: " + ($i.ConvertToDateTime($i.CreationDate))
        }
    }
    Start-ProcessInteractive -Computername $env:COMPUTERNAME -Username $Username
}

elseif ($installedversion -eq $Null) {
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
    
    $explorerprocesses = @(Get-WmiObject -Query "Select * FROM Win32_Process WHERE Name='explorer.exe'" -ErrorAction SilentlyContinue)
	if ($explorerprocesses.Count -eq 0)
    {
        "No explorer process found / Nobody interactively logged on"
    } else {
        foreach ($i in $explorerprocesses)
        {
            $Username = $i.GetOwner().User
            #$Domain = $i.GetOwner().Domain
            #$Domain + "\" + $Username + " logged on since: " + ($i.ConvertToDateTime($i.CreationDate))
        }
    }
    Start-ProcessInteractive -Computername $env:COMPUTERNAME -Username $Username
}
. Stop-Logging
