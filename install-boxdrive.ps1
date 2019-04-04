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


Function New-WPFMessageBox
{
	
	# Define Parameters
	[CmdletBinding()]
	Param
	(
		# The popup Content
		[Parameter(Mandatory = $True, Position = 0)]
		[Object]$Content,
		# The window title

		[Parameter(Mandatory = $false, Position = 1)]
		[string]$Title,
		# The buttons to add

		[Parameter(Mandatory = $false, Position = 2)]
		[ValidateSet('OK', 'OK-Cancel', 'Abort-Retry-Ignore', 'Yes-No-Cancel', 'Yes-No', 'Retry-Cancel', 'Cancel-TryAgain-Continue', 'None')]
		[array]$ButtonType = 'OK',
		# The buttons to add

		[Parameter(Mandatory = $false, Position = 3)]
		[array]$CustomButtons,
		# Content font size

		[Parameter(Mandatory = $false, Position = 4)]
		[int]$ContentFontSize = 14,
		# Title font size

		[Parameter(Mandatory = $false, Position = 5)]
		[int]$TitleFontSize = 14,
		# BorderThickness

		[Parameter(Mandatory = $false, Position = 6)]
		[int]$BorderThickness = 0,
		# CornerRadius

		[Parameter(Mandatory = $false, Position = 7)]
		[int]$CornerRadius = 8,
		# ShadowDepth

		[Parameter(Mandatory = $false, Position = 8)]
		[int]$ShadowDepth = 3,
		# BlurRadius

		[Parameter(Mandatory = $false, Position = 9)]
		[int]$BlurRadius = 20,
		# WindowHost

		[Parameter(Mandatory = $false, Position = 10)]
		[object]$WindowHost,
		# Timeout in seconds,

		[Parameter(Mandatory = $false, Position = 11)]
		[int]$Timeout,
		# Code for Window Loaded event,

		[Parameter(Mandatory = $false, Position = 12)]
		[scriptblock]$OnLoaded,
		# Code for Window Closed event,

		[Parameter(Mandatory = $false, Position = 13)]
		[scriptblock]$OnClosed
		
	)
	
	# Dynamically Populated parameters
	DynamicParam
	{
		
		# Add assemblies for use in PS Console 
		Add-Type -AssemblyName System.Drawing, PresentationCore
		
		# ContentBackground
		$ContentBackground = 'ContentBackground'
		$AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
		$ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
		$ParameterAttribute.Mandatory = $False
		$AttributeCollection.Add($ParameterAttribute)
		$RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
		$arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name
		$ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
		$AttributeCollection.Add($ValidateSetAttribute)
		$PSBoundParameters.ContentBackground = "White"
		$RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentBackground, [string], $AttributeCollection)
		$RuntimeParameterDictionary.Add($ContentBackground, $RuntimeParameter)
		
		
		# FontFamily
		$FontFamily = 'FontFamily'
		$AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
		$ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
		$ParameterAttribute.Mandatory = $False
		$AttributeCollection.Add($ParameterAttribute)
		$arrSet = [System.Drawing.FontFamily]::Families.Name | Select -Skip 1
		$ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
		$AttributeCollection.Add($ValidateSetAttribute)
		$RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($FontFamily, [string], $AttributeCollection)
		$RuntimeParameterDictionary.Add($FontFamily, $RuntimeParameter)
		$PSBoundParameters.FontFamily = "Segoe UI"
		
		# TitleFontWeight
		$TitleFontWeight = 'TitleFontWeight'
		$AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
		$ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
		$ParameterAttribute.Mandatory = $False
		$AttributeCollection.Add($ParameterAttribute)
		$arrSet = [System.Windows.FontWeights] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name
		$ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
		$AttributeCollection.Add($ValidateSetAttribute)
		$PSBoundParameters.TitleFontWeight = "Normal"
		$RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleFontWeight, [string], $AttributeCollection)
		$RuntimeParameterDictionary.Add($TitleFontWeight, $RuntimeParameter)
		
		# ContentFontWeight
		$ContentFontWeight = 'ContentFontWeight'
		$AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
		$ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
		$ParameterAttribute.Mandatory = $False
		$AttributeCollection.Add($ParameterAttribute)
		$arrSet = [System.Windows.FontWeights] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name
		$ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
		$AttributeCollection.Add($ValidateSetAttribute)
		$PSBoundParameters.ContentFontWeight = "Normal"
		$RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentFontWeight, [string], $AttributeCollection)
		$RuntimeParameterDictionary.Add($ContentFontWeight, $RuntimeParameter)
		
		
		# ContentTextForeground
		$ContentTextForeground = 'ContentTextForeground'
		$AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
		$ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
		$ParameterAttribute.Mandatory = $False
		$AttributeCollection.Add($ParameterAttribute)
		$arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name
		$ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
		$AttributeCollection.Add($ValidateSetAttribute)
		$PSBoundParameters.ContentTextForeground = "Black"
		$RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentTextForeground, [string], $AttributeCollection)
		$RuntimeParameterDictionary.Add($ContentTextForeground, $RuntimeParameter)
		
		# TitleTextForeground
		$TitleTextForeground = 'TitleTextForeground'
		$AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
		$ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
		$ParameterAttribute.Mandatory = $False
		$AttributeCollection.Add($ParameterAttribute)
		$arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name
		$ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
		$AttributeCollection.Add($ValidateSetAttribute)
		$PSBoundParameters.TitleTextForeground = "Black"
		$RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleTextForeground, [string], $AttributeCollection)
		$RuntimeParameterDictionary.Add($TitleTextForeground, $RuntimeParameter)
		
		# BorderBrush
		$BorderBrush = 'BorderBrush'
		$AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
		$ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
		$ParameterAttribute.Mandatory = $False
		$AttributeCollection.Add($ParameterAttribute)
		$arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name
		$ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
		$AttributeCollection.Add($ValidateSetAttribute)
		$PSBoundParameters.BorderBrush = "Black"
		$RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($BorderBrush, [string], $AttributeCollection)
		$RuntimeParameterDictionary.Add($BorderBrush, $RuntimeParameter)
		
		
		# TitleBackground
		$TitleBackground = 'TitleBackground'
		$AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
		$ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
		$ParameterAttribute.Mandatory = $False
		$AttributeCollection.Add($ParameterAttribute)
		$arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name
		$ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
		$AttributeCollection.Add($ValidateSetAttribute)
		$PSBoundParameters.TitleBackground = "White"
		$RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleBackground, [string], $AttributeCollection)
		$RuntimeParameterDictionary.Add($TitleBackground, $RuntimeParameter)
		
		# ButtonTextForeground
		$ButtonTextForeground = 'ButtonTextForeground'
		$AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
		$ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
		$ParameterAttribute.Mandatory = $False
		$AttributeCollection.Add($ParameterAttribute)
		$arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name
		$ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
		$AttributeCollection.Add($ValidateSetAttribute)
		$PSBoundParameters.ButtonTextForeground = "Black"
		$RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ButtonTextForeground, [string], $AttributeCollection)
		$RuntimeParameterDictionary.Add($ButtonTextForeground, $RuntimeParameter)
		
		# Sound
		$Sound = 'Sound'
		$AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
		$ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
		$ParameterAttribute.Mandatory = $False
		#$ParameterAttribute.Position = 14
		$AttributeCollection.Add($ParameterAttribute)
		$arrSet = (Get-ChildItem "$env:SystemDrive\Windows\Media" -Filter Windows* | Select -ExpandProperty Name).Replace('.wav', '')
		$ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
		$AttributeCollection.Add($ValidateSetAttribute)
		$RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($Sound, [string], $AttributeCollection)
		$RuntimeParameterDictionary.Add($Sound, $RuntimeParameter)
		
		return $RuntimeParameterDictionary
	}
	
	Begin
	{
		Add-Type -AssemblyName PresentationFramework
	}
	
	Process
	{
		
		# Define the XAML markup
		[XML]$Xaml = @"
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        x:Name="Window" Title="" SizeToContent="WidthAndHeight" WindowStartupLocation="CenterScreen" WindowStyle="None" ResizeMode="NoResize" AllowsTransparency="True" Background="Transparent" Opacity="1">
    <Window.Resources>
        <Style TargetType="{x:Type Button}">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border>
                            <Grid Background="{TemplateBinding Background}">
                                <ContentPresenter />
                            </Grid>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Border x:Name="MainBorder" Margin="10" CornerRadius="$CornerRadius" BorderThickness="$BorderThickness" BorderBrush="$($PSBoundParameters.BorderBrush)" Padding="0" >
        <Border.Effect>
            <DropShadowEffect x:Name="DSE" Color="Black" Direction="270" BlurRadius="$BlurRadius" ShadowDepth="$ShadowDepth" Opacity="0.6" />
        </Border.Effect>
        <Border.Triggers>
            <EventTrigger RoutedEvent="Window.Loaded">
                <BeginStoryboard>
                    <Storyboard>
                        <DoubleAnimation Storyboard.TargetName="DSE" Storyboard.TargetProperty="ShadowDepth" From="0" To="$ShadowDepth" Duration="0:0:1" AutoReverse="False" />
                        <DoubleAnimation Storyboard.TargetName="DSE" Storyboard.TargetProperty="BlurRadius" From="0" To="$BlurRadius" Duration="0:0:1" AutoReverse="False" />
                    </Storyboard>
                </BeginStoryboard>
            </EventTrigger>
        </Border.Triggers>
        <Grid >
            <Border Name="Mask" CornerRadius="$CornerRadius" Background="$($PSBoundParameters.ContentBackground)" />
            <Grid x:Name="Grid" Background="$($PSBoundParameters.ContentBackground)">
                <Grid.OpacityMask>
                    <VisualBrush Visual="{Binding ElementName=Mask}"/>
                </Grid.OpacityMask>
                <StackPanel Name="StackPanel" >                   
                    <TextBox Name="TitleBar" IsReadOnly="True" IsHitTestVisible="False" Text="$Title" Padding="10" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="$TitleFontSize" Foreground="$($PSBoundParameters.TitleTextForeground)" FontWeight="$($PSBoundParameters.TitleFontWeight)" Background="$($PSBoundParameters.TitleBackground)" HorizontalAlignment="Stretch" VerticalAlignment="Center" Width="Auto" HorizontalContentAlignment="Center" BorderThickness="0"/>
                    <DockPanel Name="ContentHost" Margin="0,10,0,10"  >
                    </DockPanel>
                    <DockPanel Name="ButtonHost" LastChildFill="False" HorizontalAlignment="Center" >
                    </DockPanel>
                </StackPanel>
            </Grid>
        </Grid>
    </Border>
</Window>
"@
		
		[XML]$ButtonXaml = @"
<Button xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Width="Auto" Height="30" FontFamily="Segui" FontSize="16" Background="Transparent" Foreground="White" BorderThickness="1" Margin="10" Padding="20,0,20,0" HorizontalAlignment="Right" Cursor="Hand"/>
"@
		
		[XML]$ButtonTextXaml = @"
<TextBlock xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="16" Background="Transparent" Foreground="$($PSBoundParameters.ButtonTextForeground)" Padding="20,5,20,5" HorizontalAlignment="Center" VerticalAlignment="Center"/>
"@
		
		[XML]$ContentTextXaml = @"
<TextBlock xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Text="$Content" Foreground="$($PSBoundParameters.ContentTextForeground)" DockPanel.Dock="Right" HorizontalAlignment="Center" VerticalAlignment="Center" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="$ContentFontSize" FontWeight="$($PSBoundParameters.ContentFontWeight)" TextWrapping="Wrap" Height="Auto" MaxWidth="500" MinWidth="50" Padding="10"/>
"@
		
		# Load the window from XAML
		$Window = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml))
		
		# Custom function to add a button
		Function Add-Button
		{
			Param ($Content)
			$Button = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ButtonXaml))
			$ButtonText = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ButtonTextXaml))
			$ButtonText.Text = "$Content"
			$Button.Content = $ButtonText
			$Button.Add_MouseEnter({
					$This.Content.FontSize = "17"
				})
			$Button.Add_MouseLeave({
					$This.Content.FontSize = "16"
				})
			$Button.Add_Click({
					New-Variable -Name WPFMessageBoxOutput -Value $($This.Content.Text) -Option ReadOnly -Scope Script -Force
					$Window.Close()
				})
			$Window.FindName('ButtonHost').AddChild($Button)
		}
		
		# Add buttons
		If ($ButtonType -eq "OK")
		{
			Add-Button -Content "OK"
		}
		
		If ($ButtonType -eq "OK-Cancel")
		{
			Add-Button -Content "OK"
			Add-Button -Content "Cancel"
		}
		
		If ($ButtonType -eq "Abort-Retry-Ignore")
		{
			Add-Button -Content "Abort"
			Add-Button -Content "Retry"
			Add-Button -Content "Ignore"
		}
		
		If ($ButtonType -eq "Yes-No-Cancel")
		{
			Add-Button -Content "Yes"
			Add-Button -Content "No"
			Add-Button -Content "Cancel"
		}
		
		If ($ButtonType -eq "Yes-No")
		{
			Add-Button -Content "Yes"
			Add-Button -Content "No"
		}
		
		If ($ButtonType -eq "Retry-Cancel")
		{
			Add-Button -Content "Retry"
			Add-Button -Content "Cancel"
		}
		
		If ($ButtonType -eq "Cancel-TryAgain-Continue")
		{
			Add-Button -Content "Cancel"
			Add-Button -Content "TryAgain"
			Add-Button -Content "Continue"
		}
		
		If ($ButtonType -eq "None" -and $CustomButtons)
		{
			Foreach ($CustomButton in $CustomButtons)
			{
				Add-Button -Content "$CustomButton"
			}
		}
		
		# Remove the title bar if no title is provided
		If ($Title -eq "")
		{
			$TitleBar = $Window.FindName('TitleBar')
			$Window.FindName('StackPanel').Children.Remove($TitleBar)
		}
		
		# Add the Content
		If ($Content -is [String])
		{
			# Replace double quotes with single to avoid quote issues in strings
			If ($Content -match '"')
			{
				$Content = $Content.Replace('"', "'")
			}
			
			# Use a text box for a string value...
			$ContentTextBox = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ContentTextXaml))
			$Window.FindName('ContentHost').AddChild($ContentTextBox)
		}
		Else
		{
			# ...or add a WPF element as a child
			Try
			{
				$Window.FindName('ContentHost').AddChild($Content)
			}
			Catch
			{
				$_
			}
		}
		
		# Enable window to move when dragged
		$Window.FindName('Grid').Add_MouseLeftButtonDown({
				$Window.DragMove()
			})
		
		# Activate the window on loading
		If ($OnLoaded)
		{
			$Window.Add_Loaded({
					$This.Activate()
					Invoke-Command $OnLoaded
				})
		}
		Else
		{
			$Window.Add_Loaded({
					$This.Activate()
				})
		}
		
		
		# Stop the dispatcher timer if exists
		If ($OnClosed)
		{
			$Window.Add_Closed({
					If ($DispatcherTimer)
					{
						$DispatcherTimer.Stop()
					}
					Invoke-Command $OnClosed
				})
		}
		Else
		{
			$Window.Add_Closed({
					If ($DispatcherTimer)
					{
						$DispatcherTimer.Stop()
					}
				})
		}
		
		
		# If a window host is provided assign it as the owner
		If ($WindowHost)
		{
			$Window.Owner = $WindowHost
			$Window.WindowStartupLocation = "CenterOwner"
		}
		
		# If a timeout value is provided, use a dispatcher timer to close the window when timeout is reached
		If ($Timeout)
		{
			$Stopwatch = New-object System.Diagnostics.Stopwatch
			$TimerCode = {
				If ($Stopwatch.Elapsed.TotalSeconds -ge $Timeout)
				{
					$Stopwatch.Stop()
					$Window.Close()
				}
			}
			$DispatcherTimer = New-Object -TypeName System.Windows.Threading.DispatcherTimer
			$DispatcherTimer.Interval = [TimeSpan]::FromSeconds(1)
			$DispatcherTimer.Add_Tick($TimerCode)
			$Stopwatch.Start()
			$DispatcherTimer.Start()
		}
		
		# Play a sound
		If ($($PSBoundParameters.Sound))
		{
			$SoundFile = "$env:SystemDrive\Windows\Media\$($PSBoundParameters.Sound).wav"
			$SoundPlayer = New-Object System.Media.SoundPlayer -ArgumentList $SoundFile
			$SoundPlayer.Add_LoadCompleted({
					$This.Play()
					$This.Dispose()
				})
			$SoundPlayer.LoadAsync()
		}
		
		# Display the window
		$null = $window.Dispatcher.InvokeAsync{ $window.ShowDialog() }.Wait()
		
	}
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
	Remove-Item $boxsyncstartmenufolder -Recurse -Force -ea SilentlyContinue
	
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
	$Content = "Logon to Box Drive to uninstall Sync, then reboot to complete the cleanup. All Box Sync content will be saved to Box."
	
	$Params = @{
		Content		     = $Content
		FontFamily	     = 'Verdana'
		Title		     = "Box Drive Installed"
		TitleFontSize    = 30
		#TitleTextForeground = 'Red'
		TitleFontWeight  = 'Bold'
		TitleBackground  = 'SteelBlue'
		#FontFamily = 'Lucida Console'
		ButtonType	     = 'OK'
		Timeout		     = 30
	}
	
	New-WPFMessageBox @Params
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
	Remove-Item $boxsyncstartmenufolder -Recurse -Force -ea SilentlyContinue
	
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
	$Content = "Logon to Box Drive to uninstall Sync, then reboot to complete the cleanup. All Box Sync content will be saved to Box."
	
	$Params = @{
		Content		     = $Content
		FontFamily	     = 'Verdana'
		Title		     = "Box Drive Installed"
		TitleFontSize    = 30
		#TitleTextForeground = 'Red'
		TitleFontWeight  = 'Bold'
		TitleBackground  = 'SteelBlue'
		#FontFamily = 'Lucida Console'
		ButtonType	     = 'OK'
		Timeout		     = 30
	}
	
	New-WPFMessageBox @Params
}
. Stop-Logging
# SIG # Begin signature block
# MIIfXAYJKoZIhvcNAQcCoIIfTTCCH0kCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUbkzMK9VyA8kmvIGVOViq58Hp
# TX+gghn1MIIEFDCCAvygAwIBAgILBAAAAAABL07hUtcwDQYJKoZIhvcNAQEFBQAw
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
# FgQUJka6j8+24cTD0Z3crzSMCvdOFLwwDQYJKoZIhvcNAQEBBQAEggEArvMIm5wG
# KEBUd1c/zsqdGnyBI4Bk3j/tFspKfImfZ7hrwgJsoCTchgU5iZxMMIJ5tlqJRP/0
# aExXnLzANqOni8BLNh+OGq2l7PM67kVBHUHxPvFpFHWwUhCcHxfejzU88aj3dVYz
# 8y1NjBqXVsQP3R6DLu0emwIM4vf7cVay5RPpMtYPvMmXdNREj8sqwKcqejBLivXS
# yDf/91gcjXsHrK9Nq0Em2OGpJZf5h2h8DFRWk1JG8rMFyC7KNf7dbNyurdx5fW0z
# ktvVIoMHC3r2uOiB9iXe5DSUBvBqBfTbsL22uMul/345pRidSWCAVKXmgcqHDtnD
# 0DwuYHd/BC+QqKGCAqIwggKeBgkqhkiG9w0BCQYxggKPMIICiwIBATBoMFIxCzAJ
# BgNVBAYTAkJFMRkwFwYDVQQKExBHbG9iYWxTaWduIG52LXNhMSgwJgYDVQQDEx9H
# bG9iYWxTaWduIFRpbWVzdGFtcGluZyBDQSAtIEcyAhIRIdaZp2SXPvH4Qn7pGcxT
# QRQwCQYFKw4DAhoFAKCB/TAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0xOTA0MDQyMTQ2MjZaMCMGCSqGSIb3DQEJBDEWBBQeP2RYpv+g
# vRj0bk0a5Y94A6QNVTCBnQYLKoZIhvcNAQkQAgwxgY0wgYowgYcwgYQEFGO4L6th
# 9YOQlpUFCwAknFApM+x5MGwwVqRUMFIxCzAJBgNVBAYTAkJFMRkwFwYDVQQKExBH
# bG9iYWxTaWduIG52LXNhMSgwJgYDVQQDEx9HbG9iYWxTaWduIFRpbWVzdGFtcGlu
# ZyBDQSAtIEcyAhIRIdaZp2SXPvH4Qn7pGcxTQRQwDQYJKoZIhvcNAQEBBQAEggEA
# Dj6lCc/abTN8ofzOCQpCIOMttRTD0ealY0ivWqqLvJ9SO/IHs8pdpZrvQoev5cPa
# U3YAfjX5odPBlLcBD07usNHzCEtMQkBZrhCcg1KRUk2pLVfMsxE+DW3qnM0oCe8C
# YrVgsc7GKXi4qmhMBdUt/DSXzZegQUjlvx8bWc5iemvbzxUb5H8MJ8FMVHjewwjN
# pa4yPUYK6zsh3/prCECdhU5L1tnZhjgfGgu/r17uzW9aC82rn1gLgVsgrtTPg5S8
# NU41HWPjgurDnsgEe4EPhgTi4EKtJ/XFeozRP8AGODqvyDqnrzdFXewQ8Lwc0VYd
# EDtjHdWGjgoDzJroEtWsNg==
# SIG # End signature block
