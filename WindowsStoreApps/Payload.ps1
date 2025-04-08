#Requires -RunAsAdministrator

$ComputerName = $env:computername

function Format-Output {
	param ([string]$Text)
	$Output = "[--|{0}| {1}" -f $ComputerName, $Text
    Write-Host $Output
}

function Get-PackageIsInstalled {
	param (
		[string]$Name
	)
	try {
		$Packages = Get-AppxPackage -AllUsers "$($Name)*" -ErrorAction SilentlyContinue
	}
	catch [TypeInitializationException]{
		$PShell = 'C:\Program Files\PowerShell\7\pwsh.exe'
		$PShellArgs = "Get-AppxPackage -Name $($Name)* -AllUsers"
		#Format-Output -Text "-- Get-AppxPackage failed. Running with pwsh.exe"
		$Packages = & $PShell -Command { $PShellArgs }
	}
	if ($Null -ne $Packages) {
		# return @(packages, isinstalled)
		return @($Null, $False)
	}
	foreach ($Pkg in $Packages) {
		if ($Null -ne $Pkg) {
			if ($Pkg.PackageUserInformation.Count -gt 0) {
				if ($Pkg.PackageUserInformation.Count -eq 1) {
					if ($Pkg.PackageUserInformation[0].Sid -eq 'S-1-5-18') {
						# return @(packages, isinstalled)
						return @($Null, $False)
					}
					else {
						# return @(packages, isinstalled)
						return @($Packages, $True)
					}
				}
				# return @(packages, isinstalled)
				return @($Packages, $True)
			}
		}
	}
}

function Remove-AllFolders {
	param (
		[string]$Name
	)
	$AppPath = "C:\Program Files\WindowsApps\$($Name)*"
	$Paths = Resolve-Path -Path $AppPath
	foreach ($Path in $Paths) {
		Remove-PackageFolder -FolderName $Path
	}

	if ($Name -eq 'WavesAudio') {
		$DellPath = 'C:\ProgramData\Dell\'
		Remove-PackageFolder -FolderName $DellPath
		$WavesPath = 'C:\Program Files\Waves\'
		Remove-PackageFolder -FolderName $WavesPath
	}
}

function Remove-PackageFolder {
	param (
		[string]$FolderName
	)
	if ((Test-Path -Path $FolderName -Type Container) -eq $True) {
		Remove-Item -Path $FolderName -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
		Format-Output '-- Removed App Folder'
	}
	#else {
	#	Format-Output '-- App folder not found'
	#}
}

function Invoke-RemoveAppxPackage {
	param (
		[string]$PackageName,
		[string]$UserSid,
		[bool]$All
	)
	try {
		if ($All) {			
			Remove-AppxPackage -Package $PackageName -AllUsers -ErrorAction SilentlyContinue
		}
		else {
			Remove-AppxPackage -Package $PackageName -User $UserSid -ErrorAction SilentlyContinue
		}
	}
	catch [TypeInitializationException]{
		$PShell = 'C:\Program Files\PowerShell\7\pwsh.exe'
		$PShellArgs = "Remove-AppxPackage -Package $($PackageName) -User $UserSid"
		if ($All -eq $False) {
			$PShellArgs = "Remove-AppxPackage -Package $($PackageName) -AllUsers"
		}
		#Format-Output -Text "-- Remove-AppxPackage failed. Running with pwsh.exe"
		& $PShell -Command { $PShellArgs }
	}
}

function Remove-Package {
	param (
		[string]$Name,
		[bool]$All = $True,
		[string[]]$Services = @(),
		$PackageResult
	)
	foreach ($Service in $Services) {
		Stop-Service $Service -ErrorAction SilentlyContinue
	}

	#$PackageResult = Get-PackageIsInstalled -Name $Name
	if ($Null -eq $PackageResult) {
		Format-Output '-- Skipped (Not found)'
		Remove-AllFolders -Name $Name
		return @(0, 0)
	}
	$Packages = $PackageResult[0]
	$IsInstalled = $PackageResult[1]

	if (($Null -eq $Packages) -and ($IsInstalled -eq $False)) {
		Format-Output '-- Skipped (System User)'
		Remove-AllFolders -Name $Name
		return @(0, 0)
	}

	foreach ($Pkg in $Packages) {
		foreach ($Sid in $Pkg.PackageUserInformation) {
			$UserSid = $Sid.UserSecurityId.Sid
			if ($UserSid -eq 'S-1-5-18') { continue }
			if ($All) {			
				Invoke-RemoveAppxPackage -Package $Pkg.PackageFullName -All $True
				Format-Output "-- Removed '$($Name)'"
			}
			else {
				Invoke-RemoveAppxPackage -Package $Pkg.PackageFullName -User $UserSid -All $False
				Format-Output "-- Removed '$($Name)' (User)"
			}
		}
	}

	Remove-AllFolders -Name $Name

	$PackageResult = Get-PackageIsInstalled -Name $Name
	if ($Null -eq $PackageResult) {
		Format-Output '-- Uninstalled (Not Found)'
		return @(0, 1)
	}
	$Packages = $PackageResult[0]
	$IsInstalled = $PackageResult[1]

	if (($Null -eq $Packages) -and ($IsInstalled -eq $False)) {
		Format-Output '-- Uninstalled (System User)'
		return @(0, 1)
	}

	Format-Output "-- Error"
	return @(0, 0)
}

# add or update a value in the registry
function Add-RegistryKey {
	param (
		[string]$Path,
		[string]$KeyName,
		[Microsoft.Win32.RegistryValueKind]$KeyType,
		$Value
	)
	
	# check if path exists
	if ((Test-Path -Path $Path -Type Container) -eq $False) {
		# if the path is missing, then create it
		New-Item -Path $Path | Out-Null
	}

	# create the key if missing and set the value
	Set-ItemProperty -Path $Path -Name $KeyName -Value $Value -Type $KeyType | Out-Null
}

# add registry value to stop downloading manufacturers' apps for installed devices
function Set-DisableAppsForDevices {
	# the registry values
	$RegPath = 'HKCU:\Software\Policies\Microsoft\Windows\DeviceInstall'
	$RegKeyName = 'AllowOSManagedDriverInstallationToUI'
	$RegKeyValue = 0
	# add the key
	Add-RegistryKey -Path $RegPath -KeyName $RegKeyName -Value $RegKeyValue -KeyType DWord

	# the registry values
	$RegPath = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Device Installer'
	$RegKeyName = 'DisableCoInstallers'
	$RegKeyValue = 1
	# add the key
	Add-RegistryKey -Path $RegPath -KeyName $RegKeyName -Value $RegKeyValue -KeyType DWord
}

class WindowsApp {
	[string]$Message
	[string]$PackageName
	[bool]$AllUsers
	[string[]]$Services

	WindowsApp([string]$Message, [string]$PackageName, [bool]$AllUsers, [string[]]$Services) {
		$this.Message = $Message
		$this.PackageName = $PackageName
		$this.AllUsers = $AllUsers
		$this.Services = $Services
	}
}

try {
	$SkipCount = 0
	$UninstallCount = 0

	$WindowsAppsToRemove = @()
	$WindowsAppsToRemove += [WindowsApp]::new("Uninstalling AMD Radeon Software", "AdvancedMicroDevicesInc-2.AMDRadeonSoftware", $True, @())
	$WindowsAppsToRemove += [WindowsApp]::new("Uninstalling DuckDuckGo", "DuckDuckGo", $True, @())
	$WindowsAppsToRemove += [WindowsApp]::new("Uninstalling Waves Audio (Each User)", "WavesAudio", $False, @('Waves Audio Services'))
	$WindowsAppsToRemove += [WindowsApp]::new("Uninstalling Waves Audio (All Users)", "WavesAudio", $True, @('Waves Audio Services'))
	$WindowsAppsToRemove += [WindowsApp]::new("Uninstalling Bing Wallpaper", "Microsoft.BingWallpaper", $False, @())

	Format-Output "Connected"
	foreach ($WindowsApp in $WindowsAppsToRemove) {
		$WindowsAppData = [WindowsApp]$WindowsApp
		$PackageResult = Get-PackageIsInstalled -Name $WindowsAppData.PackageName
		if ($Null -ne $PackageResult) {
			Format-Output $WindowsAppData.Message
			$Result = Remove-Package -Name $WindowsAppData.PackageName -All $WindowsAppData.AllUsers -Services $WindowsAppData.Services -PackageResult $PackageResult
			if ($Null -ne $Result) {
				$SkipCount += $Result[0]
				$UninstallCount += $Result[1]
			}
		}
	}

	Set-DisableAppsForDevices

	Format-Output "Running Hardware Inventory Cycle"
	Invoke-WmiMethod -Namespace 'root\ccm' -Class 'sms_client' -Name 'TriggerSchedule' -ArgumentList '{00000000-0000-0000-0000-000000000001}'

	Format-Output "Done`n"

	return @($SkipCount, $UninstallCount)
}
catch {
	Format-Output "-- Error caught in script. Check error file."
	Write-Error $_
}
