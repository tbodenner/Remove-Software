#Requires -RunAsAdministrator

$ComputerName = $env:computername

function Format-Output {
	param ([string]$Text)
	$Output = "[--|{0}| {1}" -f $ComputerName, $Text
    Write-Host $Output
}

function Get-AllUserPackages {
	try {
		# return our packages
		Get-AppxPackage -AllUsers
	}
	catch [TypeInitializationException] {
		# type initialization error
		Format-Output -Text "-- Get-AppxPackage failed. TypeInitializationException"
		# return null
		$Null
	}
	catch {
		# all other errors
		Format-Output -Text "-- Get-AppxPackage failed. Unknown reason."
		# return null
		$Null
	}
}

function Get-InstalledPackage {
	param (
		[string]$Name,
		[psobject[]]$AllPackages
	)
	# check if all packages is null
	if ($Null -eq $AllPackages) {
		# all packages is null, so return null
		$Null
	}
	else {
		# otherwise, get our targeted packages from all packages
		$AllPackages | Where-Object { $_.Name -like "$($Name)*" }
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
		Remove-Item -Path $FolderName -Recurse -Force #-ErrorAction SilentlyContinue | Out-Null
		Format-Output '-- Removed App Folder'
	}
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
		Format-Output -Text "-- Remove-AppxPackage failed. TypeInitializationException"
	}
	catch {
		Format-Output -Text "-- Remove-AppxPackage failed. Unknown reason"
	}
}

function Remove-Package {
	param (
		[string]$Name,
		[string[]]$Services = @(),
		[psobject[]]$Packages
	)

	# if the package has known services, stop them
	foreach ($Service in $Services) {
		Stop-Service $Service -ErrorAction SilentlyContinue
	}

	# if packages is null, we have nothing to do. remove any known folders
	if ($Null -eq $Packages) {
		Format-Output '-- Packages is NULL, removing folders'
		Remove-AllFolders -Name $Name
		return @(0, 0)
	}

	# remove each instance of the installed package
	foreach ($Pkg in $Packages) {
		# remove each user
		foreach ($User in $Pkg.PackageUserInformation) {
			$InstallState = $Pkg.PackageUserInformation.InstallState
			# pending removal, skip it
			if ($InstallState -eq 'Installed(pending removal)') {
				Format-Output "-- Pending Removal '$($Pkg.PackageFullName)'"
				continue
			}
			$UserSid = $User.UserSecurityId.Sid
			# system user, skip it
			if ($UserSid -eq 'S-1-5-18') { continue }
			# no package name, skip it
			if ($Pkg.PackageFullName -eq '') { continue }
			# no user sid, skip it
			if ($UserSid -eq '') { continue }
			# remove the package for the a user
			Format-Output "-- Removing '$($Pkg.PackageFullName)' for User"
			Invoke-RemoveAppxPackage -Package $Pkg.PackageFullName -User $UserSid -All $False
			Format-Output "-- Removed '$($Name)'"
			Format-Output "---- User '$($UserSid)'"
			}
		# remove for all users
		Format-Output "-- Removing '$($Pkg.PackageFullName)' All Users"
		Invoke-RemoveAppxPackage -Package $Pkg.PackageFullName -All $True
		Format-Output "-- Removed '$($Name)'"
	}

	Remove-AllFolders -Name $Name

	$Global:AllPackages = Get-AppxPackage -AllUsers

	$PackageResult = Get-InstalledPackage -Name $Name -AllPackages $Global:AllPackages
	if ($Null -eq $PackageResult) {
		Format-Output "-- Verified '$($Name)' was removed"
		return @(0, 1)
	}
	else {
		Format-Output "-- '$($Name)' was NOT removed"
		return @(0, 0)
	}
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
	[string[]]$Services

	WindowsApp([string]$Message, [string]$PackageName, [string[]]$Services) {
		$this.Message = $Message
		$this.PackageName = $PackageName
		$this.Services = $Services
	}
}

try {
	$SkipCount = 0
	$UninstallCount = 0

	$WindowsAppsToRemove = @()
	$WindowsAppsToRemove += [WindowsApp]::new("Uninstalling AMD Radeon Software", "AdvancedMicroDevicesInc-2.AMDRadeonSoftware", @())
	$WindowsAppsToRemove += [WindowsApp]::new("Uninstalling DuckDuckGo", "DuckDuckGo", @())
	$WindowsAppsToRemove += [WindowsApp]::new("Uninstalling Waves Audio", "WavesAudio", @('Waves Audio Services'))
	$WindowsAppsToRemove += [WindowsApp]::new("Uninstalling Bing Wallpaper", "Microsoft.BingWallpaper", @())

	Format-Output "Connected"
	# if running as powershell 7, import the appx modules
	if ($PSVersionTable.PSVersion.Major -ge 7) {
		Import-Module Appx -UseWindowsPowerShell -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
	}

	# get all the packages
	$Global:AllPackages = Get-AppxPackage -AllUsers

	foreach ($WindowsApp in $WindowsAppsToRemove) {
		$WindowsAppData = [WindowsApp]$WindowsApp
		$Packages = Get-InstalledPackage -Name $WindowsAppData.PackageName -AllPackages $Global:AllPackages
		if ($Null -eq $Packages) {
			Format-Output "-- '$($WindowsAppData.PackageName)' not found"
		}
		else {
			$WindowsAppData = [WindowsApp]$WindowsApp
			$Package = $Packages | Where-Object { $_.Name -like "$($WindowsAppData.PackageName)*" }
			if ($Null -ne $Package) {
				Format-Output $WindowsAppData.Message
				$Result = Remove-Package -Name $WindowsAppData.PackageName -Services $WindowsAppData.Services -Packages $Packages
				if ($Null -ne $Result) {
					$SkipCount += $Result[0]
					$UninstallCount += $Result[1]
				}
			}
		}
	}

	# update the reigistry to try and stop Windows from downloading the extra software packages
	Set-DisableAppsForDevices

	if ($PSVersionTable.PSVersion.Major -lt 7) {
		Format-Output "Running Hardware Inventory Cycle"
		Invoke-WmiMethod -Namespace 'root\ccm' -Class 'sms_client' -Name 'TriggerSchedule' -ArgumentList '{00000000-0000-0000-0000-000000000001}'
	}
	else {
		Format-Output "Skipped Hardware Inventory Cycle due to PowerShell7"
	}

	Format-Output "Done"
	@($SkipCount, $UninstallCount)
}
catch {
	Format-Output "-- Error caught in script. Check error file."
	Write-Error -Message "$($ComputerName): $($_)"
	$Null
}
