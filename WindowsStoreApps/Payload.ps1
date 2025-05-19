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
	if (($Null -ne $Name) -and ($Name -ne '')) {
		$AppPath = @(
			"C:\Program Files\WindowsApps\$($Name)*"
			"C:\ProgramData\Packages\$($Name)*"
		)
		$Paths = Resolve-Path -Path $AppPath
		foreach ($Path in $Paths) {
			Remove-PackageFolder -FolderName $Path
		}

		# get user folders
		$UserFolders = Get-ChildItem -Path C:\Users\ -Directory

		foreach ($User in $UserFolders) {
			# create full user folder path
			$UserPath = Join-Path -Path 'C:\Users\' -ChildPath $User 
			$UserPath = Join-Path -Path $UserPath -ChildPath "AppData\Local\Microsoft\WindowsApps\$($Name)*"
			# get all the paths that match our user path
			$UserPaths = Resolve-Path -Path $UserPath
			# remove each path
			foreach ($Path in $UserPaths) {
				Remove-PackageFolder -FolderName $Path
			}
		}

		if ($Name -eq 'WavesAudio') {
			$DellPath = 'C:\ProgramData\Dell\'
			Remove-PackageFolder -FolderName $DellPath
			$WavesPath = 'C:\Program Files\Waves\'
			Remove-PackageFolder -FolderName $WavesPath
		}

		if ($Name -eq 'AdvancedMicroDevicesInc') {
			$DellPath = 'C:\ProgramData\AMD\'
			Remove-PackageFolder -FolderName $DellPath
			$WavesPath = 'C:\Program Files\AMD\'
			Remove-PackageFolder -FolderName $WavesPath
		}
	}
}

function Remove-PackageFolder {
	param (
		[string]$FolderName
	)
	if ((Test-Path -Path $FolderName -Type Container) -eq $True) {
		Remove-Item -Path $FolderName -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
		Format-Output "-- Removed App Folder '$(Split-Path -Path $FolderName -Leaf)'"
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
		[WindowsApp]$WindowsApp,
		[psobject[]]$Packages
	)

	# if we get an empty name, don't do anything
	if (($Null -eq $WindowsApp) -or ($Null -eq $WindowsApp.PackageName) -or ($WindowsApp.PackageName -eq '')) {
		Format-Output "-- App name was null or empty"
		return @(0, 0)
	}

	# if the package has any known services, stop them
	foreach ($Service in $WindowsApp.Services) {
		Stop-Service $Service -Force -ErrorAction SilentlyContinue
	}
	# if the package has any known processes, stop them
	foreach ($Process in $WindowsApp.Processes) {
		Stop-Process -Name $Process -Force -ErrorAction SilentlyContinue
	}

	# if packages is null, we have nothing to do. remove any known folders
	if ($Null -eq $Packages) {
		Format-Output '-- Packages is NULL, removing folders'
		Remove-AllFolders -Name $WindowsApp.PackageName
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
			# check if we have a user sid
			if (($Null -eq $UserSid) -or ($UserSid -eq '')) {
				# since we have no user sid, remove for all users
				Format-Output "-- Removing '$($Pkg.PackageFullName)' for User"
				Invoke-RemoveAppxPackage -Package $Pkg.PackageFullName -All $False
				Format-Output "---- Unknown User"
			}
			else {
				# otherwise, remove the package for the user sid
				Format-Output "-- Removing '$($Pkg.PackageFullName)' for a single user"
				Invoke-RemoveAppxPackage -Package $Pkg.PackageFullName -User $UserSid -All $False
				Format-Output "---- User '$($UserSid)'"
			}
		}
		# remove for all users
		Format-Output "-- Removing '$($Pkg.PackageFullName)' for all users"
		Invoke-RemoveAppxPackage -Package $Pkg.PackageFullName -All $True
		Format-Output "-- Removed '$($Pkg.PackageFullName)'"
	}

	Remove-AllFolders -Name $WindowsApp.PackageName

	$Global:AllPackages = Get-AppxPackage -AllUsers

	$PackageResult = Get-InstalledPackage -Name $WindowsApp.PackageName -AllPackages $Global:AllPackages
	if ($Null -eq $PackageResult) {
		Format-Output "-- Verified '$($WindowsApp.PackageName)' was removed"
		return @(0, 1)
	}
	else {
		Format-Output "-- '$($WindowsApp.PackageName)' was NOT removed"
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

function Set-AppxLibrary {
	Add-Type -AssemblyName "System.EnterpriseServices"
	$publish = [System.EnterpriseServices.Internal.Publish]::new()
	
	@(
		'System.Numerics.Vectors.dll',
		'System.Runtime.CompilerServices.Unsafe.dll',
		'System.Security.Principal.Windows.dll',
		'System.Memory.dll'
	) | ForEach-Object {
		$dllPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\$_"
		$publish.GacInstall($dllPath)
	}
}

class WindowsApp {
	[string]$PackageName
	[string[]]$Services
	[string[]]$Processes

	WindowsApp([string]$PackageName, [string[]]$Services, [string[]]$Processes) {
		$this.PackageName = $PackageName
		$this.Services = $Services
		$this.Processes = $Processes
	}
}

try {
	$SkipCount = 0
	$UninstallCount = 0

	Format-Output "Connected"

	# services to stop for amd radeon
	$AmdServices = @(
		'AMD Crash Defender Service'
		'AMD External Events Utility'
		'RadeonSoftware'
		'AMDRSServ'
		'AMDRSSrcExt'
	)

	# create our windows app objects
	$AmdWindowsApp = [WindowsApp]::new("AdvancedMicroDevicesInc", $AmdServices, @())
	$DuckWindowsApp = [WindowsApp]::new("DuckDuckGo", @(), @())
	$WavesWindowsApp = [WindowsApp]::new("WavesAudio", @('Waves Audio Services'), @())
	$BingWindowsApp = [WindowsApp]::new("Microsoft.BingWallpaper", @(), @('BingWallpaper'))

	# add our apps to an array
	$WindowsAppsToRemove = @(
		$AmdWindowsApp
		$DuckWindowsApp
		$WavesWindowsApp
		$BingWindowsApp
	)
	
	# apply a fix to get Appx working in remote sessions
	Set-AppxLibrary

	# get all the packages
	$Global:AllPackages = Get-AppxPackage -AllUsers
	
	if ($Null -ne $Global:AllPackages) {
		foreach ($WindowsApp in $WindowsAppsToRemove) {
			$WindowsAppData = [WindowsApp]$WindowsApp
			$Packages = Get-InstalledPackage -Name $WindowsAppData.PackageName -AllPackages $Global:AllPackages
			if ($Null -ne $Packages) {
				$WindowsAppData = [WindowsApp]$WindowsApp
				$Package = $Packages | Where-Object { $_.Name -like "$($WindowsAppData.PackageName)*" }
				if ($Null -ne $Package) {
					Format-Output "Uninstalling $($WindowsAppData.PackageName)"
					$Result = Remove-Package -WindowsApp $WindowsAppData -Packages $Packages
					if ($Null -eq $Result) {
						Format-Output "-- Null result for $($WindowsAppData.PackageName)"
					}
					else {
						$SkipCount += $Result[0]
						$UninstallCount += $Result[1]
					}
				}
			}
		}
	}

	# update the registry to try and stop Windows from downloading the extra software packages
	Set-DisableAppsForDevices
	
	# run a hardware update cycle
	if ($PSVersionTable.PSVersion.Major -lt 7) {
		Format-Output "Running Hardware Inventory Cycle"
		Invoke-WmiMethod -Namespace 'root\ccm' -Class 'sms_client' -Name 'TriggerSchedule' -ArgumentList '{00000000-0000-0000-0000-000000000001}'
	}
	else {
		Format-Output "Skipped Hardware Inventory Cycle due to PowerShell7"
	}
	
	# done, return our results
	Format-Output "Done"
	@($SkipCount, $UninstallCount)
}
catch {
	Format-Output "-- Error caught in script. Check error file."
	Write-Error -Message "$($ComputerName): $($_)"
	$Null
}
