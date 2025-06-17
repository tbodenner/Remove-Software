#Requires -RunAsAdministrator

$ComputerName = $env:computername

function Get-AllUserPackages {
	try {
		# return our packages
		Get-AppxPackage -AllUsers
	}
	catch [TypeInitializationException] {
		# type initialization error
		Write-Host "$($ComputerName): -- Get-AppxPackage failed. TypeInitializationException"
		# return null
		$Null
	}
	catch {
		# all other errors
		Write-Host "$($ComputerName): -- Get-AppxPackage failed. Unknown reason."
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
		# our paths to remove
		$AppPaths = @(
			"C:\Program Files\WindowsApps\$($Name)*"
			"C:\ProgramData\Packages\$($Name)*"
			'C:\ProgramData\Dell\' # Waves MaxxAudio
			'C:\Program Files\Waves\' # Waves MaxxAudio
			'C:\ProgramData\AMD\' # AMD Radeon Software
			'C:\Program Files\AMD\' # AMD Radeon Software
			'C:\$WINDOWS.~BT\NewOS\Windows\System32\DriverStore\FileRepository\waves*' # Waves MaxxAudio
		)
		
		# remove each path
		foreach ($AppPath in $AppPaths) {
			# resolve our paths
			$ResolvedPaths = Resolve-Path -Path $AppPath -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
			if ($null -ne $ResolvedPaths) {
				foreach ($ResolvedPath in $ResolvedPaths) {
					Remove-PackageFolder -FolderName $ResolvedPath
				}
			}
		}

		# get user folders
		$UserFolders = Get-ChildItem -Path C:\Users\ -Directory

		foreach ($User in $UserFolders) {
			# create full user folder path
			$UserPath = Join-Path -Path 'C:\Users\' -ChildPath $User 
			$UserPath = Join-Path -Path $UserPath -ChildPath "AppData\Local\Microsoft\WindowsApps\$($Name)*"
			# get all the paths that match our user path
			$UserPaths = Resolve-Path -Path $UserPath -ErrorAction SilentlyContinue
			# remove each path
			foreach ($Path in $UserPaths) {
				Remove-PackageFolder -FolderName $Path
			}
		}

		# remove extra items for AMD, some of them are redundant and need to be cleaned up
		if ($Name -eq "AdvancedMicroDevicesInc") {
			Remove-AmdFolders
		}
	}
}

function Remove-PackageFolder {
	param (
		[string]$FolderName
	)
	if ((Test-Path -Path $FolderName) -eq $True) {
		$IsFolder = (Get-Item -Path $FolderName).PSIsContainer
		Remove-Item -Path $FolderName -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
		if ($IsFolder -eq $true) {
			Write-Host "$($ComputerName): -- Removed App Folder '$(Split-Path -Path $FolderName -Leaf)'"
		}
		else {
			Write-Host "$($ComputerName): -- Removed App File '$(Split-Path -Path $FolderName -Leaf)'"
		}
	}
}

# extra items to remove for AMD Radeon Software
function Remove-AmdFolders {
	try {
		$ProcArray = @(
			"RadeonSoftware",
			"AMDRSServ",
			"AMDRSSrcExt"
		)
		foreach ($Proc in $ProcArray) {
			if ($Null -ne (Get-Process "$($Proc)*")) {
				try {
					Stop-Process -Name $Proc -Force
					Write-Host "$($ComputerName): -- Stopped process '$($Proc)'"
				}
				catch {
					Write-Host "$($ComputerName): -- Failed to stop process '$($Proc)'"
				}
			}
		}

		$ServiceArray = @(
			"AMD Crash Defender Service",
			"AMD External Events Utility"
		)

		foreach ($Serv in $ServiceArray) {
			if ($Null -ne (Get-Service $Serv -ErrorAction SilentlyContinue)) {
				try {
					Stop-Service -Name $Serv
					Write-Host "$($ComputerName): -- Stopped service '$($Serv)'"
					sc.exe delete $Serv | Out-Null
					Write-Host "$($ComputerName): -- Removed service '$($Serv)'"
				}
				catch {
					Write-Host "$($ComputerName): -- Failed to stop or remove service '$($Serv)'"
				}
			}
		}

		$AmdPaths = @(
			"C:\Program Files\AMD\",
			"C:\ProgramData\AMD\"
		)
		$AmdPaths += Resolve-Path -Path "C:\Program Files\WindowsApps\AdvancedMicroDevicesInc*"
		$MaxPathLen = 50
		foreach ($Path in $AmdPaths) {
			if ((Test-Path -Path $Path) -eq $True) {
				Remove-Item -Path $Path -Recurse -Force -ErrorAction SilentlyContinue
				if ((Test-Path -Path $Path) -eq $True) {
					if ($Path.Length -gt $MaxPathLen) {
						Write-Host "$($ComputerName): -- Failed to remove folder '$(Split-Path -Path $Path -Leaf)'"
					}
					else {
						Write-Host "$($ComputerName): -- Failed to remove folder '$($Path)'"
					}
					if ($Path -eq "C:\Program Files\AMD\") {
						Remove-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}\0000\' -Name 'DalDCELogFilePath' -ErrorAction SilentlyContinue
						Remove-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}\0002\' -Name 'DalDCELogFilePath' -ErrorAction SilentlyContinue
						Write-Host "$($ComputerName): -- Removed 'DalDCELogFilePath' registry value"
					}
				}
				else {
					if ($Path.Length -gt $MaxPathLen) {
						Write-Host "$($ComputerName): -- Removed folder '$(Split-Path -Path $Path -Leaf)'"
					}
					else {
						Write-Host "$($ComputerName): -- Removed folder '$($Path)'"
					}
				}
			}
			Remove-Item -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\AMD Catalyst Install Manager' -ErrorAction SilentlyContinue
		}
	}
	catch {
		Write-Host "$($ComputerName): -- Error caught in Remove-AmdFolders"
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
		Write-Host "$($ComputerName): -- Remove-AppxPackage failed. TypeInitializationException"
	}
	catch {
		Write-Host "$($ComputerName): -- Remove-AppxPackage failed. Unknown reason"
	}
}

function Remove-Package {
	param (
		[WindowsApp]$WindowsApp,
		[psobject[]]$Packages
	)

	# if we get an empty name, don't do anything
	if (($Null -eq $WindowsApp) -or ($Null -eq $WindowsApp.PackageName) -or ($WindowsApp.PackageName -eq '')) {
		Write-Host "$($ComputerName): -- App name was null or empty"
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
				Write-Host "$($ComputerName): -- Pending Removal '$($Pkg.PackageFullName)'"
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
				Write-Host "$($ComputerName): -- Removing '$($Pkg.PackageFullName)' for User"
				Invoke-RemoveAppxPackage -Package $Pkg.PackageFullName -All $False
				Write-Host "$($ComputerName): ---- Unknown User"
			}
			else {
				# otherwise, remove the package for the user sid
				Write-Host "$($ComputerName): -- Removing '$($Pkg.PackageFullName)' for a single user"
				Invoke-RemoveAppxPackage -Package $Pkg.PackageFullName -User $UserSid -All $False
				Write-Host "$($ComputerName): ---- User '$($UserSid)'"
			}
		}
		# remove for all users
		Write-Host "$($ComputerName): -- Removing '$($Pkg.PackageFullName)' for all users"
		Invoke-RemoveAppxPackage -Package $Pkg.PackageFullName -All $True
		Write-Host "$($ComputerName): -- Removed '$($Pkg.PackageFullName)'"
	}

	Remove-AllFolders -Name $WindowsApp.PackageName

	$Global:AllPackages = Get-AppxPackage -AllUsers

	$PackageResult = Get-InstalledPackage -Name $WindowsApp.PackageName -AllPackages $Global:AllPackages
	if ($Null -eq $PackageResult) {
		Write-Host "$($ComputerName): -- Verified '$($WindowsApp.PackageName)' was removed"
		return @(0, 1)
	}
	else {
		Write-Host "$($ComputerName): -- '$($WindowsApp.PackageName)' was NOT removed"
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

	Write-Host "$($ComputerName): Connected"

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
					Write-Host "$($ComputerName): Uninstalling $($WindowsAppData.PackageName)"
					$Result = Remove-Package -WindowsApp $WindowsAppData -Packages $Packages
					if ($Null -eq $Result) {
						Write-Host "$($ComputerName): -- Null result for $($WindowsAppData.PackageName)"
					}
					else {
						$SkipCount += $Result[0]
						$UninstallCount += $Result[1]
					}
				}
			}
		}
	}

	# try to remove all folders even if no packages were found
	Write-Host "$($ComputerName): Removing Leftover Folders"
	Remove-AllFolders -Name "NO APP NAME"

	# update the registry to try and stop Windows from downloading the extra software packages
	Set-DisableAppsForDevices
	
	# run a hardware update cycle
	if ($PSVersionTable.PSVersion.Major -lt 7) {
		Write-Host "$($ComputerName): Triggering Hardware Inventory Cycle"
		Invoke-WmiMethod -Namespace 'root\ccm' -Class 'sms_client' -Name 'TriggerSchedule' -ArgumentList '{00000000-0000-0000-0000-000000000001}'
	}
	else {
		Write-Host "$($ComputerName): Skipped Hardware Inventory Cycle due to PowerShell7 profile issue"
	}
	
	# done, return our results
	Write-Host "$($ComputerName): Done"
	@($SkipCount, $UninstallCount)
}
catch {
	Write-Host "$($ComputerName): -- Error caught in script. Check error file."
	Write-Error -Message "$($ComputerName): $($_)"
	$Null
}
