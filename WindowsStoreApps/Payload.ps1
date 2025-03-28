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
		$ArgArray = @(
			"-Command",
			"{ Get-AppxPackage -Name ""$($Name)*"" -AllUsers }"
		)
		Format-Output -Text "-- Get-AppxPackage failed. Running with pwsh.exe"
		$Packages = Start-Process -FilePath $PShell -ArgumentList $ArgArray -Wait -PassThru
	}
	if ($Null -ne $Packages) {
		return @($Null, $False)
	}
	foreach ($Pkg in $Packages) {
		if ($Null -ne $Pkg) {
			if ($Pkg.PackageUserInformation.Count -gt 0) {
				if ($Pkg.PackageUserInformation.Count -eq 1) {
					if ($Pkg.PackageUserInformation[0].Sid -eq 'S-1-5-18') {
						return @($Null, $False)
					}
					else {
						return @($Packages, $True)
					}
				}
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
	Remove-PackageFolder -FolderName $AppPath

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
	else {
		Format-Output '-- App folder not found'
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
		$PShell = 'C:\Program Files\PowerShell\7\pwsh.exe'
		$ArgString = "{ Remove-AppxPackage -Package $($PackageName) -User $($UserSid) -AllUsers }"
		if ($All -eq $False) {
			$ArgString = "{ Remove-AppxPackage -Package $($PackageName) -User $($UserSid) }"
		}
		$ArgArray = @(
			"-Command",
			$ArgString
		)
		Format-Output -Text "-- Remove-AppxPackage failed. Running with pwsh.exe"
		Start-Process -FilePath $PShell -ArgumentList $ArgArray -Wait
	}
}

function Remove-Package {
	param (
		[string]$Name,
		[bool]$All = $True,
		[string]$Service = ""
	)
	if ($Service -ne "")
	{
		Stop-Service $Service -ErrorAction SilentlyContinue
	}

	$PackageResult = Get-PackageIsInstalled -Name $Name
	if ($Null -eq $PackageResult) {
		Format-Output '-- Skipped (Not found)'
		Remove-AllFolders -Name $Name
		return @(1, 0)
	}
	$Packages = $PackageResult[0]
	$IsInstalled = $PackageResult[1]

	if (($Null -eq $Packages) -or ($IsInstalled -eq $False)) {
		Format-Output '-- Skipped (System User)'
		Remove-AllFolders -Name $Name
		return @(1, 0)
	}

	foreach ($Pkg in $Packages) {
		foreach ($Sid in $Pkg.PackageUserInformation) {
			$UserSid = $Sid.UserSecurityId.Sid
			if ($UserSid -eq 'S-1-5-18') { continue }
			if ($All) {			
				Invoke-RemoveAppxPackage -Package $Pkg -All $True
			}
			else {
				Invoke-RemoveAppxPackage -Package $Pkg -User $UserSid -All $False
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

	if (($Null -eq $Packages) -or ($IsInstalled -eq $False)) {
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
	Add-RegistryKey -Path $RegPath -Name $RegKeyName -Value $RegKeyValue -KeyType DWord

	# the registry values
	$RegPath = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Device Installer'
	$RegKeyName = 'DisableCoInstallers'
	$RegKeyValue = 1
	# add the key
	Add-RegistryKey -Path $RegPath -Name $RegKeyName -Value $RegKeyValue -KeyType DWord
}

try {
	$SkipCount = 0
	$UninstallCount = 0

	Format-Output "Connected"
	
	Format-Output "Uninstalling AMD Radeon Software"
	$Result = Remove-Package -Name "AdvancedMicroDevicesInc-2.AMDRadeonSoftware"
	if ($Null -ne $Result) {
		$SkipCount += $Result[0]
		$UninstallCount += $Result[1]
	}

	Format-Output "Uninstalling DuckDuckGo"
	$Result = Remove-Package -Name "DuckDuckGo"
	if ($Null -ne $Result) {
		$SkipCount += $Result[0]
		$UninstallCount += $Result[1]
	}

	Format-Output "Uninstalling Waves Audio (Each User)"
	$Result = Remove-Package -Name "WavesAudio" -All $False -Service 'Waves Audio Services'
	if ($Null -ne $Result) {
		$SkipCount += $Result[0]
		$UninstallCount += $Result[1]
	}

	Format-Output "Uninstalling Waves Audio (All Users)"
	$Result = Remove-Package -Name "WavesAudio" -Service 'Waves Audio Services'
	if ($Null -ne $Result) {
		$SkipCount += $Result[0]
		$UninstallCount += $Result[1]
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
