#Requires -RunAsAdministrator

$ComputerName = $env:computername

function Format-Output {
	param ([string]$Text)
	$Output = "[--|{0}| {1}" -f $ComputerName, $Text
    Write-Host $Output
}

function Set-DisableAppsForDevices {
	# the registry values
	$RegPath = 'HKCU:\Software\Policies\Microsoft\Windows\DeviceInstall'
	$RegKeyName = 'AllowOSManagedDriverInstallationToUI'
	$RegKeyValue = 0
	# check if path exists
	if ((Test-Path -Path $RegPath -Type Container) -eq $False) {
		# if the path is missing, then create it
		New-Item -Path $RegPath | Out-Null
	}
	# set the registry value and create the key if missing
	Set-ItemProperty -Path $RegPath -Name $RegKeyName -Value $RegKeyValue | Out-Null
}

function Remove-AmdFolders {
	$ProcArray = @(
		"RadeonSoftware",
		"AMDRSServ",
		"AMDRSSrcExt"
	)
	foreach ($Proc in $ProcArray) {
		if ($Null -ne (Get-Process "$($Proc)*")) {
			try {
				Stop-Process -Name $Proc -Force
				Format-Output "-- Stopped process '$($Proc)'"
			}
			catch {
				Format-Output "-- Failed to stop process '$($Proc)'"
			}
		}
	}

	$AmdPaths = @(
		"C:\Program Files\WindowsApps\AdvancedMicroDevicesInc-2.AMDRadeonSoftware_10.22.20073.0_x64__0a9344xs7nr4m\",
		"C:\Program Files\AMD\"
	)
	foreach ($Path in $AmdPaths) {
		if ((Test-Path -Path $Path) -eq $True) {
			Remove-Item -Path $Path -Recurse -Force -ErrorAction SilentlyContinue
			if ((Test-Path -Path $Path) -eq $True) {
				Format-Output "-- Failed to remove folder '$($Path)'"
				if ($Path -eq "C:\Program Files\AMD\") {
					Remove-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}\0000\' -Name 'DalDCELogFilePath'
					Format-Output "-- Removed 'DalDCELogFilePath' registry value"
				}
			}
			else {
				Format-Output "-- Removed folder '$($Path)'"
			}
		}
		Set-DisableAppsForDevices
	}
}

try {
	$SkipCount = 0
	$UninstallCount = 0

	Format-Output "Connected"
	
	Format-Output "Removing AMD Radeon Software Folders"
	Remove-AmdFolders

	Format-Output "Running Hardware Inventory Cycle"
	Invoke-WmiMethod -Namespace 'root\ccm' -Class 'sms_client' -Name 'TriggerSchedule' -ArgumentList '{00000000-0000-0000-0000-000000000001}'

	Format-Output "Done`n"

	return @($SkipCount, $UninstallCount)
}
catch {
	Format-Output "-- Error caught in script. Check error file."
	Write-Error $_
}
