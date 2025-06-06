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
					Format-Output "-- Stopped process '$($Proc)'"
				}
				catch {
					Format-Output "-- Failed to stop process '$($Proc)'"
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
					Format-Output "-- Stopped service '$($Serv)'"
					sc.exe delete $Serv | Out-Null
					Format-Output "-- Removed service '$($Serv)'"
				}
				catch {
					Format-Output "-- Failed to stop or remove service '$($Serv)'"
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
						Format-Output "-- Failed to remove folder '$(Split-Path -Path $Path -Leaf)'"
					}
					else {
						Format-Output "-- Failed to remove folder '$($Path)'"
					}
					if ($Path -eq "C:\Program Files\AMD\") {
						Remove-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}\0000\' -Name 'DalDCELogFilePath' -ErrorAction SilentlyContinue
						Remove-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}\0002\' -Name 'DalDCELogFilePath' -ErrorAction SilentlyContinue
						Format-Output "-- Removed 'DalDCELogFilePath' registry value"
					}
				}
				else {
					if ($Path.Length -gt $MaxPathLen) {
						Format-Output "-- Removed folder '$(Split-Path -Path $Path -Leaf)'"
					}
					else {
						Format-Output "-- Removed folder '$($Path)'"
					}
				}
			}
			Remove-Item -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\AMD Catalyst Install Manager' -ErrorAction SilentlyContinue
			Set-DisableAppsForDevices
		}
		return @(0, 1)
	}
	catch {
		return @(0, 0)
	}
}

try {
	$SkipCount = 0
	$UninstallCount = 0

	Format-Output "Connected"
	
	Format-Output "Removing AMD Radeon Software Folders"
	$Result = Remove-AmdFolders
	if ($Null -ne $Result) {
		$SkipCount += $Result[0]
		$UninstallCount += $Result[1]
	}

	Format-Output "Running Hardware Inventory Cycle"
	Invoke-WmiMethod -Namespace 'root\ccm' -Class 'sms_client' -Name 'TriggerSchedule' -ArgumentList '{00000000-0000-0000-0000-000000000001}'

	Format-Output "Done"

	return @($SkipCount, $UninstallCount)
}
catch {
	Format-Output "-- Error caught in script. Check error file."
	Write-Error $_
	return @(0, 0)
}
