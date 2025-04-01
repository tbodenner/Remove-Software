#Requires -RunAsAdministrator

$ComputerName = $env:computername

function Format-Output {
	param ([string]$Text)
	$Output = "[--|{0}| {1}" -f $ComputerName, $Text
	Write-Host $Output
}

try {
	Format-Output "Connected"

	$ResultArray = $Null

	$ProductName = "ActivID ActivClient x64"
	$ProductVersion = "7.4.0"
	$ProductGuid = "{0AE0A544-7175-4CAB-97F0-922E90D07D5F}"
	$RegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$($ProductGuid)"

	# check if the uninstall string exists
	Format-Output "Checking for '$($ProductName)' v$($ProductVersion)"
	if ((Test-Path -Path $RegPath) -eq $True) {
		# get display name
		$DisplayName = (Get-ItemProperty -Path $RegPath -Name "DisplayName").DisplayName
		# get version
		$DisplayVersion = (Get-ItemProperty -Path $RegPath -Name "DisplayVersion").DisplayVersion

		# check our name and version
		if (($ProductName -eq $DisplayName) -and ($ProductVersion -eq $DisplayVersion)) {
			# remove the software
			Format-Output "Uninstalling '$($DisplayName)' v$($DisplayVersion)"
			$cmd = "msiexec.exe /x $($ProductGuid) /quiet /norestart"
			cmd /c $cmd

			# check if our uninstall string was removed
			if ((Test-Path -Path $RegPath) -eq $True) {
				# failed to remove the software
				Format-Output "-- Failed to uninstall '$($DisplayName)' v$($DisplayVersion)"
				# fail result
				$ResultArray = @(0, 0)
			}
			else {
				# software was removed
				Format-Output "-- Uninstalled '$($DisplayName)' v$($DisplayVersion)"
				# success result
				$ResultArray = @(0, 1)
			}
		}
		else {
			# wrong software found
			Format-Output "-- Software name and version mismatch"
			# fail result
			$ResultArray = @(0, 0)
		}
	}
	else {
		# software not installed
		Format-Output "-- '$($ProductName)' v$($ProductVersion) not found"
		# success result
		$ResultArray = @(0, 1)
	}

	Format-Output "Running Hardware Inventory Cycle"
	Invoke-WmiMethod -Namespace 'root\ccm' -Class 'sms_client' -Name 'TriggerSchedule' -ArgumentList '{00000000-0000-0000-0000-000000000001}'
	
	Format-Output "Done`n"

	# return the result
	$ResultArray
}
catch {
	Format-Output "Error in script"
	Format-Output $_
	Write-Error $_
	return @(0, 0)
}
