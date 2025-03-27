#Requires -RunAsAdministrator

$ComputerName = $env:computername
$SkipCount = 0
$UninstallCount = 0

function Format-Output {
	param ([string]$Text)
	$Output = "[--|$($ComputerName)| $($Text)"
	Write-Host $Output
}

try {
	$ProductName = "GoTo Opener"
	$UninstallVersionList = @("1.0.*")
	
	Format-Output "Connected"

	foreach ($UninstallCurrentVersion in $UninstallVersionList) {
		Format-Output "Getting Identifying Number and Version"
		$X64 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
		$X32 = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
		$Software = Get-ItemProperty -Path $X64, $X32
		$INum = ($Software | Where-Object { $_.DisplayName -like "$($ProductName)*" }).PSChildName
		$Version = ($Software | Where-Object { $_.DisplayName -like "$($ProductName)*" }).DisplayVersion
		Format-Output "--INum: $($INum)"
		Format-Output "--Version: '$($Version)'"

		if ($Version -like $UninstallCurrentVersion) {
			if ($INum -ne "") {
				Format-Output "Uninstalling '$($Version)'"
				$cmd = "msiexec.exe /x $($INum) /quiet /norestart"
				#Write-Host $cmd
				$Result = cmd /c $cmd
				$MsiErrorString = 'This action is only valid for products that are currently installed.'
				if ($Result -eq $MsiErrorString) {
					Format-Output "--Software not actually installed"
					# software is not installed, remove the registry value
					$Path64 = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$($INum)"
					$Path32 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$($INum)"
					# test if the registry key exists
					if (Test-Path -Path $Path64) {
						# remove the registry key
						Remove-Item -Path $Path64 -Recurse -Force
						Format-Output "--Removed x64 registry key"
					}
					# test if the registry key exists
					if (Test-Path -Path $Path32) {
						# remove the registry key
						Remove-Item -Path $Path32 -Recurse -Force
						Format-Output "--Removed x86 registry key"
					}
				}
			}

			Format-Output "Checking Installed Version"
			$Software = Get-ItemProperty -Path $X64, $X32
			$Version = ($Software | Where-Object { $_.DisplayName -like "$($ProductName)*" }).DisplayVersion
			if ($Null -ne $Version) {
				$Version = $Version.Trim()
			}
			else {
				$Version = ""
			}
			Format-Output "--Version: '$($Version)'"
			if ($Version -ne "") {
				$ErrString = "$($ComputerName): Failed To Uninstall '$($ProductName) v$($UninstallCurrentVersion)'"
				Format-Output $ErrString
				Write-Error $ErrString
			}
			else {
				# software was removed
				Format-Output  "--Version '$($UninstallCurrentVersion)' Removed"
				# update our count
				$UninstallCount += 1
			}
		}
		else {
			# nothing to do, correct version found
			Format-Output "Skipping $($UninstallCurrentVersion), nothing to uninstall"
			# update our count
			$SkipCount += 1
		}
	}

	Format-Output "Running Hardware Inventory Cycle"
	Invoke-WmiMethod -Namespace 'root\ccm' -Class 'sms_client' -Name 'TriggerSchedule' -ArgumentList '{00000000-0000-0000-0000-000000000001}'

	Format-Output "Done`n"
}
catch {
	Format-Output "Error in script"
	Format-Output $_
	Write-Error $_
}

# return an array of our counts
return @($SkipCount, $UninstallCount)