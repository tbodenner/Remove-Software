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
	$ProductName = "Citrix Screen Casting for Windows"
	$UninstallVersionList = @("19.11.100.48", "19.12.4000.19")

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
				cmd /c $cmd
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
	
	Format-Output "Done"
}
catch {
	Format-Output "Error in script"
	Format-Output $_
	Write-Error $_
}

# return an array of our counts
return @($SkipCount, $UninstallCount)