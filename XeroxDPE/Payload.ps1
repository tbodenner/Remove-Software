#Requires -RunAsAdministrator

$ComputerName = $env:computername
$SkipCount = 0
$UninstallCount = 0

try {
	$ProductName = "Xerox Desktop Print Experience"
	$UninstallVersionList = @("7.*")

	Write-Host "$($ComputerName): Connected"

	foreach ($UninstallCurrentVersion in $UninstallVersionList) {
		Write-Host "$($ComputerName): Getting Identifying Number and Version"
		$X64 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
		$X32 = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
		$Software = Get-ItemProperty -Path $X64, $X32
		$INum = ($Software | Where-Object { $_.DisplayName -like "$($ProductName)*" }).PSChildName
		$Version = ($Software | Where-Object { $_.DisplayName -like "$($ProductName)*" }).DisplayVersion
		Write-Host "$($ComputerName): --INum: $($INum)"
		Write-Host "$($ComputerName): --Version: '$($Version)'"

		if ($Version -like $UninstallCurrentVersion) {
			if ($INum -ne "") {
				Write-Host "$($ComputerName): Uninstalling '$($Version)'"
				$cmd = "msiexec.exe /x $($INum) /quiet /norestart"
				#Write-Host $cmd
				cmd /c $cmd
			}

			Write-Host "$($ComputerName): Checking Installed Version"
			$Software = Get-ItemProperty -Path $X64, $X32
			$Version = ($Software | Where-Object { $_.DisplayName -like "$($ProductName)*" }).DisplayVersion
			if ($Null -ne $Version) {
				$Version = $Version.Trim()
			}
			else {
				$Version = ""
			}
			Write-Host "$($ComputerName): --Version: '$($Version)'"
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
			Write-Host "$($ComputerName): Skipping $($UninstallCurrentVersion), nothing to uninstall"
			# update our count
			$SkipCount += 1
		}
	}

	Write-Host "$($ComputerName): Triggering Hardware Inventory Cycle"
	Invoke-WmiMethod -Namespace 'root\ccm' -Class 'sms_client' -Name 'TriggerSchedule' -ArgumentList '{00000000-0000-0000-0000-000000000001}'
	
	Write-Host "$($ComputerName): Done"
}
catch {
	Write-Host "$($ComputerName): Error in script"
	Format-Output $_
	Write-Error $_
}

# return an array of our counts
return @($SkipCount, $UninstallCount)