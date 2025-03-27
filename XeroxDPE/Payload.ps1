#Requires -RunAsAdministrator

$ComputerName = $env:computername

function Format-Output {
	param ([string]$Text)
	$Output = "[--|{0}| {1}" -f $ComputerName, $Text
	Write-Host $Output
}

try {
	$ProductName = "Xerox Desktop Print Experience"
	$UninstallVersionList = @("7.*")
	
	Format-Output "Connected"

	foreach ($UninstallCurrentVersion in $UninstallVersionList) {
		Format-Output "Getting Identifying Number and Version"
		$Object = (Get-WmiObject -Class Win32_Product | Where-Object {$_.Name -match $ProductName -and $_.Version -like $UninstallCurrentVersion})
		$INum = $Object.IdentifyingNumber
		[string]$Version = $Object.Version
		Format-Output ("--INum: {0}" -f $INum)
		Format-Output ("--Version: {0}" -f $Version)

		if ($Version -like $UninstallCurrentVersion) {
			if ($INum -ne "") {
				Format-Output ("Uninstalling '{0}'" -f $Version)
				$cmd = "msiexec.exe /x {0} /quiet /norestart" -f $INum
				cmd /c $cmd
			}

			Format-Output "Checking Installed Version"
			$Version = (Get-WmiObject -Class Win32_Product | Where-Object {$_.Name -match $ProductName -and $_.Version -like $UninstallCurrentVersion}).Version
			
			if ($Version -ne "") {
				$ErrString = "{0}: Failed To Uninstall '{1} v{2}'" -f $ComputerName, $ProductName, $UninstallCurrentVersion
				Format-Output $ErrString
				Write-Error $ErrString
			}
			else {
				Format-Output  ("--Version '{0}' Removed" -f $UninstallCurrentVersion)
			}
		}
		else {
			# nothing to do, correct version found
			Format-Output ("Skipping {0}, nothing to uninstall" -f $UninstallCurrentVersion)
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
