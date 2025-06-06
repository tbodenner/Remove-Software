#Requires -RunAsAdministrator

$ComputerName = $env:computername
$Global:UninstallStrings = $Null

$HKLMPath32 = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
$HKLMPath64 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"

function Format-Output {
	param ([string]$Text)
	$Output = "[--|{0}| {1}" -f $ComputerName, $Text
	Write-Host $Output
}

function Update-UninstallStrings {
	$Global:UninstallStrings = (Get-ItemProperty -Path "$($HKLMPath32)*", "$($HKLMPath64)*" | Select-Object -Property DisplayName,DisplayVersion,PSChildName)
	return $Global:UninstallStrings
}

function Get-UninstallStringData {
	param (
		[string]$SoftwareName,
		[string]$VersionNum,
		[PSCustomObject]$InputObject
	)

	if ($Null -eq $Global:UninstallStrings) {
		Update-UninstallStrings
	}

	$SingleString = $Global:UninstallStrings | Where-Object { $_.DisplayName -eq $SoftwareName -and $_.DisplayVersion -eq $VersionNum }

	$InputObject.Name     = $SingleString.DisplayName
	$InputObject.Version  = $SingleString.DisplayVersion
	$InputObject.GUID     = $SingleString.PSChildName
}

function Remove-Path {
	param([string]$FullPath)
	# check if the path exists
	if ((Test-Path -Path $FullPath) -eq $True) {
		# remove the path
		Remove-Item -Path $FullPath -Force -Recurse -ErrorAction SilentlyContinue
		# pause before checking
		Start-Sleep -Milliseconds 100
		# check if the path still exists
		if ((Test-Path -Path $FullPath) -eq $True) {
			# if removed, return true
			return $True
		}
		else {
			# otherwise, return false
			return $False
		}
	}
	else {
		# path doesn't exist, return true
		return $True
	}
}

function Remove-FiSeriesManuals {
	param (
		[PSCustomObject]$InputObject
	)
	# save our values for later
	$SoftwareName = $InputObject.Name
	$SoftwareVersion = $InputObject.Version

	Format-Output "Removing $($SoftwareName) v$($SoftwareVersion)"

	# check if we got a string
	if ($Null -ne $InputObject.Name) {
		# delete program files folder
		$InstallPath = "C:\Program Files (x86)\fi Series manuals\"
		if ((Remove-Path -FullPath $InstallPath) -eq $True) {
			Format-Output "-- Program files removed"
		}
		else {
			Format-Output "-- Program files NOT removed"
		}
		# delete uninstall folder
		$SetupPath = "C:\Program Files (x86)\InstallShield Installation Information\$($InputObject.GUID)\"
		if ((Remove-Path -FullPath -Path $SetupPath) -eq $True) {
			Format-Output "-- Setup files removed"
		}
		else {
			Format-Output "-- Setup files NOT removed"
		}
		# remove 32 bit uninstall string
		$RegPath32 = "$($HKLMPath32)\$($InputObject.GUID)\"
		if ((Remove-Path -FullPath -Path $RegPath32) -eq $True) {
			Format-Output "-- HKLM x86 uninstall keys removed"
		}
		else {
			Format-Output "-- HKLM x86 uninstall keys NOT removed"
		}
		# remove 64 bit uninstall string
		$RegPath64 = "$($HKLMPath64)\$($InputObject.GUID)\"
		if ((Remove-Path -FullPath -Path $RegPath64) -eq $True) {
			Format-Output "-- HKLM x64 uninstall keys removed"
		}
		else {
			Format-Output "-- HKLM x64 uninstall keys NOT removed"
		}
	}
	else {
		# otherwise the software is not installed
		Format-Output "-- $($SoftwareName) $($SoftwareVersion) not found"
	}
}

function Uninstall-Msi {
	param (
		[PSCustomObject]$InputObject
	)
	# save our values for later
	$SoftwareName = $InputObject.Name
	$SoftwareVersion = $InputObject.Version

	Format-Output "Removing $($InputObject.Name) v$($InputObject.Version)"
	# check if we got a value
	if ($Null -ne $InputObject.Name) {
		# if we got a value, try to remove it
		cmd.exe /c "msiexec.exe /x `"$($InputObject.GUID)`" /quiet /norestart"
		# update our uninstall strings
		Update-UninstallStrings
		# get our uninsall string data again
		Get-UninstallStringData -InputObject $InputObject
		# check if we removed the software and output a message
		if ($Null -eq $InputObject.Name -and $Null -eq $InputObject.Version) {
			Format-Output "-- $($SoftwareName) v$($SoftwareVersion) removed"
		}
		else {
			Format-Output "-- $($SoftwareName) v$($SoftwareVersion) NOT removed"
		}
	}
	else {
		# no uninstall string, software was not found
		Format-Output "-- $($SoftwareName) v$($SoftwareVersion) not found"
	}
}

try {
	$Products = @(
		[Tuple]::Create(1, "Fujitsu ScandAll PRO x64", "1.00.00005")
		[Tuple]::Create(2, "Fujitsu ScandAll PRO x64", "1.00.00003")
		[Tuple]::Create(3, "Fujitsu ScandAll PRO API", "1.00.00001")
		[Tuple]::Create(4, "fi Series manuals for fi-7160/7260/7180/7280", "1.04.01")
	)

	Format-Output "Connected"

	foreach ($Product in $Products) {
		$TupleId = $Product.Item1
		$ProductName = $Product.Item2
		$UninstallVersion = $Product.Item3

		$SoftwareObject = [PSCustomObject]@{
			Name     = $Null
			Version  = $Null
			GUID     = $Null
		}

		Get-UninstallStringData -SoftwareName $ProductName -VersionNum $UninstallVersion -InputObject $SoftwareObject

		if ($Null -eq $SoftwareObject.Name) {
			continue
		}

		switch ($TupleId) {
			{ $_ -in @(1,2,3) } {
				Uninstall-Msi -InputObject $SoftwareObject
			}
			4 {
				Remove-FiSeriesManuals -InputObject $SoftwareObject
			}
			Default {
				Format-Output "-- No switch found for this tuple id '$($TupleId)'"
			}
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
