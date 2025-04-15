#Requires -RunAsAdministrator

$ComputerName = $env:computername

function Format-Output {
	param ([string]$Text)
	$Output = "[--|{0}| {1}" -f $ComputerName, $Text
	Write-Host $Output
}

function Get-CmInum {
	# software package we are looking for
	$CmName = 'CM'
	# x86
	$Path32 = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
	# amd64
	$Path64 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
	# get all installed software
	$Software = Get-ItemProperty -Path $Path32, $Path64
	# look for our software in our software list
	$Inum = ($Software | Where-Object { $_.DisplayName -eq $CmName }).PSChildName
	# return the inum to be removed by msiexec
	return $Inum
}

try {
	Format-Output "Connected"
	
	# get user folders
	$UserFolders = Get-ChildItem -Path C:\Users\ -Directory

	foreach ($User in $UserFolders) {
		# create full user folder path
		$UserPath = Join-Path -Path 'C:\Users\' -ChildPath $User

		# combine user folder and zoom folder
		$ZoomRoamingPath = Join-Path -Path $UserPath -ChildPath 'AppData\Roaming\Zoom'
		# check if the zoom path exists
		if ((Test-Path -Path $ZoomRoamingPath -PathType Container) -eq $True) {
			# if the path exists, then remove it
			Remove-Item -Path $ZoomRoamingPath -Recurse -Force
			Format-Output "-- Removed roaming Zoom folder '$($User)'"
		}

		# combine user folder and zoom folder
		$ZoomLocalPath = Join-Path -Path $UserPath -ChildPath 'AppData\Local\Zoom'
		# check if the zoom path exists
		if ((Test-Path -Path $ZoomLocalPath -PathType Container) -eq $True) {
			# if the path exists, then remove it
			Remove-Item -Path $ZoomLocalPath -Recurse -Force
			Format-Output "-- Removed local Zoom folder '$($User)'"
		}

		# create full user start menu folder path
		$ZoomStartPath = Join-Path -Path $UserPath -ChildPath 'AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Zoom'
		# check if the zoom start menu path exists
		if ((Test-Path -Path $ZoomStartPath -PathType Container) -eq $True) {
			# if the path exists, then remove it
			Remove-Item -Path $ZoomStartPath -Recurse -Force
			Format-Output "-- Removed start menu Zoom folder '$($User)'"
		}

		# combine user folder and downloaded zoom file
		$ZoomFile = Join-Path -Path $UserPath -ChildPath 'Downloads\Zoom_cm*'
		# check if the downloads file exists
		if ((Test-Path -Path $ZoomFile -PathType Leaf) -eq $True) {
			# if the file exists, then remove it
			Remove-Item -Path $ZoomFile -Force
			Format-Output "-- Removed Zoom download file '$($User)'"
		}

		# combine system task folder with zoom task name
		$ZoomTask = Join-Path -Path 'C:\Windows\System32\Tasks' -ChildPath 'ZoomUpdateTask*'
		# check if the task file exists
		if ((Test-Path -Path $ZoomTask -PathType Leaf) -eq $True) {
			# if the file exists, then remove it
			Remove-Item -Path $ZoomTask -Force
			Format-Output "-- Removed Zoom task file"
		}

		# remove cm software package
		$CmINum = Get-CmInum
		# if we got an inum, try to uninstall the package
		if ($Null -ne $CmINum) {
			# uninstall the package
			cmd.exe /c "msiexec.exe /x `"$($CmINum)`" /quiet /norestart"
			# get the inum again
			$CmINum = Get-CmInum
			# check if the inum is null
			if ($Null -eq $CmINum) {
				# the software was removed
				Format-Output "-- Removed CM package"
			}
			else {
				# otherwise, the software was not removed
				Format-Output "-- Failed to removed CM package"
			}
		}

		# delete a folder
		$CompanyPath = 'C:\Program Files (x86)\My Company Name\My Product Name'
		# check if the folder exists
		if ((Test-Path -Path $CompanyPath -PathType Container) -eq $True) {
			# if the folder exists, then remove it
			Remove-Item -Path $CompanyPath -Force -Recurse
			Format-Output "-- Removed 'My Company Name' folder"
		}
	}
	
	# trigger sccm scan
	Format-Output "Running Hardware Inventory Cycle"
	Invoke-WmiMethod -Namespace 'root\ccm' -Class 'sms_client' -Name 'TriggerSchedule' -ArgumentList '{00000000-0000-0000-0000-000000000001}'
	
	Format-Output "Done"
}
catch {
	Format-Output "Error in script"
	Format-Output $_
	Write-Error $_
}
