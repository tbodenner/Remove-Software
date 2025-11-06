#Requires -RunAsAdministrator

$ComputerName = $env:computername

try {
	Write-Host "$($ComputerName): Connected"
	
	# check if zoom is running
	$ZoomProcess = Get-Process -Name Zoom -ErrorAction SilentlyContinue
	# if a zoom executable is found
	if ($null -ne $ZoomProcess) {
		# get the UDP connections
		$ZoomUPD = Get-NetUDPEndpoint -OwningProcess $ZoomProcess.Id -ErrorAction SilentlyContinue
		# remove any local connections and return a count of data being sent
		$ConnCount = ($ZoomUPD | Where-Object {$_.LocalAddress -ne '127.0.0.1'} | Measure-Object).Count
		# if the count is non-zero, zoom is in a meeting
		if ($ConnCount -gt 0) {
			# write a message
			Write-Host "$($ComputerName): -- Active Zoom meeting detected!"
			# and exit
			return
		}
		else {
			# otherwise, stop the process and continue to remove the installed application
			Stop-Process $ZoomProcess -Force -ErrorAction SilentlyContinue
			
			# wait 2 seconds for the process to end
			Start-Sleep -Seconds 2
		}
	}
	# get user folders
	$UserFolders = Get-ChildItem -Path 'C:\Users\' -Directory

	foreach ($User in $UserFolders) {
		# create full user folder path
		$UserPath = Join-Path -Path 'C:\Users\' -ChildPath $User

		# combine user folder and zoom folder
		$ZoomRoamingPath = Join-Path -Path $UserPath -ChildPath 'AppData\Roaming\Zoom'
		# check if the zoom path exists
		if ((Test-Path -Path $ZoomRoamingPath -PathType Container) -eq $True) {
			# if the path exists, then remove it
			Remove-Item -Path $ZoomRoamingPath -Recurse -Force
			Write-Host "$($ComputerName): -- Removed roaming Zoom folder '$($User)'"
		}

		# combine user folder and zoom folder
		$ZoomLocalPath = Join-Path -Path $UserPath -ChildPath 'AppData\Local\Zoom'
		# check if the zoom path exists
		if ((Test-Path -Path $ZoomLocalPath -PathType Container) -eq $True) {
			# if the path exists, then remove it
			Remove-Item -Path $ZoomLocalPath -Recurse -Force
			Write-Host "$($ComputerName): -- Removed local Zoom folder '$($User)'"
		}

		# create full user start menu folder path
		$ZoomStartPath = Join-Path -Path $UserPath -ChildPath 'AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Zoom'
		# check if the zoom start menu path exists
		if ((Test-Path -Path $ZoomStartPath -PathType Container) -eq $True) {
			# if the path exists, then remove it
			Remove-Item -Path $ZoomStartPath -Recurse -Force
			Write-Host "$($ComputerName): -- Removed start menu Zoom folder '$($User)'"
		}

		# combine user folder and downloaded zoom file
		$ZoomFile = Join-Path -Path $UserPath -ChildPath 'Downloads\Zoom_cm*'
		# check if the downloads file exists
		if ((Test-Path -Path $ZoomFile -PathType Leaf) -eq $True) {
			# if the file exists, then remove it
			Remove-Item -Path $ZoomFile -Force
			Write-Host "$($ComputerName): -- Removed Zoom download file '$($User)'"
		}

		# get all files in user's temp folder
		$EdgeDownloadFolder = Join-Path -Path $UserPath -ChildPath '\AppData\Local\Temp\'
		# check if the folder exists
		if ((Test-Path -Path $EdgeDownloadFolder -PathType Container) -eq $True) {
			# get all files in the temp folder
			$TempFilePaths = (Get-ChildItem -Path $EdgeDownloadFolder -Recurse -Force).FullName
			# loop through all the files
			foreach ($TempFile in $TempFilePaths) {
				# check if the filename contains our the item we are looking for
				if ($TempFile.Contains("Zoom_cm")) {
					# check if this is a file
					if ((Test-Path -Path $TempFile -PathType Leaf) -eq $true) {
						# delete the file
						Remove-Item $TempFile -Force -ErrorAction SilentlyContinue
						Write-Host "$($ComputerName): -- Removed Zoom file '$($User)'"
						Write-Host "$($ComputerName): ---- File '$(Split-Path -Path $TempFile -Leaf)'"
					}
					# check if this is a folder
					if ((Test-Path -Path $TempFile -PathType Container) -eq $true) {
						# delete the folder
						Remove-Item $TempFile -Force -Recurse -ErrorAction SilentlyContinue
						Write-Host "$($ComputerName): -- Removed Zoom folder '$($User)'"
						Write-Host "$($ComputerName): ---- Folder '$(Split-Path -Path $TempFile -Leaf)'"
					}
				}
			}
		}
		# combine system task folder with zoom task name
		$ZoomTask = Join-Path -Path 'C:\Windows\System32\Tasks' -ChildPath 'ZoomUpdateTask*'
		# check if the task file exists
		if ((Test-Path -Path $ZoomTask -PathType Leaf) -eq $True) {
			# if the file exists, then remove it
			Remove-Item -Path $ZoomTask -Force
			Write-Host "$($ComputerName): -- Removed Zoom task file"
		}
	}
	
	# trigger sccm scan
	Write-Host "$($ComputerName): Triggering Hardware Inventory Cycle"
	Invoke-WmiMethod -Namespace 'root\ccm' -Class 'sms_client' -Name 'TriggerSchedule' -ArgumentList '{00000000-0000-0000-0000-000000000001}'
	
	Write-Host "$($ComputerName): Done"
}
catch {
	Write-Host "$($ComputerName): Error in script"
	Write-Host "$($ComputerName): $($_)"
	Write-Error $_
}
