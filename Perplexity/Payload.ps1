#Requires -RunAsAdministrator

$ComputerName = $env:computername

try {
	Write-Host "$($ComputerName): Connected"
	
	# array of processes to stop
	$ProcessArray = @(
		'Perplexity',
		'Comet'
	)

	# this bool is set if a process was stopped
	$ProcessWasStopped = $false

	# loop through our processes
	foreach ($ProcessName in $ProcessArray) {
		# check if zoom is running
		$Process = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
		# if an executable is found
		if ($null -ne $Process) {
			# stop the process
			Stop-Process $Process -Force -ErrorAction SilentlyContinue
			Write-Host "$($ComputerName): -- Stopping process '$($ProcessName)'"
			# set our bool
			$ProcessWasStopped = $true
		}
	}

	# check if we stopped any processes
	if ($ProcessWasStopped -eq $true {
		# wait a second for processes to end
		Start-Sleep -Seconds 1
	}

	# get user folders
	$UserFolders = Get-ChildItem -Path 'C:\Users\' -Directory

	# our list of folders to remove
	$FolderArray = @(
		"AppData\Local\Programs\Perplexity",
		"AppData\Local\Perplexity"
	)
	# our list of files to remove
	$FileArray = @(
		"Downloads\Perplexity*Setup*.exe"
	)

	foreach ($User in $UserFolders) {
		# create full user folder path
		$UserPath = Join-Path -Path 'C:\Users\' -ChildPath $User

		# loop through the folder array
		foreach ($Folder in $FolderArray) {
			# combine the two folders
			$FolderJoinPath = Join-Path -Path $UserPath -ChildPath $Folder
			# resolve the folder names
			$ResolvedFolderPaths = Resolve-Path -Path $FolderJoinPath -ErrorAction SilentlyContinue
			# loop through our resolved paths
			foreach ($FolderPath in $ResolvedFolderPaths) {
				# check if the path exists
				if ((Test-Path -Path $FolderPath -PathType Container) -eq $True) {
					# if the path exists, then remove the folder
					Remove-Item -Path $FolderPath -Recurse -Force
					Write-Host "$($ComputerName): -- Removed folder '$($User)': '$(Split-Path -Path $FolderPath -Leaf)'"
				}
			}
		}

		# loop through the file array
		foreach ($File in $FileArray) {
			# combine user folder and file
			$FileJoinPath = Join-Path -Path $UserPath -ChildPath $File
			# resolve the file names
			$ResolvedFilePaths = Resolve-Path -Path $FileJoinPath -ErrorAction SilentlyContinue
			# loop through our resolved files
			foreach ($FilePath in $ResolvedFilePaths) {
				# check if the file exists
				if ((Test-Path -Path $FilePath -PathType Leaf) -eq $True) {
					# if the file exists, then remove it
					Remove-Item -Path $FilePath -Force
					Write-Host "$($ComputerName): -- Removed file '$($User)': '$(Split-Path -Path $FilePath -Leaf)'"
				}
			}
		}

		<#
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
		#>
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
