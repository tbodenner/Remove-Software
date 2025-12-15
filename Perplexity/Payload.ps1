# host name for the computer this script is running on
$ComputerName = $env:computername
# get the user who is running this script
$RunningUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name.Split('\')[1]
# get the running user's profile folder acl
$AdminAcl = Get-Acl -Path "C:\Users\$($RunningUser)"

try {
	Write-Host "$($ComputerName): Connected"
	
	# array of processes to stop
	$ProcessArray = @(
		'Perplexity',
		'Comet'
	)
	
	# our list of folders to remove
	$PathArray = @(
		"AppData\Local\Programs\Perplexity\*",
		"AppData\Local\Perplexity\*",
		"Downloads\Perplexity*Setup*.exe",
		"AppData\Local\Programs\Loom\*",
		"AppData\Local\loom-updater\pending"
	)
	# this bool is set if a process was stopped
	$ProcessWasStopped = $false

	# get user folders
	$UserFolders = Get-ChildItem -Path 'C:\Users\' -Directory

	# loop through our processes
	foreach ($ProcessName in $ProcessArray) {
		# check if the process is running
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
	if ($ProcessWasStopped -eq $true) {
		# wait a second for processes to end
		Start-Sleep -Seconds 1
	}

	# loop through each user's profile folder
	foreach ($User in $UserFolders) {
		# create full user folder path
		$UserPath = Join-Path -Path 'C:\Users\' -ChildPath $User

		# loop through the path array
		foreach ($PartialPath in $PathArray) {
			# check if the path starts in the root of c:
			if (($PartialPath.Substring(0, 3) -ne "C:\")) {
				# combine the two paths
				$PartialJoinPath = Join-Path -Path $UserPath -ChildPath $PartialPath
				# resolve the paths
				$ResolvedPaths = Resolve-Path -Path $PartialJoinPath -ErrorAction SilentlyContinue
			}
			else {
				# path is not in user's folder, don't join it
				$PartialJoinPath = $PartialPath
				# resolve the paths
				$ResolvedPaths = Resolve-Path -Path $PartialJoinPath -ErrorAction SilentlyContinue
			}

			# check if our partial path includes a wildcard character
			if (($PartialPath.Substring($PartialPath.Length - 2, 2)) -eq "\*") {
				# create our root folder path by removing the asterisk
				$RootPath = $PartialPath.Substring(0, $PartialPath.Length - 1)
				$RootJoinPath = Join-Path -Path $UserPath -ChildPath $RootPath
				# check if this is a folder and it exists
				if ((Test-Path -Path $RootJoinPath -PathType Container) -eq $True) {
					# check if our acl is not null
					if ($null -ne $AdminAcl) {
						# set admin acl on root folder
						Set-Acl -Path $RootJoinPath -AclObject $AdminAcl -ErrorAction SilentlyContinue
					}
				}
			}

			# loop through our resolved paths
			foreach ($ResolvedPath in $ResolvedPaths) {
				# check if this is a folder and it exists
				if ((Test-Path -Path $ResolvedPath -PathType Container) -eq $True) {
					# if the path exists, then remove the folder
					Remove-Item -Path $ResolvedPath -Recurse -Force
					Write-Host "$($ComputerName): -- Removed folder '$(Split-Path -Path $ResolvedPath -Leaf)' for '$($User)'"
				}
				# check if this is a file and it exists
				if ((Test-Path -Path $ResolvedPath -PathType Leaf) -eq $True) {
					# if the path exists, then remove the file
					Remove-Item -Path $ResolvedPath -Force
					Write-Host "$($ComputerName): -- Removed file '$(Split-Path -Path $ResolvedPath -Leaf)' for '$($User)'"
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
