#region PARAMETERS

param (
	# the folder that contains the payload and computer list
	[Parameter(Mandatory=$true)][string[]]$InputPath,
	# if enabled, remote sessions will be started with powershell v7
	[switch]$PowerShell7
)

#endregion PARAMETERS
#region IMPORTS

# check for our required module
$RequiredModuleName = 'LDAP-ADTools'
if($null -eq (Get-Module -Name $RequiredModuleName -ListAvailable)) {
    # the module was not found, write a message and exit
    Write-Host "Required module '$($RequiredModuleName)' not found" -ForegroundColor Red
    Write-Host "`nModule can be found at https://github.com/tbodenner/PowerShell-Modules" -ForegroundColor Red
    Write-Host "Update the PSModulePath: `$env:PSModulePath += `";<path to module>`"" -ForegroundColor Red
    Write-Host "`nExiting" -ForegroundColor Red
    exit
}
# import our module only for this script
Import-Module $RequiredModuleName -Scope Local

#endregion IMPORTS
#region PROPERTIES

# the name of our config json file
$Global:JsonConfigFileName = 'config.json'

# these properties are loaded from the config file
$Global:ConfigRootOU = $null
$Global:ConfigFilter = ""

#endregion PARAMETERS
#region FUNCTIONS

function Get-ConfigFromJson {
	# check if our config file exists
	if ((Test-Path -Path $Global:JsonConfigFileName) -eq $false)
	{
		# if we are unable to read the file, write an error message and exit
		Write-Host "ERROR: JSON config file '$($Global:JsonConfigFileName)' not found." -ForegroundColor Red
		exit
	}
	# read our data from the json config file
	$JsonConfigHashtable = Get-Content $Global:JsonConfigFileName | ConvertFrom-Json -AsHashtable

	# check if our hashtable is null
	if ($null -eq $JsonConfigHashtable) {
		# hashtable is null, exit
		Write-Host "ERROR: No config data read from JSON file '$($Global:JsonConfigFileName)'." -ForegroundColor Red
		exit
	}
	# set our variable from the config file
	$Global:ConfigRootOU = Test-ConfigValueNullOrEmpty -Hashtable $JsonConfigHashtable -Key "RootOU"
	$Global:ConfigFilter = Test-ConfigValueNullOrEmpty -Hashtable $JsonConfigHashtable -Key "Filter"
	# write our config values to the host
	Write-Host "Values read from $($Global:JsonConfigFileName):" -ForegroundColor DarkCyan
	Write-Host "  RootOU: '$($Global:ConfigRootOU -join "'`n          '")'" -ForegroundColor Cyan
	Write-Host "  Filter: '$($Global:ConfigFilter)'" -ForegroundColor Cyan
}

# check if a config value is null
function Test-ConfigValueNullOrEmpty {
	param (
		[Parameter(Mandatory=$true)][hashtable]$Hashtable,
		[Parameter(Mandatory=$true)][string]$Key
	)
	# get our value
	$Value = $Hashtable[$Key]

	# check if the value is null
	if (($null -eq $Value) -or ($Value -eq "")) {
		# if the value is null, write an error message and exit
		Write-Host "ERROR: Config '$($Key)' value is null or empty." -ForegroundColor Red
		exit
	}
	# otherwise, return the value
	$Value
}

# read computer list text file and return a unique sorted array
function Get-UniqueArrayFromFile {
	param (
		[Parameter(Mandatory=$true)][string]$FilePath
	)
	# get the list of computers from a text file
	$ComputerList = Get-Content -Path $FilePath
	# hashtable to use as a set
	$ComputerSet = @{}
	# add each computer to the hashtable
	foreach ($Computer in $ComputerList) {
		# trim our computer name
		$Computer = $Computer.Trim()
		# skip empty lines
		if (($null -eq $Computer) -or ($Computer -eq "")) { continue }
		# add the computer to the hashtable, duplicates will not be added
		$ComputerSet[$Computer] = $null
	}
	# return a sorted array of the hashtable keys
	return $ComputerSet.Keys | Sort-Object
}

# test if a computer is in AD
function Test-ComputerInAD {
	# parameters
	param (
		[Parameter(Mandatory=$true)][hashtable]$ADHashtable,
		[Parameter(Mandatory=$true)][string]$ComputerName
	)
	# check if our AD array is null
	if ($null -eq $ADHashtable) {
		# computer array is null, return false
		return $false
	}
	# check if the computer is in our hashtable
	if ($ADHashtable.Keys -match $ComputerName) {
		# if the computer is in our array, return true
		return $true
	}
	# if not found, return false
	return $false
}

# check if a computer can be found
function Find-Computer {
	# parameters
	param (
		[Parameter(Mandatory=$true)][hashtable]$ADHashtable,
		[Parameter(Mandatory=$true)][string]$ComputerName
	)
	# check if the computer is in AD
	if ((Test-ComputerInAd -ADHashtable $ADHashtable -ComputerName $ComputerName) -eq $false) {
		return @(-1, $NotInAD)
	}

	# ping the computer and save the details
	$ComputerDetails = Test-Connection -TargetName $ComputerName -Count 1 -TimeoutSeconds 3 -ErrorAction Ignore
	# check if the computer was found
	if ($null -eq $ComputerDetails) {
		# return our result
		return @(-1, $NotFound)
	}
	# get the pinged computer's ip
	$Ip = $ComputerDetails.Address
	# get the ping result
	$Latency = $ComputerDetails.Latency
	# get the status of the ping
	$Status = $ComputerDetails.Status
	# check if our status is null
	if ($null -eq $Status) {
		# if null, set our status
		$Status = $NoStatus
	}
	# check if the ping timed out
	if (($Latency -lt 0) -or ($Status -eq $NoStatus) -or ($Status -in @('TimedOut', 'DestinationHostUnreachable'))) {
		# the ping timed out, so return the result of our ping
		return @(-1, $Status)
	}
	# set our error value for our dns name
	$DnsName = $null
	# check if we got an ip
	if ($null -ne $Ip) {
		# get our dns data
		$DnsData = Resolve-DnsName -Name $Ip -ErrorAction Ignore
		# check if we got any data
		if ($null -eq $DnsData) {
			return @(-1, 'DnsNotFound')
		}
		# get our host name from the data
		$DnsName = $DnsData.NameHost
	}
	# check if we got a hostname
	if ($null -eq $DnsName) {
		# if true, write a message
		Write-Host "$($ComputerName) Unable to resolve computer name from IP address" -ForegroundColor Red
		# return our result
		return @(-1, 'NoHostName')
	}
	# otherwise, check if our computer name matches the dns name
	if ($DnsName -ne ($ADHashtable[$ComputerName]['dnshostname'])) {
		Write-Host "$($ComputerName) DNS mismatch (DNS: $($DnsName), CN: $($ADHashtable[$ComputerName]['dnshostname']))" -ForegroundColor Red
		# return the dns error
		return @($Latency, $DnsMismatch)
	}
	# return the result of our ping
	return @($Latency, $Status)
}

#endregion FUNCTIONS
#region MAIN

# if our status is empty, use this status
$NoStatus = 'NoStatus'
# if our ping fails to find a target, use this status
$NotFound = 'NotFound'
# if we have a dns mismatch, use this status
$DnsMismatch = 'DnsMismatch'
# if our computer is not in AD, use this status
$NotInAd = 'NotInAd'

# our progress activity
$Activity = "Progress"

# do the work for each input path
foreach ($IPath in $InputPath) {

	# test if our input folder exists
	if ((Test-Path -Path $IPath) -eq $false) {
		Write-Host "Error: Input folder not found." -ForegroundColor Red
		return
	}

	# log folder
	$LogFolderName = 'Logs'
	$LogFolder = Join-Path -Path $IPath -ChildPath $LogFolderName

	# test if our log folder exists
	if ((Test-Path -Path $LogFolder) -eq $false) {
		New-Item -Path $LogFolder -ItemType "directory" | Out-Null
	}

	# old log folder
	$OldLogFolderName = 'Old'
	$OldLogFolder = Join-Path -Path $LogFolder -ChildPath $OldLogFolderName

	# test if our old log folder exists
	if ((Test-Path -Path $OldLogFolder) -eq $false) {
		New-Item -Path $OldLogFolder -ItemType "directory" | Out-Null
	}

	# file to output errors
	$ErrorFile = Join-Path -Path $LogFolder -ChildPath "Errors.txt"
	# file that contains our script
	$PayloadFile = Join-Path -Path $IPath -ChildPath "Payload.ps1"
	# file that contains our computer list
	$ComputerListFile = Join-Path -Path $IPath -ChildPath "ComputerList.txt"

	# test if our required payload file exists
	if ((Test-Path -Path $PayloadFile) -eq $false) {
		Write-Host "Error: Payload file not found." -ForegroundColor Red
		return
	}
	# test if our required computer list file exists
	if ((Test-Path -Path $ComputerListFile) -eq $false) {
		Write-Host "Error: Computer list file not found." -ForegroundColor Red
		return
	}

	# an array to collect the computer names with an error
	$ErrorArray = @()

	# an array to collect the computer names that succeeded
	$SuccessArray = @()

	# get the list of computers from a text file
	$ComputerList = Get-UniqueArrayFromFile -FilePath $ComputerListFile

	# check if our list contains any computers
	if ($ComputerList.Count -le 0) {
		# write our message
		Write-Host "Computer list is empty. Skipping '$($PayloadFile)'." -ForegroundColor DarkCyan
		# skip this payload
		continue
	}

	# write our starting status
	$Plural = ""
	if ($ComputerList.Count -eq 1) {
		$Plural = "computer"
	}
	else {
		$Plural = "computers"
	}
	$StartString = "`nRunning '$($PayloadFile)' on $($ComputerList.Count) $($Plural)"
	Write-Host $StartString -ForegroundColor Green

	# clear the error list so we can write only our errors
	$Error.Clear()

	# read our config file
	Get-ConfigFromJson

	# count variables
	$TotalComputers = $ComputerList.Count
	$ComputerCount = 0
	$SkipCount = 0
	$UninstallCount = 0

	# percent complete
	$PComplete = 0.0

	# flush our dns
	Clear-DnsClientCache

	# change our default settings for our remote session used by invoke-command
	$PssOptions = New-PSSessionOption -MaxConnectionRetryCount 0 -OpenTimeout 30000 -OperationTimeout 30000

	# create an array for our root ous
	$Global:RootOU = @()
	# convert our canonical names to distinguished names
	foreach ($Root in $Global:ConfigRootOU) {
		# add the converted name to our array
		$Global:RootOU += (Convert-CanonicalToDistinguished -Path $Root)
	}

	# wite an update
	Write-Host "Getting computers from AD using LDAP" -ForegroundColor DarkCyan
	# get an array of AD computers
	$ADComputerHashtable = Get-LDAPComputer -RootOU $Global:RootOU -Computers $Global:ConfigFilter -Properties 'dnshostname'

	# loop through list of computers
	foreach ($Computer in $ComputerList) {
		try {
			# check of we have a computer name
			if (($null -eq $Computer) -or ($Computer -eq "")) { continue }
			# get the last error in the error variable
			$LastError = $Error[0]
			# update our progress
			$PComplete = ($ComputerCount / $TotalComputers) * 100
			$Status = "$ComputerCount/$TotalComputers Complete"
			Write-Progress -Activity $Activity -Status $Status -PercentComplete $PComplete
			Write-Host "`n$($Computer): Trying To Connect" -ForegroundColor Yellow
			# set our parameters for our invoke command
			$Parameters = @{
				ComputerName	= ($ADComputerHashtable[$Computer]['dnshostname'])
				FilePath		= $PayloadFile
				ErrorAction		= "SilentlyContinue"
				SessionOption	= $PssOptions
			}
			if ($PowerShell7 -eq $true) {
				$Parameters['ConfigurationName'] = "PowerShell.7"
			}
			# get our ping data
			$PingData = Find-Computer -ADHashtable $ADComputerHashtable -ComputerName $Computer
			# get the result (boolean)
			$PingLatency = $PingData[0]
			# get the ping status (string)
			$PingStatus = $PingData[1]
			# check if we can ping the computer
			if (($PingLatency -ge 0) -and ($PingStatus -ne 'TimedOut')) {
				# check if we did not have a dns mismatch
				if ($PingStatus -ne $DnsMismatch) {
					Write-Host "$($Computer) Ping ($($PingStatus))" -ForegroundColor Green
					# run the script on the target computer if we can ping the computer
					$InvokeReturn = Invoke-Command @Parameters
					# check if we got anything back from the invoke command
					if ($null -ne $InvokeReturn) {
						# should be @($SkipCount, $UninstallCount)
						# check if our skip count is an int
						if ($InvokeReturn[0] -is [int]) {
							# update our count
							$SkipCount += $InvokeReturn[0]
						}
						# check if our uninstall count is an int
						if ($InvokeReturn[1] -is [int]) {
							# update our count
							$UninstallCount += $InvokeReturn[1]
						}
					}
					else {
						# we got a null value from our invoke-command, add the computer to the error array
						$ErrorArray += $Computer
						# write the error message
						Write-Host "$($Computer) No return value from Payload script" -ForegroundColor Red
					}
				}
				else {
					# otherwise, write an error
					Write-Error -Message "$($Computer): DNS mismatch error" -Category ConnectionError -ErrorAction SilentlyContinue
				}
			}
			else {
				Write-Host "$($Computer) Ping ($($PingStatus))" -ForegroundColor Red
				# otherwise, write an error
				Write-Error -Message "$($Computer): Unable to ping $($Computer) - ($($PingStatus), $($PingLatency))" -Category ConnectionError -ErrorAction SilentlyContinue
			}
			# check if our computer is in AD before adding to either array
			if ($PingStatus -ne $NotInAd) {
				# determine if the last computer was a success or error
				if ($LastError -eq $Error[0]) {
					# if the script finished without adding an error, add the computer to our success array
					$SuccessArray += $Computer
				}
				else {
					Write-Host "Error: $($Computer)" -ForegroundColor Yellow
					# if the script added an error, add the computer to our success array
					$ErrorArray += $Computer
				}
			}
			# increment our count
			$ComputerCount += 1
		}
		catch {
			Write-Output "Caught Error"
			Write-Output "$($_)"
			Write-Host ($_ | Select-Object -Property *)
			# add the computer to our error array if an error was caught
			$ErrorArray += $Computer
		}
	}

	# complete our progress
	Write-Progress -Activity $Activity -Completed

	# create our counts array to output to the console and a file
	$CountsArray = @(
		"`nResults:"
		"    Total: $($TotalComputers)"
		"  Success: $($SuccessArray.Count)"
		"     Skip: $($SkipCount)"
		"Uninstall: $($UninstallCount)"
		"    Error: $($ErrorArray.Count)"
	)

	# write our counts
	Write-Host $CountsArray[0]
	Write-Host $CountsArray[1] -ForegroundColor Yellow
	Write-Host $CountsArray[2] -ForegroundColor Green
	Write-Host $CountsArray[3] -ForegroundColor Cyan
	Write-Host $CountsArray[4] -ForegroundColor Blue
	Write-Host $CountsArray[5] -ForegroundColor Red

	# write our output files
	Write-Host "`nWriting Output Files" -ForegroundColor Yellow

	# rename our computer list file so we can write a new one
	$DateString = Get-Date -Format "MM.dd.yyyy-HH.mm.ss"
	# get the file name for our computer list
	$ComputerListLeaf = Split-Path -Path $ComputerListFile -Leaf
	# create a new file name for our old file
	$OldComputerListFile = "$($ComputerListLeaf.Replace('.txt', ''))-$($DateString).old"
	# create the full path for our old computer list file
	$OldComputerListFile = Join-Path -Path $OldLogFolder -ChildPath $OldComputerListFile
	# move our file
	Move-Item -Path $ComputerListFile -Destination $OldComputerListFile
	Write-Host "Old Computer List: '$($OldComputerListFile)'"

	# output computer names that had the script finish without errors
	$ComputerSuccessFile = Join-Path -Path $LogFolder -ChildPath "ComputerSuccess.txt"
	Out-File -FilePath $ComputerSuccessFile -InputObject ($SuccessArray | Get-Unique)
	Write-Host "     Success List: '$($ComputerSuccessFile)'"

	# output computer names that had an error during the script
	$ComputerErrorFile = Join-Path -Path $LogFolder -ChildPath "ComputerError.txt"
	Out-File -FilePath $ComputerErrorFile -InputObject ($ErrorArray | Get-Unique)
	Write-Host "       Error List: '$($ComputerErrorFile)'"

	# write our new computer list file using our error list for the next run
	Out-File -FilePath $ComputerListFile -InputObject ($ErrorArray | Get-Unique)
	Write-Host "New Computer List: '$($ComputerListFile)'"

	# output our errors encountered during the script
	Out-File -FilePath $ErrorFile -InputObject $Error
	Write-Host "           Errors: '$($ErrorFile)'"

	# write log file
	$DateString = Get-Date -Format "MM.dd.yyyy-HH.mm.ss"
	$ResultsFile = Join-Path -Path $OldLogFolder -ChildPath "Results-$($DateString).txt"
	Out-File -FilePath $ResultsFile -InputObject $CountsArray
	Write-Host "          Results: '$($ResultsFile)'`n"
}

#endregion MAIN