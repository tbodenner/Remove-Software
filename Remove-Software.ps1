#Requires -RunAsAdministrator

# parameters
param (
	# the folder that contains the payload and computer list
	[Parameter(Mandatory=$True)][string]$InputPath
)

# read computer list text file and return a unique sorted array
function Get-UniqueArrayFromFile {
	param (
		[Parameter(Mandatory=$True)][string]$FilePath
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
		if (($Null -eq $Computer) -or ($Computer -eq "")) { continue }
		# add the computer to the hashtable, duplicates will not be added
		$ComputerSet[$Computer] = $Null
	}
	# return a sorted array of the hashtable keys
	return $ComputerSet.Keys | Sort-Object
}

# our custom line format for script output
function Format-Line {
	# parameters
	param (
	[Parameter(Mandatory=$True)][string][string]$Text,
	[string]$Computer
	)
	# check if we were given a computer name
	if ($Computer -ne "") {
		# format the line with the our text and computer
		return "[--|$($Computer)| $($Text)"
	}
	else {
		# otherwise, format the line with the our text
		return "[--|$($Text)|"
	}
}

# write a colored line to the host
function Write-ColorLine {
	# parameters
	param (
		[Parameter(Mandatory=$True)][string]$Text,
		[Parameter(Mandatory=$True)][string]$Color
	)
	# check if our color is in the color enum
	if ([ConsoleColor]::IsDefined([ConsoleColor], $Color)) {	
		# color exists, write our colored line
		Write-Host $Text -ForegroundColor $Color
	}
	else {
		# color doesn't exist, write our line without color
		Write-Host $Text
	}
}

# get an array of all our computers in AD
function Get-AdComputerArray {
	# the AD object we are going to search
	$SearchBase = 'OU=Prescott (PRE),OU=VISN18,DC=v18,DC=med,DC=va,DC=gov'
	# get our AD data
	$ADComputers = Get-ADComputer -Filter * -SearchBase $SearchBase | Select-Object Name
	# create our empty array
	$ADArray = @()
	# create a clean array from our AD data
	foreach ($Item in $ADComputers) {
		$ADArray += $Item.Name
	}
	# return our array
	$ADArray
}

# test if a computer is in AD
function Test-ComputerInAD {
	# parameters
	param (
		[Parameter(Mandatory=$True)][string[]]$ADArray,
		[Parameter(Mandatory=$True)][string]$ComputerName
	)
	# check if our AD array is null
	if ($Null -eq $ADArray) {
		# computer array is null, return false
		$False
	}
	# check if the computer is in our AD array
	if ($ADArray -contains $ComputerName) {
		# if the computer is in our array, return true
		$True
	}
	else {
		# otherwise, return false
		$False
	}
}

# check if a computer can be found
function Find-Computer {
	# parameters
	param (
		[Parameter(Mandatory=$True)][string[]]$ADArray,
		[Parameter(Mandatory=$True)][string]$ComputerName
	)
	# check if the computer is in AD
	if ((Test-ComputerInAd -ADArray $ADArray -ComputerName $ComputerName) -eq $False) {
		return @(0, $NotInAD)
	}
	# ping the computer and save the details
	try {
		$ComputerDetails = Test-Connection -TargetName $ComputerName -Count 1 -TimeoutSeconds 3
	}
	# catch the ping exception
	catch [System.Net.NetworkInformation.PingException] {
		return @(0, $NotFound)
	}
	# catch all other errors
	catch {
		# write out the exception
		Write-Host ($_.Exception | Select-Object -Property *)
	}
	# get the pinged computer's ip
	$Ip = $ComputerDetails.Address
	# get the ping result
	$Latency = $ComputerDetails.Latency
	# get the status of the ping
	$Status = $ComputerDetails.Status
	# check if our status is null
	if ($Null -eq $Status) {
		# if null, set our status
		$Status = $NoStatus
	}
	# check if the ping timed out
	if (($Latency -eq 0) -or ($Status -eq $NoStatus)) {
		# the ping timed out, so return the result of our ping
		return @($Latency, $Status)
	}
	# set our error value for our dns name
	$DnsName = $Null
	# check if we got an ip
	if ($Null -ne $Ip) {
		# get our dns data
		try {
			$DnsData = Resolve-DnsName -Name $Ip
		}
		# catch the dns error
		catch [System.ComponentModel.Win32Exception] {
			# get our error code
			$ECode = $_.Exception.NativeErrorCode
			# return our error
			if ($ECode -eq 9003) {
				return @(0, 'DnsNotFound')
			}
		}
		# catch all other errors
		catch {
			# write out the exception
			Write-Host ($_.Exception | Select-Object -Property *)
		}
		# check if we got any data
		if ($Null -eq $DnsData) {
			return @($Latency, 'GenericDnsError')
		}
		# get our host name from the data
		$NameHost = $DnsData.NameHost
		# check if we got a hostname
		if ($Null -eq $NameHost) {
			return @($Latency, 'NoHostName')
		}
		# split the host name and return it
		$DnsName = $NameHost.Split('.')
	}
	# check if our resolved name is null
	if ($Null -eq $DnsName) {
		# if true, write a message
		Write-ColorLine -Text (Format-Line -Text "Unable to resolve computer name from IP address" -Computer $ComputerName) -Color Red
	}
	else {
		# otherwise, check if our computer name matches the dns name
		if ($DnsName[0] -ne $ComputerName) {
			Write-ColorLine -Text (Format-Line -Text "DNS mismatch (DNS: $($DnsName[0]), CN: $($ComputerName))" -Computer $ComputerName) -Color Red
			# return the dns error
			return @($Latency, $DnsMismatch)
		}
	}
	# return the result of our ping
	return @($Latency, $Status)
}

# stop on errors
$ErrorActionPreference = "Stop"

# if our status is empty, use this status
$NoStatus = 'NoStatus'
# if our ping fails to find a target, use this status
$NotFound = 'NotFound'
# if we have a dns mismatch, use this status
$DnsMismatch = 'DnsMismatch'
# if our computer is not in AD, use this status
$NotInAd = 'NotInAd'

# test if our input folder exists
if ((Test-Path -Path $InputPath) -eq $False) {
	Write-ColorLine -Text "Error: Input folder not found." -Color Red
	return
}

# log folder
$LogFolderName = 'Logs'
$LogFolder = Join-Path -Path $InputPath -ChildPath $LogFolderName

# test if our log folder exists
if ((Test-Path -Path $LogFolder) -eq $False) {
	New-Item -Path $LogFolder -ItemType "directory" | Out-Null
}

# old log folder
$OldLogFolderName = 'Old'
$OldLogFolder = Join-Path -Path $LogFolder -ChildPath $OldLogFolderName

# test if our old log folder exists
if ((Test-Path -Path $OldLogFolder) -eq $False) {
	New-Item -Path $OldLogFolder -ItemType "directory" | Out-Null
}

# file to output errors
$ErrorFile = Join-Path -Path $LogFolder -ChildPath "Errors.txt"
# file that contains our script
$PayloadFile = Join-Path -Path $InputPath -ChildPath "Payload.ps1"
# file that contains our computer list
$ComputerListFile = Join-Path -Path $InputPath -ChildPath "ComputerList.txt"

# test if our required paylod file exists
if ((Test-Path -Path $PayloadFile) -eq $False) {
	Write-ColorLine -Text "Error: Payload file not found." -Color Red
	return
}
# test if our required computer list file exists
if ((Test-Path -Path $ComputerListFile) -eq $False) {
	Write-ColorLine -Text "Error: Computer list file not found." -Color Red
	return
}

# an array to collect the computer names with an errror
$ErrorArray = @()

# an array to collect the computer names that succeeded
$SuccessArray = @()

# get the list of computers from a text file
$ComputerList = Get-UniqueArrayFromFile -FilePath $ComputerListFile

# write our starting status
$Plural = ""
if ($ComputerList.Count -eq 1) {
	$Plural = "computer"
}
else {
	$Plural = "computers"
}
$StartString = "`nRunning '$($PayloadFile)' on $($ComputerList.Count) $($Plural)`n"
Write-ColorLine -Text $StartString -Color Green

# clear the error list so we can write only our errors
$Error.Clear()

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

# get an array of AD computers
$ADComputerArray = Get-AdComputerArray

# loop through list of computers
foreach ($Computer in $ComputerList) {
	try {
		# check of we have a computer name
		if (($Null -eq $Computer) -or ($Computer -eq "")) { continue }
		# get the last error in the error variable
		$LastError = $Error[0]
		# update our progress
		$PComplete = ($ComputerCount / $TotalComputers) * 100
		$Status = "$ComputerCount/$TotalComputers Complete"
		$Activity = Format-Line -Text "Progress   "
		Write-Progress -Activity $Activity -Status $Status -PercentComplete $PComplete
		Write-ColorLine -Text (Format-Line -Text "Trying To Connect" -Computer $Computer) -Color Yellow
		# set our parameters for our invoke command
		$Parameters = @{
			ComputerName	= $Computer
			FilePath		= $PayloadFile
			ErrorAction		= "SilentlyContinue"
			SessionOption	= $PssOptions
		}
		# get our ping data
		$PingData = Find-Computer -ADArray $ADComputerArray -ComputerName $Computer
		# get the result (boolean)
		$PingLatency = $PingData[0]
		# get the ping status (string)
		$PingStatus = $PingData[1]
		# check if we can ping the computer
		if ($PingLatency -gt 0) {
			# check if we did not have a dns mismatch
			if ($PingStatus -ne $DnsMismatch) {
				Write-ColorLine -Text (Format-Line -Text "Ping ($($PingStatus))" -Computer $Computer) -Color Green
				# run the script on the target computer if we can ping the computer
				$InvokeReturn = Invoke-Command @Parameters
				# check if we got anything back from the invoke command
				if ($Null -ne $InvokeReturn) {
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
			}
			else {
				# otherwise, write an error
				Write-Error -Message "DNS mismatch error" -Category ConnectionError -ErrorAction SilentlyContinue
			}
		}
		else {
			Write-ColorLine -Text (Format-Line -Text "Ping ($($PingStatus))" -Computer $Computer) -Color Red
			# otherwise, write an error
			Write-Error -Message "Unable to ping $($Computer) - ($($PingStatus), $($PingLatency))" -Category ConnectionError -ErrorAction SilentlyContinue
		}
		# check if our computer is in AD before adding to either array
		if ($PingStatus -ne $NotInAd) {
			# determine if the last computer was a success or error
			if ($LastError -eq $Error[0]) {
				# if the script finished without adding an error, add the computer to our success array
				$SuccessArray += $Computer
			}
			else {
				Write-ColorLine -Text "Error: $($Computer)`n" -Color Yellow
				# if the script added an error, add the computer to our success array
				$ErrorArray += $Computer
			}
		}
		else {
			# write an empty line after our not in AD message
			Write-Host
		}
		# increment our count
		$ComputerCount += 1
	}
	catch {
		Write-Output "Caught Error"
		Write-Output "$($_)`n"
		Write-Host ($_ | Select-Object -Property *)
		# add the computer to our error array if an error was caught
		$ErrorArray += $Computer
	}
}

# create our counts array to output to the console and a file
$CountsArray = @(
	'Results:'
	"    Total: $($TotalComputers)"
	"  Success: $($SuccessArray.Count)"
	"     Skip: $($SkipCount)"
	"Uninstall: $($UninstallCount)"
	"    Error: $($ErrorArray.Count)"
)

# write our counts
Write-Host $CountsArray[0]
Write-ColorLine -Text $CountsArray[1] -Color Yellow
Write-ColorLine -Text $CountsArray[2] -Color Green
Write-ColorLine -Text $CountsArray[3] -Color Cyan
Write-ColorLine -Text $CountsArray[4] -Color Blue
Write-ColorLine -Text $CountsArray[5] -Color Red

# write our output files
Write-Host # empty line
Write-ColorLine -Text 'Writing Output Files' -Color Yellow

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
$ComputerSucessFile = Join-Path -Path $LogFolder -ChildPath "ComputerSuccess.txt"
Out-File -FilePath $ComputerSucessFile -InputObject $SuccessArray
Write-Host "     Success List: '$($ComputerSucessFile)'"

# output computer names that had an error during the script
$ComputerErrorFile = Join-Path -Path $LogFolder -ChildPath "ComputerError.txt"
Out-File -FilePath $ComputerErrorFile -InputObject $ErrorArray
Write-Host "       Error List: '$($ComputerErrorFile)'"

# write our new computer list file using our error list for the next run
Out-File -FilePath $ComputerListFile -InputObject $ErrorArray
Write-Host "New Computer List: '$($ComputerListFile)'"

# output our errors encountered during the script
Out-File -FilePath $ErrorFile -InputObject $Error
Write-Host "           Errors: '$($ErrorFile)'"

# write log file
$DateString = Get-Date -Format "MM.dd.yyyy-HH.mm.ss"
$ResultsFile = Join-Path -Path $OldLogFolder -ChildPath "Results-$($DateString).txt"
Out-File -FilePath $ResultsFile -InputObject $CountsArray
Write-Host "          Results: '$($ResultsFile)'`n"
