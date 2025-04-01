#Requires -RunAsAdministrator

$ComputerName = $env:computername

function Format-Output {
	param ([string]$Text)
	$Output = "[--|{0}| {1}" -f $ComputerName, $Text
	Write-Host $Output
}

try {
	Format-Output "Connected"
	
	$PathArray = Resolve-Path -Path 'C:\Program Files\WindowsApps\WavesAudio*'
	foreach ($Path in $PathArray) {
		if (Test-Path -Path $path) {
			$ShortPath = Split-Path -Path $Path -Leaf
			Remove-Item -Path $Path -Force -Recurse
			if (Test-Path -Path $path) {
				Format-Output "-- Failed to remove '$($ShortPath)'"
			}
			else {
				Format-Output "-- Removed '$($ShortPath)'"
			}
		}
	}
	
	# trigger sccm scan
	Format-Output "Running Hardware Inventory Cycle"
	Invoke-WmiMethod -Namespace 'root\ccm' -Class 'sms_client' -Name 'TriggerSchedule' -ArgumentList '{00000000-0000-0000-0000-000000000001}'
	
	Format-Output "Done`n"
}
catch {
	Format-Output "Error in script"
	Format-Output $_
	Write-Error $_
}
