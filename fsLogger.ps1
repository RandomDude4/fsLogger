# Requires fsSimconnect.ps1 in the same folder
# Must be run in 32 bit Powershell with admin rights


$startKST = 1
$restartKST = 1		# Try and restart KST if it has been closed or has crashed
$kstfile = "KST_profile.kst"

#$csvfile = "Data_" + (Get-Date -Format "yyy-MM-dd_HH.mm.ss") + ".csv"
$csvfile = "Data.csv"

$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition
cd $scriptPath

if (!$global:fsConnected) {
    .\"fsSimconnect.ps1"
}

# The following .Net 4.5 library is required to find correct mime type for file extensions
[System.Reflection.Assembly]::LoadWithPartialName("System.Web") | Out-Null


Write-Host "Creating $csvfile"
New-Item "$csvfile"  -ItemType File -Force | Out-Null

# Add first line of descriptions to csv-file
$str = "Time"
For ($i=0; $i -lt $global:simnames.length; $i++) {
	$str = "$str,$($global:simnames[$i])"
}
Add-Content -Path "$csvfile" -Value ($str)

# Init-time variables
$t_prev = 0
$t_start = 0

Write-Host "Get first data point:" -NoNewline
while ($t_start -lt 1) {	# Sometimes it reads 0, wait a bit and try again?!
	Start-Sleep -m 500
	$t_start = $global:sim.ABSOLUTE_TIME
	Write-Host "." -NoNewline
}
Write-Host "OK"

# Start KST
#if (($startKST) -and ($t_new -gt 2.0) ) {
#	Write-Host "Start KST:"
#	#$startKST = 0
#	Write-Host "Starting KST"
#	#Start-Process -FilePath "$($scriptPath)\Kst\bin\kst2.exe" -ArgumentList "-F","$csvfile", "$kstfile"
#	Start-Process -FilePath "\Kst\bin\kst2.exe" -ArgumentList "$kstfile"
#}


Write-Host "Starting data collection..."
while ($global:fsConnected) {
	$data = $global:sim
	$t_new = $data.ABSOLUTE_TIME - $t_start
	

	# Check that new value is not from identical time-step as previous point
	if ($t_new -gt ($t_prev+0.001)) {
		#Write-Host $t_new
		
		# Add line of values to csv-file
		$str = "$t_new"
		For ($i=0; $i -lt $global:simnames.length; $i++) {
			#if ($global:simnames[$i] -like "Plane_Pitch_Degrees") {
			#	$sub = -$($data."$($global:simnames[$i])")
			#} else {
			$sub = $($data."$($global:simnames[$i])")
			#}
			$str = "$str,$sub"
		}
		#Add-Content -Path ./Data.csv -Value ($str)
		$str | Out-File "$csvfile" -Append -Encoding UTF8 
		
		# Start KST if it has closed
		if ($startKST) { 
			if ([math]::Round($t_new/4) -gt [math]::Round($t_prev/4) ) {
				$kstproc = Get-Process kst2 -ErrorAction SilentlyContinue
				if (!$kstproc) {
					Write-Host "Starting KST"
					if (!$restartKST) { $startKST = 0 }
					# Start-Process -FilePath "$scriptPath\Kst\bin\kst2.exe" -ArgumentList "-F","$csvfile","$kstfile"
					Start-Process -FilePath ".\Kst\bin\kst2.exe" -ArgumentList "$kstfile"
				}
			}
		}
		$t_prev = $t_new
	}
	Start-Sleep -m 75
}

Write-Host "FS disconnected"
Start-Sleep -m 4000

