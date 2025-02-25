#Requires -version 2.0


# install kasa via pip install python-kasa
#
# Enable it to run at logon automatically by creating a symlink in Windows:
#   1. Hit Win+R
#   2. type "shell:startup", hit Enter
#   3. right-click, New, Shortcut
#   4. Target:
#   	C:\Windows\System32\cmd.exe /c start /min "" powershell.exe -ExecutionPolicy Bypass -WindowStyle hidden -Command "C:\path\to\webcam.ps1"
#
# NOTE username / password is the tp-link cloud password

$kasaDeviceIP = ""    # NOTE fill in
$username = ""        # NOTE fill in
$password = ""        # NOTE fill in
$kasaBinary = ""      # NOTE fill in
$deviceType = "smart" # change if necessary
$sleepTimeMs = "250"
$regMaxDepth = 5
$user = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$regBasePath = "Registry::HKEY_USERS\$user\Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\webcam"

function KasaDeviceDo {
    param ([string[]]$action)
    & $kasaBinary --username $username --password $password --host $kasaDeviceIP --type $deviceType @action
}

function Scan-WebcamRegistry {
	param ([string]$path, [int]$depth)

	if ($depth -gt $regMaxDepth) { return $false }
	if (-not (Test-Path $path)) { return $false }
	
	try {
		$subKeys = Get-ChildItem -Path $path -ErrorAction Continue

		foreach ($subKey in $subKeys) {
			$webcamUsage = Get-ItemProperty -Path $subKey.PSPath -ErrorAction Continue
			
			if ($webcamUsage) {
				$startTimeRaw = $webcamUsage.PSObject.Properties["LastUsedTimeStart"].Value
				$stopTimeRaw = $webcamUsage.PSObject.Properties["LastUsedTimeStop"].Value
				
				# Webcam is in use if Start > 0 and Stop is 0
				if ($startTimeRaw -gt 0 -and $stopTimeRaw -eq 0) {
					return $true
				}
			}

			if (Scan-WebcamRegistry -path $subKey.PSPath -depth ($depth + 1)) {
				return $true
			}
		}
	} catch {
		Write-Output "Error reading registry: $_"
	}
	return $false
}

$stopScript = $false

# Catch termination signal
$Handler = {
    $global:stopScript = $true
}
Register-EngineEvent -SourceIdentifier "PowerShell.Exiting" -Action $Handler


KasaDeviceDo -action @("feature", "auto_off_minutes", "1")
while (-not $stopScript) {
	if (Scan-WebcamRegistry -path $regBasePath -depth 0) {
		KasaDeviceDo -action @("on")
		# simple way to reset the auto_off_at
		# we want to have the plug shut off the light if this script dies unexpectedly
		# but simply turning it on will not reset the auto-off time
		KasaDeviceDo -action @("feature", "auto_off_enabled", "False")
		KasaDeviceDo -action @("feature", "auto_off_enabled", "True")
	} else {
		KasaDeviceDo -action @("off")
	}
    Start-Sleep -Milliseconds $sleepTimeMs
}
# turn off the light if we're exiting gracefully
KasaDeviceDo -action "off"