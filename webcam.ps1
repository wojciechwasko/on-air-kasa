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

function Toggle-KasaDevice {
    param ([bool]$turnOn)
    $action = if ($turnOn) { "on" } else { "off" }
	& $kasaBinary --username $username --password $password --host $kasaDeviceIP --type $deviceType $action
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

while ($true) {
    if (Scan-WebcamRegistry -path $regBasePath -depth 0) {
        Toggle-KasaDevice -turnOn $true
    } else {
        Toggle-KasaDevice -turnOn $false
    }
    Start-Sleep -Milliseconds $sleepTimeMs
}
