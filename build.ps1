$hidapi = 'https://github.com/libusb/hidapi/releases/download/hidapi-0.9.0/hidapi-win.zip'
$ohm = 'https://openhardwaremonitor.org/files/openhardwaremonitor-v0.9.5.zip'
$libs = @(
        'hidapi-win\x86\hidapi.dll',
        'hidapi-win\x86\hidapi.lib',
        'OpenHardwareMonitor\OpenHardwareMonitorLib.dll')
$buildArgs = @(
    '-F', '-w', '-i', 'main.ico', '--uac-admin',
    '--add-data', 'OpenHardwareMonitorLib.dll;.',
    '--add-data', 'hidapi.dll;.',
    '--add-data', 'hidapi.lib;.',
    'thermostat.py')
$pyinstaller = ''
$temp = $env:TEMP

# Fetch hidapi/openhardwaremonitor if they're not already present
if ((Split-Path $libs -Leaf | Test-Path) -contains $false) {
    $hidapizip = Join-Path $temp 'hidapi.zip'
    $ohmzip = Join-Path $temp 'ohm.zip'

    # Powershell so umm, I guess gotta enable TLS, because it's disabled?
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # Download
    Invoke-WebRequest $hidapi -OutFile $hidapizip
    Invoke-WebRequest $ohm -OutFile $ohmzip

    # Extract
    Expand-Archive -Path $hidapizip -DestinationPath $temp
    Expand-Archive -Path $ohmzip -DestinationPath $temp

    # Move things around
    Copy-Item $libs.forEach({Join-Path $temp $_}) -Destination $PSScriptRoot
}

# Try to use system's version of pyinstaller
if (!$pyinstaller) {
    if (Get-Command 'pyinstaller.exe' -ErrorAction SilentlyContinue) {
        $pyinstaller = 'pyinstaller.exe'
    } else {
        $pyinstaller = 'c:\python38-32\Scripts\pyinstaller.exe'
    }
}

& $pyinstaller $buildArgs

