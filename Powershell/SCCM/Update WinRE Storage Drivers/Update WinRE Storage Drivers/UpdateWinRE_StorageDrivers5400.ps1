# Variables
$DriverName = "iaStorAC.inf"
$MountDir = "$env:SystemDrive\WinRE"
$DriverDir = "$env:SystemDrive\DrvTemp"

# Get latest version of the storage driver
$StorageDriver = Get-WindowsDriver -Online -All | Where-Object { $_.Inbox -eq $False -and $_.BootCritical -eq $True -and $_.OriginalFileName -match $DriverName } | Sort-Object Version -Descending | Select-Object -First 1

# Ensure there is a single driver of matching criteria before beginning
if ($null -ne $StorageDriver -and $StorageDriver.Count -eq 1) {
    # Create mount directory if it does not exist
    if (!(Test-Path -Path $MountDir)) {
        New-Item -Path $MountDir -ItemType Directory
    }

    # Create export directory for driver if it does not exist
    if (!(Test-Path -Path $DriverDir)) {
        New-Item -Path $DriverDir -ItemType Directory
    }

    # Export driver
    pnputil.exe /export-driver $StorageDriver.Driver $DriverDir
    # Add to Windows RE image
    ReAgentC.exe /mountre /path $MountDir
    dism /Image:$MountDir /Add-Driver /Driver:$DriverDir
    dism /Image:$MountDir /Cleanup-Image /StartComponentCleanup
    ReAgentc.exe /unmountre /path $MountDir /commit

    # Clean up
    Remove-Item -Path $DriverDir -Recurse
    Remove-Item -Path $MountDir
}
# Throw an error so you can find devices that might need manual intervention
else {
    Write-Error "Invalid quanity of drivers detected. Expect value 1."
    $StorageDriver
    Exit -1
}