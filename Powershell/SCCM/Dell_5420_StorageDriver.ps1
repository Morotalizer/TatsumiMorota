#Logfile
$Logfile = "C:\Windows\CCM\Logs\InjectWinRE_$env:computername.log"
function WriteLog
{
Param ([string]$LogString)
$Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
$LogMessage = "$Stamp $LogString"
Add-content $LogFile -value $LogMessage
}

WriteLog "Starting Injecting Storage Driver for WinRE image"
# Variables
$DriverName = "iaStorVD.inf"
$MountDir = "$env:SystemDrive\WinRE"
$DriverDir = "$env:SystemDrive\DrvTemp"

# Get latest version of the storage driver
$StorageDriver = Get-WindowsDriver -Online -All | Where-Object { $_.Inbox -eq $False -and $_.BootCritical -eq $True -and $_.OriginalFileName -match $DriverName } | Sort-Object Version -Descending | Select-Object -First 1
writeLog "Found Storage Driver: $($StorageDriver.Driver)"
WriteLog "Original Filename: $($StorageDriver.OriginalFileName)"

# Ensure there is a single driver of matching criteria before beginning
if ($null -ne $StorageDriver -and $StorageDriver.Count -eq 1) {
    # Create mount directory if it does not exist
    if (!(Test-Path -Path $MountDir)) {
        New-Item -Path $MountDir -ItemType Directory
        WriteLog "Created Mountdir"
    }

    # Create export directory for driver if it does not exist
    if (!(Test-Path -Path $DriverDir)) {
        New-Item -Path $DriverDir -ItemType Directory
        WriteLog "Created ExportDir"
    }

    # Export driver
    pnputil.exe /export-driver $StorageDriver.Driver $DriverDir
    WriteLog "Exported Driver to $($DriverDir)"
    # Add to Windows RE image
    ReAgentC.exe /mountre /path $MountDir
    WriteLog "Mounted WinRE Image"
    dism /Image:$MountDir /Add-Driver /Driver:$DriverDir
    dism /Image:$MountDir /Cleanup-Image /StartComponentCleanup
    WriteLog "Added Driver"
    ReAgentc.exe /unmountre /path $MountDir /commit
    WriteLog "Unmouted WinRE"

    # Clean up
    Remove-Item -Path $DriverDir -Recurse
    Remove-Item -Path $MountDir
    WriteLog "Cleanup Directories"
}
# Throw an error so you can find devices that might need manual intervention
elseif($StorageDriver.Count -eq 0) {
    Write-error "No drivers detected. Expect value 1."
    writeLog "No driver detected. Expect value 1."
    Exit -1
}
else {
    Write-error "Invalid quanity of drivers detected. Expect value 1."
    writeLog "Invalid quanity of drivers detected. Expect value 1."
    $StorageDriver
    Exit -1
}