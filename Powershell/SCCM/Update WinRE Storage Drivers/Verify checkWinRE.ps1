﻿#Just to check if WinRE contains andy drivers
#Set-ExecutionPolicy -ExecutionPolicy Bypass

$DriverName = "iaStorVD.inf"
$MountDir = "$env:SystemDrive\WinRE"
#$DriverDir = "$env:SystemDrive\DrvTemp"
New-Item -Path $MountDir -ItemType Directory
#New-Item -Path $DriverDir -ItemType Directory

ReAgentC.exe /mountre /path $MountDir

dism /Image:$MountDir /get-drivers

#dism /Image:$MountDir /remove-driver /Driver:oem0.inf

ReAgentc.exe /unmountre /path $MountDir /discard
#Remove-Item -Path $DriverDir -Recurse
Remove-Item -Path $MountDir