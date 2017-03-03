﻿<#  
.SYNOPSIS  
    Creates folders and sets permission for SCCM   
.DESCRIPTION  
    Creates default SCCM folder structures and sets permissions on those folders 
.NOTES  
    File Name  : SCCMFolderStructure.ps1  
    Author     : Tatsumi Morota - tatsumi.morota@rtsab.com  
.LINK

#>

#Set the Following Parameters
$Source = 'I:\SCCMSRC'
$ShareName = 'SCCMSRC$'
$NetworkAccount = 'DOMAIN\svc_sccm_naa'
$HardwareHP = Import-Csv "I:\Drift\ModelsHP.csv"
#$HardwareDell = Import-Csv "I:\Drift\ModelsDell.csv"
#$HardwareLenovo = Import-Csv "I:\Drift\ModelsLenovo.csv"
$HardwareAppsHP = "$Source\HardwareAPPS\Windows 7 x64\HP"
#$HardwareAppsDell = "$Source\HardwareAPPS\Windows 7 x64\Dell"
#$HardwareAppsLenovo = "$Source\HardwareAPPS\Windows 7 x64\Lenovo"

#Create Source Directory
New-Item -ItemType Directory -Path "$Source"


#Create Application Directory Structure
New-Item -ItemType Directory -Path "$Source\APPS"
New-Item -ItemType Directory -Path "$Source\APPS\Adobe"
New-Item -ItemType Directory -Path "$Source\APPS\Apple"
New-Item -ItemType Directory -Path "$Source\APPS\Citrix"
New-Item -ItemType Directory -Path "$Source\APPS\Microsoft"

#Create App-V Directory Structure
New-Item -ItemType Directory -Path "$Source\App-V"
New-Item -ItemType Directory -Path "$Source\App-V\Packages"
New-Item -ItemType Directory -Path "$Source\App-V\Source"

#Create Hardware Application Directory Structure
New-Item -ItemType Directory -Path "$Source\HardwareAPPS"
New-Item -ItemType Directory -Path "$Source\HardwareAPPS\Windows 7 x64"
New-Item -ItemType Directory -Path "$Source\HardwareAPPS\Windows 7 x64\Dell"
New-Item -ItemType Directory -Path "$Source\HardwareAPPS\Windows 7 x64\HP"
New-Item -ItemType Directory -Path "$Source\HardwareAPPS\Windows 7 x64\Lenovo"
New-Item -ItemType Directory -Path "$Source\HardwareAPPS\Windows 10 x64"
New-Item -ItemType Directory -Path "$Source\HardwareAPPS\Windows 10 x64\Dell"
New-Item -ItemType Directory -Path "$Source\HardwareAPPS\Windows 10 x64\HP"
New-Item -ItemType Directory -Path "$Source\HardwareAPPS\Windows 10 x64\Lenovo"

Set-Location $HardwareAppsHP

foreach ($Hardware in $HardwareHP){
    new-item $HardwareHP.Models -type directory
    #new-item ($HardwareHP.Models+"\x86") -type directory
    #new-item ($HardwareHP.Models+"\x64") -type directory
    } 
Set-Location $HardwareAppsHP

#Create Hotfix Directory Structure
New-Item -ItemType Directory -Path "$Source\Hotfix"

#Create Import Directory Structure
New-Item -ItemType Directory -Path "$Source\Import"
New-Item -ItemType Directory -Path "$Source\Import\Baselines"
New-Item -ItemType Directory -Path "$Source\Import\MOFs"
New-Item -ItemType Directory -Path "$Source\Import\Task Sequences"

#Create Log Directory Structure
New-Item -ItemType Directory -Path "$Source\Logs"
New-Item -ItemType Directory -Path "$Source\Logs\MDTLogs"
New-Item -ItemType Directory -Path "$Source\Logs\MDTLogsDL"

#Create OSD Directory Structure
New-Item -ItemType Directory -Path "$Source\OSD"
New-Item -ItemType Directory -Path "$Source\OSD\BootImages"
New-Item -ItemType Directory -Path "$Source\OSD\Branding"
New-Item -ItemType Directory -Path "$Source\OSD\Branding\WinPE Background"
New-Item -ItemType Directory -Path "$Source\OSD\Captures"
New-Item -ItemType Directory -Path "$Source\OSD\DriverPackages"
New-Item -ItemType Directory -Path "$Source\OSD\DriverPackages\Windows 7 x86"
New-Item -ItemType Directory -Path "$Source\OSD\DriverPackages\Windows 7 x86\Dell"
New-Item -ItemType Directory -Path "$Source\OSD\DriverPackages\Windows 7 x86\HP"
New-Item -ItemType Directory -Path "$Source\OSD\DriverPackages\Windows 7 x86\Lenovo"
New-Item -ItemType Directory -Path "$Source\OSD\DriverPackages\Windows 7 x86\VMWare"
New-Item -ItemType Directory -Path "$Source\OSD\DriverPackages\Windows 7 x64"
New-Item -ItemType Directory -Path "$Source\OSD\DriverPackages\Windows 7 x64\Dell"
New-Item -ItemType Directory -Path "$Source\OSD\DriverPackages\Windows 7 x64\HP"
New-Item -ItemType Directory -Path "$Source\OSD\DriverPackages\Windows 7 x64\Lenovo"
New-Item -ItemType Directory -Path "$Source\OSD\DriverPackages\Windows 7 x64\VMWare"
New-Item -ItemType Directory -Path "$Source\OSD\DriverPackages\Windows 10 x64"
New-Item -ItemType Directory -Path "$Source\OSD\DriverPackages\Windows 10 x64\Dell"
New-Item -ItemType Directory -Path "$Source\OSD\DriverPackages\Windows 10 x64\HP"
New-Item -ItemType Directory -Path "$Source\OSD\DriverPackages\Windows 10 x64\Lenovo"
New-Item -ItemType Directory -Path "$Source\OSD\DriverPackages\Windows 10 x64\VMWare"
New-Item -ItemType Directory -Path "$Source\OSD\DriverSources"
New-Item -ItemType Directory -Path "$Source\OSD\DriverSources\Windows 7 x86"
New-Item -ItemType Directory -Path "$Source\OSD\DriverSources\Windows 7 x86\Dell"
New-Item -ItemType Directory -Path "$Source\OSD\DriverSources\Windows 7 x86\HP"
New-Item -ItemType Directory -Path "$Source\OSD\DriverSources\Windows 7 x86\Lenovo"
New-Item -ItemType Directory -Path "$Source\OSD\DriverSources\Windows 7 x86\VMWare"
New-Item -ItemType Directory -Path "$Source\OSD\DriverSources\Windows 7 x64"
New-Item -ItemType Directory -Path "$Source\OSD\DriverSources\Windows 7 x64\Dell"
New-Item -ItemType Directory -Path "$Source\OSD\DriverSources\Windows 7 x64\HP"
New-Item -ItemType Directory -Path "$Source\OSD\DriverSources\Windows 7 x64\Lenovo"
New-Item -ItemType Directory -Path "$Source\OSD\DriverSources\Windows 7 x64\VMWare"
New-Item -ItemType Directory -Path "$Source\OSD\DriverSources\Windows 10 x64"
New-Item -ItemType Directory -Path "$Source\OSD\DriverSources\Windows 10 x64\Dell"
New-Item -ItemType Directory -Path "$Source\OSD\DriverSources\Windows 10 x64\HP"
New-Item -ItemType Directory -Path "$Source\OSD\DriverSources\Windows 10 x64\Lenovo"
New-Item -ItemType Directory -Path "$Source\OSD\DriverSources\Windows 10 x64\VMWare"
New-Item -ItemType Directory -Path "$Source\OSD\MDTSettings"
New-Item -ItemType Directory -Path "$Source\OSD\MDTToolkit"
New-Item -ItemType Directory -Path "$Source\OSD\OSImages"
New-Item -ItemType Directory -Path "$Source\OSD\OSInstall"
New-Item -ItemType Directory -Path "$Source\OSD\Prestart"
New-Item -ItemType Directory -Path "$Source\OSD\USMT"
New-Item -ItemType Directory -Path "$Source\OSD\SCRIPTS"

#Create Script Directory Structure
New-Item -ItemType Directory -Path "$Source\Script"

#Create State Capture Directory Structure
New-Item -ItemType Directory -Path "$Source\StateCapture"

#Create Tools Directory Structure
New-Item -ItemType Directory -Path "$Source\Tools"
New-Item -ItemType Directory -Path "$Source\Tools\PSTools"

#Create Windows Update Directory Structure
New-Item -ItemType Directory -Path "$Source\WindowsUpdates"
New-Item -ItemType Directory -Path "$Source\WindowsUpdates\Endpoint Protection"
New-Item -ItemType Directory -Path "$Source\WindowsUpdates\Lync 2010"
New-Item -ItemType Directory -Path "$Source\WindowsUpdates\Office 2010"
New-Item -ItemType Directory -Path "$Source\WindowsUpdates\Silverlight"
New-Item -ItemType Directory -Path "$Source\WindowsUpdates\Visual Studio 2008"
New-Item -ItemType Directory -Path "$Source\WindowsUpdates\Windows 7"
New-Item -ItemType Directory -Path "$Source\WindowsUpdates\Windows 8"

#Create WSUS Directory
New-Item -ItemType Directory -Path "$Source\WSUS"

#Create the Share and Permissions
New-SmbShare -Name "$ShareName” -Path “$Source” -CachingMode None -FullAccess Everyone

#Set Security Permissions
$Acl = Get-Acl "$Source\Logs"
$Ar = New-Object System.Security.AccessControl.FileSystemAccessRule("$NetworkAccount","FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
$Acl.SetAccessRule($Ar)
Set-Acl "$Source\Logs" $Acl

$Acl = Get-Acl "$Source\OSD\Captures"
$Ar = New-Object System.Security.AccessControl.FileSystemAccessRule("$NetworkAccount","FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
$Acl.SetAccessRule($Ar)
Set-Acl "$Source\OSD\Captures" $Acl

$Acl = Get-Acl "$Source\StateCapture"
$Ar = New-Object System.Security.AccessControl.FileSystemAccessRule("LOCALSERVICE","FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
$Acl.SetAccessRule($Ar)
Set-Acl "$Source\StateCapture" $Acl