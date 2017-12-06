<#
.SYNOPSIS
   Runs USMT 5.0 to migrate app/user/OS settings from Win7 to Win10 
.DESCRIPTION
   Runs USMT 5.0 to migrate app/user/OS settings from Win7 to Win10
   USMT 5 will only migrate from W7 and later.  
   Run as administrator
.PARAMETER Mode
 Pass either ScanState or LoadState to this parameter.  Any other parameter passed will throw an error.
 ScanState will scan the local machine and upload data to a network share.
 LoadState will load the uploaded data that was created from ScanState and apply the settings on the local machine.
.EXAMPLE
    ./LS_Migrate.ps1 -mode scanstate
  
 Scan and upload the local machine settings to a network location.  This is for the logged on user.
 This must be ran first before 'loadstate' mode in order to create the migration store.  
.EXAMPLE
 ./LS_Migrate.ps1 -mode loadstate
  
 Download and apply the settings from the data that was gathered using ScanState mode
.EXAMPLE
 ./LS_Migrate.ps1 -mode scanstate -VerboseOff
 ./LS_Migrate.ps1 -mode loadstate -VerboseOff
  
 Turns off Scanstate or Loadstate Verbose mode.  
  
#>
 
param(
[Parameter(Mandatory=$true)]
[ValidateSet("scanstate","loadstate")]
[string]
$Mode,
 
[Parameter(Mandatory=$false)]
[switch]
$VerboseOff = $false
)

if ($verboseOff) {$VerbosePreference="SilentlyContinue"} else {$VerbosePreference="Continue"}
 
$osArch = Get-WmiObject -Class Win32_Processor 
switch ($osArch.AddressWidth)
{
"32" {$srcFolder = "\\dp\App\USMT\Files\x86"}
"64" {$srcFolder = "\\dp\App\USMT\Files\amd64"}
} 
#$srcFolder = "\\myServer\users\usmt4"
$CurrentUserFull = Get-WMIObject -class Win32_ComputerSystem | select username
$CurrentUser = ($CurrentUserFull.username -split "\\")[1]
$usmtFolder = "c:\usmt"
$storePath = "\\sccm\usmt$\MigStore_$CurrentUser"
 
 
function scanstate {
# If USMT folder exist, do not copy from network location
if (!(Test-Path $usmtFolder)) {
Write-Verbose "Copying data from $srcFolder to $usmtFolder"
copy $srcFolder $usmtFolder -Recurse
}
 
Set-Location $usmtFolder
 
# sets UMST Scanstate to migrate Docs,Apps and User settings.  Also sets logging at the highest level and specifies to only
# migrate any users who are currently logged on.  LocalOnly parameter is specified to omit removeable or network drive data.
#$ScanStateCmd = ".\scanstate.exe $storePath /i:migdocs.xml /i:migapp.xml /i:miguser.xml /localonly /o /v:13 /uel:0 /c /l:scanstate.log"
$ScanStateCmd = ".\scanstate.exe $storePath /i:migdocs.xml /i:migapp.xml /i:miguser.xml /localonly /o /v:13 /ue:*\* /ui:*\$CurrentUser /c /l:scanstate.log"
Write-Verbose "Running command: $scanStateCmd"
Invoke-Expression $scanStateCmd
}
 
function loadstate {
# If USMT folder exist, do not copy from network location
if (!(Test-Path $usmtFolder)) {
Write-Verbose "Copying data from $srcFolder to $usmtFolder"
copy $srcFolder $usmtFolder -Recurse
}
 
Set-Location $usmtFolder
 
# Loads the user store from the network location
#$LoadStateCmd = ".\loadstate.exe $storePath /i:migdocs.xml /i:migapp.xml /i:miguser.xml /v:13 /uel:0 /c /l:loadstate.log"
$LoadStateCmd = ".\loadstate.exe $storePath /i:migdocs.xml /i:migapp.xml /i:miguser.xml /v:13 /ue:\ /ui:*\$CurrentUser /c /l:loadstate.log"
Write-Verbose "Running command: $LoadStateCmd"
Invoke-expression $LoadStateCmd
}

switch ($mode) {
"scanstate" {
scanstate
}
"loadstate" {
if (Test-Path $storePath) {
loadstate
} else {
Write-warning "User store for $CurrentUser cannot be found at $storePath `nRun script in 'ScanState' mode first."
Write-warning "Script exiting..."
Write-Host
break
}
}
}