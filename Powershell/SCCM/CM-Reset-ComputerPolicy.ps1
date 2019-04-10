#$Computername is the name of the computer having issue

$ComputerName = Read-Host 'Enter Computername'
Get-Service -Name CcmExec -ComputerName $ComputerName | Stop-service
Get-Service -Name Bits -ComputerName $ComputerName | Stop-service
Remove-Item -Path \\$ComputerName\c$\ProgramData\Microsoft\Network\Downloader\qmgr.db -Force
Get-Service -Name CcmExec -ComputerName $ComputerName | Start-Service
Invoke-WmiMethod -ComputerName $ComputerName -Namespace root\ccm -Class sms_client -Name ResetPolicy -ArgumentList @(1)