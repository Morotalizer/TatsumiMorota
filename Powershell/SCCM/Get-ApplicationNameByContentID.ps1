
$ContentID ="Content_9fc64cef-9537-4c0b-822e-6e347049a7e7"
Get-WmiObject -ComputerName SCCMapp1 -Namespace "root\sms\site_PS1" -Query "select SecuredModelName from SMS_CIToContent where ContentUniqueID = '$ContentID'" | select -first 1 | %{get-wmiobject -ComputerName SCCMAPP1 -Namespace "root\sms\site_PS1" -query "select LocalizedDisplayName from SMS_Application where ModelName = '$($_.SecuredModelName)'"} | select LocalizedDisplayName -first 1
