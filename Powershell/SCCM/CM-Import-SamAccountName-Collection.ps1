#Convert full name to samaccountname
Import-Module ActiveDirectory
$names = Get-Content 'E:\Temp\user_java.txt'
$users = @()
foreach ($name in $names) {
  $nameParts = $name -split ' '
  $lastName = $nameParts[0] -replace ',',''
  $firstName = $nameParts[1]
  $users += Get-ADUser -Filter { Surname -eq $lastName -and GivenName -eq $firstName -and SamAccountName -notlike '*A1' -and Enabled -eq $true }
}
$users | Format-Table -AutoSize -Property SamAccountName | Out-File -FilePath e:\temp\samaccount.txt


#Import samaccountname to collection
$ErrorActionPreference= 'silentlycontinue'
#Collection must pre-exist
$CollectionName = "Compliance - Java Needed"
#List of names must pre-exist
$Users = get-content "E:\Temp\samaccount.txt"
Foreach ($User in $Users)
{Add-CMUserCollectionDirectMembershipRule -collectionname $CollectionName -resourceid (Get-CMUser -name ragnsells\$User).ResourceID}



