# This script will add User objects from a collection to an existing Active Directory group
$SiteCode = "R01"
$CollectionID = "R0100143"
$ADGroupName = "GPO-Firewall-TEST"
$AdObjects = New-Object System.Collections.ArrayList

#Import the ConfigurationManager.psd1 module 
Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
# Set the current location to be the site code. 
Set-Location "$SiteCode`:"


<# USER
$CollectionMembers = Get-CMUser -CollectionId $CollectionID | Select -Property SMSID | Sort-Object SMSID

foreach ($CollectionMember in $CollectionMembers){
    $CollectionADName = $CollectionMember.SMSID.Replace("RAGNSELLS\"," ")
    Add-ADPrincipalGroupMembership -Identity $CollectionADName -MemberOf $ADGroupName
}
#>

# Computer
$CollectionMembers = Get-CMDevice -CollectionId $CollectionID | Select -Property Name | Sort-Object Name
foreach ($CollectionMember in $CollectionMembers)
    {
    $obj = Get-ADComputer $CollectionMember.name
    Add-ADGroupMember $ADGroupName -Members $obj.SamAccountName
    }
