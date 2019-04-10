#############################################################################
# Author  : Tatsumi Morota
# Website : www.advania.com
# Twitter : @TatsumiMorota
#
# Version : 1.1
# Created : 2019/02/07
# Modified : 
#
# Purpose : Create a folder structure for Windows 10 upgrade rings
# How to: Change line 16,17 for your windows Version and run script!
# Version History:¨
# 1.0 First Release
# 1.1 Added Upgrade - Assess collections to be used by Upgrade frontend Assess ps script
#############################################################################

#Change below variables to fit your Windows 10 version
#Variables
$W10YYMM = "1809" #Windows you will upgrade to format YYMM
$W10Ver = "17763" #Windows 10 Version number you will upgrade to

#Load Configuration Manager PowerShell Module
Import-module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')

#Get SiteCode
$SiteCode = Get-PSDrive -PSProvider CMSITE
Set-location $SiteCode":"

#Create Default Folders
new-item -NAme 'Servicing Windows 10' -Path $($SiteCode.Name+":\DeviceCollection")
new-item -NAme 'Operational' -Path $($SiteCode.name+":\DeviceCollection\")
new-item -NAme "Upgrade - $W10YYMM" -Path $($SiteCode.name+":\DeviceCollection\Servicing Windows 10")
new-item -NAme 'Batches' -Path $($SiteCode.name+":\DeviceCollection\")

#Create Collections
#List of Collections Query
$Collection1 = @{Name = "Workstations | Windows 10 $W10YYMM"; Query = "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_R_System.OperatingSystemNameandVersion like '%Microsoft Windows NT Workstation 10.0%' and SMS_G_System_OPERATING_SYSTEM.BuildNumber = '$W10Ver'"}
$Collection2 = @{Name = "Upgrade - Compliant for $W10YYMM"; Query = ""}
$Collection3 = @{Name = "Upgrade - Ring 1 - $W10YYMM"; Query = ""}
$Collection4 = @{Name = "Upgrade - Ring 2 - $W10YYMM"; Query = ""}
$Collection5 = @{Name = "Upgrade - Ring 3 - $W10YYMM"; Query = ""}
$Collection6 = @{Name = "Upgrade - Ring 4 - $W10YYMM"; Query = ""}
$Collection7 = @{Name = "Upgrade - Ring 5 - $W10YYMM"; Query = ""}
$Collection8 = @{Name = "Upgrade - Ring Pilot - $W10YYMM"; Query = ""}
$Collection9 = @{Name = "Upgrade - Ring PrePilot - $W10YYMM"; Query = ""}
$Collection10 = @{Name = "Upgrade - Targeted for $W10YYMM"; Query = ""}
$Collection11 = @{Name = "Upgrade - Ready for Windows 10 $W10YYMM"; Query = ""}
$Collection12 = @{Name = "ActiveClients - Targeted for $W10YYMM"; Query = ""}
$Collection13 = @{Name = "Upgrade Excluded for $W10YYMM"; Query = ""}
$Collection14 = @{Name = "Client - W10 - Ring 1"; Query = "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System where SMS_R_System.SID like "%0" OR SMS_R_System.SID like "%1""}
$Collection15 = @{Name = "Client - W10 - Ring 2"; Query = "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System where SMS_R_System.SID like "%2" OR SMS_R_System.SID like "%3""}
$Collection16 = @{Name = "Client - W10 - Ring 3"; Query = "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System where SMS_R_System.SID like "%4" OR SMS_R_System.SID like "%5""}
$Collection17 = @{Name = "Client - W10 - Ring 4"; Query = "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System where SMS_R_System.SID like "%6" OR SMS_R_System.SID like "%7""}
$Collection18 = @{Name = "Client - W10 - Ring 5"; Query = "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System where SMS_R_System.SID like "%8" OR SMS_R_System.SID like "%9""}
$Collection19 = @{Name = "Upgrade - Assess Succesfull for $W10YYMM"; Query = ""}
$Collection20 = @{Name = "Upgrade - Assess Failure for $W10YYMM"; Query = ""}



#Define possible limiting collections
$LimitingCollectionAllW10 = "Workstations | Windows 10"
$LimitingCollectionAllWorkActive = "Workstations | Active"
$LimitingCollectionAllSystems = "All Systems"

#Refresh Schedule
$Schedule = New-CMSchedule –RecurInterval Days –RecurCount 7

#Create Collection
#try{
New-CMDeviceCollection -Name $Collection1.Name -Comment "All Windows 10 Workstations $W10YYMM" -LimitingCollectionName $LimitingCollectionAllSystems -RefreshSchedule $Schedule -RefreshType 2 | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $Collection1.Name -QueryExpression $Collection1.Query -RuleName $Collection1.Name
Write-host *** Collection $Collection1.Name created ***

New-CMDeviceCollection -Name $Collection13.Name -Comment "Upgrade - Excluded for $W10YYMM" -LimitingCollectionName $LimitingCollectionAllW10 -RefreshSchedule $Schedule -RefreshType 2 | Out-Null
Add-CMDeviceCollectionExcludeMembershipRule -CollectionName $Collection13.Name -ExcludeCollectionName $Collection1.name
Write-host *** Collection $Collection13.Name created ***

New-CMDeviceCollection -Name $Collection2.Name -Comment "Upgrade - Compliant for $W10YYMM" -LimitingCollectionName $LimitingCollectionAllW10 -RefreshSchedule $Schedule -RefreshType 2 | Out-Null
Write-host *** Collection $Collection2.Name created ***

New-CMDeviceCollection -Name $Collection3.Name -Comment "Upgrade - Ring 1 - $W10YYMM" -LimitingCollectionName $LimitingCollectionAllW10 -RefreshSchedule $Schedule -RefreshType 2 | Out-Null
Add-CMDeviceCollectionExcludeMembershipRule -CollectionName $Collection3.Name -ExcludeCollectionName $Collection1.name
Write-host *** Collection $Collection3.Name created ***

New-CMDeviceCollection -Name $Collection4.Name -Comment "Upgrade - Ring 2 - $W10YYMM" -LimitingCollectionName $LimitingCollectionAllW10 -RefreshSchedule $Schedule -RefreshType 2 | Out-Null
Add-CMDeviceCollectionExcludeMembershipRule -CollectionName $Collection4.Name -ExcludeCollectionName $Collection1.name
Write-host *** Collection $Collection4.Name created ***

New-CMDeviceCollection -Name $Collection5.Name -Comment "Upgrade - Ring 3 - $W10YYMM" -LimitingCollectionName $LimitingCollectionAllW10 -RefreshSchedule $Schedule -RefreshType 2 | Out-Null
Add-CMDeviceCollectionExcludeMembershipRule -CollectionName $Collection5.Name -ExcludeCollectionName $Collection1.name
Write-host *** Collection $Collection5.Name created ***

New-CMDeviceCollection -Name $Collection6.Name -Comment "Upgrade - Ring 4 - $W10YYMM" -LimitingCollectionName $LimitingCollectionAllW10 -RefreshSchedule $Schedule -RefreshType 2 | Out-Null
Add-CMDeviceCollectionExcludeMembershipRule -CollectionName $Collection6.Name -ExcludeCollectionName $Collection1.name
Write-host *** Collection $Collection6.Name created ***

New-CMDeviceCollection -Name $Collection7.Name -Comment "Upgrade - Ring 5 - $W10YYMM" -LimitingCollectionName $LimitingCollectionAllW10 -RefreshSchedule $Schedule -RefreshType 2 | Out-Null
Add-CMDeviceCollectionExcludeMembershipRule -CollectionName $Collection7.Name -ExcludeCollectionName $Collection1.name
Write-host *** Collection $Collection7.Name created ***

New-CMDeviceCollection -Name $Collection8.Name -Comment "Upgrade - Ring Pilot - $W10YYMM" -LimitingCollectionName $LimitingCollectionAllW10  -RefreshType Manual | Out-Null
Add-CMDeviceCollectionExcludeMembershipRule -CollectionName $Collection8.Name -ExcludeCollectionName $Collection1.name
Write-host *** Collection $Collection8.Name created ***

New-CMDeviceCollection -Name $Collection9.Name -Comment "Upgrade - Ring PrePilot - $W10YYMM" -LimitingCollectionName $LimitingCollectionAllW10  -RefreshType Manual | Out-Null
Add-CMDeviceCollectionExcludeMembershipRule -CollectionName $Collection9.Name -ExcludeCollectionName $Collection1.name
Write-host *** Collection $Collection9.Name created ***

New-CMDeviceCollection -Name $Collection10.Name -Comment "Upgrade - Targeted for $W10YYMM" -LimitingCollectionName $LimitingCollectionAllW10 -RefreshSchedule $Schedule -RefreshType 2 | Out-Null
Add-CMDeviceCollectionIncludeMembershipRule -CollectionName $Collection10.Name -IncludeCollectionName $collection3.name
Add-CMDeviceCollectionIncludeMembershipRule -CollectionName $Collection10.Name -IncludeCollectionName $collection4.name
Add-CMDeviceCollectionIncludeMembershipRule -CollectionName $Collection10.Name -IncludeCollectionName $collection5.name
Add-CMDeviceCollectionIncludeMembershipRule -CollectionName $Collection10.Name -IncludeCollectionName $collection6.name
Add-CMDeviceCollectionIncludeMembershipRule -CollectionName $Collection10.Name -IncludeCollectionName $collection7.name
Add-CMDeviceCollectionIncludeMembershipRule -CollectionName $Collection10.Name -IncludeCollectionName $collection8.name
Add-CMDeviceCollectionIncludeMembershipRule -CollectionName $Collection10.Name -IncludeCollectionName $collection9.name
Write-host *** Collection $Collection10.Name created ***

New-CMDeviceCollection -Name $Collection11.Name -Comment "Upgrade - Ready for Windows 10 $W10YYMM" -LimitingCollectionName $LimitingCollectionAllW10 -RefreshSchedule $Schedule -RefreshType 2 | Out-Null
Add-CMDeviceCollectionExcludeMembershipRule -CollectionName $Collection11.Name -ExcludeCollectionName "CB - Low DiskSpace"
Add-CMDeviceCollectionExcludeMembershipRule -CollectionName $Collection11.Name -ExcludeCollectionName $Collection1.name
Add-CMDeviceCollectionExcludeMembershipRule -CollectionName $Collection11.Name -ExcludeCollectionName "Upgrade Excluded for $W10YYMM"
Add-CMDeviceCollectionExcludeMembershipRule -CollectionName $Collection11.Name -ExcludeCollectionName "Wrong MUI"
Add-CMDeviceCollectionIncludeMembershipRule -CollectionName $Collection11.Name -IncludeCollectionName $Collection2.name
Write-host *** Collection $Collection11.Name created ***

New-CMDeviceCollection -Name $Collection12.Name -Comment "ActiveClients - Targeted for $W10YYMM" -LimitingCollectionName $LimitingCollectionAllWorkActive -RefreshSchedule $Schedule -RefreshType 2 | Out-Null
Add-CMDeviceCollectionIncludeMembershipRule -CollectionName $Collection12.Name -IncludeCollectionName $Collection10.Name
Write-host *** Collection $Collection12.Name created ***

New-CMDeviceCollection -Name $Collection14.Name -Comment "Client - W10 - Ring 1" -LimitingCollectionName $LimitingCollectionAllW10 -RefreshSchedule $Schedule -RefreshType 2 | Out-Null
Add-CMDeviceCollectionIncludeMembershipRule -CollectionName $Collection14.Name -QueryExpression $Collection14.Query -RuleName $Collection14.Name
Write-host *** Collection $Collection14.Name created ***

New-CMDeviceCollection -Name $Collection15.Name -Comment "Client - W10 - Ring 2" -LimitingCollectionName $LimitingCollectionAllW10 -RefreshSchedule $Schedule -RefreshType 2 | Out-Null
Add-CMDeviceCollectionIncludeMembershipRule -CollectionName $Collection15.Name -QueryExpression $Collection15.Query -RuleName $Collection15.Name
Write-host *** Collection $Collection15.Name created ***

New-CMDeviceCollection -Name $Collection16.Name -Comment "Client - W10 - Ring 3" -LimitingCollectionName $LimitingCollectionAllW10 -RefreshSchedule $Schedule -RefreshType 2 | Out-Null
Add-CMDeviceCollectionIncludeMembershipRule -CollectionName $Collection16.Name -QueryExpression $Collection16.Query -RuleName $Collection16.Name
Write-host *** Collection $Collection16.Name created ***

New-CMDeviceCollection -Name $Collection17.Name -Comment "Client - W10 - Ring 4" -LimitingCollectionName $LimitingCollectionAllW10 -RefreshSchedule $Schedule -RefreshType 2 | Out-Null
Add-CMDeviceCollectionIncludeMembershipRule -CollectionName $Collection17.Name -QueryExpression $Collection17.Query -RuleName $Collection17.Name
Write-host *** Collection $Collection17.Name created ***

New-CMDeviceCollection -Name $Collection18.Name -Comment "Client - W10 - Ring 5" -LimitingCollectionName $LimitingCollectionAllW10 -RefreshSchedule $Schedule -RefreshType 2 | Out-Null
Add-CMDeviceCollectionIncludeMembershipRule -CollectionName $Collection18.Name -QueryExpression $Collection18.Query -RuleName $Collection18.Name
Write-host *** Collection $Collection18.Name created ***

New-CMDeviceCollection -Name $Collection19.Name -Comment "Upgrade - Assess Successfull for $W10YYMM" -LimitingCollectionName $LimitingCollectionAllW10  -RefreshType Manual | Out-Null
Add-CMDeviceCollectionExcludeMembershipRule -CollectionName $Collection19.Name -ExcludeCollectionName $Collection1.name
Write-host *** Collection $Collection19.Name created ***

New-CMDeviceCollection -Name $Collection20.Name -Comment "Upgrade - Assess Failure for $W10YYMM" -LimitingCollectionName $LimitingCollectionAllW10 -RefreshType  Manual | Out-Null
Add-CMDeviceCollectionExcludeMembershipRule -CollectionName $Collection20.Name -ExcludeCollectionName $Collection1.name
Write-host *** Collection $Collection20.Name created ***


#Move the collection to the right folder
#Operational
$FolderPath = $SiteCode.name+":\DeviceCollection\Operational"
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $Collection1.Name)


#UG - Upgrade - YYMM
$FolderPath = $SiteCode.name+":\DeviceCollection\Servicing Windows 10\Upgrade - $W10YYMM"
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $Collection2.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $Collection3.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $Collection4.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $Collection5.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $Collection6.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $Collection7.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $Collection8.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $Collection9.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $Collection10.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $Collection11.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $Collection12.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $Collection13.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $Collection19.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $Collection20.Name)

#Batches
$FolderPath = $SiteCode.name+":\DeviceCollection\Batches"
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $Collection14.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $Collection15.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $Collection16.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $Collection17.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $Collection18.Name)




