# Hämtar alla medlemmar ur valfri AD-grupp och exporterar till Excel-fil.
# Ändra  "Company Sverige Chefer" tillvalfri AD-grupp.
# Lägg till -Recursive för att få med undergrupper.
get-adgroupmember -identity "Company Sverige Chefer" | select "Name" | Export-csv -path C:\Scripts\AD-grupp.csv -NoTypeInformation

#Tar bort 2 första raderna
#$file = "c:\scripts\AD-grupp.csv"
#$newfile = "c:\scripts\AD-Grupp1.csv"

# Tar bara bort - och inte det efter.....
$oldfile = "C:\scripts\AD-grupp.csv"
$newfile = "C:\scripts\AD-grupp1.csv"
$text = (Get-Content -Path $oldfile -ReadCount 0) -join "`n"
$text -replace '-*', '' | Set-Content -Path $newfile
