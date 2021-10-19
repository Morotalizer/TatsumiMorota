# Alla dator/serverobjekt i AD
#Get-ADComputer -Filter * -Property * | Select-Object Name,LastLogonDate,OperatingSystemVersion,ipv4Address | Export-CSV ComputerLastLogonDate_Hela_AD.csv -NoTypeInformation -Encoding UTF8

# Alla i specifikt OU
#Get-ADComputer -searchbase "OU=Windows10,OU=Arbetsstationer,OU=Datorer,DC=Company,DC=net" -Filter * -Property * | Select-Object Name,LastLogonDate,OperatingSystemVersion,ipv4Address | Export-CSV Alla_Datorer_i_OU.csv -NoTypeInformation -Encoding UTF8


# Alla inaktiva datorer i ett OU
#Get-ADComputer -searchbase "OU=Windows10,OU=Arbetsstationer,OU=Datorer,DC=Company,DC=net" -Filter {(Enabled -eq $False)} -ResultPageSize 2000 -ResultSetSize $null -Properties Name,LastLogonDate,OperatingSystemversion | Export-CSV Alla_inaktiva_Datorer_i_OU.csv -NoTypeInformation -Encoding UTF8

