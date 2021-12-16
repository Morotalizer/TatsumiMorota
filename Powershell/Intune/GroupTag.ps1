Install-PackageProvider -Name Nuget -confirm:$false -Force
Install-Module Microsoft.Graph.Intune -confirm:$false -Force
Install-Module WindowsAutopilotIntune -confirm:$false -Force
Import-Module Microsoft.Graph.Intune
Import-Module WindowsAutopilotIntune

$tenant = "morotalized.onmicrosoft.com"
$authority = “https://login.windows.net/$tenant"
$clientId = "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXX"
$clientSecret = "client_app_secret"
$groupTag = "Fest"
$requestBody=
@"
    {
        groupTag: `"$groupTag`",
    }
"@

Update-MSGraphEnvironment -AppId $clientId -Quiet
Update-MSGraphEnvironment -AuthUrl $authority -Quiet
Connect-MSGraph -ClientSecret $ClientSecret -Quiet


#$deviceSerialNumber = Get-WmiObject win32_bios | select Serialnumber
$SerialNumber = Get-WmiObject win32_bios | select Serialnumber   

$SerialNumberEncoded = [Uri]::EscapeDataString($SerialNumber.Serialnumber)
$ResourceURI = "deviceManagement/windowsAutopilotDeviceIdentities?`$filter=contains(serialNumber,'$($SerialNumberEncoded)')"
#$GraphResponse = (Invoke-MSGraphOperation -Get -APIVersion "Beta" -Resource $ResourceURI).value
$ID = Invoke-MSGraphRequest -HttpMethod GET -Url "$ResourceURI"
$IDValue = $id.value.id
$URL = "deviceManagement/windowsAutopilotDeviceIdentities/$IDValue/UpdateDeviceProperties" 
Invoke-MSGraphRequest -HttpMethod POST -Content $requestBody -Url $URL 
