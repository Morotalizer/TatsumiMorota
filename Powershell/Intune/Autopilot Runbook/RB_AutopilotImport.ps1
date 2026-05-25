<#
.SYNOPSIS
    PowerShell script to import Autopilot device with managedid

.DESCRIPTION
    This script will be used in a Runbook authenticating with System Managed Identiy and send notification to teams channel.


.NOTES
    Version:        1.2
    Creation Date:  Nov 05, 2025
    Last Updated:   Dec 09, 2025
    Author:         Tatsumi Morota, Advania - Knowledge Factory
    Modified by:    Tatsumi Morota
    Notes: The foundation is made from https://mikemdm.de/2023/02/12/automate-autopilot-uploads-with-azure-automation-runbooks/
#>


param
(
[Parameter (Mandatory=$false)]
[object] $WebhookData
)

# Connect to Intune
$resourceURL = "https://graph.microsoft.com/" 
$response = [System.Text.Encoding]::Default.GetString((Invoke-WebRequest -UseBasicParsing -Uri "$($env:IDENTITY_ENDPOINT)?resource=$resourceURL" -Method 'GET' -Headers @{'X-IDENTITY-HEADER' = "$env:IDENTITY_HEADER"; 'Metadata' = 'True'}).RawContentStream.ToArray()) | ConvertFrom-Json 
#$script:authToken = $response.access_token 

$script:authToken = @{
    'Content-Type'  = 'application/json'
    'Authorization' = "Bearer " + $response.access_token
}


# --- Power Automate webhook config ---
# Prefer storing the URL in an Automation Variable 'AutopilotFlowWebhookUrl'
#   to avoid hardcoding secrets in the runbook.
try {
    $FlowWebhookUrl = Get-AutomationVariable -Name 'AutopilotFlowWebhookUrl'
} catch { $FlowWebhookUrl = $null }

if (-not $FlowWebhookUrl -or [string]::IsNullOrWhiteSpace($FlowWebhookUrl)) {
    # Fallback to the provided URL
    $FlowWebhookUrl = 'URL'
}
$script:FlowWebhookUrl = $FlowWebhookUrl

function Send-FlowWebhook {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string] $Status,          # success | error | skipped | pending
        [Parameter(Mandatory=$false)][string] $SerialNumber,
        [Parameter(Mandatory=$false)][string] $HardwareHash,
        [Parameter(Mandatory=$false)][string] $OrderIdentifier,
        [Parameter(Mandatory=$false)][string] $Action,         # SetGroupTag | ImportAutopilot | Sync | Lookup
        [Parameter(Mandatory=$false)][string] $AutopilotId,
        [Parameter(Mandatory=$false)][string] $ImportId,
        [Parameter(Mandatory=$false)][int]    $ErrorCode,
        [Parameter(Mandatory=$false)][string] $ErrorName,
        [Parameter(Mandatory=$false)][string] $Message
    )

    try {
        # Ensure TLS
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        $payload = @{
            serialNumber    = $SerialNumber
            hardwareHash    = $HardwareHash
            orderIdentifier = $OrderIdentifier
            status          = $Status
            action          = $Action
            autopilotId     = $AutopilotId
            importId        = $ImportId
            errorCode       = $ErrorCode
            errorName       = $ErrorName
            message         = $Message
            timestampUtc    = [DateTime]::UtcNow.ToString('o')
            runbookName     = $WebhookName
            source          = 'AzureAutomationRunbook'
        }

        $json = $payload | ConvertTo-Json -Depth 5
        $resp = Invoke-WebRequest -Uri $script:FlowWebhookUrl -Method POST -ContentType 'application/json' -Body $json -UseBasicParsing
        Write-Output "Flow webhook POST → HTTP $($resp.StatusCode)"
    }
    catch {
        Write-Warning "Flow webhook failed: $($_.Exception.Message)"
    }
}

####################################################Functions for Import####################################################

function Get-ErrorResponseBody {
    param([Parameter(Mandatory=$true)] $Exception)

    $body = $null
    try {
        if ($Exception -and $Exception.Response) {
            $stream = $Exception.Response.GetResponseStream()
            if ($stream) {
                $reader = New-Object System.IO.StreamReader($stream)
                $reader.BaseStream.Position = 0
                $reader.DiscardBufferedData()
                $body = $reader.ReadToEnd()
            }
        }
    } catch {
        # Swallow parsing issues, just return $null
    }
    return $body
}


Function Get-AutoPilotDeviceBySerial {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $Serial,
        [Parameter(Mandatory = $false)][switch] $Contains
    )

    $authToken = $script:authToken
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities"
    $containsMatches = @()

    try {
        $current = $uri
        do {
            $resp = Invoke-RestMethod -Uri $current -Headers $authToken -Method Get

            $batch = $resp.value
            if (-not ($batch -is [System.Collections.IEnumerable])) { $batch = ,$batch }

            if (-not $Contains) {
                $match = $batch | Where-Object { $_.serialNumber -eq $Serial } | Select-Object -First 1
                if ($match) { return $match }
            } else {
                $containsMatches += ($batch | Where-Object { $_.serialNumber -like "*$Serial*" })
            }

            $current = $resp.'@odata.nextLink'
        } while ($current)

        if ($Contains -and $containsMatches.Count -gt 0) {
            return ($containsMatches | Select-Object -First 1)
        }
        return $null
    }
    catch {
        $ex = $_.Exception
        $responseBody = Get-ErrorResponseBody -Exception $ex
        if ($responseBody) { Write-Output "Response content:`n$responseBody" }

        $statusCode = $null; $statusText = $null
        if ($ex.Response) { $statusCode = $ex.Response.StatusCode; $statusText = $ex.Response.StatusDescription }
        Write-Error "Autopilot list request failed with HTTP Status $statusCode $statusText — $($ex.Message)"
        return $null
    }
}


Function Set-AutoPilotDeviceGroupTag {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $Id,
        [Parameter(Mandatory = $true)][string] $GroupTag
    )

    $authToken = $script:authToken
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities/$Id/updateDeviceProperties"
    $body = @{ groupTag = $GroupTag } | ConvertTo-Json

    try {
        Invoke-RestMethod -Uri $uri -Headers $authToken -Method Post -Body $body -ContentType "application/json" | Out-Null
        Write-Output "Updated groupTag to '$GroupTag' on Autopilot device $Id."
    }
    catch {
        $ex = $_.Exception
        $responseBody = Get-ErrorResponseBody -Exception $ex
        if ($responseBody) { Write-Output "Response content:`n$responseBody" }

        $statusCode = $null; $statusText = $null
        if ($ex.Response) { $statusCode = $ex.Response.StatusCode; $statusText = $ex.Response.StatusDescription }
        Write-Error "Updating groupTag on $Id failed with HTTP Status $statusCode $statusText — $($ex.Message)"
        throw
    }
}


Function Get-AutoPilotDevice(){
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory=$false)] $id
    )
    
        # Defining Variables
        
        if ($id) {
            $uri = "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities/$id"
        }
        else {
            $uri = "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities/$Resource"
        }
        try {
            $response = Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get
            if ($id) {
                $response
            }
            else {
                $response.Value
            }
        }
        catch {
    
            $ex = $_.Exception
            $errorResponse = $ex.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd();
    
            Write-Output "Response content:`n$responseBody"
            Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    
            break
        }
    
    }
    

Function Get-AutoPilotImportedDevice(){
[cmdletbinding()]
param
(
    [Parameter(Mandatory=$false)] $id
)

      if ($id) {
        $uri = "https://graph.microsoft.com/v1.0/deviceManagement/importedWindowsAutopilotDeviceIdentities/$id"
    }
    else {
        $uri = "https://graph.microsoft.com/v1.0/deviceManagement/importedWindowsAutopilotDeviceIdentities/$Resource"
    }
   
        $response = Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get
        if ($id) {
            $response
        }
        else {
            $response.Value
        }
}

Function Add-AutoPilotImportedDevice(){
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory=$true)] $serialNumber,
        [Parameter(Mandatory=$true)] $hardwareIdentifier,
        [Parameter(Mandatory=$false)] $orderIdentifier
    )
    
        # Defining Variables
    
        #$uri = "https://graph.microsoft.com/v1.0/deviceManagement/importedWindowsAutopilotDeviceIdentities/$Resource"
        $uri = "https://graph.microsoft.com/v1.0/deviceManagement/importedWindowsAutopilotDeviceIdentities/"
        $json = @"
{
    "@odata.type": "#microsoft.graph.importedWindowsAutopilotDeviceIdentity",
    "groupTag": "$orderIdentifier",
    "serialNumber": "$serialNumber",
    "productKey": "",
    "hardwareIdentifier": "$hardwareIdentifier",
    "state": {
        "@odata.type": "microsoft.graph.importedWindowsAutopilotDeviceIdentityState",
        "deviceImportStatus": "pending",
        "deviceRegistrationId": "",
        "deviceErrorCode": 0,
        "deviceErrorName": ""
        }
}
"@

        try {
            $Response=(Invoke-RestMethod -Uri $uri -Headers $authToken -Method Post -Body $json -ContentType "application/json").ID
            return $Response
        }
        catch {
    
            $ex = $_.Exception
            $errorResponse = $ex.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd();
    
            Write-Output "Response content:`n$responseBody"
            Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    
            break
        }
    
    }

    
Function Remove-AutoPilotImportedDevice(){
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory=$true)] $id
    )

        $uri = "https://graph.microsoft.com/v1.0/deviceManagement/importedWindowsAutopilotDeviceIdentities/$id"

        try {
            Invoke-RestMethod -Uri $uri -Headers $authToken -Method Delete | Out-Null
        }
        catch {
    
            $ex = $_.Exception
            $errorResponse = $ex.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd();
    
            Write-Output "Response content:`n$responseBody"
            Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    
            break
        }
        
}

####################################################Import Main Function####################################################

Function Import-AutoPilotCSV(){
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true)] $Serial,
        [Parameter(Mandatory=$true)] $GroupTag,
        [Parameter(Mandatory=$true)] $hash
    )

    $ImportID = Add-AutoPilotImportedDevice -serialNumber $serial -hardwareIdentifier $hash -orderIdentifier $GroupTag
    
    # Notify import submitted
    Send-FlowWebhook -Status 'pending' `
        -SerialNumber $Serial `
        -HardwareHash $Hash `
        -OrderIdentifier $GroupTag `
        -Action 'ImportAutopilot' `
        -ImportId $ImportID `
        -Message 'Import request submitted to Autopilot'
        
    # ... existing polling loop ...

    # Generate some statistics for reporting...
    $global:successCount = 0
    $global:errorCount = 0
    $global:softErrorCount = 0
    $global:errorList = @{}
    $global:successList = @{}

    ForEach ($deviceStatus in $deviceStatuses) {

        # --- NEW: derive normalized status for Flow ---
        $st = ($deviceStatus.state.deviceImportStatus).ToLower()
        $statusForFlow = if ($st -in @('success','complete')) { 'success' }
                         elseif ($st -eq 'error') { 'error' }
                         else { 'pending' }

        if ($st -in @('success','complete')) {
            $global:successCount += 1
            $global:successList.Add($deviceStatus.serialNumber, $deviceStatus.state)
        } elseif ($st -eq 'error') {
            $global:errorCount += 1
            if ($($deviceStatus.state.deviceErrorCode) -eq 806) { $global:softErrorCount += 1 }
            $global:errorList.Add($deviceStatus.serialNumber, $deviceStatus.state)
        }

        # Display the statuses
        Write-Output "Serial number $($_.serialNumber): $($_.state.deviceImportStatus), $($_.state.deviceErrorCode), $($_.state.deviceErrorName)"

        # --- existing LA upload block ---
        $ImportedAutopilotDevice = New-Object System.Object
        $ImportedAutopilotDevice | Add-Member -MemberType NoteProperty -Name "SerialNumber" -Value $($_.serialNumber) -Force   
        $ImportedAutopilotDevice | Add-Member -MemberType NoteProperty -Name "Status" -Value $($_.state.deviceImportStatus) -Force   
        $ImportedAutopilotDevice | Add-Member -MemberType NoteProperty -Name "ErrorCode" -Value $($_.state.deviceErrorCode) -Force      
        $ImportedAutopilotDevice | Add-Member -MemberType NoteProperty -Name "ErrorName" -Value $($_.state.deviceErrorName) -Force  
        $AutopilotJson = $ImportedAutopilotDevice | ConvertTo-Json
        $ResponseLAUpload = Send-LogAnalyticsData -customerId $WorkspaceID -sharedKey $SharedKey -body ([System.Text.Encoding]::UTF8.GetBytes($AutopilotJson)) -logType $AutopilotLog -ErrorAction Stop
        Write-Output $ResponseLAUpload

        # --- NEW: send webhook to Power Automate ---
        Send-FlowWebhook -Status $statusForFlow `
            -SerialNumber $deviceStatus.serialNumber `
            -HardwareHash $hash `
            -OrderIdentifier $GroupTag `
            -Action 'ImportAutopilot' `
            -ImportId $ImportID `
            -ErrorCode ([int]$deviceStatus.state.deviceErrorCode) `
            -ErrorName $deviceStatus.state.deviceErrorName `
            -Message ("Import status: " + $deviceStatus.state.deviceImportStatus)
    }

    # Cleanup the imported device records
    $deviceStatuses | ForEach-Object {
        Remove-AutoPilotImportedDevice -id $_.id
    }
}

Function Invoke-AutopilotSync(){

    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotSettings/sync"
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $authToken -Method Post
        $response.Value
    }
    catch {

        $ex = $_.Exception
        $errorResponse = $ex.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($errorResponse)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();

        Write-Host "Response content:`n$responseBody" -f Red
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"

        # break
    }

}


####################################################Functions for Log Analytics###################################################
<#
Function New-Signature ($customerId, $sharedKey, $date, $contentLength, $method, $contentType, $resource) {
    $xHeaders = "x-ms-date:" + $date
    $stringToHash = $method + "`n" + $contentLength + "`n" + $contentType + "`n" + $xHeaders + "`n" + $resource



    $bytesToHash = [Text.Encoding]::UTF8.GetBytes($stringToHash)
    $keyBytes = [Convert]::FromBase64String($sharedKey)

    $sha256 = New-Object System.Security.Cryptography.HMACSHA256
    $sha256.Key = $keyBytes
    $calculatedHash = $sha256.ComputeHash($bytesToHash)
    $encodedHash = [Convert]::ToBase64String($calculatedHash)
    $authorization = 'SharedKey {0}:{1}' -f $customerId, $encodedHash
    Write-Host "StringToSign:`n$stringToHash"
    Write-Host "Authorization Header:`n$authorization"
    return $authorization
}#endfunction#>
Function New-Signature {
    param (
        [Parameter(Mandatory=$true)] [string] $customerId,
        [Parameter(Mandatory=$true)] [string] $sharedKey,
        [Parameter(Mandatory=$true)] [string] $date,
        [Parameter(Mandatory=$true)] [int] $contentLength,
        [Parameter(Mandatory=$true)] [string] $method,
        [Parameter(Mandatory=$true)] [string] $contentType,
        [Parameter(Mandatory=$true)] [string] $resource
    )

    # Build the canonical string according to Microsoft spec
    $xHeaders = "x-ms-date:" + $date
    $stringToSign = $method + "`n" + $contentLength + "`n" + $contentType + "`n" + $xHeaders + "`n" + $resource

    # Debug output (optional)
    Write-Host "StringToSign:`n$stringToSign"

    # Convert to bytes
    $bytesToHash = [Text.Encoding]::UTF8.GetBytes($stringToSign)
    $keyBytes = [Convert]::FromBase64String($sharedKey)

    # Compute HMAC-SHA256
    $sha256 = New-Object System.Security.Cryptography.HMACSHA256
    $sha256.Key = $keyBytes
    $calculatedHash = $sha256.ComputeHash($bytesToHash)
    $encodedHash = [Convert]::ToBase64String($calculatedHash)

    # Build Authorization header
    $authorization = "SharedKey ${customerId}:${encodedHash}"

    # Debug output (optional)
    Write-Host "Authorization Header:`n$authorization"

    return $authorization
}
Function Send-LogAnalyticsData($customerId, $sharedKey, $body, $logType) {
    $method = "POST"
    $contentType = "application/json"
    $resource = "/api/logs"
    $rfc1123date = [DateTime]::UtcNow.ToString("r")
    $contentLength = $body.Length
    $signature = New-Signature `
        -customerId $customerId `
        -sharedKey $sharedKey `
        -date $rfc1123date `
        -contentLength $contentLength `
        -method $method `
        -contentType $contentType `
        -resource $resource
    
    $uri = "https://" + $customerId + ".ods.opinsights.azure.com" + $resource + "?api-version=2016-04-01"
    
    #validate that payload data does not exceed limits
    if ($body.Length -gt (31.9 *1024*1024))
    {
        throw("Upload payload is too big and exceed the 32Mb limit for a single upload. Please reduce the payload size. Current payload size is: " + ($body.Length/1024/1024).ToString("#.#") + "Mb")
    }

    $payloadsize = ("Upload payload size is " + ($body.Length/1024).ToString("#.#") + "Kb ")
    
    $headers = @{
        "Authorization"        = $signature;
        "Log-Type"             = $logType;
        "x-ms-date"            = $rfc1123date;
        "time-generated-field" = $TimeStampField;
    }
    

    $response = Invoke-WebRequest -Uri $uri -Method $method -ContentType $contentType -Headers $headers -Body $body -UseBasicParsing 
    $statusmessage = "$($response.StatusCode) : $($payloadsize)"
    return $statusmessage 
}#endfunction


####################################################Connect to Ressources###################################################

$global:totalCount = 0
$AutopilotZTDID=""
$AutopilotMDM=""

if ($WebhookData) 
{
#Define WorkspaceID
$WorkspaceID = Get-AutomationVariable -Name 'WorkspaceID'
$SharedKey = Get-AutomationVariable -Name 'WSSharedKey'
#Define Log Analytics Workspace Subscription ID
$SubscriptionID = Get-AutomationVariable -Name 'LASubscriptionID'
#Define Autopilot Import Log Name
$AutopilotLog = "Autopilot_Import"
# DO NOT DELETE TimeStampField - IT WILL BREAK LA Injection
$TimeStampField = "" 
Connect-AzAccount -Identity -Subscription $SubscriptionID
	
# Collect properties of WebhookData
$WebhookName = $WebHookData.WebhookName
$WebhookHeaders = $WebHookData.RequestHeader
$WebhookBody = $WebHookData.RequestBody

$Input = (ConvertFrom-Json -InputObject $WebhookBody)

$SerialNumber = $Input.SerialNumber
$HardwareHash = $Input.HardwareHash
$OrderIdentifier = $Input.OrderIdentifier





	
####################################################Main logic Import###################################################

# Main logic Import

#	Import-AutoPilotCSV -Serial $SerialNumber -GroupTag $OrderIdentifier -Hash $HardwareHash
#    # Sync new devices to Intune
#    Write-output "Triggering Sync to Intune."
#    Invoke-AutopilotSync


    
####################################################Main logic Import###################################################
# Main logic: only set groupTag if device exists, otherwise import

# 1) Try to find existing Autopilot device by serial
$existingDevice = Get-AutoPilotDeviceBySerial -Serial $SerialNumber

if ($existingDevice) {
    Write-Output "Autopilot device exists (Id: $($existingDevice.id), Serial: $SerialNumber)."

    # Optional: skip if already set
    $currentTag = $existingDevice.groupTag
    if ($currentTag -and ($currentTag -eq $OrderIdentifier)) {
        Write-Output "groupTag is already '$OrderIdentifier'. No action needed."

        # Notify 'skipped' (no change)
        Send-FlowWebhook -Status 'skipped' -SerialNumber $SerialNumber -HardwareHash $HardwareHash `
            -OrderIdentifier $OrderIdentifier -Action 'SetGroupTag' -AutopilotId $existingDevice.id `
            -Message 'groupTag already set; no action'
    }
    else {
        try {
            Set-AutoPilotDeviceGroupTag -Id $existingDevice.id -GroupTag $OrderIdentifier
            Write-Output "Triggering Sync to Intune."
            Invoke-AutopilotSync

            # Notify success
            Send-FlowWebhook -Status 'success' -SerialNumber $SerialNumber -HardwareHash $HardwareHash `
                -OrderIdentifier $OrderIdentifier -Action 'SetGroupTag' -AutopilotId $existingDevice.id `
                -Message 'groupTag updated and sync triggered'
        }
        catch {
            # Notify error
            $ex = $_.Exception
            $statusCode = $null; $statusText = $null
            if ($ex.Response) { $statusCode = [int]$ex.Response.StatusCode; $statusText = $ex.Response.StatusDescription }
            Send-FlowWebhook -Status 'error' -SerialNumber $SerialNumber -HardwareHash $HardwareHash `
                -OrderIdentifier $OrderIdentifier -Action 'SetGroupTag' -AutopilotId $existingDevice.id `
                -ErrorCode $statusCode -ErrorName $statusText -Message $ex.Message
            throw
        }
    }
}
else {
    # 2) Not found → import once, with the desired tag
    Write-Output "Autopilot device not found for Serial '$SerialNumber'. Importing with groupTag '$OrderIdentifier'."
    Import-AutoPilotCSV -Serial $SerialNumber -GroupTag $OrderIdentifier -Hash $HardwareHash

    Write-Output "Triggering Sync to Intune."
    Invoke-AutopilotSync
    # Notify success
            Send-FlowWebhook -Status 'success' -SerialNumber $SerialNumber -HardwareHash $HardwareHash `
                -OrderIdentifier $OrderIdentifier -Action 'SetGroupTag' -AutopilotId $existingDevice.id `
                -Message 'groupTag updated and sync triggered'
}

}