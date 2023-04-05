# Reset and wipe an Intune managed Windows 10/11 device
# Created by Jeroen Burgerhout (@BurgerhoutJ)

# Create a tag file just so Intune knows this was installed
$Company = ""
if (-not (Test-Path "$($env:ProgramData)\$Company\ResetMDMDevice"))
{
    Mkdir "$($env:ProgramData)\$Company\ResetMDMDevice"
}
Set-Content -Path "$($env:ProgramData)\$Company\ResetMDMDevice\Reset-MDMWipeDevice_SE_PRD.ps1.tag" -Value "Installed"

# Show a messagebox where the enduser can accept or decline the reset
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Create the form.
$form  = New-Object system.Windows.Forms.Form
$form.Width   = "400"
$form.Height  = "375"
$form.Text = "Reset for Cloud Management"
$form.AutoSize = $true
$form.Topmost = $true
$form.StartPosition = "CenterScreen"
$form.BackColor = "Black"
$form.ForeColor = "RED"
# This base64 string holds the bytes that make up the Waternet icon for a 32x32 pixel image
$iconBase64      = '[insert base64 string here without brackets]'
$iconBytes       = [Convert]::FromBase64String($iconBase64)
# initialize a Memory stream holding the bytes
$stream          = [System.IO.MemoryStream]::new($iconBytes, 0, $iconBytes.Length)
$Form.Icon       = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))

$TextBox = New-Object System.Windows.Forms.label
$TextBox.Location = New-Object System.Drawing.Size(10,10)
$TextBox.Size = New-Object System.Drawing.Size(400,250)
$TextBox.Margin = New-Object System.Windows.Forms.Padding(10, 10, 10, 10)
$TextBox.BackColor = "Transparent"
$TextBox.Font = "Segoe UI, 10"
$TextBox.Text = "CAUTION!
Starting this app will reinstall your device. (approx time: 1 hour)

ATTENTION!
So only proceed if your important files are stored on OneDrive or SharePoint and the synchronization is complete! The hard drive on your device will be totally erased during this action.

IMPORTANT!
During this entire action, the laptop must be powered and connected to internet. So make sure your device is directly connected to a charger or docking station before proceeding!

Click CONTINUE to start the reset or CANCEL to stop it."

$form.Controls.Add($TextBox)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,300)
$okButton.Size = New-Object System.Drawing.Size(100,23)
$okButton.Text = 'CONTINUE'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(250,300)
$cancelButton.Size = New-Object System.Drawing.Size(100,23)
$cancelButton.Text = 'CANCEL'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::No
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$Result = $form.ShowDialog()

switch  ($Result) {

    'Yes' {

        # Dispose the form and its controls. Skip, if you want to redisplay the form later.
        $form.Close()
        $stream.Dispose()
        $form.Dispose()
        
        # And this is where the magic happens
        
        #Azure Automation Webhook URI
        $uri = ""
        #Get Hardware Hash
        $hwid = ((Get-WMIObject -Namespace root/cimv2/mdm/dmmap -Class MDM_DevDetail_Ext01 -Filter "InstanceID='Ext' AND ParentID='./DevDetail'").DeviceHardwareData)
        #Get SerialNumber
        $ser = (Get-WmiObject win32_bios).SerialNumber
        #Use Computername if SerialNumber is empty
        if([string]::IsNullOrWhiteSpace($ser)) { $ser = $env:computername}
        $orderIdentifier = "SE_UD_STANDARD_PRD"

        #Create object with the required parameters
        $devdata  = @{ SerialNumber=$ser;HardwareHash=$hwid;OrderIdentifier=$orderIdentifier}
        $body = ConvertTo-Json -InputObject $devdata 

        #Send request to Webhook
        $response = Invoke-RestMethod -Method Post -Uri $uri -Body $body
        $response.JobIds 


        $namespaceName = "root\cimv2\mdm\dmmap"
        $className = "MDM_RemoteWipe"
        $methodName = "doWipeProtectedMethod" #change this to doWipeMethod if you run this app on Surface devices

        $session = New-CimSession

        $params = New-Object Microsoft.Management.Infrastructure.CimMethodParametersCollection
        $param = [Microsoft.Management.Infrastructure.CimMethodParameter]::Create("param", "", "String", "In")
        $params.Add($param)

        $instance = Get-CimInstance -Namespace $namespaceName -ClassName $className -Filter "ParentID='./Vendor/MSFT' and InstanceID='RemoteWipe'"
        $session.InvokeMethod($namespaceName, $instance, $methodName, $params)
        
    }
    'No' {
        
        # Dispose the form and its controls. Skip, if you want to redisplay the form later.
        $form.Close()
        $stream.Dispose()
        $form.Dispose()
        #exit
    }
}