<#
.SYNOPSIS
	This Script Reset the device and preprovisioning it to an Autopilot device and sets GroupTag

.DESCRIPTION
	The script is provided via Software Center. Execution via ConfigMgr is in SYSTEM context.
.PREREQUISITES
    * Automation Runbook setup with webhook to provision AutopilotDevice
.PARAMETER [none]
	This script does not take any parameters.
.EXAMPLE
	\serviceui.exe -process:explorer.exe "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -ExecutionPolicy Bypass -WindowStyle Hidden -NoLogo -NoProfile -File ".\Reset-MDMWipeDevice.ps1"
.NOTES
	Version: 0.1 2023/01/01 initial concept version
    Version: 1.0 2023/02/22 Added Webhook, new GUI

.LINK 
	.Author Tatsumi Morota 2023/03/03
    .Inspiration Jeroen Burgerhout & niall brady
#>

Function SetVarsBasedOnDPI{
if ($readDPI -ge 96 -and $readDPI -le 144){
write-output "setting font size to work with 100-150% DPI scaling..."
$global:system_font = "12"
$global:normal_font = "12"
$global:heading_font = "22"
$global:small_font = "9"
}

if ($readDPI -ge 145 -and $readDPI -le 192){
write-output "setting font size to work with 151-200% DPI scaling..."
$global:system_font = "10"
$global:normal_font = "10"
$global:heading_font = "20"
$global:small_font = "8"
}

if ($readDPI -ge 193){
write-output "DPI was over 200%, setting a default..."
$global:system_font = "10"
$global:normal_font = "9"
$global:heading_font = "19"
$global:small_font = "7"}}
Function CreateLogsDir{
If(!(test-path $LogsFolder))
{
 New-Item -ItemType Directory -Force -Path $LogsFolder
}

# "setting ACL permissions..."

$acl = Get-Acl $LogsFolder

$AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("everyone","ExecuteFile", "ContainerInherit,ObjectInherit", "None", "Allow")
$acl.SetAccessRule($AccessRule)
$AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("everyone","ReadData", "ContainerInherit,ObjectInherit", "None", "Allow")
$acl.addAccessRule($AccessRule)
$AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("everyone","ReadPermissions", "ContainerInherit,ObjectInherit", "None", "Allow")
$acl.addAccessRule($AccessRule)
$AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("everyone","ReadAttributes", "ContainerInherit,ObjectInherit", "None", "Allow")
$acl.addAccessRule($AccessRule)
$AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("everyone","ReadExtendedAttributes", "ContainerInherit,ObjectInherit", "None", "Allow")
$acl.addAccessRule($AccessRule)
$acl | Set-Acl $LogsFolder
}

Function CheckTotal{
 
   if ($total -eq 2) {$buttonOK.Enabled = $true
   }
else {$buttonOK.Enabled = $false
}
}

Function DecodeFile {
$Content = [System.Convert]::FromBase64String($Base64File)
write-output "creating $FileToBeDecoded"
try {
Set-Content -Path $FileToBeDecoded -Value $Content -Encoding Byte
write-output "successfully created $FileToBeDecoded"}
catch{write-output "failed to create $FileToBeDecoded"}
}

# set DPI related variables
$readDPI = Get-Itemproperty "HKCU:\Control Panel\Desktop" -name "LogPixels"
if ($readDPI) {write-output "The DPI value detected was: $readDPI"
SetVarsBasedOnDPI}
else
{write-output "DPI registry key was not found, will try based on model..."
write-output "Model detected: $model"}

if ($model -eq "HP EliteBook 830 G5") {$readDPI = 150
SetVarsBasedOnDPI}
else
{$readDPI = 100
SetVarsBasedOnDPI}

##*===============================================
##* VARIABLE DECLARATION
##*===============================================
$appTitle="Migrate Computer to the cloud"
$Global:LogsFolder = "C:\Windows\Temp\CloudMigration"
$Version = "1.1"
CreateLogsDir

$FileToBeDecoded="$LogsFolder\Cloudreset.png"
$Base64File = ""
DecodeFile $FileToBeDecoded,$Base64File
$CloudMigration= "$LogsFolder\Cloudreset.png"

##*===============================================
##* END VARIABLE DECLARATION
##*===============================================



##*===============================================
##* Main UI
##*===============================================

[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
[System.Windows.Forms.Application]::EnableVisualStyles()
 
#define a tooltip object via https://petri.com/add-popup-tips-powershell-winforms-script
$tooltip1 = New-Object System.Windows.Forms.ToolTip

$ShowHelp={ 
 Switch ($this.name) {
 "label1" {$tip = ""}
 "label2" {$tip = ""}
 "label3" {$tip = "This tool allows you to safely migrate your old computer to the cloud.`r*"}
 "label4" {$tip = ""}
 "radiobutton1" {$tip = "Please confirm before continuing."}
 "radiobutton2" {$tip = "Please confirm before continuing."}
 "radiobutton3" {$tip = "Please confirm before continuing."}
 }
 $tooltip1.SetToolTip($this,$tip)
} 
 

Add-Type -AssemblyName System.Windows.forms
$form=New-Object System.Windows.Forms.Form
$form.StartPosition = "CenterScreen"
$form.Topmost = $True
$form.Width=640
$form.Height=740
$form.FormBorderStyle = 'Fixed3D'
$form.MaximizeBox = $false
$form.Text = $appTitle
# This base64 string holds the bytes that make up the Waternet icon for a 32x32 pixel image
$iconBase64 = 'inser base 64 string here'

$iconBytes = [Convert]::FromBase64String($iconBase64)
# initialize a Memory stream holding the bytes
$stream          = [System.IO.MemoryStream]::new($iconBytes, 0, $iconBytes.Length)
$Form.Icon       = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))


$label2=New-Object System.Windows.Forms.Label
$label2.Text = "Version: $Version"
$label2.Location='555,695'
$label2.Width=100
$label2.Height=30
$label2.Font = [System.Drawing.Font]::new("Segoe UI", $small_font, [System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Pixel)
$label2.name="label2"
$label2.add_MouseHover($ShowHelp)
$form.Controls.Add($label2)

$pictureBoxIcon = new-object Windows.Forms.PictureBox 
$pictureBoxIcon.width=570
$pictureBoxIcon.height=400
$pictureBoxIcon.top=20
$pictureBoxIcon.left=30
$pictureBoxIcon.Sizemode='Zoom'
$file = (get-item $CloudMigration)
$pictureBoxIcon.Image=[System.Drawing.Image]::Fromfile($file)
$form.Controls.Add($pictureBoxIcon)

$label3=New-Object System.Windows.Forms.Label
$label3.Text = 'Please confirm before reset to CloudPC, Only OneDrive synced folder content will persist after migration!'
$label3.Location='40,420'
$label3.Width=600
$label3.Height=60
$label3.Font = [System.Drawing.Font]::new("Segoe UI", $heading_font, [System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Pixel)
$label3.name="label3"
$label3.add_MouseHover($ShowHelp)
$form.Controls.Add($label3)


#draw a box around the checkbox options
$AdditionalOptionsBox=New-Object System.Windows.Forms.Label
$x = $x_1st_Col + 130
$y = $y_1st_Col + 500
$AdditionalOptionsBox.Location="$x,$y"
$AdditionalOptionsBox.Width=400
$AdditionalOptionsBox.Height=100
$AdditionalOptionsBox.Font = [System.Drawing.Font]::new("SYSTEM", $heading_font, [System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Pixel)
$AdditionalOptionsBox.ForeColor = "Black"
$AdditionalOptionsBox.name="AdditionalOptionsBox"
$AdditionalOptionsBox.add_MouseHover($ShowHelp) 
$AdditionalOptionsBox.BorderStyle = 1 

$AdditionalOptionsText=New-Object System.Windows.Forms.Label
$x = $x_1st_Col + 130
$y = $y_1st_Col + 600
$AdditionalOptionsText.Location="$x,$y"
$AdditionalOptionsText.Text = "Note: Select both check boxes to continue." 
$AdditionalOptionsText.Width=240
$AdditionalOptionsText.Height=20
$AdditionalOptionsText.Font = [System.Drawing.Font]::new("SYSTEM", $system_font, [System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Pixel)
$AdditionalOptionsText.ForeColor = "Black"
$AdditionalOptionsText.name="AdditionalOptionsText"
$AdditionalOptionsText.add_MouseHover($ShowHelp) 
$AdditionalOptionsText.BorderStyle = 0
$form.Controls.Add($AdditionalOptionsText)

$CheckBox1 = New-Object system.windows.Forms.Checkbox
$x = $x_1st_Col +170 
$y = $y_1st_Col +520
$CheckBox1.location = new-object system.drawing.point($x,$y)
$CheckBox1.Text = "My files are in OneDrive"
$CheckBox1.AutoSize = $true
$CheckBox1.Width = 104
$CheckBox1.Height = 14
$CheckBox1.Font = [System.Drawing.Font]::new("SYSTEM", $heading_font, [System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Pixel)
$CheckBox1.name="CheckBox1"
$CheckBox1.add_MouseHover($ShowHelp)
$CheckBox1.Checked = $false
$CheckBox1.add_click({
if ($CheckBox1.Checked){
$global:total=$total + 1
CheckTotal}
   else
{
$global:total=$total - 1
CheckTotal} 
   })
$Form.controls.Add($CheckBox1)

$CheckBox2 = New-Object system.windows.Forms.Checkbox
$x = $x_1st_Col +170
$y = $y_1st_Col +560
$CheckBox2.location = new-object system.drawing.point($x,$y)
$CheckBox2.Text = "I am ready to migrate to Cloud"
$CheckBox2.AutoSize = $true
$CheckBox2.Width = 104
$CheckBox2.Height = 14
$CheckBox2.Font = [System.Drawing.Font]::new("SYSTEM", $heading_font, [System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Pixel)
$CheckBox2.name="CheckBox2"
$CheckBox2.add_MouseHover($ShowHelp)
$CheckBox2.Checked = $false
$CheckBox2.add_click({
if ($CheckBox2.Checked){
$global:total=$total + 1
CheckTotal}
   else
{$global:total=$total - 1
CheckTotal} 
   })
$Form.controls.Add($CheckBox2)

$form.Controls.Add($AdditionalOptionsBox)

$buttonOK=New-Object System.Windows.Forms.Button
$buttonOK.Location='100,670'
$buttonOK.Width=120
$buttonOK.Font = [System.Drawing.Font]::new("SYSTEM", $heading_font, [System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Pixel)  
$buttonOK.Text='OK'
CheckTotal
$form.Controls.Add($buttonOK)

$buttonCancel=New-Object System.Windows.Forms.Button
$buttonCancel.Location='470,670'
$buttonCancel.Width=120
$buttonCancel.Font = [System.Drawing.Font]::new("SYSTEM", $heading_font, [System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Pixel)  
$buttonCancel.Text='Cancel'
$form.Controls.Add($buttonCancel)


$buttonOK.Add_Click(
{$appTitle="Please read this message carefully"
  $WarningMessage = 
 "
You are about to migrate your device to cloud management. The migration will take about 60 minutes, your computer will be restarted and you will be required to login during this process. After the migration process is complete, the computer will be accessible but additional policies may take up to 1 hour to fully process.`r
Click <OK> to migrate to the cloud, otherwise click <Cancel> to exit out of this process."		

######################### OK Cancel reset stuff here ##########################
$option = "OKCANCEL"
$FinalChoice = [System.Windows.Forms.MessageBox]::Show("$WarningMessage",$appTitle,$option,48)
switch  ($FinalChoice) {

  'OK' {
    write-output "The user chose OK to the Migrate warning, starting the actual Migration process now..."
    $form.Close()
    $stream.Dispose()
    $form.Dispose()
        # And this is where the magic happens
        
        #Azure Automation Webhook URI
        #Dustin.com
        $uri = "webhook url"
        #Get Hardware Hash
        $hwid = ((Get-WMIObject -Namespace root/cimv2/mdm/dmmap -Class MDM_DevDetail_Ext01 -Filter "InstanceID='Ext' AND ParentID='./DevDetail'").DeviceHardwareData)
        #Get SerialNumber
        $ser = (Get-WmiObject win32_bios).SerialNumber
        #Use Computername if SerialNumber is empty
        if([string]::IsNullOrWhiteSpace($ser)) { $ser = $env:computername}
        $orderIdentifier = "SE_UD_STANDARD"

        #Create object with the required parameters
        $devdata  = @{ SerialNumber=$ser;HardwareHash=$hwid;OrderIdentifier=$orderIdentifier}
        $body = ConvertTo-Json -InputObject $devdata 

        #Send request to Webhook
        $response = Invoke-RestMethod -Method Post -Uri $uri -Body $body
        $response.JobIds 

         #The Reset Devicebegins here
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

  'Cancel' {
    write-output "The user chose CANCEL to the Secure Wipe warning, aborting migration process..."
    $form.Close()
    $stream.Dispose()
    $form.Dispose()
    #exit
    
    }

}
######################### OK Cancel reset stuff here ##########################

write-output "We shouldn't be here...exiting buttonOK function..."

}
)


$buttonCancel.Add_Click(
	{
        # Dispose the form and its controls. Skip, if you want to redisplay the form later.
        $form.Close()
        $stream.Dispose()
        $form.Dispose()
        #exit
	}
)

# show the UI
 try
 {# put form in front of other windows
    $form.TopMost = $True
    $form.ShowDialog()
 #[system.windows.forms.application]::run($form) 
  }
 catch
 {
 $oops = $_.Exception.Message  
 write-output "An error occurred when trying to display the form, here is the error: $oops"  
 }
   
 write-output "Exiting the 'migrate my pc' script."