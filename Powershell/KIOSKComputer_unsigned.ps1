<#PSScriptInfo

.VERSION 
    3.3

.GUID 
    
.AUTHOR 
    Tatsumi Morota

.COMPANYNAME 
    RTS AB

.COPYRIGHT 
    
.TAGS 
    KIOSK Computer Shell Management

.LICENSEURI 

.PROJECTURI 

.ICONURI 

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES 

.RELEASENOTES

#>

<#
.SYNOPSIS
    This script is used on Windows 10 1511.
    The script can automatically add/remove shell on a Kiosk computer
        All user input is made in a simple Windows Form GUI.
        
    The following requirements must be installed on the computer where the script is run:
    - PowerShell v4.0 or higher executionpolicy unrestricted
    - .Net Framework
    
    NOTE: The script has only been fully tested and verified with PowerShell v5.0

.DESCRIPTION
    Use this script to add or remove shell on a kiosk computer, you can also trigger a reboot if you have moved the comptuer to another OU
    

    CHANGELOG
        Version 3.3 2017-03-07
            - Added dropdown list instead of txtbox
            - Added search from AD OU to populate dropdownlist
            - Added Error handling/message
            - Added -Asjob on invoke command
        Version 3.2 2017-03-06
            - Added error handling WINRM service not running
        Version 3.1 2017-03-03
            - Moved functions to top (to work outside ISE)
            - Added check for AD module
            - Added signing certificate
        Version 3.0 (Production Release) (2017-03-02)
            - Added computer exist check
            - Added computer is live check
            - Added Progressbar
            - Added Console feedback
        Version 2.0 (Pre-Release) (2017-02-27)
            - Initial Release

.NOTES
    Author: Tatsumi Morota

.LINK
    Author's blog:
    Author's workplace: http://www.rtsab.com

.EXAMPLE
    Run the script with administrative powershellconsole
        
    .\PersonalComputer.ps1

#>

# We set the RunInstallAs32Bit parameter to a global variable, because it can be modified in multiple functions
#[bool]$global:RunInstallAs32Bit = $RunInstallAs32Bit

#~~< Check if AD module exist >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if (Get-Module -ListAvailable -Name Microsoft.PowerShell.Management) {
    Write-Host "Powershell Module exists" -ForegroundColor Green
} else {
    Write-Host "Powershell Module does not exist, please install" -ForegroundColor Red
    exit
}

#~~< Function ButtonClick >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function ButtonOKClick
{
$OkToProceed = Validate-Form

    if ($OkToProceed)
    {
        $ProgressBar1.Visible = $true
	    $ProgressBar1.PerformStep()

        if ($radioReboot.Checked)
        {
            try
            {
                $reachable = $true
                Invoke-Command -ComputerName $listComputer.SelectedItem.ToString() -ErrorAction Stop -ScriptBlock {powershell -file "C:\Shell\reboot.ps1"} 
            }
            catch
            { 
                $reachable = $false
                write-host $textComputer.Text "ERROR: WINRM service is not running on remote computer, please start it" -ForegroundColor RED
            } 
            if ($reachable)
            {
                Invoke-Command -ComputerName $listComputer.SelectedItem.ToString() -ScriptBlock {powershell -file "C:\Shell\reboot.ps1"} -AsJob
                $ProgressBar1.PerformStep()
                Write-Host "Rebooting Computer " -NoNewline; Write-Host $textComputer.Text -ForegroundColor Green
            }
        }      
        if ($radioRemove.Checked)
        {
            try
            {
                $reachable = $true
                Invoke-Command -ComputerName $listComputer.SelectedItem.ToString() -ScriptBlock { powershell -file "C:\Shell\remove.ps1"}
            }
            catch
            {
                $reachable = $false
                write-host $textComputer.Text "ERROR: WINRM service is not running on remote computer, please start it" -ForegroundColor RED
            }
            if ($reachable)
            {
                Invoke-Command -ComputerName $listComputer.SelectedItem.ToString() -ScriptBlock { powershell -file "C:\Shell\remove.ps1"} -AsJob
                $ProgressBar1.PerformStep()
                Write-Host "Shell being deactivated and rebooting Computer " -NoNewline; Write-Host $textComputer.Text -ForegroundColor Green    
            }
        }
        if ($radioAdd.Checked)
        {
            try
            {
                $reachable = $true
                Invoke-Command -ComputerName $listComputer.SelectedItem.ToString() -ScriptBlock { powershell -File "C:\Shell\new.ps1"}
            }
            catch
            {
                $reachable = $false
                write-host $textComputer.Text "ERROR: WINRM service is not running on remote computer, please start it" -ForegroundColor RED
            }
            if ($reachable)
            {
                Invoke-Command -ComputerName $listComputer.SelectedItem.ToString() -ScriptBlock { powershell -File "C:\Shell\new.ps1"} -AsJob
                $ProgressBar1.PerformStep()
                Write-Host "Shell being activated and rebooting Computer " -NoNewline; Write-Host $textComputer.Text -ForegroundColor Green
            }
        }
        # Clear the form
		if ($reachable)
        {
            Write-Host "Done!`n" -ForegroundColor Green
            $ProgressBar1.Step = 2
		    $ProgressBar1.PerformStep()
		    Start-Sleep 5
		    Load-Form
        }
        else
        {
            Write-Host "Correct error then try again!`n" -ForegroundColor RED
            $ProgressBar1.Step = 2
		    $ProgressBar1.PerformStep()
		    Start-Sleep 5
		    Load-Form
        }
    }
}

#~~< Function Validate-Form >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Validate-Form
{
    $OkToProceed = $true
    #$ADComputer = $textComputer.Text
    $ADComputer = $listComputer.SelectedItem.ToString() #populate selected item 

    # Clear the progress bar
    $ProgressBar1.Visible = $false
	$ProgressBar1.Minimum = 1
	$ProgressBar1.Maximum = 2
	$ProgressBar1.Value = 1
	$ProgressBar1.Step = 1
    # Clear the error providers
    $ErrorProviderADComputer.Clear()
    $ErrorProviderADComputerOn.Clear()

    # Check if the computername exist
    if ($ADComputer.Length -gt 0 -and (Check-ADComputerExist $ADComputer) -eq $false)
    {
    	$OkToProceed = $false	
        $ErrorProviderADComputer.SetError($listComputer, "An AD Computer named $ADComputer doesn't exists. Please type a correct computer name.")
    }
        #Check if computer is Powered ON
    Else
    {
        If ($ADComputer.Length -gt 0 -and (Check-ADComputerOn $ADComputer) -eq $false)
        {   
    	    $OkToProceed = $false	
            $ErrorProviderADComputerOn.SetError($listComputer, "Computer $ADComputer seems to be powered OFF or is unreachable.")
        }
    }
    
# If we're ok to proceed, return True
    if ($OkToProceed)
    {
        Return $true
    }
    else
    {
        Return $false
    }
}
#~~< Function Check-ADComputerExist >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Check-ADComputerExist
{
	param(
		[Parameter(
		Position = 0)]
		$ADComputer
	)
	
    $ComputerExist = Get-ADComputer $ADComputer
	if ($ComputerExist)
	{
		return $true
	}
    else
    {
        return $false
    }
}
#~~< Function Check-ADComputerOn >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Check-ADComputerOn
{
	param(
		[Parameter(
		Position = 0)]
		$ADComputer
	)
	
    $ComputerOn = Test-Connection -ComputerName $ADComputer -Count 1 -Quiet
	if ($ComputerOn -eq  $true)
	{
		return $true
	}
    else
    {
        return $false
    }
}
#~~< Function Load-Form >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Load-Form
{
	$ProgressBar1.Visible = $false
	$ProgressBar1.Minimum = 1
	$ProgressBar1.Maximum = 2
	$ProgressBar1.Value = 1
	$ProgressBar1.Step = 1
}


Add-Type -AssemblyName System.Windows.Forms
#~~< Form >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Form = New-Object system.Windows.Forms.Form
$Form.Text = "Personal Computer Maintenace"
$Form.TopMost = $true
$Form.Width = 527
$Form.Height = 290
$Form.MaximizeBox = $False
$Form.MinimizeBox = $False
$Form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
#~~< RadioReboot >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$radioReboot = New-Object system.windows.Forms.RadioButton
$radioReboot.Text = "Reboot"
$radioReboot.AutoSize = $true
$radioReboot.Width = 104
$radioReboot.Height = 20
$radioReboot.location = new-object system.drawing.point(87,106)
$radioReboot.Font = "Microsoft Sans Serif,10"
$radioReboot.Checked = $true
$Form.controls.Add($radioReboot)
#~~< RadioRemove >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$radioRemove = New-Object system.windows.Forms.RadioButton
$radioRemove.Text = "Remove Shell and Reboot"
$radioRemove.AutoSize = $true
$radioRemove.Width = 104
$radioRemove.Height = 20
$radioRemove.location = new-object system.drawing.point(87,140)
$radioRemove.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($radioRemove)
#~~< RadioADD >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$radioAdd = New-Object system.windows.Forms.RadioButton
$radioAdd.Text = "Add Shell and Reboot"
$radioAdd.AutoSize = $true
$radioAdd.Width = 104
$radioAdd.Height = 20
$radioAdd.location = new-object system.drawing.point(86,175)
$radioAdd.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($radioAdd)
#~~< ListComputer >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ListComputer = New-Object system.windows.Forms.ComboBox
$listComputer.Width = 100
$listComputer.Height = 30
$listComputer.location = new-object system.drawing.point(190,72)
$Form.controls.Add($listComputer)
$PMComputerLists = Get-ADComputer -Filter * -SearchBase "OU=PM-Datorer_Win10,OU=Datorer,DC=[DOMAIN],DC=net" | Select-Object -ExpandProperty Name
    foreach ($PMComputerList in $PMComputerLists) {
                      $listComputer.Items.Add($PMComputerList)
                              } #end foreach
#~~< LabelComputer >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$labelComputer = New-Object system.windows.Forms.Label
$labelComputer.Text = "Computername"
$labelComputer.AutoSize = $true
$labelComputer.Width = 25
$labelComputer.Height = 10
$labelComputer.location = new-object system.drawing.point(86,73)
$labelComputer.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($labelComputer)
#~~< ButtonOK >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$buttonOK = New-Object system.windows.Forms.Button
$buttonOK.Text = "OK"
$buttonOK.Width = 60
$buttonOK.Height = 30
$buttonOK.location = new-object system.drawing.point(344,206)
$buttonOK.Font = "Microsoft Sans Serif,10"
$buttonOK.add_Click({ButtonOKClick})
$Form.controls.Add($buttonOK)
#~~< ProgressBar1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ProgressBar1 = New-Object System.Windows.Forms.ProgressBar
$ProgressBar1.Location = New-Object System.Drawing.Point(10, 15)
$ProgressBar1.Size = New-Object System.Drawing.Size(480, 22)
$ProgressBar1.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$ProgressBar1.Text = ""
$Form.Controls.Add($ProgressBar1)
#~~< ErrorProviderADComputer >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ErrorProviderADComputer = New-Object System.Windows.Forms.ErrorProvider
$ErrorProviderADComputer.BlinkStyle = [System.Windows.Forms.ErrorBlinkStyle]::NeverBlink
$ErrorProviderADComputer.ContainerControl = $Form
#~~< ErrorProviderADComputerOn >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ErrorProviderADComputerON = New-Object System.Windows.Forms.ErrorProvider
$ErrorProviderADComputerON.BlinkStyle = [System.Windows.Forms.ErrorBlinkStyle]::NeverBlink
$ErrorProviderADComputerON.ContainerControl = $Form

[void]$Form.ShowDialog()
$Form.Dispose()