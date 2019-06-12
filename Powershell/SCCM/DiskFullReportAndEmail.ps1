<#
.SYNOPSIS
	Get folder sizes in specified tree.  
.DESCRIPTION
	Script creates an HTML report with owner information, when created, 
	when last updated and folder size.  By default script will only do 1
    level of folders.  Use Recurse to do all sub-folders.
	
	Update the PARAM section to match your environment.
.PARAMETER Paths
	Specify the path(s) you wish to report on.  Specify an array of paths for
    mulitple folders to be processed, or pipe folders in.
.PARAMETER ReportPath
	Specify where you want the HTML report to be saved
.PARAMETER Sort
    Specify which column you want the script to sort by.  
    
    Valid colums are:
        Folder                  Sort by folder name
        Size                    Sort by folder size, largest to smallest
        Created                 Sort by the Created On column
        Changed                 Sort by the Last Updated column
        Owner                   Sort by Owner
.PARAMETER Descending
    Switch to control how you want things sorted.  By default the script
    will sort in an ascending order.  Use this switch to reverse that.
.PARAMETER Recurse
    Report on all sub-folders
.EXAMPLE
	.\Get-FolderSizes
    
	Run script and use defaults
.EXAMPLE
	.\Get-FolderSizes -Path "c:\Windows" -ReportPath "c:\Scripts"
    
	Run the script and report on all folers in C:\Windows.  Save the
	HTML report in C:\Scripts
.EXAMPLE
	.\Get-FolderSizes -Path "c:\Windows" -ReportPath "c:\Scripts" -Recurse
    
	Run the script and report on all folers in C:\Windows.  Save the
	HTML report in C:\Scripts.  Report on all sub-folders.
.OUTPUTS
	FolderSizes.HTML in specified Report Path
.NOTES
	Author:         Martin Pugh
	Twitter:        @thesurlyadm1n
	Spiceworks:     Martin9700
	Blog:           www.thesurlyadmin.com
	Modified by:    Tatsumi Morota
	Changelog:
        1.7         Corrected Filesize sorting, added multiple path, email report to user
        1.6         Added some error trapping to test if the path provided is valid.
        1.5         Added ability to process multiple paths, both from array input as well as
                    from the pipeline.  Added verbose logging.  Also switched away from using
                    the COM Object.  While the COM object is much faster, it will return 0 if
                    it has any problem in the folder structure (such as a typical access denied
                    when running against C:\Windows).  I figure SOME result is better than none.
                    Tightened up the final report so it looks a little better.
        1.41        @SPadminWV found a bug in the Total Size reporting.
        1.4         Add Sort and descending parameter
		1.3         Added Recurse parameter, default behavior is to now do 1 level of folders,
                    recurse will do all sub-folders.
        1.2         Added function to make the rows in the table alternating colors
		1.1         Updated to use COM Object Scripting.FileSystemObject which
		            should increase performance.  Inspired by MS Scripting Guy
					Ed Wilson.
		1.0         Initial Release
.LINK
	http://community.spiceworks.com/scripts/show/1738-get-foldersizes
.LINK
	http://community.spiceworks.com/topic/286820-how-to-export-list-all-folders-from-drive-the-list-should-include
.LINK
	http://blogs.technet.com/b/heyscriptingguy/archive/2013/01/05/weekend-scripter-sorting-folders-by-size.aspx
#>	
#requires -Version 3.0

[CmdletBinding()]
Param (
    [Parameter(ValueFromPipeline)]
	#[string[]]$Paths = (Join-Path ($UserFolder.FullName) 'OneDrive - Ragn-Sells'),
    #[string[]]$Paths = ($UserFolder.FullName),
	[string]$ReportPath = "c:\_servicedesk\Logs\",
    [ValidateSet("Folder","Folders","Size","RawSize","Created","Changed","Owner")]
    [string]$Sort = "RawSize",
    [switch]$Descending,
    [switch]$Recurse
)

Begin {
    Function AddObject {
    	Param ( 
    		$FileObject
    	)
        $RawSize = (Get-ChildItem $FileObject.FullName -Recurse | Measure-Object Length -Sum).Sum

    	If ($RawSize)
    	{	$Size = CalculateSize $RawSize
    	}
    	Else
    	{	$Size = "0.00 MB"
    	}
    	$Object = New-Object PSObject -Property @{
    		'Folder Name' = $FileObject.FullName
    		'Created on' = $FileObject.CreationTime
    		'Last Updated' = $FileObject.LastWriteTime
    		Size = $Size
    		Owner = (Get-Acl $FileObject.FullName).Owner
            RawSize = $RawSize
    	}
        Return $Object
    }

    Function CalculateSize {
    	Param (
    		[double]$Size
    	)
    	If ($Size -gt 1000000000)
    	{	$ReturnSize = "{0:N2} GB" -f ($Size / 1GB)
    	}
    	Else
    	{	$ReturnSize = "{0:N2} MB" -f ($Size / 1MB)
    	}
    	Return $ReturnSize
    }

    Function Set-AlternatingRows {
        [CmdletBinding()]
       	Param(
           	[Parameter(Mandatory=$True,ValueFromPipeline=$True)]
            [object[]]$Lines,
           
       	    [Parameter(Mandatory=$True)]
           	[string]$CSSEvenClass,
           
            [Parameter(Mandatory=$True)]
       	    [string]$CSSOddClass
       	)
    	Begin {
    		$ClassName = $CSSEvenClass
    	}
    	Process {
            ForEach ($Line in $Lines)
            {	$Line = $Line.Replace("<tr>","<tr class=""$ClassName"">")
        		If ($ClassName -eq $CSSEvenClass)
        		{	$ClassName = $CSSOddClass
        		}
        		Else
        		{	$ClassName = $CSSEvenClass
        		}
        		Return $Line
            }
    	}
    }

    #Validate sort parameter
    Switch -regex ($Sort)
    {   "^folder.?$" { $SortBy = "Folder Name";Break }
        "created" { $SortBy = "Created On";Break }
        "changed" { $SortBy = "Last Updated";Break }
        default { $SortBy = $Sort }
    }
            
    $Report = @()
    $TotalSize = 0
    $NumDirs = 0
    $Title = @()
    Write-Verbose "$(Get-Date): Script begins!"
}

Process {
    $ExcludedFolders = @('Program Files','Program Files (x86)','Windows','Users','Config.Msi','_serviceDesk','Logs','clients','Intel','PerfLogs')
    #$UserFolder = (Get-ChildItem "c:\Users" -Directory | sort LastWriteTime | select -last 1 )
    $UserFolder = (Get-ChildItem -path "c:\Users\*" -Directory | where {($_.psiscontainer -eq $true) -AND ($_.Name -eq "$env:USERNAME")})
    $LocalDriveRoot = (Get-ChildItem "c:\"| Where {($_.PSIsContainer) -and (!$ExcludedFolders.Contains($_.Name))})
    $Path1 = Join-Path ($UserFolder.FullName) 'OneDrive - Ragn-Sells\Documents'
    $Path2 = Join-Path ($UserFolder.FullName) 'OneDrive - Ragn-Sells\Pictures'
    $Path3 = Join-Path ($UserFolder.FullName) 'OneDrive - Ragn-Sells\Desktop'
    $Path4 = Join-Path ($UserFolder.FullName) 'OneDrive - Ragn-Sells\Downloads'
    $Path5 = Join-Path ($UserFolder.FullName) 'OneDrive - Ragn-Sells'
    $Path6 = Join-Path ($UserFolder.FullName) 'AppData\Roaming\Apple Computer\iTunes'
    $Path7 = Join-Path ($UserFolder.FullName) 'Dropbox'
    $Path8 = Join-Path ($UserFolder.FullName) 'Documents'
    $Path9 = Join-Path ($UserFolder.FullName) 'Downloads'
    $Path10 = Join-Path ($UserFolder.FullName) 'Desktop'
    $Path11 = Join-Path ($UserFolder.FullName) 'Pictures'
	$path12 = Join-Path ($UserFolder.FullName) 'AppData\Local\Microsoft\Outlook'
    $Paths = $Path12,$Path11,$Path10,$Path9,$Path8,$Path7,$Path6,$Path5,$Path4,$Path3,$Path2,$Path1,$LocalDriveRoot.Fullname
    ForEach ($Path in $Paths)
    {   #Test if path exists
        If (-not (Test-Path $Path))
        {   $Result += $Object = New-Object PSObject -Property @{
        		'Folder Name' = $Path
        		'Created on' = ""
        		'Last Updated' = ""
        		Size = ""
        		Owner = "Path not found"
                RawSize = 0
        	}
            $Title += $Path
            Continue
        }
            
        #First get the properties of the starting path
        $NumDirs ++
        Write-Verbose "$(Get-Date): Now working on $Path..."
        $Root = Get-Item -Path $Path 
        $Result = AddObject $Root
        $TotalSize += $Result.RawSize
        $Report += $Result
        $Title += $Path

        #Now loop through all the subfolders
        $ParamSplat = @{
            Path = $Path
            Recurse = $Recurse
        }
        ForEach ($Folder in (Get-ChildItem @ParamSplat | Where { $_.PSisContainer }))
        {	$Report += AddObject $Folder
            $NumDirs ++
        }
    }
}

End {
    #Create the HTML for our report
    $Header = @"
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
.odd  { background-color:#ffffff; }
.even { background-color:#dddddd; }
</style>
<Title>
Folder Sizes for "$Path"
</Title>
"@

    $TotalSize = CalculateSize $TotalSize
    $DiskSpaceFree = Get-WmiObject win32_logicaldisk -Filter "DeviceID='C:'" | Select-Object Freespace
    $DiskSpaceFreeGB = $DiskSpaceFree.Freespace
    $DiskSpaceFreeGB = CalculateSize $DiskSpaceFreeGB
    $Pre = "<h1>Folder size report for Computer $env:computername </h1><h3>Folders processed: ""$($Title -join ", ")""</h3>"
    $Post = "<h2><p>Total number of folders: $NumDirs<br>Total space used:  $TotalSize<br>Free disk space c:\ : $DiskSpaceFreeGB</p></h2>Run on $(Get-Date -f 'yyyy/mm/dd hh:mm:ss tt')</body></html>"

    #Create the report and save it to a file
   #$HTML = $Report | Select 'Folder Name',Owner,'Created On','Last Updated',Size | Sort $SortBy -Descending:$Descending | ConvertTo-Html -PreContent $Pre -PostContent $Post -Head $Header | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-File $ReportPath\FolderSizes.html
   #$HTML = $Report | Sort $SortBy -Descending:$Descending | Select 'Folder Name',Owner,'Created On','Last Updated',Size | ConvertTo-Html -PreContent $Pre -PostContent $Post -Head $Header | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-File $ReportPath\FolderSizes.html
    
    if(Test-Path $ReportPath)
    {
       
        $HTML = $Report | Sort $SortBy -Descending | Select 'Folder Name',Owner,'Created On','Last Updated',Size | ConvertTo-Html -PreContent $Pre -PostContent $Post -Head $Header | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-File -force $ReportPath\FolderSizes.html
    }
    else
    {
        new-item -force -path $ReportPath -type directory
        $HTML = $Report | Sort $SortBy -Descending | Select 'Folder Name',Owner,'Created On','Last Updated',Size | ConvertTo-Html -PreContent $Pre -PostContent $Post -Head $Header | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-File -force $ReportPath\FolderSizes.html
    }

    #Display the report in your default browser
    #& $ReportPath\FolderSizes.html
    
    Write-Verbose "$(Get-Date): $NumDirs folders processed"
    Write-Verbose "$(Get-Date): Script completed!"

    #Send Report via mail
    $searcher = [adsisearcher]"(samaccountname=$env:USERNAME)"
    $LoggedInUserEmail = $searcher.FindOne().Properties.mail
    $logdir = "\\STOPSCCM01\SCCMClient$\DiskFullReport"
    
    $logFile = "$logdir\$env:computername\$($myInvocation.MyCommand).log"
    Start-Transcript $logFile –Verbose
    Write-Verbose –Message "$Date Start Logging to $logFile" –Verbose
    Write-Verbose –Message "User: $LoggedInUserEmail" –Verbose
    Write-Verbose –Message "Computername: $env:computername" –Verbose

	$EmailFrom = "$LoggedInUserEmail"
	$EmailTo = "servicedesk@ragnsells.com"
	$BCC = "tatsumi.morota@ragnsells.com"
	#$SmtpServer = "ragnsells-se.mail.protection.outlook.com"
    $SmtpServer = "smtp.ragnsells.net"
	$date = Get-Date
	$Log1 = "C:\_Servicedesk\Logs\FolderSizes.html"
	
	$message = @{
		Subject	    = "$env:computername har dåligt med diskutrymme/has low diskspace"
		Body	    = "<br>English below<br><br>Ragn-Sells IT har hittat att användare $LoggedInUserEmail på dator $env:computername har för lite diskutrymme kvar, ta gärna bort de största ej företagsrelaterade filer/foldrar enligt bifogad rapport.<br>Du får det här meddelandet eftersom det inte går att utföra datorunderhåll och kan inte garantera att datorn fortsätter att fungera som förväntat. Minst 20 GB fritt utrymme behövs!<br><br>Användning av företagsdatorn för lagring av privata filer och användning av molnlagring som inte tillhandahålls av Ragn-Sells är inte tillåtet<br><br>Den loggfil som visar de största mapparna finns här:<a href=file://C:\_servicedesk\Logs\FolderSizes.html>LargeFolders</a><br>Kontakta servicedesk om du behöver hjälp<br><br>English<br>Ragn-Sells IT has found that user $LoggedInUserEmail on computer $env:computername has low diskspace, please remove none business related files/folders as seen in folder link below.<br>You receive this email because IT can not perform maintenace task and cannot guarantee that your computer will continue to work as expected. At least 20GB freespace is needed!<br><br>Using company computer for storing private files and using cloud storage other than provided by Ragn-Sells is not allowed<br><br>The log file showing the largest folders can be found here: <a href=file://C:\_servicedesk\Logs\FolderSizes.html>LargeFolders</a><br>Please contact servicedesk if you need further assistance<br><br>Regards<br>Ragn-Sells IT"
        From	    = "$EmailFrom"
		To		    = "$EmailTo"
        Bcc         = "$BCC"
		SmtpServer  = "$SmtpServer"
		Attachments = "$Log1"
            }
	
	Write-Verbose "Sending e-mail of logs" -verbose
	try{
	    Send-MailMessage @message -BodyAsHtml -Encoding utf7 -verbose
        }
    catch
        {
        Write-Verbose –Message "Send Mail failed" -verbose
        }
Stop-Transcript -verbose
}