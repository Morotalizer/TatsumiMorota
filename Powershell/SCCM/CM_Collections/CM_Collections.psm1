<#
.SYNOPSIS
    Module for importing and exportin SCCM collections
.DESCRIPTION
    Module for importing and exportin SCCM collections
.NOTES
    Version:        1.2
    Author:			Rob Looman
    Creation Date:	2013-12-05
 
.EXAMPLE
 
#>


Function Import-CMDeviceCollectionsFromXML {
  <#
  .SYNOPSIS
  This function will import device collections into SCCM
  .DESCRIPTION
  When a valid XML file has been specified this function will create the collections. 
  After the creating the collection will be moved to the correct folder. Exisisting 
  collections will not be recreated. QueryMembership Rules, Collection Variables and 
  Refresh Schedules will only be created when they not exist in SCCM but are specified 
  in the xml file.
  .EXAMPLE
  Import-CMDeviceCollectionFromXML C:\Windows\Temp\Collections.xml 
  .EXAMPLE
  Import-CMDeviceCollectionFromXML C:\Windows\Temp\Collections.xml -MoveOnly
  .EXAMPLE
  Import-CMDeviceCollectionFromXML C:\Windows\Temp\Collections.xml -CheckOnly
  .PARAMETER Path
  Path to a valid XML File
  .PARAMETER CheckOnly
  Will only check if the collections exists but will not create collections
  .PARAMETER MoveOnly
  Will only move collections and will not check or create collections.
  #>
  [CmdletBinding()]
  param
  (
        [Parameter( Mandatory=$True,  
                    ValueFromPipeline=$True, 
                    ValueFromPipelineByPropertyName=$True,   
                    HelpMessage='Specify the XML file')]
        [Alias('Path')]
        [ValidateScript({Test-Path $_})]
    [string] $XMLPath,
	    [Parameter(HelpMessage='Only check if collection exists and at right location')]	
    [switch]$CheckOnly=$false,
	    [Parameter(HelpMessage='Only moves collections')]	
    [switch]$MoveOnly=$false,
	    [Parameter(HelpMessage='Overwrite Collection Variable\Query\Schedule')]	
    [switch]$Overwrite=$false
  )
  begin{
         Import-CMCollectionsFromXML -Type "Device" -XML $XMLPath -CheckOnly:$CheckOnly -MoveOnly:$MoveOnly -Overwrite:$Overwrite
  }
}     #End Function 

Function Import-CMUserCollectionsFromXML {
  <#
  .SYNOPSIS
  This function will import user collections into SCCM
  .DESCRIPTION
  When a valid XML file has been specified this function will create the collections. 
  After the creating the collection will be moved to the correct folder. Exisisting 
  collections will not be recreated. QueryMembership Rules, Collection Variables and 
  Refresh Schedules will only be created when they not exist in SCCM but are specified 
  in the xml file.
  .EXAMPLE
  Import-CMDeviceCollectionFromXML C:\Windows\Temp\Collections.xml 
  .EXAMPLE
  Import-CMDeviceCollectionFromXML C:\Windows\Temp\Collections.xml -MoveOnly
  .EXAMPLE
  Import-CMDeviceCollectionFromXML C:\Windows\Temp\Collections.xml -CheckOnly
  .PARAMETER Path
  Path to a valid XML File
  .PARAMETER CheckOnly
  Will only check if the collections exists but will not create collections
  .PARAMETER MoveOnly
  Will only move collections and will not check or create collections.
  #>
  [CmdletBinding()]
  param
  (
        [Parameter( Mandatory=$True,  
                    ValueFromPipeline=$True, 
                    ValueFromPipelineByPropertyName=$True,   
                    HelpMessage='Specify the XML file')]
        [Alias('Path')]
        [ValidateScript({Test-Path $_})]
    [string] $XMLPath,
	    [Parameter(HelpMessage='Only check if collection exists and at right location')]	
    [switch]$CheckOnly=$false,
	    [Parameter(HelpMessage='Only moves collections')]	
    [switch]$MoveOnly=$false,
	    [Parameter(HelpMessage='Overwrite Collection Variable\Query\Schedule')]	
    [switch]$Overwrite=$false
  )
  begin{
        Import-CMCollectionsFromXML -Type "User" -XML $XMLPath -CheckOnly:$CheckOnly -MoveOnly:$MoveOnly -Overwrite:$Overwrite
    }
}       #End Function 

Function Import-CMCollectionsFromXML {
  <#
  .SYNOPSIS
  This function will import collections into SCCM
  .DESCRIPTION
  When a valid XML file has been specified this function will create the collections. 
  After the creating the collection will be moved to the correct folder. Exisisting 
  collections will not be recreated. QueryMembership Rules, Collection Variables and 
  Refresh Schedules will only be created when they not exist in SCCM but are specified 
  in the xml file.
  .EXAMPLE
  Import-CMCollectionFromXML C:\Windows\Temp\Collections.xml 
  .EXAMPLE
  Import-CMCollectionFromXML C:\Windows\Temp\Collections.xml -Type "User" -MoveOnly
  .EXAMPLE
  Import-CMCollectionFromXML C:\Windows\Temp\Collections.xml -Type "Device" -CheckOnly
  .PARAMETER Path
  Path to a valid XML File
  .PARAMETER Type
  Specify the type to import (User or Device) if none is specified both types will be 
  imported
  .PARAMETER CheckOnly
  Will only check if the collections exists but will not create collections
  .PARAMETER MoveOnly
  Will only move collections and will not check or create collections.
  #>
  [CmdletBinding()]
  param
  (
        [Parameter( Mandatory=$True)]
        [Alias('Path')]
        [ValidateScript({Test-Path $_})]
    [String] $XMLPath,
        [Parameter( Mandatory=$false)]
        [ValidateSet("Device", "User", "")]    
    [string] $Type="",
	    [Parameter(HelpMessage='Only check if collection exists and at right location')]	
    [switch]$CheckOnly=$false,
	    [Parameter(HelpMessage='Only moves collections')]	
    [switch]$MoveOnly=$false,
	    [Parameter(HelpMessage='Overwrite Collection Variable\Query\Schedule')]	
    [switch]$Overwrite=$false
  )

  begin {
    #SCCM connection
    if((Get-Location).Drive.Provider.Name -ne "CMSite"){
        Write-Error "This command cannot be run from the current drive. To run this command you must first connect to a Configuration Manager drive"
        continue
    }

    $SCCMSiteCode = (Get-Location).Drive.SiteCode
    $SCCMSiteServer = (Get-Location).Drive.SiteServer

    #object definition and clearing   
    $dicCollections = @{}
    $CMCollection  = $null   
    $CMLimiting  = $null   
    $xmlfile = $null
    $xmlDeviceCollections = $null
    $xmlCollections = $null

    #progress
    $total = 0
    $current = 0
    $progress = 0

    #Collection Variable 
    $colName = ""
    write-Verbose "Importing $Type collections to site $SCCMSiteCode" 
  }

  process {
        

        if([String]::IsNullOrEmpty($Type)){
            Import-CMCollectionsFromXML  $XMLPath "Device" -CheckOnly:$CheckOnly -MoveOnly:$MoveOnly
            Import-CMCollectionsFromXML  $XMLPath "User" -CheckOnly:$CheckOnly -MoveOnly:$MoveOnly
            continue
        }
                
        [xml] $xmlfile = Get-Content $XMLPath

        #Checking for tags
        if($xmlfile.Collections -eq $null){
            Write-Error "Missing Collections tag in XML"
            continue
        }    
        if($xmlfile.Collections."$($type)Collections" -eq $null){
            Write-Error "Missing $($type)Collections tag in XML "
            continue
        }
        if($xmlfile.Collections."$($type)Collections".GetElementsByTagName("Collection").Count -eq 0){
            Write-Warning "No $type collections found"
            continue

        }
   
        #retrieving collections 
        $xmlDeviceCollections = $xmlfile.Collections."$($type)Collections"
        $xmlCollections = $xmlDeviceCollections.GetElementsByTagName("Collection")
        $xmlFolders = $xmlDeviceCollections.GetElementsByTagName("Folder")

        #progress
        $total = $xmlCollections.Count 
        $current = 0
        $progress = 0

        if(-not $MoveOnly){
            foreach($xmlCollection in $xmlCollections){   
               
               $IsNewCollection = $false  
               $colName = $xmlCollection.name   
               $colLimiting = $xmlCollection.LimitToCollectionName
               $colRefreshType = $xmlCollection.RefreshType
               $colComment = $xmlCollection.Comment

               if($colComment -eq $null){$colComment = " "}
               #Progress                 
               $current ++
               $progress = [int] (($current/ $total)*100) 
               Write-Verbose "Collection: $colName  $current of $total"
               Write-Progress -Activity "Creating SCCM Collection" -PercentComplete $progress -CurrentOperation "$progress% Creating $colName " -Status "Please Wait.."
               
               #Checking for duplicates    
               if($dicCollections.ContainsKey($colName)){
                    Write-Warning " Duplicated collection in XML ($colName)"
                    continue
               }
               
               #Checking for Collection 
               $CMCollection = & Get-CM$($Type)Collection -Name $colName 
               if($CMCollection -eq $null){
                    #When CHeckOnly is specified only a warning is displayed    
                    if($CheckOnly){
                        Write-Warning "  $colName NOT EXISTS (CheckOnly)"
                        continue
                    }

                    #loading the limiting collection 
                    if($dicCollections.ContainsKey($colLimiting)){
                        $CMLimiting = $dicCollections[$colLimiting]
                    }

                    #try reloading 
                    if($CMLimiting.Name -eq $null){
                        $dicCollections.Remove($colLimiting)
                        $CMLimiting = & Get-CM$($Type)Collection -Name $colLimiting
                        $dicCollections.Add($colLimiting,$CMLimiting)
                    }

                    #if still failed write errror
					if($CMLimiting.Name -eq $null){	
                            Write-Error "Limiting collection not found ($($xmlCollection.LimitToCollectionName))"
                            continue
                    }
                    
                    #set refreshtype to manual if it has not been set 
                    if(-Not [String]::IsNullOrEmpty($colRefreshType )){$colRefreshType = "Manual"}

                    #creaet SCCM collection 
                            $CMCollection = & New-CM$($Type)Collection `
                                -Name $colName `
                                -Comment "$($colComment)" `
                                -LimitingCollectionName $colLimiting `
                                -RefreshType $colRefreshType
                            $IsNewCollection = $true
                            $CMLimiting = & Get-CM$($Type)Collection -Name $colLimiting
                            $CMCollection = & Get-CM$($Type)Collection -Name $colName

                            $dicCollections.Add($colName,$CMCollection)
                            Write-Host " $colName CREATED"
                   }else{
                   $dicCollections.Add($colName,$CMCollection)
                   Write-Verbose " $colName already exists"
               }

               #Creating Collection Rules
               if($CMCollection.CollectionRules -eq $null -And ($IsNewCollection -or $Overwrite)){ 
                    #Creating Query Rules
                    Foreach($colQuery in $xmlCollection.QueryMembershipRule){
                        & Add-CM$($Type)CollectionQueryMembershipRule `
                            -Collection $CMCollection `
                            -RuleName $colQuery.name `
                            -QueryExpression $colQuery.InnerText `
                            |Out-Null
                        Write-Host "  $colName Query Rule ADDED"
                    }
                    #Creating Include Collection Rule 
                    Foreach($colInclude in $xmlCollection.IncludeCollectionName){
                        & Add-CM$($Type)CollectionIncludeMembershipRule `
                            -Collection $CMCollection `
                            -IncludeCollectionName $colInclude.InnerText `
                            |Out-Null
                        Write-Host "  $colName Include Collection ADDED ($($colInclude.InnerText))"
                    }
                    #Creating Exclude Collection Rule
                    Foreach($colExclude in $xmlCollection.ExcludeCollectionName){
                        & Add-CM$($Type)CollectionExcludeMembershipRule `
                            -Collection $CMCollection `
                            -ExcludeCollectionName $colExclude.InnerText `
                            |Out-Null
                        Write-Host "  $colName Exclude Collection ADDED ($($colExclude.InnerText))"
                    }


                }

               

                

               #Creating Collection Variables
               if($CMCollection.CollectionVariablesCount -eq 0 -And ($IsNewCollection -or $Overwrite)){   
                   foreach($colVariable in $xmlCollection.CollectionVariable){
                        & New-CM$($Type)CollectionVariable `
                            -Collection $CMCollection `
                            -VariableName $colVariable.name `
                            -VariableValue $colVariable.InnerText `
                            |Out-Null
                        Write-Host "  $colName Collection Variable ADDED"
                    }
                }
                

                #Creating Collection Schedule 
               if($CMCollection.RefreshSchedule.Count -eq 0 -And ($IsNewCollection -or $Overwrite)){ 
                    foreach($colSchedule in $xmlCollection.RefreshSchedule){
                        #Load Schedule
                        $Schedule = @{}
                        foreach($property in Get-Member -InputObject $xmlCollection.RefreshSchedule -MemberType Properties){
                            $propName = $property.Name
                            $propValue = $colSchedule.$propName
                            if($propName -eq "type"){continue}
                            if(-not [String]::IsNullOrEmpty($propValue)){
                                $Schedule.Add($propName,$colSchedule.$propName)
                                continue
                            }
                        }
                        $CMCollection = & Get-CM$($Type)Collection -Name $colName
                        Add-CMScheduleToCollection $CMCollection.CollectionID $colSchedule.type $Schedule
                        Write-Host  "  $colName Collection Schedule ADDED"
                    }
                }

            }# End collections
        }# End MoveOnly

        Write-Progress -Activity "Creating SCCM Collection" -Complete       
        $current = 0

        if($CheckOnly){return}
           
        #Moving all objects
        foreach($folder in $xmlFolders){
            #retrieving full path
            $fullpath =""
            $parent = $folder.ParentNode
            While($parent.ParentNode.Folder -ne $null){
         
                $fullpath = $parent.Name +"\" + $fullpath 
                $parent = $parent.ParentNode
            }
            $fullpath = $fullpath + $folder.Name
            $fullpath = Join-Path "$($SCCMSiteCode):\$($Type)Collection\" $fullpath 

            if(Test-Path $fullpath){}else{
                New-Item -itemType directory -Path $fullpath
            }

            foreach($collection in $folder.collection){
                $current++
                $percentage = [int] (($current/ $total)*100)
                Write-Progress -Activity "Moving SCCM Collections" -PercentComplete $percentage -CurrentOperation "$percentage% Moving $($Collection.Name)" -Status "Please Wait.."
           
                if($dicCollections.ContainsKey($collection.Name)){
                    $CMCollection = $dicCollections[$collection.Name]
                }else{
                    $CMCollection = & Get-CM$($Type)Collection -Name $collection.Name
                }
                if($CMCollection.CollectionID -eq $null){$CMCollection = & Get-CM$($Type)Collection -Name $collection.Name}
                Write-Verbose "Moving collcetion $($collection.Name) to $fullpath (ID:$($CMCollection.CollectionID))"
                Move-CMObject -FolderPath $fullpath -ObjectId $CMCollection.CollectionID | Out-Null
            }
        }


  }
}           #End Function 

Function Export-CMDeviceCollectionsToXML{
<#
  .SYNOPSIS
  This function will return a XML file of device collections
  .DESCRIPTION
  Device collections will be returned into a valid XML file. A path can be 
  specified to only export certain collection
  .EXAMPLE
  Export-CMDeviceCollectionsToXML 
  .EXAMPLE
  Export-CMDeviceCollectionsToXML "GEN\ALL"
  .PARAMETER Path
  Folder path to export the collections
  #>
  [CmdletBinding()]
  param
  (
        [Parameter( Mandatory=$False,  
                    ValueFromPipeline=$True, 
                    ValueFromPipelineByPropertyName=$True,   
                    HelpMessage='Specify location of the collection to export ')]
        [Alias('Path')]
    [string] $FolderPath =""
  )
  begin 
  {
            Export-CMCollectionsToXML $FolderPath -NoUser
  }
}        #End Function 

Function Export-CMUserCollectionsToXML{
<#
  .SYNOPSIS
  This function will return a XML file of user collections
  .DESCRIPTION
  User collections will be returned into a valid XML file. A path can be 
  specified to only export certain collection
  .EXAMPLE
  Export-CMUserCollectionsToXML 
  .EXAMPLE
  Export-CMUserCollectionsToXML "GEN\ALL"
  .PARAMETER Path
  Folder path to export the collections
  #>
  [CmdletBinding()]
  param
  (
        [Parameter( Mandatory=$False,  
                    ValueFromPipeline=$True, 
                    ValueFromPipelineByPropertyName=$True,   
                    HelpMessage='Specify location of the collection to export ')]
        [Alias('Path')]
    [string] $FolderPath =""
  )
  begin 
  {
            Export-CMCollectionsToXML $FolderPath -NoDevice
  }
}          #End Function 

Function Export-CMCollectionsToXML{
<#
  .SYNOPSIS
  This function will return a XML file of collections
  .DESCRIPTION
  User collections will be returned into a valid XML file. A path can be 
  specified to only export certain collection
  .EXAMPLE
  Export-CMUserCollectionsToXML 
  .EXAMPLE
  Export-CMUserCollectionsToXML "GEN\ALL"
  .EXAMPLE
  Export-CMUserCollectionsToXML "GEN\ALL" -NoUser
  .EXAMPLE
  Export-CMUserCollectionsToXML "GEN\ALL" -NoDevice
  .PARAMETER Path
  Folder path to export the collections
  .PARAMETER NoDevice
  Will not export device collections
  .PARAMETER NoUser
  Will not export user collections
  #>
  [CmdletBinding()]
  param
  (
        [Parameter( Mandatory=$False,  
                    ValueFromPipeline=$True, 
                    ValueFromPipelineByPropertyName=$True,   
                    HelpMessage='Specify folder 
                    ')]
        [Alias('Path')]
    [string] $CollectionPath ="",
	    [Parameter(HelpMessage='Only moves collections')]	
    [switch]$NoDevice=$false,
	    [Parameter(HelpMessage='Only moves collections')]	
    [switch]$NoUser=$false
  )

  begin {

    if((Get-Location).Drive.Provider.Name -ne "CMSite"){
        Write-Error "This command cannot be run from the current drive. To run this command you must first connect to a Configuration Manager drive"
        continue
    }
    
   
    #Clearing all objects
    $SCCMSiteCode = (Get-Location).Drive.SiteCode
    $xmlStream = New-Object System.IO.MemoryStream 
    $XmlWriter = New-Object System.XMl.XmlTextWriter($xmlStream,$null) 

    write-Verbose "Exporting collections from site $sitecode"
  }

  process {  

    #Formating XML document
    $xmlWriter.Formatting = [System.Xml.Formatting]::Indented

    # Write the XML Decleration
    $xmlWriter.WriteStartDocument()
    $xmlWriter.WriteProcessingInstruction("xml-stylesheet", "type='text/xsl' href='SCCMCollections.xsl'")

    # Write Root Element    
    $xmlWriter.WriteStartElement("Collections")
    
              

    if(-not $NoUser){
        $xmlWriter.WriteStartElement("UserCollections")
       
        $FullPath = Join-Path  "$($SCCMSiteCode):\UserCollection" $CollectionPath
        If($FullPath.EndsWith("\")){$FullPath = $FullPath.Substring(0,$FullPath.Length-1)}
        If(-not (Test-Path $FullPath )){
            Write-Error "Path not found: $FullPath"
            continue
        }

        #creating folder path when path is not empty
        $PathFolder = $CollectionPath.Split("\")
        $PathCount = 0
        ForEach($folder in $pathfolder){
            if(-not[String]::IsNullOrEmpty($folder)){
                $xmlWriter.WriteStartElement("Folder")
                $xmlWriter.WriteAttributeString("name",$folder)
                $PathCount++
            }
        }
        #Connect to site
        Add-CMCollectionFolderToXml "User" $FullPath $XmlWriter

        For($i=0; $i -lt $PathCount;$i++){$xmlWriter.WriteEndElement()}
        #Finalizing XML document
        $xmlWriter.WriteEndElement()  #</UserCollections>
    }

    if(-not $NoDevice){
        $xmlWriter.WriteStartElement("DeviceCollections")
     
        $FullPath = Join-Path  "$($SCCMSiteCode):\DeviceCollection" $CollectionPath
        If($FullPath.EndsWith("\")){$FullPath = $FullPath.Substring(0,$FullPath.Length-1)}
        If(-not (Test-Path $FullPath )){
            Write-Error "Path not found: $FullPath"
            continue
        }

        #creating folder path when path is not empty
        $PathFolder = $CollectionPath.Split("\")
        $PathCount = 0
        ForEach($folder in $pathfolder){
            if(-not[String]::IsNullOrEmpty($folder)){
                $xmlWriter.WriteStartElement("Folder")
                $xmlWriter.WriteAttributeString("name",$folder)
                $PathCount++
            }
        }
        #Connect to site
        Add-CMCollectionFolderToXml "Device" $FullPath $XmlWriter

        For($i=0; $i -lt $PathCount;$i++){$xmlWriter.WriteEndElement()}
        #Finalizing XML document
        $xmlWriter.WriteEndElement()  #</DeviceCollections>
    }

    $xmlWriter.WriteEndElement()  #</Collections>
    $xmlWriter.WriteEndDocument() #EOF

    $xmlWriter.Flush() |Out-Null
    $xmlStream.Flush() |Out-Null

    #Output XML document
    $xmlStream.Position = 0    
    $xmlStreamReader = New-Object System.IO.StreamReader($xmlStream)
    return $xmlStreamReader.ReadToEnd()

  }

 }              #End Function 

Function Add-CMCollectionFolderToXml{
[CmdletBinding()]
  param
  ( $type, $CurrentPath,$xmlWriter, $level=1
  )
  begin {
    $ChildItems = Get-ChildItem $CurrentPath

    #Progress
        $current = 0
        $total = $ChildItems.Count

    Write-Progress -Activity "Exporting Folders" -Id $level -PercentComplete 0  -CurrentOperation "Exporting: $CurrentPath ($percentage%)" -Status "Please Wait.."
    ForEach($folder in $ChildItems){
            #Progress
            $current++
            $percentage = [int] (($current/$total)*100) 
            if($folder.PSIsContainer){
                Write-Progress -Activity "Exporting Folders" -Id $level -PercentComplete $percentage -CurrentOperation "Exporting: $CurrentPath ($percentage%)" -Status "Please Wait.."
             
                $NewPath = Join-Path $CurrentPath $folder.Name
           
                $xmlWriter.WriteStartElement("Folder")
                $xmlWriter.WriteAttributeString("name",$folder.Name)
                $newlevel = $level+1
                Add-CMCollectionFolderToXml  $type $NewPath $xmlWriter $newlevel 
                #Add-CMCollectionToXML  $type $NewPath $xmlWriter $level
                $xmlWriter.WriteEndElement()
            }
    } 
    
    Add-CMCollectionToXML  $type $CurrentPath $xmlWriter $level
    Write-Progress -Activity "Exporting Folders" -Id $level -Completed
 }  
 }            #End Function 

Function Add-CMCollectionToXML{
[CmdletBinding()]
  param
  ($type, $CurrentPath,$xmlWriter,$level
  )
  begin {
  
    #No PowerShell cmdlt availble using WMI
    $SCCMSiteCode = (Get-Location).Drive.SiteCode
    $SCCMServer = (Get-Location).Drive.SiteServer

    $folder = Get-Item $CurrentPath
    if([String]::IsNullOrEmpty($folder.ContainerNodeID)){
        Write-verbose "$currentPath is the root and will not export collections from the root"
        return
    }
    $folderitems = Get-WMIObject -Class SMS_objectContainerItem -NameSpace "root\sms\site_$($SCCMSiteCode)"  -ComputerName  $SCCMServer -Filter "ContainerNodeId = $($folder.ContainerNodeID)"  
    $total = $folderitems.Count
    if($total -eq $null){$total = 1}
    $current = 0
    ForEach($collection in $folderitems){

        $col = & Get-CM$($type)Collection -CollectionID $collection.InstanceKey
            #Status
            $current ++
            
            $percentage = [int] (($current/$total)*100)
            Write-Progress -Activity "Exporting Collections" -Id $level -PercentComplete $percentage -CurrentOperation $col.Name
   
        $xmlWriter.WriteStartElement("Collection")
        $xmlWriter.WriteAttributeString("name",$col.Name)
        #Collection rules
        ForEach($query in $col.CollectionRules){
            if(-not [String]::IsNullOrEmpty($query.QueryExpression)){                
                $xmlWriter.WriteStartElement("QueryMembershipRule")
                $xmlWriter.WriteAttributeString("name",$query.RuleName)
                $xmlWriter.WriteString($query.QueryExpression)
                $xmlWriter.WriteEndElement()
            }
            if(-not [String]::IsNullOrEmpty($query.IncludeCollectionID)){   
                $colInclude = & Get-CM$($type)Collection -CollectionID $query.IncludeCollectionID
                if(-not [String]::IsNullOrEmpty($colInclude.Name)){             
                    $xmlWriter.WriteStartElement("IncludeCollectionName")
                    $xmlWriter.WriteAttributeString("name",$query.RuleName)
                    $xmlWriter.WriteString($colInclude.Name)
                    $xmlWriter.WriteEndElement()
                }else{
                    Write-Warning " Include collection not exported($($col.CollectionID))"
                }
            }
            if(-not [String]::IsNullOrEmpty($query.ExcludeCollectionID)){                                 
                $colExclude = & Get-CM$($type)Collection -CollectionID $query.ExcludeCollectionID
                if(-not [String]::IsNullOrEmpty($colInclude.Name)){             
                    $xmlWriter.WriteStartElement("ExcludeCollectionName")
                    $xmlWriter.WriteAttributeString("name",$query.RuleName)
                    $xmlWriter.WriteString($colExclude.Name)
                    $xmlWriter.WriteEndElement()
                }else{
                    Write-Warning " Exclude collection not exported($($col.CollectionID))"
                }
            }

        }
        if($col.CollectionVariablesCount -ne 0){
            
            $lazycollectionitems = Get-WMIObject -Class SMS_CollectionSettings -NameSpace "root\sms\site_$($SCCMSiteCode)"  -ComputerName  $SCCMServer -Filter "CollectionID = '$($col.CollectionID)'"  
            $collectionitems = [wmi] $lazycollectionitems.__PATH
            ForEach($var in $collectionitems.CollectionVariables){
                $xmlWriter.WriteStartElement("CollectionVariable")
                $xmlWriter.WriteAttributeString("name",$var.Name)
                $xmlWriter.WriteString($var.Value)
                $xmlWriter.WriteEndElement()
                
            }
        }
        #Refresh schedule
        $i = 0
        ForEach($schedule in $col.RefreshSchedule){
            $i++
            $xmlWriter.WriteStartElement("RefreshSchedule")  
            $xmlWriter.WriteAttributeString("name","Schedule $i")  
            $xmlWriter.WriteAttributeString("type",$schedule.OverridingObjectClass.Replace("SMS_ST_",""))        
            $xmlWriter.WriteElementString("Day",$schedule.Day)
            $xmlWriter.WriteElementString("DayDuration",$schedule.DayDuration)
            $xmlWriter.WriteElementString("DaySpan",$schedule.DaySpan)
            $xmlWriter.WriteElementString("HourDuration",$schedule.HourDuration)
            $xmlWriter.WriteElementString("HourSpan",$schedule.HourSpan)
            $xmlWriter.WriteElementString("IsGMT",$schedule.IsGMT)
            $xmlWriter.WriteElementString("MinuteDuration",$schedule.MinuteDuration)
            $xmlWriter.WriteElementString("MinuteSpan",$schedule.MinuteSpan)
            $xmlWriter.WriteElementString("MonthDay",$schedule.MonthDay)
            $xmlWriter.WriteElementString("StartTime",$schedule.StartTime.ToString("yyyymmddHHMMss.000000+***"))
            $xmlWriter.WriteElementString("ForNumberOfWeeks",$schedule.ForNumberOfWeeks)
            $xmlWriter.WriteElementString("WeekOrder",$schedule.WeekOrder)
			
            $xmlWriter.WriteElementString("ForNumberOfMonths",$schedule.ForNumberOfMonths)
			
            $xmlWriter.WriteEndElement()
            
        }
        #Refresh type 


        
        $xmlWriter.WriteElementString("RefreshType",$col.RefreshType)
        $xmlWriter.WriteElementString("LimitToCollectionName",$col.LimitToCollectionName)
        $xmlWriter.WriteElementString("ServiceWindowsCount",$col.ServiceWindowsCount)
        $xmlWriter.WriteEndElement()
    }
    Write-Progress -Activity "Exporting Collections" -Completed
}                                
}                  #End Function 

Function Add-CMScheduleToCollection{
[CmdletBinding()]
  param
  ( $collectionID, $type,$dicSchedule
  )
  begin {
    if([String]::IsNullOrEmpty($collectionID)){
        Write-Warning " No schedule set because of missing collectionID"
        continue
    }
    $SCCMSiteCode = (Get-Location).Drive.SiteCode
    $SCCMSiteServer = (Get-Location).Drive.SiteServer

    $SMS_ST_RecurInterval = "SMS_ST_$Type"
    $class_SMS_ST_RecurInterval= [WMIClass] "\\$($SCCMSiteServer)\ROOT\SMS\Site_$($SCCMSiteCode):$($SMS_ST_RecurInterval)"

    $newSchedule = $class_SMS_ST_RecurInterval.CreateInstance()   
    foreach($prop in $dicSchedule.Keys){
        $value = $dicSchedule[$prop]
        if($value.GetType() -eq (Get-Date).GetType()){
            $value = $value.ToString("yyyymmddHHMMss.000000+***")
        }

        if($newSchedule.Properties.Name -contains $prop){
            $newSchedule.$prop = $value
        }
     
    }
    
    $Collection = Get-WmiObject -Class SMS_Collection -Namespace root\sms\site_$SCCMSiteCode -ComputerName $SCCMSiteServer -Filter "CollectionID ='$($collectionID)'"
    
    $Collection = [wmi]$Collection.__PATH
    $Collection.RefreshSchedule = $newSchedule
    $Collection.put() | Out-Null
}    
}             #End Function 
