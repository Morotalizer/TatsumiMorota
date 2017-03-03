<#
 Script for creating SCCM driver structure from Source directory structure. 
 The script also creates driver packages in the correct SCCM folders and imports the corresponding drivers into it.

 Base script taken from Coretech Repository and rewritten to add further functions such as individual driver folders based on Architecture/Manufacturer/Model
 V3 is for Driversource structre:
 [Manufacturer]
  |
   -----[Model]
            |
             ----[OS]
                   
 Author: Paul Raukas
 Creation Date: 2015-10-25
 Modified by:Tatsumi Morota 
 2017-02-08 Added Function Create-RecursiveTxtFile
#> 

# Define script variables

#Get the Base sourcepath (\\<SERVER>\SCCMSRC$) from the script config file that should be present in the parent directory to this one.

$config = Get-Content "$PSScriptRoot\..\ScriptConfig.ini"
$config | foreach-object -begin {$h=@{}} -process { $k = [regex]::split($_,'='); if(($k[0].CompareTo("") -ne 0) -and ($k[0].StartsWith("[") -ne $True) -and ($k[0].StartsWith("#") -ne $True)) { $h.Add($k[0], $k[1]) } }

$BaseSourcePath = ([string]$h.'SCCMSourcePath ').TrimStart()
$SCCMPrimarySiteServer = ([string]$h.'SCCMPrimarySiteServer ').TrimStart()
$SCCMSiteCode =([string]$h.'SCCMSiteCode ').TrimStart()
$DistributionPointGroup = ([string]$h.'DistributionPointGroup ').TrimStart()
$SCCMDriverSourcePath = $BaseSourcePath + "\Appsource\OSD\DriverSource"
$SCCMDriverPackageSourcePath = $BaseSourcePath + "\Appsource\OSD\DriverPackages"
$Manufacturers = Get-ChildItem -Directory -Path $SCCMDriverSourcePath

#$SCCMPrimarySiteServer = CM01
#$SCCMSiteCode = PS1

#write-Host "BaseSourcePath: " $BaseSourcepath
#$BaseSoftwareUpdateDeploymentPackagePath = $BaseSourcepath +"\WSUS\UpdateServicesPackages\"
#write-host "DeployPackageLocation" $BaseSoftwareUpdateDeploymentPackagePath
#Write-Host $SCCMDriverSourcePath
#write-host $SCCMDriverPackageSourcePath

#$DistributionPointGroup = "All DPs"

# End Define script variables

Function Get-SCCMDriverCategory
{
    [CmdletBinding()]
    PARAM
    (
        [Parameter(Position=1)] $categoryName
    )

    # Build the appropriate filter to return all categories or just one specified by name
    $filter = "CategoryTypeName = 'DriverCategories'"
    if ($categoryName -eq "" -or $categoryName -eq $null)
    {
        Write-Debug "Retriving all categories"
    }
    else
    {
        $filter += " and LocalizedCategoryInstanceName = '" + $categoryName + "'"
    }

    # Retrieve the matching list
    Get-SCCMObject SMS_CategoryInstance -filter $filter
}

Function New-SCCMDriverCategory
{
    [CmdletBinding()]
    PARAM
    (
        [Parameter(Position=1)] $categoryName
    )

    # Create a SMS_Category_LocalizedProperties instance
    $localizedClass = [wmiclass]"\\$sccmServer\$($sccmNamespace):SMS_Category_LocalizedProperties"

    # Populate the localized settings to be used with the new driver instance
    $localizedSetting = $localizedClass.psbase.CreateInstance()
    $localizedSetting.LocaleID =  1033 
    $localizedSetting.CategoryInstanceName = $categoryName
    [System.Management.ManagementObject[]] $localizedSettings += $localizedSetting

    # Create the unique ID
    $categoryGuid = [System.Guid]::NewGuid().ToString()
    $uniqueID = "DriverCategories:$categoryGuid"

    # Build the parameters for creating the collection
    $arguments = @{CategoryInstance_UniqueID = $uniqueID; LocalizedInformation = $localizedSettings; SourceSite = $sccmSiteCode; CategoryTypeName = 'DriverCategories'}

    # Create the new instance
    set-wmiinstance -class SMS_CategoryInstance -arguments $arguments -computername $sccmServer -namespace $sccmNamespace
}

Function New-SCCMDriverPackage
{
    [CmdletBinding()]
    PARAM
    (
        [Parameter(Position=1)] $name, 
        [Parameter(Position=2)] $description,
        [Parameter(Position=3)] $sourcePath
    )

    # Build the parameters for creating the collection
    $arguments = @{Name = $name; Description = $description; PkgSourceFlag = 2; PkgSourcePath = $sourcePath}
    $newPackage = set-wmiinstance -class SMS_DriverPackage -arguments $arguments -computername $sccmServer -namespace $sccmNamespace
    
    # Hack - for some reason without this we don't get the PackageID value
    $hack = $newPackage.PSBase | select * | out-null
    
    # Return the package
    $newPackage
}

Function New-SCCMFolder            
{            
  Param(                      
    $FolderName,
    $FolderType,            
    $ParentFolderID = 0
  )          
    
    
  #$FolderName = $Folder.Name # Added for troubleshooting
  #$FolderType = 25 # Added for troubleshooting
  If ($FolderType -eq "Device") { $FolderType = 5000 }
  If ($FolderType -eq "User") { $FolderType = 5001 }
                
  $SMSFolderClass = "SMS_ObjectContainerNode"             
  $Colon = ":"            
                    
  $WMIConnection = [WMIClass]"\\$sccmServer\$sccmNamespace$Colon$SMSFolderClass"            
  $CreateFolder = $WMIConnection.psbase.CreateInstance()            
  $CreateFolder.Name = $FolderName            
  $CreateFolder.ObjectType = $FolderType            
  $CreateFolder.ParentContainerNodeid = $ParentFolderID            
  $FolderResult = $CreateFolder.Put()
  
  $FolderID = $FolderResult.RelativePath.Substring($FolderResult.RelativePath.Length - 8, 8)
  
  $FolderID            
                
}

Function Move-SCCMObject            
{            
  Param(                    
    $SourceFolderID = 0,            
    $TargetFolderID,            
    $ObjectID,            
    $ObjectType
  )
          
  If ($ObjectType -eq "Device") { $ObjectType = 5000 }
  If ($ObjectType -eq "User") { $ObjectType = 5001 }     
  #$TargetFolderID = $CMDriverPackageFolderID # Added for troubleshooting      
  #$ObjectID = $CMPackage.PackageID  # Added for troubleshooting
  #$ObjectType = 23 # Added for troubleshooting
                      
  $Method = "MoveMembers"            
  $SMSObjectClass = "SMS_ObjectContainerItem"            
  $Colon = ":"            
                    
  $WMIConnection = [WMIClass]"\\$sccmServer\$sccmNamespace$Colon$SMSObjectClass"            
  $InParams = $WMIConnection.psbase.GetMethodParameters("MoveMembers")            
  $InParams.ContainerNodeID = $SourceFolderId            
  $InParams.InstanceKeys = $ObjectID           
  $InParams.ObjectType = $ObjectType            
  $InParams.TargetContainerNodeID = $TargetFolderID            
  $null = $WMIConnection.psbase.InvokeMethod($Method,$InParams,$null)
           
}

Function Get-ContentHash
{
    Param (
        $File,
        [ValidateSet("sha1","md5")]
        [string]$Algorithm="md5"
    )
	
    $content = "$($file.Name)$($file.Length)"
    $algo = [type]"System.Security.Cryptography.md5"
	$crypto = $algo::Create()
    $hash = [BitConverter]::ToString($crypto.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($content))).Replace("-", "")
    $hash
}

Function Get-FolderHash
{
    Param (
        [string]$Folder=$(throw("You must specify a folder to get the checksum of.")),
        [ValidateSet("sha1","md5")]
        [string]$Algorithm="md5"
    )
    
	Get-ChildItem $Folder -Recurse -Exclude "*.hash" | % { $content += Get-ContentHash $_ $Algorithm }
    
    
    $algo = [type]"System.Security.Cryptography.$Algorithm"
	$crypto = $algo::Create()
	$hash = [BitConverter]::ToString($crypto.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($content))).Replace("-", "")
    
    $hash
}

Function Write-Custom($message, [System.ConsoleColor]$foregroundcolor)  
{  
	
	For ($i = 2; $i -le $currentDepth; $i++)
	{
		$tab += "`t"
	}
	
	$currentColor = $Host.UI.RawUI.ForegroundColor  
	if ($foregroundcolor)
	{
		$Host.UI.RawUI.ForegroundColor = $foregroundcolor
	}
	if ($message)  
	{  
		Write-Output "$($tab)$($message)"
	}  
	$Host.UI.RawUI.ForegroundColor = $currentColor 
}

Function Write-Headline($message)
{

	$dot = "------------------------------------------------------------------------------------------------------------"
	
	For ($i = 2; $i -le $currentDepth; $i++)
	{
		$dot = $dot.Substring(0, $dot.Length-8)
	}
	Write-Custom " "
	Write-Custom "$($dot)"
	Write-Custom "$($message)"
	Write-Custom "$($dot)"
}

Function New-SCCMConnection {

    [CmdletBinding()]
    param
    (
        [Parameter(Position=1)] $serverName,
        [Parameter(Position=2)] $siteCode
    )


    # Clear the results from any previous execution

    Clear-Variable -name sccmServer -errorAction SilentlyContinue
    Clear-Variable -name sccmNamespace -errorAction SilentlyContinue
    Clear-Variable -name sccmSiteCode -errorAction SilentlyContinue
    Clear-Variable -name sccmConnection -errorAction SilentlyContinue


    # If the $serverName is not specified, use "."

    if ($serverName -eq $null -or $serverName -eq "")
    {
        $serverName = "."
    }


    # Get the pointer to the provider for the site code

    if ($siteCode -eq $null -or $siteCode -eq "")
    {
        Write-Verbose "Getting provider location for default site on server $serverName"
        $providerLocation = get-wmiobject -query "select * from SMS_ProviderLocation where ProviderForLocalSite = true" -namespace "root\sms" -computername $serverName -errorAction Stop
    }
    else
    {
        Write-Verbose "Getting provider location for site $siteName on server $serverName"
        $providerLocation = get-wmiobject -query "select * from SMS_ProviderLocation where SiteCode = '$siteCode'" -namespace "root\sms" -computername $serverName -errorAction Stop
    }


    # Split up the namespace path

    $parts = $providerLocation.NamespacePath -split "\\", 4
    Write-Verbose "Provider is located on $($providerLocation.Machine) in namespace $($parts[3])"
    $global:sccmServer = $providerLocation.Machine
    $global:sccmNamespace = $parts[3]
    $global:sccmSiteCode = $providerLocation.SiteCode


     # Make sure we can get a connection

    $global:sccmConnection = [wmi]"${providerLocation.NamespacePath}"
    Write-Verbose "Successfully connected to the specified provider"
}

Function Get-SCCMObject {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(Position=1)] $class, 
        [Parameter(Position=2)] $filter
    )

    if ($filter -eq $null -or $filter -eq "")
    {
        get-wmiobject -class $class -computername $sccmServer -namespace $sccmNamespace
    }
    else
    {
        get-wmiobject -query "select * from $class where $filter" -computername $sccmServer -namespace $sccmNamespace
    }
}


Function Import-SCCMDriverStore
{
	PARAM
    (
    [Parameter(Position=1)] $DriverStore,
    [Parameter(Position=3)] $CMPackageSource,
		#[Parameter(Position=4)] $PackageDepth,
		#[Parameter(Position=5)] $FolderDepth = ($packageDepth - 1),
		#[Parameter(Position=6)] $NameDepth = 1,
		[switch] $cleanup
    )


	
	if ($cleanup)
    {
		$currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )
		if (!$currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator ))
		{
			Write-Custom "You need to run Powershell as Administrator, to use the -Mirror switch." Red
			return;
		}
	
	}

	Write-Headline "Started Importing Driver Store: $($driverStore)"

    $path = Get-ChildItem $driverStore
    	
	Get-ChildItem $driverStore | ? {$_.psIsContainer -eq $true} | % {
	
		$global:CurrentDepth = 1

		SDS-ProcessFolder $_
		
	
	}
	
}

Function Create-RecursiveTxtFile
{
    foreach($Manufacturer in $Manufacturers)
    {
    #$Manufacturer = $Manufacturers[0]  #For testing
    $Models = Get-ChildItem -Directory -Path $Manufacturer.FullName
        foreach($Model in $Models)
        {
        # $Model = $Models[0]  # For testing
   
            $OSArchitectures = Get-ChildItem -Directory -Path $Model.FullName
            foreach($OSArchitecture in $OSArchitectures)
            {
                $filename = $model.Name + ".txt"
                $directories = Get-ChildItem -Directory -Path $OSArchitecture.FullName -Recurse
                foreach($directory in $directories)
                {
                New-Item -Path $directory.FullName -Name $filename -ItemType File -Force
                }
            }
        }
    }
}

Function SDS-ProcessFolder($path)
{
  # Go through each main OS folder
  # For test $OSPath = $path[0]
    foreach($OSPath in $path)
    {
        if($OSPath -like "WINPE")
        {
           Write-Host "Folder is WinPE - Skipping"
        }
        else
        {
        
            #$OSPath = $path ## Added for troubleshooting
            # Create the folder structure under Drivers & Driver packages
    	    $FolderPath = $OSPath.FullName.Substring($DriverStore.Length+1, $OSPath.FullName.Length-($DriverStore.Length+1))
      	    $FolderName = $OSPath.FullName.Substring($DriverStore.Length+1, $OSPath.FullName.Length-($DriverStore.Length+1))
    	    Write-Headline "Processing Folder: $($FolderName)"
            # Create the Main folder
    	    $CMDriverPackageFolderID = SDS-Folder $OSPath 0 23
            $CMDriverFolderID = SDS-Folder $OSPath 0 25
       


            # Create each manufacturer folder in Main folder
            $manufacturers = Get-ChildItem $OSPath.FullName -Directory

            foreach($manufacturer in $manufacturers)
            {
                #$manufacturer = $manufacturers[0] # Added for troubleshooting
                $headline = $FolderName + " - " + $manufacturer
                $CMDriverPackageManufacturerFolderID = SDS-Folder $manufacturer $CMDriverPackageFolderID 23
                $CMDriverManufacturerFolderID = SDS-Folder $manufacturer $CMDriverFolderID 25
                Write-Headline "Processing Folder: $($manufacturer)"

                # Create each model folder under each manufacturer
                #$manufacturer
                $models = Get-ChildItem $manufacturer.FullName -Directory
                #$model = $models[0] #for test
                foreach($model in $models)
                {   
                    #$model = $models[0] # Added for troubleshooting
                    $headline =  $FolderName + " - " + $manufacturer + " - " + $model
                    Write-Headline "Processing Folder: $($headline)"
                    $CMDriverPackageModelFolderID = SDS-Folder $model $CMDriverPackageManufacturerFolderID 23
                    $CMDriverModelFolderID = SDS-Folder $model $CMDriverManufacturerFolderID 25
                    SDS-ProcessPackage $OSPath $manufacturer $model $CMDriverModelFolderID $CMDriverPackageModelFolderID
                }

            }
         }
        
    }
<# 
    Create the driver packages
    $DriverManufacturerFolders = Get-ChildItem $path.FullName | ? {$_.psIsContainer -eq $true} 

	Get-ChildItem $path.FullName | ? {$_.psIsContainer -eq $true} | % {
		$CurrentDepth = 2
		$FolderName = $_.FullName.Substring($DriverStore.Length+1, $_.FullName.Length-($DriverStore.Length+1))
		Write-Headline "Processing Folder: $($FolderName)"
		Get-ChildItem $_.FullName | ? {$_.psIsContainer -eq $true} | % {
      $CurrentDepth = 3
      Write-Headline "Processing Model: $($_)"
      $package = $_
      #SDS-ProcessPackage $_ $FolderPath $CMFolderID
    }
	}
#>

  
}

Function SDS-Folder($folder, $parentID, $objecttype)  
{	

        #$folder = $OsPath # Added for troubleshooting
        #$parentID = 0  # Added for troubleshooting		

        #Create the folder under Driver Packages
		$CMFolder = Get-SCCMObject -Class "SMS_ObjectContainerNode" -Filter "Name = `"$($folder.Name)`" AND ParentContainerNodeID = $($parentID) AND ObjectType = $($objecttype)"
        
		
		If ($CMFolder)
		{
			$CMFolderID = $CMFolder.ContainerNodeID
		}
		Else
		{
			$CMFolderID = New-SCCMFolder -FolderName $folder.Name -FolderType $objecttype -ParentFolderID $parentID
			#Write-Custom "Created SCCM folder $($folder.Name) ($($CMFolderID))"
		}
		$CMFolderID
        
}


Function SDS-ProcessPackage($OSPath, $manufacturer, $model, $CMDriverModelFolderID, $CMDriverPackageModelFolderID)
{

    # SDS-ProcessPackage $OSPath $manufacturer $model $CMDriverModelFolderID $CMDriverPackageModelFolderID
    # Create the Packages in the correct subfolder
    $folderPath = $OSpath.Name
    #$folderID = 1677728  # Added for troubleshooting
    $package = $model

	$PackageName = $model.FullName.Substring($DriverStore.Length+1, $model.FullName.Length-($DriverStore.Length+1))
	
	#$PackageName = $PackageName.Substring($NameIndex+1, $PackageName.Length-($NameIndex+1))
	$PackageName = $PackageName.Replace("\","_")
	
	Write-Headline "Processing Driver Package: $($PackageName)"
	$PackageHash = Get-FolderHash $model.FullName
	If (Get-ChildItem $model.FullName -Filter "$($PackageHash).hash")
	{
		Write-Custom "No changes has been made to this Driver Package. Skipping." DarkGray
	}
	Else
	{
	
		$CMCategory = Get-SCCMDriverCategory -categoryName $PackageName
		if ($CMCategory -eq $null)
		{
			$CMCategory = New-SCCMDriverCategory $PackageName
			Write-Custom "Created new driver category $($PackageName)"
		}
		

		$CMPackageTrue = (get-wmiobject -query "Select * from SMS_DriverPackage join SMS_ObjectContainerItem ON SMS_ObjectContainerItem.InstanceKey = SMS_DriverPackage.PackageID WHERE SMS_ObjectContainerItem.ObjectType = 23 AND SMS_ObjectContainerItem.ContainerNodeID = `"$($CMDriverPackageModelFolderID)`" AND SMS_DriverPackage.Name = `"$($PackageName)`"" -computername $sccmServer -namespace $sccmNamespace).SMS_DriverPackage
		if ($CMPackageTrue -eq $null) { $CMPackageTrue = get-wmiobject -query "Select * from SMS_DriverPackage join SMS_ObjectContainerItem ON SMS_ObjectContainerItem.InstanceKey = SMS_DriverPackage.PackageID WHERE SMS_ObjectContainerItem.ObjectType = 23 AND SMS_ObjectContainerItem.ContainerNodeID = `"$($CMDriverPackageModelFolderID)`" AND SMS_DriverPackage.Name = `"$($PackageName)`"" -computername $sccmServer -namespace $sccmNamespace }
		$CMPackage = get-wmiobject -query "Select * from SMS_DriverPackage WHERE SMS_DriverPackage.PackageID = `"$($CMPackageTrue.PackageID)`"" -computername $sccmServer -namespace $sccmNamespace
		
		if ($CMPackage -eq $null)
		{
			Write-Custom "Creating new driver package $($PackageName)"
			$CMPackageSource = "$($CMPackageSource)\$($folderPath)\$($manufacturer)\$($model)"
			#$CMPackageSource = "$($CMPackageSource)\$($PackageName)"
			if (Test-Path $CMPackageSource)
				{
				if((Get-Item $CMPackageSource | %{$_.GetDirectories().Count + $_.GetFiles().Count}) -gt 0)
				{
					if ($cleanup)
					{
						Write-Custom "Folder already exists, removing content" Yellow
						dir $driverPackageSource | remove-item -recurse -force
					}
					else
					{
						Write-Custom "Folder already exists, remove it manually." Red
						return
					}
				}
			}
			else
			{
				$null = MkDir $CMPackageSource
			}
			
			$CMPackage = New-SCCMDriverPackage -name $PackageName -sourcePath $CMPackageSource            
			Move-SCCMObject -TargetFolderID $CMDriverPackageModelFolderID -ObjectID $CMPackage.PackageID -ObjectType 23
            $CMDriverPackageIsNew= $true
           

		}
		else
		{
			Write-Custom "Existing driver package $($PackageName) ($($CMPackage.PackageID)) retrieved." DarkGray
            $CMDriverPackageIsNew= $false
		}
		
		#$CurrentDepth += 1
		
		#$driverPackageContent = get-wmiobject -computername $sccmServer -namespace $sccmNamespace -query "SELECT * FROM SMS_Driver WHERE CI_ID IN (SELECT CTC.CI_ID FROM SMS_CIToContent AS CTC JOIN SMS_PackageToContent AS PTC ON CTC.ContentID=PTC.ContentID JOIN SMS_DriverPackage AS Pkg ON PTC.PackageID=Pkg.PackageID WHERE Pkg.PackageID='$($CMPackage.PackageID)')"
		#Get-ChildItem $package.FullName -Filter *.inf -recurse | Import-SCCMDriver -category $CMCategory -package $CMPackage | % {
		
		
		#}
		
#		$driverinpath = Get-ChildItem $package.FullName -Filter *.inf -recurse | % { SDS-ImportDriver $_ $CMCategory $CMPackage }
		Get-ChildItem $model.FullName -Filter *.inf -recurse | % { SDS-ImportDriver $_ $CMCategory $CMPackage }
#		Get-ChildItem $package.FullName -Filter *.inf -recurse | % { SDS-ImportDriver $_ $CMCategory $CMPackage }
		
		Get-ChildItem $package.FullName -Filter "*.hash"  | Remove-Item 
		$null = New-Item "$($package.FullName)\$($PackageHash).hash" -type file 
	}



 <#Add the driverpackage to DP Group and distribute it, or if the package already exists, update the DP Group

 write-host "Sleeping 1 minute before distributing the package..." -ForegroundColor Yellow
 Start-Sleep -s 60

  $DPGroupQuery = Get-WmiObject -Namespace $sccmNamespace -Class SMS_DistributionPointGroup -Filter "Name='$DistributionPointGroup'"
  if($CMDriverPackageIsNew)
  {
    $DPGroupQuery.AddPackages($CMPackage.PackageID)
  }
  else
  {
    $DPGroupQuery.ReDistributePackage($CMPackage.PackageID)
  }
#>
}

Function SDS-ImportDriver($dv, $category, $package )
{
        #$dv = $driverinpath[0] #added for troubleshooting
        #$category = $CMCategory #added for troubleshooting
        #$package = $CMPackage #added for troubleshooting

		# Split the path
        $driverINF = split-path $dv.FullName -leaf 
        $driverPath = split-path $dv.FullName

        # Create the class objects needed
        $driverClass = [WmiClass]("\\$sccmServer\$($sccmNamespace):SMS_Driver")
        $localizedClass = [WmiClass]("\\$sccmServer\$($sccmNamespace):SMS_CI_LocalizedProperties")

        # Call the CreateFromINF method
        $driver = $null
        $InParams = $driverClass.psbase.GetMethodParameters("CreateFromINF")
        $InParams.DriverPath = $driverPath
        $InParams.INFFile = $driverINF
        try
        {
            $R = $driverClass.PSBase.InvokeMethod("CreateFromINF", $inParams, $Null)

            # Get the display name out of the result
            $driverXML = [XML]$R.Driver.SDMPackageXML
            $displayName = $driverXML.DesiredConfigurationDigest.Driver.Annotation.DisplayName.Text

            # Populate the localized settings to be used with the new driver instance
            $localizedSetting = $localizedClass.psbase.CreateInstance()
            $localizedSetting.LocaleID =  1033 
            $localizedSetting.DisplayName = $displayName
            $localizedSetting.Description = ""
            [System.Management.ManagementObject[]] $localizedSettings += $localizedSetting

            # Create a new driver instance (one tied to the right namespace) and copy the needed 
            # properties to it.
            $driver = $driverClass.CreateInstance()
            $driver.SDMPackageXML = $R.Driver.SDMPackageXML
            $driver.ContentSourcePath = $R.Driver.ContentSourcePath
            $driver.IsEnabled = $true
            $driver.LocalizedInformation = $localizedSettings
            $driver.CategoryInstance_UniqueIDs = @($category.CategoryInstance_UniqueID)

            # Put the driver instance
            $null = $driver.Put()
        
            Write-Custom "New Driver: $($displayName)"
        }
        catch [System.Management.Automation.MethodInvocationException]
        {
            $e = $_.Exception.GetBaseException()
            if ($e.ErrorInformation.ErrorCode -eq 183)
            {
                
                # Look for a match on the CI_UniqueID    
                $query = "select * from SMS_Driver where CI_UniqueID = '" + $e.ErrorInformation.ObjectInfo + "'"
                $driver = get-WMIObject -query $query.Replace("\", "/") -computername $sccmServer -namespace $sccmNamespace         
 
				Write-Custom "Duplicate Driver: $($driver.LocalizedDisplayName)" DarkGray
	
                # Set the category
                if (-not $driver)
                {
                    Write-Custom "`tUnable to import and no existing driver found." Yellow
                    return
                }
                elseif ($driver.CategoryInstance_UniqueIDs -contains $category.CategoryInstance_UniqueID)
                {
                    Write-Verbose "Existing driver is already in the specified category."
                }
                else
                {
                    $driver.CategoryInstance_UniqueIDs += $category.CategoryInstance_UniqueID
                    $null = $driver.Put()
                    Write-Verbose "Adding driver to category"
                }
            }
            else
            {
                Write-Custom "`tUnexpected error, skipping INF $($infFile): $($e.ErrorInformation.Description) $($e.ErrorInformation.ErrorCode)" Yellow
                return
            }
        }
        
        # Hack - for some reason without this we don't get the CollectionID value
		$hack = $driver.PSBase | select * | out-null

        # If a package was specified, add the driver to it
        if ($package -ne $null)
        {
			$driverPackageContent = get-wmiobject -computername $sccmServer -namespace $sccmNamespace -query "SELECT * FROM SMS_Driver WHERE CI_ID IN (SELECT CTC.CI_ID FROM SMS_CIToContent AS CTC JOIN SMS_PackageToContent AS PTC ON CTC.ContentID=PTC.ContentID JOIN SMS_DriverPackage AS Pkg ON PTC.PackageID=Pkg.PackageID WHERE Pkg.PackageID='$($package.PackageID)')"
            
			$doesDriverExist = $driverPackageContent | ? {$_.CI_UniqueID -eq $driver.CI_UniqueID}
			if ($doesDriverExist -eq $null)
			{
				# Add the driver to the package since it's not already there
				Write-Verbose "Adding driver to package"
				$null = Add-SCCMDriverPackageContent -package $package -driver $driver
			}

        }

        # Write the driver object to the pipeline
        #$driver
        Move-SCCMObject -TargetFolderID $CMDriverModelFolderID -ObjectID $driver.CI_UniqueID -ObjectType 25
}

function Add-SCCMDriverPackageContent
{
    [CmdletBinding()]
    PARAM
    (
        [Parameter(Position=1)] $package,
        [Parameter(Position=2, ValueFromPipeline=$true)] $driver
    )

    Process
    {
        # Get the list of content IDs
        $idlist = @()
        $ci = $driver.CI_ID
        
        $i = 1
		$ids = Get-SCCMObject -class SMS_CIToContent -filter "CI_ID = '$ci'"

        if ($ids -eq $null)
        {
            Write-Warning "Warning: Driver not found in SMS_CIToContent"
        }
        foreach ($id in $ids)
        {
            $idlist += $id.ContentID
        }

        # Build a list of content source paths (one entry in the array)
        $sources = @($driver.ContentSourcePath)

        # Invoke the method
        try
        {
            $package.AddDriverContent($idlist, $sources, $false)
        }
        catch [System.Management.Automation.MethodInvocationException]
        {
            $e = $_.Exception.GetBaseException()
            if ($e.ErrorInformation.ErrorCode -eq 1078462229)
            {
                Write-Verbose "Driver is already in the driver package (possibly because there are multiple INFs in the same folder or the driver already was added from a different location): $($e.ErrorInformation.Description)"
            }
        }
    }

}

# Starting point for main script

Create-RecursiveTxtFile
New-SCCMConnection $SCCMPrimarySiteServer $SCCMSiteCode
#New-SCCMConnection CM01 PS1
Import-SCCMDriverStore $SCCMDriverSourcePath $SCCMDriverPackageSourcePath


#New-SCCMConnection CM01 PS1

#Import-SCCMDriverStore "\\CM01\sccmsrc$\OSD\DriverSources" "\\CM01\SCCMSRC$\OSD\DriverPackages"

# Set current drive back to where we were.
Set-Location $PSScriptRoot

