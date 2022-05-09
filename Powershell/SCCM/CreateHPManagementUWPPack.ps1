<#  Creator @gwblok - GARYTOWN.COM
    Used to download DriverPack Updates from HP, then Extract the bin file.
    This Script was created to build a DriverPack Update Package. 
    Future Scripts based on this will be one that gets the Model / Product info from the machine it's running on and pull down the correct DriverPack and run the Updater

    REQUIREMENTS:  HP Client Management Script Library
    Download / Installer: https://ftp.hp.com/pub/caps-softpaq/cmit/hp-cmsl.html  
    Docs: https://developers.hp.com/hp-client-management/doc/client-management-script-library-0
    This Script was created using version 1.1.1
#>

#Reset Vars
$DriverPack = ""
$Model = ""
$HPModelsTable = ""
$HPModelName = ""
$CurrentDownloadedVersion = ""

$OS = "Win10"
$Category = "UWPPack"
#$HPModels = @("80FC", "82CA")
$DownloadDir = "F:\Temp\HPManagement\Download"
$ExtractedDir = "F:\Temp\HPManagement\Extract"
    
        $HPModelsTable= @(
        @{ ProdCode = '83b2'; Model = "EliteBook 840 G5 UWP"}
        @{ ProdCode = '8846'; Model = "ZBook FireFly 15 G8 UWP"}
        @{ ProdCode = '8723'; Model = "EliteBook 840 G7 UWP"  }
        @{ ProdCode = '8549'; Model = "EliteBook 840 G6 UWP"  }
        @{ ProdCode = '8595'; Model = "EliteDesk 800 G5 Desktop Mini UWP"  }
        @{ ProdCode = '8079'; Model = "EliteBook 840 G3"  }
        @{ ProdCode = '829A'; Model = "EliteDesk 800 G3 UWP"  }
        @{ ProdCode = '845A'; Model = "EliteDesk 800 G4 UWP"  }
        @{ ProdCode = '8594'; Model = "EliteDesk 800 G5 UWP"  }
        @{ ProdCode = '8724'; Model = "ZBook FireFly 14 G7 UWP"  }
        #@{ ProdCode = '880D'; Model = "EliteBook 840 G8"  }
        #@{ ProdCode = '8873'; Model = "ZBook Studio 15.6 G8"  }
        @{ ProdCode = '8725'; Model = "X360 830 G7"  }
        @{ ProdCode = '8736'; Model = "ZBook Studio G7"  }
        @{ ProdCode = '83B3'; Model = "EliteBook 830 G5 UWP"  }
        @{ ProdCode = '853D'; Model = "EliteBook X360 UWP"  }        
        @{ ProdCode = '83B3'; Model = "EliteBook 830 G5 UWP"  }
        @{ ProdCode = '8711'; Model = "EliteDesk 800 G6 UWP"  }#>
        )
        $HPModelName = $HPModelsTable | ? ProdCode -eq $Model | % Model

foreach ($Model in $HPModelsTable)
    {
    Write-Output "Checking Product Code $($Model.ProdCode) for DriverPack Updates"
    #$DriverPack = Get-SoftpaqList -platform $Model.ProdCode -os $OS -category $Category
    $DriverPack = Get-SoftpaqList -platform $Model.ProdCode -category $Category
    If(!$DriverPack){
        Write-Output "No Managemenapack for $($Model.ProdCode)"
        }
    else {
    
    if (Test-Path "$($DownloadDir)\$($Model.Model)"){$CurrentDownloadedVersion = (Get-childitem -Path "$($DownloadDir)\$($Model.Model)").Name}
    $MostRecent = ($DriverPack | Measure-Object -Property "ReleaseDate" -Maximum).Maximum
    $DriverPack = $DriverPack | WHERE "ReleaseDate" -eq "$MostRecent"
    $DownloadPath = "$($DownloadDir)\$($Model.Model)\$($DriverPack.id)"
    $ExtractedPath = "$($ExtractedDir)\$($Model.Model)"
    
    if (-not (Test-Path "$($DownloadPath)"))
        {
        if ($CurrentDownloadedVersion) {Write-Output "Update Found, Replacing $([decimal]$CurrentDownloadedVersion) with $($DriverPack.id)"}
        Else {Write-Output "Update Found, Downloading: $($DriverPack.id)"}
        Write-Output "Downloading DriverPack Update for: $($Model.Model) aka $($Model.ProdCode)"
        Get-Softpaq -number $DriverPack.ID -saveAs "$($DownloadPath)\$($DriverPack.id).exe" -Verbose
        Write-Output "Creating Readme file with DriverPack Info HERE: $($DownloadPath)\$($DriverPack.ReleaseDate).txt"
        $DriverPack | Out-File -FilePath "$($DownloadPath)\$($DriverPack.ReleaseDate).txt"
        $DriverPackFileName = Get-ChildItem -Path "$($DownloadPath)\*.exe" | select -ExpandProperty "Name"
        
        if (Test-path $ExtractedPath) 
            {
            Write-Output "Deleting $($ExtractedPath) Contents before extracting new contents"
            remove-item -Path $ExtractedPath -Recurse -Force
            }
        Write-Output "Extracting Downloaded DriverPack File to: $($ExtractedPath)"
        Start-Process "$($DownloadPath)\$($DriverPackFileName)" -ArgumentList "/e /s /f`"$($ExtractedPath)`"" -wait
        #Start-Sleep -Seconds 2
        $DriverPack | Out-File -FilePath "$($ExtractedPath)\$($DriverPack.id).txt"
        # $DriverPack | Out-File -FilePath "$($ExtractedPath)\$([decimal]$DriverPack.id).txt"
        #Start-Sleep -Seconds 2
        #Write-Output "Deleting support files, leaving only the DriverPack.bin file & 64Bit Updater"
        #Remove-Item -Path "$($ExtractedPath)\*.rtf" -Verbose
        #Remove-Item -Path "$($ExtractedPath)\*.log" -Verbose
        #Remove-Item -Path "$($ExtractedPath)\Hpq*.exe" -Verbose
        #Remove-Item -Path "$($ExtractedPath)\HPDriverPackUPDREC.exe" -Verbose
        
        Start-Sleep -Seconds 3
        }
    Else
        {Write-Output "No New DriverPack Available"}
    }}