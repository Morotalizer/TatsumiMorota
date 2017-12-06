# If you have other paths you need to delete just add more $path variables  
# and copy the If statements 
 
$path = "c:\USMT"  
 
# Checks if $path exists, if so deletes it and all subfolders and files 
if (Test-Path $path) { 
     
    $path + " Exists" 
    Remove-Item -path $path -recurse 
    Write-host -foregroundcolor Red $path " Deleted" 
    [System.Threading.Thread]::Sleep(10) 
     
    } else { 
     
    Write-host -foregroundcolor Red  $path  " Does not exist" 
     
    }