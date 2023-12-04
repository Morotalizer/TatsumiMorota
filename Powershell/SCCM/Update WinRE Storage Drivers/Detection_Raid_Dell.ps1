# Specify the namespace and class name
# The value output to be compliant will be RAID 
$namespace = "root\dcim\sysman\biosattributes"
$className = "EnumerationAttribute"

# Query WMI to get the required attribute value
try {
    $attributeValue = Get-WmiObject -Namespace $namespace -Class $className | Where-Object { $_.AttributeName -eq 'EMBSataRaid' } | Select-Object -ExpandProperty CurrentValue

    if ($attributeValue -ne $null) {
        Write-Host "$attributeValue"
    } else {
        Write-Host "EMBSataRaid attribute not found."
    }
} catch {
    Write-Host "Error retrieving EMBSataRaid attribute: $_"
}