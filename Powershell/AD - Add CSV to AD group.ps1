$laptops = Get-Content "list.csv"

foreach ($laptop in $laptops) {
    $obj = Get-ADComputer $laptop
    Add-ADGroupMember -ID groupname -Members $obj
}