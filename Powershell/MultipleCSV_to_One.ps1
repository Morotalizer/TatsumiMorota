#Get the date and time and define the input/output paths
$date_time = Get-Date -Format 'yyyy-MM-dd'
$files = Get-ChildItem -Path "\\stopsccm01\logs$\PST\" -Depth 0 -Filter "*.csv"
$outfilepath = ("E:\temp\" + $date_time + "_pstoverview.csv")

#Create headers in CSV file
"Timestamp;Computer;Path;Size;Outlook version;User ID;User name;User description;User location;User department;User email;User company;User title;Manager ID;Manager email" | Out-File -FilePath $outfilepath -Append

#Import each CSV file into big CSV file
foreach ($file in $files) {
    $file | Get-Content | Out-File -FilePath $outfilepath -Append
}

#Import big CSV file in order to get userinfo from AD
$csv = Import-Csv $outfilepath -Delimiter ';'
$numberoflines = $csv.Count
$lastuserID = $null
$i = 1

#For each line in the big CSV, get the user name and query AD for additional info
foreach ($line in $csv) {
    $userID = $line.'User ID'
    write-host "Processing line $i of $numberoflines - User: $userID"
    $i++
    #Skip repeat user names
    if (($userID -ne $lastuserID) -and ($userID -ne $null)) {
        write-host "New user: $userID ... querying AD"
        #Resetting variables to avoid duplicates in case the queries fail 
        $username = $null
        $userdepartment = $null
        $userdescription = $null
        $userlocation = $null
        $usermanagerSAMID = $null
        $usermanagerEMAIL = $null
        $useremail = $null
        $usercompany = $null
        $usertitle = $null

        #Query AD and save variables
        $username = ((dsquery user -samid $userID | dsget user -display)[1]).Trim()
        $userdepartment = ((dsquery user -samid $userID | dsget user -dept)[1]).Trim()
        $userdescription = ((dsquery user -samid $userID | dsget user -desc)[1]).Trim()
        $userlocation = ((dsquery user -samid $userID | dsget user -office)[1]).Trim()
        $usermanager = (dsquery user -samid $userID | dsget user -mgr)[1].Trim()
        if ($usermanager -ne "") { 
            $usermanagerSAMID =  (($usermanager | dsget user -samid)[1]).Trim()
            $usermanagerEMAIL =  (($usermanager | dsget user -email)[1]).Trim()
        }
        $useremail = ((dsquery user -samid $userID | dsget user -email)[1]).Trim()
        $usercompany = ((dsquery user -samid $userID | dsget user -company)[1]).Trim()
        $usertitle = ((dsquery user -samid $userID | dsget user -title)[1]).Trim()
    } elseif (($userID -eq $lastuserID) -and ($userID -ne $null)) {
        write-host "Same user...skipped query"
    }
    else {
        write-host "No user found...skipped query"
    }

    $line.'User name' = $username
    $line.'User department' = $userdepartment
    $line.'User description' = $userdescription
    $line.'User location' = $userlocation
    $line.'Manager ID' = $usermanagerSAMID
    $line.'Manager email' = $usermanagerEMAIL
    $line.'User email' = $useremail
    $line.'User company' = $usercompany
    $line.'User title' = $usertitle

    $lastuserID = $userID
}

#Export CSV object to file
$csv | Export-Csv c:\temp\out.csv -Delimiter ';'