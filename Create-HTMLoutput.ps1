#-------------------------------------------------------------------------------------
#This function will create the precontent, before the table that contains the data
#-------------------------------------------------------------------------------------
function CreateHeader{

    [CmdletBinding()]

    param(
        [string]$task,
        [string]$callingScript,
        [datetime]$startTime

    )

    $text = @"
Script location: $env:computername<br>
Script/Function Name: $callingScript<br>
Start Time: $startTime<br>
Task: $task<br><br>

If contact is not listed, please check the inventory spreadsheet.<br><br>

"@

    #return the text
    $text
}#end function

#-------------------------------------------------------------------------------------
#This function will create the postcontent, after the table that contains the data
#-------------------------------------------------------------------------------------
function CreateFooter{

    [CmdletBinding()]
    param(
        [datetime]$startTime,
        [datetime]$endTime
    )

    $text = @"

<br><br>
End Time:  $endTime.<br> 
Total script time was $([math]::Round($((New-TimeSpan -Start $startTime  -End $endTime).totalMinutes),0)) minute/s.

"@

    #return the text
    $text
}#end function

#-------------------------------------------------------------------------------------
#This function will apply the styles to the webpage. It will be in the <head> tags
#-------------------------------------------------------------------------------------
function ApplyStyle1{
    [CmdletBinding()]
    param()

    $style = @"
<style>
    TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
    TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;padding: 8px;text-align: left; background-color:#dae1fd; color: #000000;}
    TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black; padding: 8px;text-align: left;}
    tr:hover{background-color:#e5e7e9; color: #0066cc;}
    h2{color: #0066cc;}
    body{font-family: Verdana, Helvetica, Arial, sans-serif;}
</style>

<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">

"@

    #return the style text
    $style

}#end function

#-------------------------------------------------------------------------------------
#This function will create all the html tags on the top part of the webpage.
#It is used by Report-MissingBackup & Report-MissingLogBackup
#-------------------------------------------------------------------------------------
function CreateHTMLhead{
    [CmdletBinding()]
    param(
        $subject
    )

    $head = @"
<html>
<head>
<title>$subject</title>

<style>
    a:link, a:visited {
	background-color: #0066cc;
	color: white;
	padding: 8px 25px;
	text-align: center;
	text-decoration: none;
	}
	
	a:hover, a:active {
	background-color: #003399;
	}

    div.backupNotRequired:target { display: block; }
    div.backupNotRequired {display: none;}

    TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
    TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;padding: 8px;text-align: left;background-color:#dae1fd; color: #000000;}
    TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black; padding: 8px;text-align: left;}
    tr:hover{background-color:#e5e7e9; color: #0066cc;}
    h2{color: #0066cc;}
    body{font-family: Verdana, Helvetica, Arial, sans-serif;}
</style>

<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">

</head>

<body>
<h2>$subject</h2>

"@

    #return the head
    $head


}#end function

#-------------------------------------------------------------------------------------
#This function will close the html body
#-------------------------------------------------------------------------------------
function CloseHTMLbody{
    [CmdletBinding()]
    param()

    $closingTags = @"
</body>
</html>
"@

    #return the tags
    $closingTags
}#end function

#-------------------------------------------------------------------------------------
#This function will create a table for missing backup & missing log backup reports.
#The table is for the databases that require to be backed up
#-------------------------------------------------------------------------------------
function CreateRequiredBackupTable{
    [CmdletBinding()]
    param(
        $RequiredBackup
    )


    $tableRequiredBackup = @"
<h3>Please investigate if database backup is required</h3>
    <table>
    <tr>
        <th>Server</th>
        <th>Instance</th>
        <th>DB</th>
        <th>Backup</th>
        <th>LstFlBk</th>
        <th>LstLgBk</th>
        <th>LstDfBk</th>
        <th>RM</th>
        <th>Created</th>
        <th>BackupRequired</th>
    </tr>
    `n`r
"@
    foreach($db in $requiredBackup){
                
        if ($db.IsRequired -eq 'Yes'){
            $temp = '<tr style="color:red;"><td>'
        }
        else{
            $temp = "<tr><td>"
        }

        $temp += $db.Server + "</td><td>" + $db.Instance + "</td><td>" + $db.DB + "</td><td>" + $db.Backup + "</td><td>" + $db.LstFlBk + "</td><td>" + $db.LstLgBk + "</td><td>" + $db.LstDfBk + "</td><td>" + $db.RM + "</td><td>" + $db.Created + "</td><td>" + $db.IsRequired + "</td></tr>"
        $temp += "`n`r"
        $tableRequiredBackup += $temp
        $tableRequiredBackup += ""
    }

    $tableRequiredBackup += '</table>'
    $tableRequiredBackup += "`n`r<br><br>" 

    #return the table
    $tableRequiredBackup


}#end function

#-------------------------------------------------------------------------------------
#This function will create a table for missing backup & missing log backup reports.
#The table is for the databases that DO NOT require to be backed up
#-------------------------------------------------------------------------------------
function CreateNotRequiredBackupTable{
    [CmdletBinding()]
    param(
        $notRequiredBackup
    )


    $tableNotRequiredBackup = @"
<div>
<a href="#backupNotRequired">Show not required backups</a>
<br><br>
<div>

<div id="backupNotRequired" class="backupNotRequired">

<table>
    <tr>
        <th>Server</th>
        <th>Instance</th>
        <th>DB</th>
        <th>Backup</th>
        <th>LstFlBk</th>
        <th>LstLgBk</th>
        <th>LstDfBk</th>
        <th>RM</th>
        <th>Created</th>
        <th>BackupRequired</th>
    </tr>
    `n`r
"@
    foreach($db in $notRequiredBackup){
                
        if ($db.IsRequired -eq 'Yes'){
            $temp = '<tr style="color:red;"><td>'
        }
        else{
            $temp = "<tr><td>"
        }

        $temp += $db.Server + "</td><td>" + $db.Instance + "</td><td>" + $db.DB + "</td><td>" + $db.Backup + "</td><td>" + $db.LstFlBk + "</td><td>" + $db.LstLgBk + "</td><td>" + $db.LstDfBk + "</td><td>" + $db.RM + "</td><td>" + $db.Created + "</td><td>" + $db.IsRequired + "</td></tr>"
        $temp += "`n`r"
        $tableNotRequiredBackup += $temp
        $tableNotRequiredBackup += ""
    }

    $tableNotRequiredBackup += '</table>'
    $tableNotRequiredBackup += "`n`r" 
    $tableNotRequiredBackup += "<br><br>"
    $tableNotRequiredBackup += "`n`r" 
    $tableNotRequiredBackup += '<a href="#">Hide not required backups</a>'
    $tableNotRequiredBackup += "`n`r"
    $tableNotRequiredBackup += "</div>"
    
    #return the table
    $tableNotRequiredBackup



}#end function
