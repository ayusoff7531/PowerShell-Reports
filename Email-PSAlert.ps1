Function Email-PSAlert {


    [cmdletbinding()]
    Param
    (
        [string]$FunctionName,
        [string]$from = "admin@email.com",
        [string[]]$to = ("user1@email.com",'user2@email.com'),
        [string]$outputPath,
        [string]$outputFile,
        [string]$errorFile,
        [string]$logFile, #can be either $global:Wmilog or $global:SMOlog
        [string]$SmtpServer = "10.10.10.10",
        [string]$subject
    )

    #Help 
  
	<#
    .SYNOPSIS 
    This function will email PowerShell daily reports if they are created
	
    .EXAMPLE
    Here is an example
	Email-PSAlert -FunctionName "Report-LowDiskSpace" -from "admin@email.com" `
    -to "admin@email.com" `
    -outputPath "output:\SQL_Alert\LowDiskSpace" -outputFile "LowDiskSpace.htm" `
    -errorFile "ErrorLowDiskSpace.txt" -logFile $global:Wmilog `
    -subject "*** LOW DISK SPACE ***" -verbose
	
    .EXAMPLE
    Here is a second example
	Email-PSAlert -FunctionName "Report-MissingFullBackup" -from "admin@email.com" `
    -to "admin@email.com" `
    -outputPath "output:\MissingBackups" -outputFile "MissingBkups.txt" `
    -errorFile "ErrorMissingBkups.txt" -logFile $global:SMOlog `
    -subject "*** MISSING FULL BACKUP ***"

	.DESCRIPTION
    This function will call other PowerShell alert report functions such as Report-LowDiskSpace, Report-MissingFullBackup etc
    If the reports create output files, this function will email them to us
    
	
    #>

    Log-Header #error log header
    $message=$jobs=$outputFpath=$date=$ErrorFname=$null
    $message += "`r`nGetting server object ... `r`n$($_ | Out-String)"#error log message
    [string]$date= (Get-Date -f "yyyy-MM-dd")


    $OutputFPath = $outputPath + '\' + $date + $outputFile
    $ErrorFName = $outputPath  + '\' + $date + $errorFile

    Write-Verbose -Message "In Email-PSAlert - OutputFPath is $OutputFPath"
    Write-Verbose -Message "In Email-PSAlert - ErrorFName is $ErrorFName"

    #Array for attachments
    #[array]$attachments = $OutputFpath,$ErrorFname
    $attachments = @()

    Try {

        #Run script agains all Servers 
        #Getting stopped services on all SQL server hosts with no Database engine"
        #Report-LowDiskSpace -outputFile $outputFile -errorFile $errorFile

        if($VerbosePreference -eq "Continue"){
            & $FunctionName -outputPath $outputPath -outputFile $outputFile -errorFile $errorFile -verbose
        }
        Else{
            & $FunctionName -outputPath $outputPath -outputFile $outputFile -errorFile $errorFile
        }

        Start-Sleep -Seconds 5


        #foreach($t in $to){

            If (Test-Path -path "$ErrorFName"){
                $attachments += "$ErrorFName"  
                 Write-Verbose -Message "In Email-PSAlert - Here is the error file - $ErrorFName"
            }

            
            If (Test-Path -path "$OutputFpath"){
                $attachments += "$OutputFpath"
                 Write-Verbose "In Email-PSAlert - Here is the LowDiskSpace report - $OutputFpath"
            }

            
            If ($attachments){
                Send-MailMessage -To $to -From $from -Subject "$subject" `
                -Body "Please see attached file. If script encountered errors, an error file is attached. `nSent from $env:COMPUTERNAME" `
                -SmtpServer $SmtpServer -Attachments $attachments
                

                foreach ($a in $attachments){
                    Write-Verbose "In Email-PSAlert - Attachment(s) - $a"
                }
            }



        #}#end foreach   

    }Catch{

    $message += "`r`n$($_ | Out-String)"
    Write-host  -foreground red "In Email-PSAlert - Unable to generate $FunctionName report";
    Write-debug ($_ | Out-String);
    $global:endTime = (Get-Date).tostring();#save end time
    Log-Footer -msg $message -LogFile $logFile;
            
    }
    Finally{}#end try/cath/finally
                   
}


