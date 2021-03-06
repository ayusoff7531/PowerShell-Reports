# ====================================================================================
# NAME   : Report-FailedSQLAgentJob.ps1
# 
# AUTHOR : Developer1, Company1
# DATE   : 10/09/2012
# UPDATED: 4/30/2018
#
# EMAIL  : user1@email.com
# Version:2.0
# 
# COMMENT: Calls Execute-sqlcommand, Get-InventoryInfo
#          This function queries SQL servers in parallel processing
#          and find any failed SQL agent jobs
# =====================================================================================

Function Report-FailedSQLAgentJob {

    [cmdletbinding()]
    Param
    (
        [string]$outputPath = "C:\Scripts\Output\SQL_Alert\FailedAgentJob",
        [string]$outputFile = "FailedAgentJob.htm",
        [string]$errorFile = "ErrorFailedAgentJob.txt",
        [string]$jobName = "OTECH_*",
        [string]$inventoryServer = "$Global:inventoryServer", #can be db server
        [string]$inventoryDatabase = "$global:inventoryDatabase" #can be another database besides SQL_Inventory

    )


    
Begin{


        #initialize variables and get current date and time
        $startTime = $endTime = $Message = $output= $fname=$ErrorFName=$ErrorMessage=$MGMTcred=$CDPHcred=$e = $null;


        #Check to see if the output path is there
        if(!(test-path "$outputPath")){
            Write-Verbose -Message "In Report-FailedSQLAgentJob - Created $outputPath folder"
            New-Item -ItemType directory -Path "$outputPath" -force | out-null
        }

        $startTime =Get-TimeStamp

        #Construct file name.  Prefix file name with date and time
        #[string]$fname= (Get-Date -f "yyyy-MM-ddHHmmss")
        [string]$date= (Get-Date -f "yyyy-MM-dd")
        #$ErrorFname= Join-Path -Path $outputPath -ChildPath "$($date +$errorFile)"
        $ErrorFName = $outputPath  + '\' + $date + $errorFile

        Write-verbose -Message "In Report-FailedSQLAgentJob - ErrorFName is $ErrorFName"

        #Delete it if it exists
        if(test-path "$ErrorFName"){Remove-Item "$ErrorFName"}

        $fName = $date + $outputFile
        #Write-verbose -Message "In Report-FailedSQLAgentJob - fName is $fName"

        $OutputFPath = $outputPath + '\' + $fname
        Write-Verbose -Message "In Report-FailedSQLAgentJob - OutputFPath is $OutputFPath"

        #Delete if the output file already exists
        if(test-path "$OutputFPath"){Remove-Item "$OutputFPath"}
        ''
        "Starting script ..."
        ''
        'Executing parallel processing, please wait...'

        $failedAgentJobs = @()
        $failedJobs = @()
        $ErrorMessage = @()
        $allErrors = @()

        #scriptblock that will be used in Invoke-Command
        $sb = {

            param(
	        $server,
	        $InstanceName,
            $strTcpPort = $null,
            $jobName = "OTECH_*"
	        )

            try{

                $ErrorActionPreference = 'Stop'

                [System.Reflection.assembly]::LoadWithPartialName("Microsoft.SqlServer.Smo") | out-null

                $SIP = $server + '\' + $InstanceName + ',' + $strTcpPort
	            $srv = new-object('Microsoft.SqlServer.Management.Smo.Server') "$SIP"

                #We comment below because some SMO in the servers do NOT have these properties
	            #$srv.ConnectionContext.LoginSecure = $true
	            #$srv.ConnectionContext.StatementTimeout = 0
	            #$srv.ConnectionContext.ApplicationName = "SQLSupport_MissingBackup_Alert"

                #Get all the jobs
                $jobs  = $srv.Jobserver.Jobs

                $failedJobs = $jobs | 
                Where {
                    #$_.Name -like $jobName will be like OTECH_INTEGRITY_CHECK_USER_DATABASES 
                    #$_.LastRunDate = at least the job has run once
                    ($_.Name -like "$jobName") -and ($_.LastRunOutcome -eq 'Failed') -and ($_.IsEnabled -eq 'True') -and ($_.LastRunDate)
                }

                #return the failed SQL agent jobs
                if($failedJobs){
	                $failedJobs
                }
            }#end try

            catch{

                $err = "$env:computername - $_.Exception"
                throw $err #throw the error
            }

        } #end scriptblock
        

} #end Begin
        
Process{       
        Try {
            $ErrorActionPreference = "Stop" # stop if error occurs 
            
            Write-Verbose -Message "In Report-FailedSQLAgentJob - Pinging $inventoryServer"
            isServerUp -fqdn $inventoryServer -EA 'Stop' | Out-Null

            
            #Gets host names form SQL inventory dtabase
            Write-Verbose -Message "In Report-FailedSQLAgentJob - Get a list of server names"
            $servers = Execute-sqlcommand -Server "$inventoryServer" -Database $inventoryDatabase -sqlcmd "Get-SQL_Alert_Standalone_Cluster_Instance" `
            -Param1 "@alert_name" -Param1Val "Missing Full Backup"
            

            #reate MGMT credential
            Write-Verbose -Message "In Report-FailedSQLAgentJob - Getting Domain admin credential using MGMT credential"
            $sqladmin= Get-AdminCred -Domain "MGMT" -InventoryServer "$inventoryServer"-InventoryDatabase $inventoryDatabase -Login $Global:PowerShellUser_name -Password $Global:PowerShellUser_pw -StoredProc 'dbo.Get-DomainAdminCred'
            $MGMTcred = Get-PSCredential -SQLAdmin $sqladmin -EA 'Stop'




            Write-Verbose -Message "In Report-FailedSQLAgentJob - Querying each server for Missing Full Backup"
            foreach ($s in $servers){
                
                If($s.FQDN.trim()){


                    #Execute the Invoke-Command below
                    #script is running locally
                    If($s.FQDN.Contains($env:computername)){

                        #The script is executed locally
                        $failedAgentJobs += Invoke-Command -ScriptBlock $sb -ArgumentList $($s.FQDN.split('.')[0]), $s.InstanceName.trim(), $s.tcpPort.trim()

                    }
                    Else{
                        
                        $remoteSvr = $s.FQDN.trim()
                        Write-Verbose -Message "In Report-FailedSQLAgentJob - remoteSvr is $remoteSvr"

                        if(Test-Path C:\Scripts\PortQryUI\PortQry.exe){
                            $portResults = C:\Scripts\PortQryUI\PortQry.exe -n $remoteSvr -e 5985 -p TCP
                            $portState = $portResults |  Select-String -Pattern 'TCP port'

                            Write-Verbose -Message "In Report-FailedSQLAgentJob - $portState"

                            If("$portState" -match "LISTENING"){


                                #script is running in CDPH servers
                                If($s.DomainName.trim() -ieq 'CDPH'){

                                    if(!($CDPHcred)){
                                        Write-Verbose -Message "In Report-FailedSQLAgentJob - Getting Domain admin credential for CDPH servers"
                                        $sqladmin= Get-AdminCred -Domain "CDPH" -InventoryServer "$inventoryServer"-InventoryDatabase $inventoryDatabase -Login $Global:PowerShellUser_name -Password $Global:PowerShellUser_pw -StoredProc 'dbo.Get-DomainAdminCred'
                                        $CDPHcred = Get-PSCredential -SQLAdmin $sqladmin -EA 'Stop'
                                    }

                                    Write-Verbose -Message "In Report-FailedSQLAgentJob - Querying CDPH servers, $remoteSvr, for Failed SQL Agent Job"
                                    Invoke-Command -ComputerName $s.FQDN.trim() -AsJob -Credential $CDPHcred -ScriptBlock $sb -ArgumentList $($s.FQDN.split('.')[0]), $s.InstanceName.trim(), $s.tcpPort.trim()


                                }
                                Else{

                                    #script is running in the rest of the servers
                                    Write-Verbose -Message "In Report-FailedSQLAgentJob - Querying $remoteSvr for Failed SQL Agent Job"
                                    Invoke-Command -ComputerName $s.FQDN.trim() -AsJob -Credential $MGMTcred -ScriptBlock $sb -ArgumentList $($s.FQDN.split('.')[0]), $s.InstanceName.trim(), $s.tcpPort.trim()

                                }




                            }#end if $portState is Listening

                            Else{

                                if(!$portState){$portState = "DNS issue - Failed to resolve name to IP address" }
                                Write-Verbose -Message "In Report-FailedSQLAgentJob - $portState. Can't run Invoke-Command to query the server."
                                $instance = $s.InstanceName.trim()
                                write-error "$remoteSvr ($instance) - $portState. Can't run Invoke-Command to query the server." -ErrorVariable +e -ErrorAction SilentlyContinue

                            }#end Else


                        }#end if Test-Path



                    }#End Else

                }#End if
      
        
            }#end foreach loop
  
            
            #sleep for 30 seconds
            Write-Verbose -Message "In Report-FailedSQLAgentJob - Start waiting for 30 seconds"
            Start-Sleep -Seconds 30

            $jobs = Microsoft.PowerShell.Core\Get-Job
            #$jobs

            Write-Verbose -Message "In Report-FailedSQLAgentJob - Compiling the output"
            foreach($j in $jobs){
                if($j.State -eq 'Completed'){
                    If($j.HasMoreData){
                        $failedAgentJobs += Receive-Job -Job $j -ErrorAction SilentlyContinue -ErrorVariable +e
                        Remove-Job -Id $j.Id
                        #break
                    }
            
            
                }
                else{
                    
                    
                    $runningtime = ((get-date) -$j.PSBeginTime).TotalSeconds
                    $server = $j.Location

                    Write-Verbose -Message "In Report-FailedSQLAgentJob - $server - PowerShell script alert has timed out ($runningtime seconds). The server is not responding in timely manner."

                    write-error "$server - PowerShell script alert has timed out ($runningtime seconds). The server is not responding in timely manner." -ErrorVariable +e -ErrorAction SilentlyContinue
                    Remove-Job -Id $j.Id -Force
            
                }

            }



            Write-Verbose 'In Report-FailedSQLAgentJob - Uncomment below to see a list of failed SQL agent jobs'
            #$failedAgentJobs

            #$jobs | remove-job

           

            if ($e){
                #$allErrors += $e
                $allErrors = $e
            }

            
            
            #If there are results in $failedAgentJobs
            If ($failedAgentJobs){
            
                Write-Verbose -Message "In Report-FailedSQLAgentJob - Creating custom PS object for Failed SQL Agent Jobs"
            
                #Process each missing db backup in dbsMissBackup                           
                ForEach ($f in $failedAgentJobs){
                    $server= $f.Parent

                    if ($server){
                    $temp = $server.Replace('[','')
                    $temp2 = $temp.Replace(']','')
                    $serverInstance = $temp2.split(',')[0]
                    $serverName = $serverInstance.split('\')[0]
                    $instance = $serverInstance.split('\')[1]
                    

                    #define empty hash array
                
                    $jobDetails = @{};
                    $jobDetails.$("Server") = $serverName
                    $jobDetails.$("Instance") = $instance
                    $jobDetails.$("Name") = $f.Name
                    $jobDetails.$("LastRunOutcome")= $f.LastRunOutcome
                    $jobDetails.$("LastRunDate")= $f.LastRunDate
                    

		            $myobj=  new-object -typename psobject -property $jobDetails
                    $failedJobs += $myobj

                    }
                
		                    
                }#end ForEach 


             }#End If
             

             #These are the errors from Invoke-Command           
             if($allErrors){
                
                foreach($e in $allErrors){
                    $ErrorMessage += (get-date).tostring()
                    $ErrorMessage +=$e.ToString()
                    $ErrorMessage +=''

                    Write-verbose -Message "In Report-FailedSQLAgentJob - error - $e"
                }

                
             }



                
            }#end Try
                    
        # capture all predefined, common, system runtime exceptions.       
        Catch [system.exception] {
           $ErrorMessage += @"                            
$(Get-TimeStamp):
$(Get-TimeStamp): -- SCRIPT PROCESSING CANCELLED
$(Get-TimeStamp): $('-' * 50)
$(Get-TimeStamp): $($NBN)
$(Get-TimeStamp): Error in $($_.InvocationInfo.ScriptName).
$(Get-TimeStamp): Line Number: $($_.InvocationInfo.ScriptLineNumber)
$(Get-TimeStamp): Offset: $($_.InvocationInfo.OffsetInLine)
$(Get-TimeStamp): Command: $($_.InvocationInfo.MyCommand)
$(Get-TimeStamp): Line: $($_.InvocationInfo.Line.Trim())
$(Get-TimeStamp): Error Details: $($_)
$(Get-TimeStamp): 
$('-' * 100)

"@
        
            #$ErrorMessage += "$($NBN) :  Getting Stopped SQL related services ... `r`n $($_ | Out-String)"#error log message
            Write-host -foreground red "Caught an  Exception.  Unable to get Failed SQL Agent Jobs"
            Write-debug ($_ | Out-String);
            }#end Catch
          
        #reset EA and display message that script has ended
        finally {
        
                 $ErrorActionPreference = "Continue";
                 $NBN=$null
                 "Ended running script on all of the servers"
                 
        }#end Finally
        
}#end Process
End{
    
    
        if($ErrorMessage){

            Write-Verbose -Message "In Report-FailedSQLAgentJob - Created error file $ErrorFName"

            $ErrorMessage >> $ErrorFName

            #notepad $Errorfname
            #Start-Sleep -Seconds 5

        }


        
        #If there are failed SQL agent jobs, display to the monitor and create the html report
        if($failedJobs){ 


            Write-Verbose -Message "In Report-FailedSQLAgentJob - Display the Failed SQL Agent Jobs"
            $failedJobs | Select Server,Instance, Name, LastRunOutcome, LastRunDate | Format-table -AutoSize
            
                    
            #$endTime = (Get-Date).tostring(); #set end time\
            $endTime = Get-TimeStamp

        
            #prepare header (top of the <body> tag)

            If ($MyInvocation.ScriptName){
                $callingScript = $MyInvocation.ScriptName
            }
            else{
                $callingScript = $MyInvocation.MyCommand.Name
            }


            $date = Get-Date -Format "MM/dd/yyyy"

            #assign the header to the header variable
            $header = CreateHeader -task "Here is a list of failed SQL Agent Jobs for $date" `
            -callingScript $callingScript -startTime $startTime

            #create the footer variable
            $footer = CreateFooter -startTime $startTime -endTime $endTime

            #assign the style
            $a = ApplyStyle1
            

            #create html report
            Write-Verbose -Message "In Report-FailedSQLAgentJob - Creating the html output for Failed SQL Agent Jobs report"
            $failedJobs  | convertto-html -head $a -body "<H2>Failed SQL Agent Jobs Alert</H2>" -property Server,Instance, Name, LastRunOutcome, LastRunDate -Title "Failed SQL Agent Jobs Report" `
            -PreContent $header -PostContent $footer | Set-Content $OutputFPath

        }#end if
            
        "Script has completed."
        "If there are any failed SQL agent job, the output file(s) will be in:"
        "    $outputPath"
        ""
        #'Start time is: ' + $startTime
        $eTime = Get-TimeStamp #end time
        #'End time is  : ' + $eTime

        "Total script time was $([math]::Round($((New-TimeSpan -Start $startTime  -End $eTime).totalMinutes),0)) minute/s."
        ''

   }#end End
   


} #end function

#Export-ModuleMember -Function Report-FailedSQLAgentJob