# ====================================================================================
# NAME   : Report-MissingLogBackup.ps1
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
#          and find any missing log backups greater than 3 days 
# =====================================================================================

Function Report-MissingLogBackup {

    [cmdletbinding()]
    Param
    (
        [string]$outputPath = "C:\Scripts\Output\SQL_Alert\MissingBackups\Log",
        [string]$outputFile = "MissingLogBackups.htm",
        [string]$errorFile = "ErrorMissingLogBackups.txt",
        [string]$inventoryServer = "$Global:inventoryServer", #can be any db server
        [string]$inventoryDatabase = "$global:inventoryDatabase", #can be another database besides SQL_Inventory
        [string]$inventoryDatabaseTemp = "$Global:inventoryDatabaseTemp" #can be another database besides SQL_InventoryTemp
    )


    
Begin{


        #initialize variables and get current date and time
        $startTime = $endTime = $Message = $output= $fname=$ErrorFName=$ErrorMessage=$MGMTcred=$CDPHcred=$e = $null;


        #Check to see if the output path is there
        if(!(test-path "$outputPath")){
            Write-Verbose -Message "In Report-MissingLogBackup - Created $outputPath folder"
            New-Item -ItemType directory -Path "$outputPath" -force | out-null
        }

        $startTime =Get-TimeStamp

        #Construct file name.  Prefix file name with date and time
        #[string]$fname= (Get-Date -f "yyyy-MM-ddHHmmss")
        [string]$date= (Get-Date -f "yyyy-MM-dd")
        #$ErrorFname= Join-Path -Path $outputPath -ChildPath "$($date +$errorFile)"
        $ErrorFName = $outputPath  + '\' + $date + $errorFile
        Write-verbose -Message "In Report-MissingLogBackup - ErrorFName is $ErrorFName"

        #Delete it if it exists
        if(test-path "$ErrorFName"){Remove-Item "$ErrorFName"}

        $fName = $date + $outputFile
        #Write-verbose -Message "In Report-MissingLogBackup - fName is $fName"

        $OutputFPath = $outputPath + '\' + $fname
        Write-Verbose -Message "In Report-MissingLogBackup - OutputFPath is $OutputFPath"

        #Delete if the output file already exists
        if(test-path "$OutputFPath"){Remove-Item "$OutputFPath"}
        ''
        "Starting script ..."
        ''
        'Executing parallel processing, please wait...'

        $dbsMissBackup = @()
        $dbs = @()
        $ErrorMessage = @()
        $allErrors = @()

        #scriptblock that will be used in Invoke-Command
        $sb = {

            param(
	        $server,
	        $InstanceName,
            $strTcpPort = $null
	        )

            try{

                $ErrorActionPreference = 'Stop'

                $dttime = Get-Date "05:00 am" #Set backups schedule date and time
	            $ckdt = $dttime.AddDays(-3) #set number of days backups are missed

                [System.Reflection.assembly]::LoadWithPartialName("Microsoft.SqlServer.Smo") | out-null

                $SIP = $server + '\' + $InstanceName + ',' + $strTcpPort
	            $srv = new-object('Microsoft.SqlServer.Management.Smo.Server') "$SIP"

                #We comment below because some SMO in the servers do NOT have these properties
	            #$srv.ConnectionContext.LoginSecure = $true
	            #$srv.ConnectionContext.StatementTimeout = 0
	            #$srv.ConnectionContext.ApplicationName = "SQLSupport_MissingBackup_Alert"

                $dbsMissBackup = $srv.databases | 
	            Where {
	                ($_.LastLogBackupDate -lt $ckdt) -and ($_.name -ne "model") -and ($_.name -ne "tempdb") -and ($_.RecoveryModel.ToString() -ne "Simple") -and ($_.status -eq 'Normal') 
	            }

                #return the databases that are missing backups
                if($dbsMissBackup){
	                $dbsMissBackup
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
            
            Write-Verbose -Message "In Report-MissingLogBackup - Pinging $inventoryServer"
            isServerUp -fqdn $inventoryServer -EA 'Stop' | Out-Null


            #Gets host names form SQL inventory dtabase
            Write-Verbose -Message "In Report-MissingLogBackup - Get a list of server names"
            $servers = Execute-sqlcommand -Server "$inventoryServer" -Database $inventoryDatabase -sqlcmd "Get-SQL_Alert_Standalone_Cluster_Instance" `
            -Param1 "@alert_name" -Param1Val "Missing Log Backup"

            #Create MGMT credential
            Write-Verbose -Message "In Report-MissingLogBackup - Getting Domain admin credential using MGMT credential"
            $sqladmin= Get-AdminCred -Domain "MGMT" -InventoryServer "$inventoryServer"-InventoryDatabase $inventoryDatabase -Login $Global:PowerShellUser_name -Password $Global:PowerShellUser_pw -StoredProc 'dbo.Get-DomainAdminCred'
            $MGMTcred = Get-PSCredential -SQLAdmin $sqladmin -EA 'Stop'

            

            Write-Verbose -Message "In Report-MissingLogBackup - Querying each server for Missing Log Backup"
            foreach ($s in $servers){
                
                If($s.FQDN.trim()){


                    #Execute the Invoke-Command below
                    #script is running locally
                    If($s.FQDN.Contains($env:computername)){

                        #The script is executed locally
                        $dbsMissBackup += Invoke-Command -ScriptBlock $sb -ArgumentList $($s.FQDN.split('.')[0]), $s.InstanceName.trim(), $s.tcpPort.trim()

                    }
                    Else{
                        
                        $remoteSvr = $s.FQDN.trim()
                        Write-Verbose -Message "In Report-MissingLogBackup - remoteSvr is $remoteSvr"

                        if(Test-Path C:\Scripts\PortQryUI\PortQry.exe){
                            $portResults = C:\Scripts\PortQryUI\PortQry.exe -n $remoteSvr -e 5985 -p TCP
                            $portState = $portResults |  Select-String -Pattern 'TCP port'

                            Write-Verbose -Message "In Report-MissingLogBackup - $portState"

                            If("$portState" -match "LISTENING"){


                                #script is running in CDPH servers
                                If($s.DomainName.trim() -ieq 'CDPH'){

                                    if(!($CDPHcred)){
                                        Write-Verbose -Message "In Report-MissingLogBackup - Getting Domain admin credential for CDPH servers"
                                        $sqladmin= Get-AdminCred -Domain "CDPH" -InventoryServer "$inventoryServer"-InventoryDatabase $inventoryDatabase -Login $Global:PowerShellUser_name -Password $Global:PowerShellUser_pw -StoredProc 'dbo.Get-DomainAdminCred'
                                        $CDPHcred = Get-PSCredential -SQLAdmin $sqladmin -EA 'Stop'
                                    }

                                    Write-Verbose -Message "In Report-MissingLogBackup - Querying CDPH servers, $remoteSvr, for Missing Log Backup"
                                    Invoke-Command -ComputerName $s.FQDN.trim() -AsJob -Credential $CDPHcred -ScriptBlock $sb -ArgumentList $($s.FQDN.split('.')[0]), $s.InstanceName.trim(), $s.tcpPort.trim()


                                }
                                Else{

                                    #script is running in the rest of the servers
                                    Write-Verbose -Message "In Report-MissingLogBackup - Querying $remoteSvr for Missing Log Backup"
                                    Invoke-Command -ComputerName $s.FQDN.trim() -AsJob -Credential $MGMTcred -ScriptBlock $sb -ArgumentList $($s.FQDN.split('.')[0]), $s.InstanceName.trim(), $s.tcpPort.trim()

                                }




                            }#end if $portState is Listening

                            Else{

                                if(!$portState){$portState = "DNS issue - Failed to resolve name to IP address" }
                                Write-Verbose -Message "In Report-MissingLogBackup - $portState. Can't run Invoke-Command to query the server."
                                $instance = $s.InstanceName.trim()
                                write-error "$remoteSvr ($instance) - $portState. Can't run Invoke-Command to query the server." -ErrorVariable +e -ErrorAction SilentlyContinue

                            }#end Else


                        }#end if Test-Path



                    }#End Else

                }#End if
      
        
            }#end foreach loop
  
            
            #sleep for 30 seconds
            Write-Verbose -Message "In Report-MissingLogBackup - Start waiting for 30 seconds"
            Start-Sleep -Seconds 30

            $jobs = Microsoft.PowerShell.Core\Get-Job
            #$jobs

            Write-Verbose -Message "In Report-MissingLogBackup - Compiling the output"
            foreach($j in $jobs){
                if($j.State -eq 'Completed'){
                    If($j.HasMoreData){
                        $dbsMissBackup += Receive-Job -Job $j -ErrorAction SilentlyContinue -ErrorVariable +e
                        Remove-Job -Id $j.Id
                        #break
                    }
            
            
                }
                else{
                    
                    
                    $runningtime = ((get-date) -$j.PSBeginTime).TotalSeconds
                    $server = $j.Location

                    Write-Verbose -Message "In Report-MissingLogBackup - $server - PowerShell script alert has timed out ($runningtime seconds). The server is not responding in timely manner."

                    write-error "$server - PowerShell script alert has timed out ($runningtime seconds). The server is not responding in timely manner." -ErrorVariable +e -ErrorAction SilentlyContinue
                    Remove-Job -Id $j.Id -Force
            
                }

            }



            <#
            #Loop to see if each job state is still running
            $running=$true
            do{
                foreach($j in $jobs){
                    if($j.State -eq 'Running'){
                        Write-Verbose -Message "In Report-MissingLogBackup - job is $j.State"
                        $running=$true
                        sleep -seconds 5
                        break
            
                    }
                    else{
                        $running=$false
            
                    }

                }

            }while($running -eq $true)
            

            Write-Verbose "In Report-MissingLogBackup - compiling any missing log backup from all the servers"
            $dbsMissBackup += Receive-Job -Job $jobs -ErrorAction SilentlyContinue -ErrorVariable +e
            #>

            Write-Verbose 'In Report-MissingLogBackup - Uncomment below to see a list of missing log backup databases'
            #$dbsMissBackup
            
            #$jobs | remove-job



            if ($e){
                #$allErrors += $e
                $allErrors = $e
            }

            
            
            #If there are results in $dbsMissBackup
            If ($dbsMissBackup){
            
                Write-Verbose -Message "In Report-MissingLogBackup - Query CMS server using Get-InventoryInfo for each database that has missing log backups"
            
                #Process each missing db backup in dbsMissBackup                           
                ForEach ($db in $dbsMissBackup){
                    $server= $db.Parent

                    if ($server){
                    $temp = $server.Replace('[','')
                    $temp2 = $temp.Replace(']','')
                    $serverInstance = $temp2.split(',')[0]
                    $serverName = $serverInstance.split('\')[0]
                    $instance = $serverInstance.split('\')[1]

                    

                    #Query the CMS server to see if the backup is required for each database in $dbsMissBackup
                    $Backup = (Execute-sqlcommand -Server "$inventoryServer" -Database "$inventoryDatabaseTemp" -sqlcmd "Get-LogBackupExclusion" `
                    -Param1 "@hostname" -Param1Val $serverName -Param2 "@dbname" -Param2Val $($db.name)).bkup

                    If($Backup -eq 'Not required'){
                        $isRequired = 'No'
                    }
                    elseif($Backup -eq 'Required'){
                        $isRequired = 'Yes'
                    }
                    else{
                        $isRequired = $Backup
                    }

                    

                    #define empty hash array
                
                    $dbDetails = @{};
                    $dbDetails.$("Server") = $serverName
                    $dbDetails.$("Instance") = $instance
                    $dbDetails.$("DB") = $db.name
                    $dbDetails.$("Backup")= (get-inventoryinfo $serverName).Backup
                    $dbDetails.$("LstFlBk") = if($db.LastBackupDate -ne "1/1/0001 12:00:00 AM"){get-date ($db.LastBackupDate) -format d} else {"None"}
                    $dbDetails.$("LstLgBk") = if($db.LastLogBackupDate -ne "1/1/0001 12:00:00 AM"){get-date ($db.LastLogBackupDate) -Format d} else {"None"}
                    $dbDetails.$("LstDfBk") = if($db.LastDifferentialBackupDate -ne "1/1/0001 12:00:00 AM"){get-date ($db.LastDifferentialBackupDate) -Format d} else {"None"}
                    $dbDetails.$("RM")= $db.RecoveryModel
                    $dbDetails.$("Created")= if($db.Createdate -ne "1/1/0001 12:00:00 AM"){get-date ($db.Createdate) -Format d} else {"Unknown"}
                    $dbDetails.$("IsRequired")= $isRequired

		            $myobj=  new-object -typename psobject -property $dbDetails
                    $dbs += $myobj

                    }
                
		                    
                }#end ForEach 

             }#End If
             

             #These are the errors from Invoke-Command           
             if($allErrors){
                
                foreach($e in $allErrors){
                    $ErrorMessage += (get-date).tostring()
                    $ErrorMessage +=$e.ToString()
                    $ErrorMessage +=''

                    Write-verbose -Message "In Report-MissingLogBackup - error - $e"
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
            Write-host -foreground red "Caught an  Exception.  Unable to get Missing backup alert"
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

            Write-Verbose -Message "In Report-MissingLogBackup - Created error file $ErrorFName"
            $ErrorMessage >> $ErrorFName

            #notepad $Errorfname
            #Start-Sleep -Seconds 5

        }


        
        #If there are missing log backups, display to the monitor and create the html report
        if($dbs){ 


            Write-Verbose -Message "In Report-MissingLogBackup - Display the missing log backup to the monitor"
            $dbs | Select Server,Instance, DB, Backup, LstFlBk, LstLgBk, LstDfBk, RM, Created, IsRequired | Format-table -AutoSize
            
                    
            #$endTime = (Get-Date).tostring(); #set end time\
            $endTime = Get-TimeStamp

        
            #prepare header (top of the <body> tag)

            If ($MyInvocation.ScriptName){
                $callingScript = $MyInvocation.ScriptName
            }
            else{
                $callingScript = $MyInvocation.MyCommand.Name
            }


            #assign to a header variable
            $header = CreateHeader -task 'Returns databases missing backups for more than three days' `
            -callingScript $callingScript -startTime $startTime

            #Create a footer (in the <body> tag) and assign to a variable
            $footer = CreateFooter -startTime $startTime -endTime $endTime

            #Creating the html tag, head tag, title tag, opening body tag
            $htmlHead = CreateHTMLhead -subject "Missing Log Backup Alert"

            #Creating the closing html tag, closing body tag
            $closingTags = CloseHTMLbody
            
            #create two variables based on the backup requirement
            $requiredBackup = $dbs | where{$_.IsRequired -ne 'No'}
            $notRequiredBackup = $dbs | where{$_.IsRequired -eq 'No'}


            #create the table for databases that require backups
            if($requiredBackup){
                $tableRequiredBackup = CreateRequiredBackupTable -RequiredBackup $requiredBackup
            }

            #create the table for databases that do NOT require backups
            If($notRequiredBackup){
                $tableNotRequiredBackup = CreateNotRequiredBackupTable -NotRequiredBackup $notRequiredBackup
            }

            #create html report
            Write-Verbose -Message "In Report-MissingLogBackup - Creating the html output for Missing Log Backup report"

            #string concatenation
            $body = $htmlHead + $header + $tableRequiredBackup + $tableNotRequiredBackup + $footer + $closingTags

            #create the html file
            $body | Set-Content $OutputFPath

        }#end if
            
        "Script has completed."
        "If there are any missing log backup, the output file(s) will be in:"
        "    $outputPath"
        ""
        #'Start time is: ' + $startTime
        $eTime = Get-TimeStamp #end time
        #'End time is  : ' + $eTime

        "Total script time was $([math]::Round($((New-TimeSpan -Start $startTime  -End $eTime).totalMinutes),0)) minute/s."
        ''

   }#end End
   


} #end function

#Export-ModuleMember -Function Report-MissingLogBackup