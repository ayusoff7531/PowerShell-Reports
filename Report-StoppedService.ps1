﻿# ====================================================================================
# NAME  : Report-StoppedService.ps1
# 
# AUTHOR: Developer1, Company1
# DATE  : 10/09/2012
# UPDATED: 4/30/2018
#
# EMAIL : user1@email.com
# Version:2.0
# 
# COMMENT: Calls Execute-sqlcommand
#          Checks stopped services with start mode of Auto
#          or manual
# =====================================================================================

Function Report-StoppedService {

    [cmdletbinding()]
    Param
    (
        [string]$outputPath = "C:\Scripts\Output\SQL_Alert\StoppedService",
        [string]$outputFile = "StoppedService.htm",
        [string]$errorFile = "ErrorStoppedService.txt",
        [string]$inventoryServer = "$Global:inventoryServer", #can be any db server
        [string]$inventoryDatabase = "$global:inventoryDatabase" #can be another database besides SQL_Inventory
    )


    
Begin{


        #initialize variables and get current date and time
        $startTime = $endTime = $Message = $output= $fname=$ErrorFName=$ErrorMessage=$MGMTcred=$CDPHcred=$e = $null;


        #Check to see if the output path is there
        if(!(test-path "$outputPath")){
            Write-Verbose -Message "In Report-StoppedService - Created $outputPath folder"
            New-Item -ItemType directory -Path "$outputPath" -force | out-null
        }

        $startTime =Get-TimeStamp

        #Construct file name.  Prefix file name with date and time
        #[string]$fname= (Get-Date -f "yyyy-MM-ddHHmmss")
        [string]$date= (Get-Date -f "yyyy-MM-dd")
        #$ErrorFname= Join-Path -Path $outputPath -ChildPath "$($date +$errorFile)"
        $ErrorFName = $outputPath  + '\' + $date + $errorFile
        Write-verbose -Message "In Report-StoppedService - ErrorFName is $ErrorFName"

        #Delete it if it exists
        if(test-path "$ErrorFName"){Remove-Item "$ErrorFName"}

        $fName = $date + $outputFile
        #Write-verbose -Message "In Report-StoppedService - fName is $fName"

        $OutputFPath = $outputPath + '\' + $fname
        Write-Verbose -Message "In Report-StoppedService - OutputFPath is $OutputFPath"

        #Delete if the output file already exists
        if(test-path "$OutputFPath"){Remove-Item "$OutputFPath"}
        ''
        "Starting script ..."
        ''
        'Executing parallel processing, please wait...'

        $stoppedServices= @()
        $SQLstopped = @()
        $ErrorMessage = @()
        $allErrors = @()

        #scriptblock that will be used in Invoke-Command
        $sb = {

            param(
                $CT = "NA" #cluster type
            )

            try{
                
                $ErrorActionPreference = 'Stop'

                $stoppedSQLServices = gwmi -class "Win32_Service" -Filter "(DisplayName like 'SQL%' OR DisplayName like 'MySQL%') and (State = 'Stopped') and (StartMode = 'Manual' OR StartMode = 'Auto') and Not (displayname Like '%Distributed Replay%' OR displayname Like '%Full-Text%' OR displayname Like '%CEIP%')"
                $count = $stoppedSQLServices.count

                if($CT -eq "AA"){

                    if ($count -ge 6){
                
                        #return stoppedServices
                        $stoppedSQLServices
                    
                    }#end inner if
                    
                 }else{
                
                    if($stoppedSQLServices){
                        #return any SQL services that are stopped
                        $stoppedSQLServices

                    }
               
                 }#end if/else

             }#end try

             catch{
                $err = "$env:computername - $_.Exception"
                throw $err #throw the error
             }#end catch


        } #end scriptblock
        

} #end Begin
        
Process{       
        Try {
            $ErrorActionPreference = "Stop" # stop if error occurs 
            
            Write-Verbose -Message "In Report-StoppedService - Pinging $inventoryServer"
            isServerUp -fqdn $inventoryServer -EA 'Stop' | Out-Null


            #Gets host names form SQL inventory dtabase
            Write-Verbose -Message "In Report-StoppedService - Get a list of server names"
            $servers = Execute-sqlcommand -Server "$inventoryServer" -Database $inventoryDatabase `
            -sqlcmd "Get-SQL_Alert_Standalone_ClusterVirtual" -Param1 "@alert_name" -Param1Val "Stopped Service"

            #Create MGMT credential
            Write-Verbose -Message "In Report-StoppedService - Getting Domain admin credential using MGMT credential"
            $sqladmin= Get-AdminCred -Domain "MGMT" -InventoryServer "$inventoryServer"-InventoryDatabase $inventoryDatabase -Login $Global:PowerShellUser_name -Password $Global:PowerShellUser_pw -StoredProc 'dbo.Get-DomainAdminCred'
            $MGMTcred = Get-PSCredential -SQLAdmin $sqladmin -EA 'Stop'

            

            Write-Verbose -Message "In Report-StoppedService - Querying each server for SQL stopped service"
            foreach ($s in $servers){
                
                
               
                #Get the cluster type for each server, but check if it is null first
                If (!($s.ClusterType.trim())){
                    #Initialize $CT (ClusterType) in case there are null ones from the query above from CMS server
                    $CT = "NA"
                }
                Else{
                    $CT = $s.ClusterType.trim()
                }



                #Execute the Invoke-Command below
                #script is running locally
                If($s.FQDN.Contains($env:computername)){

                    #The script is executed locally
                    $stoppedServices += Invoke-Command -ScriptBlock $sb -ArgumentList $CT

                }
                Else{
                    
                    $remoteSvr = $s.FQDN.trim()
                    Write-Verbose -Message "In Report-StoppedService - remoteSvr is $remoteSvr"

                    if(Test-Path C:\Scripts\PortQryUI\PortQry.exe){
                            $portResults = C:\Scripts\PortQryUI\PortQry.exe -n $remoteSvr -e 5985 -p TCP
                            $portState = $portResults |  Select-String -Pattern 'TCP port'

                            Write-Verbose -Message "In Report-StoppedService - $portState"

                            If("$portState" -match "LISTENING"){


                                #script is running in CDPH servers
                                If($s.DomainName.trim() -ieq 'CDPH'){

                                    if(!($CDPHcred)){
                                        Write-Verbose -Message "In Report-StoppedService - Getting Domain admin credential for CDPH servers"
                                        $sqladmin= Get-AdminCred -Domain "CDPH" -InventoryServer "$inventoryServer"-InventoryDatabase $inventoryDatabase -Login $Global:PowerShellUser_name -Password $Global:PowerShellUser_pw -StoredProc 'dbo.Get-DomainAdminCred'
                                        $CDPHcred = Get-PSCredential -SQLAdmin $sqladmin -EA 'Stop'
                                    }

                                    Write-Verbose -Message "In Report-StoppedService - Querying CDPH servers, $remoteSvr, for stopped SQL services"
                                    Invoke-Command -ComputerName $s.FQDN.trim() -AsJob -Credential $CDPHcred -ScriptBlock $sb -ArgumentList $CT


                                }
                                Else{

                                    #script is running in the rest of the servers
                                    Write-Verbose -Message "In Report-StoppedService - Querying $remoteSvr for stopped SQL services"
                                    Invoke-Command -ComputerName $s.FQDN.trim() -AsJob -Credential $MGMTcred -ScriptBlock $sb -ArgumentList $CT

                                }




                            }#end if $portState is Listening

                            Else{

                                if(!$portState){$portState = "DNS issue - Failed to resolve name to IP address" }
                                Write-Verbose -Message "In Report-LowDiskSpace - $portState. Can't run Invoke-Command to query the server."
                                
                                write-error "$remoteSvr - $portState. Can't run Invoke-Command to query the server." -ErrorVariable +e -ErrorAction SilentlyContinue

                            }#end Else


                        }#end if Test-Path



                }#End Else
      
        
            }#end foreach loop
  

            #sleep for 15 seconds
            Write-Verbose -Message "In Report-LowDiskSpace - Start waiting for 15 seconds"
            Start-Sleep -Seconds 15

            $jobs = Microsoft.PowerShell.Core\Get-Job
            #$jobs

            Write-Verbose -Message "In Report-LowDiskSpace - Compiling the output"
            foreach($j in $jobs){
                if($j.State -eq 'Completed'){
                    If($j.HasMoreData){
                        $stoppedServices += Receive-Job -Job $j -ErrorAction SilentlyContinue -ErrorVariable +e
                        Remove-Job -Id $j.Id
                        #break
                    }
            
            
                }
                else{
                    
                    
                    $runningtime = ((get-date) -$j.PSBeginTime).TotalSeconds
                    $server = $j.Location

                    Write-Verbose -Message "In Report-LowDiskSpace - $server - PowerShell script alert has timed out ($runningtime seconds). The server is not responding in timely manner."

                    write-error "$server - PowerShell script alert has timed out ($runningtime seconds). The server is not responding in timely manner." -ErrorVariable +e -ErrorAction SilentlyContinue
                    Remove-Job -Id $j.Id -Force
            
                }

            }
          

            Write-Verbose 'In Report-StoppedService - Uncomment below to see any stopped SQL services'
            #$stoppedServices



            if ($e){
                #$allErrors += $e
                $allErrors = $e
            }

            
            #We're not going to call Get-InventoryInfo for the stoppedServices
            #If there are results in $stoppedServices
            If ($stoppedServices){
            
                ForEach ($service in $stoppedServices){
                    $serverName = $service.PSComputerName.Split('.')[0]


                    #If $serverName is still empty or null, do the extra step
                    If(-not $serverName){$serverName = $service.__SERVER}


                    if ($serverName){
                       
                        Write-Verbose -Message "In Report-StoppedService - Creating customer PS object for $serverName"

                        #define empty hash array
                        $details = @{};
                        $details.$("Server") = $serverName
                        $details.$("SystemName") = $service.SystemName
                        $details.$("Name") = $service.name
                        $details.$("StartName") = $service.StartName
                        $details.$("Started") = $service.Started
                        $details.$("StartMode") = $service.StartMode
                        $details.$("State")= $service.State
                        

		                $myobj=  new-object -typename psobject -property $details
                        $SQLstopped += $myobj


		            }

                 }#end ForEach 

             }#End If
             

             #These are the errors from Invoke-Command           
             if($allErrors){
                
                foreach($e in $allErrors){
                    $ErrorMessage += (get-date).tostring()
                    $ErrorMessage +=$e.ToString()
                    $ErrorMessage +=''

                    Write-verbose -Message "In Report-LowDiskSpace - error - $e"
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
            Write-host -foreground red "Caught an  Exception.  Unable to get stopped services"
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

            Write-Verbose -Message "In Report-StoppedService - Created error file $ErrorFName"
            $ErrorMessage >> $ErrorFName

            #notepad $Errorfname
            #Start-Sleep -Seconds 5

        }


        
        #If there are SQL services that are stopped, display to the monitor and create the html report
        if($SQLstopped){ 


            Write-Verbose -Message "In Report-StoppedService - Display the stopped SQL services to the monitor"
            $SQLstopped | Select Server, SystemName, Name, StartName, Started, StartMode, State | Format-table -AutoSize
                        
            #$endTime = (Get-Date).tostring(); #set end time\
            $endTime = Get-TimeStamp

        
            #prepare header

            If ($MyInvocation.ScriptName){
                $callingScript = $MyInvocation.ScriptName
            }
            else{
                $callingScript = $MyInvocation.MyCommand.Name
            }


            #Create the header variable
            $header = CreateHeader -task 'Gets stopped services that have start mode Auto or Manual' `
            -callingScript $callingScript -startTime $startTime
	    
	    $header += '<i class="fa fa-exclamation-triangle" style="font-size:32px;color:red"> Please check the server!</i><br><br>' + "`n`r"

            #Create the footer variable
            $footer = CreateFooter -startTime $startTime -endTime $endTime

            #create the style variable
            $a = ApplyStyle1
            
            
            #create html report
            Write-Verbose -Message "In Report-StoppedService - Creating the html output for Stopped SQL Service report"
            $SQLstopped  | convertto-html -head $a -body "<H2>Stopped SQL Service Alert</H2>" -property Server, SystemName, Name, StartName, Started, StartMode, State -Title "Stopped SQL Service Report" `
            -PreContent $header -PostContent $footer | Set-Content $OutputFPath

        }#end if

        "Script has completed."
        "If there are any stopped SQL services, the output file(s) will be in:"
        "    $outputPath"
        ""
        #'Start time is: ' + $startTime
        $eTime = Get-TimeStamp #end time
        #'End time is  : ' + $eTime

        "Total script time was $([math]::Round($((New-TimeSpan -Start $startTime  -End $eTime).totalMinutes),0)) minute/s."
        ''
        
   }#end End
   


} #end function

#Export-ModuleMember -Function Report-StoppedService
