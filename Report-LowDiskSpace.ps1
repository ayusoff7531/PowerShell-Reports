# ====================================================================================
# NAME  : Report-LowDiskSpace.ps1
# 
# AUTHOR: Developer1, Company1
# DATE  : 03/20/2015
# UPDATED: 4/30/2018
#
# EMAIL : user1@email.com
# Version:2.0
# 
# COMMENT: Gets drives with free space < 10% for all drives and < 500MB for C drive
#          drive space info to the output stream and to a text file 
#
# NOTE: From observation, memory usage is low when this script creates the output (less than 1%)
#       CPU usage peaks about 40% within 10 seconds when this script creates the output files in MRESCMSSQL3
# =====================================================================================
Function Report-LowDiskSpace {

    [cmdletbinding()]
    Param
    (
        [string]$outputPath = "C:\Scripts\Output\SQL_Alert\LowDiskSpace",
        [string]$outputFile = "LowDiskSpace.htm",
        [string]$errorFile = "ErrorLowDiskSpace.txt",
        [string]$inventoryServer = "$Global:inventoryServer", #can be any db server
        [string]$inventoryDatabase = "$global:inventoryDatabase" #can be another database besides SQL_Inventory
    )
    
Begin{

        

        #initialize variables and get current date and time
        $startTime = $endTime = $Message = $output= $fname=$ErrorFName=$ErrorMessage=$MGMTcred=$CDPHcred=$e = $null;


        #Check to see if the output path is there
        if(!(test-path "$outputPath")){
            Write-Verbose -Message "In Report-LowDiskSpace - Created $outputPath folder"
            New-Item -ItemType directory -Path "$outputPath" -force | out-null
        }

        $startTime =Get-TimeStamp
        
        #Construct file name.  Prefix file name with date and time
        #[string]$fname= (Get-Date -f "yyyy-MM-ddHHmmss")
        [string]$date= (Get-Date -f "yyyy-MM-dd")
        #$ErrorFname= Join-Path -Path $outputPath -ChildPath "$($date +$errorFile)"
        $ErrorFName = $outputPath  + '\' + $date + $errorFile
        Write-verbose -Message "In Report-LowDiskSpace - ErrorFName is $ErrorFName"

        #Delete it if it exists
        if(test-path "$ErrorFName"){Remove-Item "$ErrorFName"}

        $fName = $date + $outputFile
        #Write-verbose -Message "In Report-LowDiskSpace - fName is $fName"

        $OutputFPath = $outputPath + '\' + $fname
        Write-Verbose -Message "In Report-LowDiskSpace - OutputFPath is $OutputFPath"

        #Delete if the output file already exists
        if(test-path "$OutputFPath"){Remove-Item "$OutputFPath"}
        ''
        "Starting script ..."
        ''
        'Executing parallel processing, please wait...'

        $driveInfo = @()
        $HDS = @()
        $allErrors = @()
        $ErrorMessage = @()




        $sb = {
            
            try{

                $ErrorActionPreference = 'Stop'

                $minC = 524288000 #minimum free space of C drive in bytes (equal to 500MB)
                $minOtherDrive = 10 #minimum free space of drives not equal to C in free percentage (equal to 10%)

                $lowDisks = @()

                $disks = gwmi -class 'win32_volume' -Filter "DriveType=3"
                $disks = $disks | where-Object {($_.Name -notlike "B*") -and ($_.Name -notlike "A*") -and ($_.Name -notlike "\\?\*")}

                foreach($d in $disks){
                    if (($d.Name -eq 'C:\' -and $d.freespace -le $minC) -or ($d.Name -ne 'C:\' -and ($d.freespace/$d.Capacity * 100) -le $minOtherDrive)){
                        $lowDisks += $d
                    }
                }
                if($lowDisks){
                    #return the disks that are below the threshold
                    $lowDisks

                }

            }#end try

            catch{
                $err = "$env:computername - $_.Exception"
                throw $err #throw the error

            }#end catch

        }#end $sb
        
        
        
}#end Begin

Process{ 
      
     Try{
            $ErrorActionPreference = "Stop" # stop if error occurs 

            Write-Verbose -Message "In Report-LowDiskSpace - Pinging $inventoryServer"
            isServerUp -fqdn $inventoryServer -EA 'Stop' | Out-Null


            #Gets host names form SQL inventory dtabase
            $hosts = Execute-sqlcommand -Server "$inventoryServer" -Database $inventoryDatabase -sqlcmd "Get-SQL_Alert_Standalone_Cluster_Node" -param1 "@alert_name" -param1val "Low Disk Space"
            

            #Create MGMT credential
            Write-Verbose -Message "In Report-LowDiskSpace - Getting Domain admin credential using MGMT credential"
            $sqladmin= Get-AdminCred -Domain "MGMT" -InventoryServer "$inventoryServer"-InventoryDatabase $inventoryDatabase -Login $Global:PowerShellUser_name -Password $Global:PowerShellUser_pw -StoredProc 'dbo.Get-DomainAdminCred'
            $MGMTcred = Get-PSCredential -SQLAdmin $sqladmin -EA 'Stop'


            Write-Verbose -Message "In Report-LowDiskSpace - Querying each server for low disk space"
            foreach ($h in $hosts){

                #script is running locally
                If($h.FQDN.Contains($env:computername)){

                    #The script is executed locally
                    $HDS += Invoke-Command -ScriptBlock $sb

                }
                Else{

                    $remoteSvr = $h.FQDN.trim()
                    Write-Verbose -Message "In Report-LowDiskSpace - remoteSvr is $remoteSvr"

                    if(Test-Path C:\Scripts\PortQryUI\PortQry.exe){
                            $portResults = C:\Scripts\PortQryUI\PortQry.exe -n $remoteSvr -e 5985 -p TCP
                            $portState = $portResults |  Select-String -Pattern 'TCP port'

                            Write-Verbose -Message "In Report-LowDiskSpace - $portState"

                            If("$portState" -match "LISTENING"){


                                #script is running in CDPH servers
                                If($h.DomainName.trim() -ieq 'CDPH'){

                                    if(!($CDPHcred)){
                                        Write-Verbose -Message "In Report-LowDiskSpace - Getting Domain admin credential for CDPH servers"
                                        $sqladmin= Get-AdminCred -Domain "CDPH" -InventoryServer "$inventoryServer"-InventoryDatabase $inventoryDatabase -Login $Global:PowerShellUser_name -Password $Global:PowerShellUser_pw -StoredProc 'dbo.Get-DomainAdminCred'
                                        $CDPHcred = Get-PSCredential -SQLAdmin $sqladmin -EA 'Stop'
                                    }

                                    Write-Verbose -Message "In Report-LowDiskSpace - Querying CDPH servers, $remoteSvr, for low disk space"
                                    Invoke-Command -ComputerName $h.FQDN.trim() -AsJob -Credential $CDPHcred -ScriptBlock $sb


                                }
                                Else{

                                    #script is running in the rest of the servers
                                    Write-Verbose -Message "In Report-LowDiskSpace - Querying $remoteSvr for low disk space"
                                    Invoke-Command -ComputerName $h.FQDN.trim() -AsJob -Credential $MGMTcred -ScriptBlock $sb

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
        

            #sleep for 30 seconds
            Write-Verbose -Message "In Report-LowDiskSpace - Start waiting for 30 seconds"
            Start-Sleep -Seconds 30

            $jobs = Microsoft.PowerShell.Core\Get-Job
            #$jobs

            Write-Verbose -Message "In Report-LowDiskSpace - Compiling the output"
            foreach($j in $jobs){
                if($j.State -eq 'Completed'){
                    If($j.HasMoreData){
                        $HDS += Receive-Job -Job $j -ErrorAction SilentlyContinue -ErrorVariable +e
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



            Write-Verbose 'In Report-LowDiskSpace - Uncomment below to see disks with low free space'
            #$HDS



            if ($e){
                #$allErrors += $e
                $allErrors = $e
            }
            
            
            
            #If there is results in $HDS
            If ($HDS){
            
                Write-Verbose -Message "In Report-LowDiskSpace - Query CMS server using Get-InventoryInfo for each server that has drive with low freespace"
            
                ForEach ($drive in $HDS){
                    $serverName = $drive.SystemName


                    #If $serverName is still empty or null, do the extra step
                    If(-not $serverName){$serverName = $drive.__SERVER}
                    If(-not $serverName){
                        $sName = ""
                        $sName = $drive.PSComputerName

                        if ($sName){
                            $splittedName = $sName.Split(".")
                            $serverName = $splittedName[0]
                        }

                    }

                    if ($serverName){
                       
                        Write-Verbose -Message "In Report-LowDiskSpace - Get-InventoryInfo on $drive.name in $serverName"

                        $serverInfo = Get-InventoryInfo $serverName

                        If($serverInfo.Type.Contains("Cluster")){
                            $clusterID = $serverInfo.ClusterMirrorID
                            $clusterIDBackupType = (get-inventoryInfo $clusterID).Backup
                            $BackupType = $clusterIDBackupType
                        }
                        Else{
                            $clusterID = "NA"
                            $BackupType = $serverInfo.Backup
                        }

                        #define empty hash array
                        $details = @{};
                        $details.$("Server") = $serverName
                        $details.$("ClusterID") = $clusterID
                        $details.$("BackupType")= $BackupType
                        $details.$("Drive") = $Drive.name
                        $details.$("Size_GB") = [math]::round($($($drive.Capacity)/1GB),2) 
                        $details.$("FreeSpace_GB") = [math]::round($($($drive.freespace)/1GB),2)
                        $details.$("FreePct") = [math]::round($($details.$("FreeSpace_GB") / $details.$("Size_GB") * 100),2)
                        $details.$("Contact")= $serverInfo.DeptContact
                        $details.$("Email_Phone")= $serverInfo.DeptPhoneEmail

		                $myobj=  new-object -typename psobject -property $details
                        $driveInfo += $myobj


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
                    
     #capture all predefined, common, system runtime exceptions.       
     Catch [system.exception] {
           $ErrorMessage += @"                            
$(Get-TimeStamp):
$(Get-TimeStamp): -- SCRIPT PROCESSING CANCELLED
$(Get-TimeStamp): $('-' * 50)
#$(Get-TimeStamp): $($hostName)
$(Get-TimeStamp): Error in $($_.InvocationInfo.ScriptName).
$(Get-TimeStamp): Line Number: $($_.InvocationInfo.ScriptLineNumber)
$(Get-TimeStamp): Offset: $($_.InvocationInfo.OffsetInLine)
$(Get-TimeStamp): Command: $($_.InvocationInfo.MyCommand)
$(Get-TimeStamp): Line: $($_.InvocationInfo.Line.Trim())
$(Get-TimeStamp): Error Details: $($_)
$(Get-TimeStamp): 
$('-' * 100)

"@     
     
            
            Write-host -foreground red "Caught an  Exception.  Unable to get low disk space info. " 
            Write-host ($_ | Out-String);
    }#end Catch

   #reset EA and display message that script has ended
   finally {
   
        $ErrorActionPreference = "Continue";
        "Ended running script on all of the servers"
        
   }#end Finally
   
}#end Process  
   
    End{
  


        
        if($ErrorMessage){

            Write-Verbose -Message "In Report-LowDiskSpace - Created error file $ErrorFName"
            $ErrorMessage >> $ErrorFName

            #notepad $Errorfname
            #Start-Sleep -Seconds 5

        }


        
        #If there are drives that have low free space, display to the monitor and create the html report
        if($driveInfo){ 

            #Write-Verbose -Message "In Report-LowDiskSpace - Creating an array of lowDrive after filtering the driveInfo array"
            #$lowDrive= $driveInfo |  ?{(($_.drive -eq 'C:\' -and $_.FreeSpace_GB  -le 0.5) -or ($_.drive -ne 'C:\' -and $_.FreePct -le 10)) } 
            Write-Verbose -Message "In Report-LowDiskSpace - Assign the variable driveInfo to lowDrive"
            $lowDrive = $driveInfo

            Write-Verbose -Message "In Report-LowDiskSpace - Display the lowDrive to the monitor"
            $lowDrive | Select Server, ClusterID, Drive, size_GB, FreeSpace_GB, FreePct, BackupType, Contact, Email_Phone | Format-table
                        
            #$endTime = (Get-Date).tostring(); #set end time\
            $endTime = Get-TimeStamp

        
            #prepare header

            If ($MyInvocation.ScriptName){
                $callingScript = $MyInvocation.ScriptName
            }
            else{
                $callingScript = $MyInvocation.MyCommand.Name
            }


            #assign the header to the header variable
            $header = CreateHeader -task 'Returns drives with free space < 10% for drives other than the C:\ drive.  Treshold for C:\drive is 500MB.' `
            -callingScript $callingScript -startTime $startTime

            #create the footer variable
            $footer = CreateFooter -startTime $startTime -endTime $endTime

            #assign the style
            $a = ApplyStyle1
            
            
            #create html report
            Write-Verbose -Message "In Report-LowDiskSpace - Creating the html output for LowDiskSpace report"
            $lowDrive  | convertto-html -head $a -body "<H2>Low Drive Alert</H2>" -property Server, ClusterID, Drive, size_GB, FreeSpace_GB, FreePct, BackupType, Contact, Email_Phone -Title "Low Drive Report" `
            -PreContent $header -PostContent $footer | Set-Content $OutputFPath
            
            
       
        }#end if
        
        
        "Script has completed."
        "If there are any low disk space issues, the output file(s) will be in:"
        "    $outputPath"
        ""
        #'Start time is: ' + $startTime
        $eTime = Get-TimeStamp #end time
        #'End time is  : ' + $eTime

        "Total script time was $([math]::Round($((New-TimeSpan -Start $startTime  -End $eTime).totalMinutes),0)) minute/s."
        ''
        
    }#end End 


} 

#Export-ModuleMember -Function Report-LowDiskSpace


