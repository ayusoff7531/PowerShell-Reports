# ====================================================================================
# NAME   : Report-ClusterNotOnDefaultNode.ps1
# 
# AUTHOR : Developer1, Company1
# DATE   : 10/09/2012
# UPDATED: 5/30/2018
#
# EMAIL  : user1@email.com
# Version:2.0
# 
# COMMENT: Calls Execute-sqlcommand, Get-InventoryInfo
#          This function queries SQL servers in parallel processing
#          and find a node that's running two instances instead of just one
# =====================================================================================

Function Report-ClusterNotOnDefaultNode {

    [cmdletbinding()]
    Param
    (
        [string]$outputPath = "C:\Scripts\Output\SQL_Alert\ClusterNotOnDefaultNode",
        [string]$outputFile = "ClusterNotOnDefaultNode.htm",
        [string]$errorFile = "ErrorClusterNotOnDefaultNode.txt",
        [string]$inventoryServer = "$Global:inventoryServer", #can be any db server
        [string]$inventoryDatabase = "$global:inventoryDatabase" #can be another database besides SQL_Inventory

    )


    
Begin{


        #initialize variables and get current date and time
        $startTime = $endTime = $Message = $output= $fname=$ErrorFName=$ErrorMessage=$MGMTcred=$CDPHcred=$e = $null;


        #Check to see if the output path is there
        if(!(test-path "$outputPath")){
            Write-Verbose -Message "In Report-ClusterNotOnDefaultNode - Created $outputPath folder"
            New-Item -ItemType directory -Path "$outputPath" -force | out-null
        }

        $startTime =Get-TimeStamp

        #Construct file name.  Prefix file name with date and time
        #[string]$fname= (Get-Date -f "yyyy-MM-ddHHmmss")
        [string]$date= (Get-Date -f "yyyy-MM-dd")
        #$ErrorFname= Join-Path -Path $outputPath -ChildPath "$($date +$errorFile)"
        $ErrorFName = $outputPath  + '\' + $date + $errorFile

        Write-verbose -Message "In Report-ClusterNotOnDefaultNode - ErrorFName is $ErrorFName"

        #Delete it if it exists
        if(test-path "$ErrorFName"){Remove-Item "$ErrorFName"}

        $fName = $date + $outputFile
        #Write-verbose -Message "In Report-ClusterNotOnDefaultNode - fName is $fName"

        $OutputFPath = $outputPath + '\' + $fname
        Write-Verbose -Message "In Report-ClusterNotOnDefaultNode - OutputFPath is $OutputFPath"

        #Delete if the output file already exists
        if(test-path "$OutputFPath"){Remove-Item "$OutputFPath"}
        ''
        "Starting script ..."
        ''
        'Executing parallel processing, please wait...'

        #this array will compile the node, clusterName, and instanceName from all of the servers
        $node_cluster = @()

        #this array will have the sorted $node_cluster
        $sorted = @()

        #this array will have the node that is running multiple SQL instances that we want to take care of
        $AACluster = @()

        #this array will have the info like the above + DeptContact column
        $AAClusterInfo = @()

        $ErrorMessage = @()
        $allErrors = @()

        #variable that hold the current node name that is being processed
        $node_i = ''

        #variable index is for the array is initialized to 0
        $index = 0

        #variable rowCount is for the total number of row
        $rowCount = 0

        #EDD contact is in the comment in the spreadsheet, so we just hard coded the value here
        $EDDContact = 'company1@email.com'

        #scriptblock that will be used in Invoke-Command
        $sb = {

            param(
	        $server,
	        $InstanceName,
            $strTcpPort = $null
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

                #Get the nodeName, clusterName, instanceName
                $srvDetails = @{}
                $srvDetails.$("Node") = $srv.ComputerNamePhysicalNetBIOS
                $srvDetails.$("ClusterName") = $srv.NetName
                $srvDetails.$("InstanceName") = $srv.InstanceName
                $myobj = New-Object -TypeName psobject -Property $srvDetails

                #return the object
                $myobj

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
            
            Write-Verbose -Message "In Report-ClusterNotOnDefaultNode - Pinging $inventoryServer"
            isServerUp -fqdn $inventoryServer -EA 'Stop' | Out-Null

            
            #Gets host names form SQL inventory dtabase
            Write-Verbose -Message "In Report-ClusterNotOnDefaultNode - Get a list of server names"
            $servers = Execute-sqlcommand -Server "$inventoryServer" -Database $inventoryDatabase -sqlcmd "Get-SQL_Alert_AA_ClusterVirtual" `
            -param1 "@alert_name" -param1val "Cluster Not on Default Node"
            

            #reate MGMT credential
            Write-Verbose -Message "In Report-ClusterNotOnDefaultNode - Getting Domain admin credential using MGMT credential"
            $sqladmin= Get-AdminCred -Domain "MGMT" -InventoryServer "$inventoryServer"-InventoryDatabase $inventoryDatabase -Login $Global:PowerShellUser_name -Password $Global:PowerShellUser_pw -StoredProc 'dbo.Get-DomainAdminCred'
            $MGMTcred = Get-PSCredential -SQLAdmin $sqladmin -EA 'Stop'




            Write-Verbose -Message "In Report-ClusterNotOnDefaultNode - Querying each server for Cluster Not on Default Node"
            foreach ($s in $servers){
                
                If($s.FQDN.trim()){


                    #Execute the Invoke-Command below
                    #script is running locally 
                    If($s.FQDN.Contains($env:computername)){

                        #The script is executed locally
                        $node_cluster += Invoke-Command -ScriptBlock $sb -ArgumentList $($s.FQDN.split('.')[0]), $s.InstanceName.trim(), $s.tcpPort.trim()

                    }
                    Else{
                        
                        $remoteSvr = $s.FQDN.trim()
                        Write-Verbose -Message "In Report-ClusterNotOnDefaultNode - remoteSvr is $remoteSvr"

                        if(Test-Path C:\Scripts\PortQryUI\PortQry.exe){
                            $portResults = C:\Scripts\PortQryUI\PortQry.exe -n $remoteSvr -e 5985 -p TCP
                            $portState = $portResults |  Select-String -Pattern 'TCP port'

                            Write-Verbose -Message "In Report-ClusterNotOnDefaultNode - $portState"

                            If("$portState" -match "LISTENING"){


                                #script is running in CDPH servers
                                If($s.DomainName.trim() -ieq 'CDPH'){

                                    if(!($CDPHcred)){
                                        Write-Verbose -Message "In Report-ClusterNotOnDefaultNode - Getting Domain admin credential for CDPH servers"
                                        $sqladmin= Get-AdminCred -Domain "CDPH" -InventoryServer "$inventoryServer"-InventoryDatabase $inventoryDatabase -Login $Global:PowerShellUser_name -Password $Global:PowerShellUser_pw -StoredProc 'dbo.Get-DomainAdminCred'
                                        $CDPHcred = Get-PSCredential -SQLAdmin $sqladmin -EA 'Stop'
                                    }

                                    Write-Verbose -Message "In Report-ClusterNotOnDefaultNode - Querying CDPH servers, $remoteSvr, for Cluster not on the default node"
                                    Invoke-Command -ComputerName $s.FQDN.trim() -AsJob -Credential $CDPHcred -ScriptBlock $sb -ArgumentList $($s.FQDN.split('.')[0]), $s.InstanceName.trim(), $s.tcpPort.trim()


                                }
                                Else{

                                    #script is running in the rest of the servers
                                    Write-Verbose -Message "In Report-ClusterNotOnDefaultNode - Querying $remoteSvr for Cluster not on the default node"
                                    Invoke-Command -ComputerName $s.FQDN.trim() -AsJob -Credential $MGMTcred -ScriptBlock $sb -ArgumentList $($s.FQDN.split('.')[0]), $s.InstanceName.trim(), $s.tcpPort.trim()

                                }




                            }#end if $portState is Listening

                            Else{

                                if(!$portState){$portState = "DNS issue - Failed to resolve name to IP address" }
                                Write-Verbose -Message "In Report-ClusterNotOnDefaultNode - $portState. Can't run Invoke-Command to query the server."
                                $instance = $s.InstanceName.trim()
                                write-error "$remoteSvr ($instance) - $portState. Can't run Invoke-Command to query the server." -ErrorVariable +e -ErrorAction SilentlyContinue

                            }#end Else


                        }#end if Test-Path



                    }#End Else

                }#End if
      
        
            }#end foreach loop
  
            
            #sleep for 30 seconds
            Write-Verbose -Message "In Report-ClusterNotOnDefaultNode - Start waiting for 30 seconds"
            Start-Sleep -Seconds 30

            $jobs = Microsoft.PowerShell.Core\Get-Job
            #$jobs

            Write-Verbose -Message "In Report-ClusterNotOnDefaultNode - Compiling the output"
            foreach($j in $jobs){
                if($j.State -eq 'Completed'){
                    If($j.HasMoreData){
                        $node_cluster += Receive-Job -Job $j -ErrorAction SilentlyContinue -ErrorVariable +e
                        Remove-Job -Id $j.Id
                        #break
                    }
            
            
                }
                else{
                    
                    
                    $runningtime = ((get-date) -$j.PSBeginTime).TotalSeconds
                    $server = $j.Location

                    Write-Verbose -Message "In Report-ClusterNotOnDefaultNode - $server - PowerShell script alert has timed out ($runningtime seconds). The server is not responding in timely manner."

                    write-error "$server - PowerShell script alert has timed out ($runningtime seconds). The server is not responding in timely manner." -ErrorVariable +e -ErrorAction SilentlyContinue
                    Remove-Job -Id $j.Id -Force
            
                }

            }



            Write-Verbose 'In Report-ClusterNotOnDefaultNode - Uncomment below to see all the nodes, clusterNames, and instanceNames'
            #$node_cluster

            #$jobs | remove-job

           

            if ($e){
                #$allErrors += $e
                $allErrors = $e
            }

            
            
            #If there are results in $node_cluster
            If ($node_cluster){

                #get the sorted object so that it will be easier to compare the nodes
                $sorted = $node_cluster | Sort-Object Node, ClusterName

                Write-verbose 'In Report-ClusterNotOnDefaultNode - Uncomment below to display $sorted below'
                Write-verbose ''
                #$sorted | ft Node, ClusterName, InstanceName

                $rowCount = $sorted.count
            
                Write-Verbose -Message "In Report-ClusterNotOnDefaultNode - Searching for a node that is running multiple instances"
                ''

                do{
                    $node= $($sorted[$index]).Node
                    write-verbose "In Report-ClusterNotOnDefaultNode - Processsing node $node"

                    #compare the first index vs the next one
                    if($($sorted[$index]).Node -eq $($sorted[$index + 1]).Node){

                        if($node_i){
                            #add another clusterName here in an array

                            if(!($array -match $($sorted[$index]).ClusterName)){
                                $array += $($sorted[$index]).ClusterName


                                $AACluster += $sorted[$index]
            
                            }
           
                            if(!($array -match $($sorted[$index + 1]).ClusterName)){
                                $array += $($sorted[$index + 1]).ClusterName

                                $AACluster += $sorted[$index + 1]
                            }

                        }
                        else{
                            $node_i = $($sorted[$index]).Node
                            write-verbose "In Report-ClusterNotOnDefaultNode - Here is `$node_i which is $node_i"

                            $array = @()
            
                            $array += $($sorted[$index]).ClusterName
                            $AACluster += $sorted[$index]

                            $array += $($sorted[$index + 1]).ClusterName
                            $AACluster += $sorted[$index + 1]

                            write-verbose "In Report-ClusterNotOnDefaultNode - `$array has this count $($array.count)"

                        }

                        if($index -eq ($rowCount - 2)){
                            If($node_i){
                                write-verbose "In Report-ClusterNotOnDefaultNode - $node_i has this and we've reached the last one in the sorted array"
                                #$array
                            }

                            $array = $null

                        }


                    }
                    else{ #node[index] is not the same as node[index + 1]
        
                        If($node_i){
                            write-verbose "In Report-ClusterNotOnDefaultNode - $node_i has this (node[index] is not the same as node[index + 1])"
                            #$array
                        }

                        $node_i = ''
                    }




                    $index++

                }while($index -lt ($rowCount - 1))


                #display $AACluster
                If($AACluster){
                    ''
                    write-verbose "In Report-ClusterNotOnDefaultNode - Here is `$AACluster"
                    #$AACluster | select Node, ClusterName, InstanceName
                


                    #Let's create an array of object that has Node, ClusterName, InstanceName, DeptContact
                    #if the Dept is EDD, then DeptContact are company1@emai.com
                    #else DeptContact is from the SQL_InventoryTemp database

                    foreach($AA in $AACluster){
                        
                        $nodeInfo = Get-InventoryInfo $AA.Node

                        #define an empty hash array
                        $details = @{}
                        $details.$("Node") = $AA.Node
                        $details.$("ClusterName") = $AA.ClusterName
                        $details.$("InstanceName")= $AA.InstanceName

                        if($nodeInfo.Dept -eq 'EDD'){
                            $details.$("DeptContact")= $EDDContact
                        }
                        else{
                            $details.$("DeptContact")= $nodeInfo.DeptContact
                        }


                        $myobj=  new-object -typename psobject -property $details
                        $AAClusterInfo += $myobj


                    }#end foreach




                }#end if $AACluster is not null


                


             }#End If
             

             #These are the errors from Invoke-Command           
             if($allErrors){
                
                foreach($e in $allErrors){
                    $ErrorMessage += (get-date).tostring()
                    $ErrorMessage +=$e.ToString()
                    $ErrorMessage +=''

                    Write-verbose -Message "In Report-ClusterNotOnDefaultNode - error - $e"
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
            Write-host -foreground red "Caught an  Exception.  Unable to generate report for Active/Active Cluster running on the same node"
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

            Write-Verbose -Message "In Report-ClusterNotOnDefaultNode - Created error file $ErrorFName"

            $ErrorMessage >> $ErrorFName

            #notepad $Errorfname
            #Start-Sleep -Seconds 5

        }


        
        #If there are Active/Active clusters running on the same node, display to the monitor and create the html report
        if($AAClusterInfo){ 


            Write-Verbose -Message "In Report-ClusterNotOnDefaultNode - Display the Failed SQL Agent Jobs"

            ''
            'Here is the node that is running multiple SQL instances'
            

            $AAClusterInfo | Select Node, ClusterName, InstanceName, DeptContact | Format-table -AutoSize
            
                    
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
            $header = CreateHeader -task "Here is a list of Active/Active Cluster running on the same node for $date" `
            -callingScript $callingScript -startTime $startTime

            #create the footer variable
            $footer = CreateFooter -startTime $startTime -endTime $endTime

            #assign the style
            $a = ApplyStyle1
            

            #create html report
            Write-Verbose -Message "In Report-ClusterNotOnDefaultNode - Creating the html output for Active/Active Cluster running on the same node report"
            $AAClusterInfo  | convertto-html -head $a -body "<H2>Active/Active Cluster running on the same node Alert</H2>" -property Node, ClusterName, InstanceName, DeptContact -Title "Active/Active Cluster running on the same node report" `
            -PreContent $header -PostContent $footer | Set-Content $OutputFPath

        }#end if
            
        "Script has completed."
        "If there are any Active/Active cluster running on the same node, the output file(s) will be in:"
        "    $outputPath"
        ""
        #'Start time is: ' + $startTime
        $eTime = Get-TimeStamp #end time
        #'End time is  : ' + $eTime

        "Total script time was $([math]::Round($((New-TimeSpan -Start $startTime  -End $eTime).totalMinutes),0)) minute/s."
        ''

   }#end End
   


} #end function

#Export-ModuleMember -Function Report-ClusterNotOnDefaultNode