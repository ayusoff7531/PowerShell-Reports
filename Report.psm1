#. \\res\Pub\Bin2\SQLServer\Utility\PowerShell\Modules\Test\Workstation_CMS\Report\Report-LowDiskSpace.ps1
#. \\res\Pub\Bin2\SQLServer\Utility\PowerShell\Modules\Test\Workstation_CMS\Report\Email-PSAlert.ps1


. "$PSScriptRoot\Report-LowDiskSpace.ps1"
. "$PSScriptRoot\Email-PSAlert.ps1"
. "$PSScriptRoot\Create-HTMLoutput.ps1"
. "$PSScriptRoot\Report-StoppedService.ps1"
. "$PSScriptRoot\Report-MissingBackup.ps1"
. "$PSScriptRoot\Report-MissingLogBackup.ps1"
. "$PSScriptRoot\Report-FailedSQLAgentJob.ps1"
. "$PSScriptRoot\Report-ClusterNotOnDefaultNode.ps1"