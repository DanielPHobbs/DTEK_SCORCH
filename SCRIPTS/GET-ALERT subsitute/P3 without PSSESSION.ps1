#========================================================================
# Created with: SAPIEN Technologies, Inc., PowerShell Studio 2012 v3.1.35
# Created on:   4/27/2016 8:16 AM
# Created by:   xdanielhobbs
# Organization: DTKEITINNOVATIONS
# Filename:     
#========================================================================

#call powershell V3
$theresults = $inobj | PowerShell {

Import-Module OperationsManager
New-SCOMManagementGroupConnection -ComputerName Pdc2scm048
 
$ctac1='LUL-CTACPROD01'
$ctac2='LUL-CTACDR02'

$ctacProdStatus="CTAC server LUL-CTACPROD01 agent is GREEN"
$ctacDRStatus="CTAC server LUL-CTACDR02 agent is GREEN"
 
$WatcherClass = Get-SCOMClass -name “Microsoft.SystemCenter.HealthServiceWatcher”
$watcherObjectInstances = Get-SCOMClassInstance -class $WatcherClass 

$watcherobject1 = $watcherObjectInstances | where {$_.Displayname -eq $ctac1}
$watcherobject2 = $watcherObjectInstances | where {$_.Displayname -eq $ctac2}

$HBalert1 = $watcherobject1|get-scomalert| Where {$_.Name -eq ‘Health Service Heartbeat Failure’ -and $_.ResolutionState -ne 255}
$HBalert2 = $watcherobject2|get-scomalert| Where {$_.Name -eq ‘Health Service Heartbeat Failure’ -and $_.ResolutionState -ne 255}

If ($HBALERT1){ 
$ctacProdStatus = "CTAC server LUL-CTACPROD01 agent is GREY"
}

If ($HBALERT2){ 
$ctacDRStatus = "CTAC server LUL-CTACDR02 agent is GREY"
}

new-object pscustomobject -property @{        
        ctacProdStatus = $ctacProdStatus
        ctacDRStatus = $ctacDRStatus      
}
}

$ctacProdStatus=$theresults.ctacProdStatus
$ctacDRStatus=$theresults.ctacDRStatus
