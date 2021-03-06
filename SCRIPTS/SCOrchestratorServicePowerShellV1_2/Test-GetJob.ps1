
###################################################################
#    Copyright (c) Microsoft. All rights reserved.
#    This code is licensed under the Microsoft Public License.
#    THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF
#    ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY
#    IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR
#    PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.
###################################################################

##################################################
# Test-GetJob
##################################################

begin {
    # import modules
     Import-Module 'E:\SCORCH_SCRIPTS\SCOrchestratorServicePowerShellV1_2\OrchestratorServiceModule.psm1'
}

process {
    # get credentials (set to $null to UseDefaultCredentials)
    $creds = $null
    #$creds = Get-Credential "DOMAIN\USERNAME"

    # create the base url to the service
    $url = Get-OrchestratorServiceUrl -server "s2012-app1"

    # Define runbook, job, server
    #$rbid = [guid]"GUID"
    #$jobid = [guid]"GUID"
    #$serverid = [guid]"GUID"

    # Various tests

    # All jobs in system
    #$jobsarray = Get-OrchestratorJob -serviceurl $url

    # All jobs with credentials passed
    #$jobsarray = Get-OrchestratorJob -serviceurl $url -credentials $creds

    # Page 1 of jobs in system
    #$jobsarray = Get-OrchestratorJob -serviceurl $url -page 1

    # All jobs with particular status
    $jobsarray = Get-OrchestratorJob -serviceurl $url -status "Pending,Running"

    # Particular job
    #$jobsarray = Get-OrchestratorJob -serviceurl $url -jobid $jobid

    # All jobs on particular server (with particular service)
    #$jobsarray = Get-OrchestratorJob -serviceurl $url -serverid $serverid
    #$jobsarray = Get-OrchestratorJob -serviceurl $url -serverid $serverid -status "Running,Pending"

    # All jobs for particular runbook (and status)
    #$jobsarray = Get-OrchestratorJob -serviceurl $url -runbookid $rbid
    #$jobsarray = Get-OrchestratorJob -serviceurl $url -runbookid $rbid -status "Running,Pending"
    
    if ($jobsarray -ne $null)
    {
        $i = 1
        foreach ($job in $jobsarray)
        {
            Write-Host $i
            Write-Host "Url_Service = " $job.Url_Service
            Write-Host "Url = " $job.Url
            Write-Host 'Published' = $job.Published
            Write-Host 'Updated' = $job.Updated
            Write-Host "Category = " $job.Category
            Write-Host "Id = " $job.Id
            Write-Host 'RunbookId' = $job.RunbookId
            Write-Host 'RunbookServers' = $job.RunbookServers
            Write-Host 'RunbookServerId' = $job.RunbookServerId
            Write-Host 'Status' = $job.Status
            Write-Host 'ParentId' = $job.ParentId
            Write-Host 'ParentIsWaiting' = $job.ParentIsWaiting
            Write-Host 'CreatedBy' = $job.CreatedBy
            Write-Host 'CreationTime' = $job.CreationTime
            Write-Host 'LastModifiedBy' = $job.LastModifiedBy
            Write-Host 'LastModifiedTime' = $job.LastModifiedTime

            Write-Host 'Url_Runbook' = $job.Url_Runbook
            Write-Host 'Url_RunbookInstances' = $job.Url_RunbookInstances
            Write-Host 'Url_RunbookServer' = $job.Url_RunbookServer
            
            Write-Host 'Parameters'
            foreach ($param in $job.Parameters)
            {
                Write-Host 'Param.Name' = $param.Name
                Write-Host 'Param.Id' = $param.Id
                Write-Host 'Param.Value' = $param.Value
            }

            $i++
        }
    }
    else
    {
        Write-Host "No Jobs"
    }
}

end {
    # remove modules
    Remove-Module OrchestratorServiceModule
}
