
###################################################################
#    Copyright (c) Microsoft. All rights reserved.
#    This code is licensed under the Microsoft Public License.
#    THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF
#    ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY
#    IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR
#    PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.
###################################################################

##################################################
# Test-GetRunbook
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

    # Use runbook id
    #$rbid = [guid]"GUID"
    #$rbarray = Get-OrchestratorRunbook -serviceurl $url -runbookid $rbid -credentials $creds

    # Use runbookpath
    $rbpath = "\tfl"
    $rbarray = Get-OrchestratorRunbook -serviceurl $url -runbookpath $rbpath -credentials $creds
        
    $i = 1
    foreach ($rb in $rbarray)
    {
        Write-Host $i
        Write-Host "Url_Service = " $rb.Url_Service
        Write-Host "Url = " $rb.Url
        Write-Host 'Url_Folder' = $rb.Url_Folder
        Write-Host 'Url_Parameters' = $rb.Url_Parameters
        Write-Host 'Url_Activities' = $rb.Url_Activities
        Write-Host 'Url_Jobs' = $rb.Url_Jobs
        Write-Host 'Url_Instances' = $rb.Url_Instances
        Write-Host 'Url_Diagram' = $rb.Url_Diagram
        Write-Host 'Name' = $rb.Name
        Write-Host 'Published' = $rb.Published
        Write-Host 'Updated' = $rb.Updated
        Write-Host "Category = " $rb.Category
        Write-Host "Id = " $rb.Id
        Write-Host 'FolderId' = $rb.FolderId
        Write-Host 'Description' = $rb.Description
        Write-Host 'CreatedBy' = $rb.CreatedBy
        Write-Host 'CreationTime' = $rb.CreationTime
        Write-Host 'LastModifiedBy' = $rb.LastModifiedBy
        Write-Host 'LastModifiedTime' = $rb.LastModifiedTime
        Write-Host 'IsMonitor' = $rb.IsMonitor
        Write-Host 'Path' = $rb.Path
        Write-Host 'CheckedOutBy' = $rb.CheckedOutBy
        Write-Host 'CheckedOutTime' = $rb.CheckedOutTime
        Write-Host ' '
        Write-Host 'Parameters'
        foreach ($param in $rb.Parameters)
        {
            Write-Host ' '
            Write-Host 'Param.Name' = $param.Name
            Write-Host 'Param.Id' = $param.Id
            Write-Host 'Param.Type' = $param.Type
            Write-Host 'Param.Description' = $param.Description
            Write-Host 'Param.RunbookUrl' = $param.RunbookUrl
            Write-Host 'Param.InstancesUrl' = $param.InstancesUrl
            Write-Host 'Param.Url' = $param.Url
            Write-Host 'Param.Updated' = $param.Updated
            Write-Host 'Param.Category' = $param.Category
            Write-Host 'Param.RunbookId' = $param.RunbookId
            Write-Host 'Param.Direction' = $param.Direction
        }

        Write-Host " "

        $i++
    }
}

end {
    # remove modules
    Remove-Module OrchestratorServiceModule
}
