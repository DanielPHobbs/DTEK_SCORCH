Function Find-Runbook {
    <#
    .Synopsis
    Find runbook.
    .DESCRIPTION
    Function retrieves a rubook based on name. 
    .PARAMETER name
    Runbook name
    .EXAMPLE
    
    #>
[Cmdletbinding()]
Param (
    [Parameter(ValueFromPipelineByPropertyName=$True, ValueFromPipeline=$True, Mandatory=$True)]
    [string[]]$name,
    [string]$scorch = 'localhost',
    [System.Management.Automation.PSCredential]$credential
    
)
BEGIN {
$port = 81
$baseurl = "http://$scorch`:$port/Orchestrator2012/Orchestrator.svc"
$resource = '/Runbooks'
}
PROCESS {
    Foreach ($n in $name) {
        $filter = "?`$filter=Name eq '$n'"
        $url = $baseurl + $resource + $filter
            Try {
                If ($credential) {
                    $response = Invoke-RestMethod -Uri $url -Method Get -Credential $credential -ErrorAction Stop
                }
                Else {
                    $response = Invoke-RestMethod -Uri $url -Method Get -UseDefaultCredentials -ErrorAction Stop
                }
                If ($response) {
                    Foreach ($resp in $response) {
                        $xmlproperties = ([xml]($resp.content.OuterXml)).content.properties
                        If ($xmlproperties.CheckedOutTime.null) {
                            $checkedin = $true
                        }
                        Else {
                            $checkedin = $false
                        }
        
                        $props = @{
                            Title   = $resp.title.'#text'
                            Guid    = $xmlproperties.id.'#text'
                            id      = $resp.id
                            Path    = $xmlproperties.path.'#text'
                            LastModifiedTime = [datetime]$xmlproperties.LastModifiedTime.'#text'
                            FolderId = $xmlproperties.FolderId.'#text'
                            CheckedIn =  $checkedin
                        }
                        New-Object -TypeName PSObject -Property $props
                    }
                }
                else {
                    Write-Error "Object $n not found"
                }
            }
            Catch {
                Throw $error[0].Exception.Message
                }
            }# end foreach
        } #end process
} #end fdunction

Function Get-Runbook {
    <#
    .Synopsis
    Retrieve runbook.
    .DESCRIPTION
    Function retrieves a rubook based on guid. 
    .PARAMETER guid
    Runbook guid
    .EXAMPLE
    
    #>
[Cmdletbinding()]
Param (
    [Parameter(ValueFromPipelineByPropertyName=$True, ValueFromPipeline=$True, Mandatory=$True)]
    [guid[]]$guid,
    [string]$scorch = 'localhost',
    [System.Management.Automation.PSCredential]$credential
    
)
BEGIN {
$port = 81
$baseurl = "http://$scorch`:$port/Orchestrator2012/Orchestrator.svc"
$resource = '/Runbooks'
}
PROCESS {
    Foreach ($id in $guid) {
        $idstring = "(guid'$id')"
        $url = $baseurl + $resource + $idstring
            Try {
                If ($credential) {
                    $response = Invoke-RestMethod -Uri $url -Method Get -Credential $credential -ErrorAction Stop
                }
                Else {
                    $response = Invoke-RestMethod -Uri $url -Method Get -UseDefaultCredentials -ErrorAction Stop
                }
                If ($response) {
                    $response = $response.entry
                    $xmlproperties = ([xml]($response.content.OuterXml)).content.properties
                    If ($xmlproperties.CheckedOutTime.null) {
                        $checkedin = $true
                    }
                    Else {
                        $checkedin = $false
                    }
                    $props = @{
                        Title   = $response.title.'#text'
                        Guid    = $xmlproperties.id.'#text'
                        id      = $response.id
                        Path    = $xmlproperties.path.'#text'
                        LastModifiedTime = [datetime]$xmlproperties.LastModifiedTime.'#text'
                        FolderId = $xmlproperties.FolderId.'#text'
                        CheckedIn = $checkedin
                    }
                    New-Object -TypeName PSObject -Property $props
                }
                else {
                    Write-Error "Object $id not found"
                }
            }
            Catch {
                Throw $error[0].Exception.Message
                }
            }# end foreach
        } #end process
} #end fdunction

Function Get-RunbookPicture {
    <#
    .Synopsis
    Retrieve runbok bitmap.
    .DESCRIPTION
    Function retrieves runbook's flow chart transparent bitmap. 
    .PARAMETER runbook
    Runbook object
    .PARAMETER folder
    Folder for the output png file.
    .EXAMPLE
    $name | Get-Runbook | Get-RunbookPicture -folder .\
    #>
    [Cmdletbinding()]
    Param (
        [Parameter(ValueFromPipelineByPropertyName=$True, ValueFromPipeline=$True, Mandatory=$True)]
        [object[]]$runbook,
        [string]$scorch = 'localhost',
        [System.Management.Automation.PSCredential]$credential,
        [string]$folder = "$env:HOMEDRIVE$env:HOMEPATH\Pictures"

    )
PROCESS {
    $guid = $runbook.guid
    $title = $runbook.title
    $filename = $title.Replace(' ','_') + '.png'
    $port = 81
    $baseurl = "http://$scorch`:$port/Orchestrator2012/Orchestrator.svc"
    $resource = '/RunbookDiagrams'
    
    $url = $baseurl + $resource + "(guid'$guid')/`$value"
    If ($credential) {
        Invoke-WebRequest -Uri $url -Method Get -Credential $credential -OutFile "$folder\$filename" -PassThru
    }
    Else {
        Invoke-WebRequest -Uri $url -Method Get -UseDefaultCredentials -OutFile "$folder\$filename" -PassThru
        
    }
} #end process
}

Function Get-RunbookInstance {
<#
.Synopsis
Retrieves runbok instances.
.DESCRIPTION
Function retrieves runbook's instances. 
.PARAMETER id
Runbook id (full URL)
.EXAMPLE

#>
[Cmdletbinding()]
Param (
    [Parameter(ValueFromPipelineByPropertyName=$True, ValueFromPipeline=$True, Mandatory=$True)]
    [string[]]$id,
    [System.Management.Automation.PSCredential]$credential
)
PROCESS {
    Foreach ($i in $id) {
            Try {
                If ($credential) {
                    $response = Invoke-RestMethod -Uri "$i/Instances" -Method Get -Credential $credential -ErrorAction Stop
                }
                Else {
                    $response = Invoke-RestMethod -Uri "$i/Instances" -Method Get -UseDefaultCredentials -ErrorAction Stop
                }
                Foreach ($resp in $response) {
                    $xmlproperties = ([xml]($resp.content.OuterXml)).content.properties
                    $props = @{
                        Guid    = $xmlproperties.id.'#text'
                        id      = $resp.id
                        JobId    = $xmlproperties.JobId.'#text'
                        CompletionTime = [datetime]$xmlproperties.CompletionTime.'#text'
                        Status = $xmlproperties.Status.'#text'
                        ParentGuid = $xmlproperties.RunbookId.'#text'
                    }
                    New-Object -TypeName PSObject -Property $props
                } #end foreach

            } #end try
            Catch {
                Throw $error[0].Exception.Message
            }
    }# end foreach
} #end process
}

Function Find-ScorchFolder {
<#
.Synopsis
Find Scorch folder.
.DESCRIPTION
Function retrieves a rubooks folder based on name. 
.PARAMETER name
Runbook name
.EXAMPLE
  
#>
[Cmdletbinding()]
Param (
    [Parameter(ValueFromPipelineByPropertyName=$True, ValueFromPipeline=$True, Mandatory=$True)]
    [string[]]$name,
    [string]$scorch = 'localhost',
    [System.Management.Automation.PSCredential]$credential
)
BEGIN {
    $port = 81
    $baseurl = "http://$scorch`:$port/Orchestrator2012/Orchestrator.svc"
    $resource = '/Folders'
}
PROCESS {
    Foreach ($n in $name) {
        $filter = "?`$filter=Name eq '$n'"
        $url = $baseurl + $resource + $filter
        Try {
            If ($credential) {
                $response = Invoke-RestMethod -Uri $url -Method Get -Credential $credential -ErrorAction Stop
            }
            Else {
                $response = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop -UseDefaultCredentials
            }
            If ($response) {
                Foreach ($resp in $response) {
                    $xmlproperties = ([xml]($resp.content.OuterXml)).content.properties
                    $props = @{
                        Title   = $resp.title.'#text'
                        Guid    = $xmlproperties.id.'#text'
                        id      = $resp.id
                        Path    = $xmlproperties.path.'#text'
                        LastModifiedTime = [datetime]$xmlproperties.LastModifiedTime.'#text'
                        ParentId = $xmlproperties.ParentID.'#text'
                    }
                    New-Object -TypeName PSObject -Property $props
                }
            }
            else {
                Write-Error "Object $n not found"
            }
        }
        Catch {
            $_.Exception.Message
            }

    }
}
}

Function Get-ScorchFolder {
<#
.Synopsis
Get Scorch folder.
.DESCRIPTION
Function retrieves a rubooks folder based on guid. 
.PARAMETER guid
Runbook guid
.EXAMPLE
      
#>
[Cmdletbinding()]
Param (
    [Parameter(ValueFromPipelineByPropertyName=$True, ValueFromPipeline=$True, Mandatory=$True)]
    [guid[]]$guid,
    [string]$scorch = 'localhost',
    [System.Management.Automation.PSCredential]$credential
)
BEGIN {
    $port = 81
    $baseurl = "http://$scorch`:$port/Orchestrator2012/Orchestrator.svc"
    $resource = '/Folders'
}
PROCESS {
    Foreach ($id in $guid) {
        $idstring = "(guid'$id')"
        $url = $baseurl + $resource + $idstring
            Try {
            If ($credential) {
                $response = Invoke-RestMethod -Uri $url -Method Get -Credential $credential -ErrorAction Stop
            }
            Else {
                $response = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop -UseDefaultCredentials
            }
            If ($response) {    
                $response = $response.entry            
                $xmlproperties = ([xml]($response.content.OuterXml)).content.properties
                $props = @{
                    Title   = $response.title.'#text'
                    Guid    = $xmlproperties.id.'#text'
                    id      = $response.id
                    Path    = $xmlproperties.path.'#text'
                    LastModifiedTime = [datetime]$xmlproperties.LastModifiedTime.'#text'
                }
                New-Object -TypeName PSObject -Property $props
            }
            else {
                Write-Error "Object $id not found"
            }
        }
        Catch {
            Throw $error[0].Exception.Message
            }
    
    }
}
}

Function Get-ScorchSubFolder {
# Get subfolders based on the parent folder's Guid
[Cmdletbinding()]
Param (
    [Parameter(ValueFromPipelineByPropertyName = $True, ValueFromPipeline = $True, Mandatory = $True)]
    [guid[]]$guid,
    [string]$scorch = 'localhost',
    [System.Management.Automation.PSCredential]$credential
)
BEGIN {
    $port = 81
    $baseurl = "http://$scorch`:$port/Orchestrator2012/Orchestrator.svc"
    $resource = '/Folders'
}
PROCESS {

    Foreach ($id in $guid) {
        $url = $baseurl + $resource + "(guid'$id')/Subfolders"
        Try {
            If ($credential) {
                $response = Invoke-RestMethod -Uri $url -Method Get -Credential $credential
            }
            Else {
                $response = Invoke-RestMethod -Uri $url -Method Get -UseDefaultCredentials

            }
            Foreach ($resp in $response) {
                $xmlproperties = ([xml]($resp.content.OuterXml)).content.properties
                $props = @{
                    Title            = $resp.title.'#text'
                    Guid             = $xmlproperties.id.'#text'
                    id               = $resp.id
                    Path             = $xmlproperties.path.'#text'
                    LastModifiedTime = [datetime]$xmlproperties.LastModifiedTime.'#text'
                }
                New-Object -TypeName PSObject -Property $props
            } # end foreach
        } # end try
        Catch {
            Throw $error[0].Exception.Message
        }
    } # end foreach id
} # end process
} # end function

Function Get-FolderRunbooks {
#
# Get all runbooks in a scorch folder

[Cmdletbinding()]
Param (
    [Parameter(ValueFromPipelineByPropertyName = $True, ValueFromPipeline = $True, Mandatory = $True)]
    [guid[]]$guid,
    [string]$scorch = 'localhost',
    [System.Management.Automation.PSCredential]$credential
)
BEGIN {
    Import-Module ActiveDirectory -ErrorAction SilentlyContinue
    $port = 81
    $baseurl = "http://$scorch`:$port/Orchestrator2012/Orchestrator.svc"
    $resource = '/Folders'
}
PROCESS {
    Foreach ($id in $guid) {
        $url = $baseurl + $resource + "(guid'$id')/Runbooks"
        Try {
            If ($credential) {
                $response = Invoke-RestMethod -Uri $url -Method Get -Credential $credential
            }
            Else {
                $response = Invoke-RestMethod -Uri $url -Method Get -UseDefaultCredentials
    
            }
            Foreach ($resp in $response) {
                $xmlproperties = ([xml]($resp.content.OuterXml)).content.properties
                If ($xmlproperties.CheckedOutTime.null) {
                    $checkedin = $true
                }
                Else {
                    $checkedin = $false
                }
                $props = @{
                    Title   = $resp.title.'#text'
                    Guid    = $xmlproperties.id.'#text'
                    id      = $resp.id
                    Path    = $xmlproperties.path.'#text'
                    LastModifiedTime = [datetime]$xmlproperties.LastModifiedTime.'#text'
                    LastModifiedBy = ($xmlproperties.LastModifiedBy.'#text' | Get-ADuser -ea SilentlyContinue).samaccountname
                    FolderId = $xmlproperties.FolderId.'#text'
                    CheckedIn =  $checkedin
                }
                New-Object -TypeName PSObject -Property $props
            }
        }
        Catch {
            Throw $error[0].Exception.Message
        }
    }
}
} # end function

Function Get-ScorchJob {
# Get jobs based on jobID
[Cmdletbinding()]
Param (
    [Parameter(ValueFromPipelineByPropertyName = $True, ValueFromPipeline = $True, Mandatory = $True)]
    [guid[]]$JobID,
    [string]$scorch = 'localhost',
    [System.Management.Automation.PSCredential]$credential
)
BEGIN {
    $port = 81
    $baseurl = "http://$scorch`:$port/Orchestrator2012/Orchestrator.svc"
    $resource = '/Jobs'
}
PROCESS {
    Foreach ($id in $jobid) {
        $url = $baseurl + $resource + "(guid'$id')"
        Try {
            If ($credential) {
                $response = Invoke-RestMethod -Uri $url -Method Get -Credential $credential
            }
            Else {
                $response = Invoke-RestMethod -Uri $url -Method Get -UseDefaultCredentials
    
            }
                $resp = $response.entry
                $xmlproperties = ([xml]($resp.content.OuterXml)).content.properties

                $props = @{
                    Guid    = $xmlproperties.id.'#text'
                    id      = $resp.id
                    RunbookID = $xmlproperties.RunbookId.'#text'
                    ParentIsWaiting = $xmlproperties.ParentIsWaiting.'#text'
                    CreationTime = [datetime]$xmlproperties.CreationTime.'#text'
                    Status = $xmlproperties.Status.'#text'
                    Updated = [datetime]$resp.updated
                    
                }
                New-Object -TypeName PSObject -Property $props
            
        }
        Catch {
            Throw $error[0].Exception.Message
        }
    }
}
} # end function
    
Function Get-ScorchRunningJobs {
[Cmdletbinding()]
Param (
[string]$scorch = 'localhost',
[System.Management.Automation.PSCredential]$credential
)
$port = 81
$baseurl = "http://$scorch`:$port/Orchestrator2012/Orchestrator.svc"
$resource = '/Jobs'
$filter = "()?`$filter=Status eq 'Running'&`$select=Id,RunbookId,LastModifiedTime,Status"
$url = $baseurl + $resource + $filter
Try {
    If ($credential) {
        $response = Invoke-RestMethod -Uri $url -Method Get -Credential $credential -ErrorAction Stop
    }
    Else {
        $response = Invoke-RestMethod -Uri $url -Method Get -UseDefaultCredentials -ErrorAction Stop
    }
    Foreach ($resp in $response) {
        $xmlproperties = ([xml]($resp.content.OuterXml)).content.properties

        $props = @{
            Guid    = $xmlproperties.id.'#text'
            id      = $resp.id
            Published = [datetime]$resp.published
            Updated = [datetime]$resp.updated
            Status = $xmlproperties.Status.'#text'             
        }
        New-Object -TypeName PSObject -Property $props
    }

}
Catch {
    Throw $error[0].Exception.Message
}
}

Function Get-ScorchEvents {
[Cmdletbinding()]
Param (
    [string]$scorch = 'localhost',
    [int]$days = 0,
    [System.Management.Automation.PSCredential]$credential
    )
    $port = 81
    $baseurl = "http://$scorch`:$port/Orchestrator2012/Orchestrator.svc"
    $resource = '/Events'
    $day = (Get-Date).AddDays(-$days).ToString("yyy-MM-dd")
    $filter = "?`$filter=CreationTime gt datetime'$day'"
    $url = $baseurl + $resource + $filter
    Try {
        If ($credential) {
            $response = Invoke-RestMethod -Uri $url -Method Get -Credential $credential -ErrorAction Stop
        }
        Else {
            $response = Invoke-RestMethod -Uri $url -Method Get -UseDefaultCredentials -ErrorAction Stop
        }
        Foreach ($resp in $response) {
            $xmlproperties = ([xml]($resp.content.OuterXml)).content.properties

            $props = @{
                Guid    = $xmlproperties.id.'#text'
                id      = $resp.id
                SourceId = $xmlproperties.Summary.'#text'
                Updated = [datetime]$resp.updated
                CreationTime = [datetime]$xmlproperties.CreationTime.'#text'
                Summary = $xmlproperties.Summary.'#text'     
                Description = $xmlproperties.Description.'#text' 
            }
            New-Object -TypeName PSObject -Property $props
        }

    }
    Catch {
        Throw $error[0].Exception.Message
    }
}

Function Get-RunbookServers {
[Cmdletbinding()]
Param (
    [string]$scorch = 'localhost',
    [System.Management.Automation.PSCredential]$credential
    )
    $port = 81
    $baseurl = "http://$scorch`:$port/Orchestrator2012/Orchestrator.svc"
    $resource = '/RunbookServers'
    $url = $baseurl + $resource
    Try {
        If ($credential) {
            $response = Invoke-RestMethod -Uri $url -Method Get -Credential $credential -ErrorAction Stop
        }
        Else {
            $response = Invoke-RestMethod -Uri $url -Method Get -UseDefaultCredentials -ErrorAction Stop
        }
        Foreach ($resp in $response) {
            $xmlproperties = ([xml]($resp.content.OuterXml)).content.properties

            $props = @{
                Guid    = $xmlproperties.id.'#text'
                id      = $resp.id
                ComputerName = $resp.title.'#text'
                Updated = [datetime]$resp.updated
            }
            New-Object -TypeName PSObject -Property $props
        }

    }
    Catch {
        Throw $error[0].Exception.Message
    }
}

Function Start-Runbook {
Param (
    [Parameter(ValueFromPipelineByPropertyName = $True, ValueFromPipeline = $True, Mandatory = $True)]
    # Runbook guid
    [guid]$Guid,
    [string]$scorch = 'localhost',
    [System.Management.Automation.PSCredential]$credential
)
#$rbParameters = @{"00000000-0000-0000-00000000000000002" = "This is the value for Param1.";" 00000000-0000-0000-00000000000000003" = " This is the value for Param2."}
 
# Create the request object
$port = 81
$baseurl = "http://$scorch`:$port/Orchestrator2012/Orchestrator.svc"
$resource = '/Jobs'
$url = $baseurl + $resource

$request = [System.Net.HttpWebRequest]::Create("$url")

# Set the credentials to default or prompt for credentials
If (!$credential) {
    $request.UseDefaultCredentials = $true
}
Else {
    $request.Credentials = $cred
}

# Build the request header
$request.Method = "POST"
$request.UserAgent = "Microsoft ADO.NET Data Services"
$request.Accept = "application/atom+xml,application/xml"
$request.ContentType = "application/atom+xml"
$request.KeepAlive = $true
$request.Headers.Add("Accept-Encoding","identity")
$request.Headers.Add("Accept-Language","en-US")
$request.Headers.Add("DataServiceVersion","1.0;NetFx")
$request.Headers.Add("MaxDataServiceVersion","2.0;NetFx")
$request.Headers.Add("Pragma","no-cache")
 
# If runbook servers are specified, format the string
$rbServerString = ""
if (-not [string]::IsNullOrEmpty($RunbookServers)) {
   $rbServerString = -join ("<d:RunbookServers>",$RunbookServers,"</d:RunbookServers>")
}
 
# Format the Runbook parameters, if any
$rbParamString = ""
if ($rbParameters -ne $null) {
   
   # Format the param string from the Parameters hashtable
   $rbParamString = "<d:Parameters><![CDATA[<Data>"
   foreach ($p in $rbParameters.GetEnumerator())
   {
      #$rbParamString = -join ($rbParamString,"&lt;Parameter&gt;&lt;ID&gt;{",$p.key,"}&lt;/ID&gt;&lt;Value&gt;",$p.value,"&lt;/Value&gt;&lt;/Parameter&gt;")
      $rbParamString = -join ($rbParamString,"<Parameter><ID>{",$p.key,"}</ID><Value>",$p.value,"</Value></Parameter>")
   }
   $rbParamString += "</Data>]]></d:Parameters>"
}
 
# Build the request body
$requestBody = @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<entry xmlns:d="http://schemas.microsoft.com/ado/2007/08/dataservices" xmlns:m="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata" xmlns="http://www.w3.org/2005/Atom">
    <content type="application/xml">
        <m:properties>
            <d:RunbookId m:type="Edm.Guid">$guid</d:RunbookId>
            $rbserverstring
            $rbparamstring
        </m:properties>
    </content>
</entry>
"@
 
# Create a request stream from the request
$requestStream = new-object System.IO.StreamWriter $Request.GetRequestStream()
    
# Sends the request to the service
$requestStream.Write($RequestBody)
$requestStream.Flush()
$requestStream.Close()
 
# Get the response from the request
[System.Net.HttpWebResponse] $response = [System.Net.HttpWebResponse] $Request.GetResponse()
 
# Write the HttpWebResponse to String
$responseStream = $Response.GetResponseStream()
$readStream = new-object System.IO.StreamReader $responseStream
$responseString = $readStream.ReadToEnd()
 
# Close the streams
$readStream.Close()
$responseStream.Close()
 
# Get the ID of the resulting job
if ($response.StatusCode -eq 'Created')
{
    $xmlDoc = [xml]$responseString
    $jobId = $xmlDoc.entry.content.properties.Id.InnerText
    $contprops = $xmlDoc.entry.content.properties

    $props = @{
        Guid    = $contprops.Id.InnerText
        id      = $xmlDoc.entry.id
        RunbookID = $contprops.RunbookId.InnerText
        ParentIsWaiting = $contprops.ParentIsWaiting.InnerText
        CreationTime = [datetime]$contprops.CreationTime.InnerText
        Status = $contprops.Status
        Updated = [datetime]$xmlDoc.entry.updated
    }
    New-Object -TypeName PSObject -Property $props

}
else
{
    Throw "Could not start runbook. Status: $($response.StatusCode)"
}
}

Function Get-ScorchDump {
# Get Scorch runbooks recursively (except _schedules).
[CmdletBinding()]
Param (
    # Folder guid
    [Parameter(ValueFromPipelineByPropertyName = $True, ValueFromPipeline = $True)]
    [string[]]$guid = '00000000-0000-0000-0000-000000000000'
)
BEGIN{}
PROCESS {
    Foreach ($id in $guid) {
        Get-FolderRunbooks -guid $id
        $folders = Get-ScorchSubFolder -guid $id   
        #$folders | where Title -ne "_schedules" | Get-ScorchDump
        $folders | Get-ScorchDump
    }
}
END{}
}
