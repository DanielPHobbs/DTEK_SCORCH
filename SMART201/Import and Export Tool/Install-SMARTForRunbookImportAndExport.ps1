<#
.SYNOPSIS
SMART - SMA Runbook Toolkit (Install-SMARTForRunbookImportAndExport.ps1)
Written by Jim Britt
Windows Server System Center, Customer and Technologies Team (WSSC CAT)
Microsoft Corporation - 10-23-2013
Updated by Jim Britt - 03-05-2014

.DESCRIPTION
This is the install wrapper for importing the SMA Runbooks related to
SMART for Runbook Import and Export.

.LINK
http://aka.ms/BuildingClouds
#>
Workflow Install-SMARunbooks
{
    param
    (
        [string]$ScriptDirectory,
        [string]$WebServiceEndpoint,
        [string]$RunbookState,
        [boolean]$ImportAssets,
        [Boolean]$overwrite        
    )

    # IMPORT SECTION
    # Builk Import from Directory with XML Files
    $ImportDirectory = $ScriptDirectory + "\RunbookXMLs"
    $Files = Get-ChildItem $ImportDirectory -File

    # We setup a while loop to ensure we process as Import/Publish/Edit/Publish (runs twice) to validate parent / child dependencies on Runbooks
    $i = 0
    while($i -le 1)
    {
        foreach -parallel($File in $Files)
        {
            # Only Process XML and PS1 in working directory
            if($File.extension -iin ".XML",".PS1")
            {
                # Get Runbook Name
                $RunbookToPublish = $File.BaseName
                
                # Modify parameters below according to your needs
                $RunbookVar = InlineScript{
                    cd $Using:ScriptDirectory
                    $parms = @{'WebServiceEndpoint'=$Using:WebServiceEndpoint;
		               'ImportDirectory'=$Using:ImportDirectory;
                       'FileName'=$Using:File.Name;
                       'overwrite'=$Using:overWrite;
                       'RunbookState'=$Using:RunbookState;
                       'ImportAssets'=$Using:ImportAssets
                    }
                    .\Import-SMARunbookfromXMLorPS1.ps1 @parms
                } 
                # Publish the Runbook into SMA that we've just imported.               
                $PublishedRunbook = Publish-SMARunbook -Name $RunbookToPublish -WebServiceEndpoint $WebServiceEndpoint
            }
        }
        # Looping once to ensure Parent and Child have dependencies built properly in SMA
        $i++
    }
    Write-Output "Completed  importing Runbooks into SMA"
}
# Determine when we started executing
$Script:startTime = get-date

# Find out where we are running
if ($MyInvocation.MyCommand.Path -ne $null)
{
    $ScriptDirectory = Split-Path ($MyInvocation.MyCommand.Path)
}

# Execute workflow to install SMA Runbooks into SMA
# Update WebServiceEndpoint if not running from SMA Server
Install-SMARunbooks -ScriptDirectory $ScriptDirectory -WebServiceEndpoint "https://localhost" -RunbookState "Published" -ImportAssets $true -overwrite $true

# Calculate elapsed time and display
$elapsedTime = new-timespan $script:StartTime $(get-date)
"Total Elapsed Time: $elapsedTime"
pause