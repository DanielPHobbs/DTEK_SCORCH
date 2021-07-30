<#
.SYNOPSIS
SMART - SMA Runbook Toolkit (Import-SMARunbookFromXMLorPS1.ps1)
Written by Jim Britt
Windows Server System Center, Customer and Technologies Team (WSSC CAT)
Microsoft Corporation - 10-16-2013
=======================================================================
    Updated by Jim Britt on 3-5-2014:
     
    Added support for importing secrets (if supplied) for variables and
    credentials.

    Fixed datatype of int to int32 in a few spots

    Added support for updating tags via webservice during edit process.
=======================================================================
=======================================================================
.DESCRIPTION
This scripted solution can be leveraged for importing
SMA Runbooks into an SMA environment from an XML based export file.
This solution leverages the existing SMA cmdlets for manipulating Runbooks
and assets and provides an atomic mechanism for sharing Runbooks between
environments.

This import solution supports the following
Runbook: Importing PS1 or XML Runbook Definition into SMA
Schedules: Any schedules assigned will be created if not already existing.
Variables: Any variables defined in the XML will be created if not already existing.
Credentials: Any credentials will be created with a default contoso\user and password.
Tags: Any tags specified on the Runbook configuration defined in the XML
Description: Description of the Runbook
Log Settings: All log settings defined for this Runbook will be imported

.EXAMPLE 
.\Import-SMARunbookfromXMLorPS1.ps1 -ImportDirectory "c:\temp\exports" -FileName "RunbookName.XML" -ImportAssets $True

Imports a Runbook into SMA named "RunbookName" from the the import directory using an XML file defined Runbook (with assets).
Assets in this case are defined as Schedules, Credentials, Variables and exist in the XML defined Runbook (not PS1). 

.EXAMPLE 

.\Import-SMARunbookfromXMLorPS1.ps1 -ImportDirectory "c:\temp\exports" -FileName "RunbookName.XML" -ImportVariables $True

Imports a Runbook into SMA named "RunbookName" from the the import directory using an XML file defined Runbook (with variables).
Encrypted variables are imported without data (just description and name)


.\Import-SMARunbookfromXMLorPS1.ps1 -ImportDirectory "c:\temp\exports" -FileName "RunbookName.XML" -ImportCredentials $True -ImportSchedules $True

Imports a Runbook into SMA named "RunbookName" from the the import directory using an XML file defined Runbook (with Creds and Schedules).
Schedules will be created if they do not exist.

.EXAMPLE 
.\Import-SMARunbookfromXMLorPS1.ps1 -ImportDirectory "c:\temp\exports" -FileName "RunbookName.XML" -ImportAssets $True -EnableScriptOutput $True

Imports a Runbook into SMA named "RunbookName" from the the import directory using an XML file defined Runbook (with all assets).
Screen status is also outputed leveraging the -EnableScriptOutput for basic details on progress.


.LINK
http://aka.ms/BuildingClouds
#>

# Parameters for main script
Param
(
    # Where your Runbooks are located for importing
    [parameter(Mandatory=$True)]
    [String]$ImportDirectory,

    # XMLFile that contains your Runbook to Import
    [parameter(Mandatory=$True)]
    [String]$FileName,

    # ex: "https://smaserver.contoso.com"
    [parameter(Mandatory=$False)]
    [string]$WebServiceEndpoint="https://localhost",

    # Windows or Basic
    [parameter(Mandatory=$False)]
    [string]$AuthenticationType="Windows",
        
    # Port 9090 defaults
    [parameter(Mandatory=$False)]
    [int]$Port=9090,
        
    # Leveraged for authenticating with alt creds
    [parameter(Mandatory=$False)]
    [pscredential]$cred,
        
    # Will overwrite existing Runbook in Draft
    [boolean]$overwrite,
        
    # Will take drafted Runbook and move to published
    [boolean]$Publish,
        
    # RunbookState - Draft or Published
    [parameter(Mandatory=$True)]
    [string]$RunbookState,

    # Will provide output to console for status
    [boolean]$EnableScriptOutput,

    # Import variables switch
    [boolean]$ImportVars,

    # Import credentials switch
    [boolean]$ImportCreds,
        
    # Import schedules switch
    [boolean]$ImportSchedules,

    # Import variables, credentials, schedules switch
    [boolean]$ImportAssets
)
Workflow Import-SMARunbookfromXMLorPS1
{
    Param
    (
        # Where your Runbooks are located for importing
        [parameter(Mandatory=$True)]
        [String]$ImportDirectory,

        # XMLFile that contains your Runbook to Import
        [parameter(Mandatory=$True)]
        [String]$FileName,

        # ex: "https://smaserver.contoso.com"
        [parameter(Mandatory=$False)]
        [string]$WebServiceEndpoint="https://localhost",

        # Windows or Basic
        [parameter(Mandatory=$False)]
        [string]$AuthenticationType="Windows",
        
        # Port 9090 defaults
        [parameter(Mandatory=$False)]
        [int]$Port=9090,
        
        # Leveraged for authenticating with alt creds
        [parameter(Mandatory=$False)]
        [pscredential]$cred,
        
        # Will overwrite existing Runbook in Draft
        [boolean]$overwrite,
        
        # Will take drafted Runbook and move to published
        [boolean]$Publish,
        
        # RunbookState - Draft or Published
        [parameter(Mandatory=$True)]
        [string]$RunbookState,

        # Will provide output to console for status
        [boolean]$EnableScriptOutput,

        # Import variables boolean
        [boolean]$ImportVars,

        # Import credentials boolean
        [boolean]$ImportCreds,
        
        # Import schedules boolean
        [boolean]$ImportSchedules,

        # Import variables, credentials, schedules boolean
        [boolean]$ImportAssets
    )
        # Display Params Passed
    ""
    $ScriptOutput = (get-date -format g).ToString() + "
ImportDirectory:$ImportDirectory`r
FileName:$FileName`r
WebServiceEndPoint:$WebServiceEndpoint`r
AuthenticationType:$AuthenticationType`r
Port:$Port`r
Cred:$cred`r
Overwrite?:$overwrite`r
Publish?:$Publish`r
RunbookState:$RunbookState`r
EnableScriptOutput?:$EnableScriptOutput`r
ImportVars?:$ImportVars`r
ImportCreds?:$ImportCreds`r
ImportSchedules?:$ImportSchedules`r
ImportAssets?:$ImportAssets`n`n"
if($EnableScriptOutput){ $ScriptOutput }
    
    # Comment out the below to show errors or change to "Continue"
    $ErrorActionPreference = "SilentlyContinue"

    #Validate the file exists
    $FileToProcess = "$ImportDirectory\$FileName"
    
    $InputFileResult = Test-Path $FileToProcess
    If ($InputFileResult -eq $false)
    {
        "$FiletoProcess doesn't exist"
        exit
    }
    
    # Get file extension of file     
    $FileExtention = $FileName.Split(".")[1]

    # It must be either a PS1 or XML to be processed
    IF(($FileExtention -ne "XML") -and ($FileExtention -ne "PS1"))
    {
        $ScriptOutput = "$FiletoProcess doesn't have a proper extension (XML or PS1) - skipping"
        if($EnableScriptOutput){ $ScriptOutput }
        exit
    }

    # Validate at least Workflow is present within XML definition or PS1
    $ValidWorkflow = Select-String -Path $FileToProcess -pattern "Workflow"
    IF($ValidWorkflow -eq $null)
    {
        $ScriptOutput = "$FiletoProcess isn't a valid Runbook - missing workflow designation"
        if($EnableScriptOutput){ $ScriptOutput }
        exit
    }
    
    # Determine PS1 or XML
    If($FileExtention -eq "XML")
    {
        # Get Runbook XML Data
        [XML]$RunbookXML = Get-Content "$FileToProcess"

        # Get Runbook Name
        $RunbookName = $RunbookXML.Runbook.Name

        # Get Runbook Content by Runbook State
        if($RunbookState -eq "Published")
        {
            If($RunbookXML.Runbook.Published.Definition -and $RunbookXML.Runbook.Published.Definition -ne "No Published Version")
            {
                $RunbookDefinition = $RunbookXML.Runbook.Published.Definition
            }
            else
            {
                "$FileToProcess - no published content found! Check Runbook XML for values."
                exit
            }
        }
        Else
        {
           if($RunbookXML.Runbook.Draft.Definition -ne "Draft Not Unique")
           {
                $RunbookDefinition = $RunbookXML.Runbook.Draft.Definition
           }
           else
           {
                "$FileToProcess - no draft content found! Check Runbook XML for values."
                exit
           }
        }
        
        # Define Runbook to Import (required by cmdlet - we are creating PS1 if XML provided)       
        $RunbookPS1 = "$ImportDirectory\$RunbookName.ps1"

        # Output Temp Runbook from Definition
        $RunbookDefinition | Out-File $RunbookPS1

        # Initialize Common Variables for Runbook Operations 
        
        # Set Tags for Runbook (only available on initial import - can't update tags with current release of cmdlets)
        # Update: We'll do this via webservice in this version of the script. :)
        $RunbookTag = $RunbookXML.Runbook.Tag

        # Define the description for the Runbook
        $RunbookDescription = $RunbookXML.Runbook.Configuration.Description

        # Set Runbook Log Options
        $RunbookLogDebug = [System.Convert]::ToBoolean($RunbookXML.Runbook.Configuration.LogDebug)
        $RunbookLogProgress = [System.Convert]::ToBoolean($RunbookXML.Runbook.Configuration.LogProgress)
        $RunbookLogVerbose = [System.Convert]::ToBoolean($RunbookXML.Runbook.Configuration.LogVerbose)
    }
    
    # Determine PS1 or XML
    If($FileExtention -eq "PS1")
    {    
        # setting base name for Runbook in SMA
        $RunbookName = $Filename.Split(".")[0] 
        
        # PS1 doesn't support Tag and Description on import
        $RunbookTag = ""   
        $RunbookDescription = ""    
    }
    
    # Name of Runbook PS1 File generated for Import
    $RunbookPS1 = "$ImportDirectory\$RunbookName.ps1"

    # Determine if Runbook already exists in environment and if so, get ID
    $GetRunbook = Get-SmaRunbook -Name $RunbookName -WebServiceEndpoint $WebServiceEndpoint -Port $Port -AuthenticationType $AuthenticationType -Credential $Cred
    If(!$GetRunbook.RunbookID.Guid)
    {
        $ImportedRunbook = Import-SmaRunbook -path $RunbookPS1 `
            -Tags $RunbookTag -WebServiceEndpoint $WebServiceEndpoint -Port $Port -AuthenticationType $AuthenticationType -Credential $cred
            $id = $ImportedRunbook.RunbookID.Guid
            $ScriptOutput = "Runbook $RunbookName was successfully imported"
            if($EnableScriptOutput){ $ScriptOutput }
    }
    else
    {
        $id = $GetRunbook.RunbookID.Guid        
        $EditStatus = Edit-SmaRunbook -Overwrite -path $RunbookPS1 `
            -Name $RunbookName -WebServiceEndpoint $WebServiceEndpoint -Port $Port -AuthenticationType $AuthenticationType -Credential $cred
        $ScriptOutput = "Runbook $RunbookName Exists - Updated"
        if($EnableScriptOutput){ $ScriptOutput }

        # Now let's update the tags via the webservice
        # Set the Runbook we are working with by ID
        $RunbookURI = "https://localhost:9090/00000000-0000-0000-0000-000000000000/Runbooks(guid'${ID}')"

# Define our XML Schema to update via the webservice
[xml]$baseXML = @'
<?xml version="1.0" encoding="utf-8"?>
<entry xmlns="http://www.w3.org/2005/Atom" xmlns:d="http://schemas.microsoft.com/ado/2007/08/dataservices" xmlns:m="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata">
    <id></id>
    <category term="Orchestrator.ResourceModel.Runbook" scheme="http://schemas.microsoft.com/ado/2007/08/dataservices/scheme" />
    <title />
    <updated></updated>
    <author>
        <name />
    </author>
    <content type="application/xml">
        <m:properties>
            <d:Tags></d:Tags>
        </m:properties>
    </content>
</entry>
'@
        # Using an inlinescript to update the webservice with tags
        $TagUpdate = InlineScript
        {
            # Set our tag data on the property in the XML schema
            $XMLValue = $Using:baseXML
            $XMLValue.entry.content.properties.Tags = [string]$Using:RunbookTag
            Invoke-RestMethod -Method Merge -Uri $Using:RunbookURI -Body $XMLValue -UseDefaultCredentials -ContentType 'application/atom+xml'
            $ScriptOutput = "Tag information updated via the webservice for $Using:RunbookName"
            if($EnableScriptOutput){ $ScriptOutput }
        }
    }

    # If XML, Update log and description data if specified in XML
    If($FileExtention -eq "XML")
    {
        # Set Runbook Configuration - Description / Log Options
        Set-SmaRunbookConfiguration -id $id -Description $RunbookDescription `
            -LogDebug $RunbookLogDebug -LogVerbose $RunbookLogVerbose -LogProgress $RunbookLogProgress `
            -WebServiceEndpoint $WebServiceEndpoint -Port $Port -AuthenticationType $AuthenticationType -Credential $cred
        $ScriptOutput = "Description and Logging information updated for $RunbookName"
        if($EnableScriptOutput){ $ScriptOutput }

        # Remove temporary PS1 if leveraging XML
        Remove-Item $RunbookPS1
        $ScriptOutput = "Temporary file $ImportDirectory\$RunbookName.ps1 was removed"
        if($EnableScriptOutput){ $ScriptOutput }

        # VARIABLES SECTION
        # ImportVars specified (or ImportAssets used)
        if($ImportVars -or $ImportAssets)
        {
            # Use $RunbookState to determine between published or draft
            # Process each variable associated in XML
            $VarsToProcess = $RunbookXML.Runbook.$RunbookState.Variables
            foreach($Var in $VarsToProcess)
            {
                $varExists = Get-SmaVariable -Name $Var.VariableName -WebServiceEndpoint $WebServiceEndpoint -Port $Port -AuthenticationType $AuthenticationType -Credential $cred
                if(!$varExists) # Only create if it doesn't exist
                {
                    $VariableName = $Var.VariableName
                    $VariableValue = $Var.VariableValue
                    $VariableDescription = $var.VariableDescription
                    $VariableType = $Var.VariableType
                    $IsEncrypted = $var.IsEncrypted

                    # Detect INT or String and set appropriately
                    if($VariableType -eq "string"){ $variablevalue = [string]$VariableValue }
                    if($VariableType -eq "Int32" -or $VariableType -eq "int"){ $variablevalue = [Int32]$VariableValue }
                    
                    # Detect if this is a $NULL variable (cannot create currently)
                    if($VariableValue -eq "" -and $IsEncrypted -eq "False")
                    { 
                        $ScriptOutput = "No value in $variableName. Must be a NULL variable. Not created."
                        if($EnableScriptOutput){ $ScriptOutput }
                        exit 
                    }

                    # Detect BOOLEAN and DATE and set appropriately
                    if($VariableType -eq "bool"){ $variablevalue = [bool]$VariableValue }
                    if($VariableType -eq "DateTime"){ $VariableValue = [DateTime]$VariableValue }
                    
                    # If this is an encrypted variable, handle appropriately
                    if($IsEncrypted -eq "True")
                    {
                        $CreateEncryptedVariable = Set-SmaVariable -Name $VariableName -Value $VariableValue -Description $VariableDescription -WebServiceEndpoint $WebServiceEndpoint -Port $Port -AuthenticationType $AuthenticationType -Credential $cred -Encrypted
                        $ScriptOutput = "Creating encrypted variable $variableName of type $VariableType.  You may need to assign value after import! If supplied in XML and imported from SMA - value is entered."
                        if($EnableScriptOutput){ $ScriptOutput }
                    }
                    Else #Otherwise Create normally (non encrypted)
                    {
                        $CreateNonEncryptedVariable = Set-SmaVariable -Name $VariableName -Value $VariableValue -Description $VariableDescription -WebServiceEndpoint $WebServiceEndpoint -Port $Port -AuthenticationType $AuthenticationType -Credential $cred
                        $ScriptOutput = "Creating standard variable $variableName of type $VariableType"
                        if($EnableScriptOutput){ $ScriptOutput }
                    }
                }     

            }
        }
        
        # SCHEDULES IMPORT SECTION
        # Import schedules if boolean used (or ImportAssets used)
        if($ImportSchedules -or $ImportAssets)
        {
            $SchedulesToProcess = $RunbookXML.Runbook.Schedules
            foreach($Schedule in $SchedulesToProcess)
            {
                $ScheduleExists = Get-SmaSchedule -Name $Schedule.ScheduleName -WebServiceEndpoint $WebServiceEndpoint -Port $Port -AuthenticationType $AuthenticationType -Credential $cred
                if(!$ScheduleExists)
                {
                    # Assign variables for Schedule with XML Data
                    [string]$ScheduleName = $Schedule.ScheduleName
                    [string]$ScheduleDescription = $Schedule.ScheduleDescription
                    [string]$ScheduleType = $Schedule.ScheduleType
                    [datetime]$ScheduleNextRun = $Schedule.ScheduleNextRun
                    [datetime]$ScheduleExpiryTime = $Schedule.ScheduleExpiryTime
                    [int32]$ScheduleDayInterval = $Schedule.ScheduleDayInterval

                    # Create schedule with details from XML - start time in the past will be set to execute on the next scheduled instance
                    # (assigning to variable to avoid displaying to screen)
                    $SMASchedule = Set-SmaSchedule -Name $ScheduleName -Description $ScheduleDescription -ScheduleType DailySchedule -StartTime $ScheduleNextRun -ExpiryTime $ScheduleExpiryTime -DayInterval $ScheduleDayInterval -WebServiceEndpoint $WebServiceEndpoint -Port $Port -AuthenticationType $AuthenticationType -Credential $cred
                    $ScriptOutput = "Schedule (" + $ScheduleName + ") doesn't exist...creating"
                    if($EnableScriptOutput){ $ScriptOutput }

                    # Associating Schedule to Runbook (assigning to variable to avoid displaying to screen)
                    $StartSMARunbook = Start-SmaRunbook -Name $RunbookName -ScheduleName $ScheduleName -WebServiceEndpoint $WebServiceEndpoint -Port $Port -AuthenticationType $AuthenticationType -Credential $cred
                    $ScriptOutput = "Associating (" + $ScheduleName + ") with $RunbookName."
                    if($EnableScriptOutput){ $ScriptOutput }
                }
                Else
                {
                    # Associating Schedule to Runbook  (assigning to variable to avoid displaying to screen)
                    $StartSMARunbook = Start-SmaRunbook -Name $RunbookName -ScheduleName $Schedule.ScheduleName -WebServiceEndpoint $WebServiceEndpoint -Port $Port -AuthenticationType $AuthenticationType -Credential $cred 
                    $ScriptOutput = "Schedule (" + $Schedule.ScheduleName + ") does exist associating with $RunbookName."
                    if($EnableScriptOutput){ $ScriptOutput }
                }
                
            }
        }
        # CREDENTIALS IMPORT SECTION
        # Import credentials if boolean used (or ImportAssets used)
        if($ImportCreds -or $ImportAssets)
        {
            # Use $RunbookState to determine between published or draft
            $CredsToProcess = $RunbookXML.Runbook.$RunbookState.Credentials
            
            foreach($Credential in $CredsToProcess)
            {
                $CredExists = Get-SmaCredential -Name $credential.CredentialName -WebServiceEndpoint $WebServiceEndpoint -Port $Port -AuthenticationType $AuthenticationType -Credential $cred
                if(!$CredExists)
                {
                    # Create credential with XML Data
                    [string]$CredentialName = $Credential.CredentialName
                    [string]$CredentialDescription = $Credential.CredentialDescription
                    
                    # If there are values, go ahead and set appropriately
                    If($Credential.CredentialUserName)
                    {
                        [string]$CredentialUserName = $Credential.CredentialUserName
                        If($Credential.CredentialPassword){[string]$CredentialPassword = $Credential.CredentialPassword}
                        else{$ScriptOutput = "Credential " + $CredentialName + " no values specified in XML for Credential Password..not created"
                        if($EnableScriptOutput){ $ScriptOutput }}
                    }

                    
                    # If the value is null or doesn't exist, create with temp creds
                    If($CredentialUserName -eq $null -or $CredentialUserName -eq "")
                    {
                        # Temporary User and Password - update after import
                        $tmpcredPassWord = ConvertTo-SecureString "Password" -AsPlainText -Force
                        $tmpcred = New-Object System.Management.Automation.PSCredential ("contoso\User", $tmpcredPassWord)
                        $ScriptOutput = "Credential " + $CredentialName + " no values specified in XML..please update after import.."
                        if($EnableScriptOutput){ $ScriptOutput }

                    }
                    elseif($CredentialUserName) #otherwise let's create it with the proper creds
                    {
                        # Credentials created with User and Password
                        $tmpcredPassWord = ConvertTo-SecureString $CredentialPassword -AsPlainText -Force
                        $tmpcred = New-Object System.Management.Automation.PSCredential ($CredentialUserName, $tmpcredPassWord)
                        $ScriptOutput = "Credential " + $CredentialName + " creating with values specified in XML.."
                        if($EnableScriptOutput){ $ScriptOutput }
                    }

                    # Create creds
                    $SMACred = Set-SmaCredential -Name $CredentialName -Value $tmpcred -Description $CredentialDescription -WebServiceEndpoint $WebServiceEndpoint -Port $Port -AuthenticationType $AuthenticationType -Credential $cred
                    $ScriptOutput = "Credential " + $CredentialName + " doesn't exist...creating"
                    if($EnableScriptOutput){ $ScriptOutput }
                }
            }
        }
    }
    If ($Publish)
    {
        # Publish SMA Runbook to SMA [leverage boolean]
        $PublishStatus = Publish-SmaRunbook -Id $id -WebServiceEndpoint $WebServiceEndpoint -Port $Port -AuthenticationType $AuthenticationType -Credential $cred
        $ScriptOutput = "Publish boolean used - $RunbookName was successfully published."
        if($EnableScriptOutput){ $ScriptOutput }
    }
    else
    {
        $ScriptOutput = "Publish boolean not used - $RunbookName was placed in draft mode"
        if($EnableScriptOutput){ $ScriptOutput }
    }
    Return $RunbookName
}

Import-SMARunbookfromXMLorPS1 -ImportDirectory $ImportDirectory -FileName $FileName -WebServiceEndpoint $WebServiceEndpoint `
-AuthenticationType $AuthenticationType -Port $Port -cred $cred -overwrite $overwrite -Publish $Publish -RunbookState $RunbookState `
-EnableScriptOutput $EnableScriptOutput -ImportVars $ImportVars -ImportCreds $ImportCreds -ImportSchedules $ImportSchedules -ImportAssets $ImportAssets