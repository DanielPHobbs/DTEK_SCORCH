<#
.SYNOPSIS
SMART - SMA Runbook Toolkit (Export-SMARunbookToXML.ps1)
Written by Jim Britt
Windows Server System Center, Customer and Technologies Team (WSSC CAT)
Microsoft Corporation - 10-16-2013
=======================================================================
    Updated by Jim Britt on 3-5-2014: 
    
    Added support for exporting secrets on export process
    From within SMA.

    Added schema updates for credentialUserName and
    CredentialPassword to support ExportSecrets update.
=======================================================================
=======================================================================
.DESCRIPTION
This scripted solution can be leveraged for exporting
SMA Runbooks from an SMA environment to an XML based export file.
This solution leverages the existing SMA cmdlets for manipulating Runbooks
and assets and provides an atomic mechanism for sharing Runbooks between
environments.

This export solution supports the following
Runbook: Exporting to an XML and PS1 as an option using -ExportPS1
Schedules: Any schedules assigned will be exported if not already existing.
Variables: Any variables defined in the XML will be exported.
Credentials: Any credentials will be exported by name.
Tags: Any tags specified on the Runbook configuration will be exported
Description: Description of the Runbook will be exported
Log Settings: All logging settings defined for this Runbook will be exported

.EXAMPLE 

.\Export-SMARunbookToXML.ps1 -RunbookName "RunbookName" -ExportDirectory "c:\temp\exports" -ExportAssets $True

Exports a Runbook from SMA named "RunbookName" to the targeted folder (with assets) to an XML file.
Assets in this case are defined as Schedules, Credentials, Variables. 

.EXAMPLE 

.\Export-SMARunbooktoXML.ps1 -RunbookName "RunbookName" -ExportDirectory "c:\temp\exports" -ExportVariables $True

Exports a Runbook from SMA named "RunbookName" to the targeted folder (with variables) to an XML file.
$Null type variables are not exported. Exported variables are exported without data (just description and name)


.\Export-SMARunbooktoXML.ps1 -RunbookName "RunbookName" -ExportDirectory "c:\temp\exports" -ExportCredentials $True -ExportSchedules $True

Exports a Runbook from SMA named "RunbookName" to the targeted folder (with variables and credentials) to an XML file.
$Null type variables are not exported. Encrypted variables are exported without data (just description and name)

.EXAMPLE 
.\Export-SMARunbookToXML.ps1 -RunbookName "RunbookName" -ExportDirectory "c:\temp\exports" -ExportPS1 $True

Exports a Runbook from SMA named "RunbookName" to the targeted folder to an XML file (and a PS1 file to a PS1 subfolder).
PS1 files created will reflect Published and Draft versions if both exist (draft with a -draft appended to name) 

.EXAMPLE 
.\Export-SMARunbookToXML.ps1 -RunbookName "RunbookName" -ExportDirectory "c:\temp\exports" -ExportAssets $True -EnableScriptOutput $True

Exports a Runbook from SMA named "RunbookName" to the targeted folder (with assets exported) to an XML file.
Screen status is also outputed leveraging the -EnableScriptOutput for basic details on progress.


.LINK
http://aka.ms/BuildingClouds
#>


param
(
    # Runbook Name as named in SMA
    [parameter(Mandatory=$True)]
    [string]$RunbookName,

    # Example: "c:\exportDirectory"
    [parameter(Mandatory=$True)]
    [string]$ExportDirectory,

    # Example: "https://smaserver.contoso.com"
    [parameter(Mandatory=$False)]
    [string]$WebServiceEndpoint="https://localhost",

    # Windows or Basic
    [parameter(Mandatory=$false)]
    [ValidateSet("Windows", "Basic")]
    [string]$AuthenticationType="Windows",
        
    # Port 9090 defaults
    [parameter(Mandatory=$False)]
    [int]$Port=9090,
        
    # Leveraged for authenticating with alt creds
    [parameter(Mandatory=$false)]
    [pscredential]$cred=$Null,

    # Will provide output to console for status
    [boolean]$EnableScriptOutput,

    # Switch provided to save exported PS1 in export directory
    [parameter(Mandatory=$false)]
    [boolean]$ExportPS1,

    # Will export variables to XML if -ExportVars specified on command line
    [boolean]$ExportVars,

    # Will export credential names to XML if -ExportCreds specified on command line
    [boolean]$ExportCreds,

    # Will export all schedules to XML if -ExportSchedules is specified on command line
    [boolean]$ExportSchedules,

    # Will export all schedules to XML if -ExportSchedules is specified on command line
    [boolean]$ExportAssets,

    # Will export all variables that are defined to SMA that exist in the Runbook (undefined will not be exported)
    [boolean]$ExportOnlyDefinedVariables
)

Workflow Export-SMARunbookToXML
{
    param
    (
        # Runbook Name as named in SMA
        [parameter(Mandatory=$True)]
        [string]$RunbookName,

        # Example: "c:\exportDirectory"
        [parameter(Mandatory=$True)]
        [string]$ExportDirectory,

        # Example: "https://smaserver.contoso.com"
        [parameter(Mandatory=$false)]
        [string]$WebServiceEndpoint="https://localhost",

        # Windows or Basic
        [parameter(Mandatory=$false)]
        [string]$AuthenticationType="Windows",
        
        # Port 9090 defaults
        [parameter(Mandatory=$False)]
        [int]$Port=9090,
        
        # Leveraged for authenticating with alt creds
        [parameter(Mandatory=$false)]
        [pscredential]$cred,

        # Will provide output to console for status
        [boolean]$EnableScriptOutput,

        # Switch provided to save exported PS1 in export directory
        [parameter(Mandatory=$false)]
        [boolean]$ExportPS1,

        # Will export variables to XML if -ExportVars specified on command line
        [boolean]$ExportVars,

        # Will export credential names to XML if -ExportCreds specified on command line
        [boolean]$ExportCreds,

        # Will export all schedules to XML if -ExportSchedules is specified on command line
        [boolean]$ExportSchedules,

        # Will export all schedules to XML if -ExportSchedules is specified on command line
        [boolean]$ExportAssets,

        # Will export all variables that are defined to SMA that exist in the Runbook (undefined will not be exported)
        [boolean]$ExportOnlyDefinedVariables
    )
       
    # Comment out the below to show errors or change to "Continue"
    $ErrorActionPreference = "SilentlyContinue"
    
    # Display Params Passed
    ""
    $ScriptOutput = (get-date -format g).ToString() + "
Runbook Name: $RunbookName`r
Export Directory: $ExportDirectory`r
WebService Endpoint:$WebServiceEndpoint`r
AuthType:$AuthenticationType`r
Port:$Port`r
PSCred:$cred`r
ExportPS1?:$ExportPS1`r
ScriptOutput?:$EnableScriptOutput`r
ExportVars?:$ExportVars`r
ExportCreds?:$ExportCreds`r
ExportSchedules?:$ExportSchedules.`r
ExportAssets?:$ExportAssets.`r
ExportOnlyDefinedVariables?:$ExportOnlyDefinedVariables`n`n"

    if($EnableScriptOutput){ $ScriptOutput }

    #Validate $ExportDirectory exists
    $DirExists = Test-Path $ExportDirectory
    If (!$DirExists)
    {
        $CreateDir = New-Item -path $ExportDirectory -ItemType directory
        $ScriptOutput = (get-date -format g).ToString() + " $ExportDirectory didn't exist - creating."
        if($EnableScriptOutput){ $ScriptOutput }
    }

    # Get Runbook and Definition details to export
    $Runbook = Get-SmaRunbook -Name $RunbookName -WebServiceEndpoint $WebServiceEndpoint -Port $Port -AuthenticationType $AuthenticationType -Credential $cred
    if(!$Runbook)
    {
        "Runbook $RunbookName Not Found!"
        exit
    }
    
    # Get both Draft and Published to export into XML
    $PublishedRunbookDefinition = InlineScript{Get-SmaRunbookDefinition -Id $Using:Runbook.RunbookID -Type Published -WebServiceEndpoint $Using:WebServiceEndpoint -Port $Using:Port -AuthenticationType $Using:AuthenticationType -Credential $Using:Cred}
    $DraftRunbookDefinition = InlineScript{Get-SmaRunbookDefinition -Id $Using:Runbook.RunbookID -Type Draft -WebServiceEndpoint $Using:WebServiceEndpoint -Port $Using:Port -AuthenticationType $Using:AuthenticationType -Credential $Using:Cred}
  
    # Evalute if there is a PUblished Runbook - if not set variable appropriately
    # These are used for XML data rather than setting "null"
    if(!$PublishedRunbookDefinition.Content)
    {
        $PublishedRunbookContent = "No Published Version"
        $ScriptOutput = "Publshed Runbook doesn't exist for $RunbookName!"
        if($EnableScriptOutput){ $ScriptOutput }
    }
    else
    {
        $PublishedRunbookContent = $PublishedRunbookDefinition.Content
    }
    
    # Evaluate content differences between Published and Draft to ensure we are gathering both
    If ($PublishedRunbookDefinition.Content -eq $DraftRunbookDefinition.Content)
    {
        $DraftRunbookContent = "Draft Not Unique" 
        $ScriptOutput = "Publshed and draft are the same or no draft found for $RunbookName"
        if($EnableScriptOutput){ $ScriptOutput }
    } 
    else
    { 
        $DraftRunbookContent = $DraftRunbookDefinition.Content
        $ScriptOutput = "Draft exists - gathering draft details for $RunbookName"
        if($EnableScriptOutput){ $ScriptOutput }
    }
    
    # If Tags don't exist - set them to blank in XMl
    If (!$Runbook.Tags)
    {
        $RunbookTags = ""
        $ScriptOutput = "Runbook tags are not present for $RunbookName"
        if($EnableScriptOutput){ $ScriptOutput }
    }
    Else
    {
        $RunbookTags = $Runbook.Tags
    }

    # If Description is blank - set it to blank in XML
    If (!$RUnbook.Description)
    {
        $RunbookDescription = ""
        $ScriptOutput = "Runbook description is not present for $RunbookName"
        if($EnableScriptOutput){ $ScriptOutput }
    }
    Else
    {
        $RunbookDescription = $Runbook.Description
    }
    
    # InlineScript required since we are working with XML and variables
    InlineScript{
        # Build XML Definition for Export
        $TemplateSchema =
@'
<?xml version="1.0" encoding="UTF-8"?>
<Runbook>
    <Name></Name>
    <Tag></Tag>
    <Configuration>
        <Description></Description>
        <LogDebug></LogDebug>
        <LogVerbose></LogVerbose>
        <LogProgress></LogProgress>
    </Configuration>
    <Published>
        <Definition></Definition>
        <Credentials>
            <CredentialName></CredentialName>
            <CredentialDescription></CredentialDescription>
        </Credentials>
        <Variables>
            <VariableName></VariableName>
            <VariableValue></VariableValue>
            <VariableDescription></VariableDescription>
            <VariableType></VariableType>
            <IsEncrypted></IsEncrypted>
        </Variables>
    </Published>
    <Draft>
        <Definition></Definition>
        <Credentials>
            <CredentialName></CredentialName>
            <CredentialDescription></CredentialDescription>
			<CredentialUserName></CredentialUserName>
            <CredentialPassword></CredentialPassword>
        </Credentials>
        <Variables>
            <VariableName></VariableName>
            <VariableValue></VariableValue>
            <VariableDescription></VariableDescription>
            <VariableType></VariableType>
            <IsEncrypted></IsEncrypted>
        </Variables>
    </Draft>
    <Schedules>
        <ScheduleName></ScheduleName>
        <ScheduleDescription></ScheduleDescription>
        <ScheduleType></ScheduleType>
        <ScheduleNextRun></ScheduleNextRun>
        <ScheduleExpiryTime></ScheduleExpiryTime>
        <ScheduleDayInterval></ScheduleDayInterval>
    </Schedules>
</Runbook>
'@
        # Assign Schema to Variable
        [XML]$XMLExportFile = $TemplateSchema 

        # Build node data for export to XML
        $XMLExportFile.Runbook.Name = "$Using:RunbookName"
        $XMLExportFile.Runbook.Published.Definition = $Using:PublishedRunbookContent
        $XMLExportFile.Runbook.Draft.Definition = $Using:DraftRunbookContent
        $XMLExportFile.Runbook.Tag = $Using:RunbookTags
        $XMLExportFile.Runbook.Configuration.Description = $Using:RunbookDescription
        
        # Log values being converted to string to be supported by XML
        $XMLExportFile.Runbook.Configuration.LogDebug = [string]$Using:Runbook.LogDebug
        $XMLExportFile.Runbook.Configuration.LogVerbose = [string]$Using:Runbook.LogVerbose
        $XMLExportFile.Runbook.Configuration.LogProgress = [string]$Using:Runbook.LogProgress
        
        # SCHEDULES SECTION
        # Export schedules (if -exportSchedules is specified) [Schedules defined per Runbook]
        if($Using:ExportSchedules -or $Using:ExportAssets)
        {
            $RunbookSch = Get-SmaRunbook -Name $Using:RunbookName -WebServiceEndpoint $Using:WebServiceEndpoint -Port $Using:Port -AuthenticationType $Using:AuthenticationType -Credential $Using:Cred
            $RBSchedules = $RunbookSch.Schedules
            $SchedulesVariable = (@($XMLExportFile.Runbook.Schedules)[0]).Clone()
            foreach($RBSchedule in $RBSchedules)
            {
                If($RBSchedule.GetType().Name -eq "DailySchedule")
                {
                    $SchedulesVariable.ScheduleName = [string]$RBSchedule.Name 
                    $SchedulesVariable.ScheduleDescription = [string]$RBSchedule.Description
                    $SchedulesVariable.ScheduleType = [string]$RBSchedule.GetType().Name
                    $SchedulesVariable.ScheduleNextRun = [string]$RBSchedule.NextRun
                    $SchedulesVariable.ScheduleExpiryTime = [string]$RBSchedule.ExpiryTime
                    $SchedulesVariable.ScheduleDayInterval = [string]$RBSchedule.DayInterval
                    $XMLExportvariable = $XMLExportFile.Runbook.AppendChild($SchedulesVariable)
                    $SchedulesVariable = $SchedulesVariable.clone()
                }
            }
        }

        # VARIABLES SECTION
        # Export Variables (if -exportVars is specified)
        if($Using:ExportVars -or $Using:ExportAssets)
        {
            function ProcessVariableData
            {
                # What content type are we evaluating (Draft or Published)
                param([string]$ContentType)
                
                # Determine type and set appropriately
                if($ContentType -eq "Published")
                {
                    $RunbookContent = $Using:PublishedRunbookContent
                    
                    # Assign New Variable for node operations in XML [clone existing structure]
                    $NewVariable = (@($XMLExportFile.Runbook.Published.Variables)[0]).Clone()
                }
                
                If($ContentType -eq "Draft")
                {
                    $RunbookContent = $Using:DraftRunbookContent
                    # Assign New Variable for node operations in XML [clone existing structure]
                    $NewVariable = (@($XMLExportFile.Runbook.Draft.Variables)[0]).Clone()
                }

                # Set string to search for variables
                $VarCheck = "Get-AutomationVariable -Name '"
                
                # Build and array to start Runbook PowerShell
                $ArrayString=@()
            
                # Split Runbook into array items using new line character
                $ArrayString = $RunbookContent.Split("`n") 

                # Locate lines that contain "Get-Automa..."
                foreach ($ArrayItem in $ArrayString)
                {
                    If($ArrayItem -match $VarCheck)
                    {
                        $VarStart = $ArrayItem.IndexOf($VarCheck) + $VarCheck.Length
                        $VarEnd = $ArrayItem.IndexOf("'",$VarStart)
                        $VariableValue = $ArrayItem.Substring($VarStart,$VarEnd-$VarStart) 
                        $AssignedVariable = $ArrayItem.Substring($VarStart,$VarEnd-$VarStart)
                        if($AssignedVariable)
                        {
                            # Get details of varable used
                            $DefinedVariable = Get-SmaVariable -Name $AssignedVariable -WebServiceEndpoint $Using:WebServiceEndpoint -Port $Using:Port -AuthenticationType $Using:AuthenticationType -Credential $Using:Cred
                            # If Undefined set appropriately (unable to get variable from SMA)
                            if(!$DefinedVariable -and !$Using:ExportOnlyDefinedVariables)
                            {
                                $NewVariable.VariableName = [string]$AssignedVariable
                                $NewVariable.VariableDescription = [string]""
                                $NewVariable.VariableValue = [string]"string"
                                $NewVariable.VariableType = [string]"String"
                                $NewVariable.IsEncrypted = [string]"False"
                                "Unable to get information on Variable $AssignedVariable.  $AssignedVariable exported by name only."
                            }
                            elseif($DefinedVariable -and $Using:ExportOnlyDefinedVariables)
                            {
                                # Didn't use -ExportOnlyDefinedVariables
                                "Unable to get information on Variable $AssignedVariable.  $AssignedVariable not exported due to -ExportOnlyDefinedVariables switch"
                            }

                            # Able to gather from SMA - Exporting Details
                            if($DefinedVariable)
                            {
                                # Set variable information from SMA in XML
                                $NewVariable.VariableName = [string]$DefinedVariable.Name
                                $NewVariable.VariableDescription = [string]$DefinedVariable.Description
                                $NewVariable.VariableValue = [string]$DefinedVariable.Value
                                $NewVariable.VariableType = [string]$DefinedVariable.Value.GetType()
                                $NewVariable.IsEncrypted = [string]$DefinedVariable.IsEncrypted
                            }
                        
                            If($ContentType -eq "Published")
                            {
                                $XMLExportvariable = $XMLExportFile.Runbook.Published.AppendChild($NewVariable) 
                            }
                        
                            If($ContentType -eq "Draft")
                            {
                                $XMLExportvariable = $XMLExportFile.Runbook.Draft.AppendChild($NewVariable) 
                            }
                            $NewVariable = $NewVariable.clone()
                        }
                    }
                }
            
            }
            # Gathering variables for Published content (if there is a published version)
            If($Using:PublishedRunbookDefinition.Content){ProcessVariableData -ContentType "Published"}
            
            # Gathering Variables for Draft Content (if there is a draft)
            If($Using:DraftRunbookDefinition.Content){ProcessVariableData -ContentType "Draft"}
        }


        if($Using:ExportCreds -or $Using:ExportAssets)
        {
            function ProcessCredentialData
            {
                # What content type are we evaluating (Draft or Published)
                param([string]$ContentType)
                
                # Determine type and set appropriately
                if($ContentType -eq "Published")
                {
                    $RunbookContent = $Using:PublishedRunbookContent
                    
                    # Assign New Variable for node operations in XML [clone existing structure]
                    $NewVariable = (@($XMLExportFile.Runbook.Published.Credentials)[0]).Clone()
                }
                
                If($ContentType -eq "Draft")
                {
                    $RunbookContent = $Using:DraftRunbookContent

                    # Assign New Variable for node operations in XML [clone existing structure]
                    $NewVariable = (@($XMLExportFile.Runbook.Draft.Credentials)[0]).Clone()
                }
            
                $ScriptOutput = "Export of creds specified.  Processing..."
                if($EnableScriptOutput){ $ScriptOutput }
                
                # Define the string to search for creds
                $CredCheck = "Get-AutomationPSCredential -Name '"
        
                # Build and array to start Runbook PowerShell
                $ArrayString=@()
            
                # Split Runbook into array items using new line character
                $ArrayString = $RunbookContent.Split("`n") 

                # Locate lines that contain "Get-Automa..."
                foreach ($ArrayItem in $ArrayString)
                {
                    If($ArrayItem -match $CredCheck)
                    {
                        $VarStart = $ArrayItem.IndexOf($CredCheck) + $CredCheck.Length
                        $VarEnd = $ArrayItem.IndexOf("'",$VarStart)
                        $VariableValue = $ArrayItem.Substring($VarStart,$VarEnd-$VarStart) 
                        $AssignedCredential = $ArrayItem.Substring($VarStart,$VarEnd-$VarStart)
                        if($AssignedCredential)
                        {
                            # Get details of credential used
                            $DefinedCredential = Get-SmaCredential -Name $AssignedCredential -WebServiceEndpoint $Using:WebServiceEndpoint -Port $Using:Port -AuthenticationType $Using:AuthenticationType -Credential $Using:Cred
                            
							# Credential not defined in SMA
							if(!$DefinedCredential)
                            {
                                $NewVariable.CredentialName = [string]$AssignedCredential
                                $NewVariable.CredentialDescription = [string]""
                                $NewVariable.CredentialUserName = [string]"Contoso\user"
                                $NewVariable.CredentialPassword = [string]"password"
                                "Unable to get information on credential $AssignedCredential.  $AssignedCredential created with temp values."                                                       
                            }
                        
                            $ Valid credential in SMA
                            if($DefinedCredential)
                            {
								# Set credential information in XML
								$NewVariable.CredentialName = [string]$DefinedCredential.Name
								$NewVariable.CredentialDescription = [string]$DefinedCredential.Description
							}

                            # Determine which Runbook type we are working with
                            if($ContentType -eq "Published")
                            {
                                $XMLExportvariable = $XMLExportFile.Runbook.Published.AppendChild($NewVariable) 
                            }
                            
                            if($ContentType -eq "Draft")
                            {
                                $XMLExportvariable = $XMLExportFile.Runbook.Draft.AppendChild($NewVariable) 
                            }
                            $NewVariable = $NewVariable.clone()
                        }
                    }
                }
            
            }
            # Gathering creds for Published content (if there is a published version)
            If($Using:PublishedRunbookDefinition.Content){ProcessCredentialData -ContentType "Published"}
                
            # Gathering creds for Draft content (if there is a draft version)
            If($Using:DraftRunbookDefinition.Content){ProcessCredentialData -ContentType "Draft"}
        }

        #clean up empty variables from published section
        $XMLExportFile.Runbook.Published.Variables | 
        Where-Object { $_.VariableName -eq "" } |
            ForEach-Object  { [void]$XMLExportFile.Runbook.Published.RemoveChild($_) }

        $XMLExportFile.Runbook.Draft.Variables | 
        Where-Object { $_.VariableName -eq "" } |
            ForEach-Object  { [void]$XMLExportFile.Runbook.Draft.RemoveChild($_) }

        $XMLExportFile.Runbook.Published.Credentials | 
        Where-Object { $_.CredentialName -eq "" } |
            ForEach-Object  { [void]$XMLExportFile.Runbook.Published.RemoveChild($_) }

        $XMLExportFile.Runbook.Draft.Credentials | 
        Where-Object { $_.CredentialName -eq "" } |
            ForEach-Object  { [void]$XMLExportFile.Runbook.Draft.RemoveChild($_) }

        $XMLExportFile.Runbook.Schedules | 
        Where-Object { $_.ScheduleName -eq "" } |
            ForEach-Object  { [void]$XMLExportFile.Runbook.RemoveChild($_) }

        # Output Runbook from SMA
        $XMLExportFile.Save("$Using:ExportDirectory\$Using:RunbookName.xml")
        $ScriptOutput = "Writing out Runbook XML to $Using:ExportDirectory\$Using:RunbookName.xml"
        if($Using:EnableScriptOutput){ $ScriptOutput }

        
        # Export PS1 if ExportPS1 switch was specified on command line
        # Export PS1 if ExportPS1 switch was specified on command line
        if($Using:ExportPS1)
        {
            # Create export dir for PS1 if it doesn't exist
            $DirExists = Test-Path "$Using:ExportDirectory\PS1"
            If (!$DirExists)
            {
                $CreatePS1Dir = New-Item -path "$Using:ExportDirectory\PS1" -ItemType directory
                $ScriptOutput = "$ExportDirectory\PS1 didn't exist - creating."
                if($EnableScriptOutput){ $ScriptOutput }
            }
            
            # Export PS1             
            If($Using:PublishedRunbookContent -ne "No Published Version"){$Using:PublishedRunbookContent | Out-File "$Using:ExportDirectory\PS1\$Using:RunbookName.ps1"}
            if($Using:DraftRunbookContent -ne "Draft Not Unique"){$Using:DraftRunbookContent | Out-File "$Using:ExportDirectory\PS1\$Using:RunbookName-draft.ps1"}
            
            $ScriptOutput = "Scripts exported to $Using:ExportDirectory\PS1."
            if($Using:EnableScriptOutput){ $ScriptOutput }
        }
    }
}
# All parameters are defined and will consume values if specified
# Note: RunbookName and ExportDirectory are the only required parameters
Export-SMARunbookToXML -RunbookName $RunbookName -ExportDirectory $ExportDirectory -WebServiceEndpoint $WebServiceEndpoint `
 -Cred $cred -EnableScriptOutput $EnableScriptOutput -ExportPS1 $ExportPS1 -ExportVars $ExportVars -ExportCreds $ExportCreds `
 -ExportSchedules $ExportSchedules -ExportAssets $ExportAssets -ExportOnlyDefinedVariables $ExportOnlyDefinedVariables

