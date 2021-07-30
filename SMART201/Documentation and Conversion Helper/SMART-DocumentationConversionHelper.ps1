##########################################################################################
# SMA Runbook Toolkit - Documentation and Conversion Helper
# Version : 2.01
# Refer to this blog post for more details :
# http://blogs.technet.com/b/privatecloud/archive/2014/05/08/updated-tool-smart-documentation-and-conversion-helper-for-your-orchestrator-runbooks.aspx
# Windows Server and System Center Customer, Architecture and Technologies (CAT) team
# Please send feedback to brunosa@microsoft.com
# Parameters (all optional) :
#              -DBServer : optional parameter to specify the server hosting the
#              Orchestrator database. If not specified, the script defaults to 'localhost',
#              and will give you the option to specific another server in the GUI
#              if the connection fails - the tool also remembers the last server used
#              -DBPort : optional parameter to specify the port on which SQL Server
#              is listening. If not specified, the script defaults to '1433',
#              and will give you the option to specific another port in the GUI
#              if the connection fails
#              -DBName : optional parameter to specify the datbase name.
#              If not specified, the script defaults to 'Orchestrator'.
#              This parameter is not exposed in the GUI, only via command line
#              and in the script.
# You may also want to update the PowerShell ISE location if needed
# (see the ISELocation variable at line 35)
##########################################################################################

    param (
    [String]$DBServer,
    [String]$DBPort,
    [String]$DBName
    )

If ($DBServer) {$DefaultDatabaseServer = $DBServer} else {$DefaultDatabaseServer ="localhost"}
If ($DBPort) {$DefaultDatabasePort = $DBPort} else {$DefaultDatabasePort ="1433"}
If ($DBName) {$DefaultDatabaseName = $DBName} else {$DefaultDatabaseName ="Orchestrator"}

$ToolVersion = "2.01"
$Global:ISELocation = "c:\windows\system32\WindowsPowerShell\v1.0\"

function ParseAndWriteProperty
# This function takes the value of an activity property, and tries to parse published data
# and variables, to replace the data by readable content (actual names of the published data
# and variables)
{
    param (
    [String]$Prefix,
    [String]$TmpProperty,
    [String]$NbTab,
    [String]$ExportMode
    )

    $myConnection3 = New-Object System.Data.SqlClient.SqlConnection $Global:SQLConnstr 
    $myConnection3.Open()

    $TmpProperty=$TmpProperty.Replace("\``d.T.~Ed/", "###PUBDATA###")
    $TmpProperty=$TmpProperty.Replace("\``d.T.~Vb/", "###VAR###")
    
    If (($TmpProperty.Contains("###PUBDATA###") -eq $True) -Or ($TmpProperty.Contains("###VAR###") -eq $True))
    
            {            
            #Let's work through the potential published data first, avoiding the first item in the split (not a published data)             
            $ComputedProperties = $TmpProperty -split("###PUBDATA###{")
            For ($i=1; $i -lt $ComputedProperties.Length; $i++){
                #PublishedData generally has one GUIDs {activity}.publisheddataincleartext, but published data Initialize Data activities are formatted like {activity}.{parameter}
                #Let's work through the first GUID in all cases
                $ComputedProperties[$i]= ($ComputedProperties[$i] -split ("###PUBDATA###"))[0]
                $OutputID = "{" + $ComputedProperties[$i].Substring(0, 37)
                #Note : Adding a specific check to make sure OutputID is a GUID. 
                #Otherwise it fails if script has a invoke-command with published data for the computername and no explicit scriptblock option followed by {
                If ($OutputID -match ("^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$")) {
                                $SqlQuery = "select name, ObjectType from objects where UniqueID = '" + $OutputID + "'"
                                $myCommand3 = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection3
                                $dr3 = $myCommand3.ExecuteReader()
                                while ($dr3.Read())
                                    {
                                    $OutputName = $dr3["name"]
                                    $OutputType = $dr3["ObjectType"].ToString
                                    If ($ExportMode -eq "DOC"){
                                        If ($Global:ActivityDependenciesActivityNames.Contains($OutputName) -eq $False) {
                                            $Global:ActivityDependenciesActivityNames += $OutputName
                                            $Global:ActivityDependenciesActivityTypes += $dr3["ObjectType"]
                                            }
                                        }
                                    }
                                $dr3.Close()
                                $TmpProperty = $TmpProperty.Replace($OutputID, "{Activity:" + $OutputName + "}")
                                #Now we can check if there is a GUID in another GUID
                                $NumberOfGUIDS = $ComputedProperties[$i].Split("{").Length - 1
                                If ($NumberOfGUIDS -eq 1) {
                                        $OutputSuffix = "{" + ($ComputedProperties[$i].Split("{")[1]).Substring(0, 36) + "}"
                                        $SqlQuery = "select value from CUSTOM_START_PARAMETERS where UniqueID = '" + $OutputSuffix + "'"
                                        $myCommand3 = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection3
                                        $dr3 = $myCommand3.ExecuteReader()
                                        while ($dr3.Read())
                                                {
                                                $OutputSuffixName = $dr3["value"]
                                                }
                                        $dr3.Close()
                                        $OutputSuffix = "}." + $OutputSuffix
                                        $TmpProperty = $TmpProperty.Replace($OutputSuffix, (".PublishedData:" + $OutputSuffixName + "}"))
                                    }
                                    else
                                    {
                                    $OutputSuffix = "}." + $ComputedProperties[$i].Split(".")[1]
                                    $TmpProperty = $TmpProperty.Replace($OutputSuffix, ".PublishedData:" + $ComputedProperties[$i].Split(".")[1] + "}")
                                    }

                }
            }
            #Let's work through the potential variables too, avoiding the first item in the split (not a variable)         
            $ComputedProperties = $TmpProperty -split("###VAR###{")
            For ($i=1; $i -lt $ComputedProperties.Length; $i++)
                {
                $OutputVariable = "{" + $ComputedProperties[$i].Substring(0, 37)
                If ($OutputVariable -match ("^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$"))
                    {
                    $SqlQuery = "select name from objects where UniqueID = '" + $OutputVariable + "'"
                    $myCommand3 = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection3
                    $dr3 = $myCommand3.ExecuteReader()
                    while ($dr3.Read()) {$OutputVariableName = $dr3["name"]}
                    $dr3.Close()
                    $SqlQuery = "select value from variables where UniqueID = '" + $OutputVariable + "'"
                    $myCommand3 = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection3
                    $dr3 = $myCommand3.ExecuteReader()
                    while ($dr3.Read())
                         {
                         $OutputVariableValue = $dr3["value"]
                         #Let's check if this is an encrypted or empty variable
                         If ([System.DBNull]::Value.Equals($dr3["value"]) -eq $False)
                            {$OutputVariableValue=$OutputVariableValue.Replace("\``d.T.~De/", "###ENCRYPTEDDATA###")}
                            else
                            {$OutputVariableValue="[Empty Value]"}
                         }
                    $dr3.Close()
                    If ($ExportMode -eq "DOC")
                         {
                         $Global:ActivityDependenciesVariableNames += $OutputVariableName
                         $Global:ActivityDependenciesVariableValues += $OutputVariableValue
                         }
                    $TmpProperty = $TmpProperty.Replace($OutputVariable, ("{Variable:" + $OutputVariableName + "}"))
                    $Global:FlagVariables = $True
                    If ($Global:FlagVariablesList.Contains("{" + $OutputVariableName + "}"))
                          {$Global:FlagVariablesNumber[$Global:FlagVariablesList.IndexOf("{" + $OutputVariableName + "}")] = $Global:FlagVariablesNumber[$Global:FlagVariablesList.IndexOf("{" + $OutputVariableName + "}")] +1}
                          else
                             {
                             $Global:FlagVariablesList+= "{" + $OutputVariableName + "}"
                             $Global:FlagVariablesNumber+= 1
                             If ($OutputVariableValue.Contains("###ENCRYPTEDDATA###"))
                                  {$Global:FlagVariablesValue+= "[Encrypted Data]"}
                                  else
                                  {$Global:FlagVariablesValue+= $OutputVariableValue}
                             }
                    }
                }
            
            #We work on the TmpProperty, to extract the published data items
            $TmpProperty = $TmpProperty.Replace("###PUBDATA###", "####")
            $TmpProperty = $TmpProperty.Replace("###VAR###", "####")
            $TmpProperty = $TmpProperty.Replace("`r`n", "")
            $ComputedProperties = $TmpProperty -split("####")
            $FullPty = ""
            ForEach ($ComputedProperty In $ComputedProperties){
                If ($ComputedProperty -ne ""){$FullPty = $FullPty + $ComputedProperty}
            }
            $Lines = $FullPty -split("`n")
            $i = 0
            ForEach ($Line In $Lines)
                {
                If ($i -eq 0)
                    {WriteToFile -ExportMode $ExportMode -Add "$Prefix $Line" -NbTab $NbTab}
                    Else
                    {WriteToFile -ExportMode $ExportMode -Add $Line -NbTab $NbTab}
                $i = $i + 1
                }
            
            }
        else
            # There is no published data or variable in the value of this property
            {
            If ($TmpProperty.Contains("\``d.T.~De/") -eq $True) { $TmpProperty = "[Encrypted Data]"}
            WriteToFile -ExportMode $ExportMode -Add "$Prefix $TmpProperty" -NbTab $NbTab
            }
    $myConnection3.Close()
}


function LinkCondition
# This function retrieves the condition on a link betwen activities (if any)
{
    param (
    [String]$LinkID
    )

    $output = ""
    $OutputID = ""
    $SqlQuery = "select condition, data, value from TRIGGERS where ParentID = '" + $LinkID + "'"
    $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
    $dr = $myCommand.ExecuteReader()
    $OffsetBracket=0
    while ($dr.Read())
            {
            $OutputID = ($dr["data"]).Substring(0, 38)
            $NumberofGUIDs = ($dr["data"]).Split("{").Count - 1
            Switch ($dr["condition"])
                {
                "isgreaterthan" {$Outputcondition = "-gt"}
                "isgreaterthanorequalto" {$Outputcondition = "-ge"}
                "islessthan" {$Outputcondition = "-lt"}
                "islessthanorequalto" {$Outputcondition = "-le"}
                "equals" {$Outputcondition = "-eq"}
                "doesnotequal" {$Outputcondition = "-ne"}
                "" {$Outputcondition = "{linkcondition:returns}" ; $OffsetBracket = 1}                
                default
                    # "contains" "doesnotcontain" "endswith" "startswith" "doesnotmatchpattern" "matchespattern"
                    {
                    $Outputcondition = "{linkcondition:" + $dr["condition"] + "}"
                    $Global:FlagStringcondition = $True
                    $OffsetBracket = 1
                    }
                }
            $Output = "If (" + $dr["data"] + " " + $Outputcondition + " `"" + $dr["value"] + "`") {"
            }
    $dr.Close()
    If ($OutputID -ne "")
            {
            #There was a condition, let's convert the published data activity (first part)
            $SqlQuery = "select name, ObjectType from objects where UniqueID = '" + $OutputID + "'"
            $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
            $dr = $myCommand.ExecuteReader()
            while ($dr.Read()) {$OutputName = $dr["name"]}
            $dr.Close()
            $Output = $Output.Replace($OutputID, "{Activity:" + $OutputName + "}")
            #Let's check if there is a second GUID to convert - only applicable when it's published data from an initialize data activity
            $NumberofGUIDs = $output.Split("{").Count - 1 -$OffsetBracket
            If ($NumberofGUIDs -eq 3)
                {
                $OutputSuffix = "{" + $output.Split("{")[2].Substring(0, 36) + "}"
                $SqlQuery = "select value from CUSTOM_START_PARAMETERS where UniqueID = '" + $OutputSuffix + "'"
                $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
                $dr = $myCommand.ExecuteReader()
                while ($dr.Read()) {$OutputSuffixName = $dr["value"]}                    
                $dr.Close()
                $OutputSuffix = "}." + $OutputSuffix
                $Output = $output.Replace($OutputSuffix, ".PublishedData" + $OutputSuffixName + "}")
                }
            }
    Return $Output
}


function AppendActivityDetails
# This function is being called to fill the details of a specific activity
# It does that in a different manner depending on the type of activity
# It also calls ParseAndWriteProperty as needed, when parsing properties values
{
    param (
    [String]$ActivityID,
    [String]$ActivityDetailsShort,
    [String]$ActivityType,
    [Int]$NbTab,
    [String]$ExportMode
    )

    $DoNotExportProperties = @("UniqueID", "ExecutionData", "CachedProps")
    $PubDataNames = @()
    $PubDataTypes = @()
    $PubDataValues = @()
    $XmlNames = @()
    $XmlValues = @()
    $XmlPatterns = @()
    $XmlRelations = @()

    $myConnection2 = New-Object System.Data.SqlClient.SqlConnection $Global:SQLConnstr 
    $myConnection2.Open()

    #First, let's retrieve the full name of the activity type
    $SqlQuery = "select Name from OBJECTTYPES where UniqueID = '{" + $ActivityType + "}'"
    $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
    $dr = $myCommand.ExecuteReader()
    while ($dr.Read()) {$ActivityTypeName = $dr["Name"]}
    $dr.Close()
    
    write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Working on activity : $ActivityDetailsShort (Activity Type : $ActivityTypeName)"
    WriteToFile -ExportMode $ExportMode -Add "" -NbTab $NbTab

    If ($SkeletonCB.IsChecked -eq $True)
        {WriteToFile -ExportMode $ExportMode -Add $ActivityDetailsShort -NbTab $NbTab}
        else
        {
        Switch ($ActivityType)
            {
            "ed7f2a41-107a-4b74-bafe-adae63632b79"
            #This is a Powershell activity
                {
                WriteToFile -ExportMode $ExportMode -Add "# START ACTIVITY - $ActivityDetailsShort (Activity Type : Run .NET Script)" -NbTab $NbTab
                #retrieve script and script type (Powershell, C#, VB.NET, JScript)
                $SqlQuery = "select ScriptType, ScriptBody from RUNDOTNETSCRIPT where UniqueID = '" + $ActivityID + "'"
                $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
                $dr = $myCommand.ExecuteReader()
                while ($dr.Read())
                        {
                        $TmpProperty = $dr["ScriptBody"]
                        $ScriptType = $dr["ScriptType"]
                        }
                $dr.Close()
                WriteToFile -ExportMode $ExportMode -Add "# Script Type = $ScriptType" -NbTab $NbTab
                If ($ScriptType -eq "PowerShell")
                    {ParseAndWriteProperty -ExportMode $ExportMode -Prefix "" -TmpProperty $TmpProperty -NbTab $NbTab}
                    else
                    {ParseAndWriteProperty -ExportMode $ExportMode -Prefix "# " -TmpProperty $TmpProperty -NbTab $NbTab}
                #retrieve published data too
                $SqlQuery = "select publisheddata from RUNDOTNETSCRIPT where UniqueID = '" + $ActivityID + "'"
                $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
                $dr = $myCommand.ExecuteReader()
                while ($dr.Read()) {[string]$TmpProperty = $dr["publisheddata"]}
                $dr.Close()
                If ($TmpProperty -eq "")
                          {WriteToFile -ExportMode $ExportMode -Add "# Published Data - None" -NbTab $NbTab}
                          else
                          {
                            $PubDataNames.Clear()
                            $PubDataTypes.Clear()
                            $PubDataValues.Clear()
                            $xmlDoc = New-Object System.Xml.XmlDocument
                            $TmpProperty = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" + $TmpProperty
                            [System.Xml.XmlDocument]$xmlDoc.LoadXml($TmpProperty)
                            $Input = New-Object System.Xml.XmlNodeReader $xmlDoc
                            While ($Input.Read()){
                              If ($Input.NodeType -eq [System.Xml.XmlNodeType]::Element){
                                   switch ($Input.Name){
                                        "Name" {$PubDataNames+=$Input.ReadString()}
                                        "Type"{$PubDataTypes+=$Input.ReadString()}
                                        "Variable"{$PubDataValues+=$Input.ReadString()}
                                    }
                               }
                            }
                            $Input.Close()
                            $PubDataOutput = ""
                            ForEach ($PubDataName In $PubDataNames){
                                  $PubDataOutput = $PubDataOutput + " Name : " + $PubDataName + " / Type : " + $PubDataTypes.Item($PubDataNames.IndexOf($PubDataName)) + " / Value : " + $PubDataValues.Item($PubDataNames.IndexOf($PubDataName)) + " - "
                            }
                            WriteToFile -ExportMode $ExportMode -Add "# Published Data - $PubDataOutput" -NbTab $NbTab
                          }
                WriteToFile -ExportMode $ExportMode -Add "# END ACTIVITY - $ActivityDetailsShort" -NbTab $NbTab
                }
            "6c576f3d-e927-417a-b145-5d3eff9c995f"
            #This is an initialize data activity
                {
                WriteToFile -ExportMode $ExportMode -Add "# START ACTIVITY - $ActivityDetailsShort (Activity Type : Initialize Data)" -NbTab $NbTab
                WriteToFile -ExportMode $ExportMode -Add "# Parameters were added in the workflow definition" -NbTab $NbTab
                $SqlQuery = "select value, type from CUSTOM_START_PARAMETERS where ParentID = '" + $ActivityID + "'"
                $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
                $dr = $myCommand.ExecuteReader()
                while ($dr.Read())
                        {
                        WriteToFile -ExportMode $ExportMode -Add ("# " + $dr["value"] + "=> [" + $dr["type"] + "]$" + $dr["value"].Replace(" ", "_")) -NbTab $NbTab
                        }
                $dr.Close()
                WriteToFile -ExportMode $ExportMode -Add "# END ACTIVITY - $ActivityDetailsShort" -NbTab $NbTab
                }
            "fa70125f-267e-4065-a4f6-d5493167d663"
            #This is a return data activity
                {
                WriteToFile -ExportMode $ExportMode -Add "# START ACTIVITY - $ActivityDetailsShort (Activity Type : Return Data)" -NbTab $NbTab
                $SqlQuery = "select [Key],Value from PUBLISH_POLICY_DATA where ParentID = '" + $ActivityID + "'"
                $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
                $dr = $myCommand.ExecuteReader()
                while ($dr.Read())
                        {
                        $TmpProperty = $dr["value"]
                        ParseAndWriteProperty -ExportMode $ExportMode -Prefix ("# " + $dr["key"] + " = ") -TmpProperty $TmpProperty -NbTab $NbTab
                        }
                $dr.Close()
                WriteToFile -ExportMode $ExportMode -Add "# END ACTIVITY - $ActivityDetailsShort" -NbTab $NbTab
                $Global:FlagReturnData = $True
                If ($Global:FlagReturnDataList.Contains($ActivityName))
                     {$Global:FlagReturnDataNumber[$Global:FlagReturnDataList.IndexOf($ActivityDetailsShort)] = $Global:FlagReturnDataNumber[$Global:FlagReturnDataList.IndexOf($ActivityDetailsShort)] +1}
                   else
                     {
                     $Global:FlagReturnDataList+= $ActivityDetailsShort
                     $Global:FlagReturnDataNumber+= 1
                     }
                }
            default
            #This is another activity
                {
                WriteToFile -ExportMode $ExportMode -Add "# START ACTIVITY - $ActivityDetailsShort (Activity Type : $ActivityTypeName)" -NbTab $NbTab
                $SqlQuery = "select PrimaryDataTable from ObjectTypes, Objects where ObjectTypes.UniqueID = Objects.ObjectType and Objects.UniqueID = '" + $ActivityID + "'"
                $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
                $dr = $myCommand.ExecuteReader()
                $TmpTableInit = $False
                while ($dr.Read())
                        {
                        If ([System.DBNull]::Value.Equals($dr["PrimaryDataTable"]) -eq $False) {
                            $TmpTable = $dr["PrimaryDataTable"]
                            $TmpTableInit = $True
                            }
                        }
                $dr.Close()
                If ($TmpTableInit -eq $True)
                        {
                        $SqlQuery = "SELECT name FROM syscolumns WHERE id = (SELECT id FROM sysobjects WHERE name= '" + $TmpTable + "') ORDER by colorder"
                        $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
                        $dr = $myCommand.ExecuteReader()
                        while ($dr.Read())
                            {
                            If ($DoNotExportProperties.Contains($dr["name"]) -eq $False) {
                                $SqlQuery = "select " + $dr["name"] + " from " + $TmpTable + " where uniqueID = '" + $ActivityID + "'"
                                $myCommand2 = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection2
                                $dr2 = $myCommand2.ExecuteReader()
                                while ($dr2.Read()) {$TmpProperty = $dr2[$dr["name"]]}
                                $dr2.Close()
                                $TmpPropertyType = "default"
                                If (([System.DBNull]::Value.Equals($tmpProperty) -eq $False) -And ($dr["name"] -eq "Filters")) {$TmpPropertyType = "Filters"}
                                If (([System.DBNull]::Value.Equals($tmpProperty) -eq $False) -And ($TmpProperty.Length -gt 26)) { If ($TmpProperty.Substring(0,26) -eq "<ItemRoot><Entry><FieldId>") {$TmpPropertyType = "Filters"}}
                                If (([System.DBNull]::Value.Equals($tmpProperty) -eq $False) -And ($TmpProperty.Length -gt 29)) { If ($TmpProperty.Substring(0,29) -eq "<ItemRoot><Entry><PropertyId>") {$TmpPropertyType = "QIKProperties"}}
                                switch($TmpPropertyType){
                                    "Filters"{
                                        $XmlNames = @()
                                        $XmlValues = @()
                                        $XmlPatterns = @()
                                        $XmlRelations = @()
                                        $xmlDoc = New-Object System.Xml.XmlDocument
                                        $TmpProperty = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" + $TmpProperty
                                        [System.Xml.XmlDocument]$xmlDoc.LoadXml($TmpProperty)
                                        $Input = New-Object System.Xml.XmlNodeReader $xmlDoc
                                        While ($Input.Read()){
                                            If ($Input.NodeType -eq [System.Xml.XmlNodeType]::Element){
                                                switch ($Input.Name){
                                                    "FieldID" {$XmlNames+=($Input.ReadString()).Split("/")[1].Split("\")[0]}
                                                    "FilterValue" {$XmlValues+=(($Input.ReadString()).Replace("\``~F/", "###DELIM###") -split("###DELIM###"))[1]}
                                                    "RelationID"
                                                        {
                                                        switch(($Input.ReadString()).Split("/")[1].Split("\")[0]){
                                                            "0" {$XmlRelations+="equals"}
                                                            "1" {$XmlRelations+="does not equal"}
                                                            "2" {$XmlRelations+="contains"}
                                                            "3" {$XmlRelations+="does not contain"}
                                                            "4" {$XmlRelations+="matches pattern"}
                                                            "5" {$XmlRelations+="does not match pattern"}
                                                            "6" {$XmlRelations+="less than or equal to"}
                                                            "7" {$XmlRelations+="greater than or equal to"}
                                                            "8" {$XmlRelations+="starts with"}
                                                            "9" {$XmlRelations+="ends with"}
                                                            "10" {$XmlRelations+="less than"}
                                                            "11" {$XmlRelations+="greater than"}
                                                            "13" {$XmlRelations+="after"}
                                                            "14" {$XmlRelations+="before"}
                                                            "default" {$XmlRelations+="unknown filter condition"}
                                                            }
                                                        }
                                                    }
                                            }
                                        }
                                        $Input.Close()
                                        $XmlOutput = ""
                                        ForEach ($XmlName In $XmlNames){
                                            ParseAndWriteProperty -ExportMode $ExportMode -Prefix ("# Filter : $XmlName [" + $XmlRelations.Item($XmlNames.IndexOf($XmlName)) + "]") -TmpProperty $XmlValues.Item($XmlNames.IndexOf($XmlName)) -NbTab $NbTab
                                        }
                                }
                                    "QIKProperties"{
                                        $XmlNames = @()
                                        $XmlValues = @()
                                        $XmlPatterns = @()
                                        $XmlRelations = @()
                                        $xmlDoc = New-Object System.Xml.XmlDocument
                                        $TmpProperty = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" + $TmpProperty
                                        [System.Xml.XmlDocument]$xmlDoc.LoadXml($TmpProperty)
                                        $Input = New-Object System.Xml.XmlNodeReader $xmlDoc
                                        While ($Input.Read()){
                                            If ($Input.NodeType -eq [System.Xml.XmlNodeType]::Element){
                                                switch ($Input.Name){
                                                    "PropertyName" {$XmlNames+=($Input.ReadString()).Split("/")[1].Split("\")[0]}
                                                    "PropertyValue" {$XmlValues+=(($Input.ReadString()).Replace("\``~F/", "###DELIM###") -split("###DELIM###"))[1]}
                                                    }
                                            }
                                        }
                                        $Input.Close()
                                        $XmlOutput = ""
                                        ForEach ($XmlName In $XmlNames){
                                            ParseAndWriteProperty -ExportMode $ExportMode -Prefix ("# $XmlName = ") -TmpProperty $XmlValues.Item($XmlNames.IndexOf($XmlName)) -NbTab $NbTab
                                        }
                                }
                                "default"
                                {ParseAndWriteProperty -ExportMode $ExportMode -Prefix ("# " + $dr["name"] + " = ") -TmpProperty $TmpProperty -NbTab $NbTab}
                                }
                                If ($dr["name"]-eq "ScheduleTemplateID") {
                                    #This is a Check Schedule activity, let's provide more details about the schedule itself
                                    $SqlQuery = "select Name from OBJECTS where uniqueID='{" + $TmpProperty + "}'"
                                    $myCommand2 = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection2
                                    $dr2 = $myCommand2.ExecuteReader()
                                    while ($dr2.Read()) {WriteToFile -ExportMode $ExportMode -Add ("# Schedule name : " + $dr2["Name"]) -NbTab $NbTab}
                                    $dr2.Close()
                                    $SqlQuery = "select * from SCHEDULES where uniqueID = '{" + $TmpProperty + "}'"
                                    $myCommand2 = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection2
                                    $dr2 = $myCommand2.ExecuteReader()
                                    while ($dr2.Read())
                                        {
                                        WriteToFile -ExportMode $ExportMode -Add ("# Schedule details : Days of week = " + $dr2["DaysOfWeek"] + " - Days of Month = "+ $dr2["DaysOfMonth"]) -NbTab $NbTab
                                        WriteToFile -ExportMode $ExportMode -Add ("# Schedule details : Monday = " + $dr2["Monday"] + " - Tuesday = "+ $dr2["Tuesday"] + " - Wednesday = "+ $dr2["Wednesday"] + " - Thursday = "+ $dr2["Thursday"] + " - Friday = "+ $dr2["Friday"] + " - Saturday = "+ $dr2["Saturday"] + " - Sunday = "+ $dr2["Sunday"]) -NbTab $NbTab
                                        WriteToFile -ExportMode $ExportMode -Add ("# Schedule details : First = " + $dr2["First"] + " - Second = "+ $dr2["Second"] + " - Third = "+ $dr2["Third"] + " - Fourth = "+ $dr2["Fourth"] + " - Last = "+ $dr2["Fourth"]) -NbTab $NbTab
                                        WriteToFile -ExportMode $ExportMode -Add ("# Schedule details : Days = " + $dr2["Days"] + " - Hours = "+ $dr2["Hours"] + " - Exceptions = "+ $dr2["Exceptions"]) -NbTab $NbTab
                                        }
                                    $dr2.Close() 
                                    $Global:FlagSchedule = $True                                  
                                    If ($Global:FlagScheduleList.Contains($ActivityName))
                                        {$Global:FlagScheduleNumber[$Global:FlagScheduleList.IndexOf($ActivityName)] = $Global:FlagScheduleNumber[$Global:FlagScheduleList.IndexOf($ActivityName)] +1}
                                    else
                                        {
                                        $Global:FlagScheduleList+= $ActivityName
                                        $Global:FlagScheduleNumber+= 1
                                        }
                                    }
                                If ($dr["name"]-eq "CounterID") {
                                    #This is a counter activity, let's provide more information on the counter name
                                    $SqlQuery = "select Name from OBJECTS where uniqueID='{" + $TmpProperty + "}'"
                                    $myCommand2 = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection2
                                    $dr2 = $myCommand2.ExecuteReader()
                                    while ($dr2.Read()) {WriteToFile -ExportMode $ExportMode -Add ("# Counter name : " + $dr2["Name"]) -NbTab $NbTab}
                                    $dr2.Close() 
                                    $SqlQuery = "select DefaultValue from COUNTERS where uniqueID = '{" + $TmpProperty + "}'"
                                    $myCommand2 = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection2
                                    $dr2 = $myCommand2.ExecuteReader()
                                    while ($dr2.Read()) {WriteToFile -ExportMode $ExportMode -Add ("# Counter default value : " + $dr2["DefaultValue"]) -NbTab $NbTab}
                                    $dr2.Close() 
                                    $Global:FlagCounter = $True                                  
                                    If ($Global:FlagCounterList.Contains($ActivityName))
                                        {$Global:FlagCounterNumber[$Global:FlagCounterList.IndexOf($ActivityName)] = $Global:FlagCounterNumber[$Global:FlagCounterList.IndexOf($ActivityName)] +1}
                                    else
                                        {
                                        $Global:FlagCounterList+= $ActivityName
                                        $Global:FlagCounterNumber+= 1
                                        }
                                    }
                                If ($dr["name"]-eq "SelectedBranch") {
                                    #This is a junction activity, let's stor the activity name to mention in the footer summary
                                    $Global:FlagJunction = $True                                  
                                    If ($Global:FlagJunctionList.Contains($ActivityName) -eq $False){
                                        $Global:FlagJunctionList+= $ActivityName
                                        $Global:FlagJunctionNumber+= 1
                                        }
                                    }
                                }
                            }
                        $dr.Close()
                        If ($ActivityType -eq "9c1bf9b4-515a-4fd2-a753-87d235d8ba1f"){
                                    #This is an invole runbook activity, we also populate the calling parameters for the invoked runbook
                                    $SqlQuery = "Select * from TRIGGER_POLICY_PARAMETERS where parentID = '" + $ActivityID + "'"
                                    $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
                                    $dr = $myCommand.ExecuteReader()
                                    while ($dr.Read())
                                            {
                                            $SqlQuery = "select * from CUSTOM_START_PARAMETERS where uniqueID = '" + $dr["parameter"] + "'"
                                            $myCommand2 = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection2
                                            $dr2 = $myCommand2.ExecuteReader()
                                            while ($dr2.Read()) {$ParamName = $dr2["value"]}
                                            $dr2.Close()
                                            If ([System.DBNull]::Value.Equals($dr["value"]) -eq $False)
                                             {ParseAndWriteProperty -ExportMode $ExportMode -Prefix ("# Input parameter : " + $ParamName + " = ") -TmpProperty $dr["value"] -NbTab $NbTab}
                                             else {WriteToFile -ExportMode $ExportMode -Add "# Input parameter : $ParamName = < no value was passed >" -NbTab $NbTab}
                                            }
                                    $dr.Close()
                                    $Global:FlagInvokeRunbook = $True                                  
                                    If ($Global:FlagInvokeRunbookList.Contains($ActivityName))
                                        {$Global:FlagInvokeRunbookNumber[$Global:FlagInvokeRunbookList.IndexOf($ActivityName)] = $Global:FlagInvokeRunbookNumber[$Global:FlagInvokeRunbookList.IndexOf($ActivityName)] +1}
                                    else
                                        {
                                        $Global:FlagInvokeRunbookList+= $ActivityName
                                        $Global:FlagInvokeRunbookNumber+= 1
                                        }
                        }

                }
                WriteToFile -ExportMode $ExportMode -Add "# END ACTIVITY - $ActivityDetailsShort" -NbTab $NbTab
                }
            }
        }
    $myConnection2.Close()
    WriteToFile -ExportMode $ExportMode -Add "" -NbTab $NbTab
  
}


function ParseRunbookFromActivity()
# This is the function we recurse on when outputting the structure of the PS1 file,
# as we go through the source Runbook in Orchestrator
# GeneratePS1() calls this function the first time, from the starting activity,
# and then it recurses from there
{

    param (
    [String]$ActivityID,
    [String]$ActivityName,
    [String]$ActivityType,
    [String]$LinkID,
    [Boolean]$InParallel,
    [Int]$CurrentTabNumber
    )

        $LinkedActivitiesID = @()
        $LinkedActivitiesNames = @()
        $LinkedActivitiesType = @()
        $Links = @()

        $NewTabNumber = $CurrentTabNumber

        $SqlQuery = "select LINKS.TargetObject, Links.UniqueID As LID, OBJECTS.ObjectType, OBJECTS.Name, OBJECTS.UniqueID As OBJID from LINKS, OBJECTS where NOT EXISTS (SELECT UniqueID FROM OBJECTS WHERE UniqueID=LINKS.UniqueID AND DELETED=1) AND LINKS.SourceObject = '" + $ActivityID + "' AND OBJECTS.UniqueID = LINKS.TargetObject AND OBJECTS.Deleted=0"
        $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
        $dr = $myCommand.ExecuteReader()
        while ($dr.Read())
            {
            $LinkedActivitiesID +=$dr["OBJID"]
            $LinkedActivitiesNames +=$dr["Name"]
            $LinkedActivitiesType +=$dr["ObjectType"]
            $Links +=$dr["LID"]
            }
        $dr.Close()
        
        #Let's check if this is the first activity in the runbook
        If ($LinkID -eq "SOURCE")
            {$LinkDetails = ""}
            else {$LinkDetails = LinkCondition($LinkID)}

       #The following is handled differently depending on the number of linked activities 
       switch ($LinkedActivitiesID.Count){
       0
            {
                If ($LinkDetails -ne ""){
                    #Condition was found on link
                    WriteToFile -ExportMode "PS1" -Add $LinkDetails -NbTab $NewTabNumber
                    $NewTabNumber = $NewTabNumber + 1
                }
                AppendActivityDetails -ActivityID $ActivityID -ActivityDetailsShort $ActivityName -ActivityType $ActivityType -NbTab $NewTabNumber -ExportMode "PS1"
                If ($LinkDetails -ne "") {
                    #Condition was found on link
                    WriteToFile -ExportMode "PS1" -Add "}" -NbTab $NewTabNumber
                    $NewTabNumber = $NewTabNumber - 1
                }
            }
        1
            {
                If (($InParallel -eq $True) -and ($LinkDetails -eq "")) {
                    WriteToFile -ExportMode "PS1" -Add "Sequence {" -NbTab $NewTabNumber
                    $NewTabNumber = $NewTabNumber + 1
                }
                If ($LinkDetails -ne "") {
                    #Condition was found on link
                    WriteToFile -ExportMode "PS1" -Add $LinkDetails -NbTab $NewTabNumber
                    $NewTabNumber = $NewTabNumber + 1
                }
                AppendActivityDetails -ActivityID $ActivityID -ActivityDetailsShort $ActivityName -ActivityType $ActivityType -NbTab $NewTabNumber -ExportMode "PS1"
                ParseRunbookFromActivity -ActivityID $LinkedActivitiesID[0] -ActivityName $LinkedActivitiesNames[0] -ActivityType $LinkedActivitiesType[0] -LinkID $Links[0] -InParallel $False -CurrentTabNumber $NewTabNumber
                If ($LinkDetails -ne "") {
                    #Condition was found on link
                    WriteToFile -ExportMode "PS1" -Add "}" -NbTab $NewTabNumber
                    $NewTabNumber = $NewTabNumber - 1
                }
                If ($InParallel -eq $True -and $LinkDetails -eq "") {
                    WriteToFile -ExportMode "PS1" -Add "}" -NbTab $NewTabNumber
                    $NewTabNumber = $NewTabNumber - 1
                }
            }
        {$_ -gt 1}
            {
                $Global:FlagParallel = $True
                If ($Global:FlagParallelList.Contains($ActivityName) -eq $False) {$Global:FlagParallelList+= $ActivityName}
                If ($LinkDetails -ne ""){
                    #Condition was found on link
                    WriteToFile -ExportMode "PS1" -Add $LinkDetails -NbTab $NewTabNumber
                    $NewTabNumber = $NewTabNumber + 1
                }
                AppendActivityDetails -ActivityID $ActivityID -ActivityDetailsShort $ActivityName -ActivityType $ActivityType -NbTab $NewTabNumber -ExportMode "PS1"
                WriteToFile -ExportMode "PS1" -Add "Parallel {" -NbTab $NewTabNumber
                $NewTabNumber = $NewTabNumber + 1
                $i = 0
                While ($i -lt $LinkedActivitiesID.Count){
                    ParseRunbookFromActivity -ActivityID $LinkedActivitiesID[$i] -ActivityName $LinkedActivitiesNames[$i] -ActivityType $LinkedActivitiesType[$i] -LinkID $Links[$i] -InParallel $True -CurrentTabNumber $NewTabNumber
                    $i = $i + 1
                }
                WriteToFile -ExportMode "PS1" -Add "}" -NbTab $NewTabNumber
                If ($LinkDetails -ne ""){
                    #Condition was found on link
                    WriteToFile -ExportMode "PS1" -Add "}" -NbTab $NewTabNumber
                    $NewTabNumber = $NewTabNumber - 1
                }
            }
        }
}



function GeneratePS1()
# This function is being called when clicking on the 'Export' button in the GUI
# It receives parameters from the TreeView
# After writing some introduction data in the PS1 file, it recurses in the source Runbook
# to find activities - leveraging AppendActivityDetails() - and finishes with
# a summary/analysis and some suggested next steps in the footer of the PS1 file
{

param (
    [String]$RunbookID,
    [String]$RunbookName,
    [String]$ExportFileName
)
        
        $myConnection = New-Object System.Data.SqlClient.SqlConnection $Global:SQLConnstr 
        $myConnection.Open()

        $Global:FlagParallel = $False
        $Global:FlagParallelList = @()
        $Global:FlagJunction = $False
        $Global:FlagJunctionList = @()
        $Global:FlagStringcondition = $False
        $Global:FlagInitializeData = $False
        $Global:FlagVariables = $False
        $Global:FlagVariablesList = @()
        $Global:FlagVariablesNumber = @()
        $Global:FlagVariablesValue =@()
        $Global:FlagInvokeRunbook = $False
        $Global:FlagInvokeRunbookList = @()
        $Global:FlagInvokeRunbookNumber = @()
        $Global:FlagReturnData = $False
        $Global:FlagReturnDataList = @()
        $Global:FlagReturnDataNumber = @()
        $Global:FlagSchedule = $False
        $Global:FlagScheduleList = @()
        $Global:FlagScheduleNumber = @()
        $Global:FlagCounter = $False
        $Global:FlagCounterList = @()
        $Global:FlagCounterNumber = @()

        $ParametersNames=@()
        $ParametersTypes=@()        

        If ($SkeletonCB.IsChecked -eq $True){$OutputPS1Name = "Invoke-" + $ExportFileName + "_Skeleton.ps1"}else{$OutputPS1Name = "Invoke-" + $ExportFileName + ".ps1"}
                
                WriteToFile -ExportMode "PS1" -Add "#################################################################################"
        WriteToFile -ExportMode "PS1" -Add "# WORKFLOW CREATED BY THE SMA RUNBOOK CONVERSION HELPER"
        If ($SkeletonCB.IsChecked -eq $True){
            WriteToFile -ExportMode "PS1" -Add "# This was created in 'Skeleton Mode', to only display"
            WriteToFile -ExportMode "PS1" -Add "# the overall structure of the source Orchestrator Runbook"
            }
            else
            {
            WriteToFile -ExportMode "PS1" -Add "# Make sure you review the summary analysis at the end of this file"
            WriteToFile -ExportMode "PS1" -Add "# to make final adjustements to the runbook"
        }

        WriteToFile -ExportMode "PS1" -Add "#################################################################################"


        WriteToFile -ExportMode "PS1" -Add "workflow Invoke-$ExportFileName {"
        WriteToFile -ExportMode "PS1" -Add ""

        $LinkID = ""

        #Find Starting Objects (there should be only one starting object, although we could find multiple links fronmm that starting object)
        $SqlQuery = "select LINKS.SourceObject, LINKS.UniqueID as LID, OBJECTS.ObjectType, OBJECTS.Name, Objects.UniqueID AS OBJID from LINKS, OBJECTS where SourceObject NOT IN (select TargetObject from LINKS) AND LINKS.SourceObject=OBJECTS.UniqueID AND OBJECTS.ParentID = '" + $RunbookID + "' AND OBJECTS.Deleted=0"
        $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
        $dr = $myCommand.ExecuteReader()
        while ($dr.Read())
            {
            $SourceID = $dr["OBJID"]
            $SourceName = $dr["Name"]
            #$LinkID = $dr["LID"]
            $LinkID = "SOURCE"
            $SourceType = $dr["ObjectType"]
            }
        $dr.Close()
        If ($LinkID -eq ""){
        #This might be a Runbook with only one single activity, let's try to find that activity
        $SqlQuery = "select UniqueID, Name, ObjectType from objects where ParentID = '" + $RunbookID + "' AND OBJECTS.Deleted=0"
        $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
        $dr = $myCommand.ExecuteReader()
        while ($dr.Read())
            {
            $SourceID = $dr["UniqueID"]
            $SourceName = $dr["Name"]
            $LinkID = "SOURCE"
            $SourceType = $dr["ObjectType"]
            }
        $dr.Close()
        }
        #Check if we are working with a an Initialize Data activity, in which case we fill out parameters
        If ($SourceType -eq "6C576F3D-E927-417A-B145-5D3EFF9C995F") {
            $SqlQuery = "select value, type from CUSTOM_START_PARAMETERS where ParentID = '" + $SourceID + "'"
            $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection 
            $dr = $myCommand.ExecuteReader()
            while ($dr.Read())
                {
                $ParametersNames+=$dr["value"]
                $ParametersTypes+=$dr["type"]
                }
            $dr.Close()
            If ($ParametersNames.Length -ne 0){
                WriteToFile -ExportMode "PS1" -Add "      ("             
                For ($i=0; $i -lt ($ParametersNames.Length-1); $i++){
                        WriteToFile -ExportMode "PS1" -Add "      [parameter(Mandatory=`$true)]"                
                        WriteToFile -ExportMode "PS1" -Add ("      [" + $ParametersTypes[$i] + "]$" + $ParametersNames[$i].Replace(" ", "_") + ",")
                }
                WriteToFile -ExportMode "PS1" -Add "      [parameter(Mandatory=`$true)]"                
                WriteToFile -ExportMode "PS1" -Add ("      [" + $ParametersTypes[$ParametersNames.Length-1] + "]$" + $ParametersNames[$ParametersNames.Length-1].Replace(" ", "_"))
            }
                        
            WriteToFile -ExportMode "PS1" -Add "         )"

            $Global:FlagInitializeData = $True
            WriteToFile -ExportMode "PS1" -Add ""
        }
        ParseRunbookFromActivity -ActivityID $SourceID -ActivityName $SourceName -ActivityType $SourceType -LinkID $LinkID -InParallel $False -CurrentTabNumber 1

        WriteToFile -ExportMode "PS1" -Add ""
        WriteToFile -ExportMode "PS1" -Add "}"

        If ($SkeletonCB.IsChecked -eq $False){
           WriteToFile -ExportMode "PS1" -Add "########################### CONVERSION HELPER SUMMARY ###########################"
            If ($Global:FlagInitializeData -eq $True) {
                 WriteToFile -ExportMode "PS1" -Add "# RUNBOOK START : Starting activity in the source Runbook is an Initialize Data activity. Actual parameters have been added at the beginning of the resulting workflow. Spaces in the parameters' names have been replaced by underscore characters."
                 }
            Else{WriteToFile -ExportMode "PS1" -Add "# RUNBOOK START : No Initialize Data activity was used in this source Runbook. It might be using a monitor activity, which might imply using schedules for examples."}
            If ($Global:FlagVariables -eq $True) {
                $AppendString = "# VARIABLES : Variables were found in the source Runbook and should be replaced by SMA variables. Here is the list of variables discovered, and how many times they were used in the source Runbook :"
                ForEach ($ArrayString In $Global:FlagVariablesList){$AppendString = $AppendString + "`r`n# - " + $ArrayString + " (x" + $Global:FlagVariablesNumber[$Global:FlagVariablesList.IndexOf($ArrayString)] + ") - Current Value = " + $Global:FlagVariablesValue[$Global:FlagVariablesList.IndexOf($ArrayString)]}
                WriteToFile -ExportMode "PS1" -Add $AppendString
                }
            Else{WriteToFile -ExportMode "PS1" -Add "# VARIABLES : No variables were used in this source Runbook."}
            If ($Global:FlagInvokeRunbook-eq $True) {
                $AppendString = "# SUBROUTINES : Subroutines are involved through Invoke Runbooks activities in the source Orchestrator Runbook. You would likely replace these subroutines by SMA subroutine Runbooks. Here is the list of invoked Runbooks discovered, and how many times they were called by the source Runbook :"
                ForEach ($ArrayString In $Global:FlagInvokeRunbookList){$AppendString = $AppendString + "`r`n# - " + $ArrayString + " (x" + $Global:FlagInvokeRunbookNumber[$Global:FlagInvokeRunbookList.IndexOf($ArrayString)] + ")"}
                WriteToFile -ExportMode "PS1" -Add $AppendString
                }
            Else{WriteToFile -ExportMode "PS1" -Add "# SUBROUTINES : No child/subroutine Runbooks were used in this source Runbook."}
            If ($Global:FlagReturnData -eq $True) {
                $AppendString = "# OUTUT DATA : Return Data activities are used in the source Orchestrator Runbook. Properties returned are embedded earlier in the script. Here is a list of names for the corresponding activities, and how many times they were called by the source Runbook :"
                ForEach ($ArrayString In $Global:FlagReturnDataList){$AppendString = $AppendString + "`r`n# - " + $ArrayString + " (x" + $Global:FlagReturnDataNumber[$Global:FlagReturnDataList.IndexOf($ArrayString)] + ")"}
                WriteToFile -ExportMode "PS1" -Add $AppendString
                }
            Else{WriteToFile -ExportMode "PS1" -Add "# OUTPUT DATA : No Return Data activities were used in this source Runbook."}
            If ($Global:FlagSchedule -eq $True) {
                $AppendString = "# SCHEDULES : Check Schedules were found in the source Runbook and documented earlier in the script. Here is a list of names for the corresponding activities, and how many times they were called by the source Runbook :"
                ForEach ($ArrayString In $Global:FlagScheduleList){$AppendString = $AppendString + "`r`n# - " + $ArrayString + " (x" + $Global:FlagScheduleNumber[$Global:FlagScheduleList.IndexOf($ArrayString)] + ")"}
                WriteToFile -ExportMode "PS1" -Add $AppendString
                }
            Else{WriteToFile -ExportMode "PS1" -Add "# SCHEDULES : No Check Schedule activities were used in this source Runbook."}
            If ($Global:FlagCounter -eq $True) {
                $AppendString = "# COUNTERS : Counter-related activities were found in the source Runbook and documented earlier in the script. Here is a list of names for the corresponding activities, and how many times they were called by the source Runbook :"
                ForEach ($ArrayString In $Global:FlagCounterList){$AppendString = $AppendString + "`r`n# - " + $ArrayString + " (x" + $Global:FlagCounterNumber[$Global:FlagCounterList.IndexOf($ArrayString)] + ")"}
                WriteToFile -ExportMode "PS1" -Add $AppendString
                }
            Else{WriteToFile -ExportMode "PS1" -Add "# COUNTERS : No Counter-related activities were used in this source Runbook."}
            If ($Global:FlagParallel -eq $True) {
                $AppendString = "# PARALLEL BRANCHES : The following activities had multiple outbound links. These have been added to the script as 'parallel' script paragraphs, but you might also want to look if the use of 'parallel' is actually needed or could be optimized."
                ForEach ($ArrayString In $Global:FlagParallelList){$AppendString = $AppendString + "`r`n# - " + $ArrayString}
                WriteToFile -ExportMode "PS1" -Add $AppendString
                }
            Else{WriteToFile -ExportMode "PS1" -Add "# PARALLEL BRANCHES : No activities has multiple outbound links."}
            If ($Global:FlagStringCondition -eq $True) {
                WriteToFile -ExportMode "PS1" -Add "# LINK CONDITIONS : Some links were including string conditions, that you may need to replace with StartsWith, CompareTo, etc. : http://technet.microsoft.com/en-us/library/ee692804.aspx"
                }
            Else{WriteToFile -ExportMode "PS1" -Add "# LINK CONDITIONS : No string conditions were used in this source Runbook's links."}
            ##Check for loops
            WriteToFile -ExportMode "PS1" -Add ("# LOOPS : If there are loops within the source Runbooks, they are listed below, and you may also need to look at the exit/non-exit conditions to convert them:")
            $SqlQuery = "select Objectlooping.UniqueID As OUID, DelaybetweenAttempts, Name from objectlooping, objects where objects.uniqueid = objectlooping.uniqueID and objectlooping.enabled = 1 AND objects.Deleted=0 AND objects.parentID = '" + $RunbookID +  "'"
            $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
            $dr = $myCommand.ExecuteReader()
            while ($dr.Read())
                {
                WriteToFile -ExportMode "PS1" -Add ("# - Activity named '" + $dr["Name"] + "' is looping every " + $dr["delaybetweenAttempts"] + " seconds.")
                }
            $dr.Close()
            ##Check for merging branches - these need to be addressed to simplify the output script (right now the output duplicates the branches)
            WriteToFile -ExportMode "PS1" -Add ("# MERGING BRANCHES : If there are activities with multiple inbound branches, they are listed below and the script output may be simplified (right now it duplicates code for each branch, in this situation)")
            $SqlQuery = "select UniqueID, Name from OBJECTS where ParentID='" + $RunbookID + "' AND DELETED=0 AND UniqueID IN (Select TargetObject from LINKS WHERE NOT EXISTS (SELECT UniqueID FROM OBJECTS WHERE OBJECTS.UniqueID=LINKS.UniqueID AND DELETED=1) group by TargetObject having count(TargetObject) > 1)"
            $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
            $dr = $myCommand.ExecuteReader()
            while ($dr.Read())
                {
                WriteToFile -ExportMode "PS1" -Add ("# - Activity named '" + $dr["Name"] + "' has more than 1 inbound links.")
                }
            $dr.Close()
            If ($Global:FlagJunction -eq $True) {
                $AppendString = "# JUNCTIONS : Junctions are used in the source Orchestrator Runbook. Here is the list of junction activities :"
                ForEach ($ArrayString In $Global:FlagJunctionList){$AppendString = $AppendString + "`r`n# - " + $ArrayString}
                WriteToFile -ExportMode "PS1" -Add $AppendString
                }
            Else{WriteToFile -ExportMode "PS1" -Add "# JUNCTIONS : No junctions were used in this source Runbook."}
        }

        $myConnection.Close()
        
        WriteToFile -ExportMode "PS1" -Add "#################################################################################"
        write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Exporting to PS1 file $OutputPS1Name"

        If ($Global:SaveAndClosePS1.IsChecked -eq $False)
            {start-process -FilePath ($Global:ISELocation + "PowerShell_ISE.exe") -ArgumentList $OutputPS1Name}

}

function GenerateVSD()

{
param (
    [String]$RunbookID,
    [String]$RunbookName,
    [String]$ExportFileName
)

        $myConnection = New-Object System.Data.SqlClient.SqlConnection $Global:SQLConnstr 
        $myConnection.Open()
        $myConnection2 = New-Object System.Data.SqlClient.SqlConnection $Global:SQLConnstr 
        $myConnection2.Open()

        $ListShapes =@()
        $ListShapes.Clear()
        $ListShapesID =@()
        $ListShapesID.Clear()

        write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Creating header"
        $vApp = New-Object -ComObject Visio.Application
        $vDoc = $vApp.Documents.Add("")
        $vStencil = $vApp.Documents.OpenEx($VisioTemplate, 4)
        $SpecificvStencil = $vStencil.Masters | where-object {$_.Name -eq "Process"}
        
        #Size the page dynamically
        $vApp.ActivePage.AutoSize = $true

        #Add a title
        $vToShape = $vApp.ActivePage.Drop($SpecificvStencil, 4, 10)
        #$vStencil.Masters.Name
        $vToShape.Text = "Runbook : $RunbookName"
        $vToShape.Cells("Char.Size").Formula = "= 30 pt."
        $vToShape.Cells("Width").Formula = "= 7"

        #Find and draw activities
        write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Drawing activities"
        $vFlowChartMaster = $vStencil.Masters | where-object {$_.Name -eq $Global:VisioStencil}
        
        $SqlQuery = "select UniqueID, ObjectType, Name, Description, PositionX, PositionY from OBJECTS WHERE ParentID = '" + $RunbookID + "' AND ObjectType <> '7A65BD17-9532-4D07-A6DA-E0F89FA0203E' AND Deleted=0"
        $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
        $dr = $myCommand.ExecuteReader()
        while ($dr.Read())
            {
            write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Working with activity :" $dr["Name"]
            $vApp.ActiveWindow.DeselectAll()
            $vToShape = $vApp.ActivePage.Drop($vFlowChartMaster, ($dr["PositionX"] / 50), (-($dr["PositionY"] / 50) + 7))
            $vToShape.Text = $dr["Name"]
            If ([System.DBNull]::Value.Equals($dr["description"]) -eq $False)
                {
                $vsoDoc1 = $vApp.Documents.OpenEx($vApp.GetBuiltInStencilFile(3, 2), 64)
                $vCallout = $vApp.ActivePage.DropCallout($vsoDoc1.Masters.ItemU($VisioCallout), $vToShape)
                $vCallout.Text = $dr["Description"]
                $vsoDoc1.Close()
                }
            $vToShape.Cells("Para.HorzAlign").Formula = "=2"
            $vToShape.Cells("LeftMargin").Formula = "=0.5"
            #get the icon for this activity
            If (Test-Path ("{" + $dr["ObjectType"] + "}.jpg"))
                {$shp1Obj = $vApp.ActivePage.Import((Get-Location -PSProvider FileSystem).ProviderPath + "\{" + $dr["ObjectType"] + "}.jpg")}
                else {$shp1Obj = $vApp.ActivePage.Import((Get-Location -PSProvider FileSystem).ProviderPath + "\default.jpg")}
            $shp2Obj = $vApp.ActivePage.Drop($shp1Obj, ($dr["PositionX"] / 50) - 0.25, -($dr["PositionY"] / 50) + 7)
            $shp1Obj.Delete() #Remove original imported reference
            $vApp.ActiveWindow.Select($vToShape, 2)
            $vApp.ActiveWindow.Select($shp2Obj, 2)
            $vSel = $vApp.ActiveWindow.Selection
            If ($FORM.FindName('GroupThumbnails').IsChecked -eq $True)
                {$vSel.Group()}
            $ListShapes += $vToShape
            $ListShapesID += $dr["UniqueID"]
            $vSel.DeselectAll()
            }
        $dr.Close()


        #Find and draw links
        write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Drawing links"
        $vConnectorMaster = $vStencil.Masters | where-object {$_.Name -eq "Dynamic Connector"}
        $SqlQuery = "Select DISTINCT LINKS.UniqueID As LID, name, deleted, objecttype, LINKS.Color, LINKS.sourceobject, LINKS.targetobject from OBJECTS, LINKS where (ObjectType='7A65BD17-9532-4D07-A6DA-E0F89FA0203E' AND ParentID='" + $RunbookID + "' AND OBJECTS.Deleted=0 AND LINKS.UniqueID=OBJECTs.UniqueID)"
        $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
        $dr = $myCommand.ExecuteReader()
        while ($dr.Read())
            {
            $vConnector = $vApp.ActivePage.Drop($vConnectorMaster, 0, 0)
            $vConnector.Cells("EndArrow").Formula = "=4"
            #write-host $dr["Color"]
            $TmpColor=[Long]$dr["Color"]
            $TmpRed = $TmpColor % 256
            $TmpColor = $TmpColor / 256
            $TmpGreen = $TmpColor % 256
            $TmpColor = $TmpColor / 256
            $TmpBlue = $TmpColor % 256
            Try {$vConnector.Cells("LineColor").Formula = "RGB($TmpRed,$TmpGreen,$TmpBlue)"}
            Catch {$vConnector.Cells("LineColor").Formula = "RGB($TmpRed;$TmpGreen;$TmpBlue)"}
            $vBeginCell = $vConnector.Cells("BeginX")
            $vFromShape = $ListShapes.Item($ListShapesID.IndexOf($dr["SourceObject"]))
            $vBeginCell.GlueTo($vFromShape.Cells("Align" + $Global:VisioGlueFrom))
            $vEndCell = $vConnector.Cells("EndX")
            $vToShape = $ListShapes.Item($ListShapesID.IndexOf($dr["TargetObject"]))
            $vEndCell.GlueTo($vToShape.Cells("Align" + $Global:VisioGlueTo))
            #LID to String?
            $SqlQuery2 = "select DISTINCT Name, LINKS.UniqueID from LINKS, OBJECTS WHERE OBJECTS.UniqueID = LINKS.UniqueID AND LINKS.UniqueID = '" + $dr["LID"] + "'"
            $myCommand2 = New-Object System.Data.SqlClient.sqlCommand $SqlQuery2, $myConnection2
            $dr2 = $myCommand2.ExecuteReader()
            while ($dr2.Read())
                {
                If ($dr2["name"] -ne "Link") {$vConnector.Text = $dr2["Name"]}
                }
            $dr2.Close()
            #$myCommand2 = Nothing
            $vConnector.SendToBack()
            }
        $dr.Close()

        #Find and draw loops
        write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Drawing loops"
        $SqlQuery = "select Objectlooping.UniqueID As OUID, DelaybetweenAttempts, Name from objectlooping, objects where objects.uniqueid = objectlooping.uniqueID and objectlooping.enabled = 1 AND OBJECTS.Deleted=0 AND objects.parentID = '" + $RunbookID + "'"
        $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
        $dr = $myCommand.ExecuteReader()
        while ($dr.Read())
            {
            $vConnector = $vApp.ActivePage.Drop($vConnectorMaster, 0, 0)
            $vConnector.Cells("EndArrow").Formula = "=4"
            $vBeginCell = $vConnector.Cells("BeginX")
            $vFromShape = $ListShapes.Item($ListShapesID.IndexOf($dr["OUID"]))
            $vBeginCell.GlueTo($vFromShape.Cells("AlignRight"))
            $vEndCell = $vConnector.Cells("EndX")
            $vToShape = $ListShapes.Item($ListShapesID.IndexOf($dr["OUID"]))
            $vEndCell.GlueTo($vToShape.Cells("AlignTop"))
            Try {If ($dr["DelayBetweenAttempts"] -ne "") {$vConnector.Text = "Loop every " + $dr["DelayBetweenAttempts"] + " seconds"}}
            Catch {$vConnector.Text = "Loop (undefined interval)"}
            $vConnector.SendToBack()
            }
        $dr.Close()

        $vApp.ActivePage.AutoSizeDrawing()

        $myConnection.Close()
        $myConnection2.Close()


    If ($FORM.FindName('SaveAndCloseVSD').IsChecked -eq $True)
        {        
        write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Exporting to Visio file $ExportFileName.VSDX"
        $vDoc.SaveAs((Get-Location -PSProvider FileSystem).ProviderPath + "\" + $ExportFileName + ".VSDX")
        $vDoc.Close()
        $vApp.Quit()
        }

}

function GenerateDOC()

{
param (
    [String]$RunbookID,
    [String]$RunbookName,
    [String]$ExportFileName
)

        $Global:FlagParallel = $False
        $Global:FlagParallelList = @()
        $Global:FlagJunction = $False
        $Global:FlagJunctionList = @()
        $Global:FlagStringcondition = $False
        $Global:FlagInitializeData = $False
        $Global:FlagVariables = $False
        $Global:FlagVariablesList = @()
        $Global:FlagVariablesNumber = @()
        $Global:FlagVariablesValue =@()
        $Global:FlagInvokeRunbook = $False
        $Global:FlagInvokeRunbookList = @()
        $Global:FlagInvokeRunbookNumber = @()
        $Global:FlagReturnData = $False
        $Global:FlagReturnDataList = @()
        $Global:FlagReturnDataNumber = @()
        $Global:FlagSchedule = $False
        $Global:FlagScheduleList = @()
        $Global:FlagScheduleNumber = @()
        $Global:FlagCounter = $False
        $Global:FlagCounterList = @()
        $Global:FlagCounterNumber = @()

        $myConnectionDOC = New-Object System.Data.SqlClient.SqlConnection $Global:SQLConnstr 
        $myConnectionDOC.Open()
        $myConnection = New-Object System.Data.SqlClient.SqlConnection $Global:SQLConnstr 
        $myConnection.Open()
        $myConnection2 = New-Object System.Data.SqlClient.SqlConnection $Global:SQLConnstr 
        $myConnection2.Open()
        $myConnection3 = New-Object System.Data.SqlClient.SqlConnection $Global:SQLConnstr 
        $myConnection3.Open()

        #Start Word and open the document template.
        $oWord = New-Object -ComObject Word.Application
        $oWord.Visible = $True
        $oDoc = $oWord.Documents.Add()
        #Insert a paragraph at the beginning of the document.
        $oPara1 = $oDoc.Content.Paragraphs.Add()
        $oPara1.Range.Text = "Runbook : " + $RunbookName
        $oPara1.Range.Font.Bold = $True
        $oPara1.Range.Font.Size = 28
        $oPara1.Format.SpaceAfter = 24    #24 pt spacing after paragraph.
        $oPara1.Range.InsertParagraphAfter()
        #Insert a 3 x 5 table, fill it with data, and make the first row bold and italic.
        $oTable = $oDoc.Tables.Add($oDoc.Bookmarks.Item("\endofdoc").Range, 1, 4)
        $oTable.Range.ParagraphFormat.SpaceAfter = 6
        $oTable.Range.Font.Size = 8
        $oTable.Range.Font.Bold = $True
        $oTable.Range.Borders.Enable = $True
        $oTable.Range.Borders.OutsideLineStyle = 7
        $oTable.Range.Borders.InsideLineStyle = 0
        $oTable.Cell(1, 1).Range.Text = "Activity"
        $oTable.Cell(1, 2).Range.Text = "Description"
        $oTable.Cell(1, 3).Range.Text = "Details"
        $oTable.Cell(1, 4).Range.Text = "Published data dependencies"
        $oTable.Columns.Item(3).Width = $oWord.InchesToPoints(3)
        $oTable.Columns.Item(2).Width = $oWord.InchesToPoints(1)
        $oTable.Columns.Item(4).Width = $oWord.InchesToPoints(1)

        $r = 2

        $SqlQueryDOC = "select UniqueID, ObjectType, Name, Description, PositionX, PositionY from OBJECTS WHERE ParentID = '" + $RunbookID + "' AND ObjectType <> '7A65BD17-9532-4D07-A6DA-E0F89FA0203E' AND Deleted=0"
        $myCommandDOC = New-Object System.Data.SqlClient.sqlCommand $SqlQueryDOC, $myConnectionDOC
        $drDOC = $myCommandDOC.ExecuteReader()
        while ($drDOC.Read())
            {
            $oTable.Rows.Add()
            $oTable.Rows.Item($r).Range.Font.Bold = $False
            $oTable.Rows.Item($r).Range.Borders.Enable = $True
            $oTable.Rows.Item($r).Range.Borders.OutsideLineStyle = 1
            $oTable.Rows.Item($r).Range.Borders.InsideLineStyle = 0
            $oPara1 = $oTable.Cell($r, 1).Range.Paragraphs.Add()
            $oPara1.Range.Text = $drDOC["Name"]
            $oPara1 = $oTable.Cell($r, 1).Range.Paragraphs.Add()
            If (Test-Path ("{" + $drDOC["ObjectType"] + "}.jpg"))
                {$oPara1.Range.InlineShapes.AddPicture((Get-Location -PSProvider FileSystem).ProviderPath + "\{" + $drDOC["ObjectType"] + "}.jpg")}
                else {$oPara1.Range.InlineShapes.AddPicture((Get-Location -PSProvider FileSystem).ProviderPath + "\default.jpg")}

            If ([System.DBNull]::Value.Equals($drDOC["description"]) -eq $False)
                {
                $oPara1 = oTable.Cell($r, 2).Range.Paragraphs.Add()
                $oPara1.Range.Text = $drDOC["Description"]
                }

            $SqlQuery2 = "select PrimaryDataTable from ObjectTypes, Objects where ObjectTypes.UniqueID = Objects.ObjectType and Objects.UniqueID = '" + $drDOC["UniqueID"] + "'"
            $myCommand2 = New-Object System.Data.SqlClient.sqlCommand $SqlQuery2, $myConnection2
            $dr2 = $myCommand2.ExecuteReader()
            $TmpTableInit = $False
            while ($dr2.Read())
                {
                If ([System.DBNull]::Value.Equals($dr2["PrimaryDataTable"]) -eq $False)
                    {
                    $TmpTable = $dr2["PrimaryDataTable"]
                    $TmpTableInit = $True
                    }
                }
            $dr2.Close()
            $Global:ActivityDependenciesActivityNames = @()
            $Global:ActivityDependenciesActivityTypes = @()
            $Global:ActivityDependenciesVariableNames = @()
            $Global:ActivityDependenciesVariableValues = @()
            AppendActivityDetails -ActivityID ($drDOC["UniqueID"]) -ActivityDetailsShort ($drDOC["Name"]) -ActivityType ($drDOC["ObjectType"]) -ExportMode "DOC"
            foreach ($ActivityDependencyActivityName in $Global:ActivityDependenciesActivityNames)
                {
                $oPara1 = $oTable.Cell($r, 4).Range.Paragraphs.Add()
                $oPara1.Range.Text = $ActivityDependencyActivityName
                $oPara1 = $oTable.Cell($r, 4).Range.Paragraphs.Add()
                If (Test-Path ("{" + $Global:ActivityDependenciesActivityTypes[$Global:ActivityDependenciesActivityNames.IndexOf($ActivityDependencyActivityName)] + "}.jpg"))
                        {$oPara1.Range.InlineShapes.AddPicture((Get-Location -PSProvider FileSystem).ProviderPath + "\{" + $Global:ActivityDependenciesActivityTypes[$Global:ActivityDependenciesActivityNames.IndexOf($ActivityDependencyActivityName)] + "}.jpg")}
                        else {$oPara1.Range.InlineShapes.AddPicture((Get-Location -PSProvider FileSystem).ProviderPath + "\default.jpg")}
                }
            foreach ($ActivityDependenciesVariableName in $Global:ActivityDependenciesVariableNames)
                {
                $oPara1 = $oTable.Cell($r, 4).Range.Paragraphs.Add()
                $oPara1.Range.Text = (" Variable: " + $ActivityDependenciesVariableName)
                $oPara1 = $oTable.Cell($r, 4).Range.Paragraphs.Add()
                $oPara1.Range.Text = (" Value: " + $Global:ActivityDependenciesVariableValues[$Global:ActivityDependenciesVariableNames.IndexOf($ActivityDependenciesVariableName)])
                }
            $r = $r + 1
            }

        $drDOC.Close()
        $myConnectionDOC.Close()
        $myConnection.Close()
        $myConnection2.Close()
        $myConnection3.Close()


    If ($FORM.FindName('SaveAndCloseDOC').IsChecked -eq $True)
        {        
        write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Exporting to Word file $ExportFileName.DOCX"
        $oDoc.SaveAs([ref]((Get-Location -PSProvider FileSystem).ProviderPath + "\" + $ExportFileName + ".DOCX"))
        $oDoc.Close()
        $oWord.Quit()
        }

}


function WriteToFile()
# This function is called throughout the tool, to write to the PS1 file
# For readability purposes in the PS1 file being generated, it includes
# a 'NbTab' parameter, to compute the tabulations as we go deeper
# into the PowerShell branches
{
    param (
    [String]$Add,
    [int]$NbTab,
    [String]$ExportMode
    )

    switch ($ExportMode)
        {
        "PS1"
            {
            $Output = ""
            If ($NbTab)
                { For ($i=1; $i -le $NbTab; $i++){$Output = "    " + $Output} }
            $Output = $Output + $Add
            Add-content $OutputPS1Name -value $Output -Force
            }
        "DOC"
            {
            $oPara1 = $oTable.Cell($r, 3).Range.Paragraphs.Add()
            $oPara1.Range.Text = $Add
            }
        }
}


function Popup()
# This function is currently not used anymore in the code
{
param (
    [String]$Message
)

$a = new-object -comobject wscript.shell
$b = $a.popup($Message,0,"SMA Runbook Conversion Helper",0)


}

function Popup2()
# This is only used a few times, to display explicit and important popups
# instead of writing to the console.
# For example when the connection to the database cannot be opened.
{

param (
    [String]$Message,
    [Boolean]$ClosedExternally,
    [Boolean]$ShowDialog,
    [int]$NbLines
)
        $MB = New-Object System.Windows.Window
        $MB.Width = 500
        $MB.Height = 55*$NbLines
        $MB.WindowStyle = "None"

        $MBLabel = New-Object System.Windows.Controls.Label
        $MBLabel.HorizontalAlignment = "Center"
        $MBLabel.FontFamily = "Verdana"
        $MBLabel.Content = $Message

        $MBStackPanel = New-Object System.Windows.Controls.StackPanel
        $MBStackPanel.VerticalAlignment = "Top"

        If ($ClosedExternally -eq $False){
            $MBButton = New-Object System.Windows.Controls.Button
            $MBButton.Content = "Close"
            $MBButton.HorizontalAlignment = "Center"
            $MBButton.Width="100"
            $MBButton.Add_Click({
            $this.parent.parent.Close()
            })
            $result=$MBStackPanel.Children.Add($MBLabel)
            $result=$MBStackPanel.Children.Add($MBButton)
            $MB.Content = $MBStackPanel
        }
        else
        {
        $MB.Content = $MBLabel
        }
        
        If ($ShowDialog){$MB.ShowDialog()} else {$MB.Show()}
        #write-host ("POPUP : " + $Message.Replace("`r`n"," "))
        Return $MB
        
}



function ListRunbooks()
# This function is being called when loading the tool
# and everytime the 'Update List' button is being clicked or called
# (the button is also called when hitting enter in the database server
# or database port textboxes)
# The function recurses through the Runbooks in the database
# to fill the TreeView in the GUI by calling the FillNode() function
{

param (
    [System.Windows.Controls.TreeView]$Tree
)

        $Global:ProgressCount = 0

        write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] Trying to connect to server $Global:DatabaseServer"
        Write-Progress -Activity "Connecting to server $Global:DatabaseServer and retrieving Runbooks..."
        $Tree.Items.Clear()
        $myConnection = New-Object System.Data.SqlClient.SqlConnection $Global:SQLConnstr 
        $eap = $ErrorActionPreference = "SilentlyContinue"
        $myConnection.Open()
        if (!$?) {
            $ErrorActionPreference =$eap
            $MB = popup2 -Message ("Runbook hierarchy cannot be displayed.`r`nConnection to database server " + $Global:DatabaseServer + " could not be opened.`r`nPlease configure or check the server name on the next screen and try again.") -ClosedExternally $False -NbLines 2 -ShowDialog $True
            $Tree.IsEnabled= $False
            }
            else{  
            $ErrorActionPreference =$eap
            write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] Connected to database, retrieving Runbooks..."
            $NodeRoot = New-Object System.Windows.Controls.TreeViewItem 
            $NodeRoot.Header = "Runbooks"
            $NodeRoot.Name = "Folder"
            $NodeRoot.Tag = "00000000-0000-0000-0000-000000000000"
            [void]$Tree.Items.Add($NodeRoot)
            FillNode($NodeRoot)
            $Tree.IsEnabled= $True
            write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] Runbooks parsing finished..."
            $myConnection.Close()
            }
        Write-Progress -Activity "Connecting to server $Global:DatabaseServer and retrieving Runbooks..." -Complete
}

function FillNode()
# This function is being called by the ListRunbooks() function
# to fill details of a specific folder in the Orchestrator
# hierarchy (subfolders and Runbooks at the root of the folder)
# For subfolders, it actually recurses on itself
# Runbooks are leaf objects in the recursion
 {
 
param (
    [System.Windows.Controls.TreeViewItem]$TreeNode
)

       
        $Global:ProgressCount += 10
        If ($Global:ProgressCount -gt 100) { $Global:ProgressCount = 0}
        Write-Progress -Activity "Connecting to server $Global:DatabaseServer and retrieving Runbooks..." -PercentComplete $Global:ProgressCount
        #Retrieve folders
        $SqlQuery = "SELECT Name, UniqueID from FOLDERS WHERE ParentID='" + $TreeNode.Tag + "' AND deleted = 'False' ORDER BY Name"
        $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
        $dr = $myCommand.ExecuteReader()
        while ($dr.Read())
            {
            $NewNode = New-Object System.Windows.Controls.TreeViewItem 
            $NewNode.Header = $dr["Name"]
            $NewNode.Tag = $dr["UniqueID"]
            $NewNode.Name = "Folder"
            [void]$TreeNode.Items.Add($NewNode)
            }
        $dr.Close()

        #Retrieve Runbooks
        $SqlQuery = "select DISTINCT POLICIES.Name AS PName, POLICIES.UniqueID AS PID, FOLDERS.Name As PFName from POLICIES, FOLDERS where FOLDERS.UniqueID = POLICIES.ParentID AND POLICIES.Deleted = 0 AND POLICIES.ParentID = '" + $TreeNode.Tag + "' ORDER BY POLICIES.NAME"
        $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
        $dr = $myCommand.ExecuteReader()
        while ($dr.Read())
            {
            $NewNode = New-Object System.Windows.Controls.TreeViewItem
            $NewNode.Header = $dr["PName"]
            $NewNode.Name = "Runbook"
            $NewNode.Tag = $dr["PID"]
            $NewNode.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Blue)
            [void]$TreeNode.Items.Add($NewNode)
            }
       $dr.Close()
       #Continue recursive search in subfolders
       ForEach ($NewNode In $TreeNode.Items)
            { If ($NewNode.Name.Substring(0,6) -eq "Folder") {FillNode($NewNode)} }
}  
 
 function LoadConfig()
 # This function is used to read the database server name and port, if the file "Config-SMART-DCH.xml"
 # exists in the same directory as the Documentation and Conversion Helper PS1 script
 {

  param (
    [String]$ConfigFileLocation

  )

        write-host -ForegroundColor gray "["(date -format "HH:mm:ss")"] Loading Configuration from file $ConfigFileLocation..."
        $XmlReader= New-Object System.Xml.XmlTextReader($ConfigFileLocation)
        While ($XmlReader.Read()){
                              If ($XmlReader.NodeType -eq [System.Xml.XmlNodeType]::Element){
                                   switch ($XmlReader.Name){
                                        "DatabaseServer" {$Global:DatabaseServer=$XmlReader.ReadString()}
                                        "DatabasePort"{$Global:DatabasePort=$XmlReader.ReadString()}
                                        "VisioTemplate"{$Global:VisioTemplate=$XmlReader.ReadString()}
                                        "VisioStencil"{$Global:VisioStencil=$XmlReader.ReadString()}
                                        "VisioCallout"{$Global:VisioCallout=$XmlReader.ReadString()}
                                        "VisioGlueFrom"{$Global:VisioGlueFrom=$XmlReader.ReadString()}
                                        "VisioGlueTo"{$Global:VisioGlueTo=$XmlReader.ReadString()}
                                    }
                               }
        }
        $XmlReader.Close()
 }

 function SaveConfig()
 # This function is called when exiting the tool, to remember the last database name and server used
 # Content is stored in Config-SMART-DCH.xml by default, in the same directory as the
 # Documentation and Conversion Helper PS1 script
 {

 param (
    [String]$ConfigFileLocation,
    [String]$DBServer,
    [String]$DBPort,
    [String]$VTemplate,
    [String]$VStencil,
    [String]$VCallout,
    [String]$VGlueFrom,
    [String]$VGlueTo
 )

 write-host -ForegroundColor gray "["(date -format "HH:mm:ss")"] Saving Configuration to file $ConfigFileLocation..."

 $XMLData = 
@”
<DefaultConfiguration>
<DatabaseServer>$DBServer</DatabaseServer>
<DatabasePort>$DBPort</DatabasePort>
<VisioTemplate>$VTemplate</VisioTemplate>
<VisioStencil>$VStencil</VisioStencil>
<VisioCallout>$VCallout</VisioCallout>
<VisioGlueFrom>$VGlueFrom</VisioGlueFrom>
<VisioGlueTo>$VGlueTo</VisioGlueTo>
</DefaultConfiguration>
“@

$XMLData | Out-File $ConfigFileLocation -Force

 }

########################################################################################
#Make sure we run elevated, or relaunch as admin
########################################################################################


$CurrentScriptDirectory = $PSCommandPath.Substring(0,$PSCommandPath.LastIndexOf("\"))
Set-Location $CurrentScriptDirectory


    #Thanks to http://gallery.technet.microsoft.com/scriptcenter/63fd1c0d-da57-4fb4-9645-ea52fc4f1dfb
    $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator") 
        if (-not $IsAdmin)  
        {  
            try 
            {  
                $ScriptToLaunch = $PSCommandPath.Substring(0,$PSCommandPath.LastIndexOf("\")) + "\SMART-DocumentationConversionHelper.ps1"
                $arg = "-file `"$($ScriptToLaunch)`"" 
                write-host -ForegroundColor yellow "["(date -format "HH:mm:ss")"] WARNING : This script should run with administrative rights - Relaunching the script in elevated mode in 3 seconds..."
                start-sleep 3
                Start-Process "$psHome\powershell.exe" -Verb Runas -ArgumentList $arg -ErrorAction 'stop'

            } 
            catch 
            { 
                write-host -ForegroundColor red "["(date -format "HH:mm:ss")"] Error : Failed to restart script with administrative rights - please make sure this script is launched elevated."  
                break               
            } 
            exit
        }
        else
        {
        write-host -ForegroundColor gray "["(date -format "HH:mm:ss")"] We are running in elevated mode, we can proceed with launching the tool."
        }

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

##########################################################################################
# Visio Settings GUI and functions
########################################################################################## 

[XML]$XAMLVisioSettings = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        ResizeMode="NoResize"
        Title="Visio Settings" Height="225" Width="870">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="30"/>
            <RowDefinition Height="30"/>
            <RowDefinition Height="30"/>
            <RowDefinition Height="30"/>
            <RowDefinition Height="30"/>
            <RowDefinition Height="30"/>
        </Grid.RowDefinitions>
        <Label FontWeight="Bold" Content="Please confirm the following Visio settings, and just close the window when finished. These settings will then be automatically applied." VerticalAlignment="Center" Grid.Row="0"></Label>
        <TextBox Text="C:\Program Files (x86)\Microsoft Office\Office15\Visio Content\1033\BASFLO_U.VSSX" Name="VisioTemplateTextBox" Grid.Row="1" HorizontalAlignment="Left" VerticalAlignment="Center"  Margin="250,0,0,0" Width="500"></TextBox>
        <Label Content="Default stencil for activities" VerticalAlignment="Center" Grid.Row="2"></Label>
        <ComboBox Name="VisioStencilComboBox" IsEditable="False" HorizontalAlignment="Left" VerticalAlignment="Center" Width="150"  Margin="250,0,0,0" Grid.Row="2"></ComboBox>
        <Label Content="Default callout for descriptions" VerticalAlignment="Center" Grid.Row="3"></Label>
        <ComboBox Name="VisioCalloutComboBox" IsEditable="False" HorizontalAlignment="Left" VerticalAlignment="Center" Width="150"  Margin="250,0,0,0" Grid.Row="3"></ComboBox>
        <Label Content="Link from" VerticalAlignment="Center" Grid.Row="4"></Label>
        <ComboBox Name="VisioGlueFromComboBox" IsEditable="False" HorizontalAlignment="Left" VerticalAlignment="Center" Width="150"  Margin="250,0,0,0" Grid.Row="4"></ComboBox>
        <Label Content="Link to" VerticalAlignment="Center" Grid.Row="5"></Label>
        <ComboBox Name="VisioGlueToComboBox" IsEditable="False" HorizontalAlignment="Left" VerticalAlignment="Center" Width="150"  Margin="250,0,0,0" Grid.Row="5"></ComboBox>
    </Grid>
</Window>

'@


##########################################################################################
# Main process defining the GUI
# and initiatlizing some global variables
# Events handling buttons and clicks are also added here
########################################################################################## 

# Form and GUI definition
[XML]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        ResizeMode="NoResize"
        Title="SMART Documentation and Conversion Helper" Height="560" Width="500">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="30"/>
            <RowDefinition Height="300"/>
            <RowDefinition Height="30"/>
            <RowDefinition Height="30"/>
            <RowDefinition Height="30"/>
            <RowDefinition Height="30"/>
            <RowDefinition Height="30"/>
            <RowDefinition Height="30"/>
            <RowDefinition Height="30"/>
        </Grid.RowDefinitions>
        <Label Content="Orchestrator databse (Server\instance) :" VerticalAlignment="Center" Grid.Row="0"></Label>
        <TextBox Text="localhost" Name="DBServer" Grid.Row="0" HorizontalAlignment="Center" VerticalAlignment="Center"  Margin="65,0,0,0" Width="120"></TextBox>
        <TextBox Text="1433" Name="DBPort" Grid.Row="0" HorizontalAlignment="Center" VerticalAlignment="Center" Margin="230,0,0,0" Width="35"></TextBox>
        <Button IsDefault="true" Content="Update List" Name="UpdateDBServer" HorizontalAlignment="Right" VerticalAlignment="Center" Width="70" Margin="0,0,10,0" Grid.Row="0"/>
        <TreeView Name="Tree" Grid.Row="1" Margin="2"/>
        <Button Content="Export" Width="150" IsEnabled="True" Height="100" HorizontalAlignment="Left" Margin="7,0,0,0" Name="ExportButton" Grid.Row="2" Grid.RowSpan="4"/>
        <CheckBox Content="Export to Visio" IsEnabled="False" Name="ExportVSDCB" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="200,0,0,0" Grid.Row="2"></CheckBox>
        <CheckBox Content="Save and Close" IsEnabled="False" Name="SaveAndCloseVSD" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="215,0,0,0" Grid.Row="3"></CheckBox>
        <CheckBox Content="Group Thumbnails" IsEnabled="False" IsChecked="True" Name="GroupThumbnails" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="330,0,0,0" Grid.Row="3"></CheckBox>
        <CheckBox Content="Export to Word" IsEnabled="False" Name="ExportDOCCB" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="200,0,0,0" Grid.Row="4"></CheckBox>
        <CheckBox Content="Save and Close" IsEnabled="False" Name="SaveAndCloseDOC" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="215,0,0,0" Grid.Row="5"></CheckBox>
        <CheckBox Content="Export to PowerShell Workflow" IsEnabled="True" Name="ExportPS1CB" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="200,0,0,0" Grid.Row="6"></CheckBox>
        <CheckBox Content="Save and Close" IsEnabled="False" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="215,0,0,0" Name="SaveAndClosePS1" IsChecked="False" Grid.Row="7"></CheckBox>  
        <CheckBox Content="Skeleton Only" IsEnabled="False" Name="SkeletonOnly" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="330,0,0,0" Grid.Row="7"></CheckBox>
      
        <Button Content="Visio Settings ..." Width="150" IsEnabled="False" Height="27" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="7,0,0,0" Name="VisioSettings" Grid.Row="6"/>
        <Label Content="[Images Exported (0)]" IsEnabled="False" Height="27" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="7,0,0,0" Name="ExtractImages" Grid.Row="7"></Label>
        

    </Grid>
</Window>

'@
$Reader = (New-Object System.XML.XMLNodeReader $XAML)
$FORM = [Windows.Markup.XAMLReader]::Load($Reader)
# Linking variables to the GUI
$Tree = New-Object System.Windows.Controls.TreeView
$Tree = $FORM.FindName('Tree')
$Tree.IsEnabled= $False
$Global:SkeletonCB = $FORM.FindName('SkeletonOnly')
$DBServerTB = $FORM.FindName('DBServer')
$DBPortTB = $FORM.FindName('DBPort')
$Global:SaveAndClosePS1 = $FORM.FindName('SaveAndClosePS1')
# GUI Events
$Tree.Add_MouseDoubleClick({
    $CurrentNode = $Tree.SelectedItem
    If ($CurrentNode.Name.Contains("Runbook")) {
        write-host "Working with Runbook " $CurrentNode.Header "(RunbookID" $CurrentNode.Tag "in the database)"
        GeneratePS1 -RunbookID $CurrentNode.Tag -RunbookName $CurrentNode.Header
    }
})

$FORM.FindName('ExportButton').Add_Click({
    $CurrentNode = $Tree.SelectedItem
    If ($CurrentNode.Name.Contains("Runbook")) {
        write-host  -ForegroundColor yellow "["(date -format "HH:mm:ss")"] Working with Runbook " $CurrentNode.Header

        $SimplifiedRunbookName = $CurrentNode.Header.Replace(" ", "")
        $SimplifiedRunbookName = $SimplifiedRunbookName.Replace("/", "")
        $SimplifiedRunbookName = $SimplifiedRunbookName.Replace("\", "")
        $SimplifiedRunbookName = $SimplifiedRunbookName.Replace(">", "")
        $SimplifiedRunbookName = $SimplifiedRunbookName.Replace("<", "")
        $SimplifiedRunbookName = $SimplifiedRunbookName.Replace(":", "")
        $SimplifiedRunbookName = $SimplifiedRunbookName.Replace("*", "")
        $SimplifiedRunbookName = $SimplifiedRunbookName.Replace("?", "")
        $SimplifiedRunbookName = $SimplifiedRunbookName.Replace("|", "")
        $SimplifiedRunbookName = $SimplifiedRunbookName.Replace("-", "")

        If ($FORM.FindName('ExportVSDCB').IsChecked)
            {
            If (Test-Path -Path $Global:VisioTemplate)
                {
                write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] Visio Export Started..."
                GenerateVSD -RunbookID $CurrentNode.Tag -RunbookName $CurrentNode.Header -ExportFileName $SimplifiedRunbookName
                write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] Visio Export Finished"
                }
                else
                {
                write-host  -ForegroundColor red "["(date -format "HH:mm:ss")"] Visio Templates path is not valid. The file path does not resolve. Please update the Visio settings with a valid path and retry..."
                }
            } else {write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] Visio Export not requested"}
        If ($FORM.FindName('ExportDOCCB').IsChecked)
            {
            write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] Word Export Started..."
            GenerateDOC -RunbookID $CurrentNode.Tag -RunbookName $CurrentNode.Header -ExportFileName $SimplifiedRunbookName
            write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] Word Export Finished"
            } else {write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] Word Export not requested"}
        If ($FORM.FindName('ExportPS1CB').IsChecked)
            {
            write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] PowerShell Export Started..."
            GeneratePS1 -RunbookID $CurrentNode.Tag -RunbookName $CurrentNode.Header -ExportFileName $SimplifiedRunbookName
            write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] PowerShell Export Finished"
            } else {write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] PowerShell Export not requested"}
        write-host  -ForegroundColor yellow "["(date -format "HH:mm:ss")"] Export process finished for Runbook " $CurrentNode.Header
        popup2 -Message ("Export process finished for Runbook " + $CurrentNode.Header) -ClosedExternally $False -NbLines 2
        }
    else {$MB = popup2 -Message "No Runbook was selected.`r`nPlease select a Runbook before trying to export to Visio.`r`nRunbooks are displayed in blue in the tree view." -ClosedExternally $False -NbLines 2}
})
$FORM.FindName('UpdateDBServer').Add_Click({
        $Global:DatabaseServer = $DBServerTB.Text
        $Global:DatabasePort = $DBPortTB.Text
        $Global:SQLConnstr = "Server=" + $Global:DatabaseServer +"," +  $Global:DatabasePort + ";Integrated Security=SSPI;database=" + $DatabaseName
        ListRunbooks($Tree)
})

$FORM.FindName('ExportVSDCB').Add_Checked({
    $FORM.FindName('SaveAndCloseVSD').IsEnabled = $True
    $FORM.FindName('GroupThumbnails').IsEnabled = $True
})

$FORM.FindName('ExportVSDCB').Add_UnChecked({
    $FORM.FindName('SaveAndCloseVSD').IsEnabled = $False
    $FORM.FindName('GroupThumbnails').IsEnabled = $False
})

$FORM.FindName('ExportDOCCB').Add_Checked({
    $FORM.FindName('SaveAndCloseDOC').IsEnabled = $True
})

$FORM.FindName('ExportDOCCB').Add_UnChecked({
    $FORM.FindName('SaveAndCloseDOC').IsEnabled = $False
})

$FORM.FindName('ExportPS1CB').Add_Checked({
    $FORM.FindName('SaveAndClosePS1').IsEnabled = $True
    $FORM.FindName('SkeletonOnly').IsEnabled = $True
})

$FORM.FindName('ExportPS1CB').Add_UnChecked({
    $FORM.FindName('SaveAndClosePS1').IsEnabled = $False
    $FORM.FindName('SkeletonOnly').IsEnabled = $False
})

$FORM.FindName('VisioSettings').Add_Click({
    $ReaderVisioSettings = (New-Object System.XML.XMLNodeReader $XAMLVisioSettings)
    $FORMVisioSettings = [Windows.Markup.XAMLReader]::Load($ReaderVisioSettings)

    $FORMVisioSettings.FindName('VisioTemplateTextBox').Text = $Global:VisioTemplate

    $FORMVisioSettings.FindName('VisioGlueFromComboBox').Items.Add("Bottom") | out-null
    $FORMVisioSettings.FindName('VisioGlueFromComboBox').Items.Add("Top") | out-null
    $FORMVisioSettings.FindName('VisioGlueFromComboBox').Items.Add("Left") | out-null
    $FORMVisioSettings.FindName('VisioGlueFromComboBox').Items.Add("Right") | out-null
    $FORMVisioSettings.FindName('VisioGlueFromComboBox').SelectedIndex = $FORMVisioSettings.FindName('VisioGlueFromComboBox').Items.IndexOf($Global:VisioGlueFrom)

    $FORMVisioSettings.FindName('VisioGlueToComboBox').Items.Add("Bottom") | out-null
    $FORMVisioSettings.FindName('VisioGlueToComboBox').Items.Add("Top") | out-null
    $FORMVisioSettings.FindName('VisioGlueToComboBox').Items.Add("Left") | out-null
    $FORMVisioSettings.FindName('VisioGlueToComboBox').Items.Add("Right") | out-null
    $FORMVisioSettings.FindName('VisioGlueToComboBox').SelectedIndex = $FORMVisioSettings.FindName('VisioGlueToComboBox').Items.IndexOf($Global:VisioGlueTo)


        $FORMVisioSettings.FindName('VisioStencilComboBox').Clear()
        $FORMVisioSettings.FindName('VisioCalloutComboBox').Items.Clear()

        $vApp = New-Object -ComObject Visio.InvisibleApp
        
        $vDoc = $vApp.Documents.Add("")

        $vStencil = $vApp.Documents.OpenEx($FORMVisioSettings.FindName('VisioTemplateTextBox').Text, 4)
        foreach ($vFlowChartMaster in $vStencil.Masters)
            {$FORMVisioSettings.FindName('VisioStencilComboBox').Items.Add($vFlowChartMaster.Name)}
        $vDoc.Close()

        If ($FORMVisioSettings.FindName('VisioStencilComboBox').Items -contains "Process")
            {$FORMVisioSettings.FindName('VisioStencilComboBox').SelectedIndex = $FORMVisioSettings.FindName('VisioStencilComboBox').Items.IndexOf($Global:VisioStencil)}

        $vDoc2 = $vApp.Documents.OpenEx($vApp.GetBuiltInStencilFile(3, 2), 64)
        foreach ($vFlowChartMaster In $vDoc2.Masters)
            {$FORMVisioSettings.FindName('VisioCalloutComboBox').Items.Add($vFlowChartMaster.Name)}
        $vDoc2.Close()
        If ($FORMVisioSettings.FindName('VisioCalloutComboBox').Items -contains "Word Balloon")
            {$FORMVisioSettings.FindName('VisioCalloutComboBox').SelectedIndex = $FORMVisioSettings.FindName('VisioCalloutComboBox').Items.IndexOf($Global:VisioCallout)}

        $vApp.Quit()

        $FORMVisioSettings.ShowDialog() | Out-Null

    $Global:VisioTemplate = $FORMVisioSettings.FindName('VisioTemplateTextBox').Text
    $Global:VisioStencil = $FORMVisioSettings.FindName('VisioStencilComboBox').Text
    $Global:VisioCallout = $FORMVisioSettings.FindName('VisioCalloutComboBox').Text
    $Global:VisioGlueFrom = $FORMVisioSettings.FindName('VisioGlueFromComboBox').Text
    $Global:VisioGlueTo = $FORMVisioSettings.FindName('VisioGlueToComboBox').Text
    $FORMVisioSettings.Close()
    SaveConfig  -ConfigFileLocation $ConfigFileLocation -DBServer $Global:DatabaseServer -DBPort $Global:DatabasePort -VTemplate $Global:VisioTemplate -VStencil $Global:VisioStencil -VCallout $Global:VisioCallout -VGlueFrom $Global:VisioGlueFrom -VGlueTo $Global:VisioGlueTo

    })


#Only enable Visio and Word options is they are installed

If (Test-Path "HKLM:\Software\Classes\.doc\Word.Document.8\")
    {
    $FORM.FindName('ExportDOCCB').IsEnabled = $True
    $NumberOfThumbnails = (Get-Item "*.jpg").Count
    $FORM.FindName('ExtractImages').Content = "[Images Exported ($NumberOfThumbnails)]"
    }

If (Test-Path "HKLM:\Software\Classes\.vsd\Visio.Drawing.11\")
    {
    $FORM.FindName('ExportVSDCB').IsEnabled = $True
    $FORM.FindName('VisioSettings').IsEnabled = $True
    $NumberOfThumbnails = (Get-Item "*.jpg").Count
    $FORM.FindName('ExtractImages').Content = "[Images Exported ($NumberOfThumbnails)]"
    If ($NumberOfThumbnails -lt 10)
        {write-host -ForegroundColor gray "["(date -format "HH:mm:ss")"] WARNING : Only $NumberOfThumbnails image(s) were found in the local folder... This may mean that you haven't run the Image Export script yet (SMA-DocumentationConversionHelper-ImageExport.ps1). Visio and Word export will still work, but will use a default image."}
    If ((Test-Path "default.jpg") -eq $false)
        {write-host -ForegroundColor gray "["(date -format "HH:mm:ss")"] WARNING : 'default.jpg' thumbnail not found in the working folder. Make sure you copied it from the download package, or that you are calling the script from the same folder where this file is located. Without this file, Visio and Word export will return errors when adding thumbnails."}
    }

##########################################################################################
# Actual start of the main process
##########################################################################################
cls
write-host -ForegroundColor green "["(date -format "HH:mm:ss")"] SMART Documentation and Conversion Helper $ToolVersion"
# We default to the database and port values from the script, as well as initialize Visio variables
$Global:DatabaseServer = $DefaultDatabaseServer
$Global:DatabasePort = $DefaultDatabasePort
$DatabaseName = $DefaultDatabaseName
$Global:VisioTemplate = "C:\Program Files (x86)\Microsoft Office\Office15\Visio Content\1033\BASFLO_U.VSSX"
$Global:VisioStencil = "Process"
$Global:VisioCallout = "Word Balloon"
$Global:VisioGlueFrom = "Bottom"
$Global:VisioGlueTo = "Top"
# Let's try to see if there is a config file and, if yes, update the global variables
$ConfigFileLocation = (Get-Location -PSProvider FileSystem).Path + "\Config-SMART-DCH.xml"
If (Test-Path -Path $ConfigFileLocation){LoadConfig -ConfigFileLocation $ConfigFileLocation}
# We can now update the GUI and define the connection string used throughout the script (also updated through $FORM.FindName('UpdateDBServer').Add_Click as needed)
$DBServerTB.Text = $Global:DatabaseServer
$DBPortTB.Text = $Global:DatabasePort
$Global:SQLConnstr = "Server=" + $Global:DatabaseServer +"," +  $Global:DatabasePort + ";Integrated Security=SSPI;database=" + $DatabaseName

ListRunbooks($Tree)
write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] Displaying GUI..."
$FORM.ShowDialog() | Out-Null
# Everything starting this line is done when exiting the tool
SaveConfig  -ConfigFileLocation $ConfigFileLocation -DBServer $Global:DatabaseServer -DBPort $Global:DatabasePort -VTemplate $Global:VisioTemplate -VStencil $Global:VisioStencil -VCallout $Global:VisioCallout -VGlueFrom $Global:VisioGlueFrom -VGlueTo $Global:VisioGlueTo
write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] Exiting..."

##########################################################################################