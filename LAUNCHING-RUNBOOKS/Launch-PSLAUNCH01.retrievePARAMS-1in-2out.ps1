##########################
# Powershell Example of calling an Orchestrator Runbook
# and retrieving the runbook output 
 
#Credentials
$secpasswd = ConvertTo-SecureString "Kerrisue1" -AsPlainText -Force
$mycreds   = New-Object System.Management.Automation.PSCredential ("dtek\danny", $secpasswd)
 
#Orchestrator Server Hostname
$OrchServer = "dtekorch16-s1"
 
# Runbook to be called with an example Property and Value
$RunBookName = "PS-LAUNCH02"

$RunbookInputProperty = "recipient"
$RunbookInputValue    = "danny@dtek.com"
 
#Begin ################
 
cls
# First Grab the GUID of the desired runbook
 
$OrchURI        = "http://$($OrchServer):81/Orchestrator2012/Orchestrator.svc/Runbooks?`$filter=Name eq '$RunBookName'" 
$ResponseObject = invoke-webrequest -Uri $OrchURI -method Get -Credential $mycreds 
$XML            = [xml] $ResponseObject.Content
$RunbookGUIDURL = $XML.feed.entry.id
 
write-host "Runbook GUID URI = " $RunbookGUIDURL
 
# User the runbook GUID to retrieve GUID values for Input Properties
# This example will retrieve an input property titled "Input1" as set in the variables
# in the header of this script
 
$ResponseObject = invoke-webrequest -Uri "$($RunbookGUIDURL)/Parameters" -method Get -Credential $mycreds 
[System.Xml.XmlDocument] $XML = $ResponseObject.Content
 
# A runbook contains a number of properties for the inputs and results returned from that runbook.
# All values within the returned XML need to be parsed to find which elements are Input properties
# with the desired name.  Once found, the "Id" or GUID for that particular property must be returned.
 
function GetScorchProperty([System.Object]$XMLString, [string]$Name, [string]$Direction, [string]$DesiredData){
 
   $nsmgr = New-Object System.XML.XmlNamespaceManager($XMLString.NameTable)    
   $nsmgr.AddNamespace('d','http://schemas.microsoft.com/ado/2007/08/dataservices')
   $nsmgr.AddNamespace('m','http://schemas.microsoft.com/ado/2007/08/dataservices/metadata')
 
 
   # Create an Array of Properties based on the 'Name' value
 
    $inputs = $XMLString.SelectNodes('//d:Name',$nsmgr)
 
   foreach ($parameter in $inputs){
      # Each 'Name' has related elements at the same level in XML
      # So the parent node is found and a new array of siblings 
      # is created.
 
      #Reset Property values 
      $obName          =""
      $obId            =""
      $obType          =""
      $obDirection     =""
      $obDescription   =""
 
      $siblings = $($parameter.ParentNode.ChildNodes)
 
      # Each of the sibling properties is identified
      foreach ($elements in $siblings){
      # write-host "Element = " $elements.ToString()
          If ($elements.ToString() -eq "Name"){
            $obName = $elements.InnerText
          }   
          If ($elements.ToString() -eq "Id"){
             $obId = $elements.InnerText
          }
          If ($elements.ToString() -eq "type"){
             $obType = $elements.InnerText
          }
          If ($elements.ToString() -eq "Direction"){
             $obDirection = $elements.InnerText
          }
         If ($elements.ToString() -eq "Description"){
            $obDescription = $elements.InnerText
         }
         If ($elements.ToString() -eq "Value"){
           # write-host "Value = "$elements.InnerText
            $obValue = $elements.InnerText
         }
       }
 
        if (($Name -eq $obName) -and ($Direction -eq $obDirection)){
          # "Correct input found"
          #Return the Requested Property
 
         If ($DesiredData -eq "Id"){
            return $obId 
         }
         If ($DesiredData -eq "Value"){
            return $obValue
         }
          }
   }
   return $Null
}
 
#The Function is called to retreive the "Id" Property
# This occurs by:
# + Passing in an XML object
# + Specifying the name of the propery being searched for (Input1)
# + Specifying that the runbook property is an "In" direction property
# + Specifying that the element neded for that property is the GUID based Id
 
$RetreivedGUID = GetScorchProperty $XML $RunbookInputProperty "In" "Id"
write-host "Property GUID = " $RetreivedGUID
 
# Derive the Runbook GUID
$urlstring = $RunbookGUIDURL
#    eg.     "http://server2012:81/Orchestrator2012/Orchestrator.svc/Runbooks(guid'c88ca155-e067-4f37-9723-ef977ac74047')"
 
$RunbookID = $RunbookGUIDURL.Substring($RunbookGUIDURL.Length - 38,36)
write-host "RunbookID = " $RunbookID 
 
# Submitting an Orchestrator Request requires a POST to call a runbook based on its GUID Id value
# Values for required inputs are submitted alongside Property GUIDs
# The XML structure of the data section will resemble:
#
# <Data>
# <Parameter>
#   <ID></ID>
#   <Value></Value>
# </Parameter>
# <Parameter>
#   <ID></ID>
#   <Value></Value>
# </Parameter>
# </Data>
#
# The XML structure uses HTML special entities to represent greater than & less than chracters.
 
$POSTBody = @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<entry xmlns:d="http://schemas.microsoft.com/ado/2007/08/dataservices" xmlns:m="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata" xmlns="http://www.w3.org/2005/Atom">
<content type="application/xml">
<m:properties>
<d:RunbookId type="Edm.Guid">{$($RunbookID)}</d:RunbookId>
<d:Parameters>&lt;Data&gt;&lt;Parameter&gt;&lt;ID&gt;{$($RetreivedGUID)}&lt;/ID&gt;&lt;Value&gt;$($RunbookInputValue)&lt;/Value&gt;&lt;/Parameter&gt;&lt;/Data&gt;</d:Parameters>
</m:properties>
</content>
</entry>
"@
 
 
# Submit Orchestrator Request
$OrchURI = "http://$($OrchServer):81/Orchestrator2012/Orchestrator.svc/Jobs/"
write-host "POST request URI " $OrchURI 
 
$ResponseObject = invoke-webrequest -Uri $OrchURI -method POST -Credential $mycreds -Body $POSTBody -ContentType "application/atom+xml" 
 
#Retrieve the Job ID from the submitted request
$XML               = [xml] $ResponseObject.Content
$RunbookJobURL     = $XML.entry.id
 
write-host "Runbook Job URI " $RunbookJobURL
 
# Runbooks will take some time to complete
# This example ues a simple loop to check if the job is still running
# a production script would use more error handling to ensure that the runbook hadn't failed
 
$status = $xml.entry.content.properties.Status
write-host "Current Status = " $status
 
do
{
	if($status -eq "Pending")
	{
		start-sleep -second 5
		$SleepCounter = $SleepCounter + 1
			if($SleepCounter -eq 20)
			{
				$DoExit="Yes"
			}
	}
	Else
	{
		$DoExit="Yes"
	}
 
    # Query the web service for the current status
     $ResponseObject = invoke-webrequest -Uri "$($RunbookJobURL)" -method Get -Credential $mycreds 
     $XML      = [xml] $ResponseObject.Content
     $RunbookJobURL = $XML.entry.id
     $status = $xml.entry.content.properties.Status
     write-host "Current Status = " $status
}While($DoExit -ne "Yes")
 
 
# As the runbook is no longer active, query the Instance of the submitted job
$ResponseObject = invoke-webrequest -Uri "$($RunbookJobURL)/Instances" -method Get -Credential $mycreds 
 
#Retrieve the Instance ID
$XML                = [xml] $ResponseObject
$RunbookInstanceURL = $XML.feed.entry.id
write-host "Runbook Instance URI " $RunbookInstanceURL
 
#The Instance can be used to retrieve the Parameters for the particular job
$ResponseObject                 = invoke-webrequest -Uri "$($RunbookInstanceURL)/Parameters" -method Get -Credential $mycreds 
[System.Xml.XmlDocument] $xml   = $ResponseObject.Content
 
$RunbookResult1                   = GetScorchProperty $xml "retdata01" "Out" "Value"
$RunbookResult2                  = GetScorchProperty $xml "retdata02" "Out" "Value"
write-host "Runbook Result 1 " $RunbookResult1
write-host "Runbook Result 2 " $RunbookResult2