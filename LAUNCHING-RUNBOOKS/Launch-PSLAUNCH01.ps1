$serviceURL = "http://dtekorch16-s1:81/Orchestrator2012/Orchestrator.svc";
$runbookPath = "\4.DTEK PROJECTS\LAUNCH-TEST\PS-LAUNCH01";
$parameters = @{"recipient" = "danny@dtek.com"; "message" = "Launched PSLAUNCH01 runbook"}

# Get the runbook based on the path
$fullUrl = [String]::Format("{0}/Runbooks?`$filter=Path eq '{1}'", $serviceURL, $runbookPath);
$request = [System.Net.HttpWebRequest]::Create($fullUrl)

# Build up a nice User Agent
$request.UserAgent = $(
   "{0} (PowerShell {1}; .NET CLR {2}; {3})" -f $UserAgent, $(if($Host.Version){$Host.Version}else{"1.0"}),
   [Environment]::Version,
   [Environment]::OSVersion.ToString().Replace("Microsoft Windows ", "Win")
)
$request.UseDefaultCredentials = $true

[System.Net.HttpWebResponse] $response = [System.Net.HttpWebResponse] $request.GetResponse()
$reader = [IO.StreamReader] $response.GetResponseStream()
$responseString = $reader.ReadToEnd()
$reader.Close()
$response.Close()
$xml = [xml]$responseString

$runbookID = $xml.feed.entry.content.properties.Id.InnerText

##------------------------------------------------------------------------------------------------

# Get the runbook parameters based on the runbook ID
$fullUrl = [String]::Format("{0}/Runbooks(guid'{1}')/Parameters", $serviceURL, $runbookId);
$request = [System.Net.HttpWebRequest]::Create($fullUrl)

# Build up a nice User Agent
$request.UserAgent = $(
   "{0} (PowerShell {1}; .NET CLR {2}; {3})" -f $UserAgent, $(if($Host.Version){$Host.Version}else{"1.0"}),
   [Environment]::Version,
   [Environment]::OSVersion.ToString().Replace("Microsoft Windows ", "Win")
)
$request.UseDefaultCredentials = $true

[System.Net.HttpWebResponse] $response = [System.Net.HttpWebResponse] $request.GetResponse()
$reader = [IO.StreamReader] $response.GetResponseStream()
$responseString = $reader.ReadToEnd()
$reader.Close()
$response.Close()
$xml = [xml]$responseString

# Format the param string from the Parameters hashtable
$rbParamString = "<d:Parameters><![CDATA[<Data>"
foreach ($entity in $xml.feed.entry)
{
   if ($entity.content.properties.Direction -eq "In")
   {
      $rbParamString = -join ($rbParamString,"<Parameter><ID>{",$entity.content.properties.Id.InnerText,"}</ID>
<Value>",$parameters[$entity.content.properties.Name],"</Value></Parameter>")
   }
}
$rbParamString += "</Data>]]></d:Parameters>"

##--------------------------------------------------------------------------------------------

# Create the request object for submitting the job
$fullUrl = [String]::Format("{0}/Jobs", $serviceURL);
$request = [System.Net.HttpWebRequest]::Create($fullUrl)

# Set the credentials to default or prompt for credentials
$request.UseDefaultCredentials = $false

# Build the request header
$request.Method = "POST"
$request.UserAgent = $(
   "{0} (PowerShell {1}; .NET CLR {2}; {3})" -f $UserAgent, $(if($Host.Version){$Host.Version}else{"1.0"}),
   [Environment]::Version,
   [Environment]::OSVersion.ToString().Replace("Microsoft Windows ", "Win")
)
$request.Accept = "application/atom+xml,application/xml"
$request.ContentType = "application/atom+xml"
$request.KeepAlive = $true
$request.Headers.Add("Accept-Encoding","identity")
$request.Headers.Add("Accept-Language","en-US")
$request.Headers.Add("DataServiceVersion","1.0;NetFx")
$request.Headers.Add("MaxDataServiceVersion","2.0;NetFx")
$request.Headers.Add("Pragma","no-cache")

# Build the request body
$requestBody = @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<entry xmlns:d="http://schemas.microsoft.com/ado/2007/08/dataservices" xmlns:m="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata" xmlns="http://www.w3.org/2005/Atom">
    <content type="application/xml">
        <m:properties>
            <d:RunbookId m:type="Edm.Guid">$runbookID</d:RunbookId>
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
    Write-Host "Successfully started runbook. Job ID: " $jobId
}
else
{
    Write-Host "Could not start runbook. Status: " $response.StatusCode
}