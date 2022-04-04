#----------------

#Eventlogsource     Scorch

#EventlogName       Orchestrator

#EventID            100-101-102

#Message            Error Message  [Trace]

#EventType          Error-Warning-Information

#-----------------
function AppendLog ([string]$Message)
{
    $script:CurrentAction = $Message
    $script:TraceLog += ((Get-Date).ToString() + "`t" + $Message + " `r`n")
}
[String]$trace=""
[String]$ErrorMessage=""
$stopwatch=  [system.diagnostics.stopwatch]::New()

#------------- Define Parameters ---------------

[String]$Eventlogsource=''
[String]$EventlogName=''
[int]$EventID=1
[String]$EventType= ''
[String]$message=''
[int16]$catagory=1


$Eventlogsource='Scorch'
$EventlogName='Orchestrator'
$EventID=102
$EventType= 'Error'
$message='This is a test SCORCH event'
$catagory=1



Try{

    $stopwatch.Start()
    $timestart=(Get-Date).ToString()
    AppendLog "Script $scriptname version $Scriptversion now executing @ $timestart in PowerShell version [$($PSVersionTable.PSVersion.ToString())] session in a [$([IntPtr]::Size * 8)] bit process"
    AppendLog "Running as user [$([Environment]::UserDomainName)\$([Environment]::UserName)] on host [$($env:COMPUTERNAME)]"
    

Write-EventLog -LogName $EventlogName `
-Source $Eventlogsource `
-EventID $EventID `
-EntryType $EventType `
-Message "$message" `
-Category $catagory `
-RawData 10,20

$ResultStatus = "Success"
throw "Bad thing happened"

$stopwatch.Stop()
$scripttime=$stopwatch.Elapsed.totalseconds

}
Catch{

    $ResultStatus = "Failed"

    $EMessage=$_.Exception.Message
    $ELine=$_.InvocationInfo.line
    $ELNum=$_.InvocationInfo.ScriptLineNumber
    
    AppendLog  "Err Message  --  $EMessage "
    AppendLog  "Err Command -- $ELine"
    AppendLog  "Err Line    --  $ELNum"

}finally{
    $stopwatch.Stop() 
    $scripttime=$stopwatch.Elapsed.totalseconds
    
        if($ErrorMessage.Length -gt 0)
        {AppendLog "Exiting Powershell session with result [$ResultStatus] and error message [$EMessage], script runtime: $scripttime seconds @ $timestart"}
        else
        { AppendLog "Exiting Powershell session with result [$ResultStatus], script runtime: $scripttime seconds @ $timestart"}
        
        $trace=$script:TraceLog
        
        $error.clear()

}

$trace