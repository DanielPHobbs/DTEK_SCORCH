# Set script parameters from runbook data bus and Orchestrator global variables
# Define any inputs here and then add to the $argsArray and script block parameters below 

$DataBusInput1 = \`d.T.~Ed/{19080144-7374-4BAF-BB6C-546EC83EF1D7}.variable1\`d.T.~Ed/
$DataBusInput2 =\`d.T.~Ed/{19080144-7374-4BAF-BB6C-546EC83EF1D7}.variable2\`d.T.~Ed/ 


#-----------------------------------------------------------------------

$ErrorMessageOutside="PSSESSION OK"
$ErrorMessage = ""
$Trace = (Get-Date).ToString() + "`t" + "Runbook activity script started" + " `r`n"
       
# Create argument array for passing data bus inputs to the external script session
$argsArray = @()
$argsArray += $DataBusInput1
$argsArray += $DataBusInput2

Try {
$Session = New-PSSession -ComputerName localhost -ea stop
	
}
catch [Exception]{
$Trace += $_.Exception|format-list -force
$PSSESSIONERROR="true"
}
Catch
{
    $ErrorMessageOutside = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
}

Try {
$ReturnArray = Invoke-Command -Session $Session -Argumentlist $argsArray -ScriptBlock {
	
	
	Param(
        [ValidateNotNullOrEmpty()]
        [string]$DataBusInput1,

        [ValidateNotNullOrEmpty()]
        [string]$DataBusInput2
    )

    # Define function to add entry to trace log variable
    function AppendLog ([string]$Message)
    {
        $script:CurrentAction = $Message
        $script:TraceLog += ((Get-Date).ToString() + "`t" + $Message + " `r`n")
    }

    # Set external session trace and status variables to defaults
    $ResultStatus = ""
    $ErrorMessage = ""
    $script:CurrentAction = ""
    $script:TraceLog = ""

    try 
    {
        # Add startup details to trace log
        AppendLog "Script now executing in external PowerShell version [$($PSVersionTable.PSVersion.ToString())] session in a [$([IntPtr]::Size * 8)] bit process"
        AppendLog "Running as user [$([Environment]::UserDomainName)\$([Environment]::UserName)] on host [$($env:COMPUTERNAME)]"
        AppendLog "Parameter values received: DataBusInput1=[$DataBusInput1]; DataBusInput2=[$DataBusInput2]"

        # The actual work the script does goes here
        AppendLog "Doing first action"
        # Do-Stuff -Value $DataBusInput1

        AppendLog "Doing second action"
        # Do-MoreStuff -Value $DataBusInput2

        # Simulate a possible error
        if($DataBusInput1 -ilike "*bad stuff*")
        {
            throw "ERROR: Encountered bad stuff in the parameter input"
        }

        # Example of custom result value
        $myCustomVariable = "$DataBusInput1 +$DataBusInput2"

        # Validate results and set return status
        AppendLog "Finished work, determining result"
        $EverythingWorked = $true
        if($EverythingWorked -eq $true)
        {
           $ResultStatus = "Success"
        }
        else
        {
            $ResultStatus = "Failed"
        }
    }
    catch
    {
        # Catch any errors thrown above here, setting the result status and recording the error message to return to the activity for data bus publishing
        $ResultStatus = "Failed"
        $ErrorMessage = $error[0].Exception.Message
        AppendLog "Exception caught during action [$script:CurrentAction]: $ErrorMessage"
    }
    finally
    {
        # Always do whatever is in the finally block. In this case, adding some additional detail about the outcome to the trace log for return
        if($ErrorMessage.Length -gt 0)
        {
            AppendLog "Exiting external session with result [$ResultStatus] and error message [$ErrorMessage]"
        }
        else
        {
            AppendLog "Exiting external session with result [$ResultStatus]"
        }
        
    }

    # Return an array of the results. Additional variables like "myCustomVariable" can be returned by adding them onto the array
    $resultArray = @()
    $resultArray += $ResultStatus
    $resultArray += $ErrorMessage
    $resultArray += $script:TraceLog
    $resultArray += $myCustomVariable
    return  $resultArray  
     
}#End Invoke-Command
	
}
	Catch {
$PSSESSIONERROR="script block Failed --$error[0].Exception.Message"
	
	}
# Get the values returned from script session for publishing to data bus
$ResultStatus = $ReturnArray[0]
$ErrorMessage = $ReturnArray[1]
$Trace += $ReturnArray[2]
$MyCustomVariable = $ReturnArray[3]

# Record end of activity script process
$Trace += (Get-Date).ToString() + "`t" + "Script finished" + " `r`n"

# Close the external session
Remove-PSSession $Session