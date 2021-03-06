

#-----------------------------------------------------------------------

$ResultStatus = ""
$ErrorMessage = ""
$Trace = (Get-Date).ToString() + "`t" + "Runbook activity script started" + " `r`n"
       

    # Define function to add entry to trace log variable
    function AppendLog ([string]$Message)
    {
        $script:CurrentAction = $Message
        $script:TraceLog += ((Get-Date).ToString() + "`t" + $Message + " `r`n")
    }

    # Set session trace and status variables to defaults
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

        ##################################################### MAIN CODE ##################################################################

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
        $myCustomVariable = "Something I want to publish back to the runbook data bus"

        ###################################################################################################################################

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
            AppendLog "Exiting session with result [$ResultStatus] and error message [$ErrorMessage]"
        }
        else
        {
            AppendLog "Exiting session with result [$ResultStatus]"
        }
        
    }

# Record end of activity script process
$Trace += (Get-Date).ToString() + "`t" + "Script finished" + " `r`n"


<#
Name	                    Type	    Variable	    Is Collection
Result Status	            String	    ResultStatus	false
Error Message	            String	    ErrorMessage	false
Trace Log	                String	    Trace	        false
My Custom Published Data	String	    MyCustomVariable	false
#>