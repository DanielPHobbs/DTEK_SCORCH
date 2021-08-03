# Set script parameters from runbook data bus and Orchestrator global variables
# Define any inputs here and then add to the $argsArray and script block parameters below 

$DataBusInput1 = "{Parameter 1 from Initialize Data}"
$DataBusInput2 = "{Global Variable 1}"


#-----------------------------------------------------------------------

## Initialize result and trace variables
# $ResultStatus provides basic success/failed indicator
# $ErrorMessage captures any error text generated by script
# $Trace is used to record a running log of actions
$ResultStatus = ""
$ErrorMessage = ""
$Trace = (Get-Date).ToString() + "`t" + "Runbook activity script started" + " `r`n"
       
# Create argument array for passing data bus inputs to the external script session
$argsArray = @()
$argsArray += $DataBusInput1
$argsArray += $DataBusInput2

# Establish an external session (to localhost) to ensure 64bit PowerShell runtime using the latest version of PowerShell installed on the runbook server
# Use this session to perform all work to ensure latest PowerShell features and behavior available
$Session = New-PSSession -ComputerName localhost

# Invoke-Command used to start the script in the external session. Variables returned by script are then stored in the $ReturnArray variable
$ReturnArray = Invoke-Command -Session $Session -Argumentlist $argsArray -ScriptBlock {
    # Define a parameter to accept each data bus input value. Recommend matching names of parameters and data bus input variables above
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

        ##################################################### MAIN CODE ##################################################################
       
        $computername="dtekaz-hw01.dtek.com"
        
        Install-Module -Name Invoke-PSSession

        
         ##############################################################
         #Create PSCredentials

         
         #############################################################
         appendlog "Gathering log file(s)"

         $RMSession = Invoke-PSSession -ComputerName $computername -Credential $PSCredential
        
         $LogContent=Invoke-Command -Session $RMSession -ScriptBlock { Get-Content F:\inetpub\logs\logfiles\W3SVC1\u_ex210414.log } 
         $myCustomVariable = $LogContent

         $myCustomVariable
         $EverythingWorked = $true
    
        
        ###################################################################################################################################

        # Validate results and set return status
        AppendLog "Finished work, determining result"
        
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

# Get the values returned from script session for publishing to data bus
$ResultStatus = $ReturnArray[0]
$ErrorMessage = $ReturnArray[1]
$Trace += $ReturnArray[2]
$MyCustomVariable = $ReturnArray[3]

# Record end of activity script process
$Trace += (Get-Date).ToString() + "`t" + "Script finished" + " `r`n"

# Close the external session
Get-PSSession | Remove-PSSession 

<#
https://automys.com/library/asset/powershell-system-center-orchestrator-practice-template
Name	                    Type	    Variable	    Is Collection
Result Status	            String	    ResultStatus	false
Error Message	            String	    ErrorMessage	false
Trace Log	                String	    Trace	        false
My Custom Published Data	String	    MyCustomVariable	false
#>
$trace