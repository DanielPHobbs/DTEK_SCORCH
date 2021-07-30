#========================================================================
# Created with: SAPIEN Technologies, Inc., PowerShell Studio 2012 v3.1.35
# Created on:   4/22/2016 2:36 PM
# Created by:   xdanielhobbs
# Organization: DTKEITINNOVATIONS
# Filename:     
#========================================================================


#========================================================================
# Created with: SAPIEN Technologies, Inc., PowerShell Studio 2012 v3.1.35
# Created on:   21/04/2016 8:42 AM
# Created by:   xdanielhobbs
# Organization: DTKEITINNOVATIONS
# Description: SCORCH GET_ALERT Activity substitute
#              edit line 103 for required data
#
# Filename:    
#========================================================================

#read-host -assecurestring | convertfrom-securestring | out-file C:\Temp\encrypt.txt`
$username = "Onelondon\xdanielhobbs"
$password = cat C:\Temp\encrypt.txt | convertto-securestring
$MyCredential = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $password		




#Set-StrictMode -Version 1.0
$resultArray=@()
$ResultStatus = ""
$ErrorMessage = ""
$Trace = (Get-Date).ToString() + "`t" + "Runbook activity script started" + " `r`n"
 
$Session = New-PSSession -ComputerName localhost -Credential $MyCredential

$MyCredential

$resultArray = Invoke-Command  -Session $Session -Argumentlist $MyCredential -ScriptBlock {
    
    Param(   
	[Parameter(Position=0)]
	[System.Management.Automation.Credential()]$mycredential
	)
	$mycredential
	
	
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
      
################     Function Connect to SCOM   #####################
			
try{
	
$ManagementServers=""
	
$MS1 = "PDC2SCM046.onelondon.tfl.local"
$MS2 = "PDC2SCM047.onelondon.tfl.local"
$MS3 = "PDC2SCM048.onelondon.tfl.local"
				

appendlog "Importing SCOM 2012 Module "				
Import-Module OperationsManager
				
appendlog "Attempting to connect to SCOM management group Onelondon "				
New-SCOMManagementGroupConnection –ComputerName $MS1  -Credential $MyCredential
appendlog "Connected to $MS1 "	
}

catch{
appendlog "Error Connecting to server $MS1 connecting to $MS2"
$ErrorMessage = $error[0].Exception.Message
AppendLog "Exception caught during action [$script:CurrentAction]: $ErrorMessage"
Try{	
Import-Module OperationsManager
New-SCOMManagementGroupConnection –ComputerName $MS2 -Credential $MyCredential
appendlog "Connected to $MS2 $Global:NewLine"		
	}
catch {
appendlog "Error Connecting to server $MS2 connecting to $MS3"	
$ErrorMessage = $error[0].Exception.Message
AppendLog "Exception caught during action [$script:CurrentAction]: $ErrorMessage"				
try{
Import-Module OperationsManager
New-SCOMManagementGroupConnection –ComputerName $MS3 -Credential $MyCredential
appendlog "Connected to $MS3 "
		}
		catch{appendlog "Error Connecting to any SCOM management server "
						}
		}           
	}

############# Get Alert ######################
			
$FalertData=Get-SCOMAlert -Criteria "Name like 'IM SCOM Remedy 100 event'" | select *
			
Return $FalertData
			
############# Close Alert ######################			
			

################## Validate results and set return status #################
        AppendLog "Finished work, determining result"
		
        $EverythingWorked = $true
        if($EverythingWorked -eq $true)
        {$ResultStatus = "Success"}
        else
        {$ResultStatus = "Failed"}
    }
    catch
    {$ResultStatus = "Failed"
     $ErrorMessage = $error[0].Exception.Message
     AppendLog "Exception caught during action [$script:CurrentAction]: $ErrorMessage"
    }
    finally
    {if($ErrorMessage.Length -gt 0)
        {AppendLog "Exiting external session with result [$ResultStatus] and error message [$ErrorMessage]" }
        else
        {AppendLog "Exiting external session with result [$ResultStatus]"}
        
    }
	
	#################  extract Data ##############################################
	
	

    ############ Return Data ###################
    $resultArray = @()
    $resultArray += $ResultStatus
    $resultArray += $ErrorMessage
    $resultArray += $script:TraceLog
    $resultArray += $alertdata
    return  $resultArray  
     
} 

# Get the values returned from script session for publishing to data bus
$ResultStatus = $resultArray[0]
$ErrorMessage = $resultArray[1]
$Trace += $resultArray[2]
$AlertResults = $resultArray[3]

# Record end of activity script process
$Trace += (Get-Date).ToString() + "`t" + "Script finished" + " `r`n"

# Close the external session
Remove-PSSession $Session

$trace
$AlertResults