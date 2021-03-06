Function Get-Log{
    <#
    .SYNOPSIS
    Gather log files from remote computers.
   
   .DESCRIPTION
    This script will gather the log files from remote computers by supplying a computer name and the type of log file you wish to collect.
   
   .PARAMETER ComputerName
    The target computer you wish to gather logs from.
   
   .PARAMETER LogType
    The log type you wish to gather from the remote computer.
   
   .EXAMPLE
    Get-Log -ComputerName computer1 -LogType WindowsUpdate
    Pulls the Windows Update Log from computer1 into your console for review.
    .EXAMPLE
    Get-Log -ComputerName computer1 -LogType DISM -Verbose
    Pulls the DISM log into your console from computer1, and uses verbose output
    to tell you what is happening on each part of the script.
   
   #>
    Param(
    [cmdletBinding()]
   
   [Parameter(
    Mandatory,
    Position=0)]
    [string]
    $ComputerName,
    
    [Parameter(
    Mandatory,
    Positon=1)]
    #You can change/add to/remove items between the () in lines 33-36 to suit your needs.
    #Make sure you wrap your text in "".
    [ValidateSet( 
    "WindowsUpdate",
    "FoG",
    "DISM")]
    [string]
    $LogType
    )
   
   Write-Verbose -Message "Testing WSMan connection and creating session..."
    #Test-WSMan Connection
    try {
    If(Test-WSMan -ComputerName $ComputerName){
   
   $Session = New-PSSession -ComputerName $ComputerName
    
    }
    }
    catch {
    
    $_.Exception.Message
    Write-Error -Message "Unable to contact remote computer via WinRM, is Powershell Remoting enabled?"
    Break
    
    }
   
   Write-Verbose -Message "Gathering requested log file(s)"
    switch ($LogType) {
    #For each log type you specified starting on line 35, construct a line just like below. 
    #Log files are static in where they are kept, so hard-coding location is typically fine.
    "WindowsUpdate" { Invoke-Command -Session $Session -ScriptBlock { Get-Content C:\Windows\WindowsUpdate.log } }
    "FoG" { Invoke-Command -Session $Session -ScriptBlock { Get-Content C:\fog.log } }
    "DISM" { Invoke-Command -Session $Session -ScriptBlock { Get-Content C:\Windows\Logs\DISM\dism.log } }
    }
    
    Write-Verbose -Message "Cleaning up remote session."
    Get-PSSession | Remove-PSSession
    
   }