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

$computername="dtekaz-hw01.dtek.com"

Write-output  "Testing WSMan connection and creating session..."
 #Test-WSMan Connection
 try {
 If(Test-WSMan -ComputerName $ComputerName){

$Session = New-PSSession -ComputerName $ComputerName
Write-output  "Connected PSSession"
 }
 }
 catch {
 
 $_.Exception.Message
 Write-Error -Message "Unable to contact remote computer via WinRM, is Powershell Remoting enabled?"
 Break
 
 }
 #############################################################
 Write-output  "Gathering log file(s)"

 $LogContent=Invoke-Command -Session $Session -ScriptBlock { Get-Content F:\inetpub\logs\logfiles\W3SVC1\u_ex210414.log }
 $LogContent
 ##############################################################

 Write-output "Cleaning up remote session."
 Get-PSSession | Remove-PSSession