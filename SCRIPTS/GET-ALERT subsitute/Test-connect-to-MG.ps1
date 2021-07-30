#========================================================================
# Created with: SAPIEN Technologies, Inc., PowerShell Studio 2012 v3.1.35
# Created on:   4/22/2016 1:08 PM
# Created by:   xdanielhobbs
# Organization: DTKEITINNOVATIONS
# Filename:     
#========================================================================



$me = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
 
$RunAsCred = get-Credential ""
 
$RMS =  ""
 
$AlertID = "8a5e738c-53d5-4a04-91af-cb8bc2d2e5d3"
 
$NewSession = new-pssession -ComputerName $env:COMPUTERNAME -Authentication Credssp -Credential (Get-Credential $me)
 
$alert = invoke-command  -session $NewSession -ScriptBlock {
 
param($RMS,$AlertID,$RunAsCred)
 
Add-PSSnapin Microsoft.EnterpriseManagement.OperationsManager.Client
 
New-PSDrive -Name:Monitoring -PSProvider:OperationsManagerMonitoring -Root:\
 
Set-Location "OperationsManagerMonitoring::"
 
new-managementGroupConnection -ConnectionString:$RMS -credential $RunAsCred | Out-Null
 
Set-Location $RMS
 
$Alert = Get-Alert -Id $AlertID
 
$Alert
 
} -ArgumentList $RMS, $AlertID, $RunAsCred
 
Remove-PSSession $NewSession
 
$alert