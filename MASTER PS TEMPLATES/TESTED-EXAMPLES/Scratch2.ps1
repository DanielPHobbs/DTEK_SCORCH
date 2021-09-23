##credssp

#Enable-WSManCredSSP -Role Client -DelegateComputer [dtekadmin3.dtek.com] -Force

#Enable-WSManCredSSP -Role Server –Force


$DataBusInput1="dhobbs-adm"

Invoke-Command -ComputerName "dtekadmin3.dtek.com" -ScriptBlock {
    
    Import-Module ActiveDirectory
    
    Get-ADUser -Identity "dhobbs-adm" -Server "dtekad05.dtek.com"  -Properties *

 }

 <#
$credential = Get-Credential -Credential dtek\danny
$session = New-PSSession -cn dtekad05.dtek.com -Credential $credential -Authentication Credssp
Invoke-Command -Session $session -ScriptBlock { Import-Module ActiveDirectory; Get-ADUser 'dhobbs-adm' }

-Server "dtekad05.dtek.com"



#>

#get-WSManCredSSP
