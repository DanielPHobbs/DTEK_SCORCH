##credssp

#Enable-WSManCredSSP -Role Client -DelegateComputer [dtekadmin3.dtek.com] -Force

#Enable-WSManCredSSP -Role Server –Force




Invoke-Command -ComputerName "dtekad05.dtek.com" -ScriptBlock {
    
    Import-Module ActiveDirectory
    $UserProperties=@()
    $UserProperties=Get-ADUser -Identity "dhobbs-adm"  -Properties *

    $UserCanonicalName=$UserProperties.CanonicalName
    $Useraddress=$UserProperties.StreetAddress
 }

 <#
$credential = Get-Credential -Credential dtek\danny
$session = New-PSSession -cn dtekad05.dtek.com -Credential $credential -Authentication Credssp
Invoke-Command -Session $session -ScriptBlock { Import-Module ActiveDirectory; Get-ADUser 'dhobbs-adm' }

-Server "dtekad05.dtek.com"



#>

#get-WSManCredSSP
