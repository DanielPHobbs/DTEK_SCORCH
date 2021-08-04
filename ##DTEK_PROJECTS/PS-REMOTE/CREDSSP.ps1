$computername="dtekaz-hw01.dtek.com"

    #$User = "DTEK\SVC-ORCH2016-RS"
    #$File = "C:\secure\ScorchRA.txt"
    #$PSCredential=New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, (Get-Content $File | ConvertTo-SecureString)

    $User = "DTEK\danny"
    $File = "C:\secure\danny.txt"
    $PSCredential=New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, (Get-Content $File | ConvertTo-SecureString)



    $LogContent=Invoke-Command -ComputerName $computername  -ScriptBlock { Get-Content F:\inetpub\logs\logfiles\W3SVC1\u_ex210414.log } -Credential $PSCredential -Authentication Credssp
    $LogContent

#WinRM get winrm/config/client

<#
start | run “gpedit.msc”
drill down to Computer-configuration | administrative templates | system | credential delegation
double click on “Allow Delgating Fresh Credentials with NTLM-only server authentication”
Enable this option
click on the show button
add in “WSMAN/*.dtek.co.uk”
#>

### client

#Set-Item WSMAN:\localhost\service\auth\credssp –value $true

### Runbook Servers

#Set-Item WSMAN:\localhost\client\auth\credssp –value $true


#enable-psremoting -force

#Enable-WSManCredSSP -Role "Client" -DelegateComputer "*.dtek.com"