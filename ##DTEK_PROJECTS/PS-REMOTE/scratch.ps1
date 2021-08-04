$computername="dtekaz-hw01.dtek.com"
        
    
##############################################################
#Create PSCredentials
$User = "dtek\danny"
$File = "C:\secure\danny.txt"
$PSCredential=New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, (Get-Content $File | ConvertTo-SecureString)


#Enter-PSSession -ComputerName $computername -Credential $PSCredential -Authentication Kerberos
