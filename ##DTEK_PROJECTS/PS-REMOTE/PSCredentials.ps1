Set-StrictMode -Version Latest

$PASSWORD = "mypassword"


# create secure string from plain-text string
$secureString = ConvertTo-SecureString -AsPlainText -Force -String $PASSWORD
Write-Host "Secure string:",$secureString
Write-Host

# convert secure string to encrypted string (for safe-ish storage to config/file/etc.)
$encryptedString = ConvertFrom-SecureString -SecureString $secureString
Write-Host "Encrypted string:",$encryptedString
Write-Host

# convert encrypted string back to secure string
$secureString = ConvertTo-SecureString -String $encryptedString
Write-Host "Secure string:",$secureString
Write-Host

# use secure string to create credential object
$credential = New-Object `
	-TypeName System.Management.Automation.PSCredential `
	-ArgumentList "myusername",$secureString

Write-Host "Credential:",$credential



################################################################################


#Exporting SecureString from Plain text with Out-File
"Kerrisue1" | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Out-File "C:\secure\Danny.txt"



#Exporting SecureString from Get-Credential
(Get-Credential).Password | ConvertFrom-SecureString | Out-File "C:\Temp 2\ScorchRA.txt"



#Exporting SecureString from Read-Host
Read-Host "Enter Password" -AsSecureString |  ConvertFrom-SecureString | Out-File "C:\Temp 2\Password.txt"




#Creating SecureString object with Get-Content and ConvertTo-SecureString
$pass = Get-Content "C:\Temp 2\Password.txt" | ConvertTo-SecureString


#Creating PSCredential object
$User = "DTEK\SVC-ORCH2016-RS"
$File = "C:\secure\ScorchRA.txt"
$MyCredential=New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, (Get-Content $File | ConvertTo-SecureString)