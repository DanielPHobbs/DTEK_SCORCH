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