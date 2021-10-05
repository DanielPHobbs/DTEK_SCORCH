Import-Module ActiveDirectory
$DataBusInput1="dhobbs-adm"
$UserProperties=Get-ADUser -Identity $DataBusInput1  -Properties *
$UserProperties


