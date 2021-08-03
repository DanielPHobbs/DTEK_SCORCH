Clear-Host
$Machine = "dtekaz-hw01.dtek.com"
Get-Eventlog -List -ComputerName $Machine

winrm help config


Install-Module -Name Invoke-PSSession
Install-Module -Name Invoke-CommandAs

get-help Invoke-PSSession -Full
get-help Invoke-CommandAs -Full


<#
#https://blogs.msdn.microsoft.com/sergey_babkins_blog/2015/03/18/another-solution-to-multi-hop-powershell-remoting/

#https://docs.microsoft.com/en-gb/archive/blogs/sergey_babkins_blog/setting-up-the-credssp-access-for-multi-hop


$null = Enable-WSManCredSSP -Role Server -Force

$null = Enable-WSManCredSSP -Role Client -DelegateComputer "*" -Force
$null = mkdir -Force "HKLM:\Software\Policies\Microsoft\Windows\CredentialsDelegation\AllowFreshCredentials"
Set-ItemProperty -LiteralPath "HKLM:\Software\Policies\Microsoft\Windows\CredentialsDelegation\AllowFreshCredentials" -Name "my" -Value "wsman/*" -Type STRING
$null = mkdir -Force "HKLM:\Software\Policies\Microsoft\Windows\CredentialsDelegation\AllowFreshCredentialsWhenNTLMOnly"
Set-ItemProperty -LiteralPath "HKLM:\Software\Policies\Microsoft\Windows\CredentialsDelegation\AllowFreshCredentialsWhenNTLMOnly" -Name "my" -Value "*" -Type STRING

#>