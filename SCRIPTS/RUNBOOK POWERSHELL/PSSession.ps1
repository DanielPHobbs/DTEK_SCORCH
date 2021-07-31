$Session = New-PSSession -ComputerName localhost
$ReturnArray = Invoke-Command -Session $Session -ScriptBlock
{
  $username = 'pureuser'
  $pwd = ConvertTo-SecureString -String 'pureuser' -AsPlainText -Force
  $Creds = New-Object System.Management.Automation.PSCredential ($username, $pwd)
  
  
  $FlashArray = New-PfaArray -EndPoint 10.1.1.10 -Credentials $Creds –IgnoreCertificateError
  New-Item -Path 'C:\Temp\Orchestrator-Log.txt' -ItemType 'File'
  New-PfaVolumeSnapshots -Array $FlashArray -Sources “Volume1" -Suffix “Suffix1" | Add-Content -Path 'C:\Temp\Orchestrator-Log.txt'
}