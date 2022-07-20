$ProtocolList       = @("SSL 2.0","SSL 3.0","TLS 1.0", "TLS 1.1", "TLS 1.2")
    $ProtocolSubKeyList = @("Client", "Server")
    $DisabledByDefault = "DisabledByDefault"
    $Enabled = "Enabled"
    $registryPath = "HKLM:\\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\"


    #https://docs.microsoft.com/en-us/system-center/orchestrator/integration-pack-for-operations-manager?view=sc-orch-2022

    
    foreach($Protocol in $ProtocolList)
    {
      Write-Host " In 1st For loop"
      foreach($key in $ProtocolSubKeyList)
      {		
          $currentRegPath = $registryPath + $Protocol + "\" + $key
          Write-Host " Current Registry Path $currentRegPath"

          if(!(Test-Path $currentRegPath))
          {
              Write-Host "creating the registry"
              New-Item -Path $currentRegPath -Force | out-Null			
          }
          if($Protocol -eq "TLS 1.2")
          {
              Write-Host "Working for TLS 1.2"
              New-ItemProperty -Path $currentRegPath -Name $DisabledByDefault -Value "0" -PropertyType DWORD -Force | Out-Null
              New-ItemProperty -Path $currentRegPath -Name $Enabled -Value "1" -PropertyType DWORD -Force | Out-Null

           }
          else
           {
               Write-Host "Working for other protocol"
               New-ItemProperty -Path $currentRegPath -Name $DisabledByDefault -Value "1" -PropertyType DWORD -Force | Out-Null
               New-ItemProperty -Path $currentRegPath -Name $Enabled -Value "0" -PropertyType DWORD -Force | Out-Null
           }
       }
     }

     Exit 0