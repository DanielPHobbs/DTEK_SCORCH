        $computername="dtekaz-hw01.dtek.com"
        
        Install-Module -Name Invoke-PSSession

        
         ##############################################################
         #Create PSCredentials
         $User = "DTEK\SVC-ORCH2016-RS"
         $File = "C:\secure\ScorchRA.txt"
         $PSCredential=New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, (Get-Content $File | ConvertTo-SecureString)

         #############################################################
         write-host "Gathering log file(s)"

        $session = Invoke-PSSession -ComputerName $computername -Credential $PSCredential 
        
        Invoke-Command -Session $Session -ScriptBlock { Get-Content F:\inetpub\logs\logfiles\W3SVC1\u_ex210414.log } 
         

         