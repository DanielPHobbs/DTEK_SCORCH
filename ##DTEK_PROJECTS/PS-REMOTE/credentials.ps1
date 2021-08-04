
$Global:sesssion = ""
$Global:server="dtekaz-hw01.dtek.com"

		$Username = & whoami
		#$password= [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
		#$password = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Password for $Username")
        #$cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $password		
		$Cred = Get-Credential -UserName $Username -Message 'Enter Password'
	#$Global:session = New-PSSession -ComputerName $Global:server -Credential $cred
    #Invoke-Command -Session $Global:session -ScriptBlock { Get-Content F:\inetpub\logs\logfiles\W3SVC1\u_ex210414.log -top 5} 

    Invoke-Command -computername $Global:server - credential $cred -ScriptBlock { Get-Content F:\inetpub\logs\logfiles\W3SVC1\u_ex210414.log -top 5} 