# Import the SCORCH module
Import-Module 'X:\SCORCH_SCRIPTS\SCOrchestratorServicePowerShellV1_2\OrchestratorServiceModule.psm1'

#$creds = Get-Credential "DOMAIN\USERNAME"

# Set URL for web service
$scorh_url = "http://s2012-app1:81/orchestrator2012/orchestrator.svc/"

# Set Folder Path for Runbook Folder
$runbook_folder = '\TFL'

# Get all runbooks within the  folder
$daily_runbooks = Get-OrchestratorRunbook -ServiceUrl $scorh_url -RunbookPath $runbook_folder

# Run through them and start each one
foreach ($runbook in $daily_runbooks)
{

$runbook_ID= $runbook | % {$_.id}
$runbook_Name= $runbook | % {$_.name}

write-host   $runbook_ID, $runbook_Name



	
	# Get the runbook object
$runbook_temp = Get-OrchestratorRunbook -ServiceUrl $scorh_url -RunbookId $runbook.Id 
foreach ($rb in $runbook_temp)

{
$rb
}

}



    
#Is it running if YES thn OK else start it  try this a few time then exclude with logged error
    
    
# Now start the runbook
#Start-OrchestratorRunbook -Runbook $runbook_temp



# Output error log
#$error | out-file "D:\logs\error_DailyRunbooks.txt"

# Output identity
#whoami | out-file "D:\logs\identity_DailyRunbooks.txt" 
