Param(
  [string]$EnvironmentName,
  [string]$EnvironmentPort,
  [string]$baseFolderPath,
  [switch]$waitForComplete
)
$modulePath = "C:\Program Files\SCOrchDev\Modules"
if(-not($Env:PSModulePath -like "*$modulePath*"))
{
	$Env:PSModulePath += ";$modulePath"
}
Import-Module scorch
$MaxConcurrency = 10
$checkRunbook = { 
Param($rb, $activeWS, $EnvironmentName)
	try
	{
		$modulePath = "C:\Program Files\SCOrchDev\Modules"
		if(-not($Env:PSModulePath -like "*$modulePath*"))
		{
			$Env:PSModulePath += ";$modulePath"
		}
		Import-Module scorch
		Write-Host "------------------------------------------------------------------------------------------------------------------------"
		Write-Host "Starting work on"  
		$rb.Path
		$rb = Get-SCORunbook $activeWS -RunbookGUID $rb.Id
		Write-Host ""
		
		#check for monitor running
		$monJob = $rb | Get-SCOJob $activeWS -jobStatus "Running" -LoadJobDetails
		
		#Monitor is running
		if(($monJob | Measure-Object).count -gt 0)
		{
			Write-Host "Monitor running in $EnvironmentName"

			#If there are more than 1 jobs running wait for all but the 'monitor' to finish
			$givenMessage = $false
			while($monJob.ActiveInstances -gt 1)
			{
				if($givenMessage)
				{
					Write-Host -NoNewline .
				}
				else
				{
					Write-Host "Waiting for jobs to complete in $EnvironmentName"
					Write-Host "Instance Count: "  $monJob.ActiveInstances
					$givenMessage = $true
				}
				#sleep -Seconds 5
				$monJob = $rb | Get-SCOJob $activeWS -jobStatus "Running" -LoadJobDetails
			}
			if($givenMessage) { Write-Host "" }
			
			Write-Host "Stopping in $EnvironmentName"
			
			#Stop the job			
			$monJob | Stop-SCOJob $activeWS | Out-Null
			$monJob = $rb | Get-SCOJob $activeWS -jobStatus "Running"
			
			#wait for job to stop
			while(($monJob | Measure-Object).count -gt 0)
			{
				#update Job Status
				Write-Host -NoNewline .
				$monJob | Stop-SCOJob $activeWS
				$monJob = $rb | Get-SCOJob $activeWS -jobStatus "Running" -LoadJobDetails
			}
		}
		else
		{
			Write-Host "Monitor not running"
		}
			
		Write-Host ""
		Write-Host "------------------------------------------------------------------------------------------------------------------------"
	}
	catch { throw }
}

$activeWS = New-SCOWebserverURL $EnvironmentName $EnvironmentPort
Write-Host "Active Environment:   $activeWS"

Write-Host ""
Write-Host "Loading $EnvironmentName Base Folders"
$activeBaseFolders = Get-SCOSubFolder $activeWS $baseFolderPath $true
$activeRBS = @()
foreach($runbook in $activeBaseFolders)
{
	$activePotentialAdd = $runbook.Runbooks | ? {$_.IsMonitor -eq $true}
	if($activePotentialAdd -ne $null)
	{
		$activeRBS += $runbook.Runbooks | ? {$_.IsMonitor -eq $true}
	}
}
$activePaths = $activeRBS | select Path;
$activePaths | ft Path

Write-Host "------------------------------------------------------------------------------------------------------------------------"
Write-Host "Checking Monitor"
foreach($rb in $activeRBS)
{
	$waited = $false
	if((Get-Job -State Running).Count -ge $MaxConcurrency) 
	{ 
		$waited = $true
		Write-Host "Max Concurrent Jobs Reached: Waiting" 
	}
	while((Get-Job -State Running).Count -ge $MaxConcurrency)
	{
		Write-Host -NoNewLine .
		sleep -Milliseconds 333
	}
	if($waited) { Write-Host "" }
	
	Write-Host "Starting work on" $rb.Path
	$j = Start-Job -ArgumentList @($rb, $activeWS, $EnvironmentName) -ScriptBlock $checkRunbook -Name $rb.path
	while($true)
	{
		$state = ($j | Get-Job).State
		if(($state -eq "Running") -or ($state -eq "Completed") -or ($state -eq "Failed"))
		{
			break
		}
	}
}
if($waitForComplete)
{
	$jArray = Get-Job
	Write-Host ""
	Write-Host "Working"
	while($true)
	{
		Write-Host -NoNewline .
		$finishedCount = ($jArray | Get-Job | ? {$_.State -eq "Completed"}).Count
		$finishedCount += ($jArray | Get-Job | ? {$_.State -eq "Failed"}).Count
		if($finishedCount -eq $jArray.Count) { break }
		sleep -Milliseconds 333
	}
	Write-Host ""
	Write-Host "------------------------------------------------------------------------------------------------------------------------"
	foreach($j in Get-Job -State Completed)
	{
		Receive-Job $j
		Remove-Job $j
	}
	Write-Host ""
	Get-Job
}
else 
{
	Write-Host ""
	Write-Host "------------------------------------------------------------------------------------------------------------------------"
}