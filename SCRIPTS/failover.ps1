Param(
  [string]$activeEnvironmentName,
  [string]$activeEnvironmentPort,
  [string]$failoverEnvironmentName,
  [string]$failoverEnvironmentPort,
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
$failoverRunbook = { 
Param($rb, $activeWS, $failoverWS, $activeEnvironmentName, $failoverEnvironmentName)
	try
	{
		$modulePath = "C:\Program Files\SCOrchDev\Modules"
		if(-not($Env:PSModulePath -like "*$modulePath*"))
		{
			$Env:PSModulePath += ";$modulePath"
		}
		Import-Module scorch
		Write-Host "------------------------------------------------------------------------------------------------------------------------"
		Write-Host "Starting work on"  $rb.Path
		$rb = Get-SCORunbook $activeWS -RunbookGUID $rb.Id
		Write-Host ""
		#Stop Job(s)
		$job = $rb | Get-SCOJob $activeWS -jobStatus "Running" -LoadJobDetails
		$job
		#Stop jobs if there are none running
		
		if(($job | Measure-Object).count -gt 0)
		{
			#If there are more than 1 jobs running wait for all but the 'monitor' to finish
			$givenMessage = $false
			while($job.ActiveInstances -gt 1)
			{
				if($givenMessage)
				{
					Write-Host -NoNewline .
				}
				else
				{
					Write-Host "Waiting for jobs to complete in $activeEnvironmentName"
					Write-Host "Instance Count: "  $job.ActiveInstances
					$givenMessage = $true
				}
				#sleep -Seconds 5
				$job = $rb | Get-SCOJob $activeWS -jobStatus "Running" -LoadJobDetails
			}
			if($givenMessage) { Write-Host "" }
			
			Write-Host "Stopping in $activeEnvironmentName"
			#Stop the job			
			$job | Stop-SCOJob $activeWS | Out-Null
			$jobStatus = ($job | Get-SCOJob $activeWS).job.Status
			Write-Host -NoNewline .
			#wait for job to stop
			while($jobStatus -eq "Running")
			{
				#update Job Status
				$job = $rb | Get-SCOJob $activeWS -jobStatus "Running" -LoadJobDetails
				$jobStatus = ($job | Get-SCOJob $activeWS).job.Status
				if($jobStatus -eq "Running") 
				{ 
					#if job Status is still running try re-sending a stop
					$job | Stop-SCOJob $activeWS
					$job = $rb | Get-SCOJob $activeWS -jobStatus "Running" -LoadJobDetails
					$jobStatus = ($job | Get-SCOJob $activeWS).job.Status
				}
				Write-Host -NoNewline .				
			}
		}
		else
		{
			Write-Host "Stopping in $activeEnvironmentName"
			Write-Host "Already stopped on $activeEnvironmentName"
		}
		Write-Host ""
		
		#Start Job in failover
		Write-Host "Starting in $failoverEnvironmentName"
		
		#check for job already running
		$failoverRB = Get-SCORunbook $failoverWS -RunbookPath $rb.Path
		$newJob = $failoverRB | Get-SCOJob $failoverWS -jobStatus "Running"
		
		#job is already running
		if(($newJob | Measure-Object).count -gt 0)
		{
			Write-Host "Already running on $failoverEnvironmentName"
		}
		else
		{
			#start the job for the runbook by path
			$newJob = $failoverRB | Start-SCORunbook $failoverWS
			$rbRunningJobs = $failoverRB | Get-SCOJob $failoverWS -jobStatus "Running"
			
			#wait for the job to start
			while(($rbRunningJobs | Measure-Object).count -gt 0)
			{
				Write-Host -NoNewline .
				$rbRunningJobs = $failoverRB | Get-SCOJob $failoverWS -jobStatus "Running"
			}
		}
		Write-Host ""
		Write-Host "------------------------------------------------------------------------------------------------------------------------"
	}
	catch { throw }
}

$activeWS = New-SCOWebserverURL $activeEnvironmentName $activeEnvironmentPort
Write-Host "Active Environment:   $activeWS"
$failoverWS = New-SCOWebserverURL $failoverEnvironmentName $failoverEnvironmentPort
Write-Host "Failover Environment: $failoverWS"

Write-Host ""
Write-Host "Loading $activeEnvironmentName Base Folders"
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

Write-Host ""
Write-Host "Loading $failoverEnvironmentName Base Folders"
$failoverBaseFolders = Get-SCOSubFolder $failoverWS $baseFolderPath $true
$failoverRBS = @()
foreach($runbook in $failoverBaseFolders)
{
	$failoverPotentialAdd = $runbook.Runbooks | ? {$_.IsMonitor -eq $true}
	if($failoverPotentialAdd -ne $null)
	{
		$failoverRBS += $runbook.Runbooks | ? {$_.IsMonitor -eq $true}
	}
}
$failoverPaths = $failoverRBS | select Path;
$failoverPaths | ft Path

$misMatch = Compare-Object -ReferenceObject $activePaths -DifferenceObject $failoverPaths

if($misMatch)
{
      Write-Host "Monitor Runbooks not idential between target environments"
      Write-Host $misMatch
}
else
{
	Write-Host "------------------------------------------------------------------------------------------------------------------------"
	Write-Host "Starting Jobs"
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
		
		Write-Host "Starting" $rb.Path
		$j = Start-Job -ArgumentList @($rb, $activeWS, $failoverWS, $activeEnvironmentName, $failoverEnvironmentName) -ScriptBlock $failoverRunbook -Name $rb.path
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
}