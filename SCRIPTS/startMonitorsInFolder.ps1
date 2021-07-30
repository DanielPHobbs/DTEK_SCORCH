Param(
  [string]$activeEnvironmentName,
  [string]$activeEnvironmentPort,
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
Param($rb, $activeWS, $activeEnvironmentName)
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
		$newJob = $rb | Get-SCOJob $activeWS -jobStatus "Running"
		
		#job is already running
		if(($newJob | Measure-Object).count -gt 0)
		{
			Write-Host "Already running in $activeEnvironmentName"
		}
		else
		{
			Write-Host "Monitor was not running: Creating new Job"
			
			#start the job for the runbook by path
			$newJob = $rb | Start-SCORunbook $activeWS
			$rbRunningJobs = $rb | Get-SCOJob $activeWS -jobStatus "Running"
			
			#wait for the job to start
			while(($rbRunningJobs | Measure-Object).count -gt 0)
			{
				Write-Host -NoNewline .
				$rbRunningJobs = $rb | Get-SCOJob $activeWS -jobStatus "Running"
			}
		}
		Write-Host ""
		Write-Host "------------------------------------------------------------------------------------------------------------------------"
	}
	catch { throw }
}

$activeWS = New-SCOWebserverURL $activeEnvironmentName $activeEnvironmentPort
Write-Host "Active Environment:   $activeWS"

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
	
	Write-Host "Starting" $rb.Path
	$j = Start-Job -ArgumentList @($rb, $activeWS, $activeEnvironmentName) -ScriptBlock $checkRunbook -Name $rb.path
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