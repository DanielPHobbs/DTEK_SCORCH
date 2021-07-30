Param(
  [string]$EnvironmentName = "localhost",
  [string]$EnvironmentPort = "81",
  [Parameter(Mandatory=$true)]
  $ActiveFolder,
  [Parameter(Mandatory=$true)]
  $DisabledFolder,
  [int]$MaxConcurrency = 10,
  [switch]$waitForComplete
)

#Ensure that the module path is loaded
$modulePath = "C:\Program Files\SCOrchDev\Modules"
if(-not($Env:PSModulePath -like "*$modulePath*"))
{
	#If the module path is not loaded, add it to the path
	$Env:PSModulePath += ";$modulePath"
}
#Import the scorch webservice module
Import-Module scorch

#Setup a script blog to pass to PowerShell Jobs
$checkRunbook = { 
Param($rb, $WS, $EnvironmentName, $shouldStart)
	try
	{		
		#Ensure that the module path is loaded
		$modulePath = "C:\Program Files\SCOrchDev\Modules"
		if(-not($Env:PSModulePath -like "*$modulePath*"))
		{
			#If the module path is not loaded, add it to the path
			$Env:PSModulePath += ";$modulePath"
		}
		#Import the scorch webservice module
		Import-Module scorch
		
		Write-Host "------------------------------------------------------------------------------------------------------------------------"
		if($shouldStart) { Write-Host "Starting" }
		else { Write-Host "Stopping" } 
		$rb.Path
		Write-Host ""
		
		#load the Runbook Object into this PowerShell instance
		$rb = Get-SCORunbook $WS -RunbookGUID $rb.Id
				
		#check for monitor running
		$monJob = $rb | Get-SCOJob $WS -jobStatus "Running"
		
		if($shouldStart)
		{
			#job is already running
			if(($monJob | Measure-Object).count -gt 0)
			{
				Write-Host "Already running in $EnvironmentName"
			}
			else
			{
				Write-Host "Monitor was not running: Creating new Job"
				
				#start the job for the runbook by path
				$monJob = $rb | Start-SCORunbook $WS
				$rbRunningJobs = $rb | Get-SCOJob $WS -jobStatus "Running"
				
				#wait for the job to start
				while(($rbRunningJobs | Measure-Object).count -gt 0)
				{
					Write-Host -NoNewline .
					$rbRunningJobs = $rb | Get-SCOJob $WS -jobStatus "Running"
				}
			}
		}
		else
		{
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
					$monJob = $rb | Get-SCOJob $WS -jobStatus "Running" -LoadJobDetails
				}
				if($givenMessage) { Write-Host "" }
				
				Write-Host "Stopping in $EnvironmentName"
				
				#Stop the job			
				$monJob | Stop-SCOJob $WS | Out-Null
				$monJob = $rb | Get-SCOJob $WS -jobStatus "Running"
				
				#wait for job to stop
				while(($monJob | Measure-Object).count -gt 0)
				{
					#update Job Status
					Write-Host -NoNewline .
					$monJob | Stop-SCOJob $WS
					$monJob = $rb | Get-SCOJob $WS -jobStatus "Running" -LoadJobDetails
				}
			}
			else
			{
				Write-Host "Monitor not running"
			}
		}
		Write-Host ""
		Write-Host "------------------------------------------------------------------------------------------------------------------------"
	}
	catch { throw }
}

#setup a variable to hold the webservice URL
$WS = New-SCOWebserverURL $EnvironmentName $EnvironmentPort

Write-Host "Environment:   $WS"
Write-Host ""
Write-Host "Loading $EnvironmentName Monitor Runbooks"

$MonitorRunbooks = Get-SCOMonitorRunbook $WS
$runbooksToStart = @()
$runbooksToStop = @()

foreach($runbook in $MonitorRunbooks)
{
	#Check the runbook for running jobs
	$job = $runbook | Get-SCOJob $WS -jobStatus Running
	
	if(($job | Measure-Object).Count -gt 0)
	{
		#The Runbook is Running
		#Check Path to see if it should be stopped or started
		
		$notFound = $true
		foreach($Path in $ActiveFolder)
		{
			if($runbook.Path.StartsWith($Path))
			{
				#Should be started and is
				$notFound = $false
				break
			}
		}
		if($notFound)
		{
			foreach($Path in $DisabledFolder)
			{
				if($runbook.Path.StartsWith($Path))
				{
					#Should be stopped and isn't
					$runbooksToStop += $runbook
					break
				}
			}
		}
	}
	else
	{
		#The Runbook is not Running
		#Check Path to see if it should be stopped or started
		
		$notFound = $true
		foreach($Path in $ActiveFolder)
		{
			if($runbook.Path.StartsWith($Path))
			{
				#Should be started and isn't
				$notFound = $false
				$runbooksToStart += $runbook
				break
			}
		}
		if($notFound)
		{
			foreach($Path in $DisabledFolder)
			{
				if($runbook.Path.StartsWith($Path))
				{
					#Should be stopped and is
					break
				}
			}
		}
	}
}

if($runbooksToStop.count -gt 0)
{
	Write-Host "Monitors to Stop"
	$runbooksToStop | ft Path
	Write-Host "------------------------------------------------------------------------------------------------------------------------"
	Write-Host "Stopping Monitors"
	foreach($rb in $runbooksToStop)
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
		
		Write-Host "Stopping" $rb.Path
		$j = Start-Job -ArgumentList @($rb, $WS, $EnvironmentName, $false) -ScriptBlock $checkRunbook -Name $rb.path
		while($true)
		{
			$state = ($j | Get-Job).State
			if(($state -eq "Running") -or ($state -eq "Completed") -or ($state -eq "Failed"))
			{
				break
			}
		}
	}
	Write-Host "------------------------------------------------------------------------------------------------------------------------"
}

if($runbooksToStart.count -gt 0)
{
	Write-Host "Monitors to Start"
	$runbooksToStart | ft Path
	Write-Host "------------------------------------------------------------------------------------------------------------------------"
	Write-Host "Starting Monitors"
	foreach($rb in $runbooksToStart)
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
		$j = Start-Job -ArgumentList @($rb, $WS, $EnvironmentName, $true) -ScriptBlock $checkRunbook -Name $rb.path
		while($true)
		{
			$state = ($j | Get-Job).State
			if(($state -eq "Running") -or ($state -eq "Completed") -or ($state -eq "Failed"))
			{
				break
			}
		}
	}
	Write-Host "------------------------------------------------------------------------------------------------------------------------"
}


if($waitForComplete)
{
	Write-Host "------------------------------------------------------------------------------------------------------------------------"
	Write-Host "Waiting for Jobs to Complete"
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