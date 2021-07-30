$basePath = "C:\Program Files\SCOrchDev\Modules"
if(-not(Test-Path $basePath))
{
	New-Item -ItemType directory -Path $basePath
}

Copy-Item -Recurse -Force -Confirm:$False -Path ".\scorch" -Destination $basePath
$machinePSModulePath = [Environment]::GetEnvironmentVariable('PSModulePath','Machine')
if (@($machinePSModulePath -split ';') -notcontains $basePath)
{
	# Add the module base path to the machine environment variable
	$machinePSModulePath += ";${basePath}"
	# Add the module base path to the process environment variable
	$env:PSModulePath += ";${basePath}"
	# Update the machine environment variable value on the local computer
	[Environment]::SetEnvironmentVariable('PSModulePath', $machinePSModulePath, 'Machine')
}
Import-Module -Name Scorch -PassThru