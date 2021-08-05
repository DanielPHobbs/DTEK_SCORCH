$Output = '@{Name=IN-VMM02; Status=Running; HostName=in-vmm.domain1.net; Owner=domain1\admin; OperatingSystem=Windows Server 2012 R2 Standard; PSComputerName=in-vm07.domain1.net; RunspaceId=b0e8adae-1831-4ab9-85bb-9f33b1373fef; PSShowComputerName=True},@{Name=IN-VMM01; Status=Running; HostName=in-vmm.domain1.net; Owner=domain1\admin; OperatingSystem=64-bit edition of Windows 7; PSComputerName=in-vm07.domain1.net; RunspaceId=b0e8adae-1831-4ab9-85bb-9f33b1373fef; PSShowComputerName=True}'
$Result = @()
FOREACH ($item in $Output.Split(','))
{
$Result += $item
}




$Output = '@{Name=IN-VMM02; Status=Running; HostName=in-vmm.domain1.net; Owner=domain1\admin; OperatingSystem=Windows Server 2012 R2 Standard; PSComputerName=in-vm07.domain1.net; RunspaceId=b0e8adae-1831-4ab9-85bb-9f33b1373fef; PSShowComputerName=True},@{Name=IN-VMM01; Status=Running; HostName=in-vmm.domain1.net; Owner=domain1\admin; OperatingSystem=64-bit edition of Windows 7; PSComputerName=in-vm07.domain1.net; RunspaceId=b0e8adae-1831-4ab9-85bb-9f33b1373fef; PSShowComputerName=True}'
$Name = @()
$Status = @()
$HostName = @()
FOREACH ($item in $Output.Split(','))
{
    $Name += (($item.Split(';')[0]).Split('='))[1]
    $Status += (($item.Split(';')[1]).Split('='))[1]
    $HostName += (($item.Split(';')[2]).Split('='))[1]
}



#https://docs.microsoft.com/en-gb/archive/blogs/privatecloud/automationorchestrator-tiptrickrun-net-script-activity-powershell-published-data-arrays



#The results of this example will look very similar to the “Flatten” functionality, but instead of using “Flatten” the delimitation is handled from within the PowerShell.
 $DelimProcessList = ""
 $Processes = Get-Process 
 foreach ($Process in $Processes) { $DelimProcessList += $Process.Name + ";" } 
 $DelimProcessList = $DelimProcessList.Substring(0,$DelimProcessList.Length-1) 
 $DelimProcessList

#The results of this example will generate Multi-Value Published Data. This means that all subsequent (downstream) Runbook Activities will consume each piece of data individually. If this is not the desired output, you can use the example above or leverage the built in “Flatten” functionality for this example’s Activity.
 $ArrayProcessList = @() 
 $Processes = Get-Process #add a where
 foreach ($Process in $Processes) { $ArrayProcessList += $Process.Name } 
 $ArrayProcessList

 <#
 Variable Declaration
Delimited List Example:

VARIABLE DECLARATION
001 002 003	$NodeID = "" $NodeName = "" $NodeState = ""
Array List Example:

VARIABLE DECLARATION
001 002 003	$NodeID = @() $NodeName = @() $NodeState = @()
 
Handling the Different Variable Data
Delimited List Example:

HANDLING THE DIFFERENT VARIABLE DATA
001 002 003 004 005 006	foreach ($Process in $Processes) { $DelimProcessList += $Process.Name + ";" } $DelimProcessList = $DelimProcessList.Substring(0,$DelimProcessList.Length-1)
Array List Example:

HANDLING THE DIFFERENT VARIABLE DATA
001 002 003 004	foreach ($Process in $Processes) { $ArrayProcessList += $Process.Name }
 #>