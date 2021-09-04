#Function
Function New-CustomEvent {
    [CmdletBinding()]
            
    # Parameters used in this function
    param
    ( 
        [Parameter(Position=0, Mandatory = $false, HelpMessage="Provide eventlog name", ValueFromPipeline = $true)] $EventLog  = "Orchestrator",
        [Parameter(Position=1, Mandatory = $false, HelpMessage="Provide event source", ValueFromPipeline = $true)]  $Source    = "Scorch",
        [Parameter(Position=2, Mandatory = $false, HelpMessage="Provide event source", ValueFromPipeline = $true)]  $EventID   = "106",
        [Parameter(Position=3, Mandatory = $true, HelpMessage="Provide event message", ValueFromPipeline = $false)] $Message,
        [Parameter(Position=4, Mandatory = $false, HelpMessage="Select event instance", ValueFromPipeline = $false)]
        [ValidateSet("Information","Warning","Error")] $EventInstance = 'Error'
    ) 
 
    #Load the event source
    If ([System.Diagnostics.EventLog]::SourceExists($Source) -eq $false) {[System.Diagnostics.EventLog]::CreateEventSource($Source, $EventLog)}


    Switch ($EventInstance){
        {$_ -match 'Error'}       {$id = New-Object System.Diagnostics.EventInstance($EventID,1,1)} #ERROR EVENT
        {$_ -match 'Warning'}     {$id = New-Object System.Diagnostics.EventInstance($EventID,1,2)} #WARNING EVENT
        {$_ -match 'Information'} {$id = New-Object System.Diagnostics.EventInstance($EventID,1)}   #INFORMATION EVENT
    }

    $Object = New-Object System.Diagnostics.EventLog;
    $Object.Log       = $EventLog;
    $Object.Source    = $Source;

    $Object.WriteEvent($id, @($Message))

}

$now=get-Date


$message ="Created New Event log for Orchestrator on $now" 
New-CustomEvent -Message $Message




#################################################################

#Events criteria
$Filter = @{
LogName      = 'orchestrator'
ProviderName = 'scorch'
ID           = 106
}

Get-WinEvent $Filter -MaxEvents 10 | select LogName,ProviderName,ID,TimeCreated,message