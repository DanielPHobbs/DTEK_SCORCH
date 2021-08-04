Get-EventLog -List


$servers = Get-TransportService;
foreach ($server in $servers)
{Write-Host "Scanning the event log of: " -NoNewLine; Write-Host $server;
Get-EventLog system -ComputerName $server -After (Get-Date).AddHours(-12) | where {($_.EntryType -Match "Error") -or ($_.EntryType -Match "Warning")} | ft  -wrap >> "C:/$server.csv";
Get-EventLog application -ComputerName $server -After (Get-Date).AddHours(-12) | where {($_.EntryType -Match "Error") -or ($_.EntryType -Match "Warning")} | ft  -wrap >> "C:/$server.csv"}


Get-EventLog application -newest 1 | Get-Member

Get-WinEvent -ListLog * | where {$_.RecordCount -gt 0}

Get-EventLog system -after (get-date).AddDays(-1) | where {$_.InstanceId -eq 7001}

$today = get-date -Hour 0 -Minute 0;
Get-EventLog system -after $today | sort -Descending | select -First 1

$logs = get-eventlog system -ComputerName <name of the monitored computer> -source Microsoft-Windows-Winlogon -After (Get-Date).AddDays(-7);
$res = @(); ForEach ($log in $logs) {if($log.instanceid -eq 7001) {$type = "Logon"} Elseif ($log.instanceid -eq 7002){$type="Logoff"} Else {Continue} $res += New-Object PSObject -Property @{Time = $log.TimeWritten; "Event" = $type; User = (New-Object System.Security.Principal.SecurityIdentifier $Log.ReplacementStrings[1]).Translate([System.Security.Principal.NTAccount])}};
$res

$DateAfter = (Get-Date).AddDays(-1)
$DateBefore = (Get-Date)
$EventLogTest = Get-EventLog -LogName Security -InstanceId 4625 -Before $DateBefore -After $DateAfter -Newest 5
$WinEventTest = Get-WinEvent -FilterHashtable @{ LogName = 'Security'; Id = 4625; StartTime = $DateAfter; EndTime = $DateBefore } -MaxEvents 5




$FilterHashTable = @{
    LogName   = 'Security'
    ProviderName= 'Microsoft-Windows-Security-Auditing' 
    #Path = <String[]>
    #Keywords = <Long[]>
    ID        = 4625
    #Level = <Int32[]>
    StartTime = (Get-Date).AddDays(-1)
    EndTime   = Get-Date
    #UserID = <SID>
    #Data = <String[]>
}
Get-WinEvent -FilterHashtable $FilterHashTable -MaxEvents 5
Get-EventLog -LogName 'Security' -Source 'Microsoft-Windows-Security-Auditing' -Newest 5




Write-Color 'Scanning Event Log with Get-EventLog' -Color Blue
$Time2 = Start-TimeLog
$Event2 = Get-EventLog -LogName 'Security' -Source 'Microsoft-Windows-Security-Auditing' -InstanceId 4625 -Before (Get-Date) -After ((Get-Date).AddDays(-1))| Where-Object { $_.Index -eq '4125545' }
Stop-TimeLog -Time $Time2
Write-Color 'Scanning Event Log with Get-WinEvent' -Color Green
$Time = Start-TimeLog
$FilterHashTable = @{
    LogName      = 'Security'
    ProviderName = 'Microsoft-Windows-Security-Auditing' 
    #Path = <String[]>
    #Keywords = <Long[]>
    ID           = 4625
    #Level = <Int32[]>
    StartTime    = (Get-Date).AddDays(-1)
    EndTime      = Get-Date
    #UserID = <SID>
    #Data = <String[]>
}
$Event1 = Get-WinEvent -FilterHashtable $FilterHashTable | Where-Object { $_.RecordID -eq '4125545'}
Stop-TimeLog -Time $Time

$WinEventTest = Get-WinEvent -FilterHashtable @{ LogName = 'Security'; Id = 4625; StartTime = $DateAfter; EndTime = $DateBefore } -MaxEvents 5
$WinEventTest[0] |Format-List *
$WinEventTest[0].Properties




$FilterHashTable = @{
    LogName   = 'Application'
    ID        = 1534
    StartTime = (Get-Date).AddHours(-1)
    EndTime   = Get-Date
}
$Events = Get-WinEvent -FilterHashtable $FilterHashTable | ForEach-Object {
    $Values = $_.Properties | ForEach-Object { $_.Value }
    
    # return a new object with the required information
    [PSCustomObject]@{
        Time      = $_.TimeCreated
        # index 0 contains the name of the update
        Event     = $Values[0]
        Component = $Values[1]
        Error     = $Values[2]
        User      = $_.UserId.Value
    }
}
$Events | Format-Table -AutoSize




$EventLog = Get-EventLog -LogName 'Application' -After ((Get-Date).AddHours(-1)) -InstanceId 1534 | ForEach-Object {   
    # return a new object with the required information
    [PSCustomObject]@{
        Time      = $_.TimeGenerated
        # index 0 contains the name of the update
        Event     = $_.ReplacementStrings[0]
        Component = $_.ReplacementStrings[1]
        Error     = $_.ReplacementStrings[2]
        User      = $_.UserName
    }
}
$EventLog | Format-Table -AutoSize



$FilterHashTable = @{
    LogName   = 'Application'
    ID        = 903
    StartTime = (Get-Date).AddDays(-1)
    EndTime   = Get-Date
}
$ComputerName = 'AD1', 'AD2'
$Event1 = foreach ($Computer in $ComputerName) {
    Get-WinEvent -FilterHashtable $FilterHashTable -ComputerName $Computer
}
