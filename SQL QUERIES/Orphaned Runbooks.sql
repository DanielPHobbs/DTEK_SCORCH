$Databaseserver='Server-SCO01';
$Orchestratordatabase='Orchestrator';
$Orchestratorservers="Server-SCO01;Server2"
$User="domain\scoservice";
$UserPassword="mygoodpassword";

[array]$Actionservers=$Orchestratorservers -split ";"

foreach ($Actionserver in $Actionservers)
    {
    if ($user -ne "")
        {
        $password = ConvertTo-SecureString  -AsPlainText $UserPassword -Force;
        $Credential = New-Object System.Management.Automation.PsCredential($user,$password);
        $session=New-PSSession -ComputerName $Actionserver -Credential $Credential;
        }
    else
        {
        $session=New-PSSession -ComputerName $Actionserver;
        }

    $updatesqlQuery=Invoke-Command -Session $session -scriptblock {
        $Databaseserver=$Args[0];
        $Orchestratordatabase=$Args[1];
        $Actionserver=$Args[2];
        $updatesqlQuery="";
        $SQLQueryrunningpolicies="SELECT Policies.[UniqueID],
                Policies.[ProcessID],
                POLICIES.JobId,
                Servertable.computer,
                policies.TimeStarted,
                Policies.Status,
                Policies.TimeEnded
                FROM [$Orchestratordatabase].[dbo].[POLICYINSTANCES] as Policies
                inner join [$Orchestratordatabase].[dbo].[ACTIONSERVERS] as Servertable on policies.ActionServer = Servertable.UniqueID
                where TimeEnded is null and Servertable.computer = '$Actionserver'
                order by Processid";
        $runningpolicies=Invoke-Sqlcmd -ServerInstance $Databaseserver -Database $Orchestratordatabase -Query $SQLQueryrunningpolicies;
        foreach ($policy in $runningpolicies)
            {
            Try
                {
                $process=Get-Process -id $policy.processid -ErrorAction Stop
                }
            Catch
                {
                $uniqueid=$policy.UniqueID;
                $jobid=$policy.JobId;
                $updatesqlQuery+="Update [$Orchestratordatabase].[dbo].[POLICYINSTANCES]
                    Set TimeEnded=GetDate(), Status='failed'
                    where Jobid='$jobid' and UniqueID='$uniqueid'"    +"`n";
                }
            }
        if ($updatesqlQuery -ne "")
            {
            Invoke-Sqlcmd -ServerInstance $Databaseserver -Database $Orchestratordatabase -Query $updatesqlQuery;
            }
        Return $updatesqlQuery
        } -ArgumentList $Databaseserver, $Orchestratordatabase, $Actionserver;
    Remove-PSSession $session;
    }