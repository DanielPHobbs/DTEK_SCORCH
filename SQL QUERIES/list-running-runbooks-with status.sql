--Just running

SELECT *
FROM [Orchestrator].[dbo].[POLICY_PUBLISH_QUEUE] ppq
LEFT JOIN Orchestrator.dbo.POLICYINSTANCES pin on pin.PolicyID = ppq.PolicyID

--with date and status check

use Orchestrator2016

Select Name, TimeStarted, TimeEnded, POLICYINSTANCES.Status
From [Microsoft.SystemCenter.Orchestrator.Runtime].Jobs AS Jobs
 INNER JOIN POLICIES ON Jobs.RunbookId = POLICIES.UniqueID
 inner join POLICYINSTANCES on jobs.Id = POLICYINSTANCES.JobId 
 where POLICYINSTANCES.Status = 'success' 
 
 --and  TimeEnded > dateadd(HOUR, -600, getdate())
 and  TimeEnded > dateadd(DAY, -2, getdate())
 
 order by TimeStarted