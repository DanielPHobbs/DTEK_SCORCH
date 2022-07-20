use Orchestrator

Select Name, TimeStarted, TimeEnded, POLICYINSTANCES.Status
From [Microsoft.SystemCenter.Orchestrator.Runtime].Jobs AS Jobs
 INNER JOIN POLICIES ON Jobs.RunbookId = POLICIES.UniqueID
 inner join POLICYINSTANCES on jobs.Id = POLICYINSTANCES.JobId 
 where POLICYINSTANCES.Status = 'success' 
 and  TimeEnded > dateadd(HOUR, -300, getdate()) 
 order by Name