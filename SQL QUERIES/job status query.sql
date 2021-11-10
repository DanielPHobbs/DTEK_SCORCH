use Orchestrator2016


SELECT POLICIES.Name,
cOUNT(*)
FROM [Microsoft.SystemCenter.Orchestrator.Runtime.Internal].Jobs INNER JOIN
POLICIES ON [Microsoft.SystemCenter.Orchestrator.Runtime.Internal].Jobs.RunbookId = POLICIES.UniqueID
WHERE ([Microsoft.SystemCenter.Orchestrator.Runtime.Internal].Jobs.StatusId NOT LIKE '4')
AND ([Microsoft.SystemCenter.Orchestrator.Runtime.Internal].Jobs.StatusId NOT LIKE '3')
AND ([Microsoft.SystemCenter.Orchestrator.Runtime.Internal].Jobs.StatusId NOT LIKE '2')
group by POLICIES.Name
order by COUNT(*) DESC

/*
StatusId is current stats of the job.
0 = Pending
1 = Running
2 = Failed
3 = Cancelled
4 = Completed
*/