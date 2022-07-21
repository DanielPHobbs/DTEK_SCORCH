--Orchestrator Job Statuses
Use Orchestrator2016
SELECT POLICIES.Name,
cOUNT(*)
FROM [Microsoft.SystemCenter.Orchestrator.Runtime.Internal].Jobs INNER JOIN
POLICIES ON [Microsoft.SystemCenter.Orchestrator.Runtime.Internal].Jobs.RunbookId = POLICIES.UniqueID
WHERE ([Microsoft.SystemCenter.Orchestrator.Runtime.Internal].Jobs.StatusId NOT LIKE '4')
AND ([Microsoft.SystemCenter.Orchestrator.Runtime.Internal].Jobs.StatusId NOT LIKE '3')
AND ([Microsoft.SystemCenter.Orchestrator.Runtime.Internal].Jobs.StatusId NOT LIKE '2')
group by POLICIES.Name
order by COUNT(*) DESC

--What runbooks are currently running and where
SELECT
J.[RunbookId],
P.[Name],
A.[Computer]
FROM
[Orchestrator2016].[Microsoft.SystemCenter.Orchestrator.Runtime.Internal].[Jobs] J
INNER JOIN [dbo].POLICIES P ON J.RunbookId = P.UniqueID
INNER JOIN [dbo].[ACTIONSERVERS] A ON J.RunbookServerId = A.UniqueID
WHERE
J.[StatusId] = 1
ORDER BY
P.[Name] DESC