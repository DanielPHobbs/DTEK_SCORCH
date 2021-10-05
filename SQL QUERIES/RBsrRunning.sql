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