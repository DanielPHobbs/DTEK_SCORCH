use Orchestrator

SELECT
[Caller Runbook].Name AS [Caller Runbook Name],
[Caller Runbook].Path AS [Caller Runbook Path],
[Called Runbook].Name AS [Called Runbook Name],
[Called Runbook].Path AS [Called Runbook Path], 
[Microsoft.SystemCenter.Orchestrator].Activities.Name AS [Activity Name]
FROM
[Microsoft.SystemCenter.Orchestrator].Activities INNER JOIN
dbo.TRIGGER_POLICY ON [Microsoft.SystemCenter.Orchestrator].Activities.Id = dbo.TRIGGER_POLICY.UniqueID INNER JOIN
[Microsoft.SystemCenter.Orchestrator].Runbooks AS [Called Runbook] ON dbo.TRIGGER_POLICY.PolicyObjectID = [Called Runbook].Id INNER JOIN
[Microsoft.SystemCenter.Orchestrator].Runbooks AS [Caller Runbook] ON [Microsoft.SystemCenter.Orchestrator].Activities.RunbookId = [Caller Runbook].Id
where [Called Runbook].Name = '1.1-Reset PW'