use Orchestrator
SELECT *
   FROM [Orchestrator].[dbo].[POLICY_PUBLISH_QUEUE] ppq
   LEFT JOIN Orchestrator.dbo.POLICYINSTANCES pin on pin.PolicyID = ppq.PolicyID
   Where AssignedActionServer IS NOT NULL
   and TimeEnded IS NULL
   and DATEDIFF(MINUTE, Heartbeat, GetDate()) > 5