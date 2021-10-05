SELECT pin.PolicyID
, pin.State
, pin.Status
, oi.ObjectID
, oi.ObjectStatus
, oi.StartTime
, oi.EndTime
FROM [Orchestrator].[dbo].[POLICYINSTANCES] pin
LEFT JOIN Orchestrator.dbo.[OBJECTS] obj on obj.ParentID = pin.PolicyID
LEFT JOIN Orchestrator.dbo.OBJECTINSTANCES oi on oi.ObjectID = obj.UniqueID
Where TimeEnded IS NULL
And Status IS NULL
and oi.EndTime IS NULL
and OI.StartTime IS NOT NULL
and DATEDIFF(MINUTE,oi.StartTime,getdate()) > 5