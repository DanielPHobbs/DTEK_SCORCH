
--List all actively-running runbooks across all Runbook Servers
SELECT *
FROM [Orchestrator2016].[dbo].[POLICY_PUBLISH_QUEUE] ppq
LEFT JOIN Orchestrator2016.dbo.POLICYINSTANCES pin on pin.PolicyID = ppq.PolicyID
Where AssignedActionServer IS NOT NULL
and TimeEnded IS NULL

--List of all currently-running Jobs that have not had their heartbeat update in the last 5 minutes
SELECT *
FROM [Orchestrator2016].[dbo].[POLICY_PUBLISH_QUEUE] ppq
LEFT JOIN Orchestrator2016.dbo.POLICYINSTANCES pin on pin.PolicyID = ppq.PolicyID
Where AssignedActionServer IS NOT NULL
and TimeEnded IS NULL
and DATEDIFF(MINUTE, Heartbeat, GetDate()) > 5

--List all running Jobs with activities that started more than 5 mins ago and have not finished
SELECT pin.PolicyID
, pin.State
, pin.Status
, oi.ObjectID
, oi.ObjectStatus
, oi.StartTime
, oi.EndTime
FROM [Orchestrator2016].[dbo].[POLICYINSTANCES] pin
LEFT JOIN Orchestrator2016.dbo.[OBJECTS] obj on obj.ParentID = pin.PolicyID
LEFT JOIN Orchestrator2016.dbo.OBJECTINSTANCES oi on oi.ObjectID = obj.UniqueID
Where TimeEnded IS NULL 
And Status IS NULL
and oi.EndTime IS NULL
and OI.StartTime IS NOT NULL
and DATEDIFF(MINUTE,oi.StartTime,getdate()) > 5

--Find out the highest number of runbooks that were run per hour at any given time in the last 30 days:
SELECT MAX(jobs.MaxInstances) FROM 
(SELECT [ActionServer]
, CONVERT(VARCHAR(19), dateadd(hour,datediff(hour,0,TimeStarted),0), 120) as Hourly
, COUNT(*) as MaxInstances
FROM [Orchestrator2016].[dbo].[POLICYINSTANCES]
WHERE (DateDiff(M, TimeStarted,GETDATE()) < 1 )
GROUP BY ActionServer, CONVERT(VARCHAR(19), dateadd(hour,datediff(hour,0,TimeStarted),0), 120)) as jobs
