SELECT MAX(jobs.MaxInstances) FROM
(SELECT [ActionServer]
, CONVERT(VARCHAR(19), dateadd(hour,datediff(hour,0,TimeStarted),0), 120) as Hourly
, COUNT(*) as MaxInstances
FROM [Opalis].[dbo].[POLICYINSTANCES]
WHERE (DateDiff(M, TimeStarted,GETDATE()) < 1 )
GROUP BY ActionServer, CONVERT(VARCHAR(19), dateadd(hour,datediff(hour,0,TimeStarted),0), 120)) as jobs