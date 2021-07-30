DECLARE @low INT, @high INT, @host nvarchar(255),@start datetime, @end datetime
 
--Which host do you want to see jobs on?
SET @host = 'MGOAPSMAP2'
--What time range do you want to see jobs for?
set @start = '2014-11-20 11:13:00'
set @end = '2014-11-20 14:35:00'
 
SELECT @low = LowKey, @high = HighKey
FROM [SMA].[Queues].[Deployment]
WHERE ComputerName = @host
 
select r.RunbookName,
count(*) as RunCount
from sma.core.vwJobs as j
inner join [SMA].[Core].[RunbookVersions] as v
on j.RunbookVersionId = v.RunbookVersionId
inner join [SMA].[Core].[Runbooks] as r
on v.RunbookKey = r.RunbookKey
where PartitionId > @low and PartitionId @start
and StartTime < @end
group by r.RunbookName