Use Orchestrator2016;

 
 --StatusId is current stats of the job.
--0 = Pending
--1 = Running
--2 = Failed
--3 = Cancelled
--4 = Completed


with RunbookPath as
(
select 'Policies\' + cast(name as varchar(max)) as [path], uniqueid from dbo.folders b
where b.ParentID='00000000-0000-0000-0000-000000000000' and disabled = 0 and deleted= 0
union all
select cast(c.[path] + '\' + cast(b.name as varchar(max)) as varchar(max)), b.uniqueid from dbo.FOLDERS b
inner join
RunbookPath c on b.ParentID = c.UniqueID
where b.Disabled = 0 and b.Deleted = 0
)
 
select f.path AS RunbookPath,r.name AS RunbookName,A.[Computer] as RunbookServer from POLICIES as r 
inner join RunbookPath as f on r.ParentID = f.UniqueID
inner join [Microsoft.SystemCenter.Orchestrator.Runtime.Internal].[Jobs] as J on J.RunbookId = r.UniqueID
INNER JOIN [dbo].[ACTIONSERVERS] A ON J.RunbookServerId = A.UniqueID
WHERE r.Deleted = 0 AND
J.[StatusId] = 4
