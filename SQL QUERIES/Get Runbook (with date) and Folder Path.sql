use orchestrator;
 
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
 
SELECT f.path AS RunbookPath,r.name AS RunbookName, r.LastModified AS 'Runbook Last Modified' from POLICIES as r 
INNER JOIN RunbookPath as f on r.ParentID = f.UniqueID
WHERE r.Deleted = 0