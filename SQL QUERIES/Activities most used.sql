use Orchestrator 

Select Count(ObjectID) as entries, ObjectID, obj.[Name] as [Action], pol.[Name] as [policy]
FROM Orchestrator.dbo.OBJECTINSTANCES as inst
join [Orchestrator].[dbo].[OBJECTS] as obj on inst.ObjectID = Obj.UniqueID
join [orchestrator].[dbo].[POLICIES] as pol on pol.UniqueID = obj.ParentID
group by obj.[Name],pol.[Name], OBJECTID
order by entries desc