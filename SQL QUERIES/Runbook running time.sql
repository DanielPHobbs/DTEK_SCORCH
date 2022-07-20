select 
Policies.name,
ins.TimeStarted,
DateDiff(second, ins.TimeStarted, GETUTCDATE()) as Totalseconds ,
DateDiff(second, ins.TimeStarted, GETUTCDATE()) / 3600 as Hours, 
(DateDiff(second, ins.TimeStarted, GETUTCDATE()) % 3600) / 60 as Minutes, 
DateDiff(second, ins.TimeStarted, GETUTCDATE()) % 60 as Seconds

from POLICYINSTANCES as Ins
inner join POLICIES on Ins.PolicyID=POLICIES.UniqueID

 where Ins.Status is null and Name not like '%-MON-%'
order by Totalseconds desc