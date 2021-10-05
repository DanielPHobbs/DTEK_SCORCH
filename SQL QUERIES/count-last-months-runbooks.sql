use Orchestrator

select	Name, count(policies.UniqueID) as Count from Policies 
inner join POLICYINSTANCES on policies.UniqueID=POLICYINSTANCES.PolicyID 
where POLICYINSTANCES.TimeEnded > DATEADD(mm,-1,GETDATE()) 
group by name 
order by Count Desc