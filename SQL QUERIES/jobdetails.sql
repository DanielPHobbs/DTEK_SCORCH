Select RB.Name, RB.Path, RT.Id, RT.RunbookServerId, A.Computer , RT.Status, RT.Parameters, RT.LastModifiedTime, RT.LastModifiedBy from 

[Microsoft.SystemCenter.Orchestrator.Runtime].[Jobs] as RT

inner join [Orchestrator2016].[Microsoft.SystemCenter.Orchestrator].[Runbooks] RB on RB.Id= RT.RunbookId 

inner join ACTIONSERVERS A on A.UniqueID = RT.RunbookServerId

order by RT.CreationTime desc