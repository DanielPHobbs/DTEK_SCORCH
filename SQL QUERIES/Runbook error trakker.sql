SELECT 
                            r.Name, r.Description, r.Path, a.Name, oid.Value ,ai.Status, ai.StartTime
                            FROM
                            [Microsoft.SystemCenter.Orchestrator].[Runbooks] r
                            INNER JOIN [Microsoft.SystemCenter.Orchestrator].[Activities] a ON a.RunbookId = r.Id
                            INNER JOIN [Microsoft.SystemCenter.Orchestrator.Runtime].[ActivityInstances] ai ON ai.ActivityId = a.Id
                            INNER JOIN [Microsoft.SystemCenter.Orchestrator].[Resources] res ON res.UniqueId = r.Id
                            INNER JOIN dbo.OBJECTS OBJ on OBJ.ParentID = r.Id
                            INNER JOIN OBJECTINSTANCES OI on OI.ObjectID = OBJ.UniqueID
                            INNER JOIN OBJECTINSTANCEDATA OID on OID.ObjectInstanceID = Oi.UniqueID 
                            WHERE 
                            ai.StartTime >= DATEADD(HOUR, -1, GETDATE()) 
                            and ai.Status = 'failed' 
                            AND OID.[Key] = 'ErrorSummary.Text' 
                            AND OID.Value <> '' 
                            --AND OID.Value <> 'Policy stopped by user.'
                            AND OI.StartTime  >= DATEADD(HOUR, -1, GETDATE())
                            GROUP BY  r.Id,r.Description, r.Path, a.Name , oid.Value ,ai.Status, ai.StartTime, r.Name