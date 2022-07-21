--If you want to check how many log files are in existence per each individual runbook that you have

SELECT COUNT(OBJECTS.Name) AS Instances,
POLICIES.Name AS Runbook, 
Objects.Name 
FROM OBJECTINSTANCEDATA
INNER JOIN OBJECTINSTANCES ON OBJECTINSTANCEDATA.ObjectInstanceID = OBJECTINSTANCES.UniqueID
INNER JOIN OBJECTS ON OBJECTINSTANCES.ObjectID = OBJECTS.UniqueID
INNER JOIN POLICIES ON OBJECTS.ParentID = POLICIES.UniqueID
GROUP BY Policies.Name,
Objects.Name
ORDER BY Instances DESC
