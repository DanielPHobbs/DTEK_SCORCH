SELECT P.Name AS [Runbook Name], O.Name AS [Activity Name], OT.Name AS [Activity Type], OA.Action, 
CASE WHEN OA.Attribute LIKE '%[0-F][0-F][0-F][0-F][0-F][0-F][0-F][0-F]-[0-F][0-F][0-F][0-F]- 
[0-F][0-F][0-F][0-F]-[0-F][0-F][0-F][0-F]-[0-F][0-F][0-F][0-F][0-F][0-F][0-F][0-F][0-F][0-F] 
[0-F][0-F]%' 
THEN 'NEW ACTIVITY' ELSE OA.Attribute END AS Attribute, OA.OldValue, OA.NewValue, CIH.DateTime AS 
[Change Timestamp], S.Account AS [User] 
FROM OBJECT_AUDIT AS OA INNER JOIN 
OBJECTS AS O ON OA.ObjectID = O.UniqueID INNER JOIN 
POLICIES AS P ON O.ParentID = P.UniqueID INNER JOIN 
OBJECTTYPES AS OT ON OA.ObjectType = OT.UniqueID INNER JOIN 
CHECK_IN_HISTORY AS CIH ON CIH.UniqueID = OA.TransactionID INNER JOIN 
SIDS AS S ON CIH.CheckInUser = S.SID 
WHERE (O.Deleted = 0) 
ORDER BY [Change Timestamp] DESC