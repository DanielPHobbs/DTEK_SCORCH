/****** Script for SelectTopNRows command from SSMS ******/
SELECT Ob.Name, [DaysOfWeek] ,[DaysOfMonth],[Monday] ,[Tuesday] ,[Wednesday] ,[Thursday] ,[Friday] ,[Saturday] ,[Sunday] ,[First] ,[Second] ,[Third] ,[Fourth]
,[Last] ,[Days] ,[Hours] FROM SCHEDULES S
Inner join OBJECTS Ob on Ob.UniqueID = S.UniqueID