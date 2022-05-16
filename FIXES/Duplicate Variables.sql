SELECT TOP (1000) [UniqueID]
      ,[ParentID]
      ,[Name]
      ,[Description]
      ,[PositionX]
      ,[PositionY]
      ,[ObjectType]
      ,[SubType]
      ,[Enabled]
      ,[Flags]
      ,[ASC_UseServiceSecurity]
      ,[ASC_ThisAccount]
      ,[ASC_Username]
      ,[ASC_Password]
      ,[HasExtenders]
      ,[CreationTime]
      ,[CreatedBy]
      ,[LastModified]
      ,[LastModifiedBy]
      ,[Deleted]
      ,[Cost]
      ,[Savings]
      ,[Number]
      ,[AlternateDisplayData]
      ,[ASW_ObjectTimeout]
      ,[ASW_NotifyOnFail]
      ,[Flatten]
      ,[FlatUseLineBreak]
      ,[FlatUseCSV]
      ,[FlatUseCustomSep]
      ,[FlatCustomSep]
  FROM [Orchestrator2016].[dbo].[OBJECTS]

 where  name ='now()'
  --where  uniqueid ='35102501-AE29-4BAC-BBAE-6F786A479AED'

  
  
  
  update [Orchestrator2016].[dbo].[OBJECTS]
  set Deleted =1
  --where  name ='SCOM-SD scan time'
  where uniqueID ='2E11AF61-5370-4C2B-835C-3BD0F0467B83'