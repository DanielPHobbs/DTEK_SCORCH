SELECT p.Name 
      ,p.CheckOutTime 
      ,p.CheckOutLocation 
      ,p.Description 
         ,s.Account 
  FROM Orchestrator.dbo.POLICIES P 
  Join Orchestrator.dbo.SIDS S on P.CheckOutUser = S.SID 
  Where CheckOutUser is not NULL and p.Deleted = 0