--TRUNCATE TABLE [Microsoft.SystemCenter.Orchestrator.Internal].AuthorizationCache

Use Orchestrator

 

Truncate table [Microsoft.SystemCenter.Orchestrator.Internal].AuthorizationCache

 

DECLARE @secToken INT

DECLARE tokenCursor CURSOR FOR

 

SELECT

Id

FROM

[Microsoft.SystemCenter.Orchestrator.Internal].SecurityTokens

 

OPEN tokenCursor

 

FETCH NEXT FROM tokenCursor

INTO @secToken

 

WHILE @@FETCH_STATUS = 0

BEGIN

PRINT ‘Computing Authorization Cache for Security Token: ‘ + Convert(Nvarchar, @secToken)

exec [Microsoft.SystemCenter.Orchestrator].ComputeAuthorizationCache @TokenId = @secToken

FETCH NEXT FROM tokenCursor

    INTO @secToken

END

 

CLOSE tokenCursor

DEALLOCATE tokenCursor