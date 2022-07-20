--Originally Written by: Narayana Vyas Kondreddi
--Modified By: Jon Mattivi
--Purpose: Search all tables and columns in the Orchestrator database to find variable instances

DECLARE @VarName nvarchar(100), @VarID nvarchar(100)
SET @VarName = 'AD-Admin'

SET @VarID = (Select VARIABLES.UniqueID
From VARIABLES
INNER JOIN OBJECTS ON OBJECTS.UniqueID = VARIABLES.UniqueID
Where OBJECTS.Name = @VarName and OBJECTS.Deleted != 1)

   
CREATE TABLE #Results (RunbookPath nvarchar(1000), RunbookName nvarchar(250), ActivityName nvarchar(250), [Table.Column] nvarchar(370))

SET NOCOUNT ON

DECLARE @TableName nvarchar(256), @ColumnName nvarchar(128), @SearchStr2 nvarchar(110)

SET  @TableName = ''
SET @SearchStr2 = QUOTENAME('%' + @VarID + '%','''')

WHILE @TableName IS NOT NULL
   
BEGIN
    SET @ColumnName = ''
    SET @TableName =
    (
        SELECT MIN(QUOTENAME(TABLE_SCHEMA) + '.' + QUOTENAME(TABLE_NAME))
        FROM     INFORMATION_SCHEMA.TABLES
        WHERE         TABLE_TYPE = 'BASE TABLE'
            AND    QUOTENAME(TABLE_SCHEMA) + '.' + QUOTENAME(TABLE_NAME) > @TableName
            AND (TABLE_SCHEMA) = 'dbo'
            AND    OBJECTPROPERTY(
                    OBJECT_ID(
                        QUOTENAME(TABLE_SCHEMA) + '.' + QUOTENAME(TABLE_NAME)
                         ), 'IsMSShipped'
                           ) = 0
    )

    WHILE (@TableName IS NOT NULL) AND (@ColumnName IS NOT NULL) AND (@TableName != '[dbo].[OBJECTINSTANCEDATA]') AND (@TableName != '[dbo].[OBJECTINSTANCES]') AND (@TableName != '[dbo].[POLICYINSTANCES]') AND ((SELECT TOP 1 COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = PARSENAME(@TableName, 1) AND (COLUMN_NAME) = 'UniqueID' AND (DATA_TYPE) = 'uniqueidentifier' AND (TABLE_SCHEMA) = 'dbo') is not null)
           
    BEGIN
        SET @ColumnName =
        (
            SELECT MIN(QUOTENAME(COLUMN_NAME))
            FROM     INFORMATION_SCHEMA.COLUMNS
            WHERE         TABLE_SCHEMA    = PARSENAME(@TableName, 2)
                AND    TABLE_NAME    = PARSENAME(@TableName, 1)
                AND    DATA_TYPE IN ('char', 'datetime', 'decimal', 'int', 'money', 'ntext', 'nvarchar', 'varbinary', 'varchar')
                AND    QUOTENAME(COLUMN_NAME) > @ColumnName
        )
        IF @ColumnName IS NOT NULL
           
        BEGIN
            INSERT INTO #Results
            EXEC
            (
                'IF EXISTS (Select TOP 1 ' + @TableName + '.' + @ColumnName + 'From ' + @TableName + ' (NOLOCK) WHERE ' + @TableName + '.' + @ColumnName + ' LIKE ' + @SearchStr2 + ')' +
                'BEGIN ' +
                'SELECT Resources2.[Path], Policy.[Name], ActObj.[Name],''' + @TableName + '.' + @ColumnName + '''' +
                'FROM ' + @TableName + ' (NOLOCK) ' +
                'INNER JOIN [dbo].[OBJECTS] ActObj (NOLOCK) ON ' + @TableName + '.[UniqueID] = ActObj.[UniqueID]' +
                'INNER JOIN [dbo].[POLICIES] Policy (NOLOCK) ON ActObj.[ParentID] = Policy.[UniqueID]' +
                'INNER JOIN [Microsoft.SystemCenter.Orchestrator.Internal].[Resources] Resources2 (NOLOCK) ON Policy.[UniqueID] = Resources2.[UniqueId]' +
                'WHERE ' + @TableName + '.' + @ColumnName + ' LIKE ' + @SearchStr2 + 'and ActObj.[Deleted] != 1 ' +
                'END'
            )
        END
    END   
END

SELECT * FROM #Results
Order By RunbookPath
   
DROP TABLE #Results