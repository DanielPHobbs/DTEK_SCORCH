Use Orchestrator

DECLARE @DBName NVARCHAR(128)    = ”          — Name of the db, empty = current catalog

DECLARE @ReorgLimit TINYINT                  = 15        — Minimum fragmentation % to recommend reorg

DECLARE @RebuildLimit TINYINT              = 30        — Minimum fragmentation % to recommend rebuild

DECLARE @PageLimit SMALLINT                               = 10        — Minimum # of Pages before you worry about the index

DECLARE @ShowAllIndexes BIT                 = 0                          — 0 = Only show reorg/rebuild recommended, 1 = Show All

 

SET NOCOUNT ON ;

SET DEADLOCK_PRIORITY LOW ;

 

BEGIN TRY

 

    DECLARE @FullName NVARCHAR(400), @SQL NVARCHAR(1000), @Rebuild NVARCHAR(1000), @DBID SMALLINT ;

    DECLARE @Error INT, @TableName NVARCHAR(128), @SchemaName NVARCHAR(128), @HasLobs TINYINT ;

    DECLARE @object_id INT, @index_id INT, @partition_number INT, @AvgFragPercent TINYINT ;

    DECLARE @IndexName NVARCHAR(128), @Partitions INT, @Print NVARCHAR(1000) ;

    DECLARE @PartSQL NVARCHAR(600), @ReOrgFlag TINYINT, @IndexTypeDesc NVARCHAR(60) ;

 

                — Get the ID of the Database Catalog

                IF @DBName = ” SET @DBName = DB_NAME();

    SET @DBID = DB_ID(@DBName) ;

 

                IF OBJECT_ID(‘tempdb..#FragLevels’) IS NOT NULL DROP TABLE #FragLevels

               

                — Create a temporary table to store results

    CREATE TABLE #FragLevels (

        [SchemaName] NVARCHAR(128) NULL, [TableName] NVARCHAR(128) NULL, [HasLOBs] TINYINT NULL,

                    [ObjectID] [int] NOT NULL, [IndexID] [int] NOT NULL, [PartitionNumber] [int] NOT NULL,

        [AvgFragPercent] [tinyint] NOT NULL, [IndexName] NVARCHAR(128) NULL, [IndexTypeDesc] NVARCHAR(60) NOT NULL ) ;

 

    — Get the initial list of indexes and partitions to work on filtering out heaps and meeting the specified thresholds

    INSERT INTO #FragLevels

       ([ObjectID], [IndexID], [PartitionNumber], [AvgFragPercent], [IndexTypeDesc])

    SELECT a.[object_id], a.[index_id], a.[partition_number], CAST(a.[avg_fragmentation_in_percent] AS TINYINT) AS [AvgFragPercent], a.[index_type_desc]

        FROM sys.dm_db_index_physical_stats(@DBID, NULL, NULL, NULL , ‘LIMITED’) AS a

                                                WHERE

                                                                ((@ShowAllIndexes = 0 AND a.[avg_fragmentation_in_percent] >= @ReorgLimit) OR (@ShowAllIndexes <> 0)) AND

                                                                a.[page_count] >= @PageLimit AND

                                                                a.[index_id] > 0

 

    — Create an index to make some of the updates & lookups faster

    CREATE INDEX [IX_#FragLevels_OBJECTID] ON #FragLevels([ObjectID]) ;

 

    — Get the Schema and Table names for each

    UPDATE #FragLevels WITH (TABLOCK)

        SET [SchemaName] = OBJECT_SCHEMA_NAME([ObjectID],@DBID),

            [TableName] = OBJECT_NAME([ObjectID],@DBID) ;

 

    — Determine if the index has a Large Object (LOB) datatype.

                — LOBs prevent reindexing and rebuilding index online

    SET @SQL = N’UPDATE #FragLevels WITH (TABLOCK) SET [HasLOBs] = (SELECT TOP 1 CASE WHEN t.[lob_data_space_id] = 0 THEN 0 ELSE 1 END ‘ +

            N’ FROM [‘ + @DBName  + N’].[sys].[tables] AS t WHERE t.[type] = ”U” AND t.[object_id] = #FragLevels.[ObjectID])’ ;

 

    EXEC(@SQL) ;

 

    —  Get the index name

    SET @SQL = N’UPDATE #FragLevels SET [IndexName] = (SELECT TOP 1 t.[name] FROM [‘ + @DBName  + N’].[sys].[indexes] AS t WHERE t.[object_id] = #FragLevels.[ObjectID] ‘ +

                        ‘ AND t.[index_id] = #FragLevels.[IndexID] )’  ;

 

    EXEC(@SQL) ;

 

                — Return the results

    SELECT

                                F.SchemaName AS [Schema Name],

                                F.TableName AS [Table Name],

                                F.IndexName AS [Index Name],

                                F.IndexTypeDesc AS [Index Type],

                                F.AvgFragPercent AS [Avg Frag (%)],

                                CASE

                                                WHEN F.AvgFragPercent >= @RebuildLimit THEN ‘Yes’

                                                ELSE ‘No’

                                END AS [Should Rebuild],

                                CASE

                                                WHEN F.AvgFragPercent >= @ReorgLimit AND F.AvgFragPercent < @RebuildLimit THEN ‘Yes’

                                                ELSE ‘No’

                                END AS [Should Reorg],

                                CASE

                                                WHEN F.HasLOBs = 1 THEN ‘Yes’

                                                ELSE ‘No’

                                END AS [Has LOBs],

                                F.ObjectID AS [Object ID],

                                F.IndexID AS [Index ID],

                                F.PartitionNumber AS [Partition Number]

                FROM #FragLevels AS F

                ORDER BY AvgFragPercent DESC

 

                IF OBJECT_ID(‘tempdb..#FragLevels’) IS NOT NULL DROP TABLE #FragLevels

               

END TRY

BEGIN CATCH

 

    SELECT

        ERROR_NUMBER() AS ErrorNumber,

        ERROR_SEVERITY() AS ErrorSeverity,

        ERROR_STATE() AS ErrorState,

        ERROR_PROCEDURE() AS ErrorProcedure,

        ERROR_LINE() AS ErrorLine,

        ERROR_MESSAGE() AS ErrorMessage;

 

END CATCH ;