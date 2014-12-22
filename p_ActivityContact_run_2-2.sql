USE [Productions_MSCRM]
GO

/****** Object:  StoredProcedure [dbo].[p_ActivityContact_run_2]    Script Date: 08/28/2014 14:48:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

ALTER PROCEDURE [dbo].[p_ActivityContact_run_2]
AS 
BEGIN 
SET QUOTED_IDENTIFIER ON

DECLARE @StartTime Datetime
SET @StartTime = GETDATE()
PRINT 'Stored Procedure starts at ' + CONVERT( VARCHAR(50), @StartTime )

DECLARE @sCatalog VARCHAR(20), @sSchema VARCHAR(10), @sTable VARCHAR(50), @sView VARCHAR(50)
DECLARE @sFullNameTable VARCHAR(80),@sFullNameView VARCHAR(50)
DECLARE @sSQL VARCHAR(MAX)
DECLARE @iMaxColumn INT, @iRange INT, @sRange VARCHAR(10), @iColumn INT, @i INT
DECLARE @iMaxNumActivityPerContact INT, @iCountContactRow INT, @iCountActivity INT, @iCountZeroContact INT, @sTotalActivity VARCHAR(50)
SET @sCatalog = 'Productions_MSCRM'
SET @sSchema = 'dbo'
SET @sTable = 'tblActivityContact'			/* Constant */
SET @sView = 'vw_ActivityContact_run_1'		/* Constant */
SET @sFullNameTable = @sCatalog + '.' + @sSchema + '.' + @sTable
SET @sFullNameView = @sCatalog + '.' + @sSchema + '.' + @sView
PRINT @sFullNameTable
PRINT @sFullNameView  

IF( EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES   
    WHERE TABLE_CATALOG = @sCatalog
      AND TABLE_SCHEMA = @sSchema 
      AND TABLE_NAME = @sTable ))
	BEGIN
		SET @sSQL = 'DROP TABLE ' + @sFullNameTable
		PRINT @sSQL;
		EXEC( @sSQL );
	END

SET @sSQL = 'CREATE TABLE ' + @sFullNameTable
SET @sSQL = @sSQL +	'( [Owner] nVARCHAR(160)'
SET @sSQL = @sSQL +	', [Contact] nVARCHAR(300)'
SET @sSQL = @sSQL +	', [Contact ID] UniqueIdentifier NOT NULL'
SET @sSQL = @sSQL +	', [Account] nVARCHAR(160)'
SET @sSQL = @sSQL +	', [Lead Source] nVARCHAR(128)'
SET @sSQL = @sSQL +	', [Contact Description] nVARCHAR(MAX)'
SET @sSQL = @sSQL +	', [Range] INT '
SET @sSQL = @sSQL +	');'
BEGIN TRANSACTION C1
PRINT @sSQL;
EXEC( @sSQL );

SET @sSQL = 'CREATE INDEX Contact_idx ON ' + @sFullNameTable + '( Contact ); ';
SET @sSQL = @sSQL + 'CREATE INDEX ContactID_idx ON ' + @sFullNameTable + '( [Contact ID] );';
PRINT @sSQL;
EXEC( @sSQL );
COMMIT TRANSACTION C1

/*********************************************************************************/

DECLARE @iCB INT, @iTy INT, @iSu INT, @sCB VARCHAR(10), @sTy VARCHAR(10), @sSu VARCHAR(10)
SELECT @iCB = MAX( S.[Length] ) FROM ( SELECT LEN( [Created By] ) AS [Length] FROM vw_ActivityContact_run_1 ) AS S
SELECT @iTy = MAX( S.[Length] ) FROM ( SELECT LEN( [Type] ) AS [Length] FROM vw_ActivityContact_run_1 ) AS S
SELECT @iSu = MAX( S.[Length] ) FROM ( SELECT LEN( [Subject] ) AS [Length] FROM vw_ActivityContact_run_1 ) AS S
SET @sCB = CONVERT( VARCHAR(10), 3 + @iCB )
SET @sTy = CONVERT( VARCHAR(10), @iTy )
SET @sSu = CONVERT( VARCHAR(10), 10 + @iSu )

SELECT @iMaxNumActivityPerContact = MAX( T.CT ) 
	FROM 
	( SELECT COUNT(*) AS CT
		FROM [Productions_MSCRM].[dbo].[vw_ActivityContact_run_1]
		GROUP BY [Contact], [Owner], [Account], [Lead Source]
	) AS T;
PRINT @sSQL;
PRINT '@iMaxNumActivityPerContact = ' + CONVERT( VARCHAR(10), @iMaxNumActivityPerContact )

IF @iMaxNumActivityPerContact <= 0
	BEGIN RETURN END

SET @i = 1
SET @iMaxColumn = 50	/* Constant */
BEGIN TRANSACTION C2	
WHILE @i <= @iMaxNumActivityPerContact AND @i <= @iMaxColumn
	BEGIN
	SET @sSQL = 'ALTER TABLE ' + @sFullNameTable + ' ADD'
	SET @sSQL = @sSQL +	' [' + CONVERT( VARCHAR(10), @i ) + ' Created On] Datetime' 
	SET @sSQL = @sSQL +	', [' + CONVERT( VARCHAR(10), @i ) + ' Created By] nVARCHAR( ' + @sCB + ' )'
	SET @sSQL = @sSQL +	', [' + CONVERT( VARCHAR(10), @i ) + ' Type] nVARCHAR( ' + @sTy + ' )'  
	SET @sSQL = @sSQL +	', [' + CONVERT( VARCHAR(10), @i ) + ' State] nVARCHAR( 9 )'  
	SET @sSQL = @sSQL +	', [' + CONVERT( VARCHAR(10), @i ) + ' Subject] nVARCHAR( ' + @sSu + ' )'  
	SET @sSQL = @sSQL +	', [' + CONVERT( VARCHAR(10), @i ) + ' Description] nVARCHAR(MAX);'
	PRINT @sSQL	
	EXEC( @sSQL )		
	SET @i = @i + 1
	END
COMMIT TRANSACTION C2

/*********************************************************************************/

DECLARE @curRow INT
DECLARE @curOwner nVARCHAR(160), @curContact nVARCHAR(4000), @curContactID UniqueIdentifier
DECLARE @curAccount nVARCHAR(160), @curLeadSource nVARCHAR(128), @curContactDescription nVARCHAR(MAX)
DECLARE @curType nVARCHAR(21), @curState nVARCHAR(9), @curSubject nVARCHAR(200), @curCreatedOn Datetime
DECLARE @curCreatedBy nVARCHAR(160), @curActivityDescription nVARCHAR(MAX)

IF CURSOR_STATUS( 'global','CurContactAct' ) >= -1	/* remove Cursor if already exists */
	BEGIN DEALLOCATE CurContactAct END
DECLARE CurContactAct CURSOR FOR	/* SQL Cursor */
	SELECT ROW_NUMBER() OVER( PARTITION BY [Contact] ORDER BY [Contact], [Created On] DESC ) AS Row
	, [Owner], [Contact], [Contact ID], [Account], [Lead Source], [Contact Description] 
	, [Created On], [Created By], [Type], [State], [Subject], [Activity Description] 
	FROM Productions_MSCRM.dbo.vw_ActivityContact_run_1
OPEN CurContactAct	
					
SELECT @i = COUNT( * ) FROM vw_ActivityContact_run_1
SET @sTotalActivity = ' of ' + CONVERT( VARCHAR(10), @i ) + ' Activities proceeded.'
SET @iCountContactRow = 0  
SET @iCountActivity = 0
SET @iCountZeroContact = 0
FETCH NEXT FROM CurContactAct INTO @curRow, @curOwner, @curContact
	, @curContactID, @curAccount, @curLeadSource, @curContactDescription
	, @curCreatedOn, @curCreatedBy, @curType, @curState, @curSubject, @curActivityDescription
WHILE @@FETCH_STATUS = 0
	BEGIN 
	IF @curOwner IS NULL BEGIN SET @curOwner = '' END
	IF @curContact IS NULL BEGIN SET @curContact = '' END
	IF @curAccount IS NULL BEGIN SET @curAccount = '' END
	IF @curLeadSource IS NULL BEGIN SET @curLeadSource = '' END
	IF @curContactDescription IS NULL BEGIN SET @curContactDescription = '' END
	IF @curSubject IS NULL BEGIN SET @curSubject = '' END
	IF @curType IS NULL BEGIN SET @curType = '' END
	IF @curCreatedBy IS NULL BEGIN SET @curCreatedBy = '' END 
	IF @curActivityDescription IS NULL BEGIN SET @curActivityDescription = '' END
	
	SET @curOwner = REPLACE( @curOwner, CHAR(39), '''''' )
	SET @curContact = REPLACE( @curContact, CHAR(39), '''''' )
	SET @curAccount = REPLACE( @curAccount, CHAR(39), '''''' )
	SET @curLeadSource = REPLACE( @curLeadSource, CHAR(39), '''''' )
	SET @curContactDescription = LTRIM( RTRIM( @curContactDescription ))
	SET @curContactDescription = REPLACE( @curContactDescription, CHAR(39), '''''' )	
	SET @curCreatedBy = REPLACE( @curCreatedBy, CHAR(39), '''''' )
	SET @curCreatedBy = REPLACE( @curCreatedBy, 'Matt Pearson', 'System' )	
	SET @curSubject = REPLACE( @curSubject, CHAR(39), '''''' )
	SET @curActivityDescription = LTRIM( RTRIM( @curActivityDescription ))
	SET @curActivityDescription = REPLACE( @curActivityDescription, CHAR(39), '''''' )
	IF @curType LIKE 'E-mail' 
		BEGIN 
		IF LEN( @curActivityDescription ) > 0
			BEGIN	
			SET @curActivityDescription = DBO.[udf_RemoveTextInAngleBracket]( @curActivityDescription ) 
			SET @curActivityDescription = REPLACE( @curActivityDescription, 'v\:*', '' )
			SET @curActivityDescription = REPLACE( @curActivityDescription, 'o\:*', '' )
			SET @curActivityDescription = REPLACE( @curActivityDescription, 'w\:*', '' )			
			SET @curActivityDescription = REPLACE( @curActivityDescription, '.shape', '' )
			SET @curActivityDescription = DBO.[udf_RemoveTextInCurlyBracket]( @curActivityDescription ) 
			SET @curActivityDescription = DBO.[udf_RemoveMTMLcodeInString]( @curActivityDescription ) 
			END
		END	
	SET @iRange = ( @curRow - 1 ) / @iMaxColumn + 1
	SET @sRange = CONVERT( VARCHAR(10), @iRange )
	IF @curRow % @iMaxColumn <> 0 
		BEGIN SET @iColumn = @curRow % @iMaxColumn END
	ELSE 
		BEGIN SET @iColumn = @iMaxColumn END
	
	IF @iColumn = 1
		BEGIN
		SET @sSQL = 'INSERT INTO ' + @sFullNameTable
		SET @sSQL = @sSQL +	' ( [Owner], [Contact], [Contact ID], [Account], [Lead Source], [Contact Description], [Range] ) '
		SET @sSQL = @sSQL +	'VALUES ( '
		SET @sSQL = @sSQL +	'''' + @curOwner + ''', ''' + @curContact + ''''
		SET @sSQL = @sSQL +	', ''' + CONVERT( VARCHAR(50), @curContactID ) + ''''
		SET @sSQL = @sSQL +	', ''' + @curAccount + ''', ''' + @curLeadSource + ''''
		SET @sSQL = @sSQL +	', ''' + @curContactDescription + ''', ' + @sRange + ' );'
		BEGIN TRANSACTION C3
		PRINT @sSQL;
		EXEC( @sSQL )
		COMMIT TRANSACTION C3
		SET @iCountContactRow = @iCountContactRow + 1
		END
	
	SET @sSQL = 'UPDATE ' + @sFullNameTable + ' '
	SET @sSQL = @sSQL + 'SET [' + CONVERT( VARCHAR(10), @iColumn ) + ' Created On] = ''' + CONVERT( VARCHAR(100), @curCreatedOn ) + ''' ' 
	SET @sSQL = @sSQL + ', [' + CONVERT( VARCHAR(10), @iColumn ) + ' Created By] = ''' + @curCreatedBy + ''' '
	SET @sSQL = @sSQL + ', [' + CONVERT( VARCHAR(10), @iColumn ) + ' Type] = ''' + @curType + ''' '
	SET @sSQL = @sSQL + ', [' + CONVERT( VARCHAR(10), @iColumn ) + ' State] = ''' + @curState + ''' '
	SET @sSQL = @sSQL + ', [' +	CONVERT( VARCHAR(10), @iColumn ) + ' Subject] = ''' + @curSubject + ''''
	SET @sSQL = @sSQL + ', [' + CONVERT( VARCHAR(10), @iColumn ) + ' Description] = ''' + @curActivityDescription + ''' '
	SET @sSQL = @sSQL + 'FROM ' + @sFullNameTable + ' '
	SET @sSQL = @sSQL + 'WHERE [Contact ID] = ''' + CONVERT( VARCHAR(50), @curContactID ) + '''' 
	SET @sSQL = @sSQL + ' AND [Range] = ''' + @sRange + ''';' 	
	BEGIN TRANSACTION C4
	PRINT @sSQL
	EXEC( @sSQL )
	COMMIT TRANSACTION C4
	SET @iCountActivity = @iCountActivity + 1
	PRINT CONVERT( VARCHAR(10), @iCountActivity ) + @sTotalActivity
	FETCH NEXT FROM CurContactAct INTO @curRow, @curOwner, @curContact
		, @curContactID, @curAccount, @curLeadSource, @curContactDescription
		, @curCreatedOn, @curCreatedBy, @curType, @curState, @curSubject, @curActivityDescription
	END
CLOSE CurContactAct
DEALLOCATE CurContactAct

/*********************************************************************************/

SELECT @iCountZeroContact = COUNT( DISTINCT C.ContactID )
	FROM ContactBase AS C
		, ContactExtensionBase AS Cext
		, ActivityPointerBase AS A
	WHERE C.ContactID = Cext.ContactID 
		AND Cext.[new_LeadSourceName] IS NOT NULL
		AND C.ContactID NOT IN 
		(
			SELECT DISTINCT [Contact ID] 
				FROM vw_ActivityContact_run_1 
		);	

PRINT 'Number of Contact Has Zero Activities: ' + CONVERT( VARCHAR(10), @iCountZeroContact )

INSERT tblActivityContact( [Owner], [Contact], [Contact ID], [Account], [Lead Source], [Contact Description] ) 
	SELECT DISTINCT ISNULL( Ow2.Name, '' ) AS [Owner]
		, C.FullName AS [Contact]
		, C.ContactID AS [Contact ID]
		, ISNULL( Ow3.Name, '' ) AS [Account]
		, Cpg.Name AS [Lead Source]
		, ISNULL( C.Description, '' ) AS [Contact Description]
		FROM ContactBase AS C
				LEFT JOIN OwnerBase AS Ow2 ON Ow2.OwnerID = C.OwnerID 
				LEFT JOIN AccountBase AS Ow3 ON Ow3.AccountID = C.ParentCustomerID
			, ContactExtensionBase AS Cext
				LEFT JOIN CampaignBase AS Cpg ON Cpg.CampaignID = Cext.[new_LeadSourceName] 
		WHERE C.ContactID = Cext.ContactID 
			AND Cext.[new_LeadSourceName] IS NOT NULL
			AND C.ContactID NOT IN 
			(
				SELECT DISTINCT [Contact ID] 
					FROM vw_ActivityContact_run_1 
			)
		ORDER BY C.FullName;	

/*********************************************************************************/

PRINT CHAR(13) + 'Time elapsed: ' +  CONVERT( VARCHAR(50), CONVERT( TIME(0),( GETDATE() - @StartTime )))
SELECT 'Activity per Contact' AS [Procedure]
	, @iCountContactRow AS [Number of Contact-Row]
	, @iCountActivity AS [Number of Activity]
	, @iMaxNumActivityPerContact AS [Max num of Activity per Contact]
	, @iCountZeroContact AS [Number of Contact Has Zero Activities]
	, CONVERT( TIME(0),( GETDATE() - @StartTime )) AS [Time elapsed]

END 

GO


