USE [Productions_MSCRM]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER PROCEDURE [dbo].[p_ActivityOpportunity_run_2]
AS 
BEGIN 
SET QUOTED_IDENTIFIER ON

DECLARE @StartTime Datetime
SET @StartTime = GETDATE()
PRINT 'Stored Procedure starts at ' + CONVERT( VARCHAR(50), @StartTime )

DECLARE @sCatalog VARCHAR(20), @sSchema VARCHAR(10), @sTable VARCHAR(50), @sView VARCHAR(50)
DECLARE @sFullNameTable VARCHAR(80),@sFullNameView VARCHAR(50)
DECLARE @sSQL VARCHAR(MAX)
DECLARE @iMaxNumOpp INT, @iMaxColumn INT, @iRange INT, @sRange VARCHAR(10), @iColumn INT, @i INT
DECLARE @iCountOpportunityRow INT, @iCountActivity INT, @iCountZeroOpportunity INT, @sTotalActivity VARCHAR(50)
SET @sCatalog = 'Productions_MSCRM'
SET @sSchema = 'dbo'
SET @sTable = 'tblActivityOpportunity'			/* Constant */
SET @sView = 'vw_ActivityOpportunity_run_1'		/* Constant */
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
SET @sSQL = @sSQL +	'( [Owner] nVARCHAR(160) NOT NULL '
SET @sSQL = @sSQL +	', [Opportunity] nVARCHAR(300) NOT NULL '
SET @sSQL = @sSQL +	', [Opportunity ID] UniqueIdentifier NULL '
SET @sSQL = @sSQL +	', [Status] nVARCHAR(10) NOT NULL '
SET @sSQL = @sSQL +	', [Invoice Customer] nVARCHAR(160)'
SET @sSQL = @sSQL +	', [Primary Contact] nVARCHAR(160)'
SET @sSQL = @sSQL +	', [Opportunity Partner] nVARCHAR(160)'
SET @sSQL = @sSQL +	', [Opportunity End User] nVARCHAR(160)'
SET @sSQL = @sSQL +	', [Lead Source] nVARCHAR(128)'
SET @sSQL = @sSQL +	', [Inside Sales Support] nVARCHAR(160)'
SET @sSQL = @sSQL +	', [Range] INT '
SET @sSQL = @sSQL +	');'
BEGIN TRANSACTION O1
PRINT @sSQL;
EXEC( @sSQL );

SET @sSQL = 'CREATE INDEX Opportunity_idx ON ' + @sFullNameTable + '( Opportunity );';
SET @sSQL = @sSQL + 'CREATE INDEX OpportunityID_idx ON ' + @sFullNameTable + '( [Opportunity ID] );';
PRINT @sSQL;
EXEC( @sSQL );
COMMIT TRANSACTION O1

DECLARE @iCB INT, @iTy INT, @iSu INT, @sCB VARCHAR(10), @sTy VARCHAR(10), @sSu VARCHAR(10)
SELECT @iCB = MAX( S.[Length] ) FROM ( SELECT LEN( [Created By] ) AS [Length] FROM vw_ActivityOpportunity_run_1 ) AS S
SELECT @iTy = MAX( S.[Length] ) FROM ( SELECT LEN( [Type] ) AS [Length] FROM vw_ActivityOpportunity_run_1 ) AS S
SELECT @iSu = MAX( S.[Length] ) FROM ( SELECT LEN( [Subject] ) AS [Length] FROM vw_ActivityOpportunity_run_1 ) AS S
SET @sCB = CONVERT( VARCHAR(10), 3 + @iCB )
SET @sTy = CONVERT( VARCHAR(10), @iTy )
SET @sSu = CONVERT( VARCHAR(10), 10 + @iSu )

SELECT @iMaxNumOpp = MAX( T.CT ) 
	FROM 
	( SELECT COUNT(*) AS CT
		FROM [Productions_MSCRM].[dbo].[vw_ActivityOpportunity_run_1]
		GROUP BY [Opportunity]
	) AS T;
PRINT @sSQL;
PRINT '@iMaxNumOpp = ' + CONVERT( VARCHAR(10), @iMaxNumOpp )

IF @iMaxNumOpp <= 0
	BEGIN RETURN END

SET @i = 1
SET @iMaxColumn = 50	/* Constant */
BEGIN TRANSACTION O2
WHILE @i <= @iMaxNumOpp AND @i <= @iMaxColumn
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
COMMIT TRANSACTION O2

/*********************************************************************************/

DECLARE @curRow INT
DECLARE @curOwner nVARCHAR(160), @curOpportunity nVARCHAR(300), @curOpportunityID UniqueIdentifier, @curStatus nVARCHAR(10)
DECLARE @curInvoiceCustomer nVARCHAR(300), @curPrimaryContact nVARCHAR(160), @curOpportunityPartner nVARCHAR(160)
DECLARE @curOpportunityEndUser nVARCHAR(160), @curLeadSource nVARCHAR(128), @curInsideSalesSupport nVARCHAR(160)

DECLARE @curType nVARCHAR(21), @curState nVARCHAR(9), @curSubject nVARCHAR(200), @curCreatedOn Datetime
DECLARE @curCreatedBy nVARCHAR(160), @curDescription nVARCHAR(MAX)

IF CURSOR_STATUS( 'global','CurOppAct' ) >= -1	/* remove Cursor if already exists */
	BEGIN DEALLOCATE CurOppAct END
DECLARE CurOppAct CURSOR FOR	/* SQL Cursor */
	SELECT ROW_NUMBER() OVER
	( 
		PARTITION BY [Opportunity], [Owner] 
		ORDER BY [Opportunity], [Owner], [Created On] DESC 
	) AS Row
	, [Owner], [Opportunity], [Opportunity ID], [Status], [Invoice Customer], [Primary Contact]
	, [Opportunity Partner], [Opportunity End User], [Lead Source], [Inside Sales Support]
	, [Created On], [Created By], [Type], [State], [Subject], [Description]
	FROM Productions_MSCRM.dbo.vw_ActivityOpportunity_run_1
	ORDER BY [Opportunity], [Owner], [Created On] DESC
OPEN CurOppAct	

SELECT @i = COUNT( * ) FROM vw_ActivityOpportunity_run_1
SET @sTotalActivity = ' of ' + CONVERT( VARCHAR(10), @i ) + ' Activities proceeded.'
SET @iCountOpportunityRow  = 0
SET @iCountActivity = 0
SET @iCountZeroOpportunity = 0
FETCH NEXT FROM CurOppAct INTO @curRow, @curOwner, @curOpportunity, @curOpportunityID, @CurStatus
	, @curInvoiceCustomer, @curPrimaryContact, @curOpportunityPartner, @curOpportunityEndUser
	, @curLeadSource, @curInsideSalesSupport, @curCreatedOn, @curCreatedBy, @curType, @curState
	, @curSubject, @curDescription
WHILE @@FETCH_STATUS = 0
	BEGIN 
	IF @curOwner IS NULL BEGIN SET @curOwner = '' END
	IF @curOpportunity IS NULL BEGIN SET @curOpportunity = '' END
	IF @curInvoiceCustomer IS NULL BEGIN SET @curInvoiceCustomer = '' END
	IF @curPrimaryContact IS NULL BEGIN SET @curPrimaryContact = '' END
	IF @curOpportunityPartner IS NULL BEGIN SET @curOpportunityPartner = '' END
	IF @curOpportunityEndUser IS NULL BEGIN SET @curOpportunityEndUser = '' END
	IF @curLeadSource IS NULL BEGIN SET @curLeadSource = '' END
	IF @curInsideSalesSupport IS NULL BEGIN SET @curInsideSalesSupport = '' END
	IF @curCreatedBy IS NULL BEGIN SET @curCreatedBy = '' END
	IF @curType IS NULL BEGIN SET @curType = '' END
	IF @curSubject IS NULL BEGIN SET @curSubject = '' END
	IF @curDescription IS NULL BEGIN SET @curDescription = '' END
	
	SET @curOwner = REPLACE( @curOwner, CHAR(39), '''''' )
	SET @curOpportunity = REPLACE( @curOpportunity, CHAR(39), '''''' )
	SET @curInvoiceCustomer = REPLACE( @curInvoiceCustomer, CHAR(39), '''''' )
	SET @curPrimaryContact = REPLACE( @curPrimaryContact, CHAR(39), '''''' )
	SET @curOpportunityPartner = REPLACE( @curOpportunityPartner, CHAR(39), '''''' )
	SET @curOpportunityEndUser = REPLACE( @curOpportunityEndUser, CHAR(39), '''''' )
	SET @curLeadSource = REPLACE( @curLeadSource, CHAR(39), '''''' )
	SET @curInsideSalesSupport = REPLACE( @curInsideSalesSupport, CHAR(39), '''''' )
	SET @curCreatedBy = REPLACE( @curCreatedBy, CHAR(39), '''''' )	
	SET @curCreatedBy = REPLACE( @curCreatedBy, 'Matt Pearson', 'System' )	
	SET @curSubject = REPLACE( @curSubject, CHAR(39), '''''' )
	SET @curDescription = LTRIM( RTRIM( @curDescription ))
	SET @curDescription = REPLACE( @curDescription, CHAR(39), '''''' )
	IF @curType LIKE 'E-mail' 
		BEGIN 
		IF LEN( @curDescription ) > 0
			BEGIN	
			SET @curDescription = DBO.[udf_RemoveTextInAngleBracket]( @curDescription ) 
			SET @curDescription = REPLACE( @curDescription, 'v\:*', '' )
			SET @curDescription = REPLACE( @curDescription, 'o\:*', '' )
			SET @curDescription = REPLACE( @curDescription, 'w\:*', '' )			
			SET @curDescription = REPLACE( @curDescription, '.shape', '' )
			SET @curDescription = DBO.[udf_RemoveTextInCurlyBracket]( @curDescription ) 
			SET @curDescription = DBO.udf_RemoveMTMLcodeInString( @curDescription ) 
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
		SET @sSQL = @sSQL +	' ( [Owner], [Opportunity], [Opportunity ID], [Status], [Invoice Customer]'
		SET @sSQL = @sSQL +	', [Primary Contact], [Opportunity Partner], [Opportunity End User]'
		SET @sSQL = @sSQL +	', [Lead Source], [Inside Sales Support], [Range] ) '
		SET @sSQL = @sSQL +	'VALUES ( '
		SET @sSQL = @sSQL +	'''' + @curOwner + ''', ''' + @curOpportunity + ''''
		SET @sSQL = @sSQL +	', ''' + CONVERT( VARCHAR(50), @curOpportunityID ) + ''''
		SET @sSQL = @sSQL +	', ''' + @curStatus + ''', ''' + @curInvoiceCustomer + ''', ''' + @curPrimaryContact + ''''
		SET @sSQL = @sSQL +	', ''' + @curOpportunityPartner + ''', ''' + @curOpportunityEndUser + ''''
		SET @sSQL = @sSQL +	', ''' + @curLeadSource + ''', ''' + @curInsideSalesSupport + ''', ' + @sRange + ' );'
		BEGIN TRANSACTION O3
		PRINT @sSQL;
		EXEC( @sSQL )
		COMMIT TRANSACTION O3
		SET @iCountOpportunityRow = @iCountOpportunityRow + 1
		END
	
	SET @sSQL = 'UPDATE ' + @sFullNameTable + ' '
	SET @sSQL = @sSQL + 'SET [' + CONVERT( VARCHAR(10), @iColumn ) + ' Created On] = ''' + CONVERT( VARCHAR(100), @curCreatedOn ) + ''' '
	SET @sSQL = @sSQL + ', [' + CONVERT( VARCHAR(10), @iColumn ) + ' Created By] = ''' + @curCreatedBy + ''' '
	SET @sSQL = @sSQL + ', [' + CONVERT( VARCHAR(10), @iColumn ) + ' Type] = ''' + @curType + ''' '
	SET @sSQL = @sSQL + ', [' + CONVERT( VARCHAR(10), @iColumn ) + ' State] = ''' + @curState + ''' '
	SET @sSQL = @sSQL + ', [' + CONVERT( VARCHAR(10), @iColumn ) + ' Subject] = ''' + @curSubject + ''''
	SET @sSQL = @sSQL + ', [' + CONVERT( VARCHAR(10), @iColumn ) + ' Description] = ''' + @curDescription + ''' '
	SET @sSQL = @sSQL + 'FROM ' + @sFullNameTable + ' '
	SET @sSQL = @sSQL + 'WHERE [Opportunity ID] = ''' + CONVERT( VARCHAR(50), @curOpportunityID ) + '''' 	
	SET @sSQL = @sSQL + ' AND [Range] = ' + @sRange + ';' 
	BEGIN TRANSACTION O4
	PRINT @sSQL
	EXEC( @sSQL )
	COMMIT TRANSACTION O4
	SET @iCountActivity  = @iCountActivity  + 1
	PRINT CONVERT( VARCHAR(10), @iCountActivity ) + @sTotalActivity
	FETCH NEXT FROM CurOppAct INTO @curRow, @curOwner, @curOpportunity, @curOpportunityID, @CurStatus
		, @curInvoiceCustomer, @curPrimaryContact, @curOpportunityPartner, @curOpportunityEndUser
		, @curLeadSource, @curInsideSalesSupport, @curCreatedOn, @curCreatedBy, @curType, @curState
		, @curSubject, @curDescription
	END
CLOSE CurOppAct
DEALLOCATE CurOppAct

/*********************************************************************************/
SELECT @iCountZeroOpportunity = COUNT( DISTINCT O.OpportunityID )
	FROM [Productions_MSCRM].[dbo].OpportunityBase AS O
		, [Productions_MSCRM].[dbo].[OpportunityExtensionBase] AS Ox		
		, [Productions_MSCRM].[dbo].[ActivityPointerBase] AS A	
	WHERE O.OpportunityID = Ox.OpportunityID 
			AND O.[CampaignID] IS NOT NULL
			AND O.OpportunityID NOT IN (
				SELECT DISTINCT [Opportunity ID] 
					FROM vw_ActivityOpportunity_run_1 
			);	

PRINT 'Number of Opportunity Has Zero Activities: ' + CONVERT( VARCHAR(10), @iCountZeroOpportunity )

INSERT tblActivityOpportunity
	( 
	[Owner], [Opportunity], [Opportunity ID], [Status]
	, [Invoice Customer], [Primary Contact], [Opportunity Partner]
	,  [Opportunity End User], [Lead Source], [Inside Sales Support] 
	)	
	SELECT DISTINCT ISNULL( Ow2.Name, '' ) AS [Owner]
		, O.Name AS [Opportunity] 
		, CAST( O.OpportunityID AS UNIQUEIDENTIFIER ) AS [Opportunity ID]
		, CASE O.StateCode
			WHEN 0 THEN 'Open'
			WHEN 1 THEN 'Won'
			WHEN 2 THEN 'Lost'
			ELSE CONVERT( VARCHAR(10), O.StateCode ) 
			END AS [Status]
	, ISNULL( O.CustomerIDName, '' ) AS [Invoice Customer] 
	, ISNULL( Ow3.FullName, '' ) AS [Primary Contact]
	, ISNULL( Ow4.Name, '' ) AS [Opportunity Partner] 
	, ISNULL( Ow5.Name, '' ) AS [Opportunity End User] 
	, ISNULL( Ow6.Name, '' ) AS [Lead Source]
	, ISNULL( Ow7.Name, '' ) AS [Inside Sales Support]
	FROM [Productions_MSCRM].[dbo].OpportunityBase AS O
			LEFT JOIN Productions_MSCRM.dbo.OwnerBase AS Ow2 ON Ow2.OwnerID = O.OwnerID
			LEFT JOIN Productions_MSCRM.dbo.CampaignBase AS Ow6	ON Ow6.CampaignID = O.[CampaignID]
		, [Productions_MSCRM].[dbo].[OpportunityExtensionBase] AS Ox		
			LEFT JOIN Productions_MSCRM.dbo.ContactBase AS Ow3 ON Ow3.ContactID = Ox.new_PrimaryContact
			LEFT JOIN Productions_MSCRM.dbo.AccountBase AS Ow4 ON Ow4.AccountID = Ox.new_opportunitypartner
			LEFT JOIN Productions_MSCRM.dbo.AccountBase AS Ow5 ON Ow5.AccountID = Ox.new_opportunityenduser
			LEFT JOIN Productions_MSCRM.dbo.OwnerBase AS Ow7 ON Ow7.OwnerID = Ox.new_InsideSalesSupport
	WHERE O.OpportunityID = Ox.OpportunityID 
			AND O.[CampaignID] IS NOT NULL
			AND O.OpportunityID NOT IN 
			(
				SELECT DISTINCT [Opportunity ID] 
					FROM vw_ActivityOpportunity_run_1 
			)	
	ORDER BY O.Name;

/*********************************************************************************/
PRINT CHAR(13) + 'Time elapsed: ' +  CONVERT( VARCHAR(50), CONVERT( TIME(0),( GETDATE() - @StartTime )))
SELECT 'Activity per Opportunity' AS [Procedure]
	, @iCountOpportunityRow AS [Number of Opportunity-Row]
	, @iCountActivity AS [Number of Activity]
	, @iMaxNumOpp AS [Max num of Activity per Opportunity]
	, @iCountZeroOpportunity AS [Number of Opportunity Has Zero Activities]
	, CONVERT( TIME(0),( GETDATE() - @StartTime )) AS [Time elapsed]
END 