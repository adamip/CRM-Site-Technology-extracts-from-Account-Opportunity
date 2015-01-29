USE [Productions_MSCRM];
GO
/*
	UPDATE Productions_MSCRM.dbo.new_siteExtensionBase
		SET cust_ProductInterest1 = @ProductInterest1, cust_ProductInterest2 = @ProductInterest2, cust_ProductInterest3 = @ProductInterest3
			, cust_ProductInterest4 = @ProductInterest4, cust_ProductInterest5 = @ProductInterest5, cust_ProductInterest6 = @ProductInterest6 
			, cust_Opportunity1GUID = @OpportunityID1, cust_Opportunity2GUID = @OpportunityID2, cust_Opportunity3GUID = @OpportunityID3
			, cust_Opportunity4GUID = @OpportunityID4, cust_Opportunity5GUID = @OpportunityID5, cust_Opportunity6GUID = @OpportunityID6
			, new_hardwareadtechglobal = @HardwareAdtech, new_supporthardware = @HardwareSupport, new_hardwaresupportlevel = @HardwareSupportLevel
			, new_softwaresupport = @SoftwareSupport, new_softwaresupportlevel = @SoftwareSupportLevel
		FROM new_siteExtensionBase 
		WHERE new_AccountGUID = @AccInHand; */
		
ALTER PROCEDURE [dbo].[p_Opportunity_Account_Site_v1] 
AS 
BEGIN 

SET ANSI_NULLS ON;
SET NOCOUNT ON; 
SET QUOTED_IDENTIFIER ON;

DECLARE @StartTime Datetime
SET @StartTime = GETDATE()
PRINT 'Stored Procedure starts at ' + CONVERT( NVARCHAR(50), @StartTime ) + CHAR(13);

/* ******************************************************************************* */

DECLARE @sSQL NVARCHAR(MAX)

BEGIN TRANSACTION S1;
IF EXISTS ( SELECT * FROM INFORMATION_SCHEMA.TABLES   
    WHERE TABLE_CATALOG = 'Productions_MSCRM'
      AND TABLE_SCHEMA = 'dbo'
      AND TABLE_NAME = 'tblOpportunityAccountSite' )
	DROP TABLE Productions_MSCRM.dbo.tblOpportunityAccountSite;

CREATE TABLE Productions_MSCRM.dbo.tblOpportunityAccountSite
	( [RecID] INT IDENTITY(1,1) 
	, [CustomerID] UniqueIdentifier NULL 
	, [OpportunityID] UniqueIdentifier NULL 
	, [ProductInterest] INT NULL 
	, [ModifiedOn] DateTime NULL );
COMMIT TRANSACTION S1;

BEGIN TRANSACTION S2;
INSERT INTO tblOpportunityAccountSite 
	SELECT O.[CustomerID], O.OpportunityID, Oe.new_ProductInterest, O.ModifiedOn
	  FROM OpportunityBase AS O, OpportunityExtensionBase AS Oe
	  WHERE O.OpportunityID = Oe.OpportunityID 
		  AND O.[CustomerID] IS NOT NULL 
		  AND Oe.new_ProductInterest IS NOT NULL
		  AND O.StateCode = 1	/* StateCode = 0 = Open; 1 = Won; 2 = Lost */
	  ORDER BY O.[CustomerID], O.ModifiedOn DESC;
COMMIT TRANSACTION S2;

DECLARE @iCountAcc INT, @iTotalAcc INT, @iCountOpp INT, @iTotalOpp INT, @iUpdate INT;
SET @iCountAcc = 0;
SELECT @iTotalAcc = COUNT( DISTINCT [CustomerID] ) FROM tblOpportunityAccountSite;
SET @iCountOpp = 0;
SELECT @iTotalOpp = COUNT( * ) FROM tblOpportunityAccountSite;
SET @iUpdate = 0;

IF CURSOR_STATUS( 'global','CurOppAcctSite' ) >= -1	/* remove Cursor if already exists */
	BEGIN DEALLOCATE CurOppAcctSite END
DECLARE CurOppAcctSite CURSOR FOR	/* SQL Cursor */
	SELECT [RecID], [CustomerID], [OpportunityID], [ProductInterest], [ModifiedOn] 
		FROM tblOpportunityAccountSite;	

/* ******************************************************************************* */

DECLARE @curRecID INT, @curCustomerID UniqueIdentifier, @curOpportunityID UniqueIdentifier, @curProductInterest INT, @curModifiedOn DateTime;
DECLARE @i INT, @iMAX INT, @isFirst INT, @isLast INT, @AccInHand UniqueIdentifier;
SET @iMAX = 6;	
SET @isFirst = 1;
SET @isLast = 0;	/* 0 means no; 1 means yes; 2 means processing completed. */
DECLARE @ModifiedOn DateTime;
DECLARE @OpportunityID1 UniqueIdentifier, @OpportunityID2 UniqueIdentifier, @OpportunityID3 UniqueIdentifier, 
	@OpportunityID4 UniqueIdentifier, @OpportunityID5 UniqueIdentifier, @OpportunityID6 UniqueIdentifier;
DECLARE @ProductInterest1 INT, @ProductInterest2 INT, @ProductInterest3 INT, @ProductInterest4 INT, @ProductInterest5 INT, @ProductInterest6 INT;
DECLARE @HardwareAdtech INT, @HardwareSupport INT, @HardwareSupportLevel INT, @SoftwareSupport INT, @SoftwareSupportLevel INT;

SET @AccInHand = '00000000-0000-0000-0000-000000000000';
OPEN CurOppAcctSite;

WHILE @isLast < 2
	BEGIN 
	FETCH NEXT FROM CurOppAcctSite INTO @curRecID, @curCustomerID, @curOpportunityID, @curProductInterest, @curModifiedOn;
	IF @@FETCH_STATUS <> 0 
		SET @isLast = 1;
	ELSE
		BEGIN
		SET @i = @i + 1;
		SET @iCountOpp= @iCountOpp + 1;
		END
	/* PRINT '@@FETCH_STATUS = ' + CONVERT( NVARCHAR(10), @@FETCH_STATUS ) + CHAR(9) + '@isFirst = ' + CONVERT( NVARCHAR(10), @isFirst ) + CHAR(9) + '@isLast = ' + CONVERT( NVARCHAR(10), @isLast ) + CHAR(9) + 'Acc = ' + CONVERT( NVARCHAR(50), @AccInHand ) + CHAR(9) + 'Customer ID = ' + CONVERT( NVARCHAR(50), @curCustomerID ); */
	IF @AccInHand <> @curCustomerID OR @isLast = 1 
		BEGIN
		IF @isFirst = 0 AND @AccInHand NOT LIKE '00000000-0000-0000-0000-000000000000' 
			BEGIN
			BEGIN TRANSACTION S3;
			SET @sSQL = 'UPDATE Productions_MSCRM.dbo.new_siteExtensionBase' + CHAR(13) + CHAR(9)
				+ 'SET cust_ProductInterest1 = ' + CONVERT( NVARCHAR(10), @ProductInterest1 )
				+ ', cust_ProductInterest2 = ' + CONVERT( NVARCHAR(10), @ProductInterest2 ) + CHAR(13) + CHAR(9) + CHAR(9)
				+ ', cust_ProductInterest3 = ' + CONVERT( NVARCHAR(10), @ProductInterest3 ) 
				+ ', cust_ProductInterest4 = ' + CONVERT( NVARCHAR(10), @ProductInterest4 ) + CHAR(13) + CHAR(9) + CHAR(9)
				+ ', cust_ProductInterest5 = ' + CONVERT( NVARCHAR(10), @ProductInterest5 )
				+ ', cust_ProductInterest6 = ' + CONVERT( NVARCHAR(10), @ProductInterest6 ) + CHAR(13) + CHAR(9) + CHAR(9)
				+ ', cust_Opportunity1GUID = ''' + CONVERT( NVARCHAR(50), @OpportunityID1 ) + '''' 
				+ ', cust_Opportunity2GUID = ''' + CONVERT( NVARCHAR(50), @OpportunityID2 ) + '''' + CHAR(13) + CHAR(9) + CHAR(9) 
				+ ', cust_Opportunity3GUID = ''' + CONVERT( NVARCHAR(50), @OpportunityID3 ) + ''''  
				+ ', cust_Opportunity4GUID = ''' + CONVERT( NVARCHAR(50), @OpportunityID4 ) + '''' + CHAR(13) + CHAR(9) + CHAR(9) 
				+ ', cust_Opportunity5GUID = ''' + CONVERT( NVARCHAR(50), @OpportunityID5 ) + '''' 
				+ ', cust_Opportunity6GUID = ''' + CONVERT( NVARCHAR(50), @OpportunityID6 ) + '''' + CHAR(13) + CHAR(9) + CHAR(9)
				+ ', new_hardwareadtechglobal = ' + CONVERT( NVARCHAR(10), @HardwareAdtech )
				+ ', new_supporthardware = ' + CONVERT( NVARCHAR(10), @HardwareSupport )
				+ ', new_hardwaresupportlevel = ' + CONVERT( NVARCHAR(20), @HardwareSupportLevel ) + CHAR(13) + CHAR(9) + CHAR(9)
				+ ', new_softwaresupport = ' + CONVERT( NVARCHAR(10), @SoftwareSupport )
				+ ', new_softwaresupportlevel = ' + CONVERT( NVARCHAR(20), @SoftwareSupportLevel ) + CHAR(13) + CHAR(9) 
				+ 'FROM new_siteExtensionBase' + CHAR(13) + CHAR(9) + CHAR(9)
				+ 'WHERE new_AccountGUID = ''' + CONVERT( NVARCHAR(50), @AccInHand ) + ''';';
			/* PRINT 'SQL Statement: ' + @sSQL; */
			EXECUTE sp_executesql @sSQL;
			SET @iUpdate = @iUpdate + 1;
			/* PRINT 'Updating Acc = ' + CONVERT( NVARCHAR(50), @AccInHand ) + CHAR(9) + '@iUpdate = ' + CONVERT( NVARCHAR(50), @iUpdate ); */
			COMMIT TRANSACTION S3;
			END
		IF @isLast = 1
			SET @isLast = 2;
		ELSE
			BEGIN /* Reset variables to prepare for the next Set record fetch */
			SET @i = 1;
			SET @iCountAcc = @iCountAcc + 1;
			SET @OpportunityID1 = NULL; SET @OpportunityID2 = NULL; SET @OpportunityID3 = NULL;
			SET @OpportunityID4 = NULL; SET @OpportunityID5 = NULL; SET @OpportunityID6 = NULL;
			SET @ProductInterest1 = 0; SET @ProductInterest2 = 0; SET @ProductInterest3 = 0;
			SET @ProductInterest4 = 0; SET @ProductInterest5 = 0; SET @ProductInterest6 = 0;
			SET @HardwareAdtech = 0; SET @HardwareSupport = 0; SET @HardwareSupportLevel = 0;
			SET @SoftwareSupport = 0; SET @SoftwareSupportLevel = 0;
			END
		END
	IF @isFirst = 1 OR @isLast = 0
		BEGIN
		SET @AccInHand = @curCustomerID;
		IF @isFirst = 1 
			SET @isFirst = 0;
		IF @i <= @iMAX /* number of opportunities per Account to be handled is @iMax */
			BEGIN
			BEGIN TRANSACTION S4;
			IF @i = 1 BEGIN SET @ProductInterest1 = @curProductInterest; SET @OpportunityID1 = @curOpportunityID; SET @ModifiedOn = @curModifiedOn; END
			ELSE IF @i = 2 BEGIN SET @ProductInterest2 = @curProductInterest; SET @OpportunityID2 = @curOpportunityID; END
			ELSE IF @i = 3 BEGIN SET @ProductInterest3 = @curProductInterest; SET @OpportunityID3 = @curOpportunityID; END
			ELSE IF @i = 4 BEGIN SET @ProductInterest4 = @curProductInterest; SET @OpportunityID4 = @curOpportunityID; END
			ELSE IF @i = 5 BEGIN SET @ProductInterest5 = @curProductInterest; SET @OpportunityID5 = @curOpportunityID; END
			ELSE IF @i = 6 BEGIN SET @ProductInterest6 = @curProductInterest; SET @OpportunityID6 = @curOpportunityID; END
			IF @curProductInterest = 279640013 BEGIN /* Hardware Install & Configuration */ SET @HardwareAdtech = 1; END
			ELSE IF @curProductInterest = 279640014 BEGIN /* Support */ SET @HardwareSupport = 1; SET @SoftwareSupport = 1; SET @HardwareSupportLevel = 279640002; SET @SoftwareSupportLevel = 279640000; END
			ELSE IF @curProductInterest = 279640033 BEGIN /* SUPR-Gold (Hardware) */ SET @HardwareSupport = 1; SET @HardwareSupportLevel = 279640000; END
			ELSE IF @curProductInterest = 279640032
				BEGIN /* SUPR-Silver (Hardware) */
				SET @HardwareSupport = 1;
				IF @HardwareSupportLevel <> 279640000 SET @HardwareSupportLevel = 279640002; 
				END
			ELSE IF @curProductInterest = 279640031 BEGIN /* SUPR-WSAS-247 Gold (Software) */ SET @SoftwareSupport = 1; SET @SoftwareSupportLevel = 279640001; END
			ELSE IF @curProductInterest = 279640035  
				BEGIN /* SUPR-WSAS-BHR Silver (Software) */ 
				SET @SoftwareSupport = 1;
				IF @SoftwareSupportLevel <> 279640001 SET @SoftwareSupportLevel = 279640000; 
				END
			PRINT 'Processing ' + CONVERT( NVARCHAR(10), @iCountAcc ) + ' of ' + CONVERT( NVARCHAR(10), @iTotalAcc ) + ' Accounts.' 
				+ CHAR(9) + 'Processing ' + CONVERT( NVARCHAR(10), @iCountOpp ) + ' of ' + CONVERT( NVARCHAR(10), @iTotalOpp ) + ' Opportunities ( #' +  CONVERT( NVARCHAR(10), @i ) + ' ).'
				+ CHAR(13) + CHAR(9) + 'Acc ' + CONVERT( NVARCHAR(50), @curCustomerID ) + ', Opp ' + CONVERT( NVARCHAR(50), @curOpportunityID ) 
				+ ', pi ' + CONVERT( NVARCHAR(20), @curProductInterest );
			COMMIT TRANSACTION S4;
			END
		ELSE
			PRINT 'Processing ' + CONVERT( NVARCHAR(10), @iCountAcc ) + ' of ' + CONVERT( NVARCHAR(10), @iTotalAcc ) + ' Accounts which has more than ' 
				+ CONVERT( NVARCHAR(10), @iMAX ) + ' Opportunities ( #' + CONVERT( NVARCHAR(10), @i ) + ' ).' 
				+ CHAR(9) + 'Processing ' + CONVERT( NVARCHAR(10), @iCountOpp ) + ' of ' + CONVERT( NVARCHAR(10), @iTotalOpp ) + ' Opportunities.'
				+ CHAR(13) + CHAR(9) + 'Acc ' + CONVERT( NVARCHAR(50), @curCustomerID ) + ', Opp ' + CONVERT( NVARCHAR(50), @curOpportunityID ) 
				+ ', pi ' + CONVERT( NVARCHAR(20), @curProductInterest );
		END
	END

/* Clean up */	
CLOSE CurOppAcctSite;
DEALLOCATE CurOppAcctSite;

IF EXISTS ( SELECT * FROM INFORMATION_SCHEMA.TABLES   
    WHERE TABLE_CATALOG = 'Productions_MSCRM'
      AND TABLE_SCHEMA = 'dbo'
      AND TABLE_NAME = 'tblOpportunityAccountSite' )
	DROP TABLE Productions_MSCRM.dbo.tblOpportunityAccountSite;
	
/* ******************************************************************************* */
PRINT CHAR(13) + 'Number of Site records updated: ' + CONVERT( NVARCHAR(10), @iUpdate );
PRINT CHAR(13) + 'Time elapsed: ' +  CONVERT( NVARCHAR(50), CONVERT( TIME(0),( GETDATE() - @StartTime )))
SELECT 'Opportunity Won Extracts Account Updates Site' AS [Procedure], CONVERT( TIME(0),( GETDATE() - @StartTime )) AS [Time elapsed]
	, 'Number of SITE Record updated' AS 'Record', CONVERT( NVARCHAR(10), @iUpdate ) AS 'Count';

END 
/* End Of File ******************************************************************* */