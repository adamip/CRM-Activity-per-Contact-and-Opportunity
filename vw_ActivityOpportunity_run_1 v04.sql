USE [Productions_MSCRM]
GO

/****** Object:  View [dbo].[vw_ActivityOpportunity_run_1]    Script Date: 10/23/2014 15:50:58 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO





ALTER VIEW [dbo].[vw_ActivityOpportunity_run_1]
AS 
SELECT ISNULL( Ow2.Name, '' ) AS [Owner] 
	, LEFT( A.RegardingObjectIDName, 300 ) AS [Opportunity] 
	, A.RegardingObjectID AS [Opportunity ID] 
	, CASE O.StateCode
		WHEN 0 THEN 'Open'
		WHEN 1 THEN 'Won'
		WHEN 2 THEN 'Lost'
		ELSE CONVERT( VARCHAR( 20 ), O.StateCode ) 
		END AS [Status]
	, LEFT( ISNULL( O.CustomerIDName, '' ), 160 ) AS [Invoice Customer] 
	, ISNULL( Ow3.FullName, '' ) AS [Primary Contact]
	, ISNULL( Ow4.Name, '' ) AS [Opportunity Partner] 
	, ISNULL( Ow5.Name, '' ) AS [Opportunity End User] 
	, ISNULL( Ow6.Name, '' ) AS [Lead Source]
	, ISNULL( Ow7.Name, '' ) AS [Inside Sales Support]
	, A.CreatedOn AS [Created On]
	, ISNULL( Ow1.Name, '' ) AS [Created By]
	, CASE A.ActivityTypeCode
		WHEN 4201 THEN 'Appointment'
		WHEN 4202 THEN 'E-mail'
		WHEN 4204 THEN 'Fax'
		WHEN 4206 THEN 'Incident Resolution'
		WHEN 4207 THEN 'Letter'
		WHEN 4210 THEN 'Phone Call'
		WHEN 4212 THEN 'Task'
		WHEN 4214 THEN 'Service Appointment'
		WHEN 4251 THEN 'Recurring Appointment'
		WHEN 4400 THEN 'Campaign'
		WHEN 4401 THEN 'Campaign Response'
		WHEN 4402 THEN 'Campaign Activity'
		WHEN 4403 THEN 'Campaign Item'
		WHEN 4404 THEN 'Campaign Activity Item'
		ELSE 'Unknown'
		END	AS [Type]
	, CASE A.StateCode
		WHEN 0 THEN 'Open'
		WHEN 1 THEN 'Completed'
		WHEN 2 THEN 'Canceled'
		WHEN 3 THEN 'Scheduled'		
		ELSE 'Unknown'
		END	AS [State]	
	, A.[Subject]	
	, A.[Description]
	FROM ActivityPointerBase AS A
		LEFT JOIN OwnerBase AS Ow1 ON Ow1.OwnerID = A.OwnerID	
	, [OpportunityBase] AS O
		LEFT JOIN OwnerBase AS Ow2 ON Ow2.OwnerID = O.OwnerID
		LEFT JOIN CampaignBase AS Ow6 ON Ow6.CampaignID = O.[CampaignID] 
	, [OpportunityExtensionBase] AS Ox		
		LEFT JOIN ContactBase AS Ow3 ON Ow3.ContactID = Ox.new_PrimaryContact
		LEFT JOIN AccountBase AS Ow4 ON Ow4.AccountID = Ox.new_opportunitypartner
		LEFT JOIN AccountBase AS Ow5 ON Ow5.AccountID = Ox.new_opportunityenduser
		LEFT JOIN OwnerBase AS Ow7 ON Ow7.OwnerID = Ox.new_InsideSalesSupport
	WHERE A.RegardingObjectID = O.OpportunityID 
		AND O.OpportunityID = Ox.OpportunityID 
		AND O.[CampaignID] IS NOT NULL		
		AND A.ActivityTypeCode IN ( 4201, 4202, 4204, 4206, 4207, 4210, 4212, 4214, 4251, 4400, 4401, 4402, 4403, 4404 ) 
		AND Ow1.Name NOT LIKE 'CRM Router%'	 /* Created By is not CRM Router */
		AND Ow1.Name NOT LIKE 'Libby Lasseigne%';


GO


