USE [Productions_MSCRM]

/****** Object:  View [dbo].[vw_ActivityContact_run_1]    Script Date: 10/23/2014 15:41:59 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

ALTER VIEW [dbo].[vw_ActivityContact_run_1]
AS 

SELECT ISNULL( Ow2.Name, '' ) AS [Owner]
	, LEFT( A.[RegardingObjectIDName], 300 ) AS [Contact]
	, C.[ContactID] AS [Contact ID]
	, Ow3.[Name] AS [Account]
	, Cpg.[Name] AS [Lead Source]
	, ISNULL( C.[Description], '' ) AS [Contact Description]
	, A.[CreatedOn] AS [Created On]
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
	, ISNULL( A.[Description], '' ) AS [Activity Description]	
	FROM ActivityPointerBase AS A
			LEFT JOIN OwnerBase AS Ow1 ON Ow1.OwnerID = A.OwnerID 
		, [ContactBase] AS C
			LEFT JOIN OwnerBase AS Ow2 ON Ow2.OwnerID = C.OwnerID 
			LEFT JOIN AccountBase AS Ow3 ON Ow3.AccountID = C.ParentCustomerID
		, [ContactExtensionBase] AS Cext	
			LEFT JOIN CampaignBase AS Cpg ON Cpg.CampaignID = Cext.[new_LeadSourceName] 
	WHERE A.RegardingObjectID = C.ContactID 
		AND C.ContactID = Cext.ContactID
		AND Cext.[new_LeadSourceName] IS NOT NULL
		AND A.ActivityTypeCode IN  ( 4201, 4202, 4204, 4206, 4207, 4210, 4212, 4214, 4251, 4400, 4401, 4402, 4403, 4404 )
		AND Ow1.Name NOT LIKE 'CRM Router%'		/* Created By is not CRM Router */
		AND Ow1.Name NOT LIKE 'Libby Lasseigne%';


GO


