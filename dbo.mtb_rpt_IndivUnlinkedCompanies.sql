USE [WorkingDB]
GO
/****** Object:  StoredProcedure [dbo].[mtb_rpt_IndivUnlinkedCompanies]    Script Date: 04/02/2010 16:47:00 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

/***********************************************************************************
* NAME: mtb_ProcedureName
*
* DESC:   
*
* REVISION HISTORY:
* 	DATE        WHO				ISSUE		DETAIL
*   ----------	-------------	---------	----------------------------
*	2010-04-02	CMurphy						Identifies Individuals in Onyx that are not 
*       									linked to a Company
************************************************************************************/

ALTER PROCEDURE [dbo].[mtb_rpt_IndivUnlinkedCompanies]

AS
BEGIN

SET NOCOUNT ON;


--Find all individual records that are not linked to a company

CREATE TABLE #NotLinked(iIndividualId INT NULL, vchFirstName NVARCHAR(255) NULL, vchLastName NVARCHAR(255) NULL, 
iCompanyId INT null, vchCompanyName NVARCHAR(255) NULL, chCountryCode NVARCHAR(10) null, vchuser6 INT,
salesTerr NVARCHAR(10)null, vchuser7 INT, insertdt DATETIME NULL, istatusid INT, iSiteId INT, iUserTypeId INT, tiRecordStatus TINYINT)

INSERT INTO #NotLinked (iIndividualId, vchFirstName, vchLastName, 
iCompanyId, vchCompanyName, chCountryCode, vchuser6, salesTerr, vchuser7, insertdt, iStatusId, iSiteId, iUserTypeId, tiRecordStatus)

SELECT [iIndividualId], vchFirstName, vchLastName, 
iCompanyId , vchCompanyName, chCountryCode, vchUser6, SalesTeam, vchUser7, dtInsertDate, iStatusId, iSiteId, iUserTypeId, tiRecordStatus
  FROM [Onyx].[dbo].[Individual] I LEFT outer JOIN Onyx.dbo.[CSuSalesPerson] C ON vchUser6 = C.salesteamnum
  WHERE tiRecordStatus = 1 
  AND iCompanyId =0




SELECT DISTINCT 
N.vchFirstName, N.vchLastName, N.vchCompanyName, N.iCompanyId, N.chCountryCode, N.salesTerr, N.insertdt
                                          
FROM         #NotLinked  AS N WITH (NOLOCK) LEFT OUTER JOIN
                      Onyx.dbo.ReferenceParameters AS r ON N.vchUser6 = r.iParameterId LEFT OUTER JOIN
                      Onyx.dbo.ReferenceParameters AS m ON N.vchUser7 = m.iParameterId LEFT OUTER JOIN
                      Onyx.dbo.CSuTerritoryAssignment AS c ON N.chCountryCode = c.chCountryCode LEFT OUTER JOIN
                      Onyx.dbo.CustomerCampaign AS g ON N.iIndividualId = g.iOwnerId LEFT OUTER JOIN
                      Onyx.dbo.TrackingCode AS t ON g.iTrackingId = t.iTrackingId LEFT OUTER JOIN
                      Onyx.dbo.CustomerCampaignAction AS a ON g.iCampaignId = a.iCampaignId LEFT OUTER JOIN
                      Onyx.dbo.Company AS co ON N.iCompanyId = co.iCompanyId
WHERE 
   
   (N.tiRecordStatus = 1) AND (N.iIndividualId IN
                          (SELECT DISTINCT N.iIndividualId
                            FROM           #NotLinked  AS N WITH (NOLOCK) LEFT OUTER JOIN
                                                   Onyx.dbo.CustomerCampaign AS c ON N.iIndividualId = c.iownerid LEFT OUTER JOIN
                                                   Onyx.dbo.TrackingCode AS t ON c.iTrackingId = t.iTrackingId
                            WHERE      (t.chCampaignCode = 'webinars') AND (a.iActionId IN (8)))) OR
										(N.tiRecordStatus = 1) AND (co.iCompanyTypeCode = 102430) OR
                      (N.tiRecordStatus = 1) AND (N.iUserTypeId = 101) OR
                      (N.tiRecordStatus = 1) 
                      
                      AND (N.iIndividualId IN
                          (SELECT     iOwnerId
                            FROM          Onyx.dbo.CustomerProfile WITH (NOLOCK)
                            WHERE      (iSurveyId = 23))) OR
                      (N.tiRecordStatus = 1) AND (N.vchCompanyName LIKE 'MTB%') OR
                      (N.tiRecordStatus = 1) 
                      
                      --Individual has an Incident
                      AND (N.iIndividualId IN
                          (SELECT     iContactId
                            FROM          Onyx.dbo.Incident WITH (NOLOCK)
                            WHERE      (tiRecordStatus = 1))) OR
                      (N.tiRecordStatus = 1) 
                      
                      --Individual has a Customer Product
                      AND (N.iIndividualId IN
                          (SELECT     iOwnerId
                            FROM          Onyx.dbo.CustomerProduct WITH (NOLOCK)
                            WHERE      (tiRecordStatus = 1))) OR
                      (N.tiRecordStatus = 1) 
                      
                    
                      AND (N.iIndividualId IN
                          (SELECT     iContactID
                            FROM          Onyx.dbo.Contact WITH (NOLOCK)
                            WHERE      (tiRecordStatus = 1) AND (iContactTypeId IN (350, 100, 104, 341, 356, 336, 340, 114, 105, 95, 325, 362, 360, 333, 344, 346, 342, 108, 369, 
                                                   219, 268, 260, 211, 223, 316, 364, 98, 208, 365, 109, 97)))) OR
                      (N.tiRecordStatus = 1) AND (N.iUserTypeId = 100133) OR
                      (N.tiRecordStatus = 1) AND (t.chCampaignCode = 'CandTester') AND (a.iActionId IN (1, 108))
   
   
   
   
   
   
     
   
       
       --SELECT * FROM #NotLinked
       DROP TABLE #NotLinked

END
 GO


GRANT EXEC ON [dbo].[mtb_rpt_IndivUnlinkedCompanies] TO DEVELOPER
GRANT EXEC ON [dbo].[mtb_rpt_IndivUnlinkedCompanies] TO reportuser

GO
