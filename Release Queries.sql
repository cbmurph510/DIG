USE MTB_MSCRM

-- DELIVER IN CSV FORMAT WITH THE FOLLOWING FIELDS - CRM ID, NAME, EMAIL, COMPANY, COUNTRY, CUSTOMER TYPE, PRODUCT MAIL FLAG AND IMPLIED CONSENT

DROP TABLE #Query1, #PT_ACAD, #PT_COMM, #ES_COMM, #ES_ACAD, #EN_COMM, #EN_ACAD, #FR_ACAD, #FR_COMM



/**********************************************************************************************************/
/*                 QUERY 1 - ALL Active MTB LC's from the Renewals Console    
/*  CAM - I fixed the create table and select statement in the INC query and I added a join to the Account table - you can modify the rest the same way */                                    */                                   
/**********************************************************************************************************/

CREATE TABLE #Query1(accountid UNIQUEIDENTIFIER,
accountName NVARCHAR(255),
MTBId INT,
Email NVARCHAR(100),
FirstName NVARCHAR(25), 
LastName NVARCHAR(50),
Country NVARCHAR(4),
SiteIdName NVARCHAR(10),
ProductUpdateFlag NVARCHAR(15)
,CustomerType NVARCHAR(15)
,impliedComsent bit
)

INSERT INTO #Query1
SELECT DISTINCT cb.ParentCustomerId, A.Name, ceb.MTB_ID_Search, CB.EMailAddress1,CB.FirstName, CB.LastName,CAB.Country,'INC',
CASE WHEN CEB.MTB_emailpref_productupdates = 908600002 THEN 'Unsubscribed' 
	 WHEN CEB.MTB_emailpref_productupdates = 908600001 THEN 'Subscribed' 
	 WHEN CEB.MTB_emailpref_productupdates = 908600000 THEN 'Not Set' END AS 'ProductUpdateFlag'
,CASE WHEN CB.CustomerTypeCode = 101970 THEN 'ACAD' 
	 WHEN CB.CustomerTypeCode = 101971 THEN 'COMM' END AS 'CustomerType'
,CASE WHEN ceb.MTB_impliedconsent = 1 THEN 'Yes'
	WHEN ceb.MTB_impliedconsent = 0 THEN 'No' END AS 'ImpliedConsent'
	 


/*INC*/
FROM INC.INC.dbo.csiActiveContracts l   
INNER JOIN INC.INC.dbo.SOP10100 ren ON l.SOPNUMBE = ren.SOPNUMBE  
INNER JOIN MTB_MSCRM.dbo.ContactExtensionBase CEB 
ON CEB.MTB_ID_Search = CAST(substring(ren.INTEGRATIONID,1,(case when charindex('-',ren.INTEGRATIONID) > 0 then charindex('-',ren.INTEGRATIONID)-1 else len(ren.INTEGRATIONID) end)) AS INTEGER)
INNER JOIN MTB_MSCRM.dbo.ContactBase CB ON CEB.ContactId = CB.ContactId
INNER JOIN MTB_MSCRM.dbo.CustomerAddressBase CAB ON CAB.ParentId = CB.ContactId
LEFT JOIN dbo.Account A ON A.AccountId = cb.ParentCustomerId /*CAM I JOINED IN THE ACCOUNT TABLE HERE TO GET ACCT NAME*/
--INNER JOIN dbo.MTB_customerproduct CP ON CB.ContactId = CP.MTB_ContactId
WHERE  (ITEMNMBR LIKE '%170A%'OR ITEMNMBR LIKE '%E010A%')--Both  MTB 17 / MTB Express suite
AND CAB.AddressNumber = 1 

AND CAB.Country NOT IN ('DE', 'KR', 'IN', 'CN', 'JP', 'TH', 'TR', 'TW', 'SG') --Exclude Partner Territories
AND ren.SOPType = 1 AND INTEGRATIONSOURCE = 0 
AND CB.StateCode = 0 --Contact is active
AND CB.CustomerTypeCode in(101970, 101971) --Contact is Academic

--Contact is LC
AND (CEB.ContactId IN
(SELECT DISTINCT Record1Id  
From MTB_MSCRM.dbo.ConnectionBase
WHERE Record1RoleId = 'C705E256-6977-E111-94EF-00155D03EC06' AND StateCode = 0)
)
SELECT * from #query1
/*LTD*/
INSERT INTO #Query1
SELECT DISTINCT cb.ParentCustomerId, CB.EMailAddress1,CB.FirstName, CB.LastName,CAB.Country,'LTD',
CASE WHEN CEB.MTB_emailpref_productupdates = 908600002 THEN 'Unsubscribed' 
	 WHEN CEB.MTB_emailpref_productupdates = 908600001 THEN 'Subscribed' 
	 WHEN CEB.MTB_emailpref_productupdates = 908600000 THEN 'Not Set' END AS 'ProductUpdateFlag',
CASE WHEN CB.CustomerTypeCode = 101970 THEN 'ACAD' 
	 WHEN CB.CustomerTypeCode = 101971 THEN 'COMM' END AS 'CustomerType'

FROM LTD.LTD.dbo.csiActiveContracts l   
INNER JOIN LTD.LTD.dbo.SOP10100 ren ON l.SOPNUMBE = ren.SOPNUMBE  
INNER JOIN MTB_MSCRM.dbo.ContactExtensionBase CEB 
ON CEB.MTB_ID_Search = CAST(substring(ren.INTEGRATIONID,1,(case when charindex('-',ren.INTEGRATIONID) > 0 then charindex('-',ren.INTEGRATIONID)-1 else len(ren.INTEGRATIONID) end)) AS INTEGER)
INNER JOIN MTB_MSCRM.dbo.ContactBase CB ON CEB.ContactId = CB.ContactId
INNER JOIN MTB_MSCRM.dbo.CustomerAddressBase CAB ON CAB.ParentId = CB.ContactId
--INNER JOIN dbo.MTB_customerproduct CP ON CB.ContactId = CP.MTB_ContactId
WHERE  (ITEMNMBR LIKE '%170A%'OR ITEMNMBR LIKE '%E010A%')--Both  MTB 17 / MTB Express suite
AND CAB.AddressNumber = 1 

AND CAB.Country NOT IN ('DE', 'KR', 'IN', 'CN', 'JP', 'TH', 'TR', 'TW', 'SG') --Exclude Partner Territories
AND ren.SOPType = 1 AND INTEGRATIONSOURCE = 0 
AND CB.StateCode = 0 --Contact is active
AND CB.CustomerTypeCode in (101970,101971) --Contact is Academic

--Contact is LC
AND (CEB.ContactId IN
(SELECT DISTINCT Record1Id  
From MTB_MSCRM.dbo.ConnectionBase
WHERE Record1RoleId = 'C705E256-6977-E111-94EF-00155D03EC06' AND StateCode = 0)
)


/*PTY*/
INSERT INTO #Query1
SELECT DISTINCT cb.ParentCustomerId, CB.EMailAddress1,CB.FirstName, CB.LastName,CAB.Country,'PTY',
CASE WHEN CEB.MTB_emailpref_productupdates = 908600002 THEN 'Unsubscribed' 
	 WHEN CEB.MTB_emailpref_productupdates = 908600001 THEN 'Subscribed' 
	 WHEN CEB.MTB_emailpref_productupdates = 908600000 THEN 'Not Set' END AS 'ProductUpdateFlag',
CASE WHEN CB.CustomerTypeCode = 101970 THEN 'ACAD' 
	 WHEN CB.CustomerTypeCode = 101971 THEN 'COMM' END AS 'CustomerType'

FROM PTY.PTY.dbo.csiActiveContracts l   
INNER JOIN PTY.PTY.dbo.SOP10100 ren ON l.SOPNUMBE = ren.SOPNUMBE  
INNER JOIN MTB_MSCRM.dbo.ContactExtensionBase CEB 
ON CEB.MTB_ID_Search = CAST(substring(ren.INTEGRATIONID,1,(case when charindex('-',ren.INTEGRATIONID) > 0 then charindex('-',ren.INTEGRATIONID)-1 else len(ren.INTEGRATIONID) end)) AS INTEGER)
INNER JOIN MTB_MSCRM.dbo.ContactBase CB ON CEB.ContactId = CB.ContactId
INNER JOIN MTB_MSCRM.dbo.CustomerAddressBase CAB ON CAB.ParentId = CB.ContactId
--INNER JOIN dbo.MTB_customerproduct CP ON CB.ContactId = CP.MTB_ContactId
WHERE  (ITEMNMBR LIKE '%170A%'OR ITEMNMBR LIKE '%E010A%')--Both  MTB 17 / MTB Express suite
AND CAB.AddressNumber = 1 

AND CAB.Country NOT IN ('DE', 'KR', 'IN', 'CN', 'JP', 'TH', 'TR', 'TW', 'SG') --Exclude Partner Territories
AND ren.SOPType = 1 AND INTEGRATIONSOURCE = 0 
AND CB.StateCode = 0 --Contact is active
AND CB.CustomerTypeCode in (101970, 101971) --Contact is Academic

--Contact is LC
AND (CEB.ContactId IN
(SELECT DISTINCT Record1Id  
From MTB_MSCRM.dbo.ConnectionBase
WHERE Record1RoleId = 'C705E256-6977-E111-94EF-00155D03EC06' AND StateCode = 0)
)

/*SARL*/
INSERT INTO #Query1
SELECT DISTINCT cb.ParentCustomerId, CB.EMailAddress1,CB.FirstName, CB.LastName,CAB.Country,'SARL',
CASE WHEN CEB.MTB_emailpref_productupdates = 908600002 THEN 'Unsubscribed' 
	 WHEN CEB.MTB_emailpref_productupdates = 908600001 THEN 'Subscribed' 
	 WHEN CEB.MTB_emailpref_productupdates = 908600000 THEN 'Not Set' END AS 'ProductUpdateFlag',
CASE WHEN CB.CustomerTypeCode = 101970 THEN 'ACAD' 
	 WHEN CB.CustomerTypeCode = 101971 THEN 'COMM' END AS 'CustomerType'

FROM SARL.SARL.dbo.csiActiveContracts l   
INNER JOIN SARL.SARL.dbo.SOP10100 ren ON l.SOPNUMBE = ren.SOPNUMBE  
INNER JOIN MTB_MSCRM.dbo.ContactExtensionBase CEB 
ON CEB.MTB_ID_Search = CAST(substring(ren.INTEGRATIONID,1,(case when charindex('-',ren.INTEGRATIONID) > 0 then charindex('-',ren.INTEGRATIONID)-1 else len(ren.INTEGRATIONID) end)) AS INTEGER)
INNER JOIN MTB_MSCRM.dbo.ContactBase CB ON CEB.ContactId = CB.ContactId
INNER JOIN MTB_MSCRM.dbo.CustomerAddressBase CAB ON CAB.ParentId = CB.ContactId
--INNER JOIN dbo.MTB_customerproduct CP ON CB.ContactId = CP.MTB_ContactId
WHERE  (ITEMNMBR LIKE '%170A%'OR ITEMNMBR LIKE '%E010A%')--Both  MTB 17 / MTB Express suite
AND CAB.AddressNumber = 1 

AND CAB.Country NOT IN ('DE', 'KR', 'IN', 'CN', 'JP', 'TH', 'TR', 'TW', 'SG') --Exclude Partner Territories
AND ren.SOPType = 1 AND INTEGRATIONSOURCE = 0 
AND CB.StateCode = 0 --Contact is active
AND CB.CustomerTypeCode IN (101970,101971) --Contact is Academic

--Contact is LC
AND (CEB.ContactId IN
(SELECT DISTINCT Record1Id  
From MTB_MSCRM.dbo.ConnectionBase
WHERE Record1RoleId = 'C705E256-6977-E111-94EF-00155D03EC06' AND StateCode = 0)
)


SELECT * FROM #Query1

/*Portuguese*/
CREATE TABLE #PT_ACAD(
Email NVARCHAR(100),
FirstName NVARCHAR(25), 
LastName NVARCHAR(50),
Country NVARCHAR(4),
ProductUpdateFlag NVARCHAR(15)
)

INSERT INTO #PT_ACAD
SELECT Email, FirstName ,LastName, Country, ProductUpdateFlag FROM #Query1
WHERE Country = 'BR'AND SiteIdName = 'INC' AND CustomerType = 'ACAD'

SELECT * FROM #PT_ACAD

CREATE TABLE #PT_COMM(
Email NVARCHAR(100),
FirstName NVARCHAR(25), 
LastName NVARCHAR(50),
Country NVARCHAR(4),
ProductUpdateFlag NVARCHAR(15)
)

INSERT INTO #PT_COMM
SELECT Email, FirstName ,LastName, Country, ProductUpdateFlag FROM #Query1
WHERE Country = 'BR'AND SiteIdName = 'INC' AND CustomerType = 'COMM'

SELECT * FROM #PT_COMM


/*French*/
CREATE TABLE #FR_ACAD(
Email NVARCHAR(100),
FirstName NVARCHAR(25), 
LastName NVARCHAR(50),
Country NVARCHAR(4),
ProductUpdateFlag NVARCHAR(15)
)
INSERT INTO #FR_ACAD
SELECT Email, FirstName ,LastName, Country, ProductUpdateFlag FROM #Query1
WHERE Country = 'FR' AND SiteIdName = 'LTD' AND CustomerType = 'ACAD'

SELECT * FROM #FR_ACAD


/*French*/
CREATE TABLE #FR_COMM(
Email NVARCHAR(100),
FirstName NVARCHAR(25), 
LastName NVARCHAR(50),
Country NVARCHAR(4),
ProductUpdateFlag NVARCHAR(15)
)
INSERT INTO #FR_COMM
SELECT Email, FirstName ,LastName, Country, ProductUpdateFlag FROM #Query1
WHERE Country = 'FR' AND SiteIdName = 'LTD' AND CustomerType = 'COMM'

SELECT * FROM #FR_COMM


/*Spanish*/
CREATE TABLE #ES_ACAD(
Email NVARCHAR(100),
FirstName NVARCHAR(25), 
LastName NVARCHAR(50),
Country NVARCHAR(4),
ProductUpdateFlag NVARCHAR(15)
)
INSERT INTO #ES_ACAD
SELECT Email, FirstName ,LastName, Country, ProductUpdateFlag FROM #Query1
WHERE (SiteIdName = 'INC' AND Country in ('MX', 'AR', 'BZ', 'BO', 'CL', 'CO', 'CR', 'EC', 'SV', 'GT', 'HN', 'NI', 'PA', 'PY', 'PE', 'UY', 'VE')
OR Country = 'ES') AND CustomerType = 'ACAD'

SELECT * FROM #ES_ACAD

/*Spanish*/
CREATE TABLE #ES_COMM(
Email NVARCHAR(100),
FirstName NVARCHAR(25), 
LastName NVARCHAR(50),
Country NVARCHAR(4),
ProductUpdateFlag NVARCHAR(15)
)
INSERT INTO #ES_COMM
SELECT Email, FirstName ,LastName, Country, ProductUpdateFlag FROM #Query1
WHERE (SiteIdName = 'INC' AND Country in ('MX', 'AR', 'BZ', 'BO', 'CL', 'CO', 'CR', 'EC', 'SV', 'GT', 'HN', 'NI', 'PA', 'PY', 'PE', 'UY', 'VE')
OR Country = 'ES')
 AND CustomerType = 'COMM'
SELECT * FROM #ES_COMM

/*ENGLISH*/

CREATE TABLE #EN_ACAD(
Email NVARCHAR(100),
FirstName NVARCHAR(25), 
LastName NVARCHAR(50),
Country NVARCHAR(4),
ProductUpdateFlag NVARCHAR(15)
)
INSERT INTO #EN_ACAD
SELECT Email, FirstName ,LastName, Country, ProductUpdateFlag FROM #Query1
WHERE Email NOT IN (SELECT Email FROM #PT_ACAD) 
AND Email NOT IN (SELECT Email FROM #FR_ACAD) 
AND Email NOT IN (SELECT Email FROM #ES_ACAD) 
 AND CustomerType = 'ACAD'
SELECT * FROM #EN_ACAD

/*ENGLISH*/

CREATE TABLE #EN_COMM(
Email NVARCHAR(100),
FirstName NVARCHAR(25), 
LastName NVARCHAR(50),
Country NVARCHAR(4),
ProductUpdateFlag NVARCHAR(15)
)
INSERT INTO #EN_COMM
SELECT Email, FirstName ,LastName, Country, ProductUpdateFlag FROM #Query1
WHERE Email NOT IN (SELECT Email FROM #PT_COMM) 
AND Email NOT IN (SELECT Email FROM #FR_COMM) 
AND Email NOT IN (SELECT Email FROM #ES_COMM) 
 AND CustomerType = 'COMM'
SELECT * FROM #EN_COMM

/**********************************************************************************************************/
/*                 QUERY 2 - Users who are part of MUL licenses above    
/*CAM you can modify the Create Table and select the same as query 1 above, account is already joined in this query */                                       */
/**********************************************************************************************************/



/*Query 2 - Contacts from the Schools in the query above*/
DROP TABLE #MtbSchool

CREATE Table #MtbSchool(
email NVARCHAR(255),
FirstName NVARCHAR(255)
,LastName NVARCHAR(255)
,Address1_Country NVARCHAR(255)
,ProductUpdateFlag NVARCHAR(15)
,ImpliedConsentFlag NVARCHAR(15)
)
INSERT INTO #MtbSchool
        
SELECT DISTINCT CB.EMailAddress1,CB.FirstName, CB.LastName,CB.Address1_Country,
CASE WHEN CB.MTB_emailpref_productupdates = 908600002 THEN 'Unsubscribed' 
	 WHEN CB.MTB_emailpref_productupdates = 908600001 THEN 'Subscribed' 
	 WHEN CB.MTB_emailpref_productupdates = 908600000 THEN 'Not Set' END AS 'ProductUpdateFlag',
CASE WHEN CB.MTB_impliedconsent = 0 THEN 'No' 
WHEN CB.MTB_impliedconsent = 1 THEN 'Yes' END AS 'Implied Consent'
--CB.MTB_emailpref_productupdates, CB.MTB_impliedconsent, CB.Department,  A.Name
--Cp.MTB_PriceListName,CEB.MTB_ID_Search, CB.FirstName, CB.LastName, CB.EMailAddress1, A.Name, A.StateCode--, CP.MTB_name
FROM MTB_MSCRM.dbo.Contact CB 
--INNER JOIN MTB_MSCRM.dbo.ContactExtensionBase CEB ON CB.ContactId = CEB.ContactId
left JOIN Account A ON CB.ParentCustomerId = A.accountid
--INNER JOIN dbo.MTB_customerproduct CP ON A.AccountId = CP.MTB_AccountId
WHERE  
CB.StateCode = 0 --Contact is active
and (MTB_EmailStatus <> 102395 --email not undeliverable
	OR MTB_EmailStatus IS NULL)
AND CB.EMailAddress1 IS NOT NULL -- email address is not null
AND A.StateCode = 0 -- Parent account is active
AND CB.MTB_emailpref_productupdates <> 908600002 -- Product Updates does not equal Unsubscribed
AND CB.Address1_Country NOT IN ('DE', 'KR', 'IN', 'CN', 'JP', 'TH', 'TR', 'TW', 'SG') --Exclude Partner Territories

AND CB.ParentCustomerId IN 

(SELECT accountid FROM #Query1)

SELECT * FROM #MtbSchool 



/**********************************************************************************************************/
/*                 QUERY 3 - Assumed 17 Users who don't fall into the MUL or LC audiences above.          */
/**********************************************************************************************************/

DROP TABLE #MtbPurch, #MtbTraining, #MtbTechSupp

/*Purchased MTB 17*/
CREATE Table #MtbPurch(
MTB_ID int
,EmailAddress nvarchar(200)
,Country nvarchar(4000)
,SiteBaseName nvarchar(160)
--,ContactId UNIQUEIDENTIFIER
)
INSERT INTO #MtbPurch
SELECT DISTINCT
                   CEB.MTB_ID_Search,
                    B.EMailAddress1,
                    CAB.Country,            
                    SB.Name	 FROM      MTB_MSCRM.dbo.ContactBase B WITH ( NOLOCK )
				    INNER JOIN MTB_MSCRM.dbo.ContactExtensionBase CEB ON B.ContactId = CEB.ContactId
				    INNER JOIN MTB_MSCRM.dbo.CustomerAddressBase CAB ON CAB.ParentId = CEB.ContactId
                    INNER JOIN MTB_MSCRM.dbo.MTB_eventattendee EA WITH ( NOLOCK ) ON B.contactid = EA.MTB_contactid AND EA.statecode = 0
                    INNER JOIN MTB_MSCRM.dbo.MTB_customerproductExtensionBase CPEB ON CPEB.MTB_ContactId = CEB.ContactId
                     INNER JOIN MTB_MSCRM.dbo.siteBase SB ON CEB.MTB_SiteId = SB.SiteId
                    WHERE B.statecode = 0 
                    AND B.EMailAddress1 NOT LIKE '%@MTB%'
                    AND  CEB.MTB_EmailStatus <> 102395 
                    AND (CEB.MTB_emailpref_productupdates = 908600001 OR CEB.MTB_impliedconsent = 1)
                    AND CAB.AddressNumber = 1 AND CAB.Country NOT IN ('DE', 'KR', 'IN', 'CN', 'JP', 'TH', 'TR', 'TW', 'SG')
                    AND CPEB.MTB_name LIKE 'MTB 17%'
                    AND B.CustomerTypeCode IN (101970, 101971)

select * from #MtbPurch
where   EmailAddress collate SQL_Latin1_General_CP1_CI_AS 
			not in(select email collate SQL_Latin1_General_CP1_CI_AS from #query1) 
		AND EmailAddress collate SQL_Latin1_General_CP1_CI_AS 
			not in(select email collate SQL_Latin1_General_CP1_CI_AS from #MtbSchool) 
			

CREATE Table #MtbTraining(
MTB_ID int,
EmailAddress nvarchar(200),
Country nvarchar(4000),
SiteBaseName nvarchar(160)
--,ContactId UNIQUEIDENTIFIER
)

/*Has attended a Training -- you'll need to look for the MTB 17 training skus*/
INSERT INTO #MtbTraining
SELECT DISTINCT
                   CEB.MTB_ID_Search,
                    B.EMailAddress1,
                    CAB.Country,          
                    SB.NAME
                     FROM      MTB_MSCRM.dbo.ContactBase B WITH ( NOLOCK )
				    INNER JOIN MTB_MSCRM.dbo.ContactExtensionBase CEB ON B.ContactId = CEB.ContactId
				    INNER JOIN MTB_MSCRM.dbo.CustomerAddressBase CAB ON CAB.ParentId = CEB.ContactId
                    INNER JOIN MTB_MSCRM.dbo.MTB_eventattendee EA WITH ( NOLOCK ) ON B.contactid = EA.MTB_contactid AND EA.statecode = 0
                    INNER JOIN MTB_MSCRM.dbo.MTB_eventcoursetime ET WITH ( NOLOCK ) ON EA.MTB_eventcoursetimeid = ET.MTB_eventcoursetimeid
                                                                             AND ET.statecode = 0
                    INNER JOIN MTB_MSCRM.dbo.Campaign C WITH ( NOLOCK ) ON ( ET.MTB_campaignid = C.campaignid
                                                               AND ( C.typecode = 3 AND C.MTB_eventtype = 908600001))
                    inner JOIN MTB_MSCRM.dbo.Product P ON P.ProductId = ET.MTB_CourseProductId
                    INNER JOIN MTB_MSCRM.dbo.siteBase SB ON CEB.MTB_SiteId = SB.SiteId
          WHERE 
          B.statecode = 0 
          AND B.EMailAddress1 NOT LIKE '%@MTB%'
          AND  CEB.MTB_EmailStatus <> 102395 --email is deliverable
          AND (CEB.MTB_emailpref_productupdates = 908600001 OR CEB.MTB_impliedconsent = 1) /*Product mail flag = Subscribed OR Implied Consent = Yes */
          AND CAB.AddressNumber = 1 AND CAB.Country NOT IN ('DE', 'KR', 'IN', 'CN', 'JP', 'TH', 'TR', 'TW', 'SG') /* Not in ILR territories*/
          AND B.CustomerTypeCode IN (101970, 101971)
          AND MTB_EventType = 908600001 --training
          AND P.ProductNumber LIKE '%17%' -- CAM I joined in the Product table to Event Course time to find all with a 17 sku

select * from #MtbTraining
where   EmailAddress collate SQL_Latin1_General_CP1_CI_AS 
			not in(select email collate SQL_Latin1_General_CP1_CI_AS from #query1) 
		AND EmailAddress collate SQL_Latin1_General_CP1_CI_AS 
			not in(select email collate SQL_Latin1_General_CP1_CI_AS from #MtbSchool) 
			

/*Has called TS for MTB 17 excluding the trial*/
CREATE Table #MtbTechSupp(
--ContactId UNIQUEIDENTIFIER,
MTBid INT
--FirstName NVARCHAR(255)
--,LastName NVARCHAR(255)
,EmailAddress1 NVARCHAR(255)
,Address1_Country NVARCHAR(255)
,MTB_SiteIdName NVARCHAR(255)

)
insert into #MtbTechSupp
select distinct  CEB.MTB_ID_Search
				,B.EMailAddress1
				,CAB.Country
				,SB.NAME	
	
   FROM      MTB_MSCRM.dbo.ContactBase B WITH ( NOLOCK ) INNER JOIN MTB_MSCRM.dbo.ContactExtensionBase CEB ON B.ContactId = CEB.ContactId
					INNER JOIN MTB_MSCRM.dbo.CustomerAddressBase CAB ON CAB.ParentId = CEB.ContactId
                    INNER JOIN MTB_MSCRM.dbo.IncidentBase I WITH ( NOLOCK ) ON I.CustomerId = B.ContactId
                    INNER JOIN MTB_MSCRM.dbo.IncidentExtensionBase IEB  WITH ( NOLOCK ) ON I.IncidentId = IEB.IncidentId
                    INNER JOIN MTB_MSCRM.dbo.siteBase SB ON CEB.MTB_SiteId = SB.SiteId
	WHERE     B.statecode = 0
				AND B.EMailAddress1 NOT LIKE '%@MTB%'	
				AND  CEB.MTB_EmailStatus <> 102395 /*Not Undeliverable Email Address*/
				AND (CEB.MTB_emailpref_productupdates = 908600001 OR CEB.MTB_impliedconsent = 1)
				AND CAB.AddressNumber = 1 AND CAB.Country NOT IN ('DE', 'KR', 'IN', 'CN', 'JP', 'TH', 'TR', 'TW', 'SG')
				AND IEB.MTB_ProductLineId IN ('C0C6D70D-6FE2-E211-B09F-0050569D0FE3') --MTB 17
				AND I.CaseTypeCode IN (101529, 102062) -- case type is Software or Statistics
 				AND IEB.MTB_subTopic <> 103961 -- subtopic NOT trialissues
 				AND B.CustomerTypeCode IN (101970, 101971)

select * from #mtbtechsupp
where   EmailAddress1 collate SQL_Latin1_General_CP1_CI_AS 
			not in(select email collate SQL_Latin1_General_CP1_CI_AS from #query1) 
		AND EmailAddress1 collate SQL_Latin1_General_CP1_CI_AS 
			not in(select email collate SQL_Latin1_General_CP1_CI_AS from #MtbSchool) 
			

/**********************************************************************************************************/
/*                 QUERY 4 - Assumed Users of Previous Releases                                         */   
/**********************************************************************************************************/

--Use the same logic from Query 3 above but it will need to be tweaked per the requirements from Cathy 
DROP TABLE #MtbPurchNot17
		  ,#MtbTechSuppNot17
		  ,#QTSubscriber
		  ,#QTSubCoord
		  ,#MtbOldTraining
		  ,#MtbWebinars
/*Purchased MTB before 17*/
CREATE Table #MtbPurchNot17(
MTB_ID int
,EmailAddress nvarchar(200)
,Country nvarchar(4000)
,SiteBaseName nvarchar(160)
--,ContactId UNIQUEIDENTIFIER
)
insert into #MtbPurchNot17
SELECT DISTINCT
                   CEB.MTB_ID_Search,
                    B.EMailAddress1,
                    CAB.Country,            
                    SB.Name	 FROM      MTB_MSCRM.dbo.ContactBase B WITH ( NOLOCK )
				    INNER JOIN MTB_MSCRM.dbo.ContactExtensionBase CEB ON B.ContactId = CEB.ContactId
				    INNER JOIN MTB_MSCRM.dbo.CustomerAddressBase CAB ON CAB.ParentId = CEB.ContactId
                    INNER JOIN MTB_MSCRM.dbo.MTB_eventattendee EA WITH ( NOLOCK ) ON B.contactid = EA.MTB_contactid AND EA.statecode = 0
                    INNER JOIN MTB_MSCRM.dbo.MTB_customerproductExtensionBase CPEB ON CPEB.MTB_ContactId = CEB.ContactId
                     INNER JOIN MTB_MSCRM.dbo.siteBase SB ON CEB.MTB_SiteId = SB.SiteId
                    WHERE B.statecode = 0 
                    AND B.EMailAddress1 NOT LIKE '%@MTB%'
                    AND  CEB.MTB_EmailStatus <> 102395 
                    AND (CEB.MTB_emailpref_productupdates = 908600001 OR CEB.MTB_impliedconsent = 1)
                    AND CAB.AddressNumber = 1 AND CAB.Country NOT IN ('DE', 'KR', 'IN', 'CN', 'JP', 'TH', 'TR', 'TW', 'SG')
                    AND CPEB.MTB_name LIKE 'MTB%'
					AND B.CustomerTypeCode IN (101970, 101971)
AND CPEB.MTB_name NOT in
(Select DISTINCT MTB_name FROM dbo.MTB_customerproduct
WHERE MTB_name LIKE '%17%')

SELECT * FROM #MtbPurchNot17
where   EmailAddress collate SQL_Latin1_General_CP1_CI_AS 
			not in(select email collate SQL_Latin1_General_CP1_CI_AS from #query1) 
		AND EmailAddress collate SQL_Latin1_General_CP1_CI_AS 
			not in(select email collate SQL_Latin1_General_CP1_CI_AS from #MtbSchool) 
			
-- tech supp not MTB 17 -- Add trial exclusion from above
CREATE Table #MtbTechSuppNot17(
--ContactId UNIQUEIDENTIFIER,
MTBid INT
--FirstName NVARCHAR(255)
--,LastName NVARCHAR(255)
,EmailAddress1 NVARCHAR(255)
,Address1_Country NVARCHAR(255)
,MTB_SiteIdName NVARCHAR(255)
)
insert into #MtbTechSuppNot17
select distinct  CEB.MTB_ID_Search
				,B.EMailAddress1
				,CAB.Country
				,SB.NAME	
	
   FROM      MTB_MSCRM.dbo.ContactBase B WITH ( NOLOCK ) INNER JOIN MTB_MSCRM.dbo.ContactExtensionBase CEB ON B.ContactId = CEB.ContactId
					INNER JOIN MTB_MSCRM.dbo.CustomerAddressBase CAB ON CAB.ParentId = CEB.ContactId
                    INNER JOIN MTB_MSCRM.dbo.IncidentBase I WITH ( NOLOCK ) ON I.CustomerId = B.ContactId
                    INNER JOIN MTB_MSCRM.dbo.IncidentExtensionBase IEB  WITH ( NOLOCK ) ON I.IncidentId = IEB.IncidentId
                    INNER JOIN MTB_MSCRM.dbo.siteBase SB ON CEB.MTB_SiteId = SB.SiteId
	WHERE     B.statecode = 0
				AND B.EMailAddress1 NOT LIKE '%@MTB%'	
				AND  CEB.MTB_EmailStatus <> 102395 /*Not Undeliverable Email Address*/
				AND (CEB.MTB_emailpref_productupdates = 908600001 OR CEB.MTB_impliedconsent = 1)
				AND CAB.AddressNumber = 1 AND CAB.Country NOT IN ('DE', 'KR', 'IN', 'CN', 'JP', 'TH', 'TR', 'TW', 'SG')
				AND IEB.MTB_ProductLineId IN ('9485C44F-C3BB-48C1-8581-0BED3C1B89A8',
													'6D8E1DAE-F622-4180-9F38-0DC5F472D5AB',
													'BE513405-B3E8-4DE0-AB4E-1C382368E90E',
													'C0EE8A58-9834-4A5C-9398-305C116C9ED8',
													'5C168F3E-6B4E-49F6-A0EE-61FECA6DE6E4',
													'680A6063-BEEE-4CC9-B77C-73E67EE43C39',
													'6FC5EEC6-EC5F-4644-BA16-A4C86183B255',
													'A630328D-5497-483E-BA03-BC7FC56D4423',
													'67C593C9-9809-40D6-88AD-CEA97FF9F65C',
													'98ED123A-A041-4155-B4FD-E6BBEDEFAA3B',
													'40B49C09-D364-4337-B90D-F7E7C227F8AF') --MTB NOT 17
				AND I.CaseTypeCode IN (101529, 102062) -- case type is Software or Statistics
 				AND IEB.MTB_subTopic <> 103961 -- subtopic NOT trialissues
 				AND B.CustomerTypeCode IN (101970, 101971)

select * from #MtbTechSuppNot17
where   EmailAddress1 collate SQL_Latin1_General_CP1_CI_AS 
			not in(select email collate SQL_Latin1_General_CP1_CI_AS from #query1) 
		AND EmailAddress1 collate SQL_Latin1_General_CP1_CI_AS 
			not in(select email collate SQL_Latin1_General_CP1_CI_AS from #MtbSchool) 
			

-- active or inactive QT subscription
CREATE Table #QTSubscriber(
--ContactId UNIQUEIDENTIFIER,
MTBid INT
--FirstName NVARCHAR(255)
--,LastName NVARCHAR(255)
,EmailAddress1 NVARCHAR(255)
,Address1_Country NVARCHAR(255)
,MTB_SiteIdName NVARCHAR(255)
)
insert into #QTSubscriber
select distinct  CEB.MTB_ID_Search
				,B.EMailAddress1
				,CAB.Country
				,SB.NAME	
	
   FROM      MTB_MSCRM.dbo.ContactBase B WITH ( NOLOCK ) INNER JOIN MTB_MSCRM.dbo.ContactExtensionBase CEB ON B.ContactId = CEB.ContactId
					INNER JOIN MTB_MSCRM.dbo.CustomerAddressBase CAB ON CAB.ParentId = CEB.ContactId
                    INNER JOIN MTB_MSCRM.dbo.IncidentBase I WITH ( NOLOCK ) ON I.CustomerId = B.ContactId
                    INNER JOIN MTB_MSCRM.dbo.IncidentExtensionBase IEB  WITH ( NOLOCK ) ON I.IncidentId = IEB.IncidentId
                    INNER JOIN MTB_MSCRM.dbo.siteBase SB ON CEB.MTB_SiteId = SB.SiteId
					INNER JOIN QualityTrainer.dbo.AssignedSubscriptions QT ON CEB.MTB_id_search = QT.CustomerId
	WHERE     B.statecode = 0
				AND B.EMailAddress1 NOT LIKE '%@MTB%'	
				AND  CEB.MTB_EmailStatus <> 102395 /*Not Undeliverable Email Address*/
				AND (CEB.MTB_emailpref_productupdates = 908600001 OR CEB.MTB_impliedconsent = 1)
				AND CAB.AddressNumber = 1 AND CAB.Country NOT IN ('DE', 'KR', 'IN', 'CN', 'JP', 'TH', 'TR', 'TW', 'SG')
				AND QT.CustomerId <> QT.CoordinatorId -- subscribers who are not coordinators
				AND B.CustomerTypeCode IN (101970, 101971)
				--AND CPEB.MTB_name LIKE 'Quality Trainer%' 


select * from #QTSubscriber
where   EmailAddress1 collate SQL_Latin1_General_CP1_CI_AS 
			not in(select email collate SQL_Latin1_General_CP1_CI_AS from #query1) 
		AND EmailAddress1 collate SQL_Latin1_General_CP1_CI_AS 
			not in(select email collate SQL_Latin1_General_CP1_CI_AS from #MtbSchool) 
			
-- active or inactive qt subscription coordinator
CREATE Table #QTSubCoord(
--ContactId UNIQUEIDENTIFIER,
MTBid INT
--FirstName NVARCHAR(255)
--,LastName NVARCHAR(255)
,EmailAddress1 NVARCHAR(255)
,Address1_Country NVARCHAR(255)
,MTB_SiteIdName NVARCHAR(255)
)
insert into #QTSubCoord
select distinct  CEB.MTB_ID_Search
				,B.EMailAddress1
				,CAB.Country
				,SB.NAME	
	
   FROM      MTB_MSCRM.dbo.ContactBase B WITH ( NOLOCK ) INNER JOIN MTB_MSCRM.dbo.ContactExtensionBase CEB ON B.ContactId = CEB.ContactId
					INNER JOIN MTB_MSCRM.dbo.CustomerAddressBase CAB ON CAB.ParentId = CEB.ContactId
                    INNER JOIN MTB_MSCRM.dbo.IncidentBase I WITH ( NOLOCK ) ON I.CustomerId = B.ContactId
                    INNER JOIN MTB_MSCRM.dbo.IncidentExtensionBase IEB  WITH ( NOLOCK ) ON I.IncidentId = IEB.IncidentId
                    INNER JOIN MTB_MSCRM.dbo.siteBase SB ON CEB.MTB_SiteId = SB.SiteId
					INNER JOIN QualityTrainer.dbo.AssignedSubscriptions QT ON CEB.MTB_id_search = QT.CoordinatorId
	WHERE     B.statecode = 0
				AND B.EMailAddress1 NOT LIKE '%@MTB%'	
				AND  CEB.MTB_EmailStatus <> 102395 /*Not Undeliverable Email Address*/
				AND (CEB.MTB_emailpref_productupdates = 908600001 OR CEB.MTB_impliedconsent = 1)
				AND CAB.AddressNumber = 1 AND CAB.Country NOT IN ('DE', 'KR', 'IN', 'CN', 'JP', 'TH', 'TR', 'TW', 'SG')
                AND B.CustomerTypeCode IN (101970, 101971)
select * from #QTSubCoord

where   EmailAddress1 collate SQL_Latin1_General_CP1_CI_AS 
			not in(select email collate SQL_Latin1_General_CP1_CI_AS from #query1) 
		AND EmailAddress1 collate SQL_Latin1_General_CP1_CI_AS 
			not in(select email collate SQL_Latin1_General_CP1_CI_AS from #MtbSchool) 
			

-- attend training < Feb 1 2014
CREATE Table #MtbOldTraining(
MTB_ID int,
EmailAddress nvarchar(200),
Country nvarchar(4000),
SiteBaseName nvarchar(160)
--,ContactId UNIQUEIDENTIFIER
)

INSERT INTO #MtbOldTraining
SELECT DISTINCT
                   CEB.MTB_ID_Search,
                    B.EMailAddress1,
                    CAB.Country,          
                    SB.NAME
                     FROM      MTB_MSCRM.dbo.ContactBase B WITH ( NOLOCK )
				    INNER JOIN MTB_MSCRM.dbo.ContactExtensionBase CEB ON B.ContactId = CEB.ContactId
				    INNER JOIN MTB_MSCRM.dbo.CustomerAddressBase CAB ON CAB.ParentId = CEB.ContactId
                    INNER JOIN MTB_MSCRM.dbo.MTB_eventattendee EA WITH ( NOLOCK ) ON B.contactid = EA.MTB_contactid AND EA.statecode = 0
                    INNER JOIN MTB_MSCRM.dbo.MTB_eventcoursetime ET WITH ( NOLOCK ) ON EA.MTB_eventcoursetimeid = ET.MTB_eventcoursetimeid
                                                                             AND ET.statecode = 0
                    INNER JOIN MTB_MSCRM.dbo.Campaign C WITH ( NOLOCK ) ON ( ET.MTB_campaignid = C.campaignid
                                                               AND ( C.typecode = 3 AND C.MTB_eventtype = 908600001))
                    INNER JOIN MTB_MSCRM.dbo.siteBase SB ON CEB.MTB_SiteId = SB.SiteId
          WHERE 
          B.statecode = 0 
          AND B.EMailAddress1 NOT LIKE '%@MTB%'
          AND  CEB.MTB_EmailStatus <> 102395 
          AND (CEB.MTB_emailpref_productupdates = 908600001 OR CEB.MTB_impliedconsent = 1) /*Product mail flag = Subscribed OR Implied Consent = Yes */
          AND CAB.AddressNumber = 1 AND CAB.Country NOT IN ('DE', 'KR', 'IN', 'CN', 'JP', 'TH', 'TR', 'TW', 'SG') /* Not in ILR territories*/
          AND B.CustomerTypeCode IN (101970, 101971)
          AND MTB_EventType = 908600001 AND MTB_product = 908600000 
          AND ET.MTB_startDate < '2/1/2014'


select * from #MtbOldTraining
where   EmailAddress collate SQL_Latin1_General_CP1_CI_AS 
			not in(select email collate SQL_Latin1_General_CP1_CI_AS from #query1) 
		AND EmailAddress collate SQL_Latin1_General_CP1_CI_AS 
			not in(select email collate SQL_Latin1_General_CP1_CI_AS from #MtbSchool) 
			
--attend any webinars since 1/1/2013
CREATE Table #MtbWebinars(
MTB_ID int,
EmailAddress nvarchar(200),
Country nvarchar(4000),
SiteBaseName nvarchar(160)
--,ContactId UNIQUEIDENTIFIER
)

INSERT INTO #MtbWebinars
SELECT DISTINCT
                   CEB.MTB_ID_Search,
                    B.EMailAddress1,
                    CAB.Country,          
                    SB.NAME
                     FROM      MTB_MSCRM.dbo.ContactBase B WITH ( NOLOCK )
				    INNER JOIN MTB_MSCRM.dbo.ContactExtensionBase CEB ON B.ContactId = CEB.ContactId
				    INNER JOIN MTB_MSCRM.dbo.CustomerAddressBase CAB ON CAB.ParentId = CEB.ContactId
                    INNER JOIN MTB_MSCRM.dbo.MTB_eventattendee EA WITH ( NOLOCK ) ON B.contactid = EA.MTB_contactid AND EA.statecode = 0
                    INNER JOIN MTB_MSCRM.dbo.MTB_eventcoursetime ET WITH ( NOLOCK ) ON EA.MTB_eventcoursetimeid = ET.MTB_eventcoursetimeid
                                                                             AND ET.statecode = 0
                    INNER JOIN MTB_MSCRM.dbo.Campaign C WITH ( NOLOCK ) ON ( ET.MTB_campaignid = C.campaignid
                                                               AND ( C.typecode = 3 AND C.MTB_eventtype = 908600002))
                    INNER JOIN MTB_MSCRM.dbo.siteBase SB ON CEB.MTB_SiteId = SB.SiteId
          WHERE 
          B.statecode = 0 
          AND B.EMailAddress1 NOT LIKE '%@MTB%'
          AND  CEB.MTB_EmailStatus <> 102395 
          AND (CEB.MTB_emailpref_productupdates = 908600001 OR CEB.MTB_impliedconsent = 1) /*Product mail flag = Subscribed OR Implied Consent = Yes */
          AND CAB.AddressNumber = 1 AND CAB.Country NOT IN ('DE', 'KR', 'IN', 'CN', 'JP', 'TH', 'TR', 'TW', 'SG') /* Not in ILR territories*/
          AND B.CustomerTypeCode = 101971
          AND MTB_EventType = 908600002 AND MTB_product = 908600000 
          AND ET.MTB_startDate > '1/1/2013'


select * from #MtbWebinars 
where   EmailAddress collate SQL_Latin1_General_CP1_CI_AS 
			not in(select email collate SQL_Latin1_General_CP1_CI_AS from #query1) 
		AND EmailAddress collate SQL_Latin1_General_CP1_CI_AS 
			not in(select email collate SQL_Latin1_General_CP1_CI_AS from #MtbSchool) 
			
/*******************************************************************
* Query 5: Anyone who has Product Updates = Subscribed (including  ILR territories)
* 
* Exclude: 
*  -- Anyone who receives ANY of the emails above
*  -- Anyone who has a mail flag of MTB News = Subscribed
*************************************************************************/
DROP TABLE #MtbAnyoneElse

CREATE Table #MtbAnyoneElse(
			MTB_ID int,
			EmailAddress nvarchar(200),
			Country nvarchar(4000),
			SiteBaseName nvarchar(160)
			--,ContactId UNIQUEIDENTIFIER
			)

INSERT INTO #MtbAnyoneElse
SELECT DISTINCT
                   CEB.MTB_ID_Search,
                    B.EMailAddress1,
                    CAB.Country,          
                    SB.NAME
                     FROM      MTB_MSCRM.dbo.ContactBase B WITH ( NOLOCK )
				    INNER JOIN MTB_MSCRM.dbo.ContactExtensionBase CEB ON B.ContactId = CEB.ContactId
				    INNER JOIN MTB_MSCRM.dbo.CustomerAddressBase CAB ON CAB.ParentId = CEB.ContactId
				    -- CAM I don't think we need to join to the Campaign and event tables?
                    --INNER JOIN MTB_MSCRM.dbo.MTB_eventattendee EA WITH ( NOLOCK ) ON B.contactid = EA.MTB_contactid AND EA.statecode = 0
                    --INNER JOIN MTB_MSCRM.dbo.MTB_eventcoursetime ET WITH ( NOLOCK ) ON EA.MTB_eventcoursetimeid = ET.MTB_eventcoursetimeid
                    --                                                         AND ET.statecode = 0
                    --INNER JOIN MTB_MSCRM.dbo.Campaign C WITH ( NOLOCK ) ON ( ET.MTB_campaignid = C.campaignid
                    --                                           AND ( C.typecode = 3 AND C.MTB_eventtype = 908600002))
                    INNER JOIN MTB_MSCRM.dbo.siteBase SB ON CEB.MTB_SiteId = SB.SiteId
          WHERE 
          B.statecode = 0 
          AND B.EMailAddress1 NOT LIKE '%@MTB%'
          
          AND  CEB.MTB_EmailStatus <> 102395 
          AND (CEB.MTB_emailpref_productupdates = 908600001)  /*Product mail flag = Subscribed OR Implied Consent = Yes */
          AND (CEB.MTB_emailpref_MTBnews <> 908600001)  /*Product mail flag = Subscribed OR Implied Consent = Yes */
          AND   EmailAddress collate SQL_Latin1_General_CP1_CI_AS 
				not in(select email collate SQL_Latin1_General_CP1_CI_AS from #query1) 
			AND EmailAddress collate SQL_Latin1_General_CP1_CI_AS 
				not in(select email collate SQL_Latin1_General_CP1_CI_AS from #MtbSchool) 
			AND EmailAddress collate SQL_Latin1_General_CP1_CI_AS 
				not in(select EmailAddress collate SQL_Latin1_General_CP1_CI_AS from #MtbPurch) 
			AND EmailAddress collate SQL_Latin1_General_CP1_CI_AS 
				not in(select EmailAddress collate SQL_Latin1_General_CP1_CI_AS from #MtbTraining) 
			AND EmailAddress collate SQL_Latin1_General_CP1_CI_AS 
				not in(select EmailAddress collate SQL_Latin1_General_CP1_CI_AS from #MtbTechSupp) 
			AND EmailAddress collate SQL_Latin1_General_CP1_CI_AS 
				not in(select EmailAddress collate SQL_Latin1_General_CP1_CI_AS from #MtbPurchNot17) 
			AND EmailAddress collate SQL_Latin1_General_CP1_CI_AS 
				not in(select EmailAddress collate SQL_Latin1_General_CP1_CI_AS from #MtbTechSuppNot17) 
			AND EmailAddress collate SQL_Latin1_General_CP1_CI_AS 
				not in(select EmailAddress collate SQL_Latin1_General_CP1_CI_AS from #QTSubscriber) 
			AND EmailAddress collate SQL_Latin1_General_CP1_CI_AS 
				not in(select EmailAddress collate SQL_Latin1_General_CP1_CI_AS from #QTSubCoord) 
			AND EmailAddress collate SQL_Latin1_General_CP1_CI_AS 
				not in(select EmailAddress collate SQL_Latin1_General_CP1_CI_AS from #MtbOldTraining) 
			AND EmailAddress collate SQL_Latin1_General_CP1_CI_AS 
				not in (select EmailAddress collate SQL_Latin1_General_CP1_CI_AS from #MtbWebinars) 

select * from #MtbAnyoneElse 
			
