/*Insert into Staging table */

DELETE FROM [IntegrationStaging].[dbo].[ContactsWithInvalidPendingAcctNames_Staging]
INSERT INTO IntegrationStaging.dbo.ContactsWithInvalidPendingAcctNames_Staging

SELECT   CEB.ContactId, CEB.MTB_ID_Search, 
ParentCustomerId, MTB_PendingAccountName, MTB_ProfileComplete
FROM  MTB_MSCRM.dbo.ContactExtensionBase CEB 
INNER JOIN MTB_MSCRM.dbo.ContactBase C ON C.ContactId = CEB.ContactId
WHERE CEB.MTB_PendingAccountName in ('Home','Personal','A','B','C','D' ,'E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V'
,'W','X','Y','Z','.','..','-','--','---','----','-----','n/a','na','none','AA','BB','CC','DD' ,'EE','FF','GG','HH','II','JJ','KK','LL','MM','NN'
,'OO','PP','QQ','RR','SS','TT','UU','VV','WW','XX','YY','ZZ', 'student', ':','school','no','0','1','2','3','4','5','6','7','8','9','00','11','22'
,'33','44','55','66','77','88','99','111','1111','11111','yes','aaa','self','self employed','abc','asdf','asd','aaaa','aaaaa','aaaaaa','aaaaaaa','aaa'
,'sss','ddd','az','?','??','???','123','/','.','..','...','....','.....','*','**','***',',',',,','000','000000','123456','bbb','ccc','n.a.'
)

/**************************************************************************************************************************************************/
/*Query Records processed in contact and staging table*/

USE MTB_MSCRM
Select MTB_ProfileComplete, MTB_PendingAccountName, ParentCustomerId, MTB_ID_Search 
from dbo.ContactBase
INNER JOIN dbo.ContactExtensionBase
ON dbo.ContactBase.ContactId = dbo.ContactExtensionBase.ContactId
WHERE ContactBase.ContactId IN 

(
SELECT ContactId
  FROM [IntegrationStaging].[dbo].[ContactsWithInvalidPendingAcctNames_Staging]

)

AND MTB_PendingAccountName IS NULL

/*********************************************************************************************************************************/
/*Select All from Staging Table*/

SELECT [ContactId]
      ,[ParentCustomerId]
      ,[MTB_PendingAccountName]
      ,[MTB_ProfileComplete]
  FROM [IntegrationStaging].[dbo].[ContactsWithInvalidPendingAcctNames_Staging]

