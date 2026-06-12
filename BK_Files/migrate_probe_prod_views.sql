/* migrate_probe_prod_views.sql — READ-ONLY probe against PROD (BrokerKnow).
   For each view that FAILED to build in BrokerKnow_Clean, try to resolve it in prod
   (SELECT TOP 0 ... FROM it). If it errors in prod too, the build "failure" is
   pre-existing prod rot reproduced faithfully — NOT a migration defect.
   Only SELECTs; no DDL/DML touches prod. */
SET NOCOUNT ON;
USE BrokerKnow;
GO
DECLARE @v TABLE (nm sysname);
INSERT @v VALUES
 ('ClientStatement'),('ClientStatementBalances'),('FullClientList2'),('FullClientListWithBalances'),
 ('BondClientList'),('EditBondProposals'),('OnlineClients'),('ClientsAboveCreditLimit'),
 ('ClientsAboveCreditLimit2'),('BalanceAccountList'),('BalanceAgentList'),
 ('BankAccList'),('BnkBranchList'),('ClientTestStatement'),('CompanyInfoList'),
 ('ContractDisplayList'),('ContractLevyList'),('CPaymentList'),('CReceiptsList'),
 ('GroupMembersList'),('OwnerCommissions'),('SubmitOfferingList'),('UnSubmitOfferingList'),
 ('SystemEntityList'),('AACommissionReport'),('AACommisionReportDaily');
DECLARE @nm sysname, @sql nvarchar(max);
DECLARE c CURSOR LOCAL FAST_FORWARD FOR SELECT nm FROM @v ORDER BY nm;
OPEN c; FETCH NEXT FROM c INTO @nm;
WHILE @@FETCH_STATUS = 0
BEGIN
    SET @sql = 'SELECT TOP 0 1 AS probe FROM dbo.' + QUOTENAME(@nm);
    BEGIN TRY
        EXEC sys.sp_executesql @sql;
        PRINT 'OK_IN_PROD       : ' + @nm;
    END TRY
    BEGIN CATCH
        PRINT 'BROKEN_IN_PROD   : ' + @nm + '  ->  ' + ERROR_MESSAGE();
    END CATCH
    FETCH NEXT FROM c INTO @nm;
END
CLOSE c; DEALLOCATE c;
GO
