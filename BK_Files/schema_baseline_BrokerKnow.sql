/* ============================================================
   BrokerKnow CLEAN SCHEMA BASELINE (dbo; [trash] excluded)
   Generated read-only from BrokerKnow_Test catalog.
   Canonical target for the phased data migration.
   ============================================================ */
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF SCHEMA_ID('app') IS NULL EXEC('CREATE SCHEMA [app]');
GO
CREATE TABLE [dbo].[_CDS_Imported_Files_] (
    [CDSFile_DPA_] int IDENTITY(1,1) NOT NULL,
    [CDSFile_EIT_] nvarchar(40) NOT NULL CONSTRAINT [DF__CDS_Imported_Files__CDSFile_EIT_] DEFAULT (newid()),
    [CDSFileName] nvarchar(50) NOT NULL,
    [CDSFileImported] bit NOT NULL CONSTRAINT [DF__CDS_Imported_Files__CDSFileImported] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[_CDS_Imported_Holdings_] (
    [ImportedID] int IDENTITY(1,1) NOT NULL,
    [TradeDate] datetime NULL,
    [TradeTime] varchar(12) NULL,
    [CDSNo] varchar(50) NULL,
    [SecurityImportCode] varchar(50) NULL,
    [Quantity] varchar(50) NULL,
    [AccountStatus] varchar(255) NULL,
    [BalanceFree] tinyint NULL,
    [Imported] tinyint NULL CONSTRAINT [DF__CDS_Imported_Holdings__Imported] DEFAULT (0),
    [isRight] bit NULL CONSTRAINT [DF__CDS_Imported_Holdings__isRight] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[_CDS_Imported_Trades_] (
    [TradeDate] char(8) NULL,
    [TradeTime] char(19) NULL,
    [CDSRef] char(10) NULL,
    [BuySell] char(1) NULL,
    [ParticipantID] char(4) NULL,
    [ParticipantType] char(1) NULL,
    [ClientPrefix] char(13) NULL,
    [ClientSuffix] char(2) NULL,
    [ClientAccountNo] char(2) NULL,
    [Category] char(1) NULL,
    [IssuerCode] char(4) NULL,
    [DebtType] char(2) NULL,
    [SecurityDescription] char(30) NULL,
    [MainType] char(1) NULL,
    [SubType] char(4) NULL,
    [Quantity] char(19) NULL,
    [Price] char(10) NULL,
    [BrokerageFee] char(17) NULL,
    [CDSFee] char(17) NULL,
    [Interest] char(13) NULL,
    [Status] char(1) NULL,
    [ContraBrokerID] char(4) NULL,
    [ContraBrokerType] char(1) NULL,
    [Crossing] char(1) NULL,
    [SettlementAmount] char(32) NULL,
    [SettlementDate] char(8) NULL,
    [TradeType] char(1) NULL,
    [Reserved] char(573) NULL,
    [RecordType] char(1) NULL,
    [CDSImport_EIT_] uniqueidentifier NOT NULL CONSTRAINT [DF__CDS_Imported_Trades__CDSImport_DPA_] DEFAULT (newid()),
    [Processed] bit NOT NULL CONSTRAINT [DF__CDS_Imported_Trades__Processed] DEFAULT (0),
    [CDSImport_DPA_] int IDENTITY(1,1) NOT NULL,
    [CommissionRate] int NULL
);
GO
CREATE TABLE [dbo].[_Initial_Table_ID_] (
    [InitialTableID_DPA_] int IDENTITY(1,1) NOT NULL,
    [InitialTableID_EIT_] nvarchar(40) NULL CONSTRAINT [DF__Initial_Table_ID__InitialTableID_EIT_] DEFAULT (newid()),
    [TableName] nvarchar(4000) NOT NULL,
    [InitialID] int NOT NULL
);
GO
CREATE TABLE [dbo].[_Parent_Child_Links_] (
    [Child] nvarchar(200) NOT NULL,
    [ChildType] int NOT NULL,
    [DeletionMessage] nvarchar(200) NOT NULL,
    [Parent] nvarchar(200) NOT NULL,
    [ParentKey] nvarchar(200) NOT NULL,
    [Link_DPA_] int IDENTITY(1,1) NOT NULL
);
GO
CREATE TABLE [dbo].[_Record_Locks_] (
    [Lock_DPA_] int NOT NULL,
    [LockID] nvarchar(40) NOT NULL,
    [LockTime] int NOT NULL,
    [RecordID] nvarchar(40) NOT NULL,
    [TableName] nvarchar(200) NOT NULL,
    [UserID] nvarchar(40) NOT NULL
);
GO
CREATE TABLE [dbo].[_ReportsParameters_] (
    [ReportsParam_DPA_] int IDENTITY(1,1) NOT NULL,
    [ReportsParam_EIT_] nvarchar(40) NULL,
    [ReportsMenuID] int NOT NULL,
    [ReportsParamCode] nvarchar(500) NOT NULL,
    [ReportsParamName] nvarchar(100) NOT NULL,
    [ReportsParamType] int NOT NULL
);
GO
CREATE TABLE [dbo].[Account] (
    [Account_DPA_] int NOT NULL,
    [AccountTypeLevel1] int NOT NULL,
    [AccountTypeLevel2] int NULL,
    [AccountTypeLevel3] int NULL,
    [AccountName] nvarchar(100) NOT NULL,
    [AccountCode] nvarchar(50) NOT NULL,
    [Account_EIT_] nvarchar(40) NULL,
    [EntityType_DPA_] int NOT NULL CONSTRAINT [DF_Account_EntityType_DPA_] DEFAULT (5),
    [AccountOpeningBal] money NOT NULL CONSTRAINT [DF_Account_AccountOpeningBal] DEFAULT (0),
    [AccountOpeningDate] datetime NOT NULL CONSTRAINT [DF_Account_AccountOpeningDate] DEFAULT (getdate()),
    [SystemMaintained] bit NOT NULL CONSTRAINT [DF_Account_SystemMaintained] DEFAULT (0),
    [ReconStartDate] datetime NOT NULL CONSTRAINT [DF_Account_ReconStartDate] DEFAULT (getdate()),
    [ReconOpeningBal] float NOT NULL CONSTRAINT [DF_Account_ReconOpeningBal] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[AccountImport] (
    [ID] nvarchar(255) NULL,
    [NAME] nvarchar(255) NULL,
    [Balance] float NULL
);
GO
CREATE TABLE [dbo].[AccountType] (
    [AccountType_DPA_] int IDENTITY(1,1) NOT NULL,
    [AccountTypeName] nvarchar(100) NOT NULL,
    [AccountType_EIT_] nvarchar(40) NULL,
    [DefaultSelection] bit NOT NULL CONSTRAINT [DF_AccountType_DefaultSelection] DEFAULT (0),
    [AccountTypeParent] int NULL,
    [Quarter1] money NULL,
    [Quarter2] money NULL,
    [Quarter3] money NULL,
    [Quarter4] money NULL
);
GO
CREATE TABLE [dbo].[Activity] (
    [Activity_DPA_] int NOT NULL,
    [Activity_EIT_] nvarchar(40) NULL,
    [ActivityDate] smalldatetime NOT NULL,
    [ActivityNotes] nvarchar(255) NULL,
    [ActvtyClass_DPA_] int NOT NULL,
    [Client_DPA_] int NOT NULL,
    [ChangedBy] int NULL,
    [TimeChanged] datetime NULL,
    [Deleted] int NULL CONSTRAINT [DF_Activity_Deleted] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[ActvtyClass] (
    [ActvtyClass_DPA_] int NOT NULL,
    [ActvtyClass_EIT_] nvarchar(40) NULL,
    [ActvtyClassDescription] nvarchar(100) NOT NULL,
    [ClientAccess] bit NULL CONSTRAINT [DF_ActvtyClass_ClientAccess] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[Agent] (
    [Agent_DPA_] int NOT NULL,
    [Agent_EIT_] nvarchar(40) NULL,
    [AgentAddr] ntext NULL,
    [AgentBDate] smalldatetime NULL,
    [AgentCellTel] nvarchar(100) NULL,
    [AgentContact] nvarchar(100) NULL,
    [AgentEmail] nvarchar(100) NULL,
    [AgentFax] nvarchar(100) NULL,
    [AgentHomeTel] nvarchar(100) NULL,
    [AgentIDPass] nvarchar(100) NULL,
    [AgentName] nvarchar(100) NOT NULL,
    [AgentOfficeTel] nvarchar(100) NULL,
    [AgentPhoto] nvarchar(100) NULL,
    [AgentSignature] nvarchar(100) NULL,
    [Branch_DPA_] int NOT NULL,
    [Commission_DPA_] int NOT NULL,
    [Gender_DPA_] int NULL,
    [Residency_DPA_] int NOT NULL,
    [DefaultSelection] bit NOT NULL CONSTRAINT [DF_Agent_DefaultSelection] DEFAULT (0),
    [EntityType_DPA_] int NOT NULL CONSTRAINT [DF_Agent_EntityType_DPA_] DEFAULT (2),
    [AgentOpeningBal] money NOT NULL CONSTRAINT [DF_Agent_BrokerOpeningBal] DEFAULT (0),
    [AgentRegDate] datetime NOT NULL CONSTRAINT [DF_Agent_BrokerRegDate] DEFAULT (getdate()),
    [GenericSetting_DPA_] int NULL,
    [GenericSetting_DPA_2] int NULL,
    [GenericSetting_DPA_3] int NULL,
    [ChangedBy] int NOT NULL CONSTRAINT [DF_Agent_ChangedBy] DEFAULT (0),
    [TimeChanged] smalldatetime NOT NULL CONSTRAINT [DF_Agent_TimeChanged] DEFAULT (getdate()),
    [Deleted] bit NOT NULL CONSTRAINT [DF_Agent_Deleted] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[APortfolio] (
    [Agent_DPA_] int NOT NULL,
    [APortfolio_DPA_] int NOT NULL,
    [APortfolio_EIT_] nvarchar(40) NULL,
    [APortfolioPDate] smalldatetime NOT NULL,
    [APortfolioPrice] money NOT NULL,
    [APortfolioQty] int NOT NULL,
    [Security_DPA_] int NOT NULL
);
GO
CREATE TABLE [dbo].[AuditTrail] (
    [AuditTrail_DPA_] int IDENTITY(1,1) NOT NULL,
    [UserID] int NOT NULL,
    [AuditTrailAction] nvarchar(100) NOT NULL,
    [AuditTrailMoment] datetime NOT NULL CONSTRAINT [DF_AuditTrail_AuditTrailMoment] DEFAULT (getdate()),
    [AuditTrailTable] nvarchar(100) NOT NULL,
    [AuditTrailSQL] nvarchar(4000) NOT NULL,
    [AuditTrailCriteria] nvarchar(4000) NULL
);
GO
CREATE TABLE [dbo].[AuditTrailItem] (
    [AuditTrailItem_DPA_] int IDENTITY(1,1) NOT NULL,
    [AuditTrail_DPA_] int NOT NULL,
    [AuditTrailItemField] nvarchar(150) NOT NULL,
    [AuditTrailItemValue] nvarchar(4000) NOT NULL
);
GO
CREATE TABLE [dbo].[Bank] (
    [Bank_DPA_] int NOT NULL,
    [Bank_EIT_] nvarchar(40) NULL,
    [BankName] nvarchar(100) NOT NULL,
    [BankCode] nvarchar(100) NULL
);
GO
CREATE TABLE [dbo].[BankAcc] (
    [BankAcc_DPA_] int IDENTITY(1,1) NOT NULL,
    [BankAcc_EIT_] nvarchar(40) NULL,
    [BankAccNumber] nvarchar(100) NULL,
    [BnkBranch] nvarchar(100) NULL,
    [Client_DPA_] int NOT NULL,
    [BankAccName] nvarchar(100) NULL,
    [BankNo] int NULL,
    [Bank_DPA_] int NULL
);
GO
CREATE TABLE [dbo].[BnkBranch] (
    [Bank_DPA_] int NOT NULL,
    [BnkBranch_DPA_] int NOT NULL,
    [BnkBranch_EIT_] nvarchar(40) NULL,
    [BnkBranchName] nvarchar(100) NOT NULL,
    [BnkBranchCode] nvarchar(20) NULL,
    [BnkBranchSwiftCode] nvarchar(20) NULL
);
GO
CREATE TABLE [dbo].[Bond] (
    [Bond_DPA_] int NOT NULL,
    [Bond_EIT_] nvarchar(40) NULL,
    [BondIDate] smalldatetime NOT NULL,
    [BondIssue] nvarchar(20) NOT NULL,
    [BondLife] int NOT NULL,
    [BondPayment] nvarchar(20) NOT NULL,
    [BondRate] real NOT NULL,
    [Security_DPA_] int NOT NULL,
    [BondMDate] smalldatetime NOT NULL,
    [Determination] nvarchar(50) NULL,
    [FaceValue] money NOT NULL,
    [ForwardRate] real NULL,
    [DateModified] datetime NULL,
    [ModifiedBy] int NULL
);
GO
CREATE TABLE [dbo].[BondConfirmation] (
    [BondConfirmation_DPA_] int NOT NULL,
    [SettlementDate] smalldatetime NULL,
    [ForwardRate] real NULL,
    [NoDaysMaturity] int NULL,
    [Convexity] float NULL,
    [Duration] int NULL,
    [ConsiderationI] nvarchar(50) NULL,
    [CommissionI] nvarchar(50) NULL,
    [CounterPartyI] nvarchar(50) NULL,
    [CounterParty_DPA_] int NULL,
    [ContactPerson] nvarchar(50) NULL,
    [Title] char(10) NULL,
    [Order_DPA_] int NULL,
    [Deleted] bit NOT NULL
);
GO
CREATE TABLE [dbo].[BondProposals] (
    [Proposal_DPA_] int NOT NULL,
    [Bond_EIT_] nvarchar(40) NULL,
    [BondIDate] smalldatetime NOT NULL,
    [BondIssue] nvarchar(20) NOT NULL,
    [BondPayment] nvarchar(20) NOT NULL,
    [CouponRate] real NOT NULL,
    [Bond_DPA_] int NULL,
    [Security_DPA_] int NOT NULL,
    [BondMDate] smalldatetime NOT NULL,
    [FaceValue] float NOT NULL,
    [Determination] nvarchar(50) NULL,
    [Client_DPA_] int NOT NULL,
    [Accepted] bit NOT NULL CONSTRAINT [DF_BondProposals_Accepted] DEFAULT (0),
    [ProposalDate] smalldatetime NULL,
    [BondCleanPrice] float NULL,
    [AccruedInterest] float NULL,
    [BondDirtyPrice] float NULL,
    [Commission] float NULL,
    [CommissionRate] float NULL,
    [Duration] float NULL,
    [Consideration] float NULL,
    [Convexity] float NULL,
    [TradeType] varchar(10) NULL,
    [SettlementDate] smalldatetime NULL,
    [ForwardRate] float NULL,
    [CouponPeriodDays] int NULL,
    [Basis] real NULL,
    [ChangeInYield] real NULL,
    [LastCouponDate] smalldatetime NULL,
    [NextCouponDate] smalldatetime NULL,
    [BondLife] real NULL,
    [BondMaturiy] real NULL,
    [PeriodicInterestRate] real NULL,
    [S/ACouponAmount] float NULL,
    [InterestDaysAccrued] int NULL,
    [RemainingDaysToCoupon] int NULL,
    [PreviousCouponPayments] int NULL,
    [YieldChange] real NULL,
    [Validity] smalldatetime NULL,
    [AccountManager] int NULL,
    [CounterParty] int NULL,
    [Salutation] varchar(20) NULL,
    [AlternatePrice] float NULL,
    [NetAmount] float NULL,
    [Canceled] bit NOT NULL CONSTRAINT [DF_BondProposals_Canceled] DEFAULT (0),
    [ModifiedBy] int NULL,
    [DateModified] datetime NULL
);
GO
CREATE TABLE [dbo].[BondTrades] (
    [BondTrades_DPA_] int NOT NULL,
    [Client_DPA_] int NOT NULL,
    [Bond_DPA_] int NOT NULL,
    [OrderType_DPA_] int NOT NULL,
    [TradeDate] smalldatetime NULL,
    [Price] money NULL,
    [Quantity] money NULL,
    [Reference] varchar(50) NULL,
    [Included] int NULL CONSTRAINT [DF_BondTrades_Included] DEFAULT (1),
    [DaysInCoupon] numeric(10,0) NULL,
    [Basis] numeric(10,0) NULL,
    [ForwardRate] float NULL,
    [ModifiedBy] int NULL,
    [DateModified] datetime NULL,
    [Deleted] int NULL
);
GO
CREATE TABLE [dbo].[Branch] (
    [Branch_DPA_] int NOT NULL,
    [Branch_EIT_] nvarchar(40) NULL,
    [BranchDescription] nvarchar(100) NOT NULL,
    [DefaultSelection] bit NOT NULL CONSTRAINT [DF_Branch_DefaultSelection] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[Broker] (
    [Broker_DPA_] int NOT NULL,
    [Broker_EIT_] nvarchar(40) NULL,
    [BrokerName] nvarchar(100) NOT NULL,
    [BrokerCode] nvarchar(5) NOT NULL,
    [BrokerAddr] nvarchar(1000) NULL,
    [BrokerOfficeTel] nvarchar(50) NULL,
    [BrokerFax] nvarchar(20) NULL,
    [EntityType_DPA_] int NOT NULL CONSTRAINT [DF_Broker_EntityType_DPA_] DEFAULT (3),
    [BrokerOpeningBal] money NOT NULL CONSTRAINT [DF_Broker_ClientOpeningBal] DEFAULT (0),
    [BrokerRegDate] datetime NULL CONSTRAINT [DF_Broker_ClientRegDate] DEFAULT (getdate())
);
GO
CREATE TABLE [dbo].[BrokerReceiptVoucher] (
    [BrokerReceiptVoucher_DPA_] int NOT NULL,
    [VoucherDate] smalldatetime NULL,
    [BrokerReceiptVoucher_EIT_] nvarchar(40) NULL,
    [VoucherPaid] bit NOT NULL CONSTRAINT [DF_BrokerReceiptVoucher_VoucherPaid] DEFAULT (0),
    [Deleted] bit NOT NULL CONSTRAINT [DF_BrokerReceiptVoucher_Deleted] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[CdsImportedHoldings] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [TradeDate] datetime2(7) NOT NULL,
    [TradeTime] nvarchar(16) NULL,
    [Exchange] nvarchar(16) NOT NULL,
    [ParticipantCode] nvarchar(32) NULL,
    [CdsNo] nvarchar(64) NULL,
    [Symbol] nvarchar(32) NOT NULL,
    [Quantity] bigint NOT NULL,
    [AccountStatus] nvarchar(64) NULL,
    [BalanceFree] nvarchar(8) NULL,
    [SourceFile] nvarchar(260) NULL,
    [SourceFileHash] nvarchar(64) NULL,
    [BatchId] int NOT NULL,
    [MatchStatus] nvarchar(20) NOT NULL CONSTRAINT [DF_CdsImportedHoldings_MatchStatus] DEFAULT ('Unmatched'),
    [UnmatchReason] nvarchar(32) NULL,
    [MatchedClientDpa] int NULL,
    [MatchedSecurityDpa] int NULL,
    [MatchedHoldingId] int NULL,
    [MatchNotes] nvarchar(500) NULL,
    [MatchedAt] datetime2(7) NULL,
    [CreatedAt] datetime2(7) NOT NULL CONSTRAINT [DF_CdsImportedHoldings_CreatedAt] DEFAULT (sysutcdatetime()),
    [CreatedBy] int NULL,
    [Deleted] bit NOT NULL CONSTRAINT [DF_CdsImportedHoldings_Deleted] DEFAULT ((0)),
    [ChangedBy] int NULL,
    [TimeChanged] datetime2(7) NULL
);
GO
CREATE TABLE [dbo].[CdsImportedTrades] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [TradeDate] datetime2(7) NOT NULL,
    [Exchange] nvarchar(16) NOT NULL,
    [Market] nvarchar(32) NULL,
    [ParticipantCode] nvarchar(32) NULL,
    [ClientCode] nvarchar(64) NULL,
    [CustodianCode] nvarchar(64) NULL,
    [Symbol] nvarchar(32) NOT NULL,
    [BuySell] nchar(1) NOT NULL,
    [Quantity] int NOT NULL,
    [Price] decimal(18,4) NOT NULL,
    [Settlement] decimal(18,2) NOT NULL,
    [SourceFile] nvarchar(260) NULL,
    [BatchId] int NOT NULL,
    [MatchStatus] nvarchar(20) NOT NULL CONSTRAINT [DF_CdsImportedTrades_MatchStatus] DEFAULT ('Unmatched'),
    [MatchedLotDpa] int NULL,
    [MatchedContractDpa] int NULL,
    [MatchNotes] nvarchar(500) NULL,
    [MatchedAt] datetime2(7) NULL,
    [UnmatchReason] nvarchar(32) NULL,
    [MatchedOrdDetailDpa] int NULL,
    [CreatedAt] datetime2(7) NOT NULL CONSTRAINT [DF_CdsImportedTrades_CreatedAt] DEFAULT (sysutcdatetime()),
    [CreatedBy] int NULL,
    [Deleted] bit NOT NULL CONSTRAINT [DF_CdsImportedTrades_Deleted] DEFAULT ((0)),
    [ChangedBy] int NULL,
    [TimeChanged] datetime2(7) NULL,
    [SourceFileHash] nvarchar(64) NULL
);
GO
CREATE TABLE [dbo].[Class] (
    [Class_DPA_] int NOT NULL,
    [Class_EIT_] nvarchar(40) NULL,
    [ClassDescription] nvarchar(100) NOT NULL,
    [DefaultSelection] bit NOT NULL CONSTRAINT [DF_Class_DefaultSelection] DEFAULT (0),
    [DefaultCommission] int NULL CONSTRAINT [DF_Class_DefaultCommission] DEFAULT (0),
    [IsCda] bit NOT NULL CONSTRAINT [DF_Class_IsCda] DEFAULT (0),
    [AgentStatus] bit NOT NULL CONSTRAINT [DF_Class_AgentStatus] DEFAULT (0),
    [Voucher] int NULL
);
GO
CREATE TABLE [dbo].[Client] (
    [Agent_DPA_] int NULL,
    [Branch_DPA_] int NOT NULL,
    [Class_DPA_] int NOT NULL,
    [Client_DPA_] int NOT NULL,
    [Client_EIT_] nvarchar(40) NULL,
    [ClientAddr] ntext NULL,
    [ClientBDate] datetime NULL,
    [ClientCellTel] nvarchar(100) NULL,
    [ClientContact] nvarchar(100) NULL,
    [ClientEmail] nvarchar(100) NULL,
    [ClientFax] nvarchar(100) NULL,
    [ClientHomeTel] nvarchar(100) NULL,
    [ClientIDPass] nvarchar(100) NULL,
    [ClientName] nvarchar(100) NOT NULL,
    [ClientOfficeTel] nvarchar(100) NULL,
    [ClientPhoto] nvarchar(100) NULL,
    [ClientVIP] bit NOT NULL,
    [Commission_DPA_] int NOT NULL,
    [Gender_DPA_] int NULL,
    [Owner_DPA_] int NULL,
    [Residency_DPA_] int NOT NULL,
    [ClientCDSNo] nvarchar(50) NULL,
    [ClientPAddr] ntext NULL,
    [ClientComment] ntext NULL,
    [EntityType_DPA_] int NOT NULL CONSTRAINT [DF__DbaMgr_Tm__Entit__4321E620] DEFAULT (1),
    [ClientOpeningBal] money NOT NULL CONSTRAINT [DF__DbaMgr_Tm__Clien__44160A59] DEFAULT (0),
    [ClientRegDate] datetime NOT NULL CONSTRAINT [DF_Client_ClientRegDate] DEFAULT (getdate()),
    [OnlineRegistration] bit NOT NULL CONSTRAINT [DF__DbaMgr_Tm__Onlin__45FE52CB] DEFAULT (0),
    [CreditLimit] money NOT NULL CONSTRAINT [DF_Client_CreditLimit] DEFAULT (0),
    [GenericSetting_DPA_] int NULL,
    [GenericSetting_DPA_2] int NULL,
    [GenericSetting_DPA_3] int NULL,
    [TimeChanged] datetime NULL CONSTRAINT [DF_Client_TimeChanged] DEFAULT (getdate()),
    [IsCustodian] bit NULL CONSTRAINT [DF_Client_IsCustodian] DEFAULT (0),
    [updateOnDebt] bit NULL CONSTRAINT [DF_Client_updateOnDebt] DEFAULT (0),
    [updateOnContract] bit NULL CONSTRAINT [DF_Client_updateOnContract] DEFAULT (0),
    [upDateAcc] bit NULL CONSTRAINT [DF_Client_SendSms] DEFAULT (0),
    [ChangedBy] int NULL,
    [ClientSignature] nvarchar(100) NULL,
    [Deleted] bit NULL CONSTRAINT [DF_Client_Deleted] DEFAULT (0),
    [UseContactNameInPortfolioReports] bit NULL CONSTRAINT [DF_Client_UseContactNameInPortfolioReports] DEFAULT (0),
    [IsNominee] bit NULL CONSTRAINT [DF_Client_IsNominee] DEFAULT (0),
    [IsFrozen] bit NULL CONSTRAINT [DF_Client_IsFrozen] DEFAULT (0),
    [CreatedBy] int NULL,
    [TimeCreated] datetime NOT NULL CONSTRAINT [DF_Client_TimeCreadted] DEFAULT (getdate()),
    [Institution] varchar(50) NULL,
    [Institution_DPA_] int NULL
);
GO
CREATE TABLE [dbo].[ClientBalances] (
    [client_DPA_] int NOT NULL,
    [CurrentBal] money NOT NULL
);
GO
CREATE TABLE [dbo].[clientBalancesTemp] (
    [client_DPA_] int NOT NULL,
    [CurrentBal] money NOT NULL
);
GO
CREATE TABLE [dbo].[ClientEmailsDocs] (
    [ID] int NULL,
    [Client_DPA_] int NULL,
    [Document_DPA_] int NULL,
    [EmailModeID] int NULL,
    [TimeChanged] datetime NOT NULL CONSTRAINT [DF_ClientEmailsDocs_TimeChanged] DEFAULT (getdate()),
    [Changedby] int NOT NULL,
    [Deleted] bit NOT NULL CONSTRAINT [DF_ClientEmailsDocs_Deleted] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[ClientTotal] (
    [Client_DPA_] int NOT NULL,
    [Total] money NOT NULL
);
GO
CREATE TABLE [dbo].[ClientVoucher] (
    [ClientVoucher_DPA_] int NOT NULL,
    [VoucherDate] smalldatetime NULL,
    [Voucher_EIT_] nvarchar(40) NULL,
    [VoucherPaid] bit NOT NULL CONSTRAINT [DF_ClientVoucher_VoucherPaid] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[Commission] (
    [Commission_DPA_] int IDENTITY(1,1) NOT NULL,
    [Commission_EIT_] nvarchar(40) NULL,
    [CommissionDescription] nvarchar(100) NULL,
    [CommissionRate] real NOT NULL CONSTRAINT [DF_Commission_CommissionRate] DEFAULT (0),
    [DefaultSelection] bit NOT NULL CONSTRAINT [DF_Commission_DefaultSelection] DEFAULT (0),
    [BondCommission] real NOT NULL CONSTRAINT [DF_Commission_BondCommission] DEFAULT (0),
    [SecurityBoundary] money NOT NULL CONSTRAINT [DF_Commission_CommissionBoundary] DEFAULT (0),
    [SecondSecurityBoundary] money NULL CONSTRAINT [DF_Commission_SecondSecurityBoundary] DEFAULT (0),
    [BondBoundary] money NOT NULL CONSTRAINT [DF_Commission_SecurityBoundary1] DEFAULT (0),
    [SecondBondBoundary] money NOT NULL CONSTRAINT [DF_Commission_SecondBondBoundary] DEFAULT (0),
    [UpperBondCommission] real NOT NULL CONSTRAINT [DF_Commission_UpperBondCommission] DEFAULT (0),
    [UpperSecurityCommission] real NOT NULL CONSTRAINT [DF_Commission_UpperSecurityCommission] DEFAULT (0),
    [MedianBondCommission] money NOT NULL CONSTRAINT [DF_Commission_MedianBondCommission] DEFAULT (0),
    [MedianSecurityCommission] money NOT NULL CONSTRAINT [DF_Commission_MedianSecurityCommission] DEFAULT (0),
    [MinimumBondCommission] money NOT NULL CONSTRAINT [DF_Commission_MinimumBondCommission] DEFAULT (0),
    [MinimumSecurityCommission] money NOT NULL CONSTRAINT [DF_Commission_MinimumSecurityCommission] DEFAULT (0),
    [CMARegulated] bit NOT NULL CONSTRAINT [DF_Commission] DEFAULT (0),
    [Immobilised] bit NOT NULL CONSTRAINT [DF_Commission_SystemMaintained1] DEFAULT (0),
    [SystemMaintained] tinyint NOT NULL CONSTRAINT [DF_Commission_SystemMaintained] DEFAULT (0),
    [Vatable] tinyint NOT NULL CONSTRAINT [DF_Commission_Vatable] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[CompanyInfo] (
    [SetupID] int NOT NULL,
    [CompanyName] nvarchar(50) NULL,
    [Address] ntext NULL,
    [City] nvarchar(50) NULL,
    [StateOrProvince] nvarchar(20) NULL,
    [PostalCode] nvarchar(20) NULL,
    [Country] nvarchar(50) NULL,
    [PhoneNumber] nvarchar(30) NULL,
    [FaxNumber] nvarchar(30) NULL,
    [DefaultPaymentTerms] nvarchar(255) NULL,
    [DefaultInvoiceDescription] ntext NULL,
    [BranchID] int NULL,
    [Photo] nvarchar(250) NULL,
    [Broker_DPA_] int NULL
);
GO
CREATE TABLE [dbo].[Contract] (
    [Contract_DPA_] int IDENTITY(1,1) NOT NULL,
    [Contract_EIT_] nvarchar(40) NULL,
    [ContractAPayNo] nvarchar(20) NULL,
    [ContractBPayNo] nvarchar(20) NULL,
    [ContractCPayNo] nvarchar(20) NULL,
    [ContractNCDate] smalldatetime NULL,
    [ContractNCertificate] nvarchar(20) NULL,
    [Status_DPA_] int NOT NULL CONSTRAINT [DF_Contract_Status_DPA_] DEFAULT (1),
    [ContractDelivered] bit NOT NULL CONSTRAINT [DF_Contract_ContractDelivered] DEFAULT (0),
    [ContractDeliveryDate] smalldatetime NULL,
    [ContractTransferNo] nvarchar(20) NULL,
    [ContractNCDelivered] bit NOT NULL CONSTRAINT [DF_Contract_ContractNCDelivered] DEFAULT (0),
    [Voucher_DPA_] int NULL,
    [ContractVouchered] bit NOT NULL CONSTRAINT [DF_Contract_ContractVouchered] DEFAULT (0),
    [ClientVoucher_DPA_] int NULL,
    [ContractClientVouchered] bit NOT NULL CONSTRAINT [DF_Contract_ContractClientVouchered] DEFAULT (0),
    [BrokerReceiptVoucher_DPA_] int NULL,
    [BrokerReceiptVouchered] bit NOT NULL CONSTRAINT [DF_Contract_ContractClientVouchered1] DEFAULT (0),
    [ContractSettlementDate] smalldatetime NULL,
    [IsInterBank] bit NOT NULL CONSTRAINT [DF_Contract_IsInterBank] DEFAULT (0),
    [Deleted] bit NOT NULL CONSTRAINT [DF_Contract_Deleted] DEFAULT (0),
    [ForwardRate] float NULL,
    [Basis] numeric(18,0) NULL,
    [DaysInCoupon] numeric(18,0) NULL,
    [Included] int NULL,
    [ModifiedBy] int NULL,
    [DateModified] datetime NULL,
    [ContractCertNo] nvarchar(50) NULL
);
GO
CREATE TABLE [dbo].[ContractApprovals] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [Action] nvarchar(20) NOT NULL,
    [Status] nvarchar(20) NOT NULL CONSTRAINT [DF_ContractApprovals_Status] DEFAULT ('Pending'),
    [ContractDpa] int NOT NULL,
    [ContractNumber] nvarchar(50) NULL,
    [ClientDpa] int NULL,
    [ClientName] nvarchar(100) NULL,
    [TotalGross] decimal(18,2) NULL,
    [SettlementDate] datetime2(7) NULL,
    [Reason] nvarchar(500) NULL,
    [RejectReason] nvarchar(500) NULL,
    [CreatedBy] int NOT NULL,
    [CreatedAt] datetime2(7) NOT NULL CONSTRAINT [DF_ContractApprovals_CreatedAt] DEFAULT (sysutcdatetime()),
    [ProcessedBy] int NULL,
    [ProcessedAt] datetime2(7) NULL
);
GO
CREATE TABLE [dbo].[CPortfolio] (
    [Client_DPA_] int NOT NULL,
    [CPortfolio_DPA_] int NOT NULL,
    [CPortfolio_EIT_] nvarchar(40) NULL,
    [CPortfolioPDate] smalldatetime NOT NULL,
    [CPortfolioPrice] money NOT NULL,
    [CPortfolioQty] int NOT NULL,
    [Security_DPA_] int NOT NULL,
    [Reference] nvarchar(500) NULL,
    [narrative] nvarchar(500) NULL,
    [Deleted] int NULL CONSTRAINT [DF_CPortfolio_Deleted] DEFAULT (0),
    [Modified] datetime NULL CONSTRAINT [DF_CPortfolio_Modified] DEFAULT (getdate()),
    [ModifiedBy] int NULL,
    [Offering_DPA_] int NULL,
    [RightIssue_DPA_] int NULL,
    [ShareQty] int NULL CONSTRAINT [DF_CPortfolio_RightsAllocation] DEFAULT (0),
    [EntitledQty] int NULL CONSTRAINT [DF_CPortfolio_EntitledQty] DEFAULT (0),
    [TotalPurchases] int NULL CONSTRAINT [DF_CPortfolio_TotalPurchases] DEFAULT (0),
    [TotalSales] int NULL CONSTRAINT [DF_CPortfolio_TotalSales] DEFAULT (0),
    [RenouncedQty] int NULL CONSTRAINT [DF_CPortfolio_RenouncedQty] DEFAULT (0),
    [ExtraAllotedQty] int NULL CONSTRAINT [DF_CPortfolio_ExtraAllotedQty] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[datastream_Market] (
    [MktRecord] nvarchar(2) NULL,
    [MktProcess] nvarchar(1) NULL,
    [MktDate] smalldatetime NULL,
    [MktCode] nvarchar(50) NULL,
    [MktClose] real NULL,
    [MktVolume] int NULL,
    [MktBid] real NULL,
    [MktOffer] real NULL,
    [MktHigh] real NULL,
    [MktLow] real NULL,
    [MktOpen] real NULL,
    [MktPrevious] real NULL,
    [MktUnique] int NOT NULL
);
GO
CREATE TABLE [dbo].[datastream_Securities] (
    [SecRecord] nvarchar(2) NULL,
    [SecProcess] nvarchar(1) NULL,
    [SecCode] nvarchar(50) NOT NULL,
    [SecNameFull] nvarchar(40) NULL,
    [SecNameShort] nvarchar(28) NULL,
    [SecIndustry] nvarchar(3) NULL,
    [SecCountry] nvarchar(3) NULL,
    [SecShares] int NULL,
    [SecMarket] nvarchar(2) NULL,
    [SecNominal] money NULL,
    [SecNominalCcy] nvarchar(3) NULL,
    [SecFirstListing] datetime NULL,
    [SecIssuePrice] money NULL,
    [SecIssueCcy] nvarchar(3) NULL,
    [SecTier] nvarchar(2) NULL,
    [SecKnow_DPA] int NULL
);
GO
CREATE TABLE [dbo].[down_File] (
    [down_File_DPA_] int IDENTITY(1,1) NOT NULL,
    [Filename] varchar(255) NOT NULL,
    [Report] varchar(100) NOT NULL,
    [CreatedBy] int NOT NULL,
    [CreatedByDesc] varchar(100) NOT NULL,
    [TimeCreated] smalldatetime NOT NULL CONSTRAINT [DF_down_File_TimeCreated] DEFAULT (getdate())
);
GO
CREATE TABLE [dbo].[dtproperties] (
    [id] int IDENTITY(1,1) NOT NULL,
    [objectid] int NULL,
    [property] varchar(64) NOT NULL,
    [value] varchar(255) NULL,
    [uvalue] nvarchar(255) NULL,
    [lvalue] image NULL,
    [version] int NOT NULL CONSTRAINT [DF__dtpropert__versi__1ED998B2] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[EmailDocs] (
    [Document_DPA_] int NOT NULL,
    [DocName] varchar(100) NULL
);
GO
CREATE TABLE [dbo].[EmailMode] (
    [ModeId] int NOT NULL,
    [ModeName] varchar(50) NOT NULL,
    [EmailMonth] int NULL CONSTRAINT [DF_EmailMode_EmailMonth] DEFAULT (0),
    [EmailDay] int NULL CONSTRAINT [DF_EmailMode_EmailDay] DEFAULT (0),
    [datetime] float NULL CONSTRAINT [DF_EmailMode_datetime] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[Entity] (
    [Entity_DPA_] int NOT NULL,
    [Entity_EIT_] nvarchar(40) NULL,
    [EntityName] nvarchar(100) NOT NULL,
    [EntityType_DPA_] int NOT NULL,
    [SystemMaintained] bit NOT NULL CONSTRAINT [DF_Entity_SystemMaintained] DEFAULT (0),
    [EntityOpeningBal] money NOT NULL CONSTRAINT [DF_Entity_BrokerOpeningBal] DEFAULT (0),
    [EntityRegDate] datetime NOT NULL CONSTRAINT [DF_Entity_BrokerRegDate] DEFAULT (getdate()),
    [LevySystemMaintained] int NULL,
    [EntityCode] nvarchar(50) NULL,
    [GenericSetting_DPA_] int NULL,
    [GenericSetting_DPA_2] int NULL,
    [GenericSetting_DPA_3] int NULL,
    [Quarter1] money NULL,
    [Quarter2] money NULL,
    [Quarter3] money NULL,
    [Quarter4] money NULL,
    [EntityGeneric1] int NULL,
    [EntityGeneric2] int NULL,
    [EntityGeneric3] int NULL
);
GO
CREATE TABLE [dbo].[EntityType] (
    [AccountType_DPA_] int NOT NULL,
    [EntityType_DPA_] int NOT NULL,
    [EntityType_EIT_] nvarchar(40) NULL,
    [EntityTypeName] nvarchar(100) NOT NULL,
    [DefaultSelection] bit NOT NULL CONSTRAINT [DF_EntityType_DefaultSelection] DEFAULT (0),
    [SystemMaintained] bit NOT NULL CONSTRAINT [DF_EntityType_SystemMaintained] DEFAULT (0),
    [Nominal] bit NOT NULL CONSTRAINT [DF_EntityType_Nominal] DEFAULT (0),
    [EntityTypeCode] nvarchar(50) NOT NULL,
    [AccountType_DPA_3] int NULL,
    [AccountType_DPA_2] int NULL
);
GO
CREATE TABLE [dbo].[excep_SummaryHoldings] (
    [Trans_DPA_] int IDENTITY(1,1) NOT NULL,
    [Code] int NOT NULL,
    [SecCode] int NOT NULL,
    [CDSQty] money NOT NULL,
    [BKQty] money NOT NULL,
    [DiffQty] money NOT NULL
);
GO
CREATE TABLE [dbo].[ForwardRate] (
    [ForwardRate_DPA_] int NOT NULL,
    [ForwardRate] float NULL,
    [ActivationDate] datetime NULL,
    [DateModified] datetime NULL,
    [ModifiedBy] int NULL,
    [Bond_DPA_] int NOT NULL
);
GO
CREATE TABLE [dbo].[Gender] (
    [Gender_DPA_] int NOT NULL,
    [Gender_EIT_] nvarchar(40) NULL,
    [GenderDescription] nvarchar(100) NOT NULL,
    [DefaultSelection] bit NOT NULL CONSTRAINT [DF_Gender_DefaultSelection] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[Generic] (
    [Generic_DPA_] int NOT NULL,
    [Generic_EIT_] nvarchar(40) NULL CONSTRAINT [DF_Generic_Generic_EIT_] DEFAULT (newid()),
    [GenericDescription] nvarchar(400) NOT NULL
);
GO
CREATE TABLE [dbo].[GenericSetting] (
    [GenericSetting_DPA_] int NOT NULL,
    [GenericSetting_EIT_] nvarchar(40) NULL,
    [GenericSettingDescription] nvarchar(400) NOT NULL,
    [Generic_DPA_] int NOT NULL,
    [EntityType_DPA_] int NOT NULL
);
GO
CREATE TABLE [dbo].[Groups] (
    [GroupID] int IDENTITY(1,1) NOT NULL,
    [GroupName] nvarchar(50) NULL,
    [Description] nvarchar(50) NULL
);
GO
CREATE TABLE [dbo].[Holdings] (
    [HoldingID] int IDENTITY(1,1) NOT NULL,
    [TradeDate] datetime NULL,
    [TradeTime] varchar(12) NULL,
    [Client_DPA_] int NULL,
    [Security_DPA_] int NULL,
    [Quantity] varchar(50) NULL,
    [AccountStatus] varchar(255) NULL,
    [BalanceFree] char(1) NULL
);
GO
CREATE TABLE [dbo].[Holidays] (
    [Holiday_DPA_] int NOT NULL,
    [Description] nvarchar(50) NOT NULL,
    [Holiday] datetime NOT NULL,
    [Recurring] bit NOT NULL CONSTRAINT [DF_Holidays_Recurring] DEFAULT ((0))
);
GO
CREATE TABLE [dbo].[Institution] (
    [Institution_DPA_] int IDENTITY(1,1) NOT NULL,
    [InstitutionName] nvarchar(50) NULL,
    [Address] nvarchar(50) NULL,
    [PostalAddress] nvarchar(50) NULL,
    [PhoneNumber] nvarchar(50) NULL,
    [fax] nvarchar(50) NULL
);
GO
CREATE TABLE [dbo].[InterTransfer] (
    [InterTransfer_DPA_] int IDENTITY(1,1) NOT NULL,
    [SourceEntityType_DPA_] int NOT NULL,
    [InterTransfer_EIT_] nvarchar(50) NULL,
    [SourceEntity_DPA_] int NOT NULL,
    [TargetEntityType_DPA_] int NOT NULL,
    [TargetEntity_DPA_] int NOT NULL,
    [TransferAmount] money NOT NULL,
    [TransferDate] smalldatetime NULL,
    [ChangedBy] int NOT NULL,
    [TimeChanged] smalldatetime NOT NULL CONSTRAINT [DF_InterTransfer_TimeChanged] DEFAULT (getdate()),
    [TransferReference] nvarchar(50) NULL,
    [TransferNarrative] nvarchar(100) NULL,
    [InterTransferType_DPA_] int NOT NULL,
    [Contract_DPA_] int NOT NULL,
    [Deleted] bit NOT NULL CONSTRAINT [DF_InterTransfer_Deleted] DEFAULT (0),
    [ReconcileDate] smalldatetime NULL
);
GO
CREATE TABLE [dbo].[InterTransferType] (
    [InterTransferType_DPA_] int NOT NULL,
    [InterTransfer_EIT_] nvarchar(100) NULL,
    [TypeDescription] nvarchar(50) NOT NULL
);
GO
CREATE TABLE [dbo].[Journal] (
    [Journal_DPA_] int IDENTITY(1,1) NOT NULL,
    [Journal_EIT_] nvarchar(40) NULL,
    [JournalDate] datetime NOT NULL,
    [JournalNarrative] nvarchar(500) NULL,
    [UserID] int NOT NULL,
    [JournalCommitted] bit NULL CONSTRAINT [DF_Journal_JournalCommitted_1] DEFAULT (0),
    [Released] int NULL CONSTRAINT [DF_Journal_Released] DEFAULT (0),
    [ReleaseDate] datetime NULL,
    [Deleted] bit NULL CONSTRAINT [DF_Journal_Deleted] DEFAULT (0),
    [ChangedBy] int NULL,
    [TimeChanged] datetime NULL
);
GO
CREATE TABLE [dbo].[JournalEntry] (
    [JournalEntry_DPA_] int NOT NULL,
    [JournalEntry_EIT_] nvarchar(40) NULL,
    [EntityType_DPA_] int NOT NULL,
    [Entity_DPA_] int NOT NULL,
    [JournalEntryDebit] money NOT NULL CONSTRAINT [DF_JournalEntry_JournalEntryDebit] DEFAULT (0),
    [JournalEntryCredit] money NOT NULL CONSTRAINT [DF_JournalEntry_JournalEntryCredit] DEFAULT (0),
    [Journal_DPA_] int NOT NULL,
    [ReconcileDate] datetime NULL,
    [Deleted] bit NULL CONSTRAINT [DF_JournalEntry_Deleted] DEFAULT (0),
    [ChangedBy] int NULL,
    [TimeChanged] datetime NULL,
    [Narrative] nvarchar(100) NULL
);
GO
CREATE TABLE [dbo].[Levy] (
    [Levy_DPA_] int IDENTITY(1,1) NOT NULL,
    [Levy_EIT_] nvarchar(40) NULL,
    [LevyActive] bit NOT NULL,
    [LevyAmount] real NOT NULL,
    [LevyAppBond] bit NOT NULL,
    [LevyAppSecurity] bit NOT NULL,
    [LevyBlock] int NULL,
    [LevyDescription] nvarchar(100) NOT NULL,
    [LevyType] nvarchar(2) NOT NULL,
    [LevyShortName] nvarchar(100) NULL,
    [SystemMaintained] int NOT NULL CONSTRAINT [DF_Levy_SystemMaintained] DEFAULT (0),
    [Vatable] varchar(20) NULL
);
GO
CREATE TABLE [dbo].[LevyContract] (
    [Contract_DPA_] int NOT NULL,
    [LevyVATAmount] money NOT NULL CONSTRAINT [DF_LevyContract_LevyVATAmount] DEFAULT (0),
    [LevyAmount] money NOT NULL,
    [LevyName] nvarchar(100) NOT NULL,
    [LevyContract_DPA_] int IDENTITY(1,1) NOT NULL,
    [LevyContract_EIT_] nvarchar(40) NULL,
    [LevyRate] real NULL,
    [LevyBlock] real NULL,
    [LevyShortName] nvarchar(50) NULL,
    [LevyRatePercentage] nvarchar(50) NULL,
    [SystemMaintained] int NULL,
    [ShortName] nvarchar(50) NULL,
    [Deleted] bit NOT NULL CONSTRAINT [DF_LevyContract_Deleted_1] DEFAULT (0),
    [ChangedBy] int NULL,
    [TimeChanged] smalldatetime NULL
);
GO
CREATE TABLE [dbo].[LevyReportOrder] (
    [LevyOrder_DPA_] int NOT NULL,
    [LevyOrder_EIT_] nvarchar(40) NULL,
    [LevyName] nvarchar(1000) NOT NULL,
    [LevyOrder] int NOT NULL CONSTRAINT [DF_LevyReportOrder_LevyOrder] DEFAULT (100)
);
GO
CREATE TABLE [dbo].[LevySecurity] (
    [LevySecurity_DPA_] int IDENTITY(1,1) NOT NULL,
    [LevySecurity_EIT_] nvarchar(40) NULL,
    [Levy_DPA_] int NOT NULL,
    [Security_DPA_] int NOT NULL
);
GO
CREATE TABLE [dbo].[Lot] (
    [Lot_DPA_] int IDENTITY(1,1) NOT NULL,
    [Contract_DPA_] int NOT NULL,
    [OrdDetail_DPA_] int NOT NULL,
    [LotPrice] money NOT NULL,
    [LotQty] int NOT NULL,
    [UniqueSlip] varchar(20) NULL,
    [LotSlipNo] nvarchar(20) NOT NULL,
    [LotTDate] datetime NOT NULL,
    [Broker_DPA_] int NOT NULL,
    [ContractNumber] nvarchar(50) NOT NULL,
    [LotGrossAmount] money NULL,
    [CDSImport_DPA_] int NULL,
    [CDSTransaction] bit NULL CONSTRAINT [DF_Lot_CDSTransaction_1] DEFAULT (0),
    [ChangedBy] int NOT NULL CONSTRAINT [DF_Lot_ChangedBy] DEFAULT (0),
    [ContractSettlementDate] datetime NOT NULL CONSTRAINT [DF_Lot_ContractSettlementDate] DEFAULT (getdate()),
    [TimeChanged] smalldatetime NOT NULL CONSTRAINT [DF_Lot_TimeChanged] DEFAULT (getdate()),
    [Deleted] bit NOT NULL CONSTRAINT [DF_Lot_Deleted] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[Lot_Deleted] (
    [Lot_DPA_] int NOT NULL,
    [Contract_DPA_] int NOT NULL,
    [OrdDetail_DPA_] int NOT NULL,
    [LotPrice] money NOT NULL,
    [LotQty] int NOT NULL,
    [UniqueSlip] varchar(20) NULL,
    [LotSlipNo] nvarchar(20) NOT NULL,
    [LotTDate] datetime NOT NULL,
    [Broker_DPA_] int NOT NULL,
    [ContractNumber] nvarchar(50) NOT NULL,
    [LotGrossAmount] money NULL,
    [CDSImport_DPA_] int NULL,
    [CDSTransaction] bit NULL CONSTRAINT [DF_Lot_Deleted_CDSTransaction_1] DEFAULT (0),
    [ChangedBy] int NOT NULL CONSTRAINT [DF_Lot_Deleted_ChangedBy] DEFAULT (0),
    [ContractSettlementDate] datetime NOT NULL CONSTRAINT [DF_Lot_Deleted_ContractSettlementDate] DEFAULT (getdate()),
    [TimeChanged] smalldatetime NOT NULL CONSTRAINT [DF_Lot_Deleted_TimeChanged] DEFAULT (getdate()),
    [Deleted] bit NOT NULL CONSTRAINT [DF_Lot_Deleted_Deleted] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[MailConfiguration] (
    [MailConfig_DPA_] int IDENTITY(1,1) NOT NULL,
    [SMTPServer] nvarchar(50) NOT NULL,
    [SMTPServerPort] int NOT NULL,
    [SMTPConnectionTimeout] int NOT NULL,
    [SMTPAuthenticate] int NOT NULL,
    [SendUserName] nvarchar(100) NOT NULL,
    [SendPassword] nvarchar(100) NOT NULL,
    [SendUsingMethod] int NOT NULL,
    [SendDisplayName] nvarchar(100) NOT NULL
);
GO
CREATE TABLE [dbo].[MarketQuotes] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [SecurityDpa] int NOT NULL,
    [QuoteDate] datetime2(7) NOT NULL,
    [Open] decimal(18,4) NULL,
    [High] decimal(18,4) NULL,
    [Low] decimal(18,4) NULL,
    [Close] decimal(18,4) NULL,
    [PreviousClose] decimal(18,4) NULL,
    [Bid] decimal(18,4) NULL,
    [Offer] decimal(18,4) NULL,
    [Volume] bigint NULL,
    [Exchange] nvarchar(16) NULL,
    [BatchId] int NULL,
    [CreatedAt] datetime2(7) NOT NULL CONSTRAINT [DF_MarketQuotes_CreatedAt] DEFAULT (sysutcdatetime()),
    [CreatedBy] int NULL,
    [TimeChanged] datetime2(7) NULL,
    [ChangedBy] int NULL
);
GO
CREATE TABLE [dbo].[MarketSector] (
    [Sector_DPA_] int NOT NULL,
    [ShortDescription] nvarchar(100) NOT NULL
);
GO
CREATE TABLE [dbo].[MenuGroups] (
    [ID] int IDENTITY(1,1) NOT NULL,
    [MenuID] int NOT NULL,
    [GroupID] int NULL,
    [CanAdd] int NULL,
    [CanEdit] int NULL,
    [CanDelete] int NULL,
    [CanSort] int NULL,
    [CanFilter] int NULL,
    [CanSearch] int NULL
);
GO
CREATE TABLE [dbo].[Menus] (
    [MenuID] int IDENTITY(1,1) NOT NULL,
    [mnuCaption] nvarchar(50) NULL,
    [mnuDescription] nvarchar(100) NULL,
    [mnuAction] nvarchar(255) NULL,
    [Submenu] int NULL,
    [ismainMenu] int NULL,
    [MainMenuID] int NULL,
    [IsReport] int NULL,
    [Image] nvarchar(255) NULL,
    [mnuType] int NULL,
    [DefaultChildID] int NOT NULL
);
GO
CREATE TABLE [dbo].[MenuTypes] (
    [MenuTypeID] int NULL,
    [Description] nvarchar(250) NULL
);
GO
CREATE TABLE [dbo].[Offerings] (
    [Offering_DPA_] int IDENTITY(1,1) NOT NULL,
    [PAL_No] nvarchar(50) NULL,
    [Client_DPA_] int NULL,
    [Offering] int NULL,
    [Offering_Price] float NULL,
    [ID_No] nvarchar(50) NULL,
    [Alloted_Rights] float NULL,
    [Accepted_Rights] float NULL,
    [Renouncee] nvarchar(50) NULL,
    [Receipt] int NULL,
    [Submitted] bit NULL CONSTRAINT [DF_Offerings_a_Submitted] DEFAULT (0),
    [Submission_Date] datetime NULL,
    [Ref_No] nvarchar(50) NULL,
    [Batch_No] int NULL,
    [Forward] bit NULL CONSTRAINT [DF_Offerings_a_Forward] DEFAULT (0),
    [Offerings_Date] datetime NULL CONSTRAINT [DF_Offerings_a_Offerings_Date] DEFAULT (getdate()),
    [OfferCheque] char(30) NULL,
    [OfferBank] nvarchar(100) NULL,
    [BankCode] nvarchar(50) NULL,
    [AccountNo] nvarchar(50) NULL,
    [TimeChanged] datetime NULL CONSTRAINT [DF_Offerings_a_TimeChanged] DEFAULT (getdate()),
    [Deleted] bit NULL CONSTRAINT [DF_Offerings_a_Deleted] DEFAULT (0),
    [ChangedBy] int NULL,
    [Offering_EIT_] nvarchar(60) NULL,
    [Paid] bit NULL CONSTRAINT [DF_Offerings_a_Paid] DEFAULT (0),
    [AppType] char(1) NULL CONSTRAINT [DF_Offerings_a_AppType] DEFAULT ('N'),
    [shNAME2] nvarchar(100) NULL,
    [BatchSeq] int NULL,
    [Payment_DPA_] int NULL,
    [Download_DPA_] int NULL,
    [Downloaded] bit NULL CONSTRAINT [DF_Offerings_a_Downloaded] DEFAULT (0),
    [LastDownLoaded] datetime NULL CONSTRAINT [DF_Offerings_a_LastDownLoaded] DEFAULT (0),
    [BatchFileName] nvarchar(50) NULL,
    [Additional] int NULL,
    [AcceptanceType] int NULL,
    [CreatedBy] int NULL,
    [DateCreated] datetime NULL,
    [Download_Date] smalldatetime NULL,
    [Branch_DPA_] int NULL,
    [Agent_DPA_] int NULL,
    [PaymentMode] int NULL,
    [BatchClosed] tinyint NULL,
    [Extra] char(10) NULL,
    [BatchPaymentMode] char(10) NULL,
    [sms] char(10) NULL,
    [ClientCellTel] nvarchar(50) NULL,
    [smsNotes] char(10) NULL,
    [CDSNumeric] char(10) NULL,
    [RefundMethod] char(10) NULL,
    [TaxExempt] char(10) NULL,
    [DividendMethod] char(10) NULL,
    [EFTBankRef] nvarchar(50) NULL,
    [EFTBranchRef] nvarchar(50) NULL,
    [EFTSortCode] varchar(20) NULL,
    [EFTAccountNo] varchar(20) NULL,
    [Category] char(10) NULL,
    [PaymentType] char(10) NULL,
    [PaymentRef] varchar(20) NULL,
    [PaymentBankRef] nvarchar(50) NULL,
    [PaymentBranchRef] nvarchar(50) NULL,
    [PaymentSortCode] varchar(20) NULL,
    [PaymentAccountNo] varchar(20) NULL,
    [OfferingType] char(10) NULL CONSTRAINT [DF_Offerings_OfferingType] DEFAULT (1),
    [CitiAccepted] char(5) NULL,
    [CitiReason] nvarchar(100) NULL,
    [Updated] int NULL,
    [CitiBatchNo] char(10) NULL,
    [CitiSerialNo] char(10) NULL,
    [Certificate] bit NOT NULL CONSTRAINT [DF_Offerings_Certificate] DEFAULT (0),
    [CDSCharge] float NOT NULL CONSTRAINT [DF_Offerings_CDSCharge] DEFAULT (0),
    [Hold] int NULL,
    [ISSMSSend] int NULL,
    [CDACode] char(5) NULL CONSTRAINT [DF_Offerings_CDACode] DEFAULT ('B23B'),
    [ReceivingBroker] int NULL,
    [AgentCode] varchar(6) NULL,
    [Status] tinyint NULL,
    [CDSPaid] tinyint NULL CONSTRAINT [DF_Offerings_CDSPaid] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[OfferType] (
    [OfferType_DPA_] int IDENTITY(1,1) NOT NULL,
    [Description] nvarchar(500) NULL,
    [Deleted] int NULL,
    [TimeModified] datetime NULL,
    [ModifiedBy] int NULL
);
GO
CREATE TABLE [dbo].[OnlineClient] (
    [Agent_DPA_] int NULL,
    [Branch_DPA_] int NULL,
    [Class_DPA_] int NULL,
    [Client_DPA_] int NOT NULL,
    [Client_EIT_] nvarchar(40) NULL,
    [ClientAddr] ntext NULL,
    [ClientBDate] smalldatetime NULL,
    [ClientCellTel] nvarchar(100) NULL,
    [ClientContact] nvarchar(100) NULL,
    [ClientEmail] nvarchar(100) NULL,
    [ClientFax] nvarchar(100) NULL,
    [ClientHomeTel] nvarchar(100) NULL,
    [ClientIDPass] nvarchar(100) NULL,
    [ClientName] nvarchar(100) NULL,
    [ClientOfficeTel] nvarchar(100) NULL,
    [ClientPhoto] nvarchar(100) NULL,
    [ClientSignature] nvarchar(100) NULL,
    [ClientVIP] bit NULL,
    [Commission_DPA_] int NULL,
    [Gender_DPA_] int NULL,
    [Owner_DPA_] int NULL,
    [Residency_DPA_] int NULL,
    [ClientCDSNo] nvarchar(20) NULL,
    [ClientPAddr] ntext NULL,
    [ClientComment] ntext NULL,
    [EntityType_DPA_] int NOT NULL CONSTRAINT [DF_OnlineClient_EntityType_DPA_] DEFAULT (1),
    [ClientOpeningBal] money NULL,
    [ClientRegDate] datetime NULL,
    [OnlineRegistration] bit NULL,
    [CreditLimit] money NULL,
    [GenericSetting_DPA_] int NULL,
    [GenericSetting_DPA_2] int NULL,
    [GenericSetting_DPA_3] int NULL,
    [OldAccountNo] int NULL,
    [ChangedBy] int NULL CONSTRAINT [DF_OnlineClient_ChangedBy] DEFAULT (125),
    [TimeChanged] smalldatetime NULL CONSTRAINT [DF_OnlineClient_TimeChanged] DEFAULT (getdate()),
    [IsCustodian] bit NULL CONSTRAINT [DF_OnlineClient_IsCustodian] DEFAULT (0),
    [Deleted] bit NULL CONSTRAINT [DF_OnlineClient_Deleted] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[OrdDetail] (
    [OrdDetail_DPA_] int NOT NULL,
    [OrdDetail_EIT_] nvarchar(40) NULL,
    [OrdDetailCertNo] nvarchar(100) NULL,
    [OrdDetailPrice] nvarchar(20) NOT NULL,
    [OrdDetailQty] int NOT NULL,
    [OrdDetailValidity] datetime NULL,
    [Order_DPA_] int NOT NULL,
    [Security_DPA_] int NOT NULL,
    [OrdDetailCompound] bit NOT NULL CONSTRAINT [DF_OrdDetail_OrdDetailCompound] DEFAULT (0),
    [BondDescription] nvarchar(100) NULL,
    [Bond_DPA_] int NULL CONSTRAINT [DF_OrdDetail_Bond_DPA_] DEFAULT (0),
    [Amount] money NULL CONSTRAINT [DF_OrdDetail_Amount] DEFAULT (0),
    [Best] bit NULL CONSTRAINT [DF_OrdDetail_Best] DEFAULT (0),
    [Deleted] bit NULL CONSTRAINT [DF_OrdDetail_Deleted] DEFAULT (0),
    [Limit] int NULL
);
GO
CREATE TABLE [dbo].[OrderHoldOptions] (
    [OrderHoldOptionID] int NOT NULL,
    [Description] nvarchar(100) NOT NULL,
    [RequiresDate] int NOT NULL CONSTRAINT [DF_OrderHoldOptions_RequiresDate] DEFAULT (0),
    [DefaultSelection] bit NULL CONSTRAINT [DF_OrderHoldOptions_DefaultSelection] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[OrderHoldType] (
    [OrderHoldType_DPA_] int IDENTITY(1,1) NOT NULL,
    [OrderHoldTypeName] nvarchar(100) NOT NULL,
    [DefaultSelection] bit NOT NULL CONSTRAINT [DF_OrderHoldType_DefaultSelection] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[OrderSecType] (
    [OrderSecType_DPA_] int NOT NULL,
    [OrderSecType_EIT_] nvarchar(40) NULL,
    [OrderSecTypeDescription] nvarchar(10) NOT NULL,
    [OrderSecTypeDisplayName] nvarchar(20) NOT NULL,
    [DefaultSelection] bit NOT NULL CONSTRAINT [DF_OrderSecType_DefaultSelection] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[OrderType] (
    [OrderType_DPA_] int NOT NULL,
    [OrderType_EIT_] nvarchar(40) NULL,
    [OrderTypeDescription] nvarchar(100) NOT NULL,
    [OrderTypeSale] tinyint NOT NULL,
    [DefaultSelection] bit NOT NULL CONSTRAINT [DF_OrderType_DefaultSelection] DEFAULT (0),
    [RequireCertificate] bit NOT NULL CONSTRAINT [DF_OrderType_RequireCertificate] DEFAULT (0),
    [HandlingFee] float NULL
);
GO
CREATE TABLE [dbo].[Owner] (
    [Owner_DPA_] int NOT NULL,
    [Owner_EIT_] nvarchar(40) NULL,
    [OwnerFname] nvarchar(100) NOT NULL,
    [OwnerLName] nvarchar(100) NOT NULL,
    [CommissionRate] real NULL CONSTRAINT [DF_Owner_CommissionRate] DEFAULT (0),
    [OwnerOpeningBal] money NOT NULL CONSTRAINT [DF_Owner_BrokerOpeningBal] DEFAULT (0),
    [OwnerRegDate] datetime NOT NULL CONSTRAINT [DF_Owner_BrokerRegDate] DEFAULT (getdate()),
    [EntityType_DPA_] int NOT NULL CONSTRAINT [DF_Owner_EntityType_DPA_] DEFAULT (7),
    [IdNo] nvarchar(50) NULL,
    [MobileNo] float NULL,
    [SendSMS] int NULL CONSTRAINT [DF_Owner_SendSMS] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[Payment] (
    [BankAccount_DPA_] int NOT NULL,
    [EntityType_DPA_] int NOT NULL,
    [Payment_DPA_] int NOT NULL,
    [Payment_EIT_] nvarchar(40) NULL,
    [PaymentAmount] money NOT NULL,
    [PaymentReceiptNo] int NULL,
    [PaymentPDate] datetime NOT NULL,
    [PaymentReference] nvarchar(20) NULL,
    [PayType_DPA_] int NOT NULL,
    [PaymentNarrative] nvarchar(200) NULL,
    [Entity_DPA_] int NOT NULL,
    [Order_DPA_] int NULL,
    [Voucher_DPA_] int NULL,
    [ClientVoucher_DPA_] int NULL,
    [BrokerReceiptVoucher_DPA_] int NULL,
    [ChequeCollectionModifyUser] int NULL,
    [ChequeCollectionUser] int NULL,
    [ChequeCollectionDate] datetime NULL,
    [ChequeCollectionModifyDate] datetime NULL,
    [ChequeCollectionNarrative] nvarchar(100) NULL,
    [ChequeCollection] nvarchar(100) NULL,
    [PaymentTypes_DPA_] int NULL,
    [Deleted] bit NULL CONSTRAINT [DF_Payment_Deleted] DEFAULT (0),
    [ChangedBy] int NULL,
    [TimeChanged] datetime NULL,
    [Contract_DPA_] int NULL,
    [ReconcileDate] datetime NULL,
    [Batch_No] int NULL,
    [TimeCreated] datetime NULL CONSTRAINT [DF_Payment_TimeCreated] DEFAULT (getdate()),
    [CreatedBy] int NULL
);
GO
CREATE TABLE [dbo].[PaymentApprovals] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [Action] nvarchar(20) NOT NULL,
    [Status] nvarchar(20) NOT NULL CONSTRAINT [DF_PaymentApprovals_Status] DEFAULT ('Pending'),
    [TargetPaymentDpa] int NULL,
    [Amount] decimal(18,2) NULL,
    [PaymentDate] datetime2(7) NULL,
    [Reference] nvarchar(50) NULL,
    [Narrative] nvarchar(500) NULL,
    [PayTypeDpa] int NULL,
    [EntityTypeDpa] int NULL,
    [EntityDpa] int NULL,
    [PaymentTypesDpa] int NULL,
    [BankAccountDpa] int NULL,
    [RejectReason] nvarchar(500) NULL,
    [CreatedBy] int NOT NULL,
    [CreatedAt] datetime2(7) NOT NULL CONSTRAINT [DF_PaymentApprovals_CreatedAt] DEFAULT (sysutcdatetime()),
    [ProcessedBy] int NULL,
    [ProcessedAt] datetime2(7) NULL,
    [PaymentRequestId] int NULL
);
GO
CREATE TABLE [dbo].[PaymentRequest] (
    [BankAccount_DPA_] int NOT NULL,
    [AccountToUse] nvarchar(50) NOT NULL,
    [PaymentRequest_DPA_] int NOT NULL,
    [PaymentAmount] money NOT NULL,
    [PaymentPDate] datetime NOT NULL,
    [PaymentReference] nvarchar(20) NULL,
    [PayType_DPA_] int NOT NULL,
    [PaymentNarrative] nvarchar(200) NULL,
    [Entity_DPA_] int NOT NULL,
    [PaymentTypes_DPA_] int NULL,
    [Deleted] bit NULL,
    [ChangedBy] int NULL,
    [TimeChanged] datetime NULL,
    [Status] nvarchar(50) NULL
);
GO
CREATE TABLE [dbo].[PaymentRequests] (
    [Request_DPA_] int NOT NULL,
    [Client_DPA_] int NOT NULL,
    [RequestDate] datetime NOT NULL,
    [RequestPayDate] datetime NULL,
    [PaidDate] datetime NULL,
    [BankAccount_DPA_] int NULL,
    [RequestNarrative] nvarchar(100) NULL,
    [PaymentNarrative] nvarchar(100) NULL,
    [PaymentReference] nvarchar(50) NULL,
    [PaymentType_DPA_] int NULL,
    [Approved] smallint NOT NULL,
    [PaymentAmount] float NULL,
    [CreatedBy] int NULL,
    [ModifiedBy] int NULL,
    [TimeModified] datetime NOT NULL,
    [Deleted] bit NOT NULL,
    [ApprovalDate] datetime NULL,
    [ApprovedBy] int NULL,
    [Contract_DPA_] char(10) NULL,
    [Processed_DPA_] int NULL,
    [ProcessedDate] datetime NULL,
    [FirstApproval] bit NOT NULL,
    [FirstBy] int NULL,
    [FirstWhen] datetime NOT NULL,
    [ReconcileDate] datetime NULL
);
GO
CREATE TABLE [dbo].[PaymentTypes] (
    [PaymentTypes_DPA_] int NOT NULL,
    [Description] nvarchar(500) NULL,
    [Reference] int NULL
);
GO
CREATE TABLE [dbo].[PaymentVoucherSettlementDates] (
    [Contract_DPA_] int NOT NULL,
    [SettlementDate] smalldatetime NOT NULL
);
GO
CREATE TABLE [dbo].[PayType] (
    [PayType_DPA_] int NOT NULL,
    [PayType_EIT_] nvarchar(40) NULL,
    [PayTypeDescription] nvarchar(100) NOT NULL,
    [PayTypeIn] bit NOT NULL
);
GO
CREATE TABLE [dbo].[PortalPaymentRequests] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [ClientDpa] int NOT NULL,
    [RequestType] nvarchar(20) NOT NULL,
    [Amount] decimal(18,2) NOT NULL,
    [Reference] nvarchar(50) NULL,
    [Narrative] nvarchar(500) NULL,
    [Status] nvarchar(20) NOT NULL CONSTRAINT [DF_PortalPaymentRequests_Status] DEFAULT ('Pending'),
    [RejectReason] nvarchar(500) NULL,
    [PortalUserId] int NULL,
    [CreatedAt] datetime2(7) NOT NULL CONSTRAINT [DF_PortalPaymentRequests_CreatedAt] DEFAULT (sysutcdatetime()),
    [CreatedBy] int NULL,
    [ProcessedBy] int NULL,
    [ProcessedAt] datetime2(7) NULL,
    [PaymentDpa] int NULL,
    [ClientBankAccountDpa] int NULL
);
GO
CREATE TABLE [dbo].[PortalRefreshTokens] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [PortalUserId] int NOT NULL,
    [Token] nvarchar(512) NOT NULL,
    [ExpiresAt] datetime2(7) NOT NULL,
    [CreatedAt] datetime2(7) NOT NULL,
    [Revoked] bit NOT NULL
);
GO
CREATE TABLE [dbo].[PortalUsers] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [Email] nvarchar(256) NOT NULL,
    [PasswordHash] nvarchar(512) NOT NULL,
    [FirstName] nvarchar(100) NOT NULL,
    [LastName] nvarchar(100) NOT NULL,
    [Phone] nvarchar(50) NULL,
    [OfficePhone] nvarchar(50) NULL,
    [HomePhone] nvarchar(50) NULL,
    [IdNumber] nvarchar(50) NULL,
    [CdsNumber] nvarchar(50) NULL,
    [DateOfBirth] datetime2(7) NULL,
    [PhysicalAddress] nvarchar(500) NULL,
    [PostalAddress] nvarchar(500) NULL,
    [ContactPerson] nvarchar(200) NULL,
    [Role] nvarchar(20) NOT NULL,
    [ClientDpa] int NULL,
    [Status] nvarchar(20) NOT NULL,
    [Active] bit NOT NULL,
    [CreatedAt] datetime2(7) NOT NULL,
    [ApprovedAt] datetime2(7) NULL,
    [ApprovedBy] int NULL,
    [RejectionReason] nvarchar(500) NULL,
    [MustChangePassword] bit NOT NULL CONSTRAINT [DF_PortalUsers_MustChangePassword] DEFAULT ((0)),
    [PasswordChangedAt] datetime2(7) NULL,
    [Username] nvarchar(100) NULL,
    [LegacyUserId] int NULL,
    [CurrentSessionId] uniqueidentifier NULL,
    [LastSeenAt] datetime2(7) NULL,
    [AgentDpa] int NULL,
    [CanCreate] bit NOT NULL CONSTRAINT [DF_PortalUsers_CanCreate] DEFAULT ((1)),
    [CanEdit] bit NOT NULL CONSTRAINT [DF_PortalUsers_CanEdit] DEFAULT ((1)),
    [CanDelete] bit NOT NULL CONSTRAINT [DF_PortalUsers_CanDelete] DEFAULT ((1)),
    [CanApprove] bit NOT NULL CONSTRAINT [DF_PortalUsers_CanApprove] DEFAULT ((1))
);
GO
CREATE TABLE [dbo].[PortfoliosQuantities] (
    [Client_DPA_] int NOT NULL,
    [TradeDate] smalldatetime NOT NULL,
    [SlipNo] nvarchar(50) NOT NULL,
    [ContractNumber] nvarchar(50) NOT NULL,
    [Security_DPA_] int NOT NULL,
    [SecurityCode] char(10) NOT NULL,
    [Type] char(10) NOT NULL,
    [Quantity] float NOT NULL CONSTRAINT [DF_PortfoliosQuantities_Quantity] DEFAULT (0),
    [Price] money NOT NULL CONSTRAINT [DF_PortfoliosQuantities_Price] DEFAULT (0),
    [Balance] money NOT NULL CONSTRAINT [DF_PortfoliosQuantities_Balance] DEFAULT (0),
    [AveragePrice] money NOT NULL CONSTRAINT [DF_PortfoliosQuantities_AveragePrice] DEFAULT (0),
    [EntryQuantity] float NOT NULL CONSTRAINT [DF_PortfoliosQuantities_EntryQuantity] DEFAULT (0),
    [EntryPrice] money NOT NULL CONSTRAINT [DF_PortfoliosQuantities_EntryPrice] DEFAULT (0),
    [ValueDate] smalldatetime NOT NULL,
    [Cash] money NOT NULL CONSTRAINT [DF_PortfoliosQuantities_Cash] DEFAULT (0),
    [MarketPrice] money NOT NULL CONSTRAINT [DF_PortfoliosQuantities_MarketPrice] DEFAULT (0),
    [BookValue] money NOT NULL CONSTRAINT [DF_PortfoliosQuantities_BookValue] DEFAULT (0),
    [CurrentValue] money NOT NULL CONSTRAINT [DF_PortfoliosQuantities_CurrentValue] DEFAULT (0),
    [PL] money NOT NULL CONSTRAINT [DF_PortfoliosQuantities_PL] DEFAULT (0),
    [PL_] float NOT NULL CONSTRAINT [DF_PortfoliosQuantities_PL_] DEFAULT (0),
    [SecurityBookValue] money NOT NULL CONSTRAINT [DF_PortfoliosQuantities_SecurityBookValue] DEFAULT (0),
    [SecurityCurrentValue] money NOT NULL CONSTRAINT [DF_PortfoliosQuantities_SecurityCurrentValue] DEFAULT (0),
    [ClientBookValue] money NOT NULL CONSTRAINT [DF_PortfoliosQuantities_ClientBookValue] DEFAULT (0),
    [ClientCurrentValue] money NOT NULL CONSTRAINT [DF_PortfoliosQuantities_ClientCurrentValue] DEFAULT (0),
    [TotalPortfolio] money NOT NULL CONSTRAINT [DF_PortfoliosQuantities_TotalPortfolio] DEFAULT (0),
    [LastQuantity] float NOT NULL CONSTRAINT [DF_PortfoliosQuantities_LastQuantity] DEFAULT (0),
    [LastPrice] money NOT NULL CONSTRAINT [DF_PortfoliosQuantities_LastPrice] DEFAULT (0),
    [Cost] money NOT NULL CONSTRAINT [DF_PortfoliosQuantities_Cost] DEFAULT (0),
    [Sales] money NOT NULL CONSTRAINT [DF_PortfoliosQuantities_Sales] DEFAULT (0),
    [Profit] money NOT NULL CONSTRAINT [DF_PortfoliosQuantities_Profit] DEFAULT (0),
    [Profit_] float NOT NULL CONSTRAINT [DF_PortfoliosQuantities_Profit_] DEFAULT (0),
    [Portfolio_DPA_] int IDENTITY(1,1) NOT NULL,
    [entryid] char(10) NOT NULL,
    [userid] int NOT NULL,
    [NetAmount] money NOT NULL CONSTRAINT [DF_PortfoliosQuantities_NetAmount] DEFAULT (0),
    [LastPL] money NOT NULL CONSTRAINT [DF_PortfoliosQuantities_LastPL] DEFAULT (0),
    [LastPL_] float NOT NULL CONSTRAINT [DF_PortfoliosQuantities_LastPL_] DEFAULT (0),
    [ClientPL] money NOT NULL CONSTRAINT [DF_PortfoliosQuantities_ClientPL] DEFAULT (0),
    [ClientPL_] float NOT NULL CONSTRAINT [DF_PortfoliosQuantities_ClientPL_] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[PriceImportBatches] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [QuoteDate] datetime2(7) NOT NULL,
    [Exchange] nvarchar(16) NOT NULL,
    [SourceFile] nvarchar(260) NULL,
    [SourceFileHash] nvarchar(64) NULL,
    [Status] nvarchar(20) NOT NULL CONSTRAINT [DF_PriceImportBatches_Status] DEFAULT ('Pending'),
    [RowCount] int NOT NULL CONSTRAINT [DF_PriceImportBatches_RowCount] DEFAULT ((0)),
    [MatchedCount] int NOT NULL CONSTRAINT [DF_PriceImportBatches_MatchedCount] DEFAULT ((0)),
    [UnmatchedCount] int NOT NULL CONSTRAINT [DF_PriceImportBatches_UnmatchedCount] DEFAULT ((0)),
    [CommittedCount] int NOT NULL CONSTRAINT [DF_PriceImportBatches_CommittedCount] DEFAULT ((0)),
    [CreatedAt] datetime2(7) NOT NULL CONSTRAINT [DF_PriceImportBatches_CreatedAt] DEFAULT (sysutcdatetime()),
    [CreatedBy] int NULL,
    [CommittedAt] datetime2(7) NULL,
    [CommittedBy] int NULL,
    [Notes] nvarchar(500) NULL
);
GO
CREATE TABLE [dbo].[PriceImportRows] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [BatchId] int NOT NULL,
    [SourceSymbol] nvarchar(50) NOT NULL,
    [SourceName] nvarchar(200) NULL,
    [SecurityDpa] int NULL,
    [MatchStatus] nvarchar(20) NOT NULL CONSTRAINT [DF_PriceImportRows_MatchStatus] DEFAULT ('Unmatched'),
    [UnmatchReason] nvarchar(50) NULL,
    [Open] decimal(18,4) NULL,
    [High] decimal(18,4) NULL,
    [Low] decimal(18,4) NULL,
    [Close] decimal(18,4) NULL,
    [PreviousClose] decimal(18,4) NULL,
    [Bid] decimal(18,4) NULL,
    [Offer] decimal(18,4) NULL,
    [Volume] bigint NULL,
    [CreatedAt] datetime2(7) NOT NULL CONSTRAINT [DF_PriceImportRows_CreatedAt] DEFAULT (sysutcdatetime())
);
GO
CREATE TABLE [dbo].[PrimaryIssues] (
    [PrimaryIssues_DPA_] int NOT NULL,
    [Client_DPA_] int NOT NULL,
    [Bond_DPA_] int NOT NULL,
    [ValueDate] datetime NULL,
    [FaceValue] money NULL,
    [Price] money NULL,
    [ForwardRate] real NULL,
    [ApplicationNo] nvarchar(500) NULL,
    [AcceptedAmount] money NULL,
    [PaymentTypes_DPA_] int NOT NULL,
    [PaymentDate] datetime NULL,
    [Reference] nvarchar(500) NULL,
    [PaidAmount] money NULL,
    [BankAcc_DPA_] int NULL,
    [Narrative] nvarchar(500) NULL,
    [Included] int NULL CONSTRAINT [DF_PrimaryIssues_Included] DEFAULT (1),
    [Basis] numeric(18,0) NULL,
    [DaysInCoupon] numeric(18,0) NULL,
    [ModifiedBy] int NULL,
    [DateModified] datetime NULL,
    [Deleted] int NULL
);
GO
CREATE TABLE [dbo].[Printers] (
    [Printer_DPA_] int NOT NULL,
    [PrinterActualName] nvarchar(100) NOT NULL,
    [PrinterName] nvarchar(100) NOT NULL,
    [Description] nvarchar(200) NOT NULL,
    [Active] bit NOT NULL
);
GO
CREATE TABLE [dbo].[Residency] (
    [Residency_DPA_] int NOT NULL,
    [Residency_EIT_] nvarchar(40) NULL,
    [ResidencyDescription] nvarchar(100) NOT NULL,
    [DefaultSelection] bit NOT NULL CONSTRAINT [DF_Residency_DefaultSelection] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[SecTransFee] (
    [SecTransFee_DPA_] int NOT NULL,
    [SecTransFee_EIT_] nvarchar(40) NULL,
    [SecTransFeeActive] bit NOT NULL CONSTRAINT [DF_SecTransFee_SecTransFeeActive] DEFAULT (0),
    [SecTransFeeADate] datetime NOT NULL,
    [SecTransFeeFee] money NOT NULL,
    [Security_DPA_] int NOT NULL
);
GO
CREATE TABLE [dbo].[Security] (
    [Security_DPA_] int NOT NULL,
    [Security_EIT_] nvarchar(40) NULL,
    [SecurityAddr] nvarchar(100) NULL,
    [SecurityCode] nvarchar(50) NOT NULL,
    [SecurityMktPrice] money NOT NULL,
    [SecurityName] nvarchar(100) NOT NULL,
    [OrderSecType_DPA_] int NOT NULL,
    [Immobilised] bit NOT NULL,
    [Sector_DPA_] int NOT NULL CONSTRAINT [DF_Security_Sector_DPA_] DEFAULT (0),
    [ImportCode] nvarchar(50) NULL,
    [CanTrade] bit NULL CONSTRAINT [DF_Security_CanTrade] DEFAULT (1),
    [IsOffering] bit NULL CONSTRAINT [DF_Security_IsOffering] DEFAULT (0),
    [NSEName] varchar(100) NULL,
    [ExpiryDate] datetime NULL,
    [Offerings] int NOT NULL CONSTRAINT [DF_Security_Offerings] DEFAULT (0),
    [BankAccount_DPA_] int NULL,
    [ClosingDate] datetime NULL,
    [OfferType_DPA_] int NULL,
    [TimeModified] datetime NULL CONSTRAINT [DF_Security_TimeModified] DEFAULT (getdate()),
    [ModifiedBy] int NULL,
    [ParentSecurity_DPA_] int NULL,
    [BatchSize] int NULL,
    [RequiresExtra] int NULL,
    [DefaultSelection] int NULL,
    [Ratio] money NULL,
    [MinimumQty] int NULL,
    [StepQty] int NULL,
    [RequiresHoldings] int NULL
);
GO
CREATE TABLE [dbo].[security2] (
    [Security_DPA_] int NOT NULL,
    [Security_EIT_] nvarchar(40) NULL,
    [SecurityAddr] nvarchar(100) NULL,
    [SecurityCode] nvarchar(50) NOT NULL,
    [SecurityMktPrice] money NOT NULL,
    [SecurityName] nvarchar(100) NOT NULL,
    [OrderSecType_DPA_] int NOT NULL,
    [Immobilised] bit NOT NULL,
    [Sector_DPA_] int NOT NULL,
    [ImportCode] nvarchar(50) NULL,
    [CanTrade] bit NULL,
    [IsOffering] bit NULL,
    [NSEName] varchar(100) NULL,
    [ExpiryDate] datetime NULL,
    [Offerings] int NOT NULL,
    [BankAccount_DPA_] int NULL,
    [ClosingDate] datetime NULL,
    [OfferType_DPA_] int NULL,
    [TimeModified] datetime NULL,
    [ModifiedBy] int NULL,
    [ParentSecurity_DPA_] int NULL,
    [BatchSize] int NULL,
    [RequiresExtra] int NULL,
    [DefaultSelection] int NULL,
    [Ratio] money NULL,
    [MinimumQty] int NULL,
    [StepQty] int NULL,
    [RequiresHoldings] int NULL
);
GO
CREATE TABLE [dbo].[SecurityOffering] (
    [Security_DPA_] int NOT NULL,
    [Security_EIT_] nvarchar(40) NULL,
    [SecurityAddr] nvarchar(100) NULL,
    [SecurityCode] nvarchar(10) NOT NULL,
    [SecurityMktPrice] money NOT NULL,
    [SecurityName] nvarchar(100) NOT NULL,
    [OrderSecType_DPA_] int NOT NULL,
    [Immobilised] bit NOT NULL,
    [CanTrade] bit NOT NULL,
    [Deleted] bit NULL CONSTRAINT [DF_SecurityOffering_Deleted] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[Share] (
    [Security_DPA_] int NOT NULL,
    [Share_DPA_] int NOT NULL,
    [Share_EIT_] nvarchar(40) NULL,
    [ShareAnnouncement] smalldatetime NOT NULL,
    [ShareClosing] smalldatetime NOT NULL,
    [SharePDate] smalldatetime NOT NULL,
    [ShareYEnd] smalldatetime NOT NULL,
    [ShareAnnouncementType_DPA_] int NULL
);
GO
CREATE TABLE [dbo].[ShareAnnouncementType] (
    [ShareAnnouncementType_DPA_] int IDENTITY(1,1) NOT NULL,
    [ShareAnnouncementTypeName] nvarchar(50) NOT NULL
);
GO
CREATE TABLE [dbo].[Status] (
    [Status_DPA_] int NOT NULL,
    [Status_EIT_] nvarchar(40) NULL,
    [StatusDescription] nvarchar(10) NOT NULL
);
GO
CREATE TABLE [dbo].[StockWatch] (
    [StockWatch_DPA_] int NOT NULL,
    [StockWatch_EIT_] nvarchar(50) NULL,
    [Client_DPA_] int NOT NULL,
    [Security_DPA_] int NOT NULL,
    [Changedby] int NOT NULL,
    [TimeChanged] datetime NOT NULL CONSTRAINT [DF_StockWatch_TimeChanged] DEFAULT (getdate()),
    [Deleted] bit NOT NULL CONSTRAINT [DF_StockWatch_Deleted] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[SystemNotification] (
    [SysConfig_DPA_] int IDENTITY(1,1) NOT NULL,
    [Entity_DPA_] int NULL,
    [Description] nvarchar(500) NULL,
    [Title] nvarchar(100) NULL
);
GO
CREATE TABLE [dbo].[tblApplDetails] (
    [PayRef] nvarchar(50) NULL,
    [AgentID] nvarchar(6) NULL,
    [PayRecID] int NOT NULL,
    [ScdID] int NULL,
    [ItemNo] smallint NULL,
    [AppType] nvarchar(1) NOT NULL,
    [NoShares] money NULL,
    [AmtPayed] money NULL,
    [ChqNo] nvarchar(50) NULL,
    [BankCode] nvarchar(60) NOT NULL,
    [DrwAcc] nvarchar(60) NULL,
    [shNAME1] nvarchar(100) NOT NULL,
    [shNAME2] nvarchar(35) NULL,
    [CapturedBy] nvarchar(10) NULL,
    [CDate] smalldatetime NOT NULL CONSTRAINT [DF_tblApplDetails_CDate] DEFAULT (getdate()),
    [EditedBy] nvarchar(50) NULL,
    [EditDate] nvarchar(50) NULL,
    [Banked] bit NOT NULL CONSTRAINT [DF_tblApplDetails_Banked] DEFAULT (1),
    [DateBanked] smalldatetime NULL,
    [Balanced] bit NOT NULL CONSTRAINT [DF_tblApplDetails_Balanced] DEFAULT (1),
    [Diff] int NOT NULL CONSTRAINT [DF_tblApplDetails_Diff] DEFAULT (0),
    [balRef] nvarchar(60) NULL,
    [AgentCode] nvarchar(6) NULL,
    [OfferingDPA] int NULL
);
GO
CREATE TABLE [dbo].[tblCompleteEntityList] (
    [EntityName] nvarchar(201) NULL,
    [EntityType] nvarchar(100) NOT NULL,
    [Entity_DPA_] int NOT NULL,
    [EntityType_DPA_] int NOT NULL,
    [EntityCode] nvarchar(4000) NULL,
    [EntityNameEx] nvarchar(4000) NULL
);
GO
CREATE TABLE [dbo].[tblstatementlist] (
    [Client_DPA_] int NOT NULL,
    [TransDate] datetime NULL,
    [REF] nvarchar(500) NULL,
    [Particulars] nvarchar(500) NULL,
    [Debit] float NULL,
    [Credit] money NULL,
    [CreditBal] money NULL,
    [IsOpeningBalance] int NOT NULL,
    [PaymentReceiptNo] nvarchar(50) NULL,
    [ReceiptType] int NOT NULL,
    [Balance] float NULL
);
GO
CREATE TABLE [dbo].[tblTurnOver] (
    [Turnover_DPA_] bigint IDENTITY(1,1) NOT NULL,
    [TradeDate] datetime NOT NULL,
    [MarketTurnOver] decimal(23,2) NOT NULL,
    [CreatedBy] int NOT NULL,
    [TimeCreated] datetime NOT NULL CONSTRAINT [DF_tblTurnOver_TimeCreated] DEFAULT (getdate()),
    [ChangedBy] int NULL,
    [TimeChanged] datetime NULL,
    [Deleted] tinyint NOT NULL CONSTRAINT [DF_tblTurnOver_Deleted] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[tbOrder] (
    [Branch_DPA_] int NOT NULL,
    [Client_DPA_] int NOT NULL,
    [Order_DPA_] int NOT NULL,
    [Order_EIT_] nvarchar(40) NULL,
    [OrderDate] smalldatetime NOT NULL,
    [OrderHold] bit NOT NULL CONSTRAINT [DF_tbOrder_OrderHold] DEFAULT (0),
    [OrderRef] nvarchar(100) NOT NULL,
    [OrderSecType_DPA_] int NOT NULL,
    [OrderType_DPA_] int NOT NULL,
    [OrderCanceled] bit NOT NULL CONSTRAINT [DF_tbOrder_OrderCanceled] DEFAULT (0),
    [OrderHoldType_DPA_] int NOT NULL CONSTRAINT [DF_tbOrder_OrderHoldType_DPA_] DEFAULT (2),
    [OrderAutoReleaseDate] smalldatetime NULL,
    [OrderCompounded] bit NOT NULL CONSTRAINT [DF_tbOrder_OrdDetailCompound] DEFAULT (0),
    [OrderReleasedBy] int NULL,
    [OrderDateReleased] datetime NULL,
    [Deleted] bit NULL CONSTRAINT [DF_tbOrder_Deleted] DEFAULT (0),
    [InterBank] bit NULL CONSTRAINT [DF_tbOrder_InterBank] DEFAULT (0),
    [ChangedBy] int NULL,
    [TimeChanged] datetime NULL CONSTRAINT [DF_tbOrder_TimeChanged] DEFAULT (getdate()),
    [Proposal_DPA_] int NULL,
    [IsCustodian] bit NULL CONSTRAINT [DF_tbOrder_IsCustodian] DEFAULT (0),
    [Agent_DPA_] int NULL,
    [PartialAmount] float NULL CONSTRAINT [DF_tbOrder_PartialAmount] DEFAULT (0),
    [PayOption] int NULL CONSTRAINT [DF_tbOrder_PayOption] DEFAULT (0),
    [ToPay] tinyint NULL CONSTRAINT [DF_tbOrder_ToPay] DEFAULT (0),
    [PT] tinyint NULL CONSTRAINT [DF_tbOrder_PT] DEFAULT (0),
    [Remarks] varchar(200) NULL,
    [AgentReturnable] tinyint NULL,
    [TimeCreated] datetime NULL CONSTRAINT [DF_tbOrder_TimeCreated] DEFAULT (getdate()),
    [CreatedBy] int NULL
);
GO
CREATE TABLE [dbo].[temp_MarchContractSchedule] (
    [Contract_DPA_] nvarchar(255) NULL,
    [Agent] money NULL,
    [Basic] money NULL,
    [Commission] money NULL,
    [MSEComm] money NULL,
    [VAT] money NULL,
    [Total] money NULL,
    [AgentName] nvarchar(100) NULL,
    [IsCustodian] bit NULL,
    [ContractSettlementDate] datetime NOT NULL,
    [LotTDate] smalldatetime NULL,
    [Contract] int NOT NULL,
    [ContractNumber] nvarchar(50) NOT NULL,
    [OrderTypeSale] tinyint NOT NULL,
    [SecurityCode] nvarchar(50) NULL,
    [OrderSecType_DPA_] int NOT NULL,
    [BrokerCode] nvarchar(5) NOT NULL,
    [LotSlipNo] nvarchar(20) NOT NULL,
    [LotPrice] money NOT NULL,
    [LotQty] int NOT NULL,
    [ClientName] nvarchar(100) NOT NULL,
    [Client_DPA_] int NOT NULL,
    [LotGrossAmount] money NULL
);
GO
CREATE TABLE [dbo].[TimeLimit] (
    [TimeLimit_DPA_] int NOT NULL,
    [TimeLimit_EIT_] nvarchar(40) NULL,
    [TimeLimitAction] nvarchar(100) NOT NULL,
    [TimeLimitLimDaysInt] int NOT NULL,
    [TimeLimitLimDaysNSE] int NOT NULL
);
GO
CREATE TABLE [dbo].[UserGroups] (
    [MemberID] int IDENTITY(1,1) NOT NULL,
    [GroupID] int NULL,
    [UserID] int NULL
);
GO
CREATE TABLE [dbo].[UserPageAccess] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [PortalUserId] int NOT NULL,
    [PageKey] nvarchar(64) NOT NULL,
    [CanView] bit NOT NULL CONSTRAINT [DF_UserPageAccess_View] DEFAULT ((0)),
    [CanCreate] bit NOT NULL CONSTRAINT [DF_UserPageAccess_Create] DEFAULT ((0)),
    [CanEdit] bit NOT NULL CONSTRAINT [DF_UserPageAccess_Edit] DEFAULT ((0)),
    [CanDelete] bit NOT NULL CONSTRAINT [DF_UserPageAccess_Delete] DEFAULT ((0)),
    [CanApprove] bit NOT NULL CONSTRAINT [DF_UserPageAccess_Approve] DEFAULT ((0))
);
GO
CREATE TABLE [dbo].[Users] (
    [UserID] int IDENTITY(1,1) NOT NULL,
    [UserName] nvarchar(50) NULL,
    [Password] nvarchar(255) NULL,
    [Surname] nvarchar(50) NULL,
    [OtherNames] nvarchar(50) NULL,
    [StaffID] nvarchar(50) NULL,
    [Removed] int NULL,
    [FirstTime] int NULL,
    [Description] nvarchar(200) NULL,
    [Expires] nvarchar(50) NULL,
    [Enabled] int NULL,
    [RemoteUser] int NOT NULL,
    [Client_DPA_] int NULL,
    [SecretQuestion] nvarchar(100) NULL,
    [SecretAnswer] nvarchar(100) NULL,
    [RequiresSecretQuestion] int NOT NULL CONSTRAINT [DF_Users_RequiresSecretQuestion] DEFAULT (0),
    [email] nvarchar(50) NULL,
    [Accepted] bit NOT NULL CONSTRAINT [DF_Users_Accepted] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[Voucher] (
    [Voucher_DPA_] int NOT NULL,
    [VoucherDate] datetime NULL,
    [Voucher_EIT_] nvarchar(40) NULL,
    [VoucherPaid] bit NOT NULL CONSTRAINT [DF_Voucher_VoucherPaid] DEFAULT (0),
    [Deleted] bit NOT NULL CONSTRAINT [DF_Voucher_Deleted] DEFAULT (0)
);
GO
CREATE TABLE [dbo].[WebtbOrder] (
    [Client_DPA_] int NOT NULL,
    [Order_DPA_] int NOT NULL,
    [Order_EIT_] nvarchar(40) NULL,
    [OrderDate] smalldatetime NOT NULL,
    [OrderRef] nvarchar(100) NOT NULL,
    [OrderSecType_DPA_] int NOT NULL,
    [OrderType_DPA_] int NOT NULL,
    [OrderCanceled] bit NOT NULL CONSTRAINT [DF_WebtbOrder_OrderCanceled] DEFAULT (0),
    [OrderCompounded] bit NOT NULL CONSTRAINT [DF_WebtbOrder_OrderCompounded] DEFAULT (0),
    [OrderReleasedBy] int NULL,
    [OrderDateReleased] datetime NULL,
    [OrdDetailPrice] nvarchar(20) NULL,
    [OrdDetailQty] int NULL,
    [OrdDetailValidity] smalldatetime NULL,
    [Security_DPA_] int NULL,
    [Accepted] bit NOT NULL CONSTRAINT [DF_WebtbOrder_Accepted] DEFAULT (0),
    [ApprovalAction] nvarchar(100) NULL,
    [Reason] nvarchar(100) NULL,
    [UserName] nvarchar(100) NULL,
    [ActionDate] smalldatetime NULL,
    [ActionTime] smalldatetime NULL,
    [Rejected] bit NOT NULL CONSTRAINT [DF_WebtbOrder_Rejected] DEFAULT (0),
    [Action] int NOT NULL CONSTRAINT [DF_WebtbOrder_Action] DEFAULT (3)
);
GO
-- ===== PRIMARY KEYS =====
ALTER TABLE [dbo].[_CDS_Imported_Files_] ADD CONSTRAINT [PK__CDS_Imported_Files_] PRIMARY KEY CLUSTERED ([CDSFile_DPA_]);
GO
ALTER TABLE [dbo].[_CDS_Imported_Trades_] ADD CONSTRAINT [PK__CDS_Imported_Trades_] PRIMARY KEY CLUSTERED ([CDSImport_DPA_]);
GO
ALTER TABLE [dbo].[_Initial_Table_ID_] ADD CONSTRAINT [PK__Initial_Table_ID_] PRIMARY KEY CLUSTERED ([InitialTableID_DPA_]);
GO
ALTER TABLE [dbo].[_Parent_Child_Links_] ADD CONSTRAINT [PK__Parent_Child_Links_] PRIMARY KEY CLUSTERED ([Link_DPA_]);
GO
ALTER TABLE [dbo].[Account] ADD CONSTRAINT [PK_Account] PRIMARY KEY CLUSTERED ([Account_DPA_]);
GO
ALTER TABLE [dbo].[AccountType] ADD CONSTRAINT [PK_AccountType] PRIMARY KEY CLUSTERED ([AccountType_DPA_]);
GO
ALTER TABLE [dbo].[Activity] ADD CONSTRAINT [PK_Activity] PRIMARY KEY CLUSTERED ([Activity_DPA_]);
GO
ALTER TABLE [dbo].[ActvtyClass] ADD CONSTRAINT [PK_ActvtyClass] PRIMARY KEY CLUSTERED ([ActvtyClass_DPA_]);
GO
ALTER TABLE [dbo].[Agent] ADD CONSTRAINT [PK_Agent] PRIMARY KEY CLUSTERED ([Agent_DPA_]);
GO
ALTER TABLE [dbo].[APortfolio] ADD CONSTRAINT [PK_APortfolio] PRIMARY KEY CLUSTERED ([APortfolio_DPA_]);
GO
ALTER TABLE [dbo].[AuditTrail] ADD CONSTRAINT [PK_AuditTrail] PRIMARY KEY CLUSTERED ([AuditTrail_DPA_]);
GO
ALTER TABLE [dbo].[AuditTrailItem] ADD CONSTRAINT [PK_AuditTrailItem] PRIMARY KEY CLUSTERED ([AuditTrailItem_DPA_]);
GO
ALTER TABLE [dbo].[Bank] ADD CONSTRAINT [PK_Bank] PRIMARY KEY CLUSTERED ([Bank_DPA_]);
GO
ALTER TABLE [dbo].[BankAcc] ADD CONSTRAINT [PK_BankAcc] PRIMARY KEY NONCLUSTERED ([BankAcc_DPA_]);
GO
ALTER TABLE [dbo].[BnkBranch] ADD CONSTRAINT [PK_BnkBranch] PRIMARY KEY CLUSTERED ([BnkBranch_DPA_]);
GO
ALTER TABLE [dbo].[Branch] ADD CONSTRAINT [PK_Branch] PRIMARY KEY CLUSTERED ([Branch_DPA_]);
GO
ALTER TABLE [dbo].[Broker] ADD CONSTRAINT [PK_Broker] PRIMARY KEY CLUSTERED ([Broker_DPA_]);
GO
ALTER TABLE [dbo].[CdsImportedHoldings] ADD CONSTRAINT [PK_CdsImportedHoldings] PRIMARY KEY CLUSTERED ([Id]);
GO
ALTER TABLE [dbo].[CdsImportedTrades] ADD CONSTRAINT [PK_CdsImportedTrades] PRIMARY KEY CLUSTERED ([Id]);
GO
ALTER TABLE [dbo].[Class] ADD CONSTRAINT [PK_Class] PRIMARY KEY CLUSTERED ([Class_DPA_]);
GO
ALTER TABLE [dbo].[Client] ADD CONSTRAINT [PK_Client] PRIMARY KEY CLUSTERED ([Client_DPA_]);
GO
ALTER TABLE [dbo].[clientBalancesTemp] ADD CONSTRAINT [PK_clientBalancesTemp] PRIMARY KEY CLUSTERED ([client_DPA_]);
GO
ALTER TABLE [dbo].[Commission] ADD CONSTRAINT [PK_Commission] PRIMARY KEY CLUSTERED ([Commission_DPA_]);
GO
ALTER TABLE [dbo].[Contract] ADD CONSTRAINT [PK_Contract] PRIMARY KEY CLUSTERED ([Contract_DPA_]);
GO
ALTER TABLE [dbo].[ContractApprovals] ADD CONSTRAINT [PK_ContractApprovals] PRIMARY KEY CLUSTERED ([Id]);
GO
ALTER TABLE [dbo].[CPortfolio] ADD CONSTRAINT [PK_CPortfolio] PRIMARY KEY CLUSTERED ([CPortfolio_DPA_]);
GO
ALTER TABLE [dbo].[datastream_Securities] ADD CONSTRAINT [PK_datastream_Securities] PRIMARY KEY CLUSTERED ([SecCode]);
GO
ALTER TABLE [dbo].[down_File] ADD CONSTRAINT [PK_down_File] PRIMARY KEY CLUSTERED ([down_File_DPA_]);
GO
ALTER TABLE [dbo].[dtproperties] ADD CONSTRAINT [pk_dtproperties] PRIMARY KEY CLUSTERED ([id], [property]);
GO
ALTER TABLE [dbo].[EmailDocs] ADD CONSTRAINT [PK_EmailDocs] PRIMARY KEY CLUSTERED ([Document_DPA_]);
GO
ALTER TABLE [dbo].[EmailMode] ADD CONSTRAINT [PK_EmailMode] PRIMARY KEY CLUSTERED ([ModeId]);
GO
ALTER TABLE [dbo].[Entity] ADD CONSTRAINT [PK_Entity] PRIMARY KEY CLUSTERED ([Entity_DPA_]);
GO
ALTER TABLE [dbo].[EntityType] ADD CONSTRAINT [PK_EntityType] PRIMARY KEY CLUSTERED ([EntityType_DPA_]);
GO
ALTER TABLE [dbo].[excep_SummaryHoldings] ADD CONSTRAINT [PK_excep_SummaryHoldings] PRIMARY KEY CLUSTERED ([Trans_DPA_]);
GO
ALTER TABLE [dbo].[Gender] ADD CONSTRAINT [PK_Gender] PRIMARY KEY CLUSTERED ([Gender_DPA_]);
GO
ALTER TABLE [dbo].[Generic] ADD CONSTRAINT [PK_Generic] PRIMARY KEY CLUSTERED ([Generic_DPA_]);
GO
ALTER TABLE [dbo].[GenericSetting] ADD CONSTRAINT [PK_GenericSetting] PRIMARY KEY CLUSTERED ([GenericSetting_DPA_]);
GO
ALTER TABLE [dbo].[Holidays] ADD CONSTRAINT [PK_Holidays] PRIMARY KEY CLUSTERED ([Holiday_DPA_]);
GO
ALTER TABLE [dbo].[InterTransfer] ADD CONSTRAINT [PK_InterTransfer] PRIMARY KEY CLUSTERED ([InterTransfer_DPA_]);
GO
ALTER TABLE [dbo].[Journal] ADD CONSTRAINT [PK_Journal] PRIMARY KEY CLUSTERED ([Journal_DPA_]);
GO
ALTER TABLE [dbo].[JournalEntry] ADD CONSTRAINT [PK_JournalEntry] PRIMARY KEY CLUSTERED ([JournalEntry_DPA_]);
GO
ALTER TABLE [dbo].[Levy] ADD CONSTRAINT [PK_Levy] PRIMARY KEY CLUSTERED ([Levy_DPA_]);
GO
ALTER TABLE [dbo].[LevyContract] ADD CONSTRAINT [PK_LevyContract] PRIMARY KEY CLUSTERED ([LevyContract_DPA_]);
GO
ALTER TABLE [dbo].[LevyReportOrder] ADD CONSTRAINT [PK_LevyReportOrder] PRIMARY KEY CLUSTERED ([LevyOrder_DPA_]);
GO
ALTER TABLE [dbo].[LevySecurity] ADD CONSTRAINT [PK_LevySecurity] PRIMARY KEY CLUSTERED ([LevySecurity_DPA_]);
GO
ALTER TABLE [dbo].[Lot] ADD CONSTRAINT [PK_Lot] PRIMARY KEY CLUSTERED ([Lot_DPA_]);
GO
ALTER TABLE [dbo].[Lot_Deleted] ADD CONSTRAINT [PK_Lot_Deleted] PRIMARY KEY CLUSTERED ([Lot_DPA_]);
GO
ALTER TABLE [dbo].[MailConfiguration] ADD CONSTRAINT [PK_MailConfiguration] PRIMARY KEY CLUSTERED ([MailConfig_DPA_]);
GO
ALTER TABLE [dbo].[MarketQuotes] ADD CONSTRAINT [PK_MarketQuotes] PRIMARY KEY CLUSTERED ([Id]);
GO
ALTER TABLE [dbo].[Offerings] ADD CONSTRAINT [PK_Offerings] PRIMARY KEY CLUSTERED ([Offering_DPA_]);
GO
ALTER TABLE [dbo].[OrdDetail] ADD CONSTRAINT [PK_OrdDetail] PRIMARY KEY CLUSTERED ([OrdDetail_DPA_]);
GO
ALTER TABLE [dbo].[OrderHoldType] ADD CONSTRAINT [PK_OrderHoldType] PRIMARY KEY CLUSTERED ([OrderHoldType_DPA_]);
GO
ALTER TABLE [dbo].[OrderSecType] ADD CONSTRAINT [PK_OrderSecType] PRIMARY KEY CLUSTERED ([OrderSecType_DPA_]);
GO
ALTER TABLE [dbo].[OrderType] ADD CONSTRAINT [PK_OrderType] PRIMARY KEY CLUSTERED ([OrderType_DPA_]);
GO
ALTER TABLE [dbo].[Owner] ADD CONSTRAINT [PK_Owner] PRIMARY KEY CLUSTERED ([Owner_DPA_]);
GO
ALTER TABLE [dbo].[Payment] ADD CONSTRAINT [PK_Payment] PRIMARY KEY NONCLUSTERED ([Payment_DPA_]);
GO
ALTER TABLE [dbo].[PaymentApprovals] ADD CONSTRAINT [PK_PaymentApprovals] PRIMARY KEY CLUSTERED ([Id]);
GO
ALTER TABLE [dbo].[PayType] ADD CONSTRAINT [PK_PayType] PRIMARY KEY CLUSTERED ([PayType_DPA_]);
GO
ALTER TABLE [dbo].[PortalPaymentRequests] ADD CONSTRAINT [PK_PortalPaymentRequests] PRIMARY KEY CLUSTERED ([Id]);
GO
ALTER TABLE [dbo].[PortalRefreshTokens] ADD CONSTRAINT [PK_PortalRefreshTokens] PRIMARY KEY CLUSTERED ([Id]);
GO
ALTER TABLE [dbo].[PortalUsers] ADD CONSTRAINT [PK_PortalUsers] PRIMARY KEY CLUSTERED ([Id]);
GO
ALTER TABLE [dbo].[PortfoliosQuantities] ADD CONSTRAINT [PK_PortfoliosQuantities] PRIMARY KEY CLUSTERED ([Portfolio_DPA_]);
GO
ALTER TABLE [dbo].[PriceImportBatches] ADD CONSTRAINT [PK_PriceImportBatches] PRIMARY KEY CLUSTERED ([Id]);
GO
ALTER TABLE [dbo].[PriceImportRows] ADD CONSTRAINT [PK_PriceImportRows] PRIMARY KEY CLUSTERED ([Id]);
GO
ALTER TABLE [dbo].[Residency] ADD CONSTRAINT [PK_Residency] PRIMARY KEY CLUSTERED ([Residency_DPA_]);
GO
ALTER TABLE [dbo].[Security] ADD CONSTRAINT [PK_Security] PRIMARY KEY NONCLUSTERED ([Security_DPA_]);
GO
ALTER TABLE [dbo].[Share] ADD CONSTRAINT [PK_Share] PRIMARY KEY CLUSTERED ([Share_DPA_]);
GO
ALTER TABLE [dbo].[ShareAnnouncementType] ADD CONSTRAINT [PK_ShareAnnouncementType] PRIMARY KEY CLUSTERED ([ShareAnnouncementType_DPA_]);
GO
ALTER TABLE [dbo].[Status] ADD CONSTRAINT [PK_Status] PRIMARY KEY CLUSTERED ([Status_DPA_]);
GO
ALTER TABLE [dbo].[tblTurnOver] ADD CONSTRAINT [PK_tblTurnOver] PRIMARY KEY CLUSTERED ([Turnover_DPA_]);
GO
ALTER TABLE [dbo].[tbOrder] ADD CONSTRAINT [PK_tbOrder] PRIMARY KEY CLUSTERED ([Order_DPA_]);
GO
ALTER TABLE [dbo].[TimeLimit] ADD CONSTRAINT [PK_TimeLimit] PRIMARY KEY CLUSTERED ([TimeLimit_DPA_]);
GO
ALTER TABLE [dbo].[UserPageAccess] ADD CONSTRAINT [PK_UserPageAccess] PRIMARY KEY CLUSTERED ([Id]);
GO
ALTER TABLE [dbo].[Users] ADD CONSTRAINT [PK_Users] PRIMARY KEY CLUSTERED ([UserID]);
GO
ALTER TABLE [dbo].[Voucher] ADD CONSTRAINT [PK_Voucher] PRIMARY KEY CLUSTERED ([Voucher_DPA_]);
GO
ALTER TABLE [dbo].[WebtbOrder] ADD CONSTRAINT [PK_WebtbOrder] PRIMARY KEY CLUSTERED ([Order_DPA_]);
GO
-- ===== FOREIGN KEYS =====
ALTER TABLE [dbo].[Client] ADD CONSTRAINT [FK_Client_Agent] FOREIGN KEY ([Agent_DPA_]) REFERENCES [dbo].[Agent] ([Agent_DPA_]);
GO
ALTER TABLE [dbo].[Client] ADD CONSTRAINT [FK_Client_Branch] FOREIGN KEY ([Branch_DPA_]) REFERENCES [dbo].[Branch] ([Branch_DPA_]);
GO
ALTER TABLE [dbo].[Client] ADD CONSTRAINT [FK_Client_Class] FOREIGN KEY ([Class_DPA_]) REFERENCES [dbo].[Class] ([Class_DPA_]);
GO
ALTER TABLE [dbo].[Client] ADD CONSTRAINT [FK_Client_Commission] FOREIGN KEY ([Commission_DPA_]) REFERENCES [dbo].[Commission] ([Commission_DPA_]);
GO
ALTER TABLE [dbo].[Client] ADD CONSTRAINT [FK_Client_EntityType] FOREIGN KEY ([EntityType_DPA_]) REFERENCES [dbo].[EntityType] ([EntityType_DPA_]);
GO
ALTER TABLE [dbo].[Client] ADD CONSTRAINT [FK_Client_Gender] FOREIGN KEY ([Gender_DPA_]) REFERENCES [dbo].[Gender] ([Gender_DPA_]);
GO
ALTER TABLE [dbo].[Client] ADD CONSTRAINT [FK_Client_Residency] FOREIGN KEY ([Residency_DPA_]) REFERENCES [dbo].[Residency] ([Residency_DPA_]);
GO
ALTER TABLE [dbo].[Contract] ADD CONSTRAINT [FK_Contract_Status] FOREIGN KEY ([Status_DPA_]) REFERENCES [dbo].[Status] ([Status_DPA_]);
GO
ALTER TABLE [dbo].[Lot] ADD CONSTRAINT [FK_Lot_Broker] FOREIGN KEY ([Broker_DPA_]) REFERENCES [dbo].[Broker] ([Broker_DPA_]);
GO
ALTER TABLE [dbo].[Lot] ADD CONSTRAINT [FK_Lot_Contract] FOREIGN KEY ([Contract_DPA_]) REFERENCES [dbo].[Contract] ([Contract_DPA_]);
GO
ALTER TABLE [dbo].[Lot] ADD CONSTRAINT [FK_Lot_OrdDetail] FOREIGN KEY ([OrdDetail_DPA_]) REFERENCES [dbo].[OrdDetail] ([OrdDetail_DPA_]);
GO
ALTER TABLE [dbo].[OrdDetail] ADD CONSTRAINT [FK_OrdDetail_Security] FOREIGN KEY ([Security_DPA_]) REFERENCES [dbo].[Security] ([Security_DPA_]);
GO
ALTER TABLE [dbo].[OrdDetail] ADD CONSTRAINT [FK_OrdDetail_tbOrder] FOREIGN KEY ([Order_DPA_]) REFERENCES [dbo].[tbOrder] ([Order_DPA_]);
GO
ALTER TABLE [dbo].[Payment] ADD CONSTRAINT [FK_Payment_EntityType] FOREIGN KEY ([EntityType_DPA_]) REFERENCES [dbo].[EntityType] ([EntityType_DPA_]);
GO
ALTER TABLE [dbo].[Payment] ADD CONSTRAINT [FK_Payment_PayType] FOREIGN KEY ([PayType_DPA_]) REFERENCES [dbo].[PayType] ([PayType_DPA_]);
GO
ALTER TABLE [dbo].[PortalRefreshTokens] ADD CONSTRAINT [FK_PortalRefreshTokens_PortalUsers] FOREIGN KEY ([PortalUserId]) REFERENCES [dbo].[PortalUsers] ([Id]);
GO
ALTER TABLE [dbo].[PriceImportRows] ADD CONSTRAINT [FK_PriceImportRows_Batch] FOREIGN KEY ([BatchId]) REFERENCES [dbo].[PriceImportBatches] ([Id]);
GO
ALTER TABLE [dbo].[tbOrder] ADD CONSTRAINT [FK_tbOrder_Branch] FOREIGN KEY ([Branch_DPA_]) REFERENCES [dbo].[Branch] ([Branch_DPA_]);
GO
ALTER TABLE [dbo].[tbOrder] ADD CONSTRAINT [FK_tbOrder_Client] FOREIGN KEY ([Client_DPA_]) REFERENCES [dbo].[Client] ([Client_DPA_]);
GO
ALTER TABLE [dbo].[tbOrder] ADD CONSTRAINT [FK_tbOrder_OrderHoldType] FOREIGN KEY ([OrderHoldType_DPA_]) REFERENCES [dbo].[OrderHoldType] ([OrderHoldType_DPA_]);
GO
ALTER TABLE [dbo].[tbOrder] ADD CONSTRAINT [FK_tbOrder_OrderSecType] FOREIGN KEY ([OrderSecType_DPA_]) REFERENCES [dbo].[OrderSecType] ([OrderSecType_DPA_]);
GO
ALTER TABLE [dbo].[tbOrder] ADD CONSTRAINT [FK_tbOrder_OrderType] FOREIGN KEY ([OrderType_DPA_]) REFERENCES [dbo].[OrderType] ([OrderType_DPA_]);
GO
-- ===== INDEXES (non-PK) =====
CREATE NONCLUSTERED INDEX [IX_CdsImportedHoldings_BatchId] ON [dbo].[CdsImportedHoldings] ([BatchId]);
GO
CREATE NONCLUSTERED INDEX [IX_CdsImportedHoldings_MatchStatus] ON [dbo].[CdsImportedHoldings] ([MatchStatus]);
GO
CREATE NONCLUSTERED INDEX [IX_CdsImportedHoldings_SourceFileHash] ON [dbo].[CdsImportedHoldings] ([SourceFileHash]);
GO
CREATE NONCLUSTERED INDEX [IX_CdsImportedHoldings_TradeDate_Symbol] ON [dbo].[CdsImportedHoldings] ([TradeDate], [Symbol]);
GO
CREATE NONCLUSTERED INDEX [IX_CdsImportedTrades_BatchId] ON [dbo].[CdsImportedTrades] ([BatchId]);
GO
CREATE NONCLUSTERED INDEX [IX_CdsImportedTrades_MatchStatus] ON [dbo].[CdsImportedTrades] ([MatchStatus]);
GO
CREATE NONCLUSTERED INDEX [IX_CdsImportedTrades_SourceFileHash] ON [dbo].[CdsImportedTrades] ([SourceFileHash]);
GO
CREATE NONCLUSTERED INDEX [IX_CdsImportedTrades_TradeDate_Symbol] ON [dbo].[CdsImportedTrades] ([TradeDate], [Symbol]);
GO
CREATE NONCLUSTERED INDEX [IX_Client_Agent] ON [dbo].[Client] ([Agent_DPA_]);
GO
CREATE UNIQUE NONCLUSTERED INDEX [IX_ClientBalances] ON [dbo].[ClientBalances] ([client_DPA_]);
GO
CREATE NONCLUSTERED INDEX [IX_Contract_Status] ON [dbo].[Contract] ([Status_DPA_]);
GO
CREATE NONCLUSTERED INDEX [IX_ContractApprovals_ContractDpa] ON [dbo].[ContractApprovals] ([ContractDpa]);
GO
CREATE NONCLUSTERED INDEX [IX_ContractApprovals_Status_CreatedAt] ON [dbo].[ContractApprovals] ([Status], [CreatedAt]);
GO
CREATE NONCLUSTERED INDEX [IX_Holdings_Client] ON [dbo].[Holdings] ([Client_DPA_]);
GO
CREATE NONCLUSTERED INDEX [IX_Holdings_Security] ON [dbo].[Holdings] ([Security_DPA_]);
GO
CREATE NONCLUSTERED INDEX [IX_JournalEntry_Entity] ON [dbo].[JournalEntry] ([EntityType_DPA_], [Entity_DPA_]) INCLUDE ([JournalEntryDebit], [JournalEntryCredit]);
GO
CREATE NONCLUSTERED INDEX [IX_JournalEntry_Journal] ON [dbo].[JournalEntry] ([Journal_DPA_]);
GO
CREATE NONCLUSTERED INDEX [IX_LevyContract_Contract] ON [dbo].[LevyContract] ([Contract_DPA_]);
GO
CREATE NONCLUSTERED INDEX [IX_Lot_Broker] ON [dbo].[Lot] ([Broker_DPA_]);
GO
CREATE NONCLUSTERED INDEX [IX_Lot_Contract] ON [dbo].[Lot] ([Contract_DPA_]);
GO
CREATE NONCLUSTERED INDEX [IX_Lot_OrdDetail] ON [dbo].[Lot] ([OrdDetail_DPA_]);
GO
CREATE NONCLUSTERED INDEX [IX_MarketQuotes_QuoteDate] ON [dbo].[MarketQuotes] ([QuoteDate]);
GO
CREATE UNIQUE NONCLUSTERED INDEX [UX_MarketQuotes_Security_QuoteDate] ON [dbo].[MarketQuotes] ([SecurityDpa], [QuoteDate]);
GO
CREATE NONCLUSTERED INDEX [IX_OrdDetail_Order] ON [dbo].[OrdDetail] ([Order_DPA_]);
GO
CREATE NONCLUSTERED INDEX [IX_OrdDetail_Security] ON [dbo].[OrdDetail] ([Security_DPA_]);
GO
CREATE NONCLUSTERED INDEX [IX_Payment_Contract] ON [dbo].[Payment] ([Contract_DPA_]);
GO
CREATE NONCLUSTERED INDEX [IX_Payment_Entity] ON [dbo].[Payment] ([EntityType_DPA_], [Entity_DPA_], [PayType_DPA_]) INCLUDE ([PaymentAmount], [PaymentPDate], [Deleted]);
GO
CREATE NONCLUSTERED INDEX [IX_Payment_Order] ON [dbo].[Payment] ([Order_DPA_]);
GO
CREATE NONCLUSTERED INDEX [IX_PaymentApprovals_Status_CreatedAt] ON [dbo].[PaymentApprovals] ([Status], [CreatedAt]);
GO
CREATE NONCLUSTERED INDEX [IX_PaymentApprovals_TargetPaymentDpa] ON [dbo].[PaymentApprovals] ([TargetPaymentDpa]);
GO
CREATE NONCLUSTERED INDEX [IX_PortalPaymentRequests_ClientDpa] ON [dbo].[PortalPaymentRequests] ([ClientDpa]);
GO
CREATE NONCLUSTERED INDEX [IX_PortalPaymentRequests_PortalUserId] ON [dbo].[PortalPaymentRequests] ([PortalUserId]);
GO
CREATE NONCLUSTERED INDEX [IX_PortalPaymentRequests_Status_CreatedAt] ON [dbo].[PortalPaymentRequests] ([Status], [CreatedAt]);
GO
CREATE UNIQUE NONCLUSTERED INDEX [IX_PortalRefreshTokens_Token] ON [dbo].[PortalRefreshTokens] ([Token]);
GO
CREATE UNIQUE NONCLUSTERED INDEX [IX_PortalUsers_Email] ON [dbo].[PortalUsers] ([Email]) WHERE ([Email] IS NOT NULL AND [Email]<>'');
GO
CREATE UNIQUE NONCLUSTERED INDEX [IX_PortalUsers_Username] ON [dbo].[PortalUsers] ([Username]) WHERE ([Username] IS NOT NULL);
GO
CREATE NONCLUSTERED INDEX [IX_PriceImportBatches_QuoteDate_Exchange] ON [dbo].[PriceImportBatches] ([QuoteDate], [Exchange]);
GO
CREATE NONCLUSTERED INDEX [IX_PriceImportBatches_SourceFileHash] ON [dbo].[PriceImportBatches] ([SourceFileHash]);
GO
CREATE NONCLUSTERED INDEX [IX_PriceImportRows_BatchId] ON [dbo].[PriceImportRows] ([BatchId]);
GO
CREATE NONCLUSTERED INDEX [IX_PriceImportRows_SecurityDpa] ON [dbo].[PriceImportRows] ([SecurityDpa]);
GO
CREATE UNIQUE NONCLUSTERED INDEX [IX_tblTurnOver] ON [dbo].[tblTurnOver] ([TradeDate]);
GO
CREATE NONCLUSTERED INDEX [IX_tbOrder_Agent] ON [dbo].[tbOrder] ([Agent_DPA_]);
GO
CREATE NONCLUSTERED INDEX [IX_tbOrder_Client] ON [dbo].[tbOrder] ([Client_DPA_]);
GO
-- ===== VIEWS / PROCEDURES / FUNCTIONS =====
CREATE VIEW dbo.AA_Inst_Comm_Report
AS
SELECT        Q1.Contract_DPA_, 'Malawi' AS Country, 'African Alliance Securities Malawi' AS Broker, Q1.InstitutionName, Q1.ClientName, Q1.ResidencyDescription,
                         SUM(Q1.LevyAmount) AS BrokerComm, SUM(Q2.LevyAmount) AS AgentComm, SUM(Q1.LevyAmount) - SUM(Q2.LevyAmount) AS NetComm, Q1.LotTDate, Q1.LotQty,
                         Q1.LotPrice, Q1.LotGrossAmount, Q1.SecurityName
FROM            (SELECT        I.InstitutionName, dbo.LevyContract.Contract_DPA_, dbo.Client.ClientName, dbo.Residency.ResidencyDescription, SUM(dbo.LevyContract.LevyAmount)
                                                    AS LevyAmount, dbo.LevyContract.LevyShortName, dbo.Lot.LotTDate, dbo.Lot.LotQty, dbo.Lot.LotPrice, dbo.Lot.LotGrossAmount,
                                                    dbo.Security.SecurityName
                          FROM            dbo.tbOrder INNER JOIN
                                                    dbo.Client ON dbo.Client.Client_DPA_ = dbo.tbOrder.Client_DPA_ INNER JOIN
                                                    dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                                                    dbo.Security ON dbo.OrdDetail.Security_DPA_ = Security.Security_DPA_ INNER JOIN
                                                    dbo.Lot ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                                                    dbo.Contract ON dbo.Contract.Contract_DPA_ = dbo.Lot.Contract_DPA_ INNER JOIN
                                                    dbo.LevyContract ON dbo.LevyContract.Contract_DPA_ = dbo.Contract.Contract_DPA_ INNER JOIN
                                                    dbo.Residency ON dbo.Residency.Residency_DPA_ = dbo.Client.Residency_DPA_ INNER JOIN
                                                    dbo.Institution AS I ON I.Institution_DPA_ = dbo.Client.Institution_DPA_
                          WHERE        (dbo.LevyContract.LevyShortName = 'Commission') AND (dbo.Client.Institution_DPA_ IS NOT NULL) AND (dbo.tbOrder.Deleted = 0)
                          GROUP BY I.InstitutionName, dbo.LevyContract.Contract_DPA_, dbo.Residency.ResidencyDescription, dbo.LevyContract.LevyShortName, dbo.Lot.LotTDate,
                                                    dbo.Lot.LotQty, dbo.Lot.LotPrice, dbo.Lot.LotGrossAmount, dbo.Security.SecurityName, dbo.Client.ClientName) AS Q1 INNER JOIN
                             (SELECT        I.InstitutionName, dbo.LevyContract.Contract_DPA_, dbo.Client.ClientName, dbo.Residency.ResidencyDescription,
                                                         SUM(dbo.LevyContract.LevyAmount) AS LevyAmount, dbo.LevyContract.LevyShortName, dbo.Lot.LotTDate, dbo.Lot.LotQty, dbo.Lot.LotPrice,
                                                         dbo.Lot.LotGrossAmount, dbo.Security.SecurityName
                               FROM            dbo.tbOrder INNER JOIN
                                                         dbo.Client ON dbo.Client.Client_DPA_ = dbo.tbOrder.Client_DPA_ INNER JOIN
                                                         dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                                                         dbo.Security ON dbo.OrdDetail.Security_DPA_ = Security.Security_DPA_ INNER JOIN
                                                         dbo.Lot ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                                                         dbo.Contract ON dbo.Contract.Contract_DPA_ = dbo.Lot.Contract_DPA_ INNER JOIN
                                                         dbo.LevyContract ON dbo.LevyContract.Contract_DPA_ = dbo.Contract.Contract_DPA_ INNER JOIN
                                                         dbo.Residency ON dbo.Residency.Residency_DPA_ = dbo.Client.Residency_DPA_ INNER JOIN
                                                         dbo.Institution AS I ON I.Institution_DPA_ = dbo.Client.Institution_DPA_
                               WHERE        (dbo.LevyContract.LevyShortName = 'Agent') AND (dbo.Client.Institution_DPA_ IS NOT NULL) AND (dbo.tbOrder.Deleted = 0)
                               GROUP BY I.InstitutionName, dbo.LevyContract.Contract_DPA_, dbo.Residency.ResidencyDescription, dbo.LevyContract.LevyShortName, dbo.Lot.LotTDate,
                                                         dbo.Lot.LotQty, dbo.Lot.LotPrice, dbo.Lot.LotGrossAmount, dbo.Security.SecurityName, dbo.Client.ClientName) AS Q2 ON
                         Q1.Contract_DPA_ = Q2.Contract_DPA_
GROUP BY Q1.InstitutionName, Q1.ClientName, Q1.ResidencyDescription, Q1.LotTDate, Q1.LotQty, Q1.LotPrice, Q1.LotGrossAmount, Q1.SecurityName,
                         Q1.Contract_DPA_
UNION
SELECT        Q1.Contract_DPA_, 'Malawi' AS Country, 'African Alliance Securities Malawi' AS Broker, 'Retail' AS Expr1, Q1.ClientName, Q1.ResidencyDescription,
                         SUM(Q1.LevyAmount) AS BrokerComm, SUM(Q2.LevyAmount) AS AgentComm, SUM(Q1.LevyAmount) - SUM(Q2.LevyAmount) AS NetComm, Q1.LotTDate, Q1.LotQty,
                         Q1.LotPrice, Q1.LotGrossAmount, Q1.SecurityName
FROM            (SELECT        dbo.LevyContract.Contract_DPA_, dbo.Client.ClientName, dbo.Residency.ResidencyDescription, SUM(dbo.LevyContract.LevyAmount) AS LevyAmount,
                                                    dbo.LevyContract.LevyShortName, dbo.Lot.LotTDate, dbo.Lot.LotQty, dbo.Lot.LotPrice, dbo.Lot.LotGrossAmount, dbo.Security.SecurityName
                          FROM            dbo.tbOrder INNER JOIN
                                                    dbo.Client ON dbo.Client.Client_DPA_ = dbo.tbOrder.Client_DPA_ INNER JOIN
                                                    dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                                                    dbo.Security ON dbo.OrdDetail.Security_DPA_ = Security.Security_DPA_ INNER JOIN
                                                    dbo.Lot ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                                                    dbo.Contract ON dbo.Contract.Contract_DPA_ = dbo.Lot.Contract_DPA_ INNER JOIN
                                                    dbo.LevyContract ON dbo.LevyContract.Contract_DPA_ = dbo.Contract.Contract_DPA_ INNER JOIN
                                                    dbo.Residency ON dbo.Residency.Residency_DPA_ = dbo.Client.Residency_DPA_
                          WHERE        (dbo.LevyContract.LevyShortName = 'Commission') AND (dbo.Client.Institution_DPA_ IS NULL) AND (dbo.tbOrder.Deleted = 0)
                          GROUP BY dbo.LevyContract.Contract_DPA_, dbo.Residency.ResidencyDescription, dbo.LevyContract.LevyShortName, dbo.Lot.LotTDate, dbo.Lot.LotQty,
                                                    dbo.Lot.LotPrice, dbo.Lot.LotGrossAmount, dbo.Security.SecurityName, dbo.Client.ClientName) AS Q1 INNER JOIN
                             (SELECT        LevyContract_1.Contract_DPA_, Client_1.ClientName, Residency_1.ResidencyDescription, SUM(LevyContract_1.LevyAmount) AS LevyAmount,
                                                         LevyContract_1.LevyShortName, Lot_1.LotTDate, Lot_1.LotQty, Lot_1.LotPrice, Lot_1.LotGrossAmount, Security_1.SecurityName
                               FROM            dbo.tbOrder AS tbOrder_1 INNER JOIN
                                                         dbo.Client AS Client_1 ON Client_1.Client_DPA_ = tbOrder_1.Client_DPA_ INNER JOIN
                                                         dbo.OrdDetail AS OrdDetail_1 ON tbOrder_1.Order_DPA_ = OrdDetail_1.Order_DPA_ INNER JOIN
                                                         dbo.Security AS Security_1 ON OrdDetail_1.Security_DPA_ = Security_1.Security_DPA_ INNER JOIN
                                                         dbo.Lot AS Lot_1 ON Lot_1.OrdDetail_DPA_ = OrdDetail_1.OrdDetail_DPA_ INNER JOIN
                                                         dbo.Contract AS Contract_1 ON Contract_1.Contract_DPA_ = Lot_1.Contract_DPA_ INNER JOIN
                                                         dbo.LevyContract AS LevyContract_1 ON LevyContract_1.Contract_DPA_ = Contract_1.Contract_DPA_ INNER JOIN
                                                         dbo.Residency AS Residency_1 ON Residency_1.Residency_DPA_ = Client_1.Residency_DPA_
                               WHERE        (LevyContract_1.LevyShortName = 'Agent') AND (Client_1.Institution_DPA_ IS NULL) AND (tbOrder_1.Deleted = 0)
                               GROUP BY LevyContract_1.Contract_DPA_, Residency_1.ResidencyDescription, LevyContract_1.LevyShortName, Lot_1.LotTDate, Lot_1.LotQty, Lot_1.LotPrice,
                                                         Lot_1.LotGrossAmount, Security_1.SecurityName, Client_1.ClientName) AS Q2 ON Q1.Contract_DPA_ = Q2.Contract_DPA_
GROUP BY Q1.ResidencyDescription, Q1.ClientName, Q1.LotTDate, Q1.LotQty, Q1.LotPrice, Q1.LotGrossAmount, Q1.SecurityName, Q1.Contract_DPA_

GO
CREATE VIEW dbo.AA_View_Comm_By_Client
AS
SELECT     'Malawi' AS Country, Q1.InstitutionName, Q1.ResidencyDescription, SUM(Q1.LevyAmount) AS BrokerComm, SUM(Q2.LevyAmount) AS AgentComm,
                      SUM(Q1.LevyAmount) - SUM(Q2.LevyAmount) AS NetComm, Q1.LotTDate
FROM         (SELECT     I.InstitutionName, dbo.LevyContract.Contract_DPA_, dbo.Residency.ResidencyDescription, SUM(dbo.LevyContract.LevyAmount) AS LevyAmount,
                                              dbo.LevyContract.LevyShortName, dbo.Lot.LotTDate
                       FROM          dbo.tbOrder INNER JOIN
                                              dbo.Client ON dbo.Client.Client_DPA_ = dbo.tbOrder.Client_DPA_ INNER JOIN
                                              dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                                              dbo.Lot ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                                              dbo.Contract ON dbo.Contract.Contract_DPA_ = dbo.Lot.Contract_DPA_ INNER JOIN
                                              dbo.LevyContract ON dbo.LevyContract.Contract_DPA_ = dbo.Contract.Contract_DPA_ INNER JOIN
                                              dbo.Residency ON dbo.Residency.Residency_DPA_ = dbo.Client.Residency_DPA_ INNER JOIN
                                              dbo.Institution AS I ON I.Institution_DPA_ = dbo.Client.Institution_DPA_
                       WHERE      (dbo.LevyContract.LevyShortName = 'Commission') AND (dbo.Client.Institution_DPA_ IS NOT NULL)
                       GROUP BY I.InstitutionName, dbo.LevyContract.Contract_DPA_, dbo.Residency.ResidencyDescription, dbo.LevyContract.LevyShortName, dbo.Lot.LotTDate)
                      AS Q1 INNER JOIN
                          (SELECT     I.InstitutionName, dbo.LevyContract.Contract_DPA_, dbo.Residency.ResidencyDescription, SUM(dbo.LevyContract.LevyAmount) AS LevyAmount,
                                                   dbo.LevyContract.LevyShortName, dbo.Lot.LotTDate
                            FROM          dbo.tbOrder INNER JOIN
                                                   dbo.Client ON dbo.Client.Client_DPA_ = dbo.tbOrder.Client_DPA_ INNER JOIN
                                                   dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                                                   dbo.Lot ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                                                   dbo.Contract ON dbo.Contract.Contract_DPA_ = dbo.Lot.Contract_DPA_ INNER JOIN
                                                   dbo.LevyContract ON dbo.LevyContract.Contract_DPA_ = dbo.Contract.Contract_DPA_ INNER JOIN
                                                   dbo.Residency ON dbo.Residency.Residency_DPA_ = dbo.Client.Residency_DPA_ INNER JOIN
                                                   dbo.Institution AS I ON I.Institution_DPA_ = dbo.Client.Institution_DPA_
                            WHERE      (dbo.LevyContract.LevyShortName = 'Agent') AND (dbo.Client.Institution_DPA_ IS NOT NULL)
                            GROUP BY I.InstitutionName, dbo.LevyContract.Contract_DPA_, dbo.Residency.ResidencyDescription, dbo.LevyContract.LevyShortName, dbo.Lot.LotTDate)
                      AS Q2 ON Q1.Contract_DPA_ = Q2.Contract_DPA_
GROUP BY Q1.InstitutionName, Q1.ResidencyDescription, Q1.LotTDate
UNION
SELECT     'Malawi' AS Country, 'Retail' AS Expr1, Q1.ResidencyDescription, SUM(Q1.LevyAmount) AS BrokerComm, SUM(Q2.LevyAmount) AS AgentComm,
                      SUM(Q1.LevyAmount) - SUM(Q2.LevyAmount) AS NetComm, Q1.LotTDate
FROM         (SELECT     dbo.LevyContract.Contract_DPA_, dbo.Residency.ResidencyDescription, SUM(dbo.LevyContract.LevyAmount) AS LevyAmount,
                                              dbo.LevyContract.LevyShortName, dbo.Lot.LotTDate
                       FROM          dbo.tbOrder INNER JOIN
                                              dbo.Client ON dbo.Client.Client_DPA_ = dbo.tbOrder.Client_DPA_ INNER JOIN
                                              dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                                              dbo.Lot ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                                              dbo.Contract ON dbo.Contract.Contract_DPA_ = dbo.Lot.Contract_DPA_ INNER JOIN
                                              dbo.LevyContract ON dbo.LevyContract.Contract_DPA_ = dbo.Contract.Contract_DPA_ INNER JOIN
                                              dbo.Residency ON dbo.Residency.Residency_DPA_ = dbo.Client.Residency_DPA_
                       WHERE      (dbo.LevyContract.LevyShortName = 'Commission') AND (dbo.Client.Institution_DPA_ IS NULL)
                       GROUP BY dbo.LevyContract.Contract_DPA_, dbo.Residency.ResidencyDescription, dbo.LevyContract.LevyShortName, dbo.Lot.LotTDate) AS Q1 INNER JOIN
                          (SELECT     LevyContract_1.Contract_DPA_, Residency_1.ResidencyDescription, SUM(LevyContract_1.LevyAmount) AS LevyAmount,
                                                   LevyContract_1.LevyShortName, Lot_1.LotTDate
                            FROM          dbo.tbOrder AS tbOrder_1 INNER JOIN
                                                   dbo.Client AS Client_1 ON Client_1.Client_DPA_ = tbOrder_1.Client_DPA_ INNER JOIN
                                                   dbo.OrdDetail AS OrdDetail_1 ON tbOrder_1.Order_DPA_ = OrdDetail_1.Order_DPA_ INNER JOIN
                                                   dbo.Lot AS Lot_1 ON Lot_1.OrdDetail_DPA_ = OrdDetail_1.OrdDetail_DPA_ INNER JOIN
                                                   dbo.Contract AS Contract_1 ON Contract_1.Contract_DPA_ = Lot_1.Contract_DPA_ INNER JOIN
                                                   dbo.LevyContract AS LevyContract_1 ON LevyContract_1.Contract_DPA_ = Contract_1.Contract_DPA_ INNER JOIN
                                                   dbo.Residency AS Residency_1 ON Residency_1.Residency_DPA_ = Client_1.Residency_DPA_
                            WHERE      (LevyContract_1.LevyShortName = 'Agent') AND (Client_1.Institution_DPA_ IS NULL)
                            GROUP BY LevyContract_1.Contract_DPA_, Residency_1.ResidencyDescription, LevyContract_1.LevyShortName, Lot_1.LotTDate) AS Q2 ON
                      Q1.Contract_DPA_ = Q2.Contract_DPA_
GROUP BY Q1.ResidencyDescription, Q1.LotTDate

GO

	CREATE VIEW [dbo].[AA_View_Comm_By_Institution]
	AS

	Select
		'Malawi' As Country,
		Q1.InstitutionName,
		Q1.ResidencyDescription,
		SUM(Q1.LevyAmount) BrokerComm,
		SUM(Q2.LevyAmount) AgentComm,
		SUM(Q1.LevyAmount) - SUM(Q2.LevyAmount) AS NetComm,
		Q1.LotTDate
	FROM
		(SELECT
			I.InstitutionName,
			dbo.LevyContract.Contract_DPA_,
			dbo.Residency.ResidencyDescription,
			Sum(dbo.LevyContract.LevyAmount) LevyAmount,
			dbo.LevyContract.LevyShortName,
			dbo.Lot.LotTDate
		FROM
			dbo.tbOrder
		INNER JOIN
			dbo.Client
		ON
			dbo.Client.Client_DPA_ = dbo.tbOrder.Client_DPA_
		INNER JOIN
			dbo.OrdDetail
		ON
			dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_
		INNER JOIN
			dbo.Lot
		ON
			dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_
		INNER JOIN
			dbo.Contract
		ON
			dbo.Contract.Contract_DPA_ = dbo.Lot.Contract_DPA_
		INNER JOIN
			dbo.LevyContract
		ON
			dbo.LevyContract.Contract_DPA_ = dbo.Contract.Contract_DPA_
		INNER JOIN
			dbo.Residency
		ON
			dbo.Residency.Residency_DPA_ = dbo.Client.Residency_DPA_
		Inner Join
			Institution I
		On
			I.Institution_DPA_ = Client.Institution_DPA_
		Where
			LevyShortName = 'Commission'
		And
			dbo.Client.Institution_DPA_ is not null
		Group By
			I.InstitutionName,
			dbo.LevyContract.Contract_DPA_,
			dbo.Residency.ResidencyDescription,
			dbo.LevyContract.LevyShortName,
			dbo.Lot.LotTDate) Q1
		Inner Join
			(SELECT
				I.InstitutionName,
				dbo.LevyContract.Contract_DPA_,
				dbo.Residency.ResidencyDescription,
				Sum(dbo.LevyContract.LevyAmount) LevyAmount,
				dbo.LevyContract.LevyShortName,
				dbo.Lot.LotTDate
			FROM
				dbo.tbOrder
			INNER JOIN
				dbo.Client
			ON
				dbo.Client.Client_DPA_ = dbo.tbOrder.Client_DPA_
			INNER JOIN
				dbo.OrdDetail
			ON
				dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_
			INNER JOIN
				dbo.Lot
			ON
				dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_
			INNER JOIN
				dbo.Contract
			ON
				dbo.Contract.Contract_DPA_ = dbo.Lot.Contract_DPA_
			INNER JOIN
				dbo.LevyContract
			ON
				dbo.LevyContract.Contract_DPA_ = dbo.Contract.Contract_DPA_
			INNER JOIN
				dbo.Residency
			ON
				dbo.Residency.Residency_DPA_ = dbo.Client.Residency_DPA_
			Inner Join
				Institution I
			On
				I.Institution_DPA_ = Client.Institution_DPA_
			Where
				LevyShortName = 'Agent'
			And
				dbo.Client.Institution_DPA_ is not null
			Group By
				I.InstitutionName,
				dbo.LevyContract.Contract_DPA_,
				dbo.Residency.ResidencyDescription,
				dbo.LevyContract.LevyShortName,
				dbo.Lot.LotTDate
			)Q2
		ON
			Q1.Contract_DPA_ = Q2.Contract_DPA_
		Group By
			Q1.InstitutionName,
			Q1.ResidencyDescription,
			Q1.LotTDate
		--Order By
		--	 Q1.Institution


GO
CREATE VIEW dbo.AACommisionReportDaily
AS
SELECT     AACommissionReport_2.TransDate, AACommissionReport_2.Gross, AACommissionReport_2.Commission,
                      AACommissionReport_2.Commission - AACommissionReport_1.Commission AS NetCommission
FROM         dbo.AACommissionReport AS AACommissionReport_2 INNER JOIN
                      dbo.AACommissionReport AS AACommissionReport_1 ON AACommissionReport_2.TransDate = AACommissionReport_1.TransDate
WHERE     (AACommissionReport_1.LevyShortName = N'Agent') AND (AACommissionReport_2.LevyShortName = N'Commission')

GO
CREATE VIEW dbo.AALeviesReport
AS
SELECT     TOP 100 PERCENT CAST(FLOOR(CAST(TransDate AS float)) AS datetime) AS TransDate, SUM(LotPrice) AS Gross, SUM(LevyAmount) AS Commission,
                      SystemMaintained, LevyShortName
FROM         dbo.LeviesReport
GROUP BY CAST(FLOOR(CAST(TransDate AS float)) AS datetime), SystemMaintained, LevyShortName

GO


CREATE VIEW [dbo].[AALeviesReport]
AS
/*SELECT     TOP 100 PERCENT dbo.Contract.Contract_DPA_, dbo.Contract.ContractTransferNo, dbo.Contract.ContractDeliveryDate, dbo.Lot.Lot_DPA_,
                      dbo.Lot.LotSlipNo, CAST(FLOOR(CAST(dbo.Lot.LotTDate AS float)) AS datetime) AS TransDate, dbo.Lot.LotQty, dbo.Lot.LotGrossAmount AS LotPrice,
                      dbo.Lot.ContractNumber, dbo.Contract.ContractDelivered, dbo.Contract.ContractNCertificate, dbo.Contract.ContractNCDate,
                      dbo.Contract.ContractNCDelivered, dbo.Contract.Voucher_DPA_, dbo.Contract.ContractVouchered, dbo.LevyContract.LevyAmount,
                      dbo.LevyContract.LevyName, dbo.tbOrder.OrderType_DPA_ AS OrdDetailType, dbo.Broker.BrokerCode, dbo.Client.ClientName AS OrdDetailClient,
                      dbo.Security.SecurityName AS OrdDetailSecurity, dbo.Security.OrderSecType_DPA_ AS OrdDetailSecType, dbo.Security.SecurityCode,
                      dbo.LevyContract.SystemMaintained, dbo.LevyContract.LevyShortName
FROM         dbo.Contract INNER JOIN
                      dbo.Lot ON dbo.Contract.Contract_DPA_ = dbo.Lot.Contract_DPA_ INNER JOIN
                      dbo.LevyContract ON dbo.Contract.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Broker ON dbo.Lot.Broker_DPA_ = dbo.Broker.Broker_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_
WHERE     (dbo.Client.Deleted = 0) AND (dbo.tbOrder.Deleted = 0) AND (dbo.LevyContract.Deleted = 0) AND (dbo.Contract.Deleted = 0) AND (dbo.Lot.Deleted = 0)
                      AND (dbo.OrdDetail.Deleted = 0)
ORDER BY dbo.tbOrder.OrderType_DPA_


*/



SELECT
	TOP 100 PERCENT Contract.Contract_DPA_, Contract.ContractTransferNo,
	Contract.ContractDeliveryDate, Lot.Lot_DPA_, Lot.LotSlipNo,
    CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime) AS TransDate,
	Lot.LotQty, Lot.LotGrossAmount AS LotPrice, Lot.ContractNumber,
	Contract.ContractDelivered, Contract.ContractNCertificate,
	Contract.ContractNCDate, Contract.ContractNCDelivered, Contract.Voucher_DPA_,
    Contract.ContractVouchered,
    CASE
		WHEN SystemMaintained = 11
		THEN LevyContract.LevyAmount
--		- (SELECT levyamount FROM	LevyContract
--        WHERE      systemMaintained = 25 AND deleted = 0 AND Contract_DPA_ = Contract.Contract_DPA_)
          ELSE dbo.LevyContract.LevyAmount END AS LevyAmount, LevyContract.LevyName, tbOrder.OrderType_DPA_ AS OrdDetailType, Broker.BrokerCode,
          Client.ClientName AS OrdDetailClient, Security.SecurityName AS OrdDetailSecurity, Security.OrderSecType_DPA_ AS OrdDetailSecType,
          Security.SecurityCode, LevyContract.SystemMaintained, LevyContract.LevyShortName
FROM         Contract INNER JOIN
          Lot ON Contract.Contract_DPA_ = Lot.Contract_DPA_ INNER JOIN
          LevyContract ON Contract.Contract_DPA_ = LevyContract.Contract_DPA_ INNER JOIN
          OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
          tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN
          Broker ON Lot.Broker_DPA_ = Broker.Broker_DPA_ INNER JOIN
          Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
          Security ON OrdDetail.Security_DPA_ = Security.Security_DPA_
WHERE     (Client.Deleted = 0) AND (tbOrder.Deleted = 0) AND (LevyContract.Deleted = 0) AND (Contract.Deleted = 0) AND (Lot.Deleted = 0) AND
          (OrdDetail.Deleted = 0)
ORDER BY tbOrder.OrderType_DPA_


GO
CREATE VIEW dbo.AccountEntityList
AS
SELECT     dbo.Account.AccountName AS EntityName, dbo.EntityType.EntityTypeName AS EntityType, dbo.Account.Account_DPA_ AS Entity_DPA_,
                      dbo.EntityType.EntityType_DPA_, dbo.Account.AccountCode AS EntityCode, dbo.Account.AccountCode + SPACE
                          ((SELECT     ISNULL(MAX(LEN(AccountCode)), 0) AS MAXLEN
                              FROM         dbo.Account) - LEN(dbo.Account.AccountCode)) + ' : ' + dbo.Account.AccountName AS EntityNameEx
FROM         dbo.Account INNER JOIN
                      dbo.EntityType ON dbo.Account.EntityType_DPA_ = dbo.EntityType.EntityType_DPA_

GO

CREATE VIEW dbo.AccountList
AS
SELECT     dbo.Account.AccountCode, dbo.Account.AccountName, dbo.AccountTypeLevel1.AccountTypeName AS AccountTypeLevel1,
                      dbo.AccountTypeLevel2.AccountTypeName AS AccountTypeLevel2, dbo.AccountTypeLevel3.AccountTypeName AS AccountTypeLevel3,
                      dbo.Account.Account_DPA_, dbo.AccountTypeLevel1.AccountType_DPA_ AS AccountTypeLevel1_DPA_
FROM         dbo.AccountTypeLevel1 INNER JOIN
                      dbo.Account ON dbo.AccountTypeLevel1.AccountType_DPA_ = dbo.Account.AccountTypeLevel1 LEFT OUTER JOIN
                      dbo.AccountTypeLevel2 ON dbo.Account.AccountTypeLevel2 = dbo.AccountTypeLevel2.AccountType_DPA_ LEFT OUTER JOIN
                      dbo.AccountTypeLevel3 ON dbo.Account.AccountTypeLevel3 = dbo.AccountTypeLevel3.AccountType_DPA_


GO

CREATE VIEW dbo.AccountManagerCreditors
AS
SELECT TOP 100 PERCENT dbo.DebtorCreditor.Balance, dbo.DebtorCreditor.Client_DPA_, dbo.DebtorCreditor.LastDate, dbo.Client.Owner_DPA_,
               dbo.Client.ClientName
FROM  dbo.DebtorCreditor INNER JOIN
               dbo.Client ON dbo.DebtorCreditor.Client_DPA_ = dbo.Client.Client_DPA_
WHERE (dbo.DebtorCreditor.Balance > 0)
ORDER BY dbo.Client.ClientName


GO

CREATE VIEW dbo.AccountManagerDebtors
AS
SELECT TOP 100 PERCENT dbo.DebtorCreditor.Balance, dbo.DebtorCreditor.Client_DPA_, dbo.DebtorCreditor.LastDate, dbo.Client.ClientName,
               dbo.Client.Owner_DPA_
FROM  dbo.DebtorCreditor INNER JOIN
               dbo.Client ON dbo.DebtorCreditor.Client_DPA_ = dbo.Client.Client_DPA_
WHERE (dbo.DebtorCreditor.Balance < 0)
ORDER BY dbo.Client.ClientName


GO
CREATE VIEW dbo.AccountManagerTotals
AS
SELECT     ISNULL(dbo.Owner.Owner_DPA_, 0) AS [Account Manager Code], ISNULL(dbo.Owner.OwnerFname + ' ' + dbo.Owner.OwnerLName, 'NONE')
                      AS [Account Manager], SUM(dbo.Lot.LotQty) AS Volume, SUM(dbo.LevyContract.LevyAmount) AS Commission
FROM         dbo.Lot INNER JOIN
                      dbo.OrdDetail ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.LevyContract ON dbo.Lot.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ LEFT OUTER JOIN
                      dbo.Owner ON dbo.Client.Owner_DPA_ = dbo.Owner.Owner_DPA_
GROUP BY dbo.Owner.Owner_DPA_, dbo.Owner.OwnerFname, dbo.Owner.OwnerLName



GO

CREATE VIEW dbo.AccountTypeLevel1
AS
SELECT     AccountTypeName, AccountType_DPA_, DefaultSelection, AccountTypeParent
FROM         dbo.AccountType
WHERE     (AccountTypeParent IS NULL)




GO

CREATE VIEW dbo.AccountTypeLevel2
AS
SELECT     AccountType_DPA_, AccountTypeName, DefaultSelection, AccountTypeParent, ISNULL(Quarter1, 0) AS Quarter1, ISNULL(Quarter2, 0) AS Quarter2,
                      ISNULL(Quarter3, 0) AS Quarter3, ISNULL(Quarter4, 0) AS Quarter4
FROM         dbo.AccountType
WHERE     (AccountTypeParent IN
                          (SELECT     AccountType_DPA_
                            FROM          AccountTypeLevel1))


GO
CREATE VIEW dbo.AccountTypeLevel3
AS
SELECT     AccountType_DPA_, AccountTypeName, DefaultSelection, AccountTypeParent
FROM         dbo.AccountType
WHERE     (AccountTypeParent IN
                          (SELECT     AccountType_DPA_
                            FROM          AccountTypeLevel2))

GO


CREATE VIEW dbo.AccountTypeList
AS
SELECT     AccountType_DPA_, AccountTypeName AS AccountTypeParent, 'Level1' AS AccountTypeLevel, AccountTypeName AS AccountTypeSetting, 1 AS AccountTypeLevel_DPA_
FROM         dbo.AccountTypeLevel1
UNION ALL
SELECT     dbo.AccountTypeLevel2.AccountType_DPA_, dbo.AccountTypeLevel1.AccountTypeName AS AccountTypeParent, 'Level2' AS AccountTypeLevel,
                      dbo.AccountTypeLevel2.AccountTypeName AS AccountTypeSetting, 2 AS AccountTypeLevel_DPA_
FROM         dbo.AccountTypeLevel2 INNER JOIN
                      dbo.AccountTypeLevel1 ON dbo.AccountTypeLevel2.AccountTypeParent = dbo.AccountTypeLevel1.AccountType_DPA_
UNION ALL
SELECT     dbo.AccountTypeLevel3.AccountType_DPA_, dbo.AccountTypeLevel2.AccountTypeName AS AccountTypeParent, 'Level3' AS AccountTypeLevel,
                      dbo.AccountTypeLevel3.AccountTypeName AS AccountTypeSetting, 3 AS AccountTypeLevel_DPA_
FROM         dbo.AccountTypeLevel3 INNER JOIN
                      dbo.AccountTypeLevel2 ON dbo.AccountTypeLevel3.AccountTypeParent = dbo.AccountTypeLevel2.AccountType_DPA_






GO
CREATE VIEW dbo.ActivatedClientList
AS
SELECT     dbo.FullClientList.*
FROM         dbo.FullClientList

GO
CREATE VIEW dbo.ActivityList
AS
SELECT     TOP 100 PERCENT dbo.ClientList.ClientName, dbo.ActvtyClass.ActvtyClassDescription AS ActivityName, CONVERT(DATETIME,
                      dbo.Activity.ActivityDate, 108) AS ActivityDate, dbo.Activity.Activity_DPA_, dbo.ClientList.Client_DPA_, dbo.Activity.ActivityNotes AS Notes,
                      dbo.UserList.UserName, dbo.Activity.TimeChanged, dbo.Activity.Deleted
FROM         dbo.ActvtyClass INNER JOIN
                      dbo.Activity INNER JOIN
                      dbo.ClientList ON dbo.Activity.Client_DPA_ = dbo.ClientList.Client_DPA_ ON
                      dbo.ActvtyClass.ActvtyClass_DPA_ = dbo.Activity.ActvtyClass_DPA_ LEFT OUTER JOIN
                      dbo.UserList ON dbo.Activity.ChangedBy = dbo.UserList.UserID
WHERE     (dbo.Activity.Deleted = 0) OR
                      (dbo.Activity.Deleted IS NULL)
ORDER BY dbo.ClientList.ClientName, dbo.ActvtyClass.ActvtyClassDescription

GO

CREATE VIEW dbo.ActvtyClassList
AS
SELECT     TOP 100 PERCENT ActvtyClassDescription, ActvtyClass_DPA_, ClientAccess
FROM         dbo.ActvtyClass
ORDER BY ActvtyClassDescription


GO

CREATE VIEW dbo.AgentBalanceList
AS
SELECT DISTINCT
                      TOP 100 PERCENT dbo.Agent.AgentName AS Agent, a.CurrentBal, CAST(FLOOR(CAST(dbo.AgentStatement.TransDate AS Float)) AS DateTime)
                      AS TransDate
FROM         dbo.AgentStatement INNER JOIN
                      dbo.Agent ON dbo.AgentStatement.Agent_DPA_ = dbo.Agent.Agent_DPA_ INNER JOIN
                          (SELECT     SUM(ISNULL(dbo.AgentStatement.Credit - dbo.AgentStatement.Debit, 0)) + dbo.Agent.AgentOpeningBal AS CurrentBal,
                                                   dbo.AgentStatement.Agent_DPA_
                            FROM          dbo.AgentStatement INNER JOIN
                                                   dbo.Agent ON dbo.AgentStatement.Agent_DPA_ = dbo.Agent.Agent_DPA_
                            WHERE      (dbo.Agent.Deleted = 0)
                            GROUP BY dbo.AgentStatement.Agent_DPA_, dbo.Agent.AgentOpeningBal) a ON dbo.Agent.Agent_DPA_ = a.Agent_DPA_
WHERE     (dbo.Agent.Deleted = 0)


GO
CREATE VIEW [dbo].[AgentCommissionList]
AS
SELECT     CAST(LotQty AS nvarchar) + ' ' + SecurityName + ' @ ' + CAST(LotPrice AS nvarchar) AS SecurityName, ContractNumber, LotTDate, Agent_DPA_,
                      LotGrossAmount AS Gross, LevyAmount AS AgentCommission
FROM         (SELECT     dbo.Lot.ContractNumber, dbo.Lot.LotTDate, dbo.Agent.Agent_DPA_, dbo.Lot.LotGrossAmount, SUM(dbo.LevyContract.LevyAmount)
                                              AS LevyAmount, dbo.Lot.LotQty, dbo.Lot.LotPrice, dbo.Security.SecurityName
                       FROM          dbo.Lot INNER JOIN
                                              dbo.LevyContract ON dbo.Lot.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                                              dbo.OrdDetail ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                                              dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                                              dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                                              dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                                              dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_


INNER JOIN    dbo.Agent ON dbo.Client.Agent_DPA_ = dbo.Agent.Agent_DPA_
--INNER JOIN dbo.Agent ON dbo.tbOrder.Agent_DPA_ = dbo.Agent.Agent_DPA_



                       WHERE      dbo.LevyContract.SystemMaintained = 12 AND dbo.Lot.Deleted <> 1
AND LEVYAMOUNT<>0
                       GROUP BY dbo.Lot.ContractNumber, dbo.Lot.LotTDate, dbo.Agent.Agent_DPA_, dbo.Lot.LotGrossAmount, dbo.Lot.LotQty, dbo.Lot.LotPrice,
                                              dbo.Security.SecurityName


) innerTable

GO

CREATE VIEW dbo.AgentCommissionList1
AS
SELECT     CAST(LotQty AS nvarchar) + ' ' + SecurityName + ' @ ' + CAST(LotPrice AS nvarchar) AS SecurityName, ContractNumber, LotTDate, Agent_DPA_,
                      LotGrossAmount AS Gross, LevyAmount AS AgentCommission
FROM         (SELECT     dbo.Lots.ContractNumber, dbo.Lots.LotTDate, dbo.Agent.Agent_DPA_, dbo.Lots.LotGrossAmount, SUM(dbo.LevyContract.LevyAmount)
                                              AS LevyAmount, dbo.Lots.LotQty, dbo.Lots.LotPrice, dbo.Security.SecurityCode as SecurityName
                       FROM          dbo.Lots INNER JOIN
                                              dbo.LevyContract ON dbo.Lots.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                                              dbo.OrdDetail ON dbo.Lots.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                                              dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                                              dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                                              dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                                              dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                                              dbo.Agent ON dbo.Client.Agent_DPA_ = dbo.Agent.Agent_DPA_
                       WHERE      dbo.LevyContract.SystemMaintained = 12 AND LevyContract.Deleted = 0 AND Lots.Deleted = 0 AND OrdDetail.Deleted = 0 AND
                                              tbOrder.Deleted = 0 AND Client.Deleted = 0 AND Agent.Deleted = 0
                       GROUP BY dbo.Lots.ContractNumber, dbo.Lots.LotTDate, dbo.Agent.Agent_DPA_, dbo.Lots.LotGrossAmount, dbo.Lots.LotQty, dbo.Lots.LotPrice,
                                              dbo.Security.SecurityCode) innerTable





GO

CREATE VIEW dbo.AgentCommissions
AS
SELECT     TOP 100 PERCENT dbo.Lots.ContractNumber, CAST(FLOOR(CAST(dbo.Lots.LotTDate AS Float)) AS DateTime) AS LotTDate, dbo.tbOrder.Client_DPA_,
                      dbo.OrderType.OrderTypeSale, dbo.LevyContract.LevyAmount, CAST(dbo.Lots.LotQty AS nvarchar)
                      + ' ' + dbo.Security.SecurityCode + ' @ ' + CAST(dbo.Lots.LotPrice AS nvarchar) AS SecurityName, dbo.Payment.PaymentReceiptNo,
                      dbo.LevyContract.Contract_DPA_, dbo.AgentList.AgentName, dbo.AgentList.Agent_DPA_, dbo.Client.ClientName, dbo.Lots.LotGrossAmount
FROM         dbo.Lots INNER JOIN
                      dbo.LevyContract ON dbo.Lots.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.Lots.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.AgentList ON dbo.Client.Agent_DPA_ = dbo.AgentList.Agent_DPA_ LEFT OUTER JOIN
                      dbo.Payment ON dbo.tbOrder.Order_DPA_ = dbo.Payment.Order_DPA_
WHERE     (dbo.tbOrder.OrderCanceled = 0) AND (dbo.LevyContract.SystemMaintained <> 8) AND (dbo.LevyContract.SystemMaintained <> 12) AND
                      (dbo.LevyContract.LevyShortName LIKE N'%Commission%') AND (dbo.Client.Deleted = 0) AND (dbo.OrdDetail.Deleted = 0) AND
                      (dbo.tbOrder.Deleted = 0)
ORDER BY dbo.AgentList.Agent_DPA_, dbo.Lots.LotTDate


GO
CREATE VIEW dbo.AgentCommissionTotal
AS
SELECT     Contract_DPA_, SUM(LevyAmount) AS LevyTotal
FROM         dbo.LevyContract
WHERE     (SystemMaintained = 12) AND (Deleted = 0)
GROUP BY Contract_DPA_

GO
CREATE VIEW dbo.AgentContractCompounded
AS
SELECT     TOP 100 PERCENT dbo.Contract.Contract_DPA_, dbo.LotList.Lot_DPA_, dbo.LotList.BrokerCode, dbo.LotList.OrdDetailSecType,
                      dbo.LotList.OrdDetailType, dbo.LotList.OrderTypeSale, dbo.LotList.LotSlipNo, dbo.LotList.LotTDate, dbo.LotList.LotQty, dbo.LotList.LotPrice,
                      dbo.LotList.ContractNumber, dbo.LotList.OrdDetailClient, dbo.LotList.OrdDetailSecurity, dbo.LevyContract.LevyAmount, dbo.LevyContract.LevyName,
                      dbo.tbOrder.OrderRef, dbo.Client.ClientAddr, dbo.LotList.LotGrossAmount AS GrossAmount, dbo.tbOrder.Order_DPA_,
                      dbo.LevyContract.SystemMaintained, CASE WHEN ISNULL(CHARINDEX('%', dbo.LevyContract.LevyRatePercentage), 0)
                      > 0 THEN '%' ELSE '' END AS LevyRatePercentage, dbo.LevyContract.LevyShortName, dbo.Agent.Agent_DPA_, dbo.Agent.AgentName,
                      CAST(dbo.Agent.AgentAddr AS NVARCHAR) AS AgentAddr, dbo.Client.Client_DPA_, dbo.Client.ClientCDSNo, dbo.LotList.OrdDetail_DPA_,
                      dbo.LevyContract.LevyVATAmount, dbo.LotList.ContractSettlementDate
FROM         dbo.Contract INNER JOIN
                      dbo.LotList ON dbo.Contract.Contract_DPA_ = dbo.LotList.Contract_DPA_ INNER JOIN
                      dbo.LevyContract ON dbo.Contract.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.LotList.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.LotList.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Agent ON dbo.Agent.Agent_DPA_ = dbo.Client.Agent_DPA_
WHERE     (dbo.tbOrder.OrderCompounded = 1) AND (ISNULL(dbo.Client.Agent_DPA_, 0) <> 0)
ORDER BY dbo.LotList.LotTDate, dbo.LotList.OrdDetailSecurity, dbo.LotList.LotSlipNo

GO
CREATE VIEW dbo.AgentContracts
AS
SELECT     dbo.Contract.Contract_DPA_, dbo.Contract.ContractTransferNo, dbo.Contract.ContractDeliveryDate, dbo.LotList.Lot_DPA_, dbo.LotList.BrokerCode,
                      dbo.LotList.OrdDetailSecType, dbo.LotList.OrdDetailType, dbo.LotList.OrderTypeSale, dbo.LotList.LotSlipNo, dbo.LotList.LotTDate, dbo.LotList.LotQty,
                      dbo.LotList.LotPrice, dbo.LotList.LotGrossAmount, dbo.LotList.ContractNumber, dbo.LotList.StatusDescription, dbo.LotList.OrdDetailClient,
                      dbo.Contract.ContractDelivered, dbo.LotList.OrdDetailSecurity, dbo.Contract.ContractNCertificate, dbo.Contract.ContractNCDate,
                      dbo.Contract.ContractNCDelivered, dbo.Contract.Voucher_DPA_, dbo.Contract.ContractVouchered, dbo.LevyContract.LevyAmount,
                      dbo.LevyContract.LevyName, dbo.LevyContract.SystemMaintained, CASE WHEN ISNULL(CHARINDEX('%', dbo.LevyContract.LevyRatePercentage), 0)
                      > 0 THEN dbo.LevyContract.LevyRatePercentage ELSE '' END AS LevyRatePercentage, dbo.LevyContract.LevyShortName, dbo.Agent.Agent_DPA_,
                      dbo.Agent.AgentName, dbo.Agent.AgentAddr, dbo.Commission.CommissionDescription, dbo.Commission.MinimumSecurityCommission,
                      dbo.LevyContract.LevyVATAmount, dbo.LotList.ContractSettlementDate
FROM         dbo.Contract INNER JOIN
                      dbo.LotList ON dbo.Contract.Contract_DPA_ = dbo.LotList.Contract_DPA_ INNER JOIN
                      dbo.LevyContract ON dbo.Contract.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.LotList.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.Client.Client_DPA_ = dbo.LotList.Client_DPA_ INNER JOIN
                      dbo.Agent ON dbo.Agent.Agent_DPA_ = dbo.Client.Agent_DPA_ INNER JOIN
                      dbo.Commission ON dbo.Client.Commission_DPA_ = dbo.Commission.Commission_DPA_
WHERE     (dbo.tbOrder.OrderCompounded = 0) AND (ISNULL(dbo.Client.Agent_DPA_, 0) <> 0)

GO

CREATE VIEW dbo.AgentCreditors
AS
SELECT     TOP 100 PERCENT dbo.Client.Agent_DPA_ AS Agent, dbo.DebtorCreditor.Client_DPA_, dbo.DebtorCreditor.LastDate, dbo.DebtorCreditor.Balance,
                      dbo.Client.ClientName
FROM         dbo.Client INNER JOIN
                      dbo.DebtorCreditor ON dbo.Client.Client_DPA_ = dbo.DebtorCreditor.Client_DPA_
WHERE     (dbo.DebtorCreditor.Balance > 0) AND (dbo.Client.Deleted = 0)
ORDER BY dbo.Client.ClientName


GO

CREATE VIEW dbo.AgentDebtors
AS
SELECT     TOP 100 PERCENT dbo.Client.Agent_DPA_ AS Agent, dbo.DebtorCreditor.Client_DPA_, dbo.DebtorCreditor.LastDate, dbo.DebtorCreditor.Balance,
                      dbo.Client.ClientName
FROM         dbo.Client INNER JOIN
                      dbo.DebtorCreditor ON dbo.Client.Client_DPA_ = dbo.DebtorCreditor.Client_DPA_
WHERE     (dbo.DebtorCreditor.Balance < 0) AND (dbo.Client.Deleted = 0)
ORDER BY dbo.Client.ClientName


GO
CREATE VIEW dbo.AgentEntityList
AS
SELECT     dbo.Agent.AgentName AS EntityName, dbo.EntityType.EntityTypeName AS EntityType, dbo.Agent.Agent_DPA_ AS Entity_DPA_,
                      dbo.EntityType.EntityType_DPA_, CAST(dbo.Agent.Agent_DPA_ AS NVARCHAR(4000)) AS EntityCode, CAST(dbo.Agent.Agent_DPA_ AS NVARCHAR(4000))
                      + SPACE
                          ((SELECT     ISNULL(MAX(LEN(CAST(dbo.Agent.Agent_DPA_ AS NVARCHAR(4000)))), 0) AS MAXLEN
                              FROM         dbo.Agent) - LEN(CAST(dbo.Agent.Agent_DPA_ AS NVARCHAR(4000)))) + ' : ' + dbo.Agent.AgentName AS EntityNameEx
FROM         dbo.Agent INNER JOIN
                      dbo.EntityType ON dbo.Agent.EntityType_DPA_ = dbo.EntityType.EntityType_DPA_
WHERE     (dbo.Agent.AgentName <> N'_none_')

GO
CREATE VIEW dbo.AgentList
AS
SELECT     TOP 100 PERCENT dbo.Agent.AgentName, dbo.Agent.AgentContact, dbo.Agent.AgentOfficeTel, dbo.Agent.AgentCellTel, dbo.Agent.AgentEmail,
                      REPLACE(CAST(dbo.Agent.AgentAddr AS NVARCHAR(4000)), CHAR(13) + CHAR(10), ' ') AS AgentAddr, dbo.Agent.Agent_DPA_, dbo.Agent.DefaultSelection,
                      dbo.Commission.CommissionRate AS AgentCommission, dbo.Agent.TimeChanged, dbo.UserList.[USER] AS ChangedBy
FROM         dbo.Agent INNER JOIN
                      dbo.Commission ON dbo.Agent.Commission_DPA_ = dbo.Commission.Commission_DPA_ LEFT OUTER JOIN
                      dbo.UserList ON dbo.Agent.ChangedBy = dbo.UserList.UserID
WHERE     (dbo.Agent.Deleted = 0) OR
                      (dbo.Agent.Deleted IS NULL)
ORDER BY dbo.Agent.AgentName

GO

CREATE VIEW dbo.AgentRCommissionList
AS
SELECT     dbo.Commission.CommissionRate AS Commission, dbo.Client.Client_DPA_, dbo.Class.AgentStatus, dbo.Agent.Agent_DPA_,
                      dbo.Commission.Commission_DPA_
FROM         dbo.Commission INNER JOIN
                      dbo.Agent ON dbo.Commission.Commission_DPA_ = dbo.Agent.Commission_DPA_ INNER JOIN
                      dbo.Class INNER JOIN
                      dbo.Client ON dbo.Class.Class_DPA_ = dbo.Client.Class_DPA_ ON dbo.Agent.Agent_DPA_ = dbo.Client.Agent_DPA_
WHERE     (dbo.Client.Client_DPA_ = 101946)


GO
CREATE VIEW dbo.AgentReport
AS

SELECT     TOP 100 PERCENT dbo.Agent.AgentName As Agent, dbo.Agent.AgentContact As Contact, dbo.Agent.AgentOfficeTel AS [Office Phone], dbo.Agent.AgentCellTel As [Cell Phone], dbo.Agent.AgentEmail AS Email,
                      REPLACE(CAST(dbo.Agent.AgentAddr AS NVARCHAR(4000)), CHAR(13) + CHAR(10), ' ') AS Address,
                      dbo.Commission.CommissionRate AS Commission
FROM         dbo.Agent INNER JOIN
                      dbo.Commission ON dbo.Agent.Commission_DPA_ = dbo.Commission.Commission_DPA_
ORDER BY dbo.Agent.AgentName


GO

CREATE VIEW dbo.AgentReports
AS
SELECT DISTINCT TOP 100 PERCENT a.Client_DPA_, MAX(a.TransDate) AS LastDate, SUM(a.Balance) AS Balance
FROM         (SELECT     *
                       FROM          dbo.ClientTransactionList) a INNER JOIN
                      dbo.Client ON a.Client_DPA_ = dbo.Client.Client_DPA_
WHERE     (dbo.Client.Deleted = 0)
GROUP BY a.Client_DPA_
ORDER BY MAX(a.TransDate) DESC


GO

CREATE   VIEW app.Agents AS
SELECT  Agent_DPA_         AS AgentId,
        AgentName          AS Name,
        AgentEmail         AS Email,
        AgentCellTel       AS Mobile,
        Branch_DPA_        AS BranchId,
        Commission_DPA_    AS CommissionId,
        AgentOpeningBal    AS OpeningBalance,
        AgentRegDate       AS RegisteredOn
FROM dbo.Agent
WHERE Deleted = 0 OR Deleted IS NULL;

GO
CREATE VIEW dbo.AgentStatement
AS
SELECT     TOP (100) PERCENT COUNT(*) AS ClientTransaction_DPA_, Agent_DPA_, CAST(FLOOR(CAST(TransDate AS Float)) AS DateTime) AS TransDate, REF, Particulars, Debit,
                      Credit, SUM(Balance) AS Balance, IsOpeningBalance, Gross
FROM         (SELECT     Agent_DPA_, TransDate, REF, Particulars, Debit, Credit, CreditBal, IsOpeningBalance, Gross, Balance
                       FROM          dbo.AgentTransactionList) AS a
GROUP BY IsOpeningBalance, Agent_DPA_, TransDate, REF, Particulars, Debit, Credit, Balance, Gross
ORDER BY IsOpeningBalance DESC, Agent_DPA_, TransDate, Particulars, REF, Debit, Credit, Balance

GO

CREATE VIEW dbo.AgentStatement1
AS
SELECT     TOP 100 PERCENT COUNT(*) AS ClientTransaction_DPA_, Agent_DPA_, TransDate, REF, Particulars, Debit, Credit, SUM(Balance) AS Balance,
                      IsOpeningBalance
FROM         (SELECT     *
                       FROM          dbo.AgentTransactionList1) a
GROUP BY IsOpeningBalance, Agent_DPA_, TransDate, REF, Particulars, Debit, Credit, Balance
ORDER BY IsOpeningBalance DESC, Agent_DPA_, TransDate, Particulars, REF, Debit, Credit, Balance

GO

CREATE VIEW dbo.AgentTransactionList
AS
SELECT     AgentTransactions.*, CreditBal - Debit AS Balance
FROM         (SELECT     Agent_DPA_, AgentRegDate AS TransDate, '' AS REF, ' Opening Balance' AS Particulars, 0 AS Debit, 0 AS Credit,
                                              AgentOpeningBal AS CreditBal, 1 AS IsOpeningBalance, 0 AS Gross
                       FROM          dbo.Agent
                       WHERE      Deleted = 0
                       UNION ALL
                       SELECT     dbo.Payment.Entity_DPA_ AS Agent_DPA_, dbo.Payment.PaymentPDate AS TransDate, dbo.Payment.PaymentReference AS Ref,
                                             dbo.Payment.PaymentNarrative AS Particulars, CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS CreditBal, 0 AS IsOpeningBalance, 0 AS Gross
                       FROM         dbo.Payment INNER JOIN
                                             dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
                       WHERE     EntityType_DPA_ = 2 AND Payment.Deleted = 0
                       UNION ALL
                       SELECT     Agent_DPA_, LotTDate AS TransDate, ContractNumber AS Ref, SecurityName AS Particulars, 0 AS Debit, isnull(AgentCommission, 0)
                                             AS Credit, isnull(AgentCommission, 0) AS CreditBal, 0 AS IsOpeningBalance, Gross AS Gross
                       FROM         dbo.AgentCommissionList
                       UNION ALL
                       SELECT     dbo.JournalList.Entity_DPA_ AS Agent_DPA_, dbo.JournalList.JournalDate AS TransDate,
                                             CAST(dbo.JournalList.JournalEntry_DPA_ AS NVARCHAR(500)) AS Ref, dbo.JournalList.JournalNarrative AS Particulars,
                                             JournalList.JournalEntryDebit AS Debit, JournalList.JournalEntryCredit AS Credit, JournalList.JournalEntryCredit AS CreditBal,
                                             0 AS IsOpeningBalance, 0 AS Gross
                       FROM         dbo.JournalList
                       WHERE     EntityType_DPA_ = 2) AgentTransactions


GO

CREATE VIEW dbo.AgentTransactionList1
AS
SELECT     AgentTransactions.*, CreditBal - Debit AS Balance
FROM         (SELECT     Agent_DPA_, AgentRegDate AS TransDate, '' AS REF, ' Opening Balance' AS Particulars, 0 AS Debit, 0 AS Credit,
                                              AgentOpeningBal AS CreditBal, 1 AS IsOpeningBalance
                       FROM          dbo.Agent
                       WHERE      Deleted = 0
                       UNION ALL
                       SELECT     dbo.Payment.Entity_DPA_ AS Agent_DPA_, dbo.Payment.PaymentPDate AS TransDate, dbo.Payment.PaymentReference AS Ref,
                                             dbo.Payment.PaymentNarrative AS Particulars, CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS CreditBal, 0 AS IsOpeningBalance
                       FROM         dbo.Payment INNER JOIN
                                             dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
                       WHERE     EntityType_DPA_ = 2 AND Payment.Deleted = 0
                       UNION ALL
                       SELECT     Agent_DPA_, LotTDate AS TransDate, ContractNumber AS Ref, SecurityName AS Particulars, 0 AS Debit, isnull(AgentCommission, 0)
                                             AS Credit, isnull(AgentCommission, 0) AS CreditBal, 0 AS IsOpeningBalance
                       FROM         dbo.AgentCommissionList1
                       UNION ALL
                       SELECT     dbo.JournalList.Entity_DPA_ AS Agent_DPA_, dbo.JournalList.JournalDate AS TransDate,
                                             CAST(dbo.JournalList.JournalEntry_DPA_ AS NVARCHAR(500)) AS Ref, dbo.JournalList.JournalNarrative AS Particulars,
                                             JournalList.JournalEntryDebit AS Debit, JournalList.JournalEntryCredit AS Credit, JournalList.JournalEntryCredit AS CreditBal,
                                             0 AS IsOpeningBalance
                       FROM         dbo.JournalList
                       WHERE     EntityType_DPA_ = 2) AgentTransactions



GO
CREATE VIEW dbo.AgentVolumeTotals
AS
SELECT     dbo.Agent.Agent_DPA_ AS [Agent Code], dbo.Agent.AgentName AS [Agent Name], SUM(dbo.Lot.LotQty) AS Volume, SUM(dbo.LevyContract.LevyAmount)
                      AS Commission
FROM         dbo.Lot INNER JOIN
                      dbo.OrdDetail ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.LevyContract ON dbo.Lot.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.Agent ON dbo.Client.Agent_DPA_ = dbo.Agent.Agent_DPA_
GROUP BY dbo.Agent.Agent_DPA_, dbo.Agent.AgentName

GO

CREATE VIEW dbo.AllocationSchedule
AS
SELECT DISTINCT
                      TOP 100 PERCENT dbo.Client.Client_DPA_, dbo.Client.ClientName, dbo.Agent.Agent_DPA_ AS AgentCode, dbo.Agent.AgentName,
                      dbo.Client.ClientCDSNo, dbo.LotList.Order_DPA_, dbo.LotList.BalanceQty, dbo.LotList.OrderDate, dbo.LotList.SecurityCode, dbo.LotList.OrdDetailPrice,
                      dbo.LotList.OrderTypeSale, dbo.LotList.Security_DPA_, dbo.Security.SecurityName, dbo.LotList.Best
FROM         dbo.LotList INNER JOIN
                      dbo.Client ON dbo.LotList.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Security ON dbo.LotList.Security_DPA_ = dbo.Security.Security_DPA_ LEFT OUTER JOIN
                      dbo.Agent ON dbo.Client.Agent_DPA_ = dbo.Agent.Agent_DPA_
WHERE     (dbo.LotList.BalanceQty > 0)
ORDER BY dbo.LotList.Security_DPA_, dbo.LotList.OrderTypeSale DESC, dbo.LotList.Best DESC, dbo.LotList.OrdDetailPrice


GO
CREATE VIEW dbo.APortfolioList AS SELECT APortfolio.APortfolioPrice AS APortfolioPrice, APortfolio.APortfolioQty AS APortfolioQty, APortfolio.APortfolio_DPA_ AS APortfolio_DPA_
FROM APortfolio

GO

CREATE VIEW dbo.BalanceAccountList
AS
SELECT DISTINCT
               dbo.ClientStatement.Client_DPA_, dbo.ClientStatement.Balance, dbo.Client.ClientName, dbo.Owner.Owner_DPA_,
               dbo.Client.ClientCellTel + '/' + dbo.Client.ClientOfficeTel + '/' + dbo.Client.ClientHomeTel + '/' + dbo.Client.ClientEmail AS Contacts,
               dbo.Owner.OwnerFname + '  ' + dbo.Owner.OwnerLName AS Owner, dbo.ClientStatement.TransDate
FROM  dbo.ClientStatement INNER JOIN
               dbo.Client ON dbo.ClientStatement.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
               dbo.Owner ON dbo.Client.Owner_DPA_ = dbo.Owner.Owner_DPA_


GO

CREATE VIEW dbo.BalanceAgentList
AS
SELECT dbo.ClientStatement.Client_DPA_, dbo.ClientStatement.Balance, dbo.Client.ClientName,
               dbo.Client.ClientCellTel + '/' + dbo.Client.ClientOfficeTel + '/' + dbo.Client.ClientHomeTel + '/' + dbo.Client.ClientEmail AS Contacts,
               dbo.Agent.AgentName AS Owner, dbo.ClientStatement.TransDate, dbo.Client.Agent_DPA_
FROM  dbo.ClientStatement INNER JOIN
               dbo.Client ON dbo.ClientStatement.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
               dbo.Agent ON dbo.Client.Agent_DPA_ = dbo.Agent.Agent_DPA_


GO

CREATE VIEW dbo.BankAccList
AS
SELECT     TOP 100 PERCENT dbo.Bank.BankName + ' ' + dbo.BankAcc.BankAccNumber + ' ' + dbo.BnkBranch.BnkBranchName AS BankAccAccount,
                      dbo.Client.ClientName AS BankAccClient, dbo.BankAcc.BankAcc_DPA_ AS BankAcc_DPA_
FROM         dbo.Client INNER JOIN
                      dbo.Bank INNER JOIN
                      dbo.BnkBranch ON dbo.Bank.Bank_DPA_ = dbo.BnkBranch.Bank_DPA_ INNER JOIN
                      dbo.BankAcc ON dbo.BnkBranch.BnkBranch_DPA_ = dbo.BankAcc.BnkBranch_DPA_ ON dbo.Client.Client_DPA_ = dbo.BankAcc.Client_DPA_
ORDER BY dbo.Client.ClientName


GO

CREATE VIEW dbo.BankAccountList
AS
SELECT     dbo.Account.Account_DPA_, dbo.Account.AccountName, dbo.Account.AccountCode, dbo.Account.AccountCode + SPACE
                          ((SELECT     ISNULL(MAX(LEN(AccountCode)), 0) AS MAXLEN
                              FROM         dbo.Account) - LEN(dbo.Account.AccountCode)) + ' : ' + dbo.Account.AccountName AS AccountNameEx
FROM         dbo.Account INNER JOIN
                      dbo.AccountType ON dbo.Account.AccountTypeLevel1 = dbo.AccountType.AccountType_DPA_
WHERE     (dbo.Account.AccountTypeLevel1 = 7) OR
                      (dbo.Account.Account_DPA_ = 4)


GO

CREATE VIEW dbo.BankAccountStatement
AS
SELECT TOP 100 PERCENT a.BankTransaction_DPA_, a.BankAccount_DPA_, a.TransDate, a.Ref, a.Particulars, a.Debit, a.Credit,
               CASE WHEN (SUM(b.Balance) >= 0) THEN CONVERT(NVARCHAR, SUM(b.Balance)) + ' Cr' ELSE CONVERT(NVARCHAR, ABS(SUM(b.Balance)))
               + ' Dr' END AS Balance, a.OpeningBalance AS IsOpeningBalance, a.ReceiptNo
FROM  BankTransactionList a CROSS JOIN
               BankTransactionList b
WHERE (b.BankTransaction_DPA_ <= a.BankTransaction_DPA_) AND (a.BankAccount_DPA_ = b.BankAccount_DPA_)
GROUP BY a.BankAccount_DPA_, a.TransDate, a.BankTransaction_DPA_, a.Ref, a.Particulars, a.Debit, a.Credit, a.OpeningBalance, a.ReceiptNo
ORDER BY a.BankAccount_DPA_, a.BankTransaction_DPA_

GO

CREATE VIEW dbo.BankList
AS
SELECT     TOP 100 PERCENT dbo.Bank.BankCode, dbo.Bank.BankName, dbo.Bank.Bank_DPA_, dbo.BnkBranch.BnkBranch_DPA_,
                      dbo.BnkBranch.BnkBranch_EIT_, dbo.BnkBranch.BnkBranchName, dbo.BnkBranch.BnkBranchCode, dbo.BnkBranch.BnkBranchSwiftCode
FROM         dbo.Bank INNER JOIN
                      dbo.BnkBranch ON dbo.Bank.Bank_DPA_ = dbo.BnkBranch.Bank_DPA_
ORDER BY dbo.Bank.BankName


GO

CREATE VIEW dbo.BankTransactionList
AS
SELECT     TOP 100 PERCENT COUNT(*) AS BankTransaction_DPA_, a.BankAccount_DPA_, a.TransDate, a.Ref, a.Particulars, a.Debit, a.Credit, a.Balance,
                      a.OpeningBalance, a.ReceiptNo
FROM         (SELECT     AccountsStatement.*, Credit - Debit AS Balance
                       FROM          (SELECT     Account_DPA_ AS BankAccount_DPA_, AccountOpeningDate AS TransDate, '' AS Ref, 'Opening Balance' AS Particulars,
                                                                      CASE WHEN (AccountOpeningBal < 0) THEN AccountOpeningBal ELSE 0 END AS Debit, CASE WHEN (AccountOpeningBal >= 0)
                                                                      THEN AccountOpeningBal ELSE 0 END AS Credit, 1 AS OpeningBalance, '' AS ReceiptNo
                                               FROM          dbo.Account
                                               WHERE      AccountTypeLevel1 = 7
                                               UNION ALL
                                               SELECT     dbo.Payment.BankAccount_DPA_, dbo.Payment.PaymentPDate AS TransDate, dbo.Payment.PaymentReference AS Ref,
                                                                     dbo.Payment.PaymentNarrative AS Particulars,
                                                                     CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                                                                     CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Credit, 0 AS OpeningBalance,
                                                                     PaymentReceiptNo AS ReceiptNo
                                               FROM         dbo.Payment INNER JOIN
                                                                     dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
                                               UNION ALL
                                               SELECT     dbo.Payment.Entity_DPA_ AS Entity_DPA_, dbo.Payment.PaymentPDate AS TransDate, dbo.Payment.PaymentReference AS Ref,
                                                                     dbo.Payment.PaymentNarrative AS Particulars,
                                                                     CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                                                                     CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit, 0 AS IsOpeningBalance,
                                                                     PaymentReceiptNo AS ReceiptNo
                                               FROM         dbo.Payment INNER JOIN
                                                                     dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
                                               WHERE     (EntityType_DPA_ = 5) AND (Entity_DPA_ IN
                                                                         (SELECT     Account_DPA_
                                                                           FROM          ACCOUNT
                                                                           WHERE      (Account_DPA_ IN
                                                                                                      (SELECT     Entity_DPA_
                                                                                                        FROM          Payment
                                                                                                        WHERE      EntityType_DPA_ = 5))))
                                               UNION ALL
                                               SELECT     dbo.JournalList.Entity_DPA_, dbo.JournalList.JournalDate AS TransDate,
                                                                     CAST(dbo.JournalList.JournalEntry_DPA_ AS SQL_VARIANT) AS Ref, dbo.JournalList.JournalNarrative AS Particulars,
                                                                     JournalList.JournalEntryDebit AS Debit, JournalList.JournalEntryCredit AS Credit, 0 AS IsOpeningBalance, '' AS ReceiptNo
                                               FROM         dbo.JournalList
                                               WHERE     (EntityType_DPA_ = 5) ) AccountsStatement) a
GROUP BY a.BankAccount_DPA_, a.TransDate, a.Ref, a.Particulars, a.Debit, a.Credit, a.Balance, a.OpeningBalance, a.ReceiptNo
ORDER BY a.BankAccount_DPA_, a.TransDate, a.Ref, a.Particulars, a.Debit, a.Credit, a.Balance, a.OpeningBalance



GO
CREATE VIEW dbo.BatchedApplications
AS
SELECT DISTINCT
                      TOP 100 PERCENT dbo.Offerings.Batch_No, COUNT(dbo.Offerings.Offering_DPA_) AS TotalNo, dbo.Security.SecurityName,
                      SUM(dbo.Offerings.Alloted_Rights) AS TotalQty, SUM(dbo.Offerings.Alloted_Rights * dbo.Offerings.Offering_Price) AS TotalAmt,
                      dbo.Security.ClosingDate, MAX(dbo.Offerings.Offering_DPA_) AS Offering_DPA_
FROM         dbo.Offerings INNER JOIN
                      dbo.Security ON dbo.Offerings.Offering = dbo.Security.Security_DPA_
WHERE     (dbo.Offerings.Forward <> 1) AND (dbo.Offerings.Deleted = 0)
GROUP BY dbo.Offerings.Batch_No, dbo.Security.SecurityName, dbo.Security.ClosingDate
HAVING      (NOT (dbo.Offerings.Batch_No IS NULL))
ORDER BY dbo.Security.ClosingDate DESC, dbo.Offerings.Batch_No DESC

GO
CREATE VIEW dbo.BatchedForwards
AS
SELECT     TOP 100 PERCENT dbo.Offerings.Batch_No, COUNT(dbo.Offerings.Offering_DPA_) AS TotalNo, dbo.Security.SecurityName,
                      SUM(dbo.Offerings.Alloted_Rights) AS TotalQty, SUM(dbo.Offerings.Alloted_Rights * dbo.Offerings.Offering_Price) AS TotalAmt,
                      dbo.Security.ClosingDate, MAX(dbo.Offerings.Offering_DPA_) AS Offering_DPA_, dbo.Offerings.Forward
FROM         dbo.Offerings INNER JOIN
                      dbo.Security ON dbo.Offerings.Offering = dbo.Security.Security_DPA_
WHERE     (dbo.Offerings.Forward = 1)
GROUP BY dbo.Offerings.Batch_No, dbo.Security.SecurityName, dbo.Security.ClosingDate, dbo.Offerings.Forward
HAVING      (NOT (dbo.Offerings.Batch_No IS NULL))
ORDER BY dbo.Security.ClosingDate DESC, dbo.Offerings.Batch_No DESC

GO
CREATE VIEW dbo.BnkBranchList AS SELECT Bank.BankName + ", " + BnkBranch.BnkBranchName AS BnkBranchName, BnkBranch.BnkBranch_DPA_ AS BnkBranch_DPA_
FROM Bank INNER JOIN BnkBranch ON Bank.Bank_DPA_ = BnkBranch.Bank_DPA_

GO
CREATE VIEW dbo.BondClientList
AS
SELECT     dbo.FullClientList2.*
FROM         dbo.FullClientList2

GO

CREATE VIEW dbo.BondConfirmationList
AS
SELECT     TOP 100 PERCENT dbo.tbOrder.Order_DPA_, MIN(dbo.Lots.ContractNumber) AS ContractNumber, SUM(dbo.Lots.LotGrossAmount) AS Gross,
                      CAST(FLOOR(CAST(dbo.Lots.LotTDate AS float)) AS datetime) AS TransDate, MIN(dbo.Lots.LotSlipNo) AS SlipNo, dbo.Security.SecurityName,
                      dbo.Bond.BondIssue, CASE (dbo.tbOrder.OrderType_DPA_) WHEN 2 THEN SUM(dbo.Lots.LotGrossAmount)
                      - a.Commission ELSE SUM(dbo.Lots.LotGrossAmount) + a.Commission END AS NetAmount, a.Commission AS Commission, CONVERT(float,
                      SUM(dbo.Lots.LotGrossAmount)) / CONVERT(float, SUM(dbo.Lots.LotQty)) * 100 AS Price, dbo.Bond.BondIDate AS IssueDate,
                      dbo.Bond.BondMDate AS MaturityDate, '' AS ProposalNo, SUM(dbo.Lots.LotQty) AS FaceValue, SUM(dbo.Lots.LotQty) * (SUM(dbo.Lots.LotGrossAmount)
                       / SUM(dbo.Lots.LotQty) * 100) AS Consideration, Client.ClientName, Client.Client_DPA_, CASE (dbo.tbOrder.OrderType_DPA_)
                      WHEN 2 THEN 'SALE' ELSE 'PURCHASE' END AS Type, BrokerName, BrokerCode, BrokerOfficeTel,
                      MIN(CAST(FLOOR(CAST(dbo.Contracts.ContractSettlementDate AS float)) AS datetime)) AS SettlementDate, Client.ClientContact,
                      MIN(Owner.OwnerFname + '  ' + Owner.OwnerLName) AS ClientOwner, MIN(CAST(FLOOR(CAST(b.SettlementDate AS float)) AS datetime))
                      AS BondSettlementDate, c.Commission AS AgentCommission, MIN(Lots.Contract_DPA_) AS ContractNo
FROM         dbo.Lots INNER JOIN
                      dbo.Contracts ON dbo.Lots.Contract_DPA_ = dbo.Contracts.Contract_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.Lots.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ LEFT OUTER JOIN
                          (SELECT     *
                            FROM          BondConfirmation
                            WHERE      Deleted = 0) b ON tbOrder.Order_DPA_ = b.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.Bond ON dbo.OrdDetail.Bond_DPA_ = dbo.Bond.Bond_DPA_ INNER JOIN
                          (SELECT     SUM(dbo.LevyContract.LevyAmount) AS Commission, dbo.OrdDetail.Order_DPA_
                            FROM          dbo.OrdDetail INNER JOIN
                                                   dbo.Lots ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.Lots.OrdDetail_DPA_ INNER JOIN
                                                   dbo.LevyContract ON dbo.Lots.Contract_DPA_ = dbo.LevyContract.Contract_DPA_
                            WHERE      (LevyContract.SystemMaintained = 11)
                            GROUP BY dbo.OrdDetail.Order_DPA_) a ON dbo.tbOrder.Order_DPA_ = a.Order_DPA_ INNER JOIN
                      Broker ON Lots.Broker_DPA_ = Broker.Broker_DPA_ LEFT OUTER JOIN
                      Owner ON Client.Owner_DPA_ = Owner.Owner_DPA_ INNER JOIN
                          (SELECT     SUM(LevyContract.LevyAmount) AS Commission, OrdDetail.Order_DPA_
                            FROM          OrdDetail INNER JOIN
                                                   Lots ON OrdDetail.OrdDetail_DPA_ = Lots.OrdDetail_DPA_ INNER JOIN
                                                   LevyContract ON Lots.Contract_DPA_ = LevyContract.Contract_DPA_
                            WHERE      (LevyContract.SystemMaintained = 12)
                            GROUP BY OrdDetail.Order_DPA_) c ON dbo.tbOrder.Order_DPA_ = c.Order_DPA_
WHERE     (dbo.Security.OrderSecType_DPA_ = 1) AND (dbo.tbOrder.OrderCompounded = 1) AND (dbo.OrdDetail.Deleted = 0) AND (dbo.tbOrder.Deleted = 0)
GROUP BY Client.ClientName, Client.Client_DPA_, Client.ClientContact, dbo.tbOrder.Order_DPA_, CAST(FLOOR(CAST(dbo.Lots.LotTDate AS float)) AS datetime),
                      dbo.Security.SecurityName, dbo.Bond.BondIssue, a.Commission, c.Commission, dbo.tbOrder.OrderType_DPA_, dbo.Bond.BondIDate,
                      dbo.Bond.BondMDate, BrokerName, BrokerCode, BrokerOfficeTel
ORDER BY CAST(FLOOR(CAST(dbo.Lots.LotTDate AS float)) AS datetime) DESC, ContractNo DESC

GO

CREATE VIEW dbo.BondList
AS
SELECT     TOP 100 PERCENT CONVERT(DATETIME, dbo.Bond.BondIDate, 108) AS BondIDate,
                      dbo.Bond.BondIssue + ', ' + dbo.Security.SecurityName AS BondIssue, dbo.Bond.BondRate AS BondRate, dbo.Bond.Bond_DPA_ AS Bond_DPA_
FROM         dbo.Security INNER JOIN
                      dbo.Bond ON dbo.Security.Security_DPA_ = dbo.Bond.Security_DPA_
ORDER BY dbo.Bond.BondIssue + ', ' + dbo.Security.SecurityName


GO

CREATE VIEW dbo.BondListBond
AS
SELECT     TOP 100 PERCENT dbo.SecurityListFilter.SecurityName, dbo.SecurityListFilter.SecurityCode, dbo.SecurityListFilter.SecurityMktPrice,
                      REPLACE(REPLACE(REPLACE(dbo.SecurityListFilter.SecurityAddr, CHAR(13), ','), CHAR(10), ''), CHAR(34), ' ') AS SecurityAddr,
                      dbo.SecurityListFilter.OrderSecTypeDescription, dbo.SecurityListFilter.Security_DPA_, dbo.SecurityListFilter.OrderSecTypeDisplayName,
                      dbo.SecTransFee.SecTransFeeFee, dbo.SecurityListFilter.OrderSecType_DPA_, dbo.SecurityListFilter.SecurityNameEx,
                      dbo.SecurityListFilter.Immobilised
FROM         dbo.SecurityListFilter INNER JOIN
                      dbo.SecTransFee ON dbo.SecurityListFilter.Security_DPA_ = dbo.SecTransFee.Security_DPA_ AND
                      dbo.SecurityListFilter.MaxOfSecTransFeeADate = dbo.SecTransFee.SecTransFeeADate INNER JOIN
                      dbo.OrderSecType ON dbo.SecurityListFilter.OrderSecType_DPA_ = dbo.OrderSecType.OrderSecType_DPA_
WHERE     (dbo.OrderSecType.OrderSecType_DPA_ = 1)
GROUP BY dbo.SecurityListFilter.SecurityName, dbo.SecurityListFilter.SecurityCode, dbo.SecurityListFilter.SecurityMktPrice, dbo.SecurityListFilter.SecurityAddr,
                      dbo.SecurityListFilter.OrderSecTypeDescription, dbo.SecurityListFilter.Security_DPA_, dbo.SecurityListFilter.OrderSecTypeDisplayName,
                      dbo.SecTransFee.SecTransFeeFee, dbo.SecurityListFilter.OrderSecType_DPA_, dbo.SecurityListFilter.SecurityNameEx,
                      dbo.SecurityListFilter.Immobilised
ORDER BY dbo.SecurityListFilter.SecurityName


GO



CREATE VIEW dbo.BondOffer
AS
SELECT     dbo.BondProposals.Salutation AS Account, dbo.Client.ClientAddr AS AccountAddress, dbo.Client.ClientContact AS Contact,
                      dbo.OwnerList.OwnerName AS Owner, dbo.Client.ClientFax AS Fax, dbo.BondProposals.*, dbo.Client.ClientName AS ClientName
FROM         dbo.BondProposals INNER JOIN
                      dbo.Client ON dbo.BondProposals.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Bond ON dbo.BondProposals.Bond_DPA_ = dbo.Bond.Bond_DPA_ LEFT OUTER JOIN
                      dbo.OwnerList ON dbo.Client.Owner_DPA_ = dbo.OwnerList.Owner_DPA_
WHERE     (dbo.BondProposals.Canceled = 0)




GO


CREATE VIEW dbo.BondOrderList
AS
SELECT     TOP 100 PERCENT dbo.BondProposals.ProposalDate, dbo.BondProposals.Proposal_DPA_, dbo.BondProposals.BondIDate, dbo.tbOrder.Client_DPA_,
                      dbo.tbOrder.Order_DPA_, dbo.tbOrder.OrderHold, dbo.tbOrder.OrderRef, dbo.tbOrder.OrderType_DPA_, dbo.tbOrder.OrderCanceled,
                      dbo.tbOrder.OrderHoldType_DPA_, dbo.tbOrder.OrderCompounded, dbo.tbOrder.OrderDateReleased, dbo.tbOrder.InterBank, dbo.tbOrder.TimeChanged,
                      dbo.tbOrder.OrderDate, dbo.OrdDetailList.OrdDetailPrice, dbo.OrdDetailList.OrdDetailQty, dbo.OrdDetailList.OrdDetailSecType,
                      dbo.OrdDetailList.BalanceQty, dbo.OrdDetailList.OrderSecType_DPA_, dbo.OrdDetailList.OrdDetailSecurity, dbo.OrdDetailList.OrderTypeSale,
                      dbo.OrdDetailList.OrdDetail_DPA_, dbo.OrdDetail.Best, dbo.Client.ClientCDSNo, dbo.Client.ClientName, dbo.tbOrder.OrderReleasedBy,
                      dbo.tbOrder.ChangedBy, dbo.Users.OtherNames + '  ' + dbo.Users.Surname AS UserName, dbo.OrdDetail.BondDescription,
                      dbo.Security.SecurityCode
FROM         dbo.Users RIGHT OUTER JOIN
                      dbo.tbOrder ON dbo.Users.UserID = dbo.tbOrder.ChangedBy LEFT OUTER JOIN
                      dbo.OrdDetail RIGHT OUTER JOIN
                      dbo.OrdDetailList ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.OrdDetailList.OrdDetail_DPA_ ON
                      dbo.tbOrder.Order_DPA_ = dbo.OrdDetailList.Order_DPA_ LEFT OUTER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ RIGHT OUTER JOIN
                      dbo.BondProposals ON dbo.tbOrder.Proposal_DPA_ = dbo.BondProposals.Proposal_DPA_ LEFT OUTER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_
WHERE     (dbo.BondProposals.Canceled <> 1)
ORDER BY dbo.BondProposals.Proposal_DPA_ DESC, dbo.BondProposals.ProposalDate





GO
CREATE VIEW dbo.BondProposalList
AS
SELECT     dbo.BondProposals.*, dbo.Client.ClientName, dbo.OwnerList.OwnerName, dbo.Security.SecurityCode AS SecurityCode,
                      dbo.Security.SecurityName AS SecurityName, dbo.Commission.CommissionDescription AS CommDesc,
                      dbo.Commission.Commission_DPA_ AS Comm_DPA, dbo.UserList.UserName
FROM         dbo.BondProposals INNER JOIN
                      dbo.Client ON dbo.BondProposals.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.OwnerList ON dbo.Client.Owner_DPA_ = dbo.OwnerList.Owner_DPA_ INNER JOIN
                      dbo.Security ON dbo.BondProposals.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.Commission ON dbo.Client.Commission_DPA_ = dbo.Commission.Commission_DPA_ LEFT OUTER JOIN
                      dbo.UserList ON dbo.BondProposals.ModifiedBy = dbo.UserList.UserID

GO
CREATE VIEW dbo.BondsEvaluationList
AS
SELECT     dbo.Contract.Included, dbo.Contract.DaysInCoupon, a.ForwardRate, dbo.Contract.Basis, 'Con' + CAST(dbo.Contract.Contract_DPA_ AS nvarchar)
                      AS Evaluation_DPA_, 'Contract' AS Type, dbo.Lot.ContractNumber AS Reference, dbo.Lot.LotTDate AS TradeDate, dbo.tbOrder.Client_DPA_,
                      dbo.OrdDetail.Bond_DPA_, dbo.Bond.BondIssue, dbo.Bond.BondIDate AS IssueDate, dbo.Client.ClientName, dbo.Bond.BondRate AS CouponRate,
                      dbo.Lot.LotQty AS FaceValue, dbo.Lot.LotPrice AS Bprice, dbo.Bond.BondMDate AS MaturityDate, dbo.Contract.ModifiedBy, dbo.Contract.DateModified,
                      dbo.Users.UserName AS [User], a.ForwardRate_DPA_
FROM         dbo.Contract INNER JOIN
                      dbo.Lot ON dbo.Contract.Contract_DPA_ = dbo.Lot.Contract_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Bond ON dbo.OrdDetail.Bond_DPA_ = dbo.Bond.Bond_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ LEFT OUTER JOIN
                      dbo.Users ON dbo.Contract.ModifiedBy = dbo.Users.UserID LEFT OUTER JOIN
                          (SELECT     TOP 100 PERCENT FRate.ForwardRate_DPA_, FRate.ForwardRate, FRate.ActivationDate, FRate.Bond_DPA_, dbo.Users.UserName,
                                                   FRate.DateModified, FRate.ModifiedBy
                            FROM          dbo.ForwardRate FRate LEFT OUTER JOIN
                                                   dbo.Users ON FRate.ModifiedBy = dbo.Users.UserID
                            WHERE      (FLOOR(CAST(FRate.ActivationDate AS Float)) =
                                                       (SELECT     MAX(Floor(cast(ActivationDate AS Float))) AS UniqueDate
                                                         FROM          dbo.ForwardRate
                                                         WHERE      ActivationDate IS NOT NULL AND Frate.Bond_DPA_ = Bond_DPA_))) a ON dbo.Bond.Bond_DPA_ = a.Bond_DPA_
WHERE     (dbo.Bond.BondMDate >= GETDATE())
UNION
SELECT     (CASE AcceptedAmount WHEN 0 THEN 0 ELSE Included END) AS Included, dbo.PrimaryIssues.DaysInCoupon, a.ForwardRate,
                      dbo.PrimaryIssues.Basis, 'Pry' + CAST(dbo.PrimaryIssues.PrimaryIssues_DPA_ AS nvarchar) AS Evaluation_DPA_, 'Primary Issue' AS Type,
                      dbo.PrimaryIssues.Reference, dbo.PrimaryIssues.PaymentDate AS TradeDate, dbo.PrimaryIssues.Client_DPA_, dbo.Bond.Bond_DPA_,
                      dbo.Bond.BondIssue, dbo.Bond.BondIDate AS IssueDate, dbo.Client.ClientName, dbo.Bond.BondRate AS CouponRate,
                      dbo.PrimaryIssues.AcceptedAmount AS FaceValue, dbo.PrimaryIssues.Price AS Bprice, dbo.Bond.BondMDate AS MaturityDate,
                      dbo.PrimaryIssues.ModifiedBy, dbo.PrimaryIssues.DateModified, dbo.Users.UserName AS [User], a.ForwardRate_DPA_
FROM         dbo.PrimaryIssues INNER JOIN
                      dbo.Client ON dbo.PrimaryIssues.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Bond ON dbo.PrimaryIssues.Bond_DPA_ = dbo.Bond.Bond_DPA_ LEFT OUTER JOIN
                      dbo.Users ON dbo.PrimaryIssues.ModifiedBy = dbo.Users.UserID LEFT OUTER JOIN
                          (SELECT     TOP 100 PERCENT FRate.ForwardRate_DPA_, FRate.ForwardRate, FRate.ActivationDate, FRate.Bond_DPA_, dbo.Users.UserName,
                                                   FRate.DateModified, FRate.ModifiedBy
                            FROM          dbo.ForwardRate FRate LEFT OUTER JOIN
                                                   dbo.Users ON FRate.ModifiedBy = dbo.Users.UserID
                            WHERE      (FLOOR(CAST(FRate.ActivationDate AS Float)) =
                                                       (SELECT     MAX(Floor(cast(ActivationDate AS Float))) AS UniqueDate
                                                         FROM          dbo.ForwardRate
                                                         WHERE      ActivationDate IS NOT NULL AND Frate.Bond_DPA_ = Bond_DPA_))) a ON dbo.Bond.Bond_DPA_ = a.Bond_DPA_
WHERE     (dbo.Bond.BondMDate >= GETDATE()) AND (NOT (dbo.PrimaryIssues.PaymentDate IS NULL)) AND (dbo.PrimaryIssues.Deleted <> 1) OR
                      (dbo.PrimaryIssues.Deleted IS NULL)
UNION
SELECT     dbo.BondTrades.Included, dbo.BondTrades.DaysInCoupon, a.ForwardRate, dbo.BondTrades.Basis,
                      'Ext' + CAST(dbo.BondTrades.BondTrades_DPA_ AS nvarchar) AS Evaluation_DPA_, 'External Trade' AS Type, dbo.BondTrades.Reference,
                      dbo.BondTrades.TradeDate, dbo.BondTrades.Client_DPA_, dbo.BondTrades.Bond_DPA_, dbo.Bond.BondIssue, dbo.Bond.BondIDate AS IssueDate,
                      dbo.Client.ClientName, dbo.Bond.BondRate AS CouponRate, dbo.BondTrades.Quantity AS FaceValue, dbo.BondTrades.Price AS Bprice,
                      dbo.Bond.BondMDate AS MaturityDate, dbo.BondTrades.ModifiedBy, dbo.BondTrades.DateModified, dbo.Users.UserName AS [User],
                      a.ForwardRate_DPA_
FROM         dbo.BondTrades INNER JOIN
                      dbo.Client ON dbo.BondTrades.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Bond ON dbo.BondTrades.Bond_DPA_ = dbo.Bond.Bond_DPA_ LEFT OUTER JOIN
                      dbo.Users ON dbo.BondTrades.ModifiedBy = dbo.Users.UserID LEFT OUTER JOIN
                          (SELECT     TOP 100 PERCENT FRate.ForwardRate_DPA_, FRate.ForwardRate, FRate.ActivationDate, FRate.Bond_DPA_, dbo.Users.UserName,
                                                   FRate.DateModified, FRate.ModifiedBy
                            FROM          dbo.ForwardRate FRate LEFT OUTER JOIN
                                                   dbo.Users ON FRate.ModifiedBy = dbo.Users.UserID
                            WHERE      (FLOOR(CAST(FRate.ActivationDate AS Float)) =
                                                       (SELECT     MAX(Floor(cast(ActivationDate AS Float))) AS UniqueDate
                                                         FROM          dbo.ForwardRate
                                                         WHERE      ActivationDate IS NOT NULL AND Frate.Bond_DPA_ = Bond_DPA_))) a ON dbo.Bond.Bond_DPA_ = a.Bond_DPA_
WHERE     (dbo.Bond.BondMDate >= GETDATE()) AND (dbo.BondTrades.Deleted <> 1) OR
                      (dbo.BondTrades.Deleted IS NULL)

GO


CREATE VIEW dbo.BondsTradeList
AS
SELECT     dbo.BondTrades.*, dbo.Security.SecurityCode AS BondType, dbo.Client.ClientName AS ClientName, dbo.Bond.BondIssue AS BondIssue,
                      dbo.Bond.Security_DPA_ AS Security_DPA_, dbo.OrderType.OrderTypeDescription AS TradeType, dbo.Users.UserName AS [User]
FROM         dbo.BondTrades INNER JOIN
                      dbo.Client ON dbo.BondTrades.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Bond ON dbo.BondTrades.Bond_DPA_ = dbo.Bond.Bond_DPA_ INNER JOIN
                      dbo.Security ON dbo.Bond.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.BondTrades.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ LEFT OUTER JOIN
                      dbo.Users ON dbo.BondTrades.ModifiedBy = dbo.Users.UserID
WHERE     (dbo.BondTrades.Deleted <> 1) OR
                      (dbo.BondTrades.Deleted IS NULL)



GO
CREATE VIEW dbo.BranchList
AS
SELECT     TOP 100 PERCENT BranchDescription AS BranchName, Branch_DPA_, DefaultSelection
FROM         dbo.Branch
ORDER BY BranchDescription

GO

CREATE VIEW dbo.BrokerBalanceList

AS

SELECT  TOP 100 PERCENT    dbo.Broker.BrokerName AS Broker, CAST(dbo.Broker.BrokerAddr AS NVARCHAR(1000)) As Address,
	(SELECT TOP 1 Balance FROM dbo.BrokerStatement WHERE Broker_DPA_ = a.Broker_DPA_ ORDER BY TransDate DESC) AS Balance,
	dbo.Broker.BrokerOfficeTel AS Telephone
FROM         dbo.BrokerStatement a INNER JOIN dbo.Broker ON dbo.Broker.Broker_DPA_ = a.Broker_DPA_
GROUP BY  a.Broker_DPA_, dbo.Broker.BrokerName, CAST(dbo.Broker.BrokerAddr AS NVARCHAR(1000)),
		dbo.Broker.BrokerOfficeTel
ORDER BY  dbo.Broker.BrokerName






GO
CREATE VIEW dbo.BrokerCommissionList
AS
SELECT     TOP 100 PERCENT dbo.Security.SecurityCode, dbo.Lots.LotTDate, dbo.Lots.ContractNumber, dbo.Lots.LotQty, dbo.Lots.LotPrice,
                      dbo.OrdDetail.Order_DPA_, dbo.LevyContract.LevyAmount, dbo.Users.Surname AS ChangedBy, dbo.LevyContract.LevyRate,
                      dbo.Contracts.Contract_DPA_, dbo.Client.ClientName, dbo.Lots.LotGrossAmount, dbo.LevyContract.TimeChanged, dbo.Broker.Broker_DPA_,
                      dbo.tbOrder.OrderType_DPA_, CAST(LotQty AS nvarchar) + ' ' + CASE WHEN len(Security.SecurityName) > 20 THEN LEFT(Security.SecurityName, 20)
                      + '...' ELSE Security.SecurityName END + ' @ ' + CAST(LotPrice AS nvarchar) AS SecurityName
FROM         dbo.Lots INNER JOIN
                      dbo.Contracts ON dbo.Lots.Contract_DPA_ = dbo.Contracts.Contract_DPA_ INNER JOIN
                      dbo.LevyContract ON dbo.Lots.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.Lots.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Broker ON dbo.Lots.Broker_DPA_ = dbo.Broker.Broker_DPA_ LEFT OUTER JOIN
                      dbo.Users ON dbo.LevyContract.ChangedBy = dbo.Users.UserID
WHERE     (dbo.LevyContract.LevyShortName LIKE N'%Commission%')
ORDER BY dbo.Lots.LotTDate DESC, dbo.Contracts.Contract_DPA_ DESC




GO
CREATE VIEW dbo.BrokerCommissionStatement
AS
SELECT     Entity_DPA_, TransDate, Ref, Particulars, EntityName, Credit, Debit, CreditBal AS Balance, 1 AS IsOpeningBalance
FROM         (SELECT     dbo.Entity.Entity_DPA_, CAST('3/1/2005' AS Datetime) AS TransDate, '' AS Ref, 'Opening Balance' AS Particulars, dbo.Entity.EntityName,
                                              0 AS Credit, 0 AS Debit, dbo.Entity.EntityOpeningBal AS CreditBal
                       FROM          dbo.Entity INNER JOIN
                                              dbo.EntityType ON dbo.Entity.EntityType_DPA_ = dbo.EntityType.EntityType_DPA_
                       WHERE      (dbo.Entity.EntityType_DPA_ = 4)) AS a
UNION
SELECT     TOP 100 PERCENT Entity_2.Entity_DPA_, CAST(FLOOR(CAST(dbo.Lots.LotTDate AS float)) AS datetime) AS TransDate, dbo.Lots.ContractNumber AS REF,
                       CAST(dbo.Lots.LotQty AS nvarchar) + ' ' + dbo.Security.SecurityCode + ' @ ' + CAST(dbo.Lots.LotPrice AS nvarchar) AS Particulars,
                      Entity_2.EntityName, dbo.LevyContract.LevyAmount AS Credit, 0 AS Debit, dbo.LevyContract.LevyAmount AS Balance, 0 AS IsOpeningBalance
FROM         dbo.LevyContract INNER JOIN
                      dbo.Contract ON dbo.LevyContract.Contract_DPA_ = dbo.Contract.Contract_DPA_ INNER JOIN
                      dbo.Lots ON dbo.Contract.Contract_DPA_ = dbo.Lots.Contract_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.Lots.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.Entity AS Entity_2 ON dbo.LevyContract.LevyShortName = Entity_2.EntityName
WHERE     (Entity_2.Entity_DPA_ = 4)
UNION
SELECT     TOP 100 PERCENT Entity_DPA_ AS Client_DPA_, CAST(FLOOR(CAST(JournalDate AS float)) AS datetime) AS TransDate,
                      CAST(JournalEntry_DPA_ AS Nvarchar(500)) AS Expr1, JournalNarrative AS Particulars, '' AS Expr2, JournalEntryCredit AS Credit,
                      JournalEntryDebit AS Debit, JournalEntryCredit - JournalEntryDebit AS CreditBal, 0 AS IsOpeningBalance
FROM         dbo.JournalList
WHERE     (EntityType_DPA_ = 4)
UNION
SELECT     TOP 100 PERCENT 4 AS Entity_DPA_, CAST(FLOOR(CAST(Lots_1.LotTDate AS float)) AS datetime) AS TransDate, Lots_1.ContractNumber AS REF,
                      'Agent : ' + CAST(Lots_1.LotQty AS nvarchar) + ' ' + Security_1.SecurityCode + ' @ ' + CAST(Lots_1.LotPrice AS nvarchar) AS Particulars,
                      'Agent Commission' AS Expr1, 0 AS Credit, LevyContract_1.LevyAmount AS Debit, 0 - LevyContract_1.LevyAmount AS Balance,
                      0 AS IsOpeningBalance
FROM         dbo.LevyContract AS LevyContract_1 INNER JOIN
                      dbo.Contract AS Contract_1 ON LevyContract_1.Contract_DPA_ = Contract_1.Contract_DPA_ INNER JOIN
                      dbo.Lots AS Lots_1 ON Contract_1.Contract_DPA_ = Lots_1.Contract_DPA_ INNER JOIN
                      dbo.OrdDetail AS OrdDetail_1 ON Lots_1.OrdDetail_DPA_ = OrdDetail_1.OrdDetail_DPA_ INNER JOIN
                      dbo.Security AS Security_1 ON OrdDetail_1.Security_DPA_ = Security_1.Security_DPA_ INNER JOIN
                      dbo.tbOrder ON OrdDetail_1.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_
WHERE     (LevyContract_1.SystemMaintained = 12) AND (dbo.Client.Owner_DPA_ IS NULL) OR
                      (LevyContract_1.SystemMaintained = 12) AND (dbo.Client.Owner_DPA_ = 8)
UNION
SELECT     TOP 100 PERCENT Entity_DPA_, TransDate, Ref, Particulars, entityname, Credit, Debit, Credit - Debit AS balance, 0 AS IsOpeningBalance
FROM         (SELECT     dbo.Payment.Entity_DPA_, CAST(FLOOR(CAST(dbo.Payment.PaymentPDate AS float)) AS datetime) AS TransDate,
                                              dbo.Payment.PaymentReference AS Ref, dbo.Payment.PaymentNarrative AS Particulars, '' AS entityname,
                                              CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit,
                                              CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit
                       FROM          dbo.Payment INNER JOIN
                                              dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_ INNER JOIN
                                              dbo.PaymentClientStatementNos ON dbo.Payment.Payment_DPA_ = dbo.PaymentClientStatementNos.Payment_DPA_
                       WHERE      (dbo.Payment.EntityType_DPA_ = 4) AND (dbo.Payment.Deleted = 0)) AS a_1

GO
CREATE VIEW dbo.BrokerEntityList
AS
SELECT     dbo.Broker.BrokerName AS EntityName, dbo.EntityType.EntityTypeName AS EntityType, dbo.Broker.Broker_DPA_ AS Entity_DPA_,
                      dbo.EntityType.EntityType_DPA_, dbo.Broker.BrokerCode AS EntityCode, dbo.Broker.BrokerCode + SPACE
                          ((SELECT     ISNULL(MAX(LEN(BrokerCode)), 0) AS MAXLEN
                              FROM         dbo.Broker) - LEN(dbo.Broker.BrokerCode)) + ' : ' + dbo.Broker.BrokerName AS EntityNameEx
FROM         dbo.Broker INNER JOIN
                      dbo.EntityType ON dbo.Broker.EntityType_DPA_ = dbo.EntityType.EntityType_DPA_

GO

CREATE VIEW dbo.BrokerList
AS
SELECT     TOP 100 PERCENT BrokerName, Broker_DPA_, BrokerOfficeTel, BrokerFax, REPLACE(REPLACE(BrokerAddr, CHAR(13), ','), CHAR(10), '') AS BrokerAddr,
                       BrokerCode, BrokerCode + SPACE
                          ((SELECT     ISNULL(MAX(LEN(BrokerCode)), 0) AS MAXLEN
                              FROM         dbo.Broker) - LEN(BrokerCode)) + ' : ' + BrokerName AS BrokerNameEx
FROM         dbo.Broker
ORDER BY BrokerName


GO
CREATE VIEW dbo.BrokerReport
AS

SELECT     TOP 100 PERCENT BrokerName AS Broker, BrokerOfficeTel AS [Office Phone], BrokerFax AS [Fax], REPLACE(REPLACE(BrokerAddr, CHAR(13), ','), CHAR(10), '') AS Address,
                       BrokerCode AS [Code]
FROM         dbo.Broker
ORDER BY BrokerName


GO

CREATE   VIEW app.Brokers AS
SELECT  Broker_DPA_        AS BrokerId,
        BrokerCode         AS Code,
        BrokerName         AS Name,
        BrokerOpeningBal   AS OpeningBalance,
        BrokerRegDate      AS RegisteredOn
FROM dbo.Broker;     -- no Deleted column

GO

CREATE VIEW dbo.BrokerStatement
AS

SELECT     TOP 100 PERCENT COUNT(*) AS ClientTransaction_DPA_,
		a.Broker_DPA_,
		a.TransDate, a.Ref,
		a.Particulars,
		a.Debit,
		a.Credit,
		CASE WHEN (SUM(b.Balance) >= 0) Then CONVERT(NVARCHAR, SUM(b.Balance)) + ' Cr' ELSE CONVERT(NVARCHAR,  ABS(SUM(b.Balance))) + ' Dr' END   AS Balance,
                a.IsOpeningBalance
FROM         (SELECT * FROM dbo.BrokerTransactionList) a CROSS JOIN
                          (SELECT * FROM dbo.BrokerTransactionList) b
WHERE a.TransDate >= b.TransDate AND a.Broker_DPA_ = b.Broker_DPA_
GROUP BY a.Broker_DPA_, a.TransDate, a.Ref, a.Particulars, a.Debit, a.Credit, a.Balance, a.IsOpeningBalance
ORDER BY a.TransDate, a.Broker_DPA_, a.Particulars, a.Ref, a.Debit, a.Credit, a.Balance, a.IsOpeningBalance





GO

CREATE VIEW dbo.BrokerStatement1
AS

SELECT     TOP 100 PERCENT COUNT(*) AS ClientTransaction_DPA_,
		a.Broker_DPA_,
		a.TransDate, a.Ref,
		a.Particulars,
		a.Debit,
		a.Credit,
		 CASE WHEN SUM(b.Balance) >= 0 THEN SUM(b.Balance) ELSE ABS(SUM(b.Balance)) END AS Balance,
                a.IsOpeningBalance
FROM         (SELECT * FROM dbo.BrokerTransactionList1) a CROSS JOIN
                          (SELECT * FROM dbo.BrokerTransactionList1) b
WHERE a.TransDate >= b.TransDate AND a.Broker_DPA_ = b.Broker_DPA_
GROUP BY a.Broker_DPA_, a.TransDate, a.Ref, a.Particulars, a.Debit, a.Credit, a.Balance, a.IsOpeningBalance
ORDER BY a.TransDate, a.Broker_DPA_, a.Particulars, a.Ref, a.Debit, a.Credit, a.Balance, a.IsOpeningBalance








GO

CREATE VIEW dbo.BrokerTransactionList
AS

SELECT   BrokerTransactions.*, Credit - Debit AS Balance
FROM         (SELECT     Broker_DPA_, BrokerRegDate AS TransDate, '' AS REF, ' Opening Balance' AS Particulars, case when BrokerOpeningBal < 0 then BrokerOpeningBal else 0 end AS Debit, case when BrokerOpeningBal >= 0 then BrokerOpeningBal else 0 end AS Credit,
                                              1 AS IsOpeningBalance
                       FROM          dbo.Broker
                       UNION ALL
                       SELECT     dbo.Payments.Entity_DPA_ AS Broker_DPA_, dbo.Payments.PaymentPDate AS TransDate, dbo.Payments.PaymentReference AS Ref,
                                             dbo.Payments.PaymentNarrative AS Particulars, CASE PayType.PayTypeIn WHEN 0 THEN Payments.PaymentAmount ELSE 0 END AS Debit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payments.PaymentAmount ELSE 0 END AS Credit, 0 AS IsOpeningBalance
                       FROM         dbo.Payments INNER JOIN
                                             dbo.PayType ON dbo.Payments.PayType_DPA_ = dbo.PayType.PayType_DPA_
                       WHERE     EntityType_DPA_ = 3

			UNION ALL
		       SELECT     dbo.JournalList.Entity_DPA_ AS Broker_DPA_, dbo.JournalList.JournalDate AS TransDate, CAST(dbo.JournalList.JournalEntry_DPA_ AS NVARCHAR(500))AS Ref,
		                             dbo.JournalList.JournalNarrative AS Particulars, JournalList.JournalEntryDebit AS Debit,
		                             JournalList.JournalEntryCredit AS Credit, 0 AS IsOpeningBalance
		       FROM         dbo.JournalList
		       WHERE     EntityType_DPA_ = 3) BrokerTransactions






GO

CREATE VIEW dbo.BrokerTransactionList1
AS

SELECT TOP 100 PERCENT     BrokerTransactions.*, Credit - Debit AS Balance
FROM         (SELECT     Broker_DPA_, BrokerRegDate AS TransDate, '' AS REF, ' Opening Balance' AS Particulars,
                                              CASE WHEN BrokerOpeningBal < 0 THEN BrokerOpeningBal ELSE 0 END AS Debit,
                                              CASE WHEN BrokerOpeningBal >= 0 THEN BrokerOpeningBal ELSE 0 END AS Credit, 1 AS IsOpeningBalance
                       FROM          dbo.Broker
                       UNION ALL
                       SELECT     dbo.Payments.Entity_DPA_ AS Broker_DPA_, dbo.Payments.PaymentPDate AS TransDate, dbo.Payments.PaymentReference AS Ref,
		         CASE PayType.PayTypeIn WHEN 0 THEN 'PAY: ' + dbo.Payments.PaymentNarrative ELSE 'REC: ' + dbo.Payments.PaymentNarrative END AS Particulars,
			CASE PayType.PayTypeIn WHEN 0 THEN Payments.PaymentAmount ELSE 0 END AS Debit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payments.PaymentAmount ELSE 0 END AS Credit, 0 AS IsOpeningBalance
                       FROM         dbo.Payments INNER JOIN
                                             dbo.PayType ON dbo.Payments.PayType_DPA_ = dbo.PayType.PayType_DPA_
                       WHERE     EntityType_DPA_ = 3
                       UNION ALL
                       SELECT     dbo.JournalList.Entity_DPA_ AS Broker_DPA_, dbo.JournalList.JournalDate AS TransDate,
                                             CAST(dbo.JournalList.JournalEntry_DPA_ AS NVARCHAR(500)) AS Ref, dbo.JournalList.JournalNarrative AS Particulars,
                                             JournalList.JournalEntryDebit AS Debit, JournalList.JournalEntryCredit AS Credit, 0 AS IsOpeningBalance
                       FROM         dbo.JournalList
                       WHERE     EntityType_DPA_ = 3
                       UNION ALL
                      SELECT     Broker_DPA_, LotTDate AS TransDate, ContractNumber AS Ref, SecurityName AS Particulars,
                                             CASE OrderType_DPA_ WHEN 2 THEN isnull(LotGrossAmount, 0) ELSE 0 END AS Debit,
                                             CASE OrderType_DPA_ WHEN 1 THEN isnull(LotGrossAmount, 0) ELSE 0 END AS Credit, 0 AS IsOpeningBalance
                       FROM         dbo.BrokerCommissionList) BrokerTransactions
ORDER BY BrokerTransactions.TransDate DESC










GO

CREATE VIEW dbo.BrokerTransactionsSubList
AS


SELECT CAST(LotQty AS nvarchar) + ' ' + SecurityName + ' @ ' + CAST(LotPrice AS nvarchar) AS SecurityName, ContractNumber
	, LotTDate, Broker_DPA_, LotGrossAmount As Gross, OrderTypeSale
FROM
(SELECT     dbo.Lots.ContractNumber, dbo.Lots.LotTDate, dbo.Lots.Broker_DPA_, dbo.Lots.LotGrossAmount, dbo.OrderType.OrderTypeSale,
                      dbo.Lots.LotQty, dbo.Lots.LotPrice, dbo.Security.SecurityName
FROM         dbo.Lots INNER JOIN
                      dbo.LevyContract ON dbo.Lots.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.Lots.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_
GROUP BY dbo.Lots.ContractNumber, dbo.Lots.LotTDate, dbo.Lots.Broker_DPA_, dbo.OrderType.OrderTypeSale, dbo.Lots.LotGrossAmount, dbo.Lots.LotQty,
                      dbo.Lots.LotPrice, dbo.Security.SecurityName) innerTable


GO

CREATE VIEW dbo.BrokerTransactionsSubList1
AS


SELECT     CAST(LotQty AS nvarchar) + ' ' + SecurityName + ' @ ' + CAST(LotPrice AS nvarchar) AS SecurityName, ContractNumber, LotTDate, Broker_DPA_,
                      LotGrossAmount AS Gross, OrderTypeSale
FROM         (SELECT     dbo.Lots.ContractNumber, dbo.Lots.LotTDate, dbo.Lots.Broker_DPA_, dbo.Lots.LotGrossAmount, dbo.OrderType.OrderTypeSale, dbo.Lots.LotQty,
                                              dbo.Lots.LotPrice, dbo.Security.SecurityCode AS SecurityName
                       FROM          dbo.Lots INNER JOIN
                                              dbo.LevyContract ON dbo.Lots.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                                              dbo.OrdDetail ON dbo.Lots.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                                              dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                                              dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                                              dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_
                       GROUP BY dbo.Lots.ContractNumber, dbo.Lots.LotTDate, dbo.Lots.Broker_DPA_, dbo.OrderType.OrderTypeSale, dbo.Lots.LotGrossAmount, dbo.Lots.LotQty,
                                              dbo.Lots.LotPrice, dbo.Security.Securitycode) innerTable





GO
CREATE VIEW dbo.BrokerVolumeTotals
AS
SELECT     dbo.Broker.BrokerCode AS [Broker Code ], dbo.Broker.BrokerName AS [Broker Name ], SUM(dbo.Lot.LotQty) AS Volume,
                      SUM(dbo.LevyContract.LevyAmount) AS Commission
FROM         dbo.Lot INNER JOIN
                      dbo.LevyContract ON dbo.Lot.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.Broker ON dbo.Lot.Broker_DPA_ = dbo.Broker.Broker_DPA_
GROUP BY dbo.Broker.BrokerCode, dbo.Broker.BrokerName

GO

CREATE VIEW dbo.BudgetReport
AS
SELECT     dbo.AccountTypeLevel2.AccountType_DPA_, dbo.AccountTypeLevel2.AccountTypeName, dbo.AccountTypeLevel2.DefaultSelection,
                      dbo.AccountTypeLevel2.AccountTypeParent, dbo.AccountTypeLevel2.Quarter1, dbo.AccountTypeLevel2.Quarter2, dbo.AccountTypeLevel2.Quarter3,
                      dbo.AccountTypeLevel2.Quarter4, dbo.AccountType.AccountTypeName AS AccountTypeParentName,
                      dbo.AccountType.AccountType_DPA_ AS AccountTypeParent_DPA_
FROM         dbo.AccountTypeLevel2 INNER JOIN
                      dbo.AccountType ON dbo.AccountTypeLevel2.AccountTypeParent = dbo.AccountType.AccountType_DPA_


GO
CREATE VIEW dbo.CanceledOrderList
AS
SELECT     TOP 100 PERCENT dbo.tbOrder.Order_DPA_ AS [Order No], CONVERT(DATETIME, dbo.tbOrder.OrderDate, 108) AS [Order Date],
                      dbo.ClientList.ClientName + '	[' + CONVERT(nvarchar(4000), dbo.tbOrder.Client_DPA_) + ']' AS Client, dbo.OrdDetailList.OrdDetailSecurity AS Security,
                      dbo.OrdDetailList.OrdDetailPrice AS [Order Price], dbo.OrdDetailList.OrdDetailQty AS [Order Qty]
FROM         dbo.tbOrder INNER JOIN
                      dbo.ClientList ON dbo.tbOrder.Client_DPA_ = dbo.ClientList.Client_DPA_ INNER JOIN
                      dbo.OrderTypeList ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderTypeList.OrderType_DPA_ INNER JOIN
                      dbo.OrdDetailList ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetailList.Order_DPA_ INNER JOIN
                      dbo.OrderHoldType ON dbo.tbOrder.OrderHoldType_DPA_ = dbo.OrderHoldType.OrderHoldType_DPA_ LEFT OUTER JOIN
                      dbo.UserList ON dbo.tbOrder.OrderReleasedBy = dbo.UserList.UserID
WHERE     (dbo.tbOrder.OrderCanceled = 1)
ORDER BY dbo.tbOrder.Order_DPA_ DESC, dbo.OrdDetailList.OrdDetailSecurity

GO

CREATE VIEW dbo.CBKBondConfirmationList
AS
SELECT     TOP 100 PERCENT dbo.tbOrder.Order_DPA_, CAST(FLOOR(CAST(dbo.Lots.LotTDate AS float)) AS datetime) AS TransDate, Lots.LotSlipNo AS SlipNo,
                      (CONVERT(float, Lots.LotGrossAmount) / CONVERT(float, dbo.Lots.LotQty)) * 100 AS Price, dbo.Bond.BondIDate AS IssueDate,
                      dbo.Bond.BondMDate AS MaturityDate, Lots.LotQty AS FaceValue, Client.ClientName, Client.Client_DPA_, CASE (dbo.tbOrder.OrderType_DPA_)
                      WHEN 2 THEN 'SALE' ELSE 'PURCHASE' END AS Type, BrokerName, BrokerCode, BrokerOfficeTel,
                      CAST(FLOOR(CAST(dbo.Contracts.ContractSettlementDate AS float)) AS datetime) AS SettlementDate, Client.ClientContact,
                      Owner.OwnerFname + '  ' + Owner.OwnerLName AS ClientOwner, Bond.BondIssue
FROM         dbo.Lots INNER JOIN
                      dbo.Contracts ON dbo.Lots.Contract_DPA_ = dbo.Contracts.Contract_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.Lots.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.Bond ON dbo.OrdDetail.Bond_DPA_ = dbo.Bond.Bond_DPA_ INNER JOIN
                          (SELECT     SUM(dbo.LevyContract.LevyAmount) AS Commission, dbo.OrdDetail.Order_DPA_
                            FROM          dbo.OrdDetail INNER JOIN
                                                   dbo.Lots ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.Lots.OrdDetail_DPA_ INNER JOIN
                                                   dbo.LevyContract ON dbo.Lots.Contract_DPA_ = dbo.LevyContract.Contract_DPA_
                            GROUP BY dbo.OrdDetail.Order_DPA_) a ON dbo.tbOrder.Order_DPA_ = a.Order_DPA_ INNER JOIN
                      Broker ON Lots.Broker_DPA_ = Broker.Broker_DPA_ INNER JOIN
                      Owner ON Client.Owner_DPA_ = Owner.Owner_DPA_
WHERE     (dbo.Security.OrderSecType_DPA_ = 1) AND (dbo.tbOrder.OrderCompounded = 1) AND (dbo.OrdDetail.Deleted = 0) AND (dbo.tbOrder.Deleted = 0)
ORDER BY CAST(FLOOR(CAST(dbo.Lots.LotTDate AS float)) AS datetime) DESC, dbo.Bond.BondIssue


GO

CREATE VIEW dbo.CDASettlement
AS
SELECT     TOP 100 PERCENT dbo.Lots.LotGrossAmount, dbo.Lots.LotSlipNo, dbo.OrderType.OrderTypeDescription, dbo.Client.Client_DPA_, dbo.Client.ClientName,
                      dbo.tbOrder.OrderType_DPA_, dbo.Security.SecurityCode, dbo.Lots.LotPrice, dbo.Lots.LotQty, dbo.Lots.LotTDate, dbo.Lots.ContractNumber,
                      dbo.Client.IsCustodian, dbo.Security.OrderSecType_DPA_
FROM         dbo.tbOrder INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                      dbo.Contract INNER JOIN
                      dbo.Lots ON dbo.Contract.Contract_DPA_ = dbo.Lots.Contract_DPA_ ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.Lots.OrdDetail_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_
WHERE     (dbo.Security.OrderSecType_DPA_ = 2)
ORDER BY dbo.Client.IsCustodian, dbo.tbOrder.OrderType_DPA_ DESC, dbo.Lots.LotSlipNo



GO

CREATE VIEW dbo.CDSControlStatement
AS

SELECT     a.Entity_DPA_ AS Client_DPA_, a.TransDate, a.Ref, a.Particulars, a.Credit, a.Debit, a.CreditBal AS Balance, 1 AS IsOpeningBalance
FROM         (SELECT     dbo.Entity.Entity_DPA_, CAST('3/1/2005' AS Datetime) AS TransDate, '' AS Ref, 'Opening Balance' AS Particulars, 0 AS Credit, 0 AS Debit,
                                              dbo.Entity.EntityOpeningBal AS CreditBal
                       FROM          dbo.Entity INNER JOIN
                                              dbo.EntityType ON dbo.Entity.EntityType_DPA_ = dbo.EntityType.EntityType_DPA_
                       WHERE      (dbo.Entity.EntityType_DPA_ = 8)) a
UNION
SELECT     a.Client_DPA_, a.TransDate, a.Ref, a.Particulars, a.Credit, a.Debit, a.Credit - a.Debit AS Balance, 0 AS IsOpeningBalance
FROM         (SELECT     14 AS Client_DPA_, cast(floor(cast(LotTDate AS float)) AS Datetime) AS TransDate, ContractNumber AS Ref, SecurityName AS Particulars,
                                              CASE OrderTypeSale WHEN 0 THEN Gross ELSE 0 END AS Credit, CASE OrderTypeSale WHEN 1 THEN Gross ELSE 0 END AS Debit,
                                              0 AS Balance, 0 AS IsOpeningBalance
                       FROM          dbo.ClientTransactionsSubList1) a
UNION
SELECT     a.Client_DPA_, a.TransDate, a.Ref, a.Particulars, a.Credit, a.Debit, a.Credit - a.Debit AS Balance, 0 AS IsOpeningBalance
FROM         (SELECT     14 AS Client_DPA_, CAST(FLOOR(CAST(InterTransfer.TransferDate AS Float)) AS DateTime) AS TransDate, Lots.ContractNumber AS REF,
                                              ISNULL(InterTransfer.TransferNarrative, '') + '  ' + InterTransferType.TypeDescription AS Particulars, CASE (OrderType.OrderTypeSale)
                                              WHEN 1 THEN InterTransfer.TransferAmount ELSE 0 END AS Credit, CASE (OrderType.OrderTypeSale)
                                              WHEN 0 THEN InterTransfer.TransferAmount ELSE 0 END AS Debit, 0 AS Balance, 0 AS IsOpeningBalance
                       FROM          OrdDetail INNER JOIN
                                              Contracts INNER JOIN
                                              InterTransfer ON Contracts.Contract_DPA_ = InterTransfer.Contract_DPA_ INNER JOIN
                                              InterTransfers ON Contracts.Contract_DPA_ = InterTransfers.Contract_DPA_ INNER JOIN
                                              InterTransferType ON InterTransfer.InterTransferType_DPA_ = InterTransferType.InterTransferType_DPA_ INNER JOIN
                                              Lots ON Contracts.Contract_DPA_ = Lots.Contract_DPA_ ON OrdDetail.OrdDetail_DPA_ = Lots.OrdDetail_DPA_ INNER JOIN
                                              OrderType INNER JOIN
                                              tbOrder ON OrderType.OrderType_DPA_ = tbOrder.OrderType_DPA_ ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_
                       WHERE      (InterTransfer.Deleted = 0) AND (Contracts.Deleted = 0) AND (Lots.Deleted = 0) AND (tbOrder.Deleted = 0) AND (OrdDetail.Deleted = 0))
                      a
UNION
SELECT     a.Client_DPA_, a.TransDate, a.Ref, a.Particulars, a.Credit, a.Debit, a.Credit - a.Debit AS Balance, 0 AS IsOpeningBalance
FROM         (SELECT     dbo.JournalList.Entity_DPA_ AS Client_DPA_, cast(floor(cast(dbo.JournalList.JournalDate AS float)) AS DateTime) AS TransDate,
                                              CAST(dbo.JournalList.JournalEntry_DPA_ AS Nvarchar(500)) AS Ref, dbo.JournalList.JournalNarrative AS Particulars,
                                              JournalList.JournalEntryCredit AS Credit, JournalList.JournalEntryDebit AS Debit, 0 AS Balance, 0 AS IsOpeningBalance
                       FROM          dbo.JournalList
                       WHERE      EntityType_DPA_ = 8) a
UNION
SELECT     a.Entity_DPA_ AS Client_DPA_, a.TransDate, a.Ref, a.Particulars, a.Credit, a.Debit, a.Credit - a.Debit AS Balance, 0 AS IsOpeningBalance
FROM         (SELECT     dbo.Payments.Entity_DPA_, cast(floor(cast(dbo.Payments.PaymentPDate AS float)) AS datetime) AS TransDate,
                                              ISNULL(Lots.ContractNumber, '') + ' ' + Payments.PaymentReference AS Ref,
                                              UPPER(CASE PayType.PayTypeIn WHEN 1 THEN 'RECEIPT:' ELSE 'PAID OUT:' END + ISNULL(Security.SecurityCode + ' @ ' + CAST(Lots.LotPrice
                                               AS nvarchar), '') + ' ' + Payments.PaymentNarrative) AS Particulars,
                                              CASE PayType.PayTypeIn WHEN 1 THEN Payments.PaymentAmount ELSE 0 END AS Credit,
                                              CASE PayType.PayTypeIn WHEN 0 THEN Payments.PaymentAmount ELSE 0 END AS Debit,
                                              CASE PayType.PayTypeIn WHEN 1 THEN Payments.PaymentAmount ELSE 0 END AS CreditBal, 0 AS OpeningBalance
                       FROM          OrdDetail INNER JOIN
                                              Lots ON OrdDetail.OrdDetail_DPA_ = Lots.OrdDetail_DPA_ INNER JOIN
                                              Security ON OrdDetail.Security_DPA_ = Security.Security_DPA_ RIGHT OUTER JOIN
                                              Payments INNER JOIN
                                              PayType ON Payments.PayType_DPA_ = PayType.PayType_DPA_ ON Lots.Contract_DPA_ = Payments.Contract_DPA_
                       WHERE      EntityType_DPA_ = 8 AND Entity_DPA_ = 14) a




















GO
CREATE VIEW dbo.CDSMatchedHoldings
AS
SELECT     dbo.Client.Client_DPA_ AS ClientCode, dbo.Security.SecurityCode, dbo.Security.ImportCode, dbo._CDS_Imported_Holdings_.*
FROM         dbo.Client INNER JOIN
                      dbo._CDS_Imported_Holdings_ ON dbo.Client.ClientCDSNo = dbo._CDS_Imported_Holdings_.CDSNo INNER JOIN
                      dbo.Security ON dbo._CDS_Imported_Holdings_.SecurityImportCode = dbo.Security.ImportCode
WHERE     (dbo._CDS_Imported_Holdings_.Imported <> 1)

GO

CREATE VIEW dbo.CDSMatchedTradesList
AS
SELECT     TOP 100 PERCENT dbo._CDS_Imported_Trades_.CDSImport_DPA_, dbo.OrdDetailList.OrdDetail_DPA_, dbo.OrdDetailList.Order_DPA_,
                      dbo.OrdDetailList.OrderDate, dbo.OrdDetailList.OrderTypeSale, dbo.OrdDetailList.CDSOrderTypeSale, dbo.OrdDetailList.OrdDetailClient,
                      dbo.OrdDetailList.SecurityCode, dbo.OrdDetailList.BalanceQty, dbo._CDS_Imported_Trades_.CDSRef, dbo._CDS_Imported_Trades_.Quantity,
                      dbo._CDS_Imported_Trades_.Price, REPLACE(dbo._CDS_Imported_Trades_.ContraBrokerID, 'B', '') AS BrokerCode,
                      dbo._CDS_Imported_Trades_.ContraBrokerID AS CDSBrokerCode, dbo.OrdDetailList.OrdDetailType, dbo.OrdDetailList.OrdDetailSecType,
                      dbo.OrdDetailList.Security_DPA_, dbo.OrdDetailList.CommissionRate, dbo.OrdDetailList.OrdDetailQty, dbo.OrdDetailList.VolumeRate,
                      dbo.OrdDetailList.VolumeBoundary, dbo.OrdDetailList.MinimumCommission, dbo.OrdDetailList.CMARegulated,
                      dbo.OrdDetailList.PostImmobilisedRate, dbo.OrdDetailList.SecurityImmobilised, dbo.Broker.Broker_DPA_, dbo._CDS_Imported_Trades_.TradeTime,
                      dbo.OrdDetailList.AgentCommission, dbo.OrdDetailList.StaffCommission, CONVERT(smalldatetime,
                      SUBSTRING(dbo._CDS_Imported_Trades_.TradeDate, 3, 2) + '/' + SUBSTRING(dbo._CDS_Imported_Trades_.TradeDate, 1, 2)
                      + '/' + SUBSTRING(dbo._CDS_Imported_Trades_.TradeDate, 5, 4)) AS TradeDate, CONVERT(smalldatetime,
                      SUBSTRING(dbo._CDS_Imported_Trades_.SettlementDate, 3, 2) + '/' + SUBSTRING(dbo._CDS_Imported_Trades_.SettlementDate, 1, 2)
                      + '/' + SUBSTRING(dbo._CDS_Imported_Trades_.SettlementDate, 5, 4)) AS SettlementDate, dbo._CDS_Imported_Trades_.SettlementAmount,
                      dbo.OrdDetailList.EntityType_DPA_, dbo.OrdDetailList.IsCustodian, dbo.OrdDetailList.InterBank, dbo.OrdDetailList.Client_DPA_,
                      dbo.OrdDetailList.Class, dbo._CDS_Imported_Trades_.CommissionRate AS Commission
FROM         dbo.OrdDetailList RIGHT OUTER JOIN
                      dbo._CDS_Imported_Trades_ ON dbo.OrdDetailList.BalanceQty >= dbo._CDS_Imported_Trades_.Quantity AND
                      dbo.OrdDetailList.ClientCDSNo = dbo._CDS_Imported_Trades_.ClientPrefix + dbo._CDS_Imported_Trades_.ClientSuffix AND
                      dbo.OrdDetailList.SecurityCode = dbo._CDS_Imported_Trades_.SecurityDescription AND
                      dbo.OrdDetailList.CDSOrderTypeSale = dbo._CDS_Imported_Trades_.BuySell INNER JOIN
                      dbo.Broker ON LTRIM(RTRIM(REPLACE(dbo._CDS_Imported_Trades_.ContraBrokerID, 'B', ''))) = dbo.Broker.BrokerCode
WHERE     (dbo.OrdDetailList.BalanceQty > 0) AND (dbo._CDS_Imported_Trades_.Processed = 0) AND (dbo.OrdDetailList.OrderHold = 0)
ORDER BY dbo.OrdDetailList.SecurityCode, dbo._CDS_Imported_Trades_.CDSRef, dbo.OrdDetailList.OrdDetailClient


GO


CREATE VIEW dbo.CDSSettlements
AS
SELECT     TOP 100 PERCENT Account_DPA_ AS BankAccount_DPA_, cast(floor(cast(AccountOpeningDate AS float)) AS DateTime) AS TransDate, '' AS Ref,
                      'Opening Balance' AS Particulars, 0 AS Debit, 0 AS Credit, AccountOpeningBal AS CreditBal, 1 AS IsOpeningBalance, '' AS ReceiptNo, 0 AS CDS
FROM         dbo.Account
WHERE     Account_DPA_ = 4
UNION ALL
SELECT     TOP 100 PERCENT 4 AS Account_DPA_, CAST(FLOOR(CAST(SettlementDate AS float)) AS DateTime) AS TransDate, '' AS Ref,
                      'CDS Settlement: ' + CONVERT(Char(12), MIN(LotTDate)) AS Particulars, 0, 0, SUM(CASE (OrderType_DPA_) WHEN 2 THEN LotGrossAmount ELSE 0 END)
                       - SUM(CASE (OrderType_DPA_) WHEN 1 THEN LotGrossAmount ELSE 0 END) AS DepositeAmount, 0 AS IsOpeningBalance, '' AS ReceiptNo,
                      1 AS CDS
FROM         SettlementSchedule
WHERE     (IsCustodian = 0) AND OrderSecType_DPA_ = 2
GROUP BY CAST(FLOOR(CAST(SettlementDate AS float)) AS DateTime)
UNION ALL
SELECT     TOP 100 PERCENT dbo.Payments.Entity_DPA_ AS Entity_DPA_, cast(floor(cast(dbo.Payments.PaymentPDate AS float)) AS DateTime) AS TransDate,
                      dbo.Payments.PaymentReference AS Ref,
                      CASE PayType.PayTypeIn WHEN 0 THEN ' RECEIVED: ' ELSE ' PAID OUT: ' END + isnull(dbo.Payments.PaymentNarrative, '') AS Particulars,
                      CASE PayType.PayTypeIn WHEN 0 THEN Payments.PaymentAmount ELSE 0 END AS Debit,
                      CASE PayType.PayTypeIn WHEN 1 THEN Payments.PaymentAmount ELSE 0 END AS Credit,
                      CASE PayType.PayTypeIn WHEN 1 THEN Payments.PaymentAmount ELSE 0 END AS CreditBal, 0 AS IsOpeningBalance,
                      PaymentReceiptNo AS ReceiptNo, 0 AS CDS
FROM         dbo.Payments INNER JOIN
                      dbo.PayType ON dbo.Payments.PayType_DPA_ = dbo.PayType.PayType_DPA_
WHERE     (EntityType_DPA_ = 5) AND Entity_DPA_ = 4
UNION ALL
SELECT     TOP 100 PERCENT dbo.JournalList.Entity_DPA_, Cast(floor(Cast(dbo.JournalList.JournalDate AS float)) AS DateTime) AS TransDate,
                      CAST(dbo.JournalList.JournalEntry_DPA_ AS SQL_VARIANT) AS Ref, dbo.JournalList.JournalNarrative AS Particulars,
                      JournalList.JournalEntryDebit AS Debit, JournalList.JournalEntryCredit AS Credit, JournalList.JournalEntryCredit AS CreditBal, 0 AS IsOpeningBalance,
                      '' AS ReceiptNo, 0 AS CDS
FROM         dbo.JournalList
WHERE     (EntityType_DPA_ = 5)
ORDER BY IsOpeningBalance DESC, TransDate ASC



GO


CREATE VIEW dbo.CDSUnmatchedHoldings
AS
SELECT     dbo.CDSMatchedHoldings.ClientCode, dbo._CDS_Imported_Holdings_.*
FROM         dbo.CDSMatchedHoldings RIGHT OUTER JOIN
                      dbo._CDS_Imported_Holdings_ ON dbo.CDSMatchedHoldings.ImportedID = dbo._CDS_Imported_Holdings_.ImportedID
WHERE     (dbo.CDSMatchedHoldings.ClientCode IS NULL) AND (dbo._CDS_Imported_Holdings_.Imported = 0)



GO
CREATE VIEW dbo.CDSUnMatchedTradesList
AS
SELECT DISTINCT
                      TOP 100 PERCENT dbo._CDS_Imported_Trades_.CDSImport_DPA_, dbo._CDS_Imported_Trades_.BuySell AS CDSOrderTypeSale,
                      dbo._CDS_Imported_Trades_.CDSRef, dbo._CDS_Imported_Trades_.Quantity, dbo._CDS_Imported_Trades_.Price,
                      REPLACE(dbo._CDS_Imported_Trades_.ContraBrokerID, 'B', '') AS BrokerCode, dbo._CDS_Imported_Trades_.ContraBrokerID AS CDSBrokerCode,
                      dbo._CDS_Imported_Trades_.TradeTime, CONVERT(smalldatetime, SUBSTRING(dbo._CDS_Imported_Trades_.TradeDate, 3, 2)
                      + '/' + SUBSTRING(dbo._CDS_Imported_Trades_.TradeDate, 1, 2) + '/' + SUBSTRING(dbo._CDS_Imported_Trades_.TradeDate, 5, 4)) AS TradeDate,
                      CONVERT(smalldatetime, SUBSTRING(dbo._CDS_Imported_Trades_.SettlementDate, 3, 2)
                      + '/' + SUBSTRING(dbo._CDS_Imported_Trades_.SettlementDate, 1, 2) + '/' + SUBSTRING(dbo._CDS_Imported_Trades_.SettlementDate, 5, 4))
                      AS SettlementDate, dbo._CDS_Imported_Trades_.SettlementAmount,
                      dbo._CDS_Imported_Trades_.ClientPrefix + + dbo._CDS_Imported_Trades_.ClientSuffix AS ClientDescription,
                      dbo._CDS_Imported_Trades_.SecurityDescription,
                      CASE WHEN dbo._CDS_Imported_Trades_.ClientPrefix + + dbo._CDS_Imported_Trades_.ClientSuffix NOT IN
                          (SELECT     ClientCDSNo
                            FROM          client
                            WHERE      Clientcdsno <> 'none') THEN 'Invalid Client CDS No.' ELSE CASE WHEN dbo._CDS_Imported_Trades_.SecurityDescription NOT IN
                          (SELECT     SecurityCode
                            FROM          Security)
                      THEN 'Invalid Security Code.' ELSE CASE WHEN dbo.OrdDetailList.BalanceQty <= 0 THEN 'Order Balance Less than or Equal to Zero' ELSE CASE WHEN
                       dbo.OrdDetailList.OrderHold = 1 THEN 'Order Held' ELSE CASE WHEN isnull(dbo.OrdDetailList.BalanceQty, 0)
                      < dbo._CDS_Imported_Trades_.Quantity THEN 'Trade Quantity More than Order balance' ELSE CASE WHEN REPLACE(dbo._CDS_Imported_Trades_.ContraBrokerID,
                       'B', '') NOT IN
                          (SELECT     BrokerCode
                            FROM          Broker) THEN 'Broker not Found' ELSE 'Unknown' END END END END END END AS Reason
FROM         dbo._CDS_Imported_Trades_ INNER JOIN
                      dbo.NotMatched ON dbo._CDS_Imported_Trades_.CDSImport_DPA_ = dbo.NotMatched.CDSImport_DPA_ LEFT OUTER JOIN
                      dbo.OrdDetailList ON dbo.OrdDetailList.ClientCDSNo = dbo._CDS_Imported_Trades_.ClientPrefix + dbo._CDS_Imported_Trades_.ClientSuffix AND
                      dbo.OrdDetailList.SecurityCode = dbo._CDS_Imported_Trades_.SecurityDescription AND
                      dbo.OrdDetailList.CDSOrderTypeSale = dbo._CDS_Imported_Trades_.BuySell
WHERE     (dbo._CDS_Imported_Trades_.Processed <> 1 AND Ltrim(dbo._CDS_Imported_Trades_.SettlementDate) <> '')
ORDER BY dbo._CDS_Imported_Trades_.CDSRef

GO
CREATE VIEW dbo.ChequeCollection
AS
SELECT     dbo.Client.ClientName, dbo.Payment.PaymentReference AS ChequeNo, dbo.Payment.PaymentAmount, dbo.Payment.Payment_DPA_,
                      dbo.Payment.ChequeCollection, dbo.Payment.ChequeCollectionDate, dbo.Client.Client_DPA_, dbo.Payment.PaymentNarrative,
                      dbo.Payment.ChequeCollectionModifyUser, dbo.Payment.ChequeCollectionModifyDate, dbo.Payment.ChequeCollectionUser,
                      dbo.Payment.ChequeCollectionNarrative AS ChequeCollectionNarrative, dbo.Users.Surname AS ModifyUser, Users_1.Surname AS OriginalUser
FROM         dbo.Payment INNER JOIN
                      dbo.Client ON dbo.Payment.Entity_DPA_ = dbo.Client.Client_DPA_ LEFT OUTER JOIN
                      dbo.Users ON dbo.Payment.ChequeCollectionModifyUser = dbo.Users.UserID LEFT OUTER JOIN
                      dbo.Users Users_1 ON dbo.Payment.ChequeCollectionUser = Users_1.UserID
WHERE     (dbo.Payment.ChequeCollection = N'COLLECTED') AND (dbo.Payment.EntityType_DPA_ = 1) AND (dbo.Payment.Deleted = 0) OR
                      (dbo.Payment.Deleted IS NULL)

GO

CREATE VIEW dbo.ClassList
AS
SELECT     dbo.Class.ClassDescription AS ClassClass, dbo.Class.Class_DPA_, dbo.Class.DefaultSelection, ISNULL(dbo.CommissionList.CommissionDisplay, '')
                      AS DefaultCommission, ISNULL(dbo.CommissionList.Commission_DPA_, 0) AS Commission_DPA_, dbo.Class.IsCda, dbo.Class.AgentStatus,
                      dbo.Class.Voucher
FROM         dbo.Class LEFT OUTER JOIN
                      dbo.CommissionList ON dbo.Class.DefaultCommission = dbo.CommissionList.Commission_DPA_


GO
CREATE VIEW [dbo].[Client Listing NEW]
AS
SELECT     dbo.Client.Client_DPA_, dbo.Client.ClientName, dbo.Client.ClientIDPass, dbo.Client.ClientCDSNo, dbo.Client.IsCustodian, dbo.Agent.AgentName,
                      dbo.Institution.InstitutionName AS Expr2, dbo.Client.ClientCellTel, dbo.Client.ClientEmail, dbo.Class.ClassDescription
FROM         dbo.Client LEFT OUTER JOIN
                      dbo.Agent ON dbo.Client.Agent_DPA_ = dbo.Agent.Agent_DPA_ LEFT OUTER JOIN
                      dbo.Institution ON dbo.Client.Institution_DPA_ = dbo.Institution.Institution_DPA_ LEFT OUTER JOIN
                      dbo.Class ON dbo.Client.Class_DPA_ = dbo.Class.Class_DPA_



GO

CREATE VIEW dbo.ClientCommissions
AS
SELECT     TOP 100 PERCENT dbo.Lots.ContractNumber, CAST(FLOOR(CAST(dbo.Lots.LotTDate AS Float)) AS DateTime) AS TransDate, dbo.tbOrder.Client_DPA_,
                      dbo.LevyContract.LevyAmount, CAST(dbo.Lots.LotQty AS nvarchar) + ' ' + dbo.Security.SecurityCode + ' @ ' + CAST(dbo.Lots.LotPrice AS nvarchar)
                      AS SecurityName, dbo.Payments.PaymentReceiptNo, dbo.LevyContract.Contract_DPA_, dbo.Client.ClientName, dbo.Lots.LotGrossAmount,
                      dbo.Security.SecurityCode, dbo.OrderType.OrderTypeDescription, dbo.Broker.BrokerCode AS BrokerName, dbo.Lots.LotSlipNo, dbo.Lots.LotQty,
                      dbo.Lots.LotPrice, dbo.Agent.Agent_DPA_, dbo.Agent.AgentName
FROM         dbo.Lots INNER JOIN
                      dbo.LevyContract ON dbo.Lots.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.Lots.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ LEFT OUTER JOIN
                      dbo.Agent ON dbo.Client.Agent_DPA_ = dbo.Agent.Agent_DPA_ LEFT OUTER JOIN
                      dbo.Broker ON dbo.Lots.Broker_DPA_ = dbo.Broker.Broker_DPA_ LEFT OUTER JOIN
                      dbo.Payments ON dbo.tbOrder.Order_DPA_ = dbo.Payments.Order_DPA_
WHERE     (dbo.tbOrder.OrderCanceled = 0) AND (dbo.LevyContract.SystemMaintained <> 8) AND (dbo.LevyContract.SystemMaintained <> 12) AND
                      (dbo.LevyContract.LevyShortName LIKE N'%Commission%') AND (dbo.Client.Deleted = 0) AND (dbo.OrdDetail.Deleted = 0) AND
                      (dbo.tbOrder.Deleted = 0) AND (dbo.LevyContract.Deleted = 0)
ORDER BY dbo.tbOrder.Client_DPA_, dbo.Lots.LotTDate


GO

CREATE VIEW dbo.ClientCommissionsByDate
AS
SELECT     TOP 100 PERCENT SUM(LevyAmount) AS Commission, SUM(LotGrossAmount) AS Gross, Client_DPA_, TransDate, ClientName
FROM         dbo.ClientCommissions
GROUP BY Client_DPA_, TransDate, ClientName
ORDER BY Client_DPA_, TransDate


GO
CREATE VIEW dbo.ClientCompundedContractsms
AS
SELECT     TOP 100 PERCENT dbo.Contract.Contract_DPA_, dbo.LotList.Lot_DPA_, dbo.LotList.BrokerCode, dbo.LotList.OrdDetailType,
                      dbo.LotList.OrdDetailSecType, dbo.LotList.OrderTypeSale, dbo.LotList.LotSlipNo, dbo.LotList.LotTDate, dbo.LotList.LotQty, dbo.LotList.LotPrice,
                      dbo.LotList.ContractNumber, dbo.LotList.OrdDetailClient, dbo.LotList.OrdDetailSecurity, dbo.LevyContract.LevyAmount, dbo.LevyContract.LevyName,
                      dbo.tbOrder.OrderRef, dbo.Client.ClientCellTel, dbo.Client.Client_DPA_, dbo.LotList.LotGrossAmount AS GrossAmount, dbo.tbOrder.Order_DPA_,
                      dbo.LevyContract.SystemMaintained, CASE WHEN ISNULL(CHARINDEX('%', dbo.LevyContract.LevyRatePercentage), 0)
                      > 0 THEN '%' ELSE '' END AS LevyRatePercentage, dbo.LevyContract.LevyShortName
FROM         dbo.Contract INNER JOIN
                      dbo.LotList ON dbo.Contract.Contract_DPA_ = dbo.LotList.Contract_DPA_ INNER JOIN
                      dbo.LevyContract ON dbo.Contract.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.LotList.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.LotList.Client_DPA_ = dbo.Client.Client_DPA_
WHERE     (dbo.tbOrder.OrderCompounded = 1)
ORDER BY dbo.LotList.LotTDate, dbo.LotList.OrdDetailSecurity, dbo.LotList.LotSlipNo

GO
CREATE VIEW dbo.ClientContractCompounded
AS
SELECT     TOP 100 PERCENT dbo.Contract.Contract_DPA_, dbo.LotList.Lot_DPA_, dbo.LotList.BrokerCode, dbo.LotList.OrdDetailType,
                      dbo.LotList.OrdDetailSecType, dbo.LotList.OrderTypeSale, dbo.LotList.LotSlipNo, dbo.LotList.LotTDate, dbo.LotList.LotQty, dbo.LotList.LotPrice,
                      dbo.LotList.ContractNumber, dbo.LotList.OrdDetailClient, dbo.LotList.OrdDetailSecurity, dbo.LevyContract.LevyName, dbo.tbOrder.OrderRef,
                      dbo.Client.ClientAddr, dbo.Client.Client_DPA_, dbo.LotList.LotGrossAmount AS GrossAmount, dbo.tbOrder.Order_DPA_,
                      dbo.LevyContract.SystemMaintained, CASE WHEN ISNULL(CHARINDEX('%', dbo.LevyContract.LevyRatePercentage), 0)
                      > 0 THEN '%' ELSE '' END AS LevyRatePercentage, dbo.LevyContract.LevyShortName,
                      CASE WHEN LotList.OrdDetailSecType = 'Fixed' THEN LevyContract.LevyRate * LotList.LotGrossAmount / 100 ELSE LevyContract.LevyAmount END AS LevyAmount,
                       dbo.LotList.ClientCDSNo, dbo.LotList.OrdDetail_DPA_, dbo.LevyContract.LevyVATAmount, dbo.LotList.ContractSettlementDate,dbo.LotList.Security_DPA_
FROM         dbo.Contract INNER JOIN
                      dbo.LotList ON dbo.Contract.Contract_DPA_ = dbo.LotList.Contract_DPA_ INNER JOIN
                      dbo.LevyContract ON dbo.Contract.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.LotList.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.LotList.Client_DPA_ = dbo.Client.Client_DPA_
WHERE     (dbo.tbOrder.OrderCompounded = 1)
ORDER BY dbo.LotList.LotTDate, dbo.LotList.OrdDetailSecurity, dbo.LotList.LotSlipNo


GO
CREATE VIEW dbo.ClientContractCompounded
AS

SELECT     TOP 100 PERCENT dbo.Contract.Contract_DPA_, dbo.LotList.Lot_DPA_, dbo.LotList.BrokerCode, dbo.LotList.OrdDetailType, dbo.LotList.OrdDetailSecType, dbo.LotList.OrderTypeSale,
                      dbo.LotList.LotSlipNo, dbo.LotList.LotTDate, dbo.LotList.LotQty, dbo.LotList.LotPrice, dbo.LotList.ContractNumber, dbo.LotList.OrdDetailClient,
                      dbo.LotList.OrdDetailSecurity, dbo.LevyContract.LevyAmount, dbo.LevyContract.LevyName, dbo.tbOrder.OrderRef, dbo.Client.ClientAddr,dbo.Client.Client_DPA_,
                      dbo.LotList.LotGrossAmount AS GrossAmount, dbo.tbOrder.Order_DPA_, dbo.LevyContract.SystemMaintained,
		      CASE
				WHEN  ISNULL(CHARINDEX('%', dbo.LevyContract.LevyRatePercentage), 0) > 0 THEN
					'%'
				ELSE
					''
			END As 	LevyRatePercentage, dbo.LevyContract.LevyShortName
FROM         dbo.Contract INNER JOIN
                      dbo.LotList ON dbo.Contract.Contract_DPA_ = dbo.LotList.Contract_DPA_ INNER JOIN
                      dbo.LevyContract ON dbo.Contract.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.LotList.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.LotList.Client_DPA_ = dbo.Client.Client_DPA_
WHERE     (dbo.tbOrder.OrderCompounded = 1)
ORDER BY dbo.LotList.LotTDate, dbo.LotList.OrdDetailSecurity, dbo.LotList.LotSlipNo








GO
CREATE VIEW dbo.ClientEmailDocsList
AS
SELECT     dbo.ClientEmailsDocs.ID, dbo.Client.Client_DPA_, dbo.Client.ClientCDSNo, dbo.Client.ClientName, dbo.EmailDocs.DocName,
                      dbo.EmailMode.ModeName, dbo.Users.OtherNames + N'  ' + dbo.Users.Surname AS username, dbo.ClientEmailsDocs.TimeChanged
FROM         dbo.ClientEmailsDocs INNER JOIN
                      dbo.EmailDocs ON dbo.ClientEmailsDocs.Document_DPA_ = dbo.EmailDocs.Document_DPA_ INNER JOIN
                      dbo.EmailMode ON dbo.ClientEmailsDocs.EmailModeID = dbo.EmailMode.ModeId INNER JOIN
                      dbo.Client ON dbo.ClientEmailsDocs.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Users ON dbo.Client.Client_DPA_ = dbo.Users.Client_DPA_
WHERE     (dbo.ClientEmailsDocs.Deleted <> 1)

GO
CREATE VIEW dbo.ClientEntityList
AS
SELECT     dbo.Client.ClientName AS EntityName, dbo.EntityType.EntityTypeName AS EntityType, dbo.Client.Client_DPA_ AS Entity_DPA_,
                      dbo.EntityType.EntityType_DPA_, CAST(dbo.Client.Client_DPA_ AS NVARCHAR(4000)) AS EntityCode, CAST(dbo.Client.Client_DPA_ AS NVARCHAR(4000))
                       + SPACE
                          ((SELECT     ISNULL(MAX(LEN(CAST(dbo.Client.Client_DPA_ AS NVARCHAR(4000)))), 0) AS MAXLEN
                              FROM         dbo.Client) - LEN(CAST(dbo.Client.Client_DPA_ AS NVARCHAR(4000)))) + ' : ' + dbo.Client.ClientName AS EntityNameEx
FROM         dbo.Client INNER JOIN
                      dbo.EntityType ON dbo.Client.EntityType_DPA_ = dbo.EntityType.EntityType_DPA_

GO
CREATE VIEW dbo.ClientList
AS
SELECT     TOP 100 PERCENT dbo.ActivatedClientList.*
FROM         dbo.ActivatedClientList

GO
CREATE VIEW dbo.ClientPortfolioAveragePriceList
AS
SELECT     Client_DPA_, ClientName, Security_DPA_, SecurityName, SUM(Quantity) AS TotalQty, SUM(Amount) AS TotalAmount, SUM(Amount) / SUM(Quantity)
                      AS AveragePrice
FROM         dbo.ClientPortfolioPurchaseList
GROUP BY Client_DPA_, ClientName, Security_DPA_, SecurityName

GO
CREATE VIEW dbo.ClientPortfolioPurchase
AS
SELECT     dbo.ClientSecurityList.ClientName, CASE isnull(dbo.ClientPortfolioAveragePriceList.AveragePrice, - 1)
                      WHEN - 1 THEN dbo.ClientSecurityList.SecurityName ELSE dbo.ClientPortfolioAveragePriceList.SecurityName END AS SecurityName,
                      '@ ' + CONVERT(nvarchar(500), dbo.ClientPortfolioAveragePriceList.AveragePrice) AS AveragePrice,
                          (SELECT     SUM(dbo.ClientPortfolioPurchaseList.Quantity) AS Purchases
                            FROM          dbo.ClientPortfolioPurchaseList
                            WHERE      dbo.ClientPortfolioAveragePriceList.Client_DPA_ = dbo.ClientPortfolioPurchaseList.Client_DPA_ AND
                                                   dbo.ClientPortfolioAveragePriceList.Security_DPA_ = dbo.ClientPortfolioPurchaseList.Security_DPA_ AND
                                                   dbo.ClientPortfolioPurchaseList.IsPortfolio = 0) AS Purchases, dbo.ClientSecurityList.Client_DPA_,
                      dbo.ClientSecurityList.Security_DPA_
FROM         dbo.ClientPortfolioAveragePriceList RIGHT OUTER JOIN
                      dbo.ClientSecurityList ON dbo.ClientPortfolioAveragePriceList.Client_DPA_ = dbo.ClientSecurityList.Client_DPA_ AND
                      dbo.ClientPortfolioAveragePriceList.Security_DPA_ = dbo.ClientSecurityList.Security_DPA_

GO
CREATE VIEW dbo.ClientPortfolioPurchaseList
AS
SELECT     Client_DPA_, OrdDetailClient AS ClientName, Security_DPA_, OrdDetailSecurity AS SecurityName, LotPrice, SUM(LotQty) AS Quantity, SUM(LotPrice * LotQty) AS Amount,
                      0 AS IsPortfolio
FROM         dbo.LotList
WHERE     (OrderTypeSale = 0)
GROUP BY Client_DPA_, OrdDetailClient, Security_DPA_, OrdDetailSecurity, LotPrice
UNION
SELECT     Client_DPA_, CPortfolioClient AS ClientName, Security_DPA_, CPortfolioSecurity AS SecurityName,  CPortfolioPrice AS Price, CPortfolioQty AS Quantity,CPortfolioPrice * CPortfolioQty AS Amount,
                      1 AS IsPortfolio
FROM         dbo.CPortfolioList




GO

CREATE VIEW dbo.ClientPortfolioQuantities
AS
SELECT     TOP 100 PERCENT dbo.CPortfolioList.CPortfolio_DPA_ AS Order_DPA_, dbo.Security.SecurityCode, dbo.CPortfolioList.Client_DPA_,
                      dbo.CPortfolioList.Security_DPA_, 0 AS OrderTypeSale, 'E' + CAST(dbo.CPortfolioList.CPortfolio_DPA_ AS char(2)) AS ContractNumber,
                      'Unknown' AS LotSlipNo, CONVERT(Char(7), dbo.CPortfolioList.CPortfolioPrice) AS OrdDetailPrice, dbo.CPortfolioList.CPortfolioQty AS OrdDetailQty,
                      dbo.Client.ClientName AS OrdDetailClient, dbo.CPortfolioList.CPortfolio_DPA_ AS OrdDetail_DPA_, 0 AS Amount, dbo.Security.SecurityMktPrice,
                      dbo.CPortfolioList.CPortfolio_DPA_ AS Lot_DPA_, dbo.CPortfolioList.CPortfolioQty AS LotQty, dbo.CPortfolioList.CPortfolioPrice AS LotPrice,
                      dbo.CPortfolioList.CPortfolioPDate AS LotTDate, dbo.CPortfolioList.CPortfolioQty * dbo.CPortfolioList.CPortfolioPrice AS LotGrossAmount
FROM         dbo.CPortfolioList INNER JOIN
                      dbo.Client ON dbo.CPortfolioList.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Security ON dbo.CPortfolioList.Security_DPA_ = dbo.Security.Security_DPA_
WHERE     (dbo.Client.Deleted = 0)


GO
CREATE VIEW dbo.ClientPortfolioSaleList
AS
SELECT     Client_DPA_, OrdDetailClient AS ClientName, Security_DPA_, OrdDetailSecurity AS SecurityName, LotPrice, SUM(LotQty) AS Quantity
FROM         dbo.LotList
WHERE     (OrderTypeSale = 1)
GROUP BY Client_DPA_, OrdDetailClient, Security_DPA_, OrdDetailSecurity, LotPrice



GO
CREATE VIEW dbo.ClientPortfolioSold
AS
SELECT     dbo.ClientSecurityList.ClientName, CASE isnull(SUM(dbo.ClientPortfolioSaleList.Quantity), - 1)
                      WHEN - 1 THEN dbo.ClientSecurityList.SecurityName ELSE dbo.ClientPortfolioSaleList.SecurityName END AS SecurityName,
                      SUM(dbo.ClientPortfolioSaleList.Quantity) AS Sold, dbo.ClientSecurityList.Client_DPA_, dbo.ClientSecurityList.Security_DPA_
FROM         dbo.ClientPortfolioSaleList RIGHT OUTER JOIN
                      dbo.ClientSecurityList ON dbo.ClientPortfolioSaleList.Client_DPA_ = dbo.ClientSecurityList.Client_DPA_ AND
                      dbo.ClientPortfolioSaleList.Security_DPA_ = dbo.ClientSecurityList.Security_DPA_
GROUP BY dbo.ClientPortfolioSaleList.ClientName, dbo.ClientPortfolioSaleList.SecurityName, dbo.ClientSecurityList.Client_DPA_,
                      dbo.ClientSecurityList.Security_DPA_, dbo.ClientSecurityList.ClientName, dbo.ClientSecurityList.SecurityName

GO
CREATE VIEW dbo.ClientPortfolioStatement
AS
SELECT     TOP 100 PERCENT dbo.ClientPortfolioTransactions.ClientName AS [Client],  dbo.ClientPortfolioTransactions.Client_DPA_ AS [Code], dbo.ClientPortfolioTransactions.SecurityName AS [Security],
                      dbo.ClientPortfolioTransactions.AveragePrice AS [Average Price], dbo.ClientPortfolioTransactions.Purchases, dbo.ClientPortfolioTransactions.Sold,
                      dbo.ClientPortfolioTransactions.Balance + ISNULL(dbo.CPortfolio.CPortfolioQty, 0) AS Balance
FROM         dbo.ClientPortfolioTransactions LEFT OUTER JOIN
                      dbo.CPortfolio ON dbo.CPortfolio.Client_DPA_ = dbo.ClientPortfolioTransactions.Client_DPA_ AND
                      dbo.CPortfolio.Security_DPA_ = dbo.ClientPortfolioTransactions.Security_DPA_

UNION

SELECT     TOP 100 PERCENT dbo.ClientList.ClientName, dbo.ClientList.Client_DPA_ AS [Code], dbo.SecurityList.SecurityName, ISNULL(dbo.ClientPortfolioTransactions.AveragePrice,
                      '@ ' + CONVERT(nvarchar(500), dbo.CPortfolio.CPortfolioPrice)) AS AveragePrice, ISNULL(dbo.ClientPortfolioTransactions.Purchases, 0) AS Purchases,
                      ISNULL(dbo.ClientPortfolioTransactions.Sold, 0) AS Sold, ISNULL(dbo.ClientPortfolioTransactions.Balance, 0) + ISNULL(dbo.CPortfolio.CPortfolioQty, 0)
                      AS Balance
FROM         dbo.ClientPortfolioTransactions RIGHT OUTER JOIN
                      dbo.ClientList INNER JOIN
                      dbo.SecurityList INNER JOIN
                      dbo.CPortfolio ON dbo.SecurityList.Security_DPA_ = dbo.CPortfolio.Security_DPA_ ON dbo.ClientList.Client_DPA_ = dbo.CPortfolio.Client_DPA_ ON
                      dbo.ClientPortfolioTransactions.Client_DPA_ = dbo.CPortfolio.Client_DPA_ AND
                      dbo.ClientPortfolioTransactions.Security_DPA_ = dbo.CPortfolio.Security_DPA_



GO
CREATE VIEW dbo.ClientPortfolioTransactions
AS
SELECT     TOP 100 PERCENT dbo.ClientPortfolioPurchase.ClientName, CASE ISNULL(dbo.ClientPortfolioPurchase.Purchases, - 1)
                      WHEN - 1 THEN dbo.ClientPortfolioSold.SecurityName ELSE dbo.ClientPortfolioPurchase.SecurityName END AS SecurityName,
                      ISNULL(dbo.ClientPortfolioPurchase.AveragePrice, '') AS AveragePrice, ISNULL(dbo.ClientPortfolioPurchase.Purchases, 0) AS Purchases,
                      ISNULL(dbo.ClientPortfolioSold.Sold, 0) AS Sold, ISNULL(dbo.ClientPortfolioPurchase.Purchases, 0) - ISNULL(dbo.ClientPortfolioSold.Sold, 0)
                      AS Balance, dbo.ClientPortfolioPurchase.Client_DPA_, dbo.ClientPortfolioPurchase.Security_DPA_
FROM         dbo.ClientPortfolioSold RIGHT OUTER JOIN
                      dbo.ClientPortfolioPurchase ON dbo.ClientPortfolioSold.Client_DPA_ = dbo.ClientPortfolioPurchase.Client_DPA_ AND
                      dbo.ClientPortfolioSold.Security_DPA_ = dbo.ClientPortfolioPurchase.Security_DPA_
WHERE     (dbo.ClientPortfolioPurchase.Purchases IS NOT NULL) OR
                      (dbo.ClientPortfolioSold.Sold IS NOT NULL)
ORDER BY dbo.ClientPortfolioPurchase.ClientName, dbo.ClientPortfolioPurchase.SecurityName
GO
CREATE VIEW dbo.ClientReport
AS
SELECT     TOP 100 PERCENT dbo.Client.ClientName AS Name, dbo.Client.ClientCDSNo AS [CDS Number], dbo.Client.ClientOfficeTel AS [Office Telephone],
                      dbo.Client.ClientCellTel AS [Cell Phone], dbo.Client.ClientEmail AS Email, dbo.Client.ClientAddr AS Address, dbo.Client.IsCustodian
FROM         dbo.Client LEFT OUTER JOIN
                      dbo.OwnerList ON dbo.Client.Owner_DPA_ = dbo.OwnerList.Owner_DPA_
ORDER BY dbo.Client.ClientName

GO

CREATE   VIEW app.Clients AS
SELECT  Client_DPA_        AS ClientId,
        ClientName         AS Name,
        ClientCDSNo        AS CdsNumber,
        ClientIDPass       AS IdNumber,
        ClientEmail        AS Email,
        ClientCellTel      AS Mobile,
        Agent_DPA_         AS AgentId,
        Branch_DPA_        AS BranchId,
        Class_DPA_         AS ClassId,
        Commission_DPA_    AS CommissionId,
        Residency_DPA_     AS ResidencyId,
        ClientOpeningBal   AS OpeningBalance,
        CreditLimit        AS CreditLimit,
        ClientVIP          AS IsVip,
        IsCustodian        AS IsCustodian,
        IsNominee          AS IsNominee,
        IsFrozen           AS IsFrozen,
        ClientRegDate      AS RegisteredOn
FROM dbo.Client
WHERE Deleted = 0 OR Deleted IS NULL;

GO
CREATE VIEW dbo.ClientsAboveCreditLimit
AS
SELECT     TOP 100 PERCENT dbo.Client.ClientName AS Client, CAST(dbo.Client.ClientAddr AS NVARCHAR(1000)) AS Address,
                      dbo.ClientsAboveCreditLimit2.Balance, dbo.Client.CreditLimit AS [Credit Limit],
                      dbo.ClientsAboveCreditLimit2.Balance - dbo.Client.CreditLimit AS [Excess Amount], dbo.Client.ClientOfficeTel AS Telephone
FROM         dbo.Client INNER JOIN
                      dbo.ClientsAboveCreditLimit2 ON dbo.Client.Client_DPA_ = dbo.ClientsAboveCreditLimit2.Client_DPA_
GROUP BY dbo.Client.ClientName, CAST(dbo.Client.ClientAddr AS NVARCHAR(1000)), dbo.Client.CreditLimit, dbo.Client.ClientOfficeTel,
                      dbo.ClientsAboveCreditLimit2.Balance, dbo.Client.Deleted
HAVING      (dbo.ClientsAboveCreditLimit2.Balance > dbo.Client.CreditLimit) AND (dbo.Client.Deleted = 0) OR
                      (dbo.Client.Deleted IS NULL)
ORDER BY dbo.Client.ClientName

GO
CREATE VIEW dbo.ClientsAboveCreditLimit2
AS
SELECT     TOP 100 PERCENT Client_DPA_, CAST(REPLACE(Balance, ' Dr', '') AS Money) AS Balance
FROM         dbo.ClientStatement CStatement
WHERE     (CAST(TransDate AS Float) =
                          (SELECT     MAX(cast(TransDate AS Float)) AS UniqueDate
                            FROM          dbo.ClientStatement
                            WHERE      CStatement.Client_DPA_ = Client_DPA_)) AND (Balance NOT LIKE '%Cr%')

GO
CREATE VIEW dbo.ClientSecurityList
AS
SELECT     dbo.ClientList.*, dbo.SecurityList.*
FROM         dbo.ClientList CROSS JOIN
                      dbo.SecurityList

GO

CREATE VIEW dbo.ClientsHoldings
AS
SELECT     dbo.Holdings.*, dbo.Security.SecurityCode AS SecurityCode
FROM         dbo.Holdings INNER JOIN
                      dbo.Security ON dbo.Holdings.Security_DPA_ = dbo.Security.Security_DPA_


GO
CREATE VIEW dbo.ClientStatement
AS
SELECT     TOP 100 PERCENT COUNT(*) AS ClientTransaction_DPA_, a.Client_DPA_, a.TransDate, a.Ref, a.Particulars, a.Debit, a.Credit,
                      CASE a.IsOpeningBalance WHEN 1 THEN CASE WHEN a.Balance >= 0 THEN CONVERT(NVARCHAR(400), a.Balance)
                      + ' Cr' ELSE CONVERT(NVARCHAR(400), a.Balance) + ' Dr' END ELSE CASE WHEN SUM(b.Balance) >= 0 THEN CONVERT(NVARCHAR(400),
                      SUM(b.Balance)) + ' Cr' ELSE CONVERT(NVARCHAR(400), ABS(SUM(b.Balance))) + ' Dr' END END AS Balance, a.IsOpeningBalance
FROM         (SELECT     TOP 100 PERCENT *
                       FROM          dbo.ClientTransactionList) a CROSS JOIN
                          (SELECT     TOP 100 PERCENT *
                            FROM          dbo.ClientTransactionList) b
WHERE     a.TransDate >= b.TransDate AND a.Client_DPA_ = b.Client_DPA_
GROUP BY a.Client_DPA_, a.TransDate, a.Ref, a.Particulars, a.Debit, a.Credit, a.Balance, a.IsOpeningBalance
ORDER BY a.Client_DPA_, a.IsOpeningBalance DESC, a.ClientTransaction_DPA_, a.Particulars, a.Ref, a.Debit, a.Credit, a.Balance

GO
CREATE VIEW dbo.ClientStatementBalances
AS
SELECT     TOP 100 PERCENT dbo.ClientStatement.Client_DPA_, SUM(dbo.ClientStatement.Credit - dbo.ClientStatement.Debit)
                      + dbo.Client.ClientOpeningBal AS Balance
FROM         dbo.ClientStatement INNER JOIN
                      dbo.Client ON dbo.ClientStatement.Client_DPA_ = dbo.Client.Client_DPA_
GROUP BY dbo.ClientStatement.Client_DPA_, dbo.Client.ClientOpeningBal
ORDER BY dbo.ClientStatement.Client_DPA_

GO
CREATE VIEW dbo.ClientStatementTransactionList
AS
SELECT     ClientsStatement.*, CreditBal - Debit AS Balance
FROM         (SELECT     Client_DPA_, ClientRegDate AS TransDate, '' AS REF, ' Opening Balance' AS Particulars, 0 AS Debit, 0 AS Credit,
                                              ClientOpeningBal AS CreditBal, 1 AS IsOpeningBalance, '' AS PaymentReceiptNo, 0 AS ReceiptType
                       FROM          dbo.Client
                       WHERE      deleted = 0
                       UNION
                       SELECT     dbo.Payment.Entity_DPA_ AS Client_DPA_, dbo.Payment.PaymentPDate AS TransDate, dbo.Payment.PaymentReference AS Ref,
                                             dbo.Payment.PaymentNarrative AS Particulars, CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS CreditBal, 0 AS IsOpeningBalance,
                                             cast(dbo.PaymentClientStatementNos.ReceiptNo AS varchar), dbo.PaymentClientStatementNos.type AS receipttype
                       FROM         dbo.Payment INNER JOIN
                                             dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_ INNER JOIN
                                             dbo.PaymentClientStatementNos ON Payment.Payment_DPA_ = PaymentClientStatementNos.Payment_DPA_
                       WHERE     EntityType_DPA_ = 1 AND Payment.deleted = 0
                       UNION
                       SELECT     Client_DPA_, LotTDate AS TransDate, cast(ContractNumber AS varchar) AS Ref, SecurityName AS Particulars,
                                             CASE OrderTypeSale WHEN 0 THEN NetAmt ELSE 0 END AS Debit, CASE OrderTypeSale WHEN 1 THEN NetAmt ELSE 0 END AS Credit,
                                             CASE OrderTypeSale WHEN 1 THEN NetAmt ELSE 0 END AS CreditBal, 0 AS IsOpeningBalance, cast(ContractNumber AS varchar)
                                             AS PaymentReceiptNo, 2 AS receipttype
                       FROM         dbo.ClientTransactionsSubList
                       UNION
                       SELECT     dbo.JournalList.Entity_DPA_ AS Client_DPA_, dbo.JournalList.JournalDate AS TransDate,
                                             CAST(dbo.JournalList.JournalEntry_DPA_ AS Nvarchar(500)) AS Ref, dbo.JournalList.JournalNarrative AS Particulars,
                                             JournalList.JournalEntryDebit AS Debit, JournalList.JournalEntryCredit AS Credit, JournalList.JournalEntryCredit AS CreditBal,
                                             0 AS IsOpeningBalance, '' AS PaymentReceiptNo, 4 AS receipttype
                       FROM         dbo.JournalList
                       WHERE     EntityType_DPA_ = 1
                       UNION
                       SELECT     dbo.InterTransfers.Client_DPA_, dbo.InterTransfer.TransferDate, dbo.InterTransfer.TransferReference AS REF,
                                             isnull(dbo.InterTransfer.TransferNarrative, '') + '  ' + InterTransferType.TypeDescription AS Particulars, 0 AS Debit,
                                             dbo.InterTransfer.TransferAmount AS Credit, dbo.InterTransfer.TransferAmount AS CreditBal, 0 AS IsOpeningBalance,
                                             cast(dbo.InterTransfers.ContractNumber AS varchar) AS PaymentReceiptNo, 5 AS receipttype
                       FROM         dbo.Contract INNER JOIN
                                             dbo.InterTransfer ON dbo.Contract.Contract_DPA_ = dbo.InterTransfer.Contract_DPA_ INNER JOIN
                                             dbo.InterTransfers ON dbo.Contract.Contract_DPA_ = dbo.InterTransfers.Contract_DPA_ INNER JOIN
                                             dbo.InterTransferType ON dbo.InterTransfer.InterTransferType_DPA_ = dbo.InterTransferType.InterTransferType_DPA_
                       WHERE     (dbo.InterTransfers.OrderTypeSale = 0) AND intertransfer.deleted = 0
                       UNION
                       SELECT     dbo.InterTransfers.Client_DPA_, dbo.InterTransfer.TransferDate, dbo.InterTransfer.TransferReference AS REF,
                                             isnull(dbo.InterTransfer.TransferNarrative, '') + '  ' + InterTransferType.TypeDescription AS Particulars,
                                             dbo.InterTransfer.TransferAmount AS Debit, 0 AS Credit, 0 AS CreditBal, 0 AS IsOpeningBalance,
                                             cast(dbo.InterTransfers.ContractNumber AS varchar) AS PaymentReceiptNo, 5 AS receipttype
                       FROM         dbo.Contract INNER JOIN
                                             dbo.InterTransfer ON dbo.Contract.Contract_DPA_ = dbo.InterTransfer.Contract_DPA_ INNER JOIN
                                             dbo.InterTransfers ON dbo.Contract.Contract_DPA_ = dbo.InterTransfers.Contract_DPA_ INNER JOIN
                                             dbo.InterTransferType ON dbo.InterTransfer.InterTransferType_DPA_ = dbo.InterTransferType.InterTransferType_DPA_
                       WHERE     (dbo.InterTransfers.OrderTypeSale = 1) AND intertransfer.deleted = 0
                       UNION
                       SELECT     Offerings.Client_DPA_, CAST(FLOOR(CAST(Offerings.Offerings_Date AS float)) AS DateTime) AS TransDate, Offerings.PAL_No AS REF,
                                             CAST(FLOOR(CAST(Offerings.Alloted_Rights AS INT)) AS Char) + ' ' + CASE WHEN len(Security.SecurityName)
                                             > 20 THEN LEFT(Security.SecurityName, 20) + '...' ELSE Security.SecurityName END + ' @ ' + CAST(Offerings.Offering_Price AS char(5))
                                             AS Particulars, Offerings.Alloted_Rights * Offerings.Offering_Price AS Debit, 0 AS Credit, 0 AS CreditBal, 0 AS IsOpeningbal,
                                             '' AS PaymentReceiptNo, - 1 AS ReceiptType
                       FROM         Offerings INNER JOIN
                                             Security ON Offerings.Offering = Security.Security_DPA_
                       WHERE     (Offerings.Deleted = 0)) ClientsStatement

GO

CREATE VIEW dbo.ClientStatementTransactionList1
AS
SELECT     ClientsStatement.*, CreditBal - Debit AS Balance
FROM         (SELECT     Client_DPA_, ClientRegDate AS TransDate, '' AS REF, ' Opening Balance' AS Particulars, 0 AS Debit, 0 AS Credit,
                                              ClientOpeningBal AS CreditBal, 1 AS IsOpeningBalance, '' AS PaymentReceiptNo, 0 AS ReceiptType
                       FROM          dbo.Client
                       WHERE      deleted = 0
                       UNION
                       SELECT     dbo.Payment.Entity_DPA_ AS Client_DPA_, dbo.Payment.PaymentPDate AS TransDate, dbo.Payment.PaymentReference AS Ref,
                                             dbo.Payment.PaymentNarrative AS Particulars, CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS CreditBal, 0 AS IsOpeningBalance,
                                             dbo.PaymentClientStatementNos.ReceiptNo, dbo.PaymentClientStatementNos.type AS receipttype
                       FROM         dbo.Payment INNER JOIN
                                             dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_ INNER JOIN
                                             dbo.PaymentClientStatementNos ON Payment.Payment_DPA_ = PaymentClientStatementNos.Payment_DPA_
                       WHERE     EntityType_DPA_ = 1 AND Payment.deleted = 0
                       UNION
                       SELECT     Client_DPA_, LotTDate AS TransDate, ContractNumber AS Ref, SecurityName AS Particulars,
                                             CASE OrderTypeSale WHEN 0 THEN NetAmt ELSE 0 END AS Debit, CASE OrderTypeSale WHEN 1 THEN NetAmt ELSE 0 END AS Credit,
                                             CASE OrderTypeSale WHEN 1 THEN NetAmt ELSE 0 END AS CreditBal, 0 AS IsOpeningBalance, ContractNumber AS PaymentReceiptNo,
                                             2 AS receipttype
                       FROM         dbo.ClientTransactionsSubList1
                       UNION
                       SELECT     dbo.JournalList.Entity_DPA_ AS Client_DPA_, dbo.JournalList.JournalDate AS TransDate,
                                             CAST(dbo.JournalList.JournalEntry_DPA_ AS Nvarchar(500)) AS Ref, dbo.JournalList.JournalNarrative AS Particulars,
                                             JournalList.JournalEntryDebit AS Debit, JournalList.JournalEntryCredit AS Credit, JournalList.JournalEntryCredit AS CreditBal,
                                             0 AS IsOpeningBalance, '' AS PaymentReceiptNo, 4 AS receipttype
                       FROM         dbo.JournalList
                       WHERE     EntityType_DPA_ = 1
                       UNION
                       SELECT     dbo.InterTransfers.Client_DPA_, dbo.InterTransfer.TransferDate, dbo.InterTransfer.TransferReference AS REF,
                                             isnull(dbo.InterTransfer.TransferNarrative, '') + '  ' + InterTransferType.TypeDescription AS Particulars, 0 AS Debit,
                                             dbo.InterTransfer.TransferAmount AS Credit, dbo.InterTransfer.TransferAmount AS CreditBal, 0 AS IsOpeningBalance,
                                             dbo.InterTransfers.ContractNumber AS PaymentReceiptNo, 5 AS Type
                       FROM         dbo.Contract INNER JOIN
                                             dbo.InterTransfer ON dbo.Contract.Contract_DPA_ = dbo.InterTransfer.Contract_DPA_ INNER JOIN
                                             dbo.InterTransfers ON dbo.Contract.Contract_DPA_ = dbo.InterTransfers.Contract_DPA_ INNER JOIN
                                             dbo.InterTransferType ON dbo.InterTransfer.InterTransferType_DPA_ = dbo.InterTransferType.InterTransferType_DPA_
                       WHERE     (dbo.InterTransfers.OrderTypeSale = 0) AND intertransfer.deleted = 0
                       UNION
                       SELECT     dbo.InterTransfers.Client_DPA_, dbo.InterTransfer.TransferDate, dbo.InterTransfer.TransferReference AS REF,
                                             isnull(dbo.InterTransfer.TransferNarrative, '') + '  ' + InterTransferType.TypeDescription AS Particulars,
                                             dbo.InterTransfer.TransferAmount AS Debit, 0 AS Credit, 0 AS CreditBal, 0 AS IsOpeningBalance,
                                             dbo.InterTransfers.ContractNumber AS PaymentReceiptNo, 5 AS Type
                       FROM         dbo.Contract INNER JOIN
                                             dbo.InterTransfer ON dbo.Contract.Contract_DPA_ = dbo.InterTransfer.Contract_DPA_ INNER JOIN
                                             dbo.InterTransfers ON dbo.Contract.Contract_DPA_ = dbo.InterTransfers.Contract_DPA_ INNER JOIN
                                             dbo.InterTransferType ON dbo.InterTransfer.InterTransferType_DPA_ = dbo.InterTransferType.InterTransferType_DPA_
                       WHERE     (dbo.InterTransfers.OrderTypeSale = 1) AND intertransfer.deleted = 0) ClientsStatement






GO
CREATE VIEW dbo.ClientStatusSummary
AS


SELECT LTRIM(REPLACE(Security, '^', '')) AS Security, Quantity, Price, Validity, Certificate, [Slip No.], [Trade Date], [Trade Qty], [Trade Price], Broker, Contract


FROM (SELECT TOP 100 PERCENT * FROM

(SELECT     dbo.OrdDetail.OrdDetail_DPA_,
	 '       ' + LTRIM(dbo.Client.ClientName) AS Security,
	  '' AS Quantity,
	  '' AS Price, '' AS Validity, '' AS Certificate, '' AS [Slip No.], '' AS [Trade Date],
                      '' AS [Trade Qty], '' AS [Trade Price], '' AS Broker, '' AS Contract
FROM         dbo.tbOrder INNER JOIN
                      dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_

UNION ALL

SELECT     dbo.OrdDetail.OrdDetail_DPA_,
	 '      ' + CAST(dbo.Client.ClientAddr AS NVARCHAR(500)) AS Security,
	  '' AS Quantity,
	  '' AS Price, '' AS Validity, '' AS Certificate, '' AS [Slip No.], '' AS [Trade Date],
                      '' AS [Trade Qty], '' AS [Trade Price], '' AS Broker, '' AS Contract
FROM         dbo.tbOrder INNER JOIN
                      dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_


UNION ALL


SELECT     dbo.OrdDetail.OrdDetail_DPA_,
	   '    ' + dbo.OrderType.OrderTypeDescription + ' Order: ' + dbo.tbOrder.OrderRef  AS Security,
	   '' AS Quantity,
		'' AS Price,
		'' AS Validity,
		'' AS Certificate,
		'' [Slip No.],
		'' [Trade Date],
		'' [Trade Qty],
		'' [Trade Price],
		'' [Broker],
		'' [Contract]
FROM         dbo.tbOrder INNER JOIN
                      dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_
UNION ALL



SELECT     dbo.OrdDetail.OrdDetail_DPA_,
	   '   Order Date: ' + CAST(dbo.tbOrder.OrderDate AS CHAR(12))  AS Security,
	    '' AS Quantity,
		'' AS Price,
		'' AS Validity,
		'' AS Certificate,
		'' [Slip No.],
		'' [Trade Date],
		'' [Trade Qty],
		'' [Trade Price],
		'' [Broker],
		'' [Contract]
FROM         dbo.tbOrder INNER JOIN
                      dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_

UNION ALL

SELECT    dbo.OrdDetail.OrdDetail_DPA_,
	  '  ' + dbo.Security.SecurityName Security,
	   CAST(dbo.OrdDetail.OrdDetailQty AS NVARCHAR(500)) Quantity,
	   CAST(dbo.OrdDetail.OrdDetailPrice AS NVARCHAR(500)) Price,
           CAST(dbo.OrdDetail.OrdDetailValidity AS CHAR(12)) Validity,
	   dbo.OrdDetail.OrdDetailCertNo Certificate,
	   '' [Slip No.],
	'' [Trade Date],
	'' [Trade Qty],
	'' [Trade Price],
	'' [Broker],
	'' [Contract]

FROM         dbo.tbOrder INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_

UNION ALL

SELECT  dbo.OrdDetail.OrdDetail_DPA_,
	'^' Security, '' Quantity, '' Price, '' Validity, '' Certificate,
	dbo.Lot.LotSlipNo [Slip No.],
	CAST(dbo.Lot.LotTDate AS CHAR(12)) [Trade Date],
	CAST(dbo.Lot.LotQty AS NVARCHAR(500)) [Trade Qty],
	CAST(dbo.Lot.LotPrice  AS NVARCHAR(500)) [Trade Price],
	dbo.Broker.BrokerName [Broker],
	dbo.Lot.ContractNumber [Contract]
FROM          dbo.Lot INNER JOIN
                      dbo.Broker ON dbo.Lot.Broker_DPA_ = dbo.Broker.Broker_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_) INNERTBL
ORDER BY  OrdDetail_DPA_, Security, CAST([Trade Date] AS DATETIME)) OUTERTBL
GO
CREATE VIEW dbo.ClientTestStatement
AS
SELECT     TOP 100 PERCENT COUNT(*) AS ClientTransaction_DPA_, a.Client_DPA_, a.TransDate, a.Ref, a.Particulars, a.Debit, a.Credit, a.Creditbal,
                      CASE a.IsOpeningBalance WHEN 1 THEN CASE WHEN a.Balance >= 0 THEN CONVERT(NVARCHAR(400), a.Balance)
                      + ' Cr' ELSE CONVERT(NVARCHAR(400), a.Balance) + ' Dr' END ELSE CASE WHEN SUM(b.Balance) >= 0 THEN CONVERT(NVARCHAR(400),
                      SUM(b.Balance)) + ' Cr' ELSE CONVERT(NVARCHAR(400), ABS(SUM(b.Balance))) + ' Dr' END END AS Balance, a.IsOpeningBalance,
                      a.PaymentReceiptNo
FROM         (SELECT     TOP 100 PERCENT *
                       FROM          dbo.ClientTransactionList) a CROSS JOIN
                          (SELECT     TOP 100 PERCENT *
                            FROM          dbo.ClientTransactionList) b
WHERE     a.TransDate >= b.TransDate AND a.Client_DPA_ = b.Client_DPA_
GROUP BY a.Client_DPA_, a.TransDate, a.Ref, a.Particulars, a.Debit, a.Credit, a.Balance, a.IsOpeningBalance, a.Creditbal, a.PaymentReceiptNo
ORDER BY a.Client_DPA_, a.IsOpeningBalance DESC, a.ClientTransaction_DPA_, a.Particulars, a.Ref, a.Debit, a.Credit, a.Balance

GO

CREATE VIEW dbo.ClientTestTransactionList
AS
SELECT     ClientsStatement.*, CreditBal - Debit AS Balance
FROM         (SELECT     Client_DPA_, ClientRegDate AS TransDate, '' AS REF, ' Opening Balance' AS Particulars, 0 AS Debit, 0 AS Credit,
                                              ClientOpeningBal AS CreditBal, 1 AS IsOpeningBalance, '' AS PaymentReceiptNo
                       FROM          dbo.Client
                       WHERE      deleted = 0
                       UNION
                       SELECT     dbo.Payment.Entity_DPA_ AS Client_DPA_, dbo.Payment.PaymentPDate AS TransDate, dbo.Payment.PaymentReference AS Ref,
                                             OrderType.OrderTypeDescription + ': ' + dbo.Payment.PaymentNarrative AS Particulars,
                                             CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS CreditBal, 0 AS IsOpeningBalance,
                                             dbo.Payment.PaymentReceiptNo
                       FROM         dbo.Payment INNER JOIN
                                             dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_ LEFT OUTER JOIN
                                             tborder ON Payment.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN
                                             OrderType ON tbOrder.OrderType_DPA_ = OrderType.Ordertype_DPA_
                       WHERE     EntityType_DPA_ = 1 AND Payment.Deleted = 0
                       UNION
                       SELECT     Client_DPA_, LotTDate AS TransDate, ContractNumber AS Ref, OrderType.OrderTypeDescription + ': ' + SecurityName AS Particulars,
                                             CASE ClientTransactionsSubList.OrderTypeSale WHEN 0 THEN NetAmt ELSE 0 END AS Debit,
                                             CASE ClientTransactionsSubList.OrderTypeSale WHEN 1 THEN NetAmt ELSE 0 END AS Credit,
                                             CASE ClientTransactionsSubList.OrderTypeSale WHEN 1 THEN NetAmt ELSE 0 END AS CreditBal, 0 AS IsOpeningBalance,
                                            '' As  PaymentReceiptNo
                       FROM         dbo.ClientTransactionsSubList LEFT OUTER JOIN
                                             dbo.OrderType ON dbo.ClientTransactionsSubList.OrderTypeSale = dbo.OrderType.OrderTypeSale
                       UNION
                       SELECT     dbo.JournalList.Entity_DPA_ AS Client_DPA_, dbo.JournalList.JournalDate AS TransDate,
                                             CAST(dbo.JournalList.JournalEntry_DPA_ AS Nvarchar(500)) AS Ref, dbo.JournalList.JournalNarrative AS Particulars,
                                             JournalList.JournalEntryDebit AS Debit, JournalList.JournalEntryCredit AS Credit, JournalList.JournalEntryCredit AS CreditBal,
                                             0 AS IsOpeningBalance, '' AS PaymentReceiptNo
                       FROM         dbo.JournalList
                       WHERE     EntityType_DPA_ = 1) ClientsStatement


GO

CREATE VIEW dbo.ClientTotals
AS
SELECT     SUM(ISNULL(BalanceQty * OrdDetailPrice, 0)) AS Total, Client_DPA_
FROM         (SELECT DISTINCT
                                              dbo.LotList.BalanceQty, CASE (dbo.OrdDetail.Best) WHEN 1 THEN dbo.datastream_SecurityPriceList.Price * 1.020825 ELSE CONVERT(MONEY,
                                              dbo.LotList.OrdDetailPrice) END AS OrdDetailPrice, dbo.LotList.Order_DPA_, dbo.LotList.Client_DPA_
                       FROM          dbo.LotList INNER JOIN
                                              dbo.OrdDetail ON dbo.LotList.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                                              dbo.datastream_SecurityPriceList ON dbo.OrdDetail.Security_DPA_ = dbo.datastream_SecurityPriceList.Security_DPA_
                       WHERE      (RTRIM(dbo.LotList.OrdDetailType) LIKE '%Purchase%') AND BalanceQty > 0) a
GROUP BY Client_DPA_






GO

CREATE VIEW dbo.ClientTradeVolumeBondsTotals
AS
SELECT     dbo.Client.Client_DPA_ AS [Client Code], dbo.Client.ClientName AS [Client Name], SUM(dbo.Lots.LotQty) AS Volume, SUM(LevyContract.LevyAmount)
                      AS Commission, SUM(dbo.Lots.LotGrossAmount) AS Gross, CAST(FLOOR(CAST(dbo.Lots.LotTDate AS float)) AS datetime) AS TransDate
FROM         dbo.Lots INNER JOIN
                      dbo.OrdDetail ON dbo.Lots.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                          (SELECT     Contract_DPA_, SUM(LevyAmount) AS LevyAmount
                            FROM          dbo.LevyContract
                            WHERE      (Deleted = 0)
                            GROUP BY Contract_DPA_) LevyContract ON dbo.Lots.Contract_DPA_ = LevyContract.Contract_DPA_
WHERE     (dbo.Client.Deleted = 0) AND (dbo.tbOrder.Deleted = 0)  AND (dbo.OrdDetail.Deleted = 0) AND
                      (dbo.tbOrder.OrderSecType_DPA_ <> 2)
GROUP BY dbo.Client.Client_DPA_, dbo.Client.ClientName, CAST(FLOOR(CAST(dbo.Lots.LotTDate AS float)) AS datetime)



GO

CREATE VIEW dbo.ClientTradeVolumeSharesTotals
AS
SELECT DISTINCT
                      dbo.Client.Client_DPA_ AS [Client Code], dbo.Client.ClientName AS [Client Name], SUM(dbo.Lots.LotQty) AS Volume, SUM(LevyContract.LevyAmount)
                      AS Commission, SUM(dbo.Lots.LotGrossAmount) AS Gross, CAST(FLOOR(CAST(dbo.Lots.LotTDate AS float)) AS datetime) AS TransDate
FROM         dbo.Lots INNER JOIN
                      dbo.OrdDetail ON dbo.Lots.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                          (SELECT     Contract_DPA_, SUM(LevyAmount) AS LevyAmount
                            FROM          dbo.LevyContract
                            WHERE      (Deleted = 0)
                            GROUP BY Contract_DPA_) LevyContract ON dbo.Lots.Contract_DPA_ = LevyContract.Contract_DPA_
WHERE     (dbo.Client.Deleted = 0) AND (dbo.tbOrder.Deleted = 0) AND (dbo.Lots.Deleted = 0) AND (dbo.OrdDetail.Deleted = 0) AND
                      (dbo.tbOrder.OrderSecType_DPA_ = 2)
GROUP BY dbo.Client.Client_DPA_, dbo.Client.ClientName,CAST(FLOOR(CAST(dbo.Lots.LotTDate AS float)) AS datetime)




GO


CREATE VIEW dbo.ClientTradeVolumeTotals
AS
SELECT     dbo.Client.Client_DPA_ AS [Client Code], dbo.Client.ClientName AS [Client Name], SUM(dbo.Lot.LotQty) AS Volume, SUM(dbo.LevyContract.LevyAmount)
                      AS Commission
FROM         dbo.Lot INNER JOIN
                      dbo.OrdDetail ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.LevyContract ON dbo.Lot.Contract_DPA_ = dbo.LevyContract.Contract_DPA_
GROUP BY dbo.Client.Client_DPA_, dbo.Client.ClientName



GO
CREATE VIEW dbo.ClientTransactionList
AS
SELECT     ClientsStatement.*, Credit - Debit AS Balance
FROM         (SELECT     Client_DPA_, ClientRegDate AS TransDate, '' AS REF, ' Opening Balance' AS Particulars, 0 AS Debit, ClientOpeningBal AS Credit,
                                              1 AS IsOpeningBalance
                       FROM          dbo.Client
                       UNION ALL
                       SELECT     dbo.Payment.Entity_DPA_ AS Client_DPA_, dbo.Payment.PaymentPDate AS TransDate, dbo.Payment.PaymentReference AS Ref,
                                             dbo.Payment.PaymentNarrative AS Particulars, CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit, 0 AS IsOpeningBalance
                       FROM         dbo.Payment INNER JOIN
                                             dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
                       WHERE     EntityType_DPA_ = 1
                       UNION ALL
                       SELECT     Client_DPA_, LotTDate AS TransDate, ContractNumber AS Ref, SecurityName AS Particulars,
                                             CASE OrderTypeSale WHEN 0 THEN NetAmt ELSE 0 END AS Debit, CASE OrderTypeSale WHEN 1 THEN NetAmt ELSE 0 END AS Credit,
                                             0 AS IsOpeningBalance
                       FROM         dbo.ClientTransactionsSubList
                       UNION ALL
                       SELECT     dbo.JournalList.Entity_DPA_ AS Client_DPA_, dbo.JournalList.JournalDate AS TransDate,
                                             CAST(dbo.JournalList.JournalEntry_DPA_ AS Nvarchar(500)) AS Ref, dbo.JournalList.JournalNarrative AS Particulars,
                                             JournalList.JournalEntryDebit AS Debit, JournalList.JournalEntryCredit AS Credit, 0 AS IsOpeningBalance
                       FROM         dbo.JournalList
                       WHERE     EntityType_DPA_ = 1
	          UNION ALL
                       SELECT     Offerings.Client_DPA_, CAST(FLOOR(CAST(Offerings.Offerings_Date AS float)) AS DateTime) AS TransDate, Offerings.PAL_No AS REF,
                                             CAST(FLOOR(CAST(Offerings.Alloted_Rights AS INT)) AS Char) + ' ' + CASE WHEN len(Security.SecurityName)
                                             > 20 THEN LEFT(Security.SecurityName, 20) + '...' ELSE Security.SecurityName END + ' @ ' + CAST(Offerings.Offering_Price AS char(5))
                                             AS Particulars, Offerings.Alloted_Rights * Offerings.Offering_Price AS Debit, 0 AS Credit, 0 AS IsOpeningBalance
                       FROM         Offerings INNER JOIN
                                             Security ON Offerings.Offering = Security.Security_DPA_
                       WHERE     (Offerings.Deleted = 0) AND (Offerings.Forward <> 1)

) ClientsStatement


GO
CREATE VIEW dbo.ClientTransactionsSubList
AS
/*
SELECT     CAST(LotQty AS nvarchar) + ' ' + SecurityName + ' @ ' + CAST(LotPrice AS nvarchar) AS SecurityName, ContractNumber, LotTDate, Client_DPA_,
                      LotGrossAmount AS Gross, OrderTypeSale, LevyAmount AS TotalLevies,
                      CASE OrderTypeSale WHEN 1 THEN LotGrossAmount - LevyAmount ELSE LotGrossAmount + LevyAmount END AS NetAmt,
                      ReturnableCommission
FROM         (SELECT     dbo.Lot.ContractNumber, dbo.Lot.LotTDate, dbo.tbOrder.Client_DPA_, dbo.Lot.LotGrossAmount, dbo.OrderType.OrderTypeSale,
                                              SUM(CASE LevyContract.SystemMaintained WHEN 12 THEN 0 ELSE LevyContract.LevyAmount END) AS LevyAmount, dbo.Lot.LotQty,
                                              dbo.Lot.LotPrice, dbo.Security.SecurityName,
                                              SUM(CASE LevyContract.SystemMaintained WHEN 12 THEN LevyContract.LevyAmount ELSE 0 END) AS ReturnableCommission
                       FROM          dbo.Lot INNER JOIN
                                              dbo.LevyContract ON dbo.Lot.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                                              dbo.OrdDetail ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                                              dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                                              dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                                              dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_
                       WHERE      dbo.tbOrder.OrderCanceled = 0 and dbo.Lot.Deleted <> 1
                       GROUP BY dbo.Lot.ContractNumber, dbo.Lot.LotTDate, dbo.tbOrder.Client_DPA_, dbo.OrderType.OrderTypeSale, dbo.Lot.LotGrossAmount,
                                              dbo.Lot.LotQty, dbo.Lot.LotPrice, dbo.Security.SecurityName) innerTable
*/


SELECT     CAST(LotQty AS nvarchar) + ' ' + SecurityName + ' @ ' + CAST(LotPrice AS nvarchar) AS SecurityName, ContractNumber, LotTDate, Client_DPA_,
                      LotGrossAmount AS Gross, OrderTypeSale, LevyAmount AS TotalLevies,
                      CASE OrderTypeSale WHEN 1 THEN LotGrossAmount - LevyAmount ELSE LotGrossAmount + LevyAmount END AS NetAmt,
                      ReturnableCommission
FROM         (SELECT     Lot.ContractNumber, Lot.LotTDate, tbOrder.Client_DPA_, Lot.LotGrossAmount, OrderType.OrderTypeSale,
                      SUM(CASE LevyContract.SystemMaintained WHEN 12 THEN 0 ELSE CASE LevyContract.SystemMaintained WHEN 25 THEN 0 ELSE CASE LevyContract.SystemMaintained
                       WHEN 99 THEN 0 ELSE LevyContract.LevyAmount + LevyContract.LevyVATAmount END END END) AS LevyAmount, Lot.LotQty, Lot.LotPrice,
                      Security.SecurityName, SUM(CASE LevyContract.SystemMaintained WHEN 12 THEN LevyContract.LevyAmount ELSE 0 END)
                      AS ReturnableCommission
FROM         Lot INNER JOIN

                  LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_ INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN
                      OrderType ON tbOrder.OrderType_DPA_ = OrderType.OrderType_DPA_ INNER JOIN
                      Security ON OrdDetail.Security_DPA_ = Security.Security_DPA_
WHERE     (tbOrder.OrderCanceled = 0) AND (Lot.Deleted <> 1)
GROUP BY Lot.ContractNumber, Lot.LotTDate, tbOrder.Client_DPA_, OrderType.OrderTypeSale, Lot.LotGrossAmount, Lot.LotQty, Lot.LotPrice,
                      Security.SecurityName
) innerTable



GO

CREATE VIEW dbo.ClientTransactionsSubList1
AS



SELECT     CAST(LotQty AS nvarchar) + ' ' + SecurityName + ' @ ' + CAST(LotPrice AS nvarchar) AS SecurityName, ContractNumber, LotTDate, Client_DPA_,
                      LotGrossAmount AS Gross, OrderTypeSale, LevyAmount AS TotalLevies, PaymentReceiptNo,
                      CASE OrderTypeSale WHEN 1 THEN LotGrossAmount - LevyAmount ELSE LotGrossAmount + LevyAmount END AS NetAmt
FROM         (SELECT     dbo.Lots.ContractNumber, dbo.Lots.LotTDate, dbo.tbOrder.Client_DPA_, dbo.Lots.LotGrossAmount, dbo.OrderType.OrderTypeSale,
                                              SUM(dbo.LevyContract.LevyAmount) AS LevyAmount, dbo.Lots.LotQty, dbo.Lots.LotPrice, dbo.Security.SecurityCode as SecurityName,
                                              dbo.Payment.PaymentReceiptNo
                       FROM          dbo.Lots INNER JOIN
                                              dbo.LevyContract ON dbo.Lots.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                                              dbo.OrdDetail ON dbo.Lots.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                                              dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                                              dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                                              dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ LEFT OUTER JOIN
                                              dbo.Payment ON dbo.tbOrder.Order_DPA_ = dbo.Payment.Order_DPA_
                       WHERE      (dbo.tbOrder.OrderCanceled = 0) AND LevyContract.SystemMaintained <> 8 AND LevyContract.SystemMaintained <> 12 AND
                                              LevyContract.Deleted = 0 AND OrdDetail.Deleted = 0 AND tbOrder.Deleted = 0 AND Lots.Deleted = 0
                       GROUP BY dbo.Lots.ContractNumber, dbo.Lots.LotTDate, dbo.tbOrder.Client_DPA_, dbo.OrderType.OrderTypeSale, dbo.Lots.LotGrossAmount,
                                              dbo.Lots.LotQty, dbo.Lots.LotPrice, dbo.Security.SecurityCode, dbo.Payment.PaymentReceiptNo) innerTable








GO
CREATE VIEW dbo.ClientVolumeTotals
AS
SELECT     dbo.Security.SecurityCode AS SecurityCode, dbo.Security.SecurityName AS SecurityName, SUM(dbo.Lot.LotQty) AS Volume,
                      SUM(dbo.Lot.LotGrossAmount) AS Gross, dbo.Client.Client_DPA_ AS ClientCode, dbo.Client.ClientName
FROM         dbo.Lot INNER JOIN
                      dbo.OrdDetail ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_
GROUP BY dbo.Security.SecurityCode, dbo.Security.SecurityName, dbo.Client.Client_DPA_, dbo.Client.ClientName

GO
CREATE VIEW dbo.ClientVoucherList
AS
SELECT     dbo.ClientVoucher.ClientVoucher_DPA_, SUM(dbo.ContractListWithNet.NetAmount) AS VoucherAmount, dbo.ClientVoucher.VoucherDate,
                      dbo.ClientVoucher.VoucherPaid, dbo.ClientList.ClientName, dbo.ClientList.Client_DPA_
FROM         dbo.ContractListWithNet INNER JOIN
                      dbo.ClientList ON dbo.ContractListWithNet.Client_DPA_ = dbo.ClientList.Client_DPA_ INNER JOIN
                      dbo.ClientVoucher ON dbo.ContractListWithNet.ClientVoucher_DPA_ = dbo.ClientVoucher.ClientVoucher_DPA_
GROUP BY dbo.ClientVoucher.ClientVoucher_DPA_, dbo.ClientVoucher.VoucherDate, dbo.ClientList.ClientName, dbo.ClientList.Client_DPA_,
                      dbo.ClientVoucher.VoucherPaid




GO


CREATE VIEW dbo.CommissionList
AS
/*
SELECT     TOP 100 PERCENT CommissionRate, Commission_DPA_, CommissionDescription, CommissionDescription + ' (' + CONVERT(nvarchar,
                      CommissionRate) + '; Bonds:  ' + CONVERT(nvarchar, BondCommission) + ')' AS CommissionDisplay, DefaultSelection, BondCommission,
                      SecurityBoundary, BondBoundary, UpperBondCommission, UpperSecurityCommission, MinimumBondCommission, MinimumSecurityCommission,
                      CMARegulated, Immobilised, SystemMaintained
FROM         dbo.Commission
ORDER BY CommissionRate*/

SELECT     TOP 100 PERCENT CommissionRate, Commission_DPA_, CommissionDescription, CommissionDescription + ' (' + CONVERT(nvarchar,
                      CommissionRate) + '; Bonds:  ' + CONVERT(nvarchar, BondCommission) + ')' AS CommissionDisplay, DefaultSelection, BondCommission,
                      SecurityBoundary, SecondSecurityBoundary, BondBoundary, SecondBondBoundary, UpperBondCommission, UpperSecurityCommission,
                      MedianSecurityCommission, MedianBondCommission, MinimumBondCommission, MinimumSecurityCommission, CMARegulated, Immobilised,
                      SystemMaintained
FROM         Commission
ORDER BY CommissionRate



GO

CREATE VIEW dbo.CommissionStatement
AS
SELECT     a.Entity_DPA_, a.TransDate, a.Ref, a.Particulars, a.EntityName, a.Credit, a.Debit, a.CreditBal AS Balance, 1 AS IsOpeningBalance
FROM         (SELECT     dbo.Entity.Entity_DPA_, CAST('3/1/2005' AS Datetime) AS TransDate, '' AS Ref, 'Opening Balance' AS Particulars, dbo.Entity.EntityName,
                                              0 AS Credit, 0 AS Debit, dbo.Entity.EntityOpeningBal AS CreditBal
                       FROM          dbo.Entity INNER JOIN
                                              dbo.EntityType ON dbo.Entity.EntityType_DPA_ = dbo.EntityType.EntityType_DPA_
                       WHERE      (dbo.Entity.EntityType_DPA_ = 6)) a
UNION
SELECT     TOP 100 PERCENT Entity.Entity_DPA_, CAST(FLOOR(CAST(Lots.LotTDate AS float)) AS datetime) AS TransDate, Lots.ContractNumber AS REF,
                      CAST(Lots.LotQty AS nvarchar) + ' ' + Security.SecurityCode + ' @ ' + CAST(Lots.LotPrice AS nvarchar) AS Particulars, Entity.EntityName,
                      LevyContract.LevyAmount AS Credit, 0 AS Debit, LevyContract.LevyAmount AS Balance, 0 AS IsOpeningBalance
FROM         LevyContract INNER JOIN
                      Contract ON LevyContract.Contract_DPA_ = Contract.Contract_DPA_ INNER JOIN
                      Lots ON Contract.Contract_DPA_ = Lots.Contract_DPA_ INNER JOIN
                      OrdDetail ON Lots.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      Security ON OrdDetail.Security_DPA_ = Security.Security_DPA_ INNER JOIN
                      Entity ON LevyContract.LevyShortName = Entity.EntityName OR LevyContract.LevyName = Entity.EntityName
WHERE     (Entity.Entity_DPA_ <> 1)
UNION
SELECT     TOP 100 PERCENT Entity_DPA_ AS Client_DPA_, cast(floor(cast(JournalDate AS float)) AS datetime) AS TransDate,
                      CAST(JournalEntry_DPA_ AS Nvarchar(500)) AS Ref, JournalNarrative AS Particulars, '', JournalEntryCredit AS Credit, JournalEntryDebit AS Debit,
                      JournalEntryCredit - JournalEntryDebit AS CreditBal, 0 AS IsOpeningBalance
FROM         JournalList
WHERE     (EntityType_DPA_ = 6)
UNION
SELECT     TOP 100 PERCENT a.*, a.Credit - a.Debit AS balance, 0 AS IsOpeningBalance
FROM         (SELECT     dbo.Payment.Entity_DPA_ AS Entity_DPA_, cast(floor(cast(dbo.Payment.PaymentPDate AS float)) AS datetime) AS TransDate,
                                              dbo.Payment.PaymentReference AS Ref, dbo.Payment.PaymentNarrative AS Particulars, Entity.EntityName AS entityname,
                                              CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit,
                                              CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit
                       FROM          dbo.Payment INNER JOIN
                                              dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_ INNER JOIN
                                              dbo.PaymentClientStatementNos ON Payment.Payment_DPA_ = PaymentClientStatementNos.Payment_DPA_ INNER JOIN
                                              Entity ON Payment.Entity_DPA_ = Entity.Entity_DPA_
                       WHERE      Payment.EntityType_DPA_ = 6 AND Payment.deleted = 0) a
order by IsOpeningBalance Desc






GO
CREATE VIEW dbo.CompanyInfoList AS SELECT CompanyInfo.SetupID, CompanyInfo.CompanyName, CompanyInfo.Address, CompanyInfo.City + " " +  CompanyInfo.StateOrProvince + " " +  CompanyInfo.Country AS Location, CompanyInfo.PhoneNumber, CompanyInfo.FaxNumber
FROM CompanyInfo

GO
CREATE VIEW dbo.CompanyInfoView
AS
SELECT     dbo.CompanyInfo.CompanyName, dbo.CompanyInfo.Address, dbo.CompanyInfo.City, dbo.CompanyInfo.StateOrProvince, dbo.CompanyInfo.PostalCode,
                      dbo.CompanyInfo.Country, dbo.CompanyInfo.PhoneNumber, dbo.CompanyInfo.FaxNumber, dbo.CompanyInfo.Photo,
                      dbo.Branch.BranchDescription AS Branch, dbo.CompanyInfo.Broker_DPA_, dbo.Broker.BrokerName, dbo.Broker.BrokerCode, dbo.Branch.Branch_DPA_
FROM         dbo.CompanyInfo INNER JOIN
                      dbo.Branch ON dbo.CompanyInfo.BranchID = dbo.Branch.Branch_DPA_ INNER JOIN
                      dbo.Broker ON dbo.CompanyInfo.Broker_DPA_ = dbo.Broker.Broker_DPA_


GO
CREATE VIEW dbo.CompletedOrderList
AS
SELECT     Order_DPA_ AS [Order No], OrderDate AS [Order Date], OrdDetailClient + '	[' + CONVERT(nvarchar(4000), Client_DPA_) + ']' AS Client,
                      OrdDetailSecurity AS Security, OrdDetailQty AS [Order Qty], OrdDetailPrice AS [Order Price], LotSlipNo AS [Slip No], LotTDate AS [Trade Date],
                      LotQty AS Qty, LotPrice AS Price, BrokerCode AS Broker, ContractNumber AS Contract
FROM         dbo.LotList
WHERE     (Order_DPA_ NOT IN
                          (SELECT DISTINCT Order_DPA_
                            FROM          dbo.LotList
                            WHERE      (BalanceQty > 0)))



GO

CREATE VIEW dbo.CompleteEntityList
AS
SELECT LTRIM(RTRIM(EntityName)) AS EntityName, EntityType, Entity_DPA_, EntityType_DPA_, EntityCode, EntityNameEx
FROM  dbo.ClientEntityList
UNION
SELECT LTRIM(RTRIM(EntityName)) AS EntityName, EntityType, Entity_DPA_, EntityType_DPA_, EntityCode, EntityNameEx
FROM  EntityList
UNION
SELECT LTRIM(RTRIM(EntityName)) AS EntityName, EntityType, Entity_DPA_, EntityType_DPA_, EntityCode, EntityNameEx
FROM  BrokerEntityList
UNION
SELECT LTRIM(RTRIM(EntityName)) AS EntityName, EntityType, Entity_DPA_, EntityType_DPA_, EntityCode, EntityNameEx
FROM  AgentEntityList
UNION
SELECT LTRIM(RTRIM(EntityName)) AS EntityName, EntityType, Entity_DPA_, EntityType_DPA_, EntityCode, EntityNameEx
FROM  OwnerEntityList
UNION
SELECT LTRIM(RTRIM(EntityName)) AS EntityName, EntityType, Entity_DPA_, EntityType_DPA_, EntityCode, EntityNameEx
FROM  AccountEntityList







GO

CREATE VIEW dbo.ComputerStatement
AS
SELECT     Entity_DPA_, CAST(FLOOR(CAST(JournalDate AS float)) AS DateTime) AS TransDate, CAST(JournalEntry_DPA_ AS Nvarchar(500)) AS Ref,
                      JournalNarrative AS Particulars, JournalEntryDebit AS Debit, JournalEntryCredit AS Credit, JournalEntryCredit AS CreditBal, 0 AS IsOpeningBalance,
                      '' AS PaymentReceiptNo, 4 AS receipttype
FROM         dbo.JournalList
WHERE     (EntityType_DPA_ = 11)




GO
CREATE VIEW dbo.cont_ContractTotals
AS
SELECT     Contract_DPA_,
                          (SELECT     LevyAmount
                            FROM          LevyContract
                            WHERE      (Contract_DPA_ = ProperTable.Contract_DPA_) AND (SystemMaintained = 12)) AS AgentComm,
                          (SELECT     LevyAmount
                            FROM          LevyContract
                            WHERE      (Contract_DPA_ = ProperTable.Contract_DPA_) AND (SystemMaintained = 100)) AS BasicFee,
                          (SELECT     LevyAmount
                            FROM          LevyContract
                            WHERE      (Contract_DPA_ = ProperTable.Contract_DPA_) AND (SystemMaintained = 11)) -
                          (SELECT     LevyAmount
                            FROM          LevyContract
                            WHERE      (Contract_DPA_ = ProperTable.Contract_DPA_) AND (SystemMaintained = 25)) AS BrokerComm,
                          (SELECT     LevyAmount
                            FROM          LevyContract
                            WHERE      (Contract_DPA_ = ProperTable.Contract_DPA_) AND (SystemMaintained = 25)) AS MSEComm,
                          (SELECT     LevyAmount
                            FROM          LevyContract
                            WHERE      (Contract_DPA_ = ProperTable.Contract_DPA_) AND (SystemMaintained = 99)) AS VAT,
                          (SELECT     LotGrossAmount
                            FROM          Lot
                            WHERE      (Contract_DPA_ = ProperTable.Contract_DPA_)) AS Gross, CASE WHEN
                          (SELECT     LEFT(ContractNumber, 1)
                            FROM          Lot
                            WHERE      (Contract_DPA_ = ProperTable.Contract_DPA_)) = 'P' THEN
                          (SELECT     LotGrossAmount
                            FROM          Lot
                            WHERE      (Contract_DPA_ = ProperTable.Contract_DPA_)) +
                          (SELECT     LevyAmount
                            FROM          LevyContract
                            WHERE      (Contract_DPA_ = ProperTable.Contract_DPA_) AND (SystemMaintained = 100)) +
                          (SELECT     LevyAmount
                            FROM          LevyContract
                            WHERE      (Contract_DPA_ = ProperTable.Contract_DPA_) AND (SystemMaintained = 11)) +
                          (SELECT     LevyAmount
                            FROM          LevyContract
                            WHERE      (Contract_DPA_ = ProperTable.Contract_DPA_) AND (SystemMaintained = 99)) ELSE
                          (SELECT     LotGrossAmount
                            FROM          Lot
                            WHERE      (Contract_DPA_ = ProperTable.Contract_DPA_)) -
                          ((SELECT     LevyAmount
                              FROM         LevyContract
                              WHERE     (Contract_DPA_ = ProperTable.Contract_DPA_) AND (SystemMaintained = 100)) +
                          (SELECT     LevyAmount
                            FROM          LevyContract
                            WHERE      (Contract_DPA_ = ProperTable.Contract_DPA_) AND (SystemMaintained = 11)) +
                          (SELECT     LevyAmount
                            FROM          LevyContract
                            WHERE      (Contract_DPA_ = ProperTable.Contract_DPA_) AND (SystemMaintained = 99))) END AS Net
FROM         dbo.Lot ProperTable

GO

CREATE VIEW dbo.ContractCompensationByDate
AS
SELECT     TOP 100 PERCENT CAST(FLOOR(CAST(LotTDate AS float)) AS datetime) AS TransDate, SUM(LotPrice) AS Gross, SUM(LevyAmount)
                      AS Commission
FROM         dbo.ContractCompensationFunds
GROUP BY CAST(FLOOR(CAST(LotTDate AS float)) AS datetime)
ORDER BY CAST(FLOOR(CAST(LotTDate AS float)) AS datetime)


GO

CREATE VIEW dbo.ContractCompensationFunds
AS
SELECT     dbo.Contract.Contract_DPA_, dbo.Contract.ContractTransferNo, dbo.Contract.ContractDeliveryDate, dbo.Lot.Lot_DPA_, dbo.Lot.LotSlipNo,
                      dbo.Lot.LotTDate, dbo.Lot.LotQty, dbo.Lot.LotGrossAmount AS LotPrice, dbo.Lot.ContractNumber, dbo.Contract.ContractDelivered,
                      dbo.Contract.ContractNCertificate, dbo.Contract.ContractNCDate, dbo.Contract.ContractNCDelivered, dbo.Contract.Voucher_DPA_,
                      dbo.Contract.ContractVouchered, dbo.Lot.ContractNumber AS Expr1, dbo.LevyContract.LevyAmount, dbo.LevyContract.LevyName,
                      dbo.tbOrder.OrderType_DPA_ AS OrdDetailType, dbo.Broker.BrokerCode, dbo.Client.ClientName AS OrdDetailClient,
                      dbo.Security.SecurityName AS OrdDetailSecurity, dbo.Security.OrderSecType_DPA_ AS OrdDetailSecType, dbo.Security.SecurityCode
FROM         dbo.Contract INNER JOIN
                      dbo.Lot ON dbo.Contract.Contract_DPA_ = dbo.Lot.Contract_DPA_ INNER JOIN
                      dbo.LevyContract ON dbo.Contract.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Broker ON dbo.Lot.Broker_DPA_ = dbo.Broker.Broker_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_
WHERE     (dbo.LevyContract.SystemMaintained = 15) AND (dbo.Client.Deleted = 0) AND (dbo.tbOrder.Deleted = 0) AND (dbo.LevyContract.Deleted = 0) AND
                      (dbo.Contract.Deleted = 0) AND (dbo.Lot.Deleted = 0) AND (dbo.OrdDetail.Deleted = 0)


GO

CREATE VIEW dbo.ContractCompounded
AS
SELECT DISTINCT
                      TOP 100 PERCENT a.Gross, a.Quantity, dbo.Security.Security_DPA_, dbo.Security.SecurityCode, dbo.Lots.LotPrice, a.Order_DPA_, a.TransDate
FROM         (SELECT     SUM(dbo.Lots.LotQty * dbo.Lots.LotPrice) AS Gross, dbo.tbOrder.Order_DPA_, SUM(dbo.Lots.LotQty) AS Quantity,
                                              CAST(FLOOR(CAST(dbo.Lots.LotTDate AS float)) AS datetime) AS TransDate
                       FROM          dbo.tbOrder INNER JOIN
                                              dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                                              dbo.Lots ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.Lots.OrdDetail_DPA_ INNER JOIN
                                              dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_
                       GROUP BY dbo.tbOrder.Order_DPA_, CAST(FLOOR(CAST(dbo.Lots.LotTDate AS float)) AS datetime)) a INNER JOIN
                      dbo.OrdDetail ON a.Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                      dbo.Lots ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.Lots.OrdDetail_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_
WHERE     (dbo.OrdDetail.Deleted = 0) AND (dbo.Lots.Deleted = 0)
ORDER BY a.TransDate DESC, a.Order_DPA_


GO
CREATE VIEW dbo.ContractCompoundedClients
AS
SELECT     dbo.Lot.Contract_DPA_, dbo.Lot.Lot_DPA_, dbo.Broker.BrokerCode, dbo.OrderType.OrderTypeDescription AS OrdDetailType,
                      dbo.OrderSecType.OrderSecTypeDisplayName AS OrdDetailSecType, dbo.OrderType.OrderTypeSale, dbo.Lot.LotSlipNo, dbo.Lot.LotTDate,
                      dbo.Lot.LotQty, dbo.Lot.LotPrice, dbo.Lot.ContractNumber, dbo.Client.ClientName AS OrdDetailClient,
                      dbo.Security.SecurityName + ' ' + ISNULL(dbo.Bond.BondIssue, '') AS OrdDetailSecurity, dbo.LevyContract.LevyName, dbo.tbOrder.OrderRef,
                      dbo.Client.ClientAddr, dbo.Client.Client_DPA_, dbo.Lot.LotGrossAmount AS GrossAmount, dbo.tbOrder.Order_DPA_,
                      dbo.LevyContract.SystemMaintained, CASE WHEN ISNULL(CHARINDEX('%', dbo.LevyContract.LevyRatePercentage), 0)
                      > 0 THEN '%' ELSE '' END AS LevyRatePercentage, dbo.LevyContract.LevyShortName,
                      CASE WHEN ltrim(rtrim(OrderSecType.OrderSecTypeDisplayName))
                      = 'Fixed' THEN LevyContract.LevyRate * Lot.LotGrossAmount / 100 ELSE LevyContract.LevyAmount END AS LevyAmount, dbo.Client.ClientCDSNo,
                      dbo.OrdDetail.OrdDetail_DPA_, dbo.LevyContract.LevyVATAmount, dbo.Lot.ContractSettlementDate
FROM         dbo.tbOrder INNER JOIN
                      dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                      dbo.Lot ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.Lot.OrdDetail_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                      dbo.OrderSecType ON dbo.tbOrder.OrderSecType_DPA_ = dbo.OrderSecType.OrderSecType_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.LevyContract ON dbo.Lot.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ LEFT OUTER JOIN
                      dbo.Bond ON dbo.OrdDetail.Bond_DPA_ = dbo.Bond.Bond_DPA_ LEFT OUTER JOIN
                      dbo.Broker ON dbo.Lot.Broker_DPA_ = dbo.Broker.Broker_DPA_
WHERE     (dbo.tbOrder.OrderCompounded = 1) AND (dbo.tbOrder.OrderHold = 0) AND (dbo.Client.Deleted = 0) AND (dbo.OrdDetail.Deleted = 0) AND
                      (dbo.tbOrder.Deleted = 0) AND (dbo.Lot.Deleted = 0) AND (dbo.tbOrder.OrderCanceled = 0) AND (dbo.LevyContract.Deleted = 0)

GO
CREATE VIEW dbo.ContractDisplayList
AS
SELECT     dbo.Lot.ContractNumber, dbo.Contract.Contract_DPA_ AS Contract_DPA_, dbo.OrdDetailList.OrdDetailType, dbo.Lot.LotPrice, dbo.Lot.LotQty,
                      dbo.Lot.LotSlipNo, dbo.Lot.LotTDate, dbo.Client.ClientName, dbo.Security.SecurityName, dbo.LevyContract.LevyName, dbo.LevyContract.LevyAmount,
                      dbo.OrdDetailList.OrderTypeSale, dbo.OrdDetailList.Order_DPA_
FROM         dbo.Security INNER JOIN
                      dbo.OrdDetailList ON dbo.Security.Security_DPA_ = dbo.OrdDetailList.Security_DPA_ LEFT OUTER JOIN
                      dbo.[Order] ON dbo.OrdDetailList.Order_DPA_ = dbo.[Order].Order_DPA_ LEFT OUTER JOIN
                      dbo.Client ON dbo.[Order].Client_DPA_ = dbo.Client.Client_DPA_ RIGHT OUTER JOIN
                      dbo.Contract INNER JOIN
                      dbo.Lot ON dbo.Contract.Contract_DPA_ = dbo.Lot.Contract_DPA_ INNER JOIN
                      dbo.LevyContract ON dbo.Contract.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ ON dbo.OrdDetailList.OrdDetail_DPA_ = dbo.Lot.OrdDetail_DPA_

GO
CREATE VIEW dbo.ContractLevyList
AS
SELECT     dbo.Contract.Contract_DPA_, dbo.Contract.ContractTransferNo, dbo.Contract.ContractCertificate, dbo.Contract.ContractDeliveryDate,
                      dbo.LotList.Lot_DPA_, dbo.LotList.BrokerCode, dbo.LotList.OrdDetailType, dbo.LotList.OrderTypeSale, dbo.LotList.LotSlipNo, dbo.LotList.LotTDate,
                      dbo.LotList.LotQty, dbo.LotList.LotPrice, dbo.LotList.ContractNumber, dbo.LotList.StatusDescription, dbo.LotList.OrdDetailClient,
                      dbo.Contract.ContractDelivered, dbo.LotList.OrdDetailSecurity, dbo.Contract.ContractNCertificate, dbo.Contract.ContractNCDate,
                      dbo.Contract.ContractNCDelivered, dbo.Contract.Voucher_DPA_, dbo.Contract.ContractVouchered, dbo.LotList.ContractNumber AS Expr1,
                      dbo.LevyContract.LevyAmount, dbo.LevyContract.LevyName
FROM         dbo.Contract INNER JOIN
                      dbo.LotList ON dbo.Contract.Contract_DPA_ = dbo.LotList.Contract_DPA_ INNER JOIN
                      dbo.LevyContract ON dbo.Contract.Contract_DPA_ = dbo.LevyContract.Contract_DPA_

GO
CREATE VIEW dbo.ContractLevyTotal
AS
/*
SELECT     Contract_DPA_, SUM(LevyAmount) AS LevyTotal
FROM         dbo.LevyContract
WHERE     (SystemMaintained <> 12) AND (SystemMaintained <> 8) AND (SystemMaintained <> 99)
GROUP BY Contract_DPA_
*/

SELECT     Contract_DPA_, SUM(case SystemMaintained when 25 then  LevyAmount else LevyAmount+ LevyVATAmount end) AS LevyTotal
FROM         dbo.LevyContract
WHERE     (SystemMaintained <> 12) AND (SystemMaintained <> 8) AND (SystemMaintained <> 99)
GROUP BY Contract_DPA_




GO
CREATE VIEW dbo.ContractList
AS
SELECT     dbo.Contract.Contract_DPA_, dbo.Contract.ContractTransferNo, dbo.Contract.ContractDeliveryDate, dbo.LotList.Lot_DPA_, dbo.LotList.BrokerCode,
                      dbo.LotList.OrdDetailType, dbo.LotList.OrderTypeSale, dbo.LotList.LotSlipNo, dbo.LotList.LotTDate, dbo.LotList.LotQty, dbo.LotList.LotPrice,
                      dbo.LotList.ContractNumber, dbo.LotList.StatusDescription, dbo.LotList.OrdDetailClient, dbo.Contract.ContractDelivered, dbo.LotList.OrdDetailSecurity,
                      dbo.Contract.ContractNCertificate, dbo.Contract.ContractNCDate, dbo.Contract.ContractNCDelivered, dbo.Contract.Voucher_DPA_,
                      dbo.Contract.ContractVouchered, dbo.LotList.LotQty * dbo.LotList.LotPrice AS ContractAmount, dbo.LotList.Client_DPA_,
                      dbo.OrdDetail.OrdDetailCertNo AS ContractCertificate, dbo.Contract.ClientVoucher_DPA_, dbo.Contract.ContractClientVouchered,
                      dbo.Contract.BrokerReceiptVoucher_DPA_, dbo.Contract.BrokerReceiptVouchered, dbo.LotList.CDSTransaction, dbo.Contract.IsInterBank,
                      dbo.InterTransfer.InterTransfer_DPA_ AS Intertransfer, DATEADD([day], 7, dbo.LotList.LotTDate) AS Settlementdate
FROM         dbo.Contract INNER JOIN
                      dbo.LotList ON dbo.Contract.Contract_DPA_ = dbo.LotList.Contract_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.LotList.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ LEFT OUTER JOIN
                      dbo.InterTransfer ON dbo.Contract.Contract_DPA_ = dbo.InterTransfer.Contract_DPA_
WHERE     (dbo.Contract.Deleted = 0 OR
                      dbo.Contract.Deleted IS NULL) AND (dbo.OrdDetail.Deleted = 0 OR
                      dbo.OrdDetail.Deleted IS NULL)

GO
CREATE VIEW dbo.ContractListWithNet
AS
SELECT     dbo.ContractList.*,
                      CAST(CASE dbo.ContractList.OrderTypeSale WHEN 1 THEN dbo.ContractList.ContractAmount - dbo.ContractLevyTotal.LevyTotal ELSE dbo.ContractList.ContractAmount
                       + dbo.ContractLevyTotal.LevyTotal END AS real) AS NetAmount
FROM         dbo.ContractList INNER JOIN
                      dbo.ContractLevyTotal ON dbo.ContractList.Contract_DPA_ = dbo.ContractLevyTotal.Contract_DPA_
GO
CREATE VIEW dbo.ContractNewCertList
AS
SELECT     dbo.Contract.Contract_DPA_, dbo.Contract.ContractTransferNo, dbo.Contract.ContractDeliveryDate, dbo.LotList.Lot_DPA_, dbo.LotList.BrokerCode,
                      dbo.LotList.OrdDetailType, dbo.LotList.OrderTypeSale, dbo.LotList.LotSlipNo, dbo.LotList.LotTDate, dbo.LotList.LotQty, dbo.LotList.LotPrice,
                      dbo.LotList.ContractNumber, dbo.LotList.StatusDescription, dbo.LotList.OrdDetailClient, dbo.Contract.ContractDelivered, dbo.LotList.OrdDetailSecurity,
                      dbo.Contract.ContractNCertificate, dbo.Contract.ContractNCDate, dbo.Contract.ContractNCDelivered, dbo.LotList.SecurityCode, dbo.Contract.Deleted,
                      DATEADD([day], 7, dbo.LotList.LotTDate) AS Settlementdate
FROM         dbo.Contract INNER JOIN
                      dbo.LotList ON dbo.Contract.Contract_DPA_ = dbo.LotList.Contract_DPA_
WHERE     (dbo.LotList.OrderTypeSale = 0) AND (dbo.Contract.Deleted = 0 OR
                      dbo.Contract.Deleted IS NULL)

GO

CREATE   VIEW app.Contracts AS
SELECT  Contract_DPA_           AS ContractId,
        Status_DPA_             AS StatusId,
        ContractSettlementDate  AS SettlementDate,
        IsInterBank             AS IsInterBank,
        Voucher_DPA_            AS VoucherId,
        ClientVoucher_DPA_      AS ClientVoucherId
FROM dbo.Contract
WHERE Deleted = 0 OR Deleted IS NULL;

GO

CREATE VIEW dbo.Contracts
AS
SELECT     dbo.Contract.*
FROM         dbo.Contract
WHERE     (Deleted = 0)


GO
CREATE VIEW dbo.ContractStamps
AS
SELECT     dbo.Contract.Contract_DPA_, dbo.Contract.ContractTransferNo, dbo.Contract.ContractDeliveryDate, dbo.Lot.Lot_DPA_, dbo.Lot.LotSlipNo,
                      dbo.Lot.LotTDate, dbo.Lot.LotQty, dbo.Lot.LotGrossAmount AS LotPrice, dbo.Lot.ContractNumber, dbo.Contract.ContractDelivered,
                      dbo.Contract.ContractNCertificate, dbo.Contract.ContractNCDate, dbo.Contract.ContractNCDelivered, dbo.Contract.Voucher_DPA_,
                      dbo.Contract.ContractVouchered, dbo.Lot.ContractNumber AS Expr1, dbo.LevyContract.LevyAmount, dbo.LevyContract.LevyName,
                      dbo.tbOrder.OrderType_DPA_ AS OrdDetailType, dbo.Broker.BrokerCode, dbo.Client.ClientName AS OrdDetailClient,
                      dbo.Security.SecurityName AS OrdDetailSecurity, dbo.Security.OrderSecType_DPA_ AS OrdDetailSecType, dbo.Security.SecurityCode
FROM         dbo.Contract INNER JOIN
                      dbo.Lot ON dbo.Contract.Contract_DPA_ = dbo.Lot.Contract_DPA_ INNER JOIN
                      dbo.LevyContract ON dbo.Contract.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Broker ON dbo.Lot.Broker_DPA_ = dbo.Broker.Broker_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_
WHERE     (dbo.LevyContract.SystemMaintained = 10)

GO

CREATE VIEW dbo.ContractStampsByDate
AS
SELECT     TOP 100 PERCENT CAST(FLOOR(CAST(LotTDate AS float)) AS datetime) AS TransDate, SUM(LotPrice) AS Gross, SUM(LevyAmount)
                      AS Commission
FROM         dbo.ContractStamps
GROUP BY CAST(FLOOR(CAST(LotTDate AS float)) AS datetime)
ORDER BY CAST(FLOOR(CAST(LotTDate AS float)) AS datetime)


GO
CREATE VIEW dbo.ControlAccountHistory
AS

SELECT
	'******1' AS [Entity_DPA_],
	GETDATE() AS TransDate,'' AS Ref, '' AS Particulars,
	CASE
		WHEN SUM(Type1.Balance) < 0 THEN 0 - SUM(Type1.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(Type1.Balance) >= 0 THEN SUM(Type1.Balance)
		ELSE 0 END AS Credit,
		0 AS IsOpeningBalance,
		SUM(Type1.Balance) AS Balance,
	(SELECT EntityTypeName FROM dbo.FullEntityTypeList WHERE EntityType_DPA_ = 1) AS EntityName

	 FROM


(SELECT     ClientsStatement.*, Credit - Debit AS Balance
                       FROM          (SELECT     Client_DPA_, ClientRegDate AS TransDate, '' AS REF, ' Opening Balance' AS Particulars, 0 AS Debit, ClientOpeningBal AS Credit,
                              1 AS IsOpeningBalance
       FROM          dbo.Client
       UNION ALL
       SELECT     dbo.Payment.Entity_DPA_ AS Client_DPA_, dbo.Payment.PaymentPDate AS TransDate, dbo.Payment.PaymentReference AS Ref,
                             dbo.Payment.PaymentNarrative AS Particulars,
                             CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.Payment INNER JOIN
                             dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
       WHERE     EntityType_DPA_ = 1
       UNION ALL
       SELECT     Client_DPA_, LotTDate AS TransDate, ContractNumber AS Ref, SecurityName AS Particulars,
                             CASE OrderTypeSale WHEN 0 THEN NetAmt ELSE 0 END AS Debit,
                             CASE OrderTypeSale WHEN 1 THEN NetAmt ELSE 0 END AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.ClientTransactionsSubList
       UNION ALL
       SELECT     dbo.JournalList.Entity_DPA_ AS Client_DPA_, dbo.JournalList.JournalDate AS TransDate, dbo.JournalList.JournalEntry_DPA_ AS Ref,
                             dbo.JournalList.JournalNarrative AS Particulars, JournalList.JournalEntryDebit AS Debit,
                             JournalList.JournalEntryCredit AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.JournalList
       WHERE     EntityType_DPA_ = 1) ClientsStatement) Type1

UNION ALL

SELECT
	'******2' AS [Account Code],
	GETDATE() AS TransDate,'' AS Ref, '' AS Particulars,
	CASE
		WHEN SUM(Type2.Balance) < 0 THEN 0 - SUM(Type2.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(Type2.Balance) >= 0 THEN SUM(Type2.Balance)
		ELSE 0 END AS Credit,
	0 AS IsOpeningBalance,
	SUM(Type2.Balance) AS Balance,
	(SELECT EntityTypeName FROM dbo.FullEntityTypeList WHERE EntityType_DPA_ = 2) AS AccountName
FROM
	(SELECT     AgentTransactions.*, Credit - Debit AS Balance
	FROM         (SELECT     Agent_DPA_, AgentRegDate AS TransDate, '' AS REF, ' Opening Balance' AS Particulars,
                              CASE WHEN AgentOpeningBal < 0 THEN (0 - AgentOpeningBal) ELSE 0 END AS Debit,
                              CASE WHEN AgentOpeningBal >= 0 THEN AgentOpeningBal ELSE 0 END AS Credit, 1 AS IsOpeningBalance
       FROM          dbo.Agent
       UNION ALL
       SELECT     dbo.Payment.Entity_DPA_ AS Agent_DPA_, dbo.Payment.PaymentPDate AS TransDate, dbo.Payment.PaymentReference AS Ref,
                             dbo.Payment.PaymentNarrative AS Particulars, CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.Payment INNER JOIN
                             dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
       WHERE     EntityType_DPA_ = 2
       UNION ALL
       SELECT     Agent_DPA_, LotTDate AS TransDate, ContractNumber AS Ref, SecurityName AS Particulars, 0 AS Debit, isnull(AgentCommission, 0)
                             AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.AgentCommissionList
       UNION ALL
	 SELECT     dbo.JournalList.Entity_DPA_ AS Agent_DPA_, dbo.JournalList.JournalDate AS TransDate, dbo.JournalList.JournalEntry_DPA_ AS Ref,
                             dbo.JournalList.JournalNarrative AS Particulars, JournalList.JournalEntryDebit AS Debit,
                             JournalList.JournalEntryCredit AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.JournalList
       WHERE     EntityType_DPA_ = 2) AgentTransactions) Type2


union all



SELECT
	'******3' AS [Account Code],
	GETDATE() AS TransDate,'' AS Ref, '' AS Particulars,
	CASE
		WHEN SUM(Type3.Balance) < 0 THEN 0 - SUM(Type3.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(Type3.Balance) >= 0 THEN SUM(Type3.Balance)
		ELSE 0 END AS Credit,
	0 AS IsOpeningBalance,
	 SUM(Type3.Balance) AS Balance,
	(SELECT EntityTypeName FROM dbo.FullEntityTypeList WHERE EntityType_DPA_ = 3) AS AccountName

	FROM
(SELECT     BrokerTransactions.*, Credit - Debit AS Balance
FROM         (SELECT     Broker_DPA_, BrokerRegDate AS TransDate, '' AS REF, ' Opening Balance' AS Particulars, case when BrokerOpeningBal < 0 then BrokerOpeningBal else 0 end AS Debit, case when BrokerOpeningBal >= 0 then BrokerOpeningBal else 0 end AS Credit,
                                              1 AS IsOpeningBalance
                       FROM          dbo.Broker
                       UNION ALL
                       SELECT     dbo.Payment.Entity_DPA_ AS Broker_DPA_, dbo.Payment.PaymentPDate AS TransDate, dbo.Payment.PaymentReference AS Ref,
                                             dbo.Payment.PaymentNarrative AS Particulars, CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit, 0 AS IsOpeningBalance
                       FROM         dbo.Payment INNER JOIN
                                             dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
                       WHERE     EntityType_DPA_ = 3
                       UNION ALL
                       SELECT     Broker_DPA_, LotTDate AS TransDate, ContractNumber AS Ref, SecurityName AS Particulars,
                                             CASE OrderTypeSale WHEN 1 THEN Gross ELSE 0 END AS Debit, CASE OrderTypeSale WHEN 0 THEN Gross ELSE 0 END AS Credit,
                                             0 AS IsOpeningBalance
                       FROM         dbo.BrokerTransactionsSubList
			UNION ALL
		       SELECT     dbo.JournalList.Entity_DPA_ AS Broker_DPA_, dbo.JournalList.JournalDate AS TransDate, dbo.JournalList.JournalEntry_DPA_ AS Ref,
		                             dbo.JournalList.JournalNarrative AS Particulars, JournalList.JournalEntryDebit AS Debit,
		                             JournalList.JournalEntryCredit AS Credit, 0 AS IsOpeningBalance
		       FROM         dbo.JournalList
		       WHERE     EntityType_DPA_ = 3) BrokerTransactions) Type3

UNION ALL

SELECT
	'******7' AS [Account Code],
	GETDATE() AS TransDate,'' AS Ref, '' AS Particulars,
	CASE
		WHEN SUM(Type7.Balance) < 0 THEN 0 - SUM(Type7.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(Type7.Balance) >= 0 THEN SUM(Type7.Balance)
		ELSE 0 END AS Credit,
	0 AS IsOpeningBalance,
	SUM(Type7.Balance) AS Balance,
	(SELECT EntityTypeName FROM dbo.FullEntityTypeList WHERE EntityType_DPA_ = 7) AS AccountName
FROM
	(SELECT     OwnerTransactions.*, Credit - Debit AS Balance
	FROM         (SELECT     Owner_DPA_, OwnerRegDate AS TransDate, '' AS REF, ' Opening Balance' AS Particulars,
                              CASE WHEN OwnerOpeningBal < 0 THEN (0 - OwnerOpeningBal) ELSE 0 END AS Debit,
                              CASE WHEN OwnerOpeningBal >= 0 THEN OwnerOpeningBal ELSE 0 END AS Credit, 1 AS IsOpeningBalance
       FROM          dbo.Owner
       UNION ALL
       SELECT     dbo.Payment.Entity_DPA_ AS Owner_DPA_, dbo.Payment.PaymentPDate AS TransDate, dbo.Payment.PaymentReference AS Ref,
                             dbo.Payment.PaymentNarrative AS Particulars, CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.Payment INNER JOIN
                             dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
       WHERE     EntityType_DPA_ = 7
       UNION ALL
       SELECT     Owner_DPA_, LotTDate AS TransDate, ContractNumber AS Ref, SecurityName AS Particulars, 0 AS Debit, isnull(OwnerCommission, 0)
                             AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.OwnerCommissionList
       UNION ALL
	 SELECT     dbo.JournalList.Entity_DPA_ AS Owner_DPA_, dbo.JournalList.JournalDate AS TransDate, dbo.JournalList.JournalEntry_DPA_ AS Ref,
                             dbo.JournalList.JournalNarrative AS Particulars, JournalList.JournalEntryDebit AS Debit,
                             JournalList.JournalEntryCredit AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.JournalList
       WHERE     EntityType_DPA_ = 7) OwnerTransactions) Type7



UNION ALL
SELECT
	'******' + CAST(dbo.FullEntityTypeList.EntityType_DPA_ AS VARCHAR(500)) AS [Account Code],
	GETDATE() AS TransDate,'' AS Ref, '' AS Particulars,
	CASE
		WHEN SUM(dbo.EntityTransactionList.Balance) < 0 THEN 0 - SUM(dbo.EntityTransactionList.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.EntityTransactionList.Balance) >= 0 THEN SUM(dbo.EntityTransactionList.Balance)
		ELSE 0 END AS Credit,
	0 AS IsOpeningBalance,
        SUM(dbo.EntityTransactionList.Balance)  AS Balance,
	dbo.FullEntityTypeList.EntityTypeName AS AccountName

FROM         dbo.Entity INNER JOIN
                      dbo.FullEntityTypeList ON dbo.Entity.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_ INNER JOIN
                      dbo.EntityTransactionList ON dbo.Entity.Entity_DPA_ = dbo.EntityTransactionList.Entity_DPA_
GROUP BY dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName, dbo.FullEntityTypeList.EntityType_DPA_

GO
CREATE VIEW CPaymentList AS SELECT CPayment.CPaymentAmount AS CPaymentAmount, CPayment.CPaymentPDate AS CPaymentPDate, CPayment.CPaymentReference + ', ' + Client.ClientName AS CPaymentPayment, CPayment.CPayment_DPA_ AS CPayment_DPA_
FROM Client INNER JOIN CPayment ON Client.Client_DPA_=CPayment.Client_DPA_

GO
CREATE VIEW dbo.CPortfolioList
AS
SELECT     TOP 100 PERCENT dbo.Client.ClientName AS CPortfolioClient, dbo.CPortfolio.CPortfolioPrice, dbo.CPortfolio.CPortfolioQty,
                      dbo.Security.SecurityName AS CPortfolioSecurity, dbo.CPortfolio.CPortfolio_DPA_, dbo.CPortfolio.Client_DPA_, dbo.CPortfolio.Security_DPA_,
                      dbo.CPortfolio.CPortfolioPDate, dbo.CPortfolio.Reference
FROM         dbo.Security INNER JOIN
                      dbo.Client INNER JOIN
                      dbo.CPortfolio ON dbo.Client.Client_DPA_ = dbo.CPortfolio.Client_DPA_ ON dbo.Security.Security_DPA_ = dbo.CPortfolio.Security_DPA_
WHERE     (dbo.Client.Deleted = 0) OR
                      (dbo.Client.Deleted IS NULL)
ORDER BY dbo.Client.ClientName

GO
CREATE VIEW CReceiptsList AS SELECT CPayment.CPaymentAmount AS CPaymentAmount, CPayment.CPaymentPDate AS CPaymentPDate, CPayment.CPaymentReference + ', ' + Client.ClientName AS CPaymentPayment, CPayment.CPayment_DPA_ AS CPayment_DPA_
FROM Client INNER JOIN CPayment ON Client.Client_DPA_ = CPayment.Client_DPA_

GO

CREATE VIEW dbo.Creditors
AS
SELECT Client_DPA_, LastDate, Balance
FROM  dbo.DebtorCreditor
WHERE (Balance > 0)


GO

CREATE VIEW dbo.CurrentBalances
AS
SELECT     SUM(ISNULL(dbo.StatementList.Credit - dbo.StatementList.Debit, 0)) + dbo.Client.ClientOpeningBal AS CurrentBal, dbo.StatementList.Client_DPA_
FROM         dbo.StatementList INNER JOIN
                      dbo.Client ON dbo.StatementList.Client_DPA_ = dbo.Client.Client_DPA_
WHERE     (dbo.Client.Deleted = 0)
GROUP BY dbo.StatementList.Client_DPA_, dbo.Client.ClientOpeningBal


GO
CREATE VIEW dbo.CustodianCharges_Final
AS

SELECT     CustodianNetCharges.Agent_DPA_, CustodianNetCharges.AgentName, CustodianNetCharges.Client_DPA_, CustodianNetCharges.ClientName,
                      CustodianNetCharges.[Date], CustodianNetCharges.Type, CustodianNetCharges.Security, CustodianNetCharges.Broker, CustodianNetCharges.Contract,
                      CustodianNetCharges.Slip, CustodianNetCharges.Price, CustodianNetCharges.Qty, CustodianNetCharges.Gross,
                      CustodianTotalCharges.[Total Charges] AS TotalCharges, CustodianNetCharges.[Proper Charges] AS ProperCharges,
                      CustodianTotalCharges.[Total Charges] - CustodianNetCharges.[Proper Charges] AS [AdditionalCharge]
FROM         (SELECT     TOP 100 PERCENT Agent.Agent_DPA_, Agent.AgentName, tborder.Client_DPA_, Client.ClientName, Lot.LotTDate AS Date,
                                              OrderType.OrderTypeDescription AS Type, Security.SecurityCode AS Security, Broker.BrokerCode AS Broker,
                                              Lot.ContractNumber AS Contract, Lot.LotSlipNo AS Slip, Lot.LotPrice AS Price, Lot.LotQty AS Qty, Lot.LotGrossAmount AS Gross,
                                              SUM(CASE WHEN LevyContract.SystemMaintained = 7 THEN ROUND(LevyContract.LevyAmount / 2, 2)
                                              ELSE LevyContract.LevyAmount END) AS [Proper Charges]
                       FROM          Lot INNER JOIN
                                              OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                                              tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN
                                              Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
                                              LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_ INNER JOIN
                                              OrderType ON tbOrder.OrderType_DPA_ = OrderType.OrderType_DPA_ INNER JOIN
                                              Security ON OrdDetail.Security_DPA_ = Security.Security_DPA_ INNER JOIN
                                              Broker ON Lot.Broker_DPA_ = Broker.Broker_DPA_ LEFT OUTER JOIN
                                              Agent ON Client.Agent_DPA_ = Agent.Agent_DPA_
                       WHERE      (Lot.Deleted <> 1) AND (LevyContract.Deleted <> 1) AND (LevyContract.SystemMaintained <> 12) AND
                                              (LevyContract.SystemMaintained <> 16) AND (tbOrder.IsCustodian = 1)
                       GROUP BY Agent.Agent_DPA_, Agent.AgentName, tborder.Client_DPA_, Client.ClientName, Lot.LotTDate, OrderType.OrderTypeDescription,
                                              Security.SecurityCode, Broker.BrokerCode, Lot.ContractNumber, Lot.LotSlipNo, Lot.LotPrice, Lot.LotQty, Lot.LotGrossAmount,
                                              Lot.Contract_DPA_
                       ORDER BY Lot.Contract_DPA_) CustodianNetCharges INNER JOIN
                          (SELECT     TOP 100 PERCENT Lot.ContractNumber AS Contract, SUM(LevyContract.LevyAmount) AS [Total Charges]
                            FROM          Lot INNER JOIN
                                                   OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                                                   tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN
                                                   Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
                                                   LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_ INNER JOIN
                                                   OrderType ON tbOrder.OrderType_DPA_ = OrderType.OrderType_DPA_ INNER JOIN
                                                   Security ON OrdDetail.Security_DPA_ = Security.Security_DPA_ INNER JOIN
                                                   Broker ON Lot.Broker_DPA_ = Broker.Broker_DPA_ LEFT OUTER JOIN
                                                   Agent ON Client.Agent_DPA_ = Agent.Agent_DPA_
                            WHERE      (Lot.Deleted <> 1) AND (LevyContract.Deleted <> 1) AND (LevyContract.SystemMaintained <> 16) AND (tbOrder.IsCustodian = 1)
                            GROUP BY Lot.ContractNumber, Lot.Contract_DPA_
                            ORDER BY Lot.Contract_DPA_) CustodianTotalCharges ON CustodianNetCharges.Contract = CustodianTotalCharges.Contract

GO
CREATE VIEW dbo.CustodianCharges_ProperCharges
AS

SELECT   top 100 percent  Lot.LotTDate AS Date, OrderType.OrderTypeDescription AS Type, Security.SecurityCode AS Security, Broker.BrokerCode AS Broker,
                      Lot.ContractNumber AS Contract, Lot.LotSlipNo AS Slip, Lot.LotPrice AS Price, Lot.LotQty AS Qty, Lot.LotGrossAmount AS Gross,
                      SUM(CASE WHEN LevyContract.SystemMaintained = 7 THEN ROUND(LevyContract.LevyAmount / 2, 2) ELSE LevyContract.LevyAmount END)
                      AS [Proper Charges]
FROM         Lot INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN
                      Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_ INNER JOIN
                      OrderType ON tbOrder.OrderType_DPA_ = OrderType.OrderType_DPA_ INNER JOIN
                      Security ON OrdDetail.Security_DPA_ = Security.Security_DPA_ INNER JOIN
                      Broker ON Lot.Broker_DPA_ = Broker.Broker_DPA_ LEFT OUTER JOIN
                      Agent ON Client.Agent_DPA_ = Agent.Agent_DPA_
WHERE     (Lot.Deleted <> 1) AND (LevyContract.Deleted <> 1) AND (LevyContract.SystemMaintained <> 12) AND (LevyContract.SystemMaintained <> 16) AND
                      (tbOrder.IsCustodian = 1)
GROUP BY Lot.LotTDate, OrderType.OrderTypeDescription, Security.SecurityCode, Broker.BrokerCode, Lot.ContractNumber, Lot.LotSlipNo, Lot.LotPrice,
                      Lot.LotQty, Lot.LotGrossAmount, Lot.Contract_DPA_
ORDER BY Lot.Contract_DPA_
GO
CREATE VIEW dbo.CustodianClient
AS
SELECT     TOP 100 PERCENT dbo.Client.Client_DPA_, dbo.Client.ClientName
FROM         dbo.tbOrder INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_
WHERE     (dbo.tbOrder.IsCustodian = 1) AND (dbo.tbOrder.Deleted = 0) AND (dbo.Client.Deleted = 0)
GROUP BY dbo.Client.Client_DPA_, dbo.Client.ClientName
ORDER BY dbo.Client.ClientName

GO
CREATE VIEW dbo.CustodianCommission
AS
SELECT     TOP 100 PERCENT Contract_DPA_, TransDate, OrderTypeDescription, SecurityName, BrokerName, ContractNumber, LotSlipNo, LotPrice, LotQty,
                      CAST(SUM(HalfCompensationFund + TheOthers) AS FLOAT) AS NetCharges, LotGrossAmount, Client_DPA_, AgentName, Agent_DPA_,
                      SecurityCode
FROM         (SELECT     TOP 100 PERCENT Lots.ContractNumber, CAST(FLOOR(CAST(Lots.LotTDate AS Float)) AS DateTime) AS TransDate, tbOrder.Client_DPA_,
                                              CAST(Lots.LotQty AS nvarchar) + ' ' + Security.SecurityCode + ' @ ' + CAST(Lots.LotPrice AS nvarchar) AS SecurityName,
                                              Payments.PaymentReceiptNo, LevyContract.Contract_DPA_, Client.ClientName, Lots.LotGrossAmount, Security.SecurityCode,
                                              OrderType.OrderTypeDescription, Broker.BrokerCode AS BrokerName, Lots.LotSlipNo, Lots.LotQty, Lots.LotPrice, Agent.Agent_DPA_,
                                              Agent.AgentName, 0.5 * LevyContract.LevyAmount AS HalfCompensationFund, 0 AS TheOthers
                       FROM          Lots INNER JOIN
                                              LevyContract ON Lots.Contract_DPA_ = LevyContract.Contract_DPA_ INNER JOIN
                                              OrdDetail ON Lots.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                                              tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN
                                              OrderType ON tbOrder.OrderType_DPA_ = OrderType.OrderType_DPA_ INNER JOIN
                                              Security ON OrdDetail.Security_DPA_ = Security.Security_DPA_ INNER JOIN
                                              Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ LEFT OUTER JOIN
                                              Agent ON Client.Agent_DPA_ = Agent.Agent_DPA_ LEFT OUTER JOIN
                                              Broker ON Lots.Broker_DPA_ = Broker.Broker_DPA_ LEFT OUTER JOIN
                                              Payments ON tbOrder.Order_DPA_ = Payments.Order_DPA_
                       WHERE      (tbOrder.OrderCanceled = 0) AND (LevyContract.SystemMaintained = 7) AND (Client.Deleted = 0) AND (OrdDetail.Deleted = 0) AND
                                              (tbOrder.Deleted = 0) AND (LevyContract.Deleted = 0)
                       GROUP BY Lots.ContractNumber, Lots.LotTDate, tbOrder.Client_DPA_, LevyContract.LevyAmount, Lots.LotQty, Security.SecurityCode,
                                              Lots.LotGrossAmount, Lots.LotPrice, Payments.PaymentReceiptNo, LevyContract.Contract_DPA_, Client.ClientName,
                                              OrderType.OrderTypeDescription, Broker.BrokerCode, Lots.LotSlipNo, Agent.Agent_DPA_, Agent.AgentName
                       UNION ALL
                       SELECT DISTINCT
                                             TOP 100 PERCENT dbo.Lots.ContractNumber, CAST(FLOOR(CAST(dbo.Lots.LotTDate AS Float)) AS DateTime) AS TransDate,
                                             dbo.tbOrder.Client_DPA_, CAST(dbo.Lots.LotQty AS nvarchar) + ' ' + dbo.Security.SecurityCode + ' @ ' + CAST(dbo.Lots.LotPrice AS nvarchar)
                                             AS SecurityName, dbo.Payments.PaymentReceiptNo, dbo.LevyContract.Contract_DPA_, dbo.Client.ClientName, dbo.Lots.LotGrossAmount,
                                             dbo.Security.SecurityCode, dbo.OrderType.OrderTypeDescription, dbo.Broker.BrokerCode AS BrokerName, dbo.Lots.LotSlipNo,
                                             dbo.Lots.LotQty, dbo.Lots.LotPrice, dbo.Agent.Agent_DPA_, dbo.Agent.AgentName, 0 AS HalfCompensationFund,
                                             SUM(dbo.LevyContract.LevyAmount) AS TheOthers
                       FROM         dbo.Lots INNER JOIN
                                             dbo.LevyContract ON dbo.Lots.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                                             dbo.OrdDetail ON dbo.Lots.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                                             dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                                             dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                                             dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                                             dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ LEFT OUTER JOIN
                                             dbo.Agent ON dbo.Client.Agent_DPA_ = dbo.Agent.Agent_DPA_ LEFT OUTER JOIN
                                             dbo.Broker ON dbo.Lots.Broker_DPA_ = dbo.Broker.Broker_DPA_ LEFT OUTER JOIN
                                             dbo.Payments ON dbo.tbOrder.Order_DPA_ = dbo.Payments.Order_DPA_
                       WHERE     (dbo.tbOrder.OrderCanceled = 0) AND (dbo.LevyContract.SystemMaintained IN (6, 5, 10, 11)) AND (dbo.Client.Deleted = 0) AND
                                             (dbo.OrdDetail.Deleted = 0) AND (dbo.tbOrder.Deleted = 0) AND (dbo.LevyContract.Deleted = 0)
                       GROUP BY dbo.Lots.ContractNumber, dbo.Lots.LotTDate, dbo.tbOrder.Client_DPA_, dbo.LevyContract.LevyAmount, dbo.Lots.LotQty,
                                             dbo.Security.SecurityCode, dbo.Lots.LotGrossAmount, dbo.Lots.LotPrice, dbo.Payments.PaymentReceiptNo,
                                             dbo.LevyContract.Contract_DPA_, dbo.Client.ClientName, dbo.OrderType.OrderTypeDescription, dbo.Broker.BrokerCode, dbo.Lots.LotSlipNo,
                                             dbo.Agent.Agent_DPA_, dbo.Agent.AgentName) DERIVEDTBL
GROUP BY Contract_DPA_, TransDate, OrderTypeDescription, SecurityName, BrokerName, ContractNumber, LotSlipNo, LotPrice, LotQty, LotGrossAmount,
                      Client_DPA_, AgentName, Agent_DPA_, SecurityCode

GO

CREATE VIEW dbo.Datastream_MuhammedPriceUpdate
AS
SELECT     TOP 100 PERCENT dbo.datastream_Securities.SecNameShort AS Security, dbo.datastream_Market.MktDate AS [Date],
                      dbo.datastream_Market.MktClose AS Price, dbo.datastream_Market.MktVolume AS Volume
FROM         dbo.datastream_Securities INNER JOIN
                      dbo.datastream_Market ON dbo.datastream_Securities.SecCode = dbo.datastream_Market.MktCode
WHERE     (NOT (dbo.datastream_Market.MktClose IS NULL))
ORDER BY dbo.datastream_Securities.SecNameShort, dbo.datastream_Market.MktDate


GO
CREATE VIEW dbo.datastream_SecurityPriceList
AS
SELECT DISTINCT
                      TOP 100 PERCENT dsM.MktDate AS [Date], ISNULL(dsM.MktClose, 0) AS Price, dbo.Security.Security_DPA_, dbo.Security.SecurityCode,
                      dbo.datastream_Securities.SecCode, dbo.Security.SecurityName, dsM.MktUnique
FROM         dbo.datastream_Market dsM INNER JOIN
                      dbo.datastream_Securities ON dsM.MktCode = dbo.datastream_Securities.SecCode INNER JOIN
                      dbo.MarketView ON FLOOR(CAST(dsM.MktDate AS float)) = dbo.MarketView.MktUnique AND dsM.MktCode = dbo.MarketView.Code RIGHT OUTER JOIN
                      dbo.Security ON dbo.datastream_Securities.SecKnow_DPA = dbo.Security.Security_DPA_
ORDER BY dbo.Security.Security_DPA_

GO
CREATE VIEW dbo.datastream_SecurityPriceListAllPrices
AS
SELECT     TOP 100 PERCENT dbo.Security.SecurityCode, dbo.Security.Security_DPA_, dsM.MktDate AS [Date], ISNULL(dsM.MktClose, 0) AS Price, dsM.MktDate,
                      dsM.MktClose, dbo.Security.SecurityName, dbo.datastream_Securities.SecCode, dsM.MktUnique
FROM         dbo.datastream_Market dsM INNER JOIN
                      dbo.datastream_Securities ON dsM.MktCode = dbo.datastream_Securities.SecCode RIGHT OUTER JOIN
                      dbo.Security ON dbo.datastream_Securities.SecKnow_DPA = dbo.Security.Security_DPA_
WHERE     (dsM.MktUnique =
                          (SELECT     TOP 1 MktUnique
                            FROM          dbo.datastream_Market
                            WHERE      MktCode = dsm.MktCode AND MktClose > 0
                            ORDER BY MktUnique DESC)) OR
                      (dbo.datastream_Securities.SecCode IS NULL)
ORDER BY dbo.Security.SecurityCode DESC

GO

CREATE VIEW dbo.datastream_SecurityPriceListLastestUpdate
AS
SELECT DISTINCT TOP 100 PERCENT dbo.datastream_Securities.SecKnow_DPA AS [BrokerKnow DPA], MAX(dbo.datastream_Market.MktDate) AS Date
FROM         dbo.Security INNER JOIN
                      dbo.datastream_Securities ON dbo.Security.Security_DPA_ = dbo.datastream_Securities.SecKnow_DPA INNER JOIN
                      dbo.datastream_Market ON dbo.datastream_Securities.SecCode = dbo.datastream_Market.MktCode
WHERE     (NOT (dbo.datastream_Market.MktClose IS NULL))
GROUP BY dbo.datastream_Securities.SecKnow_DPA
ORDER BY MAX(dbo.datastream_Market.MktDate) DESC


GO

CREATE VIEW dbo.datastream_SecurityPriceListz
AS
SELECT TOP 100 PERCENT dbo.datastream_SecurityPriceListAllPrices.SecurityCode, dbo.datastream_SecurityPriceListAllPrices.Security_DPA_,
               dbo.datastream_SecurityPriceListAllPrices.MktDate AS [Date], dbo.datastream_SecurityPriceListAllPrices.MktClose AS Price
FROM  dbo.datastream_SecurityPriceListAllPrices INNER JOIN
               dbo.datastream_SecurityPriceListLastestUpdate ON
               dbo.datastream_SecurityPriceListAllPrices.Security_DPA_ = dbo.datastream_SecurityPriceListLastestUpdate.[BrokerKnow DPA] AND
               dbo.datastream_SecurityPriceListAllPrices.MktDate = dbo.datastream_SecurityPriceListLastestUpdate.[Date]
ORDER BY dbo.datastream_SecurityPriceListAllPrices.SecurityCode


GO

CREATE VIEW dbo.DB_Affirmation
AS
SELECT     TOP 100 PERCENT a.*, Owner.OwnerFName + ' ' + OwnerLName AS AccountManager
FROM         (SELECT     TOP 100 PERCENT dbo.Security.Security_DPA_, dbo.Lots.LotSlipNo, dbo.Lots.ContractNumber, dbo.Lots.LotQty, dbo.Lots.LotPrice,
                                              dbo.tbOrder.OrderRef, dbo.tbOrder.Order_DPA_, dbo.Client.ClientAddr AS AccountAddress, dbo.Client.ClientName AS Account,
                                              Cast(Floor(Cast(dbo.Lots.LotTDate AS Float)) AS DateTime) AS LotTDate, dbo.Client.Client_DPA_, dbo.Client.ClientCDSNo,
                                              dbo.Lots.LotGrossAmount AS SettlementGrossAmount, ISNULL(dbo.Client.ClientContact, '') AS Owner, dbo.Client.ClientFax,
                                              dbo.Security.SecurityName, BalanceQty = CASE ISNULL(dbo.OrdDetailContractedQtyList.ContractQty, 1)
                                              WHEN 1 THEN dbo.OrdDetail.OrdDetailQty ELSE dbo.OrdDetail.OrdDetailQty - dbo.OrdDetailContractedQtyList.ContractQty END,
                                              Security.SecurityName AS ordDetailSecurity, Class.ClassDescription, Contracts.ContractSettlementDate, Owner.Owner_DPA_,
                                              OrderType.OrderTypeSale,
                                              CASE dbo.OrderType.OrderTypeSale WHEN 0 THEN 'PURCHASE OF ' + dbo.Security.SecurityName ELSE 'SALE OF ' + dbo.Security.SecurityName
                                               END AS ReferenceHeader, CAST(CASE dbo.OrderType.OrderTypeSale WHEN 1 THEN cast(dbo.Lots.LotGrossAmount AS float)
                                              - Cast(dbo.ContractLevyTotal.LevyTotal AS float) ELSE cast(dbo.Lots.LotGrossAmount AS float)
                                              + cast(dbo.ContractLevyTotal.LevyTotal AS float) END AS float) AS SettlementAmount, Client.Class_DPA_
                       FROM          Lots INNER JOIN
                                              OrdDetail ON Lots.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                                              OrdDetailContractedQtyList ON OrdDetail.OrdDetail_DPA_ = OrdDetailContractedQtyList.OrdDetail_DPA_ INNER JOIN
                                              tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN
                                              Contracts ON Lots.Contract_DPA_ = Contracts.Contract_DPA_ INNER JOIN
                                              Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
                                              ContractLevyTotal ON Lots.Contract_DPA_ = ContractLevyTotal.Contract_DPA_ INNER JOIN
                                              Security ON OrdDetail.Security_DPA_ = Security.Security_DPA_ INNER JOIN
                                              Class ON Client.Class_DPA_ = Class.Class_DPA_ INNER JOIN
                                              OrderType ON tbOrder.OrderType_DPA_ = OrderType.OrderType_DPA_ LEFT OUTER JOIN
                                              Owner ON Client.Owner_DPA_ = Owner.Owner_DPA_
                       WHERE      (tbOrder.Deleted = 0) AND (Client.Deleted = 0) AND (OrdDetail.Deleted = 0)
                       ORDER BY tbOrder.Order_DPA_) AS a INNER JOIN
                      tbOrder ON a.Order_DPA_ = tbOrder.Order_DPA_ LEFT OUTER JOIN
                      Owner ON a.Owner_DPA_ = Owner.Owner_DPA_
WHERE     tbOrder.OrderCompounded = 1
ORDER BY a.Order_DPA_, a.ContractNumber

GO
CREATE VIEW dbo.DB_Affirmation2
AS
SELECT     TOP 100 PERCENT a.*, Owner.OwnerFName + ' ' + OwnerLName AS AccountManager
FROM         (SELECT     TOP 100 PERCENT dbo.Security.Security_DPA_, dbo.Lots.LotSlipNo, dbo.Lots.ContractNumber, dbo.Lots.LotQty, dbo.Lots.LotPrice,
                                              dbo.tbOrder.OrderRef, dbo.tbOrder.Order_DPA_, dbo.Client.ClientAddr AS AccountAddress, dbo.Client.ClientName AS Account,
                                              Cast(Floor(Cast(dbo.Lots.LotTDate AS Float)) AS DateTime) AS LotTDate, dbo.Client.Client_DPA_, dbo.Client.ClientCDSNo,
                                              dbo.Lots.LotGrossAmount AS SettlementGrossAmount, ISNULL(dbo.Client.ClientContact, '') AS Owner, dbo.Client.ClientFax,
                                              dbo.Security.SecurityName, BalanceQty = CASE ISNULL(dbo.OrdDetailContractedQtyList.ContractQty, 1)
                                              WHEN 1 THEN dbo.OrdDetail.OrdDetailQty ELSE dbo.OrdDetail.OrdDetailQty - dbo.OrdDetailContractedQtyList.ContractQty END,
                                              Security.SecurityName AS ordDetailSecurity, Class.ClassDescription, Contracts.ContractSettlementDate, Owner.Owner_DPA_,
                                              OrderType.OrderTypeSale,
                                              CASE dbo.OrderType.OrderTypeSale WHEN 0 THEN 'PURCHASE OF ' + dbo.Security.SecurityName ELSE 'SALE OF ' + dbo.Security.SecurityName
                                               END AS ReferenceHeader, CAST(CASE dbo.OrderType.OrderTypeSale WHEN 1 THEN cast(dbo.Lots.LotGrossAmount AS float)
                                              - Cast(dbo.ContractLevyTotal.LevyTotal AS float) + cast(dbo.AgentCommissionTotal.LevyTotal AS float)
                                              ELSE cast(dbo.Lots.LotGrossAmount AS float) + cast(dbo.ContractLevyTotal.LevyTotal AS float)
                                              - cast(dbo.AgentCommissionTotal.LevyTotal AS float) END AS float) AS SettlementAmount, Client.Class_DPA_,
                                              dbo.AgentCommissionTotal.LevyTotal, dbo.ContractLevyTotal.LevyTotal AS Comm, contracts.ContractSettlementDate AS SettlementDate,
                                              Security.SecurityCode, Contracts.Contract_DPA_
                       FROM          Lots INNER JOIN
                                              OrdDetail ON Lots.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                                              OrdDetailContractedQtyList ON OrdDetail.OrdDetail_DPA_ = OrdDetailContractedQtyList.OrdDetail_DPA_ INNER JOIN
                                              tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN
                                              Contracts ON Lots.Contract_DPA_ = Contracts.Contract_DPA_ INNER JOIN
                                              Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
                                              ContractLevyTotal ON Lots.Contract_DPA_ = ContractLevyTotal.Contract_DPA_ INNER JOIN
                                              AgentCommissionTotal ON Lots.Contract_DPA_ = AgentCommissionTotal.Contract_DPA_ INNER JOIN
                                              Security ON OrdDetail.Security_DPA_ = Security.Security_DPA_ INNER JOIN
                                              Class ON Client.Class_DPA_ = Class.Class_DPA_ INNER JOIN
                                              OrderType ON tbOrder.OrderType_DPA_ = OrderType.OrderType_DPA_ LEFT OUTER JOIN
                                              Owner ON Client.Owner_DPA_ = Owner.Owner_DPA_
                       WHERE      (tbOrder.Deleted = 0) AND (Client.Deleted = 0) AND (OrdDetail.Deleted = 0)
                       ORDER BY tbOrder.Order_DPA_) AS a INNER JOIN
                      tbOrder ON a.Order_DPA_ = tbOrder.Order_DPA_ LEFT OUTER JOIN
                      Owner ON a.Owner_DPA_ = Owner.Owner_DPA_
ORDER BY a.ContractSettlementDate DESC

GO

CREATE VIEW dbo.DB_AgentCommissionBonds
AS
SELECT     TOP 100 PERCENT dbo.Lots.ContractNumber, CAST(FLOOR(CAST(dbo.Lots.LotTDate AS Float)) AS DateTime) AS LotTDate, dbo.tbOrder.Client_DPA_,
                      dbo.OrderType.OrderTypeSale, dbo.LevyContract.LevyAmount, CAST(dbo.Lots.LotQty AS nvarchar)
                      + ' ' + dbo.Security.SecurityCode + ' @ ' + CAST(dbo.Lots.LotPrice AS nvarchar) AS SecurityName, dbo.Payment.PaymentReceiptNo,
                      dbo.LevyContract.Contract_DPA_, dbo.AgentList.AgentName, dbo.AgentList.Agent_DPA_, dbo.Client.ClientName, dbo.Lots.LotGrossAmount
FROM         dbo.Lots INNER JOIN
                      dbo.LevyContract ON dbo.Lots.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.Lots.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.AgentList ON dbo.Client.Agent_DPA_ = dbo.AgentList.Agent_DPA_ LEFT OUTER JOIN
                      dbo.Payment ON dbo.tbOrder.Order_DPA_ = dbo.Payment.Order_DPA_
WHERE     (dbo.tbOrder.OrderCanceled = 0) AND (dbo.LevyContract.SystemMaintained <> 8) AND (dbo.LevyContract.SystemMaintained <> 12) AND
                      (dbo.LevyContract.LevyShortName LIKE N'%Commission%') AND (dbo.Client.Deleted = 0) AND (dbo.OrdDetail.Deleted = 0) AND
                      (dbo.tbOrder.Deleted = 0) AND (dbo.Client.Owner_DPA_ IS NULL) AND (dbo.Security.OrderSecType_DPA_ = 1)
ORDER BY dbo.AgentList.Agent_DPA_, dbo.Lots.LotTDate


GO

CREATE VIEW dbo.DB_AgentCommissions
AS
SELECT     TOP 100 PERCENT dbo.Lots.ContractNumber, CAST(FLOOR(CAST(dbo.Lots.LotTDate AS Float)) AS DateTime) AS LotTDate, dbo.tbOrder.Client_DPA_,
                      dbo.OrderType.OrderTypeSale, dbo.LevyContract.LevyAmount, CAST(dbo.Lots.LotQty AS nvarchar)
                      + ' ' + dbo.Security.SecurityCode + ' @ ' + CAST(dbo.Lots.LotPrice AS nvarchar) AS SecurityName, dbo.Payment.PaymentReceiptNo,
                      dbo.LevyContract.Contract_DPA_, dbo.AgentList.AgentName, dbo.AgentList.Agent_DPA_, dbo.Client.ClientName, dbo.Lots.LotGrossAmount
FROM         dbo.Lots INNER JOIN
                      dbo.LevyContract ON dbo.Lots.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.Lots.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.AgentList ON dbo.Client.Agent_DPA_ = dbo.AgentList.Agent_DPA_ LEFT OUTER JOIN
                      dbo.Payment ON dbo.tbOrder.Order_DPA_ = dbo.Payment.Order_DPA_
WHERE     (dbo.tbOrder.OrderCanceled = 0) AND (dbo.LevyContract.SystemMaintained <> 8) AND (dbo.LevyContract.SystemMaintained <> 12) AND
                      (dbo.LevyContract.LevyShortName LIKE N'%Commission%') AND (dbo.Client.Deleted = 0) AND (dbo.OrdDetail.Deleted = 0) AND
                      (dbo.tbOrder.Deleted = 0) AND (dbo.Client.Owner_DPA_ IS NULL)
ORDER BY dbo.AgentList.Agent_DPA_, dbo.Lots.LotTDate


GO

CREATE VIEW dbo.DB_BalanceQtys
AS
SELECT dbo.OrdDetail.OrdDetail_DPA_, CASE (dbo.OrdDetail.OrdDetailQty)
               WHEN 0 THEN dbo.OrdDetail.Amount ELSE dbo.OrdDetail.OrdDetailQty END - SUM(ISNULL(dbo.Lot.LotQty, 0)) AS BalanceQty
FROM  dbo.OrdDetail LEFT OUTER JOIN
               dbo.Lot ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.Lot.OrdDetail_DPA_
GROUP BY dbo.OrdDetail.OrdDetail_DPA_, dbo.OrdDetail.OrdDetailQty, dbo.OrdDetail.Amount


GO

CREATE VIEW dbo.DB_BankAccountStatement
AS
SELECT     AccountsStatement.*, CreditBal - Debit AS Balance
FROM         (SELECT     Account_DPA_ AS BankAccount_DPA_, cast(floor(cast(AccountOpeningDate AS float)) AS DateTime) AS TransDate, '' AS Ref,
                                              UPPER('Opening Balance') AS Particulars, 0 AS Debit, 0 AS Credit, AccountOpeningBal AS CreditBal, 1 AS IsOpeningBalance, '' AS ReceiptNo
                       FROM          dbo.Account
UNION ALL
                       SELECT     dbo.Payments.BankAccount_DPA_, cast(floor(cast(dbo.Payments.PaymentPDate AS float)) AS datetime) AS TransDate,
                                             ISNULL(Lots.ContractNumber, '') + ' ' + Payments.PaymentReference AS Ref,UPPER(CASE PayType.PayTypeIn WHEN 1 THEN 'RECEIPT:' Else 'PAID OUT:' end +
                                             ISNULL(Security.SecurityCode + ' @ ' + CAST(Lots.LotPrice AS nvarchar), '') + ' ' + Payments.PaymentNarrative)  AS Particulars,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payments.PaymentAmount ELSE 0 END AS Debit,
                                             CASE PayType.PayTypeIn WHEN 0 THEN Payments.PaymentAmount ELSE 0 END AS Credit,
                                             CASE PayType.PayTypeIn WHEN 0 THEN Payments.PaymentAmount ELSE 0 END AS CreditBal, 0 AS OpeningBalance,
                                             PaymentReceiptNo AS ReceiptNo
                       FROM         OrdDetail INNER JOIN
                                             Lots ON OrdDetail.OrdDetail_DPA_ = Lots.OrdDetail_DPA_ INNER JOIN
                                             Security ON OrdDetail.Security_DPA_ = Security.Security_DPA_ RIGHT OUTER JOIN
                                             Payments INNER JOIN
                                             PayType ON Payments.PayType_DPA_ = PayType.PayType_DPA_ ON Lots.Contract_DPA_ = Payments.Contract_DPA_
UNION  ALL
                       SELECT     dbo.Payments.Entity_DPA_ AS Entity_DPA_, cast(floor(cast(dbo.Payments.PaymentPDate AS float)) AS DateTime) AS TransDate,
                                             dbo.Payments.PaymentReference AS Ref, UPPER(CASE PayType.PayTypeIn WHEN 1 then ' RECEIVED: '  else ' PAID OUT: ' end + isnull(dbo.Payments.PaymentNarrative,''))  AS Particulars,
                                             CASE PayType.PayTypeIn WHEN 0 THEN Payments.PaymentAmount ELSE 0 END AS Debit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payments.PaymentAmount ELSE 0 END AS Credit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payments.PaymentAmount ELSE 0 END AS CreditBal, 0 AS IsOpeningBalance,
                                             PaymentReceiptNo AS ReceiptNo
                       FROM         dbo.Payments INNER JOIN
                                             dbo.PayType ON dbo.Payments.PayType_DPA_ = dbo.PayType.PayType_DPA_
                       WHERE     (EntityType_DPA_ = 5)
UNION  ALL
                       SELECT     dbo.JournalList.Entity_DPA_, Cast(floor(Cast(dbo.JournalList.JournalDate AS float)) AS DateTime) AS TransDate,
                                             CAST(dbo.JournalList.JournalEntry_DPA_ AS SQL_VARIANT) AS Ref, UPPER(dbo.JournalList.JournalNarrative) AS Particulars,
                                             JournalList.JournalEntryDebit AS Debit, JournalList.JournalEntryCredit AS Credit, JournalList.JournalEntryCredit AS CreditBal,
                                             0 AS IsOpeningBalance, '' AS ReceiptNo
                       FROM         dbo.JournalList
                       WHERE     (EntityType_DPA_ = 5)
UNION  ALL
SELECT     Security.BankAccount_DPA_, CAST(FLOOR(CAST(Offerings.Offerings_Date AS float)) AS DateTime) AS TransDate, Offerings.PAL_No AS REF,
                      CAST(Offerings.Alloted_Rights AS char(6)) + ' ' + CASE WHEN len(Security.SecurityName) > 20 THEN LEFT(Security.SecurityName, 20)
                      + '...' ELSE Security.SecurityName END + ' @ ' + CAST(Offerings.Offering_Price AS char(5)) AS Particulars, 0 AS Debit,
                      Offerings.Alloted_Rights * Offerings.Offering_Price AS Credit, Offerings.Alloted_Rights * Offerings.Offering_Price AS CreditBal, 0 AS IsOpeningbal,
                      '' AS PaymentReceiptNo
FROM         Offerings INNER JOIN
                      Security ON Offerings.Offering = Security.Security_DPA_
WHERE     (Offerings.Deleted = 0) and (Offerings.Forward<>1)

                      ) AccountsStatement












GO

CREATE VIEW dbo.DB_BankTransactionList
AS
SELECT     AccountsStatement.*, CreditBal - Debit AS Balance
FROM         (SELECT     Account_DPA_ AS BankAccount_DPA_, AccountOpeningDate AS TransDate, '' AS Ref, 'Opening Balance' AS Particulars, 0 AS Debit,
                                              0 AS Credit, AccountOpeningBal AS CreditBal, 1 AS OpeningBalance, '' AS ReceiptNo
                       FROM          dbo.Account
                       WHERE      AccountTypeLevel1 = 7
                       UNION ALL
                       SELECT     dbo.Payment.BankAccount_DPA_, dbo.Payment.PaymentPDate AS TransDate, dbo.Payment.PaymentReference AS Ref,
                                             dbo.Payment.PaymentNarrative AS Particulars, CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                                             CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Credit,
                                             CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS CreditBal, 0 AS OpeningBalance,
                                             PaymentReceiptNo AS ReceiptNo
                       FROM         dbo.Payment INNER JOIN
                                             dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
                       UNION ALL
                       SELECT     dbo.Payment.Entity_DPA_ AS Entity_DPA_, dbo.Payment.PaymentPDate AS TransDate, dbo.Payment.PaymentReference AS Ref,
                                             dbo.Payment.PaymentNarrative AS Particulars, CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS CreditBal, 0 AS IsOpeningBalance,
                                             PaymentReceiptNo AS ReceiptNo
                       FROM         dbo.Payment INNER JOIN
                                             dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
                       WHERE     (EntityType_DPA_ = 5) AND (Entity_DPA_ IN
                                                 (SELECT     Account_DPA_
                                                   FROM          ACCOUNT
                                                   WHERE      (Account_DPA_ IN
                                                                              (SELECT     Entity_DPA_
                                                                                FROM          Payment
                                                                                WHERE      EntityType_DPA_ = 5))))
                       UNION ALL
                       SELECT     dbo.JournalList.Entity_DPA_, dbo.JournalList.JournalDate AS TransDate, CAST(dbo.JournalList.JournalEntry_DPA_ AS SQL_VARIANT) AS Ref,
                                              dbo.JournalList.JournalNarrative AS Particulars, JournalList.JournalEntryDebit AS Debit, JournalList.JournalEntryCredit AS Credit,
                                             JournalList.JournalEntryCredit AS CreditBal, 0 AS IsOpeningBalance, '' AS ReceiptNo
                       FROM         dbo.JournalList
                       WHERE     (EntityType_DPA_ = 5) ) AccountsStatement


GO
CREATE VIEW dbo.DB_DataStreamPriceList
AS
SELECT     dbo.EquityList.Security_DPA_ AS SecKnow_DPA, dbo.EquityList.SecurityCode, dbo.EquityList.Price AS MktClose
FROM         dbo.EquityList INNER JOIN
                          (SELECT     Security_DPA_, MAX([Date]) AS LastDate
                            FROM          (SELECT DISTINCT
                                                                           TOP 100 PERCENT Security.Security_DPA_, Security.SecurityCode, Security.SecurityCode AS SecCode,
                                                                           Security.SecurityName, datastream_Market.MktDate AS Date, datastream_Market.MktClose AS Price,
                                                                           dbo.datastream_Market.MktUnique
                                                    FROM          Security INNER JOIN
                                                                           datastream_Securities ON Security.Security_DPA_ = datastream_Securities.SecKnow_DPA INNER JOIN
                                                                           datastream_Market ON Security.SecurityCode = datastream_Market.MktCode
                                                    UNION
                                                    SELECT     TOP 100 PERCENT Security_DPA_, SecurityCode, SecurityName, SecurityCode AS SecCode, CAST(FLOOR(CAST(GETDATE()
                                                                          AS Float)) AS DateTime) AS [Date], NULL AS Price, 0 AS MktUnique
                                                    FROM         dbo.Security
                                                    WHERE     security_DPA_ NOT IN
                                                                              (SELECT DISTINCT Security.Security_DPA_
                                                                                FROM          Security INNER JOIN
                                                                                                       datastream_Securities ON Security.Security_DPA_ = datastream_Securities.SecKnow_DPA INNER JOIN
                                                                                                       datastream_Market ON Security.SecurityCode = datastream_Market.MktCode
                                                                                WHERE      (FLOOR(CAST(datastream_Market.MktDate AS Float)) = FLOOR(CAST(GETDATE() AS Float))))) EquityList
                            WHERE      (NOT (Price IS NULL))
                            GROUP BY Security_DPA_) LatestTransactions ON dbo.EquityList.Security_DPA_ = LatestTransactions.Security_DPA_ AND
                      dbo.EquityList.[Date] = LatestTransactions.LastDate

GO

CREATE VIEW dbo.DB_DataStreamPriceList
AS
SELECT DISTINCT TOP 100 PERCENT Dsm.MktClose, dbo.datastream_Securities.SecKnow_DPA
FROM         dbo.datastream_Market Dsm INNER JOIN
                      dbo.datastream_Securities ON Dsm.MktCode = dbo.datastream_Securities.SecCode
WHERE     (FLOOR(CAST(Dsm.MktDate AS float)) =
                          (SELECT     MAX(floor(cast(MktDate AS float))) AS MktUnique
                            FROM          dbo.datastream_Market
                            WHERE      (NOT (MktClose IS NULL)) AND MktCode = Dsm.[MktCode]))


GO

CREATE VIEW dbo.DB_FineTradingSchedule
AS
SELECT DISTINCT
                      TOP 100 PERCENT dbo.OrderType.OrderTypeDescription AS OrdDetailType, dbo.OrdDetail.OrdDetailPrice, dbo.OrdDetail.Order_DPA_,
                      dbo.tbOrder.OrderRef, dbo.tbOrder.OrderDate, dbo.Security.SecurityCode, dbo.Client.Client_DPA_ AS Code, dbo.Client.ClientName AS Client,
                      dbo.OrdDetail.Best, dbo.OrdDetail.OrdDetailValidity AS Validity, dbo.OrdDetail.Amount, dbo.DB_DataStreamPriceList.MktClose AS Price,
                      dbo.DB_OrdDetailContractedQtyList.BalanceQty, CASE (OrderTypeSale) WHEN 0 THEN (- isnull(dbo.Client.CreditLimit, 0)
                      - (isnull(dbo.ClientBalances.CurrentBal, 0) - isnull(dbo.ClientTotal.Total, 0))) ELSE (isnull(SalesOrdersTotals.Total, 0) - isnull(0, 0)) END AS Excess,
                      dbo.ClientBalances.CurrentBal, isnull(dbo.ClientTotal.Total, 0) AS Total, dbo.OrderSecType.OrderSecTypeDisplayName AS OrdDetailSecType,
                      dbo.Security.SecurityName AS ordDetailSecurity, dbo.Client.CreditLimit, dbo.OrderType.OrderTypeSale, dbo.OrdDetail.Security_DPA_,
                      dbo.Client.ClientCDSNo, CASE WHEN len(dbo.Agent.AgentName) > 10 THEN LEFT(dbo.Agent.AgentName, 10)
                      + '...' ELSE dbo.Agent.AgentName END AS AgentName, dbo.Agent.Agent_DPA_ AS AgentCode, '' AS BalanceFree, dbo.Bond.BondIssue,
                      dbo.Client.Iscustodian, dbo.OrdDetail.Limit
FROM         OrdDetail LEFT OUTER JOIN
                      Bond ON OrdDetail.Bond_DPA_ = Bond.Bond_DPA_ INNER JOIN
                      Security ON OrdDetail.Security_DPA_ = Security.Security_DPA_ INNER JOIN
                      tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN
                      OrderType ON tbOrder.OrderType_DPA_ = OrderType.OrderType_DPA_ INNER JOIN
                      OrderSecType ON tbOrder.OrderSecType_DPA_ = OrderSecType.OrderSecType_DPA_ INNER JOIN
                      Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ LEFT OUTER JOIN
                      TotalSalesOrderholdings ON Client.Client_DPA_ = TotalSalesOrderholdings.Client_DPA_ AND
                      Security.Security_DPA_ = TotalSalesOrderholdings.Security_DPA_ INNER JOIN
                      SalesOrdersTotals ON OrdDetail.Order_DPA_ = SalesOrdersTotals.Order_DPA_ AND
                      Security.Security_DPA_ = SalesOrdersTotals.Security_DPA_ INNER JOIN
                      DB_OrdDetailContractedQtyList ON OrdDetail.OrdDetail_DPA_ = DB_OrdDetailContractedQtyList.OrdDetail_DPA_ LEFT OUTER JOIN
                      Agent ON Client.Agent_DPA_ = Agent.Agent_DPA_ LEFT OUTER JOIN
                      ClientBalances ON Client.Client_DPA_ = ClientBalances.client_DPA_ LEFT OUTER JOIN
                      ClientTotal ON Client.Client_DPA_ = ClientTotal.Client_DPA_ LEFT OUTER JOIN
                      DB_DataStreamPriceList ON Security.Security_DPA_ = DB_DataStreamPriceList.SecKnow_DPA
WHERE     (DB_OrdDetailContractedQtyList.BalanceQty > 0) AND (tbOrder.OrderHold = 0) AND (tbOrder.OrderCanceled = 0) AND (tbOrder.Deleted = 0) AND
                      (OrdDetail.Deleted = 0)
GO

CREATE VIEW dbo.DB_InterBankTransferList
AS
SELECT DISTINCT
                      TOP 100 PERCENT dbo.InterBankTransferList.*, dbo.OrdDetail.Amount AS Amount, dbo.OrdDetail.Best AS Best,
                      dbo.datastream_SecurityPriceList.Price AS Price
FROM         dbo.InterBankTransferList INNER JOIN
                      dbo.OrdDetail ON dbo.InterBankTransferList.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.Security ON dbo.InterBankTransferList.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.datastream_SecurityPriceList ON dbo.Security.Security_DPA_ = dbo.datastream_SecurityPriceList.Security_DPA_
ORDER BY dbo.InterBankTransferList.Order_DPA_ DESC, dbo.InterBankTransferList.ContractNumber DESC



GO

CREATE VIEW dbo.DB_OrdDetailContractedQtyList
AS
SELECT     dbo.OrdDetail.OrdDetail_DPA_, dbo.OrdDetail.OrdDetailQty - SUM(ISNULL(dbo.Lots.LotQty, 0)) AS BalanceQty
FROM         dbo.OrdDetail LEFT OUTER JOIN
                      dbo.Lots ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.Lots.OrdDetail_DPA_
GROUP BY dbo.OrdDetail.OrdDetail_DPA_, dbo.OrdDetail.OrdDetailQty


GO

CREATE VIEW dbo.DB_Portfolios
AS
SELECT DISTINCT
                      TOP 100 PERCENT dbo.tbOrder.Order_DPA_, dbo.OrdDetail.Security_DPA_, dbo.Security.SecurityCode, dbo.tbOrder.Client_DPA_,
                      dbo.OrderType.OrderTypeSale, dbo.Lot.ContractNumber, dbo.Lot.LotSlipNo, dbo.OrdDetail.OrdDetailPrice, dbo.OrdDetail.OrdDetailQty,
                      dbo.Client.ClientName AS OrdDetailClient, dbo.OrdDetail.OrdDetail_DPA_, dbo.OrdDetail.Amount, isnull(dbo.DB_DataStreamPriceList.MktClose, 0)
                      AS SecurityMktPrice, dbo.Lot.Lot_DPA_, dbo.Lot.LotQty, dbo.Lot.LotPrice, cast(floor(cast(dbo.Lot.LotTDate AS float)) AS Datetime) AS TransDate,
                      dbo.Lot.LotGrossAmount, dbo.ClientBalances.CurrentBal,
                      CAST(CASE dbo.OrderType.OrderTypeSale WHEN 1 THEN dbo.Lot.LotGrossAmount - dbo.ContractLevyTotal.LevyTotal ELSE dbo.Lot.LotGrossAmount + dbo.ContractLevyTotal.LevyTotal
                       END AS real) AS NetAmount, Security.SecurityName, dbo.OrderType.OrderTypeDescription
FROM         dbo.OrderSecType INNER JOIN
                      dbo.tbOrder ON dbo.OrderSecType.OrderSecType_DPA_ = dbo.tbOrder.OrderSecType_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                      dbo.Lot INNER JOIN
                      dbo.OrdDetail ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.ContractLevyTotal ON dbo.Lot.Contract_DPA_ = dbo.ContractLevyTotal.Contract_DPA_ INNER JOIN
                      dbo.ClientBalances ON dbo.Client.Client_DPA_ = dbo.ClientBalances.client_DPA_ LEFT OUTER JOIN
                      dbo.DB_DataStreamPriceList ON dbo.Security.Security_DPA_ = dbo.DB_DataStreamPriceList.SecKnow_DPA
WHERE     (dbo.Security.ExpiryDate IS NULL OR cast(floor(cast(dbo.Security.ExpiryDate AS float)) AS Datetime) <= cast(floor(cast(GetDATE() AS float)) AS Datetime))
UNION
SELECT DISTINCT
                      TOP 100 PERCENT 1 AS order_Dpa_, dbo.CPortfolio.Security_DPA_, dbo.Security.SecurityCode, dbo.CPortfolio.Client_DPA_, 0 AS orderTypeSale,
                      isnull(dbo.CPortfolio.Reference, '') AS Cont, '' AS LotSlipNo, CONVERT(char(20), dbo.CPortfolio.CPortfolioPrice) AS Price, dbo.CPortfolio.CPortfolioQty,
                      dbo.Client.ClientName AS CPortfolioClient, 0 AS OrderDetail_DPA_, 0 AS amount, isnull(dbo.DB_DataStreamPriceList.MktClose, 0) AS SecurityMktPrice,
                      0 AS Lot_DPA_, dbo.CPortfolio.CPortfolioQty AS LotQty, dbo.CPortfolio.CPortfolioPrice AS LotPrice,
                      cast(floor(cast(dbo.CPortfolio.CPortfolioPDate AS float)) AS DateTime) AS TransDate,
                      dbo.CPortfolio.CPortfolioPrice * dbo.CPortfolio.CPortfolioQty AS LotGrossAmount, dbo.ClientBalances.CurrentBal,
                      dbo.CPortfolio.CPortfolioPrice * dbo.CPortfolio.CPortfolioQty AS NetAmount, dbo.Security.SecurityName AS SecurityName, 'Purchase' AS DESCr
FROM         dbo.Security INNER JOIN
                      dbo.Client INNER JOIN
                      dbo.CPortfolio ON dbo.Client.Client_DPA_ = dbo.CPortfolio.Client_DPA_ ON dbo.Security.Security_DPA_ = dbo.CPortfolio.Security_DPA_ INNER JOIN
                      dbo.ClientBalances ON dbo.Client.Client_DPA_ = dbo.ClientBalances.client_DPA_ LEFT OUTER JOIN
                      dbo.DB_DataStreamPriceList ON dbo.Security.Security_DPA_ = dbo.DB_DataStreamPriceList.SecKnow_DPA
WHERE     (dbo.Client.Deleted = 0)
AND    (dbo.Security.ExpiryDate IS NULL OR cast(floor(cast(dbo.Security.ExpiryDate AS float)) AS Datetime) <= cast(floor(cast(GetDATE() AS float)) AS Datetime))
ORDER BY TransDate








GO

CREATE VIEW dbo.DB_Reconciliation
AS
SELECT     AccountsStatement.*, CreditBal - Debit AS Balance
FROM         (SELECT     Account_DPA_ AS BankAccount_DPA_, AccountOpeningDate AS ReconcileDate, AccountOpeningDate AS TransDate, '' AS Ref,
                                              'Opening Balance' AS Particulars, 0 AS Debit, 0 AS Credit, AccountOpeningBal AS CreditBal, 1 AS IsOpeningBalance, '' AS ReceiptNo
                       FROM          dbo.Account
                       WHERE      AccountTypeLevel1 = 7
                       UNION ALL
                       SELECT     dbo.Payment.BankAccount_DPA_, dbo.Payment.ReconcileDate, dbo.Payment.PaymentPDate AS TransDate,
                                             ISNULL(Lots.ContractNumber, '')
                                             + ' ' + Payment.PaymentReference AS Ref, ISNULL(Security.SecurityCode + ' @ ' + CAST(Lots.LotPrice AS nvarchar), '')
                                             + ' ' + Payment.PaymentNarrative AS Particulars,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                                             CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Credit,
                                             CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS CreditBal, 0 AS OpeningBalance,
                                             PaymentReceiptNo AS ReceiptNo
                        FROM         OrdDetail INNER JOIN
                                             Lots ON OrdDetail.OrdDetail_DPA_ = Lots.OrdDetail_DPA_ INNER JOIN
                                             Security ON OrdDetail.Security_DPA_ = Security.Security_DPA_ RIGHT OUTER JOIN
                                             Payment INNER JOIN
                                             PayType ON Payment.PayType_DPA_ = PayType.PayType_DPA_ ON Lots.Contract_DPA_ = Payment.Contract_DPA_
                       UNION ALL
                       SELECT     TOP 100 PERCENT 4 AS Bankcode, CAST(FLOOR(CAST(InterTransfer.ReconcileDate AS float)) AS DateTime) AS ReconcileDate,
                                             CAST(FLOOR(CAST(InterTransfer.TransferDate AS float)) AS DateTime) AS TransDate, Lot.contractNumber AS Ref,
                                             isnull(InterTransfer.TransferNarrative, '') + CASE (OrderType.OrderTypeSale)
                                             WHEN 1 THEN 'SALE ' ELSE ' PURCHASE ' END + InterTransferType.TypeDescription AS Particulars, CASE (OrderType.OrderTypeSale)
                                             WHEN 0 THEN intertransfer.Transferamount ELSE 0 END AS Debit, CASE (OrderType.OrderTypeSale)
                                             WHEN 1 THEN intertransfer.Transferamount ELSE 0 END AS Credit, CASE (OrderType.OrderTypeSale)
                                             WHEN 1 THEN intertransfer.Transferamount ELSE 0 END AS CreditBal, 0 AS IsOpeningBal, isnull(intertransfer.transferreference, '')
                                             AS Receiptno
                       FROM         OrdDetail INNER JOIN
                                             Lot ON OrdDetail.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN
                                             tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN
                                             OrderType ON tbOrder.OrderType_DPA_ = OrderType.OrderType_DPA_ INNER JOIN
                                             InterTransfer INNER JOIN
                                             InterTransferType ON InterTransfer.InterTransferType_DPA_ = InterTransferType.InterTransferType_DPA_ ON
                                             Lot.Contract_DPA_ = InterTransfer.Contract_DPA_
                       UNION ALL
                       SELECT     dbo.Payment.Entity_DPA_ AS Entity_DPA_, dbo.Payment.ReconcileDate AS ReconcileDate, dbo.Payment.PaymentPDate AS TransDate,
                                             dbo.Payment.PaymentReference AS Ref, dbo.Payment.PaymentNarrative AS Particulars,
                                             CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS CreditBal, 0 AS IsOpeningBalance,
                                             PaymentReceiptNo AS ReceiptNo
                       FROM         dbo.Payment INNER JOIN
                                             dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
                       WHERE     (EntityType_DPA_ = 5) AND (Entity_DPA_ IN
                                                 (SELECT     Account_DPA_
                                                   FROM          ACCOUNT
                                                   WHERE      (Account_DPA_ IN
                                                                              (SELECT     Entity_DPA_
                                                                                FROM          Payment
                                                                                WHERE      EntityType_DPA_ = 5))))
                       UNION ALL
                       SELECT     dbo.JournalList.Entity_DPA_, dbo.JournalList.ReconcileDate AS Reconciledate, dbo.JournalList.JournalDate AS TransDate,
                                             CAST(dbo.JournalList.JournalEntry_DPA_ AS SQL_VARIANT) AS Ref, dbo.JournalList.JournalNarrative AS Particulars,
                                             JournalList.JournalEntryDebit AS Debit, JournalList.JournalEntryCredit AS Credit, JournalList.JournalEntryCredit AS CreditBal,
                                             0 AS IsOpeningBalance, '' AS ReceiptNo
                       FROM         dbo.JournalList
                       WHERE     (EntityType_DPA_ = 5) AND (Entity_DPA_ IN
                                                 (SELECT     Account_DPA_
                                                   FROM          ACCOUNT
                                                   WHERE      (Account_DPA_ IN
                                                                              (SELECT     Entity_DPA_
                                                                                FROM          Payment
                                                                                WHERE      EntityType_DPA_ = 5))))) AccountsStatement





GO

CREATE VIEW dbo.Db_TradeAffirmationBond
AS
SELECT DISTINCT
                      a.*, dbo.TradeAffirmationBond.LotPrice, dbo.TradeAffirmationBond.Owner, dbo.TradeAffirmationBond.AccountManager,
                      dbo.TradeAffirmationBond.Account, dbo.TradeAffirmationBond.MaturityDate, dbo.TradeAffirmationBond.BrokerName,
                      dbo.TradeAffirmationBond.IssueDate, dbo.TradeAffirmationBond.OrderRef, dbo.TradeAffirmationBond.TradeAction,
                      dbo.TradeAffirmationBond.BondIssue, dbo.TradeAffirmationBond.OrderTypeSale, dbo.TradeAffirmationBond.Client_DPA_
FROM         (SELECT     Order_DPA_, SUM(Commission) AS TotalCommission, SUM(dbo.TradeAffirmationBond.LotQty) AS Quantity
                       FROM          TradeAffirmationBond
                       WHERE      (LotTDate = '24-Mar-2005')
                       GROUP BY Order_DPA_) a INNER JOIN
                      dbo.TradeAffirmationBond ON a.Order_DPA_ = dbo.TradeAffirmationBond.Order_DPA_


GO

CREATE VIEW dbo.Db_TrialBalance
AS
SELECT     dbo.StatementList1.Credit, dbo.StatementList1.Debit, dbo.Client.ClientOpeningBal AS OpeningBal, dbo.StatementList1.Client_DPA_,
                      dbo.Client.ClientName, dbo.StatementList1.TransDate, 'Clients' AS Entity
FROM         dbo.StatementList1 INNER JOIN
                      dbo.Client ON dbo.StatementList1.Client_DPA_ = dbo.Client.Client_DPA_
UNION
SELECT     dbo.AgentStatement1.Credit, dbo.AgentStatement1.Debit, dbo.Agent.AgentOpeningBal AS OpeningBal, dbo.AgentStatement1.Agent_DPA_,
                      dbo.Agent.AgentName, dbo.AgentStatement1.TransDate, 'Agents' AS Entity
FROM         dbo.AgentStatement1 INNER JOIN
                      dbo.Agent ON dbo.AgentStatement1.Agent_DPA_ = dbo.Agent.Agent_DPA_
UNION
SELECT     dbo.BrokerStatement1.Credit, dbo.BrokerStatement1.Debit, dbo.Broker.BrokerOpeningBal AS OpeningBal, dbo.BrokerStatement1.Broker_DPA_,
                      dbo.Broker.BrokerName, dbo.BrokerStatement1.TransDate, 'Brokers' AS Entity
FROM         dbo.BrokerStatement1 INNER JOIN
                      dbo.Broker ON dbo.BrokerStatement1.Broker_DPA_ = dbo.Broker.Broker_DPA_
UNION
SELECT     dbo.OwnerStatement.Credit, dbo.OwnerStatement.Debit, dbo.Owner.OwnerOpeningBal AS OpeningBal, dbo.OwnerStatement.Owner_DPA_,
                      dbo.Owner.OwnerFName + ' ' + dbo.Owner.OwnerLName, dbo.OwnerStatement.TransDate, 'Account Managers' AS Entity
FROM         dbo.OwnerStatement INNER JOIN
                      dbo.Owner ON dbo.OwnerStatement.Owner_DPA_ = dbo.Owner.Owner_DPA_


GO
CREATE VIEW dbo.Debita
AS
SELECT DISTINCT
                      TOP 100 PERCENT a.Client_DPA_, MAX(a.TransDate) AS LastDate, b.CurrentBal AS Balance, dbo.Client.CreditLimit, dbo.Client.ClientCellTel
FROM         (SELECT     *
                       FROM          dbo.ClientTransactionList
                       WHERE      Transdate <= getdate()) a INNER JOIN
                      dbo.Client ON a.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                          (SELECT     SUM(ISNULL(dbo.StatementList.Credit - dbo.StatementList.Debit, 0)) + dbo.Client.ClientOpeningBal AS CurrentBal,
                                                   dbo.StatementList.Client_DPA_
                            FROM          dbo.StatementList INNER JOIN
                                                   dbo.Client ON dbo.StatementList.Client_DPA_ = dbo.Client.Client_DPA_
                            WHERE      (dbo.Client.Deleted = 0 AND Transdate <= getdate())
                            GROUP BY dbo.StatementList.Client_DPA_, dbo.Client.ClientOpeningBal) b ON dbo.Client.Client_DPA_ = b.Client_DPA_
WHERE     (dbo.Client.Deleted = 0)
GROUP BY a.Client_DPA_, b.CurrentBal, dbo.Client.CreditLimit, dbo.Client.ClientCellTel, dbo.Client.updateOnDebt
HAVING      (b.CurrentBal < 0) AND (dbo.Client.updateOnDebt = 1) AND (dbo.Client.ClientCellTel IS NOT NULL) AND (RTRIM(LTRIM(dbo.Client.ClientCellTel)) <> '')

GO
CREATE VIEW dbo.Debitb
AS
SELECT DISTINCT
                      TOP 100 PERCENT a.Client_DPA_, MAX(a.TransDate) AS LastDate, b.CurrentBal AS Balance, dbo.Client.CreditLimit, dbo.Client.ClientCellTel
FROM         (SELECT     *
                       FROM          dbo.ClientTransactionList
                       WHERE      Transdate <= dateadd(d,-5,getdate())) a INNER JOIN
                      dbo.Client ON a.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                          (SELECT     SUM(ISNULL(dbo.StatementList.Credit - dbo.StatementList.Debit, 0)) + dbo.Client.ClientOpeningBal AS CurrentBal,
                                                   dbo.StatementList.Client_DPA_
                            FROM          dbo.StatementList INNER JOIN
                                                   dbo.Client ON dbo.StatementList.Client_DPA_ = dbo.Client.Client_DPA_
                            WHERE      (dbo.Client.Deleted = 0 AND Transdate <= dateadd(d,-5,getdate()) )
                            GROUP BY dbo.StatementList.Client_DPA_, dbo.Client.ClientOpeningBal) b ON dbo.Client.Client_DPA_ = b.Client_DPA_
WHERE     (dbo.Client.Deleted = 0)
GROUP BY a.Client_DPA_, b.CurrentBal, dbo.Client.CreditLimit, dbo.Client.ClientCellTel, dbo.Client.updateOnDebt
HAVING      (b.CurrentBal < 0) AND (dbo.Client.updateOnDebt = 1) AND (dbo.Client.ClientCellTel IS NOT NULL) AND (RTRIM(LTRIM(dbo.Client.ClientCellTel)) <> '')


GO

CREATE VIEW dbo.DebtorCreditor
AS
SELECT DISTINCT TOP 100 PERCENT a.Client_DPA_, MAX(a.TransDate) AS LastDate, dbo.CurrentBalances.CurrentBal AS Balance, dbo.Client.ClientName
FROM         (SELECT     *
                       FROM          dbo.ClientTransactionList) a INNER JOIN
                      dbo.Client ON a.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.CurrentBalances ON dbo.Client.Client_DPA_ = dbo.CurrentBalances.Client_DPA_
WHERE     (dbo.Client.Deleted = 0)
GROUP BY a.Client_DPA_, dbo.CurrentBalances.CurrentBal, dbo.Client.ClientName
ORDER BY dbo.Client.ClientName


GO

CREATE VIEW dbo.Debtors
AS
SELECT Client_DPA_, LastDate, Balance
FROM  dbo.DebtorCreditor
WHERE (Balance < 0)


GO
CREATE VIEW dbo.DeliveredContractList
AS
SELECT     dbo.LotView.LotSlipNo, dbo.LotView.LotTDate, dbo.Security.SecurityCode, SUM(dbo.LotView.LotQty) AS LotQty, SUM(dbo.LotView.LotPrice) AS LotPrice,
                      dbo.Contract.ContractTransferNo, dbo.OrderType.OrderTypeSale, dbo.OrderSecType.OrderSecTypeDisplayName,
                      dbo.OrderSecType.OrderSecType_DPA_, dbo.OrdDetail.OrdDetailCertNo AS ContractNCertificate, dbo.Contract.ContractDeliveryDate
FROM         dbo.OrdDetail INNER JOIN
                      dbo.LotView ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.LotView.OrdDetail_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.Contract ON dbo.LotView.Contract_DPA_ = dbo.Contract.Contract_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                      dbo.OrderSecType ON dbo.Security.OrderSecType_DPA_ = dbo.OrderSecType.OrderSecType_DPA_
WHERE     (dbo.Contract.ContractDelivered = 1)
GROUP BY dbo.LotView.LotSlipNo, dbo.LotView.LotTDate, dbo.Security.SecurityCode, dbo.LotView.LotQty, dbo.LotView.LotPrice, dbo.Contract.ContractTransferNo,
                      dbo.Contract.ContractNCertificate, dbo.OrderType.OrderTypeSale, dbo.OrderSecType.OrderSecTypeDisplayName,
                      dbo.OrderSecType.OrderSecType_DPA_, dbo.OrdDetail.OrdDetailCertNo, dbo.Contract.ContractDeliveryDate

GO
CREATE VIEW dbo.DeliverySlips
AS
SELECT     dbo.LotView.LotSlipNo, dbo.LotView.LotTDate, dbo.Security.SecurityCode, SUM(dbo.LotView.LotQty) AS LotQty, dbo.LotView.LotPrice AS LotPrice,
                      dbo.Contract.ContractTransferNo, dbo.OrderType.OrderTypeSale, dbo.OrderSecType.OrderSecTypeDisplayName,
                      dbo.OrderSecType.OrderSecType_DPA_, dbo.OrdDetail.OrdDetailCertNo AS ContractNCertificate
FROM         dbo.OrdDetail INNER JOIN
                      dbo.LotView ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.LotView.OrdDetail_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.Contract ON dbo.LotView.Contract_DPA_ = dbo.Contract.Contract_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                      dbo.OrderSecType ON dbo.Security.OrderSecType_DPA_ = dbo.OrderSecType.OrderSecType_DPA_
GROUP BY dbo.LotView.LotSlipNo, dbo.LotView.LotTDate, dbo.Security.SecurityCode, dbo.LotView.LotQty, dbo.LotView.LotPrice, dbo.Contract.ContractTransferNo,
                      dbo.Contract.ContractNCertificate, dbo.OrderType.OrderTypeSale, dbo.OrderSecType.OrderSecTypeDisplayName,
                      dbo.OrderSecType.OrderSecType_DPA_, dbo.OrdDetail.OrdDetailCertNo

GO

CREATE VIEW dbo.EditBondProposals
AS
SELECT     dbo.BondProposals.*, dbo.BondClientList.ClientName AS ClientName, dbo.OwnerList.OwnerName AS OwnerName,
                      dbo.Security.SecurityCode AS SecurityCode, dbo.Commission.CommissionDescription AS CommDesc,
                      dbo.Commission.Commission_DPA_ AS Comm_DPA, dbo.Bond.BondIssue AS Bond
FROM         dbo.BondProposals INNER JOIN
                      dbo.BondClientList ON dbo.BondProposals.Client_DPA_ = dbo.BondClientList.Client_DPA_ LEFT OUTER JOIN
                      dbo.OwnerList ON dbo.BondClientList.Owner_DPA_ = dbo.OwnerList.Owner_DPA_ INNER JOIN
                      dbo.Security ON dbo.BondProposals.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.Commission ON dbo.BondClientList.Commission_DPA_ = dbo.Commission.Commission_DPA_ INNER JOIN
                      dbo.Bond ON dbo.BondProposals.Bond_DPA_ = dbo.Bond.Bond_DPA_


GO
CREATE VIEW dbo.EditOffering
AS
SELECT     dbo.Payment.BankAccount_DPA_, dbo.Payment.PaymentAmount, dbo.Payment.PaymentPDate, dbo.OfferingsList.PAL_No,
                      dbo.OfferingsList.Client_DPA_, dbo.OfferingsList.Offering, dbo.OfferingsList.Offering_DPA_, dbo.OfferingsList.Offering_Price, dbo.OfferingsList.ID_No,
                      dbo.OfferingsList.Alloted_Rights, dbo.OfferingsList.Accepted_Rights, dbo.OfferingsList.Receipt, dbo.OfferingsList.Renouncee,
                      dbo.OfferingsList.Submitted, dbo.OfferingsList.Submission_Date, dbo.OfferingsList.Ref_No, dbo.OfferingsList.ClientName,
                      dbo.OfferingsList.Amount_Payable, dbo.OfferingsList.SecurityName, dbo.OfferingsList.SecurityCode, dbo.OfferingsList.Batch_No,
                      dbo.OfferingsList.TimeChanged, dbo.OfferingsList.ModifiedBy, dbo.Payment.PaymentReference, dbo.Payment.PaymentNarrative,
                      dbo.Payment.PaymentReceiptNo, dbo.Payment.Payment_DPA_
FROM         dbo.OfferingsList INNER JOIN
                      dbo.Payment ON dbo.OfferingsList.Receipt = dbo.Payment.Payment_DPA_

GO

CREATE VIEW dbo.EntityList
AS
SELECT     dbo.Entity.EntityName, dbo.EntityType.EntityTypeName AS EntityType, dbo.Entity.Entity_DPA_, dbo.EntityType.EntityType_DPA_,
                      dbo.AccountType.AccountTypeName, dbo.AccountType.AccountType_DPA_, ISNULL(dbo.Entity.EntityCode,
                      CAST(dbo.Entity.Entity_DPA_ AS nvarchar(4000))) AS EntityCode, ISNULL(dbo.Entity.EntityCode, CAST(dbo.Entity.Entity_DPA_ AS nvarchar(4000)))
                      + SPACE(ISNULL(MAX(LEN(ISNULL(dbo.Entity.EntityCode, CAST(dbo.Entity.Entity_DPA_ AS nvarchar(4000))))), 0) - LEN(ISNULL(dbo.Entity.EntityCode,
                      CAST(dbo.Entity.Entity_DPA_ AS nvarchar(4000))))) + ' : ' + dbo.Entity.EntityName AS EntityNameEx, dbo.Entity.EntityOpeningBal,
                      CAST(FLOOR(CAST(dbo.Entity.EntityRegDate AS Float)) AS DateTime) AS RegistrationDate
FROM         dbo.EntityType INNER JOIN
                      dbo.Entity ON dbo.EntityType.EntityType_DPA_ = dbo.Entity.EntityType_DPA_ INNER JOIN
                      dbo.AccountType ON dbo.EntityType.AccountType_DPA_ = dbo.AccountType.AccountType_DPA_
GROUP BY dbo.Entity.EntityName, dbo.EntityType.EntityTypeName, dbo.Entity.Entity_DPA_, dbo.EntityType.EntityType_DPA_,
                      dbo.AccountType.AccountTypeName, dbo.AccountType.AccountType_DPA_, dbo.Entity.EntityCode, dbo.Entity.EntityOpeningBal,
                      CAST(FLOOR(CAST(dbo.Entity.EntityRegDate AS Float)) AS DateTime)



GO

CREATE VIEW dbo.EntityTransactionList
AS



(SELECT     Entity_DPA_, EntityRegDate AS TransDate, '' AS REF, ' Opening Balance' AS Particulars, case when EntityOpeningBal < 0 then EntityOpeningBal else 0 end AS Debit, case when EntityOpeningBal >= 0 then EntityOpeningBal else 0 end AS Credit,
                                              1 AS IsOpeningBalance,case when EntityOpeningBal < 0 then 0 - EntityOpeningBal else EntityOpeningBal end AS Balance
                       FROM          dbo.Entity WHERE SystemMaintained = 0)


union all

(SELECT     dbo.Payment.Entity_DPA_ AS Entity_DPA_, dbo.Payment.PaymentPDate AS TransDate, dbo.Payment.PaymentReference AS Ref,
                     dbo.Payment.PaymentNarrative AS Particulars, CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                     CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit, 0 AS IsOpeningBalance, CASE PayType.PayTypeIn WHEN 0 THEN 0 - Payment.PaymentAmount ELSE Payment.PaymentAmount END AS Balance
FROM         dbo.Payment INNER JOIN
                     dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_ INNER JOIN
                      dbo.Entity ON (dbo.Payment.Entity_DPA_ = dbo.Entity.Entity_DPA_) AND (dbo.Payment.EntityType_DPA_ = dbo.Entity.EntityType_DPA_)
WHERE     (dbo.Entity.SystemMaintained = 0))

UNION ALL

(SELECT     dbo.JournalList.Entity_DPA_, dbo.JournalList.JournalDate AS TransDate, CAST(dbo.JournalList.JournalEntry_DPA_ AS NVARCHAR(500)) AS Ref,
                     dbo.JournalList.JournalNarrative AS Particulars, JournalList.JournalEntryDebit AS Debit,
                     JournalList.JournalEntryCredit AS Credit, 0 AS IsOpeningBalance,JournalList.JournalEntryCredit - JournalList.JournalEntryDebit AS Balance
FROM         dbo.JournalList INNER JOIN
                      dbo.Entity ON (dbo.JournalList.Entity_DPA_ = dbo.Entity.Entity_DPA_) AND (dbo.JournalList.EntityType_DPA_ = dbo.Entity.EntityType_DPA_)
WHERE     (dbo.Entity.SystemMaintained = 0))



GO

CREATE VIEW dbo.EntityTypeList
AS
SELECT     dbo.AccountType.AccountTypeName AS EntityTypeAccountType, dbo.EntityType.EntityTypeName, dbo.EntityType.EntityType_DPA_,
                      dbo.EntityType.DefaultSelection, dbo.EntityType.EntityTypeCode
FROM         dbo.AccountType INNER JOIN
                      dbo.EntityType ON dbo.AccountType.AccountType_DPA_ = dbo.EntityType.AccountType_DPA_
WHERE     (dbo.EntityType.SystemMaintained = 0)


GO
CREATE VIEW dbo.EquityList
AS

SELECT DISTINCT
                      TOP 100 PERCENT Security.Security_DPA_, Security.SecurityCode, Security.SecurityCode AS SecCode, Security.SecurityName,
                      datastream_Market.MktDate AS Date, datastream_Market.MktClose AS Price,dbo.datastream_Market.MktUnique
FROM         Security INNER JOIN
                      datastream_Securities ON Security.Security_DPA_ = datastream_Securities.SecKnow_DPA INNER JOIN
                      datastream_Market ON Security.SecurityCode = datastream_Market.MktCode
UNION
SELECT     TOP 100 PERCENT Security_DPA_, SecurityCode, SecurityName, SecurityCode AS SecCode, CAST(FLOOR(CAST(GETDATE() AS Float)) AS DateTime)
                      AS [Date], NULL AS Price ,0 AS MktUnique
FROM         dbo.Security
WHERE     security_DPA_ NOT IN
                          (SELECT DISTINCT Security.Security_DPA_
                            FROM          Security INNER JOIN
                                                   datastream_Securities ON Security.Security_DPA_ = datastream_Securities.SecKnow_DPA INNER JOIN
                                                   datastream_Market ON Security.SecurityCode = datastream_Market.MktCode
                            WHERE      (FLOOR(CAST(datastream_Market.MktDate AS Float)) = FLOOR(CAST(GETDATE() AS Float))))






GO

CREATE VIEW dbo.ExistingOnlineUserList
AS
SELECT     dbo.Users.UserID, dbo.Users.OtherNames + '  ' + dbo.Users.Surname AS [USER], dbo.Users.UserName, dbo.Users.Description,
                      dbo.Users.SecretQuestion, dbo.Users.SecretAnswer, dbo.Users.RequiresSecretQuestion, dbo.Users.Client_DPA_, dbo.Client.ClientName,
                      dbo.Users.Password, dbo.Users.Surname, dbo.Users.OtherNames, dbo.Users.StaffID, dbo.Users.Removed, dbo.Users.FirstTime, dbo.Users.Expires,
                      dbo.Users.Enabled, dbo.Users.RemoteUser, dbo.Client.ClientEmail, dbo.Users.Accepted, dbo.Users.Email
FROM         dbo.Users INNER JOIN
                      dbo.Client ON dbo.Users.Client_DPA_ = dbo.Client.Client_DPA_
WHERE     (dbo.Users.RemoteUser = 1) AND (dbo.Users.Accepted = 1)



GO

CREATE VIEW dbo.FineTradingSchedule
AS
SELECT TOP 100 PERCENT OrdDetailSecurity, OrdDetailType, OrdDetailPrice, Order_DPA_, BalanceQty, OrderRef, OrderDate, SecurityCode,
               OrdDetailSecType, Client, code, Validity, Best, Amount, Price, Total, CurrentBal, CreditLimit, Excess
FROM  (SELECT dbo.Client.ClientName AS OrdDetailClient, dbo.Security.SecurityCode + ' : ' + CONVERT(NVARCHAR(1000), dbo.OrdDetail.OrdDetailQty)
                              AS OrdDetailItem, dbo.OrdDetail.OrdDetailPrice AS OrdDetailPrice, dbo.OrdDetail.OrdDetailQty AS OrdDetailQty,
                              BalanceQty = CASE ISNULL(dbo.OrdDetailContractedQtyList.ContractQty, 1)
                              WHEN 1 THEN dbo.OrdDetail.OrdDetailQty ELSE dbo.OrdDetail.OrdDetailQty - dbo.OrdDetailContractedQtyList.ContractQty END,
                              dbo.OrderSecType.OrderSecTypeDisplayName AS OrdDetailSecType, dbo.OrderSecType.OrderSecType_DPA_ AS OrderSecType_DPA_,
                              dbo.Security.SecurityName AS OrdDetailSecurity, dbo.OrderType.OrderTypeDescription AS OrdDetailType,
                              dbo.OrdDetail.OrdDetail_DPA_ AS OrdDetail_DPA_, dbo.[OrderList].Order_DPA_, dbo.[OrderList].OrderDate, dbo.[OrderList].OrderCanceled,
                              dbo.[OrderList].OrderHold, dbo.[OrderList].OrderRef, dbo.Security.Security_DPA_, dbo.Commission.CommissionRate,
                              dbo.OrderType.OrderTypeSale, dbo.OrdDetail.OrdDetailCompound, CONVERT(DATETIME, dbo.OrdDetail.OrdDetailValidity, 108) AS Validity,
                              dbo.OrdDetail.OrdDetailCertNo, dbo.Security.SecurityName, dbo.Client.Client_DPA_, dbo.Security.SecurityCode, Client.ClientName AS Client,
                              Client.Client_DPA_ AS Code, OrdDetail.Best, OrdDetail.amount, datastream_SecurityPriceList.Price, dbo.ClientTotal.Total,
                              dbo.ClientBalances.CurrentBal, dbo.Client.CreditLimit,
                              isnull((- dbo.Client.CreditLimit - (dbo.ClientBalances.CurrentBal - dbo.ClientTotal.Total)), 0) AS Excess
               FROM   dbo.[OrderList] INNER JOIN
                              dbo.OrderSecType ON dbo.[OrderList].OrderSecType_DPA_ = dbo.OrderSecType.OrderSecType_DPA_ INNER JOIN
                              dbo.OrderType ON dbo.[OrderList].OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                              dbo.Client ON dbo.[OrderList].Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                              dbo.OrdDetail ON dbo.[OrderList].Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                              dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                              dbo.Commission ON dbo.Client.Commission_DPA_ = dbo.Commission.Commission_DPA_ LEFT OUTER JOIN
                              dbo.OrdDetailContractedQtyList ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.OrdDetailContractedQtyList.OrdDetail_DPA_ LEFT OUTER JOIN
                              datastream_SecurityPriceList ON OrdDetail.Security_DPA_ = .datastream_SecurityPriceList.Security_DPA_ LEFT OUTER JOIN
                             ClientBalances ON Client.Client_DPA_ = ClientBalances.Client_DPA_ LEFT OUTER JOIN
                              ClientTotal ON Client.Client_DPA_ = Clienttotal.Client_DPA_) innerTBL
WHERE (innerTBL.OrderHold = 0 AND innerTBL.BalanceQty > 0)
GROUP BY OrdDetailSecurity, OrdDetailType, OrdDetailPrice, Order_DPA_, OrdDetail_DPA_, BalanceQty, OrderRef, OrderDate, SecurityCode,
               OrdDetailSecType, Client, code, Validity, Best, Amount, Price, Total, CurrentBal, CreditLimit, Excess
ORDER BY OrdDetail_DPA_


GO


CREATE VIEW dbo.ForwardRatesList
AS
SELECT     dbo.Bond.Bond_DPA_, dbo.Bond.BondIDate AS IssueDate, dbo.Bond.BondIssue, dbo.Security.SecurityCode AS BondType,
                      dbo.Bond.BondMDate AS MaturityDate, dbo.Bond.FaceValue, dbo.Bond.Security_DPA_, a.ForwardRate, a.ActivationDate, a.ModifiedBy, a.DateModified,
                      a.UserName AS [User], a.ForwardRate_DPA_
FROM         dbo.Bond INNER JOIN
                      dbo.Security ON dbo.Bond.Security_DPA_ = dbo.Security.Security_DPA_ LEFT OUTER JOIN
                          (SELECT     TOP 100 PERCENT FRate.ForwardRate_DPA_, FRate.ForwardRate, FRate.ActivationDate, FRate.Bond_DPA_, dbo.Users.UserName,
                                                   FRate.DateModified, FRate.ModifiedBy
                            FROM          dbo.ForwardRate FRate LEFT OUTER JOIN
                                                   dbo.Users ON FRate.ModifiedBy = dbo.Users.UserID
                            WHERE      (FLOOR(CAST(FRate.ActivationDate AS Float)) =
                                                       (SELECT     MAX(Floor(cast(ActivationDate AS Float))) AS UniqueDate
                                                         FROM          dbo.ForwardRate
                                                         WHERE      ActivationDate IS NOT NULL AND Frate.Bond_DPA_ = Bond_DPA_))) a ON dbo.Bond.Bond_DPA_ = a.Bond_DPA_



GO

CREATE VIEW dbo.ForwardsList
AS
SELECT     TOP 100 PERCENT dbo.Offerings.PAL_No, dbo.Offerings.Client_DPA_, dbo.Offerings.Offering_DPA_, dbo.Offerings.Offering,
                      dbo.Offerings.Offering_Price, dbo.Offerings.ID_No, dbo.Offerings.Alloted_Rights, dbo.Offerings.Accepted_Rights, dbo.Offerings.Renouncee,
                      dbo.Offerings.Receipt, dbo.Offerings.Submitted, dbo.Offerings.Submission_Date, dbo.Offerings.Ref_No, dbo.Client.ClientName,
                      dbo.Offerings.Alloted_Rights * dbo.Offerings.Offering_Price AS Amount_Payable, dbo.Security.SecurityName, dbo.Security.SecurityCode,
                      dbo.Offerings.Batch_No, dbo.Offerings.Offerings_Date, dbo.Offerings.OfferCheque, dbo.Offerings.OfferBank,
                      dbo.Users.Surname + ' ' + dbo.Users.OtherNames AS ModifiedBy, dbo.Offerings.TimeChanged, dbo.Offerings.BatchSeq,
                      ISNULL(dbo.Security.BatchSize, 0) AS BatchSize, CAST(FLOOR(CAST(dbo.Security.ClosingDate AS Float)) AS datetime) AS ClosingDate,
                      dbo.Offerings.Forward, dbo.Broker.BrokerName, dbo.Client.ClientIDPass, dbo.Offerings.Deleted, dbo.Offerings.PaymentRef,
                      dbo.Offerings.PaymentBankRef, dbo.Offerings.PaymentBranchRef, dbo.Offerings.PaymentSortCode, dbo.Offerings.PaymentAccountNo,
                      dbo.Offerings.EFTAccountNo, dbo.Offerings.EFTSortCode, dbo.Offerings.Status, dbo.Offerings.CDSPaid, dbo.Offerings.AgentCode,
                      dbo.Offerings.ReceivingBroker, dbo.Offerings.CDACode, dbo.Offerings.TaxExempt, dbo.Offerings.RefundMethod, dbo.Offerings.BatchFileName,
                      dbo.Offerings.LastDownLoaded, dbo.Offerings.Additional, dbo.Offerings.AcceptanceType, dbo.Offerings.Download_Date, dbo.Offerings.DateCreated,
                      dbo.Offerings.EFTBranchRef, dbo.Offerings.EFTBankRef, dbo.Offerings.PaymentType, dbo.Offerings.Category, dbo.Offerings.OfferingType,
                      dbo.Offerings.Updated, dbo.Offerings.Downloaded, dbo.Offerings.DividendMethod, dbo.Offerings.Certificate,
                      Users_1.Surname + ' ' + Users_1.OtherNames AS CreatedBy
FROM         dbo.Offerings INNER JOIN
                      dbo.Client ON dbo.Offerings.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Security ON dbo.Offerings.Offering = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.Broker ON dbo.Offerings.ReceivingBroker = dbo.Broker.Broker_DPA_ INNER JOIN
                      dbo.Users Users_1 ON dbo.Offerings.CreatedBy = Users_1.UserID LEFT OUTER JOIN
                      dbo.Users ON dbo.Offerings.ChangedBy = dbo.Users.UserID
WHERE     (dbo.Offerings.Forward = 1) AND (dbo.Offerings.Deleted = 0)
ORDER BY dbo.Offerings.Offering_DPA_ DESC



GO
CREATE VIEW dbo.FullClientList
AS
SELECT     TOP 100 PERCENT CAST(dbo.Client.Client_DPA_ AS NVARCHAR(4000)) + SPACE
                          ((SELECT     ISNULL(MAX(LEN(CAST(dbo.Client.Client_DPA_ AS NVARCHAR(4000)))), 0) AS MAXLEN
                              FROM         dbo.Client) - LEN(CAST(dbo.Client.Client_DPA_ AS NVARCHAR(4000)))) + ' : ' + dbo.Client.ClientName AS ClientNameEx,
                      LTRIM(RTRIM(dbo.Client.ClientName)) AS ClientName, dbo.OwnerList.OwnerName, dbo.Client.ClientContact, dbo.Client.ClientOfficeTel,
                      dbo.Client.ClientCellTel, dbo.Client.ClientEmail, dbo.Client.ClientAddr, dbo.Client.ClientVIP, dbo.Client.Client_DPA_, dbo.Client.ClientCDSNo,
                      dbo.Client.OnlineRegistration, dbo.Client.CreditLimit, dbo.Client.ClientFax, dbo.Client.Owner_DPA_, dbo.Agent.Agent_DPA_, dbo.Agent.AgentName,
                      dbo.Client.TimeChanged, dbo.UserList.[USER] AS ChangedBy, dbo.Client.IsCustodian, dbo.Client.Commission_DPA_
FROM         dbo.Client LEFT OUTER JOIN
                      dbo.UserList ON dbo.Client.ChangedBy = dbo.UserList.UserID LEFT OUTER JOIN
                      dbo.Agent ON dbo.Client.Agent_DPA_ = dbo.Agent.Agent_DPA_ LEFT OUTER JOIN
                      dbo.OwnerList ON dbo.Client.Owner_DPA_ = dbo.OwnerList.Owner_DPA_
WHERE     (dbo.Client.Deleted = 0)
ORDER BY dbo.Client.ClientName

GO

CREATE VIEW dbo.FullClientList2
AS
SELECT     TOP 100 PERCENT CAST(dbo.Client.Client_DPA_ AS NVARCHAR(4000)) + SPACE
                          ((SELECT     ISNULL(MAX(LEN(CAST(dbo.Client.Client_DPA_ AS NVARCHAR(4000)))), 0) AS MAXLEN
                              FROM         dbo.Client) - LEN(CAST(dbo.Client.Client_DPA_ AS NVARCHAR(4000)))) + ' : ' + dbo.Client.ClientName AS ClientNameEx,
                      LTRIM(RTRIM(dbo.Client.ClientName)) AS ClientName, dbo.OwnerList.OwnerName, dbo.Client.ClientContact, dbo.Client.ClientOfficeTel,
                      dbo.Client.ClientCellTel, dbo.Client.ClientEmail, dbo.Client.ClientAddr, dbo.Client.ClientVIP, dbo.Client.Client_DPA_, dbo.Client.ClientCDSNo,
                      dbo.Client.OnlineRegistration, dbo.Client.CreditLimit, dbo.Client.ClientFax, dbo.Client.Owner_DPA_, ISNULL(dbo.Client.Agent_DPA_, 0) AS Agent_DPA_,
                      ISNULL(dbo.Agent.AgentName, '') AS AgentName, dbo.Client.TimeChanged, dbo.UserList.[USER] AS ChangedBy, dbo.ClientStatementBalances.Balance,
                       dbo.Client.IsCustodian, dbo.Client.Commission_DPA_
FROM         dbo.Client LEFT OUTER JOIN
                      dbo.ClientStatementBalances ON dbo.Client.Client_DPA_ = dbo.ClientStatementBalances.Client_DPA_ LEFT OUTER JOIN
                      dbo.UserList ON dbo.Client.ChangedBy = dbo.UserList.UserID LEFT OUTER JOIN
                      dbo.Agent ON dbo.Client.Agent_DPA_ = dbo.Agent.Agent_DPA_ LEFT OUTER JOIN
                      dbo.OwnerList ON dbo.Client.Owner_DPA_ = dbo.OwnerList.Owner_DPA_
WHERE     (dbo.Client.Deleted = 0) OR
                      (dbo.Client.Deleted IS NULL)
ORDER BY dbo.Client.ClientName


GO
CREATE VIEW dbo.FullClientList2
AS
SELECT     TOP 100 PERCENT CAST(dbo.Client.Client_DPA_ AS NVARCHAR(4000)) + SPACE
                          ((SELECT     ISNULL(MAX(LEN(CAST(dbo.Client.Client_DPA_ AS NVARCHAR(4000)))), 0) AS MAXLEN
                              FROM         dbo.Client) - LEN(CAST(dbo.Client.Client_DPA_ AS NVARCHAR(4000)))) + ' : ' + dbo.Client.ClientName AS ClientNameEx,
                      LTRIM(RTRIM(dbo.Client.ClientName)) AS ClientName, dbo.OwnerList.OwnerName, dbo.Client.ClientContact, dbo.Client.ClientOfficeTel,
                      dbo.Client.ClientCellTel, dbo.Client.ClientEmail, dbo.Client.ClientAddr, dbo.Client.ClientVIP, dbo.Client.Client_DPA_, dbo.Client.ClientCDSNo,
                      dbo.Client.OnlineRegistration, dbo.Client.CreditLimit, dbo.Client.ClientFax, dbo.Client.Owner_DPA_, ISNULL(dbo.Client.Agent_DPA_, 0) AS Agent_DPA_,
                      ISNULL(dbo.Agent.AgentName, '') AS AgentName, dbo.Client.TimeChanged, dbo.UserList.[USER] AS ChangedBy, dbo.ClientStatementBalances.Balance,
                       dbo.Client.IsCustodian, dbo.Client.Commission_DPA_
FROM         dbo.Client LEFT OUTER JOIN
                      dbo.ClientStatementBalances ON dbo.Client.Client_DPA_ = dbo.ClientStatementBalances.Client_DPA_ LEFT OUTER JOIN
                      dbo.UserList ON dbo.Client.ChangedBy = dbo.UserList.UserID LEFT OUTER JOIN
                      dbo.Agent ON dbo.Client.Agent_DPA_ = dbo.Agent.Agent_DPA_ LEFT OUTER JOIN
                      dbo.OwnerList ON dbo.Client.Owner_DPA_ = dbo.OwnerList.Owner_DPA_
WHERE     (dbo.Client.Deleted = 0) OR
                      (dbo.Client.Deleted IS NULL)
ORDER BY dbo.Client.ClientName

GO

CREATE VIEW dbo.FullEntityTypeList
AS
SELECT     dbo.AccountType.AccountTypeName AS EntityTypeAccountType, dbo.EntityType.EntityTypeName, dbo.EntityType.EntityType_DPA_,
                      dbo.EntityType.DefaultSelection, dbo.EntityType.EntityTypeCode, dbo.AccountType.AccountType_DPA_
FROM         dbo.AccountType INNER JOIN
                      dbo.EntityType ON dbo.AccountType.AccountType_DPA_ = dbo.EntityType.AccountType_DPA_


GO
CREATE VIEW dbo.FullOrderList
AS
SELECT DISTINCT
                      TOP 100 PERCENT dbo.OrdDetailList.OrdDetail_DPA_, CONVERT(DATETIME, dbo.tbOrder.OrderDate, 108) AS OrderDate, dbo.tbOrder.OrderRef,
                      dbo.OrdDetailList.OrdDetailSecurity, dbo.OrdDetailList.OrdDetailPrice, dbo.OrdDetailList.OrdDetailQty, dbo.OrdDetailList.OrdDetailValidity,
                      dbo.tbOrder.OrderHold, dbo.OrdDetailList.BalanceQty, dbo.tbOrder.Order_DPA_, dbo.OrderHoldType.OrderHoldTypeName,
                      dbo.tbOrder.OrderAutoReleaseDate, dbo.tbOrder.OrderSecType_DPA_, dbo.tbOrder.Client_DPA_, dbo.tbOrder.OrderCanceled,
                      dbo.tbOrder.OrderType_DPA_, dbo.tbOrder.OrderCompounded AS OrdDetailCompound, dbo.tbOrder.OrderDateReleased,
                      dbo.OrdDetailList.SecurityCode,
                          (SELECT     COUNT(SubTbl.OrdDetail_DPA_) AS DetailPos
                            FROM          dbo.OrdDetailList SubTbl
                            WHERE      (SubTbl.Order_DPA_ = dbo.OrdDetailList.Order_DPA_) AND (SubTbl.OrdDetail_DPA_ <= dbo.OrdDetailList.OrdDetail_DPA_))
                      AS OrdDetailPos, dbo.OrdDetailList.OrdDetailSecType, dbo.OrdDetailList.BondDescription, dbo.Security.SecurityMktPrice,
                      dbo.datastream_SecurityPriceList.Price, dbo.OrdDetailList.Best, dbo.tbOrder.TimeChanged, UserList_1.[USER] AS ChangedBy, dbo.tbOrder.InterBank,
                      UserList_2.[USER] AS OrderReleasedBy, dbo.Client.ClientName, dbo.Client.ClientCDSNo, dbo.OrderType.OrderTypeDescription AS OrderTypeName,
                      dbo.OrderType.OrderTypeSale, dbo.OrdDetailList.Amount AS Amount
FROM         dbo.tbOrder INNER JOIN
                      dbo.OrdDetailList ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetailList.Order_DPA_ INNER JOIN
                      dbo.OrderHoldType ON dbo.tbOrder.OrderHoldType_DPA_ = dbo.OrderHoldType.OrderHoldType_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetailList.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.datastream_SecurityPriceList ON dbo.Security.Security_DPA_ = dbo.datastream_SecurityPriceList.Security_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ LEFT OUTER JOIN
                      dbo.UserList UserList_1 ON dbo.tbOrder.ChangedBy = UserList_1.UserID LEFT OUTER JOIN
                      dbo.UserList UserList_2 ON dbo.tbOrder.OrderReleasedBy = UserList_2.UserID
WHERE     (dbo.tbOrder.OrderCanceled = 0)
ORDER BY dbo.tbOrder.Order_DPA_ DESC, dbo.OrdDetailList.OrdDetailSecurity, dbo.OrdDetailList.OrdDetail_DPA_ DESC

GO
CREATE VIEW dbo.GenderList
AS
SELECT     GenderDescription AS GenderGender, Gender_DPA_, DefaultSelection
FROM         dbo.Gender

GO

CREATE VIEW dbo.GenericList
AS
SELECT     Generic_DPA_, Generic_EIT_, GenericDescription
FROM         dbo.Generic




GO

CREATE VIEW dbo.GenericSettingList
AS
SELECT     dbo.GenericSetting.GenericSetting_DPA_, dbo.GenericSetting.GenericSettingDescription, dbo.Generic.GenericDescription,
                      dbo.FullEntityTypeList.EntityTypeName, dbo.GenericSetting.Generic_DPA_, dbo.GenericSetting.EntityType_DPA_
FROM         dbo.GenericSetting INNER JOIN
                      dbo.Generic ON dbo.GenericSetting.Generic_DPA_ = dbo.Generic.Generic_DPA_ INNER JOIN
                      dbo.FullEntityTypeList ON dbo.GenericSetting.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_




GO
CREATE VIEW GroupAccessList AS SELECT TOP 100 PERCENT Groups.GroupName, MenuGroups.ID, Menus.mnuCaption
FROM (Groups INNER JOIN MenuGroups ON Groups.GroupID=MenuGroups.groupID) INNER JOIN Menus ON MenuGroups.MenuID=Menus.MenuID
WHERE (((Menus.mnuCaption)<>'Add' And (Menus.mnuCaption)<>'Delete' And (Menus.mnuCaption)<>'Edit') And (IsMainMenu=1))
ORDER BY Groups.GroupName, Menus.mnuCaption

GO
CREATE VIEW dbo.GroupMembersList AS SELECT UserGroups.MemberID, Groups.GroupName, Users.OtherNames + " " + Users.SurName AS Member
FROM (Groups INNER JOIN UserGroups ON Groups.GroupID = UserGroups.GroupID) INNER JOIN Users ON UserGroups.UserID = Users.UserID

GO

CREATE VIEW GroupsList AS SELECT TOP 100 PERCENT Groups.GroupID, Groups.GroupName, Groups.Description
FROM Groups
ORDER BY Groups.GroupName


GO
CREATE VIEW dbo.HeldOrderList
AS
SELECT     Order_DPA_ AS [Order No], OrderDate AS [Order Date], ClientName + '	[' + CONVERT(nvarchar(4000), Client_DPA_) + ']' AS Client,
                      OrdDetailSecurity AS Security, OrdDetailQty AS [Order Qty], OrdDetailPrice AS [Order Price]
FROM         dbo.FullOrderList
WHERE     (OrderHold = 1)



GO
CREATE VIEW dbo.HistoryTransactionsList
AS




SELECT * FROM dbo.LevyTransactionHistoryList

UNION ALL

SELECT  CAST(Account_DPA_ AS NVARCHAR(500)) AS Account_DPA_, TransDate, CAST(REF AS NVARCHAR(500)) AS REF, Particulars, Debit, Credit, IsOpeningBalance, Balance,
	(SELECT AccountName FROM Account dd WHERE dd.Account_DPA_ = dbo.NominalTransactionList.Account_DPA_) EntityName FROM dbo.NominalTransactionList
GO
CREATE VIEW dbo.HoldingExceptions
AS
SELECT     dbo.excep_SummaryHoldings.Code, dbo.Client.ClientName, dbo.Security.SecurityName, dbo.excep_SummaryHoldings.CDSQty,
                      dbo.excep_SummaryHoldings.BKQty, dbo.Client.ClientCDSNo, dbo.excep_SummaryHoldings.DiffQty, dbo.Client.IsNominee, dbo.Client.IsCustodian,
                      dbo.Security.Security_DPA_
FROM         dbo.excep_SummaryHoldings INNER JOIN
                      dbo.Client ON dbo.excep_SummaryHoldings.Code = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Security ON dbo.excep_SummaryHoldings.SecCode = dbo.Security.Security_DPA_
WHERE     (dbo.Client.IsFrozen <> 1) AND (dbo.Client.Deleted <> 1)

GO




CREATE VIEW dbo.HoldingsSecurityPriceList
AS
SELECT     dbo.datastream_SecurityPriceList.Price, dbo.Holdings.*, dbo.Security.SecurityCode AS SecurityCode
FROM         dbo.datastream_SecurityPriceList INNER JOIN
                      dbo.Holdings ON dbo.datastream_SecurityPriceList.Security_DPA_ = dbo.Holdings.Security_DPA_ INNER JOIN
                      dbo.Security ON dbo.datastream_SecurityPriceList.Security_DPA_ = dbo.Security.Security_DPA_





GO

CREATE VIEW dbo.HoldOrderList
AS
SELECT     OrderDate, OrderTypeName, OrderRef, ClientName, OrderHold, Order_DPA_, Client_DPA_, OrderHoldType_DPA_, ISNULL(CONVERT(nvarchar,
                      OrderAutoReleaseDate), '') AS OrderAutoReleaseDate, ClientCDSNo, SecurityName, OrdDetailQty, Quantity, CurrentBal, BalanceFree,
                      OrderDateReleased, OrdDetail_DPA_, SecurityType, BondIssue
FROM         dbo.OrderListPlain
WHERE     (Order_DPA_ IN
                          (SELECT     Order_DPA_
                            FROM          LotList))


GO

CREATE VIEW dbo.HolidayList
AS
SELECT     Holiday_DPA_, Description, CONVERT(NVARCHAR(3), DATEPART(DD, Holiday)) + ' ' + DATENAME(MM, Holiday) AS HOLIDAY
FROM         dbo.Holidays


GO
create VIEW [dbo].[InstitutionList]
AS
SELECT     dbo.Institution.*
FROM         dbo.Institution
GO

CREATE VIEW dbo.Interbanks
AS
SELECT     TOP 100 PERCENT dbo.OrdDetailList.Order_DPA_, dbo.OrdDetailList.OrdDetail_DPA_, dbo.OrdDetailList.OrdDetailType,
                      dbo.OrdDetailList.OrdDetailClient, dbo.OrdDetailList.OrdDetailSecurity, dbo.OrdDetailList.OrdDetailQty, dbo.OrdDetailList.OrdDetailPrice,
                      dbo.OrdDetailList.BalanceQty, dbo.Lots.LotSlipNo, CONVERT(SMALLDATETIME, CAST(dbo.Lots.LotTDate AS CHAR(12))) AS LotTDate, dbo.Lots.LotQty,
                      dbo.Lots.LotPrice, dbo.Lots.LotGrossAmount, dbo.Broker.BrokerCode, dbo.Broker.BrokerName, dbo.Contracts.Contract_DPA_,
                      dbo.Status.StatusDescription, Lot_DPA_ = CASE ISNULL(dbo.Lots.Lot_DPA_, - 1) WHEN - 1 THEN 0 ELSE dbo.Lots.Lot_DPA_ END,
                      dbo.tbOrder.OrderDate, dbo.Lots.ContractNumber, dbo.OrdDetailList.OrdDetailSecType, dbo.OrdDetailList.OrderTypeSale,
                      dbo.OrdDetailList.Security_DPA_, dbo.OrdDetailList.CommissionRate, dbo.OrdDetailList.Client_DPA_, dbo.OrdDetailList.AgentCommission,
                      dbo.OrdDetailList.StaffCommission, dbo.OrdDetailList.SecurityCode, dbo.OrdDetailList.VolumeBoundary, dbo.OrdDetailList.VolumeRate,
                      dbo.OrdDetailList.MinimumCommission, dbo.OrdDetailList.CMARegulated, dbo.OrdDetailList.PostImmobilisedRate,
                      dbo.OrdDetailList.SecurityImmobilised, dbo.OrdDetailList.BondDescription, dbo.Lots.CDSTransaction, Contracts.ContractSettlementDate,
                      OrdDetailList.ClientCDSNo, OrdDetailList.InterBank, Contracts.IsInterBank, OrdDetailList.IsCustodian
FROM         dbo.Status RIGHT OUTER JOIN
                      dbo.Contracts RIGHT OUTER JOIN
                      dbo.Broker RIGHT OUTER JOIN
                      dbo.Lots RIGHT OUTER JOIN
                      dbo.OrdDetailList ON dbo.Lots.OrdDetail_DPA_ = dbo.OrdDetailList.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetailList.Order_DPA_ = dbo.tbOrder.Order_DPA_ ON dbo.Broker.Broker_DPA_ = dbo.Lots.Broker_DPA_ ON
                      dbo.Contracts.Contract_DPA_ = dbo.Lots.Contract_DPA_ ON dbo.Status.Status_DPA_ = dbo.Contracts.Status_DPA_
WHERE     (dbo.OrdDetailList.OrderHold = 0) AND tborder.interbank = 1
ORDER BY dbo.Contracts.Contract_DPA_ ASC



GO

CREATE VIEW dbo.InterbankSchedule
AS
SELECT     TOP 100 PERCENT dbo.Lot.LotGrossAmount, dbo.Lot.LotSlipNo, dbo.OrderType.OrderTypeDescription, dbo.Client.Client_DPA_, dbo.Client.ClientName,
                      dbo.tbOrder.OrderType_DPA_, dbo.Security.SecurityCode, dbo.Lot.LotPrice, dbo.Lot.LotQty, CAST(FLOOR(CAST(dbo.Lot.LotTDate AS Float))
                      AS DateTime) AS LotTDate, dbo.Lot.ContractNumber, dbo.Client.IsCustodian, dbo.Security.OrderSecType_DPA_,
                      CAST(FLOOR(CAST(dbo.Contract.ContractSettlementDate AS Float)) AS DateTime) AS SettlementDate, a.Contract_DPA_,
                      dbo.InterTransferType.TypeDescription, dbo.InterTransferType.InterTransferType_DPA_, dbo.InterTransfer.TransferDate
FROM         dbo.tbOrder INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                      dbo.Contract INNER JOIN
                      dbo.Lot ON dbo.Contract.Contract_DPA_ = dbo.Lot.Contract_DPA_ ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.Lot.OrdDetail_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                      dbo.InterTransfer ON dbo.Contract.Contract_DPA_ = dbo.InterTransfer.Contract_DPA_ INNER JOIN
                      dbo.InterTransferType ON dbo.InterTransfer.InterTransferType_DPA_ = dbo.InterTransferType.InterTransferType_DPA_ LEFT OUTER JOIN
                          (SELECT     *
                            FROM          Intertransfer
                            WHERE      InterTransferType_DPA_ = 1) a ON dbo.Contract.Contract_DPA_ = a.Contract_DPA_
WHERE     (dbo.InterTransferType.InterTransferType_DPA_ <> 2)
ORDER BY dbo.tbOrder.OrderType_DPA_ DESC, dbo.Lot.LotSlipNo


GO
CREATE VIEW dbo.InterBankTransferList
AS
SELECT     dbo.InterTransfers.*, 'Inter Bank' AS TransferType, dbo.InterTransfer.TransferDate AS InterBankDate,
                      dbo.InterTransfer.TransferReference AS InterBankReference, dbo.InterTransfer.TransferNarrative AS InterBankNarrative
FROM         dbo.InterTransfers LEFT OUTER JOIN
                      dbo.InterTransfer ON dbo.InterTransfers.Contract_DPA_ = dbo.InterTransfer.Contract_DPA_
WHERE     (dbo.InterTransfers.InterBank = 1) AND (dbo.InterTransfers.IsInterBank = 1) AND (dbo.InterTransfer.Deleted = 0 OR
                      dbo.InterTransfer.Deleted IS NULL)

GO

CREATE VIEW dbo.InterTransfers
AS
SELECT     TOP 100 PERCENT dbo.OrdDetailList.Order_DPA_, dbo.OrdDetailList.OrdDetail_DPA_, dbo.OrdDetailList.OrdDetailType,
                      dbo.OrdDetailList.OrdDetailClient, dbo.OrdDetailList.OrdDetailSecurity, dbo.OrdDetailList.OrdDetailQty, dbo.OrdDetailList.OrdDetailPrice,
                      dbo.OrdDetailList.BalanceQty, dbo.Lots.LotSlipNo, CONVERT(SMALLDATETIME, CAST(dbo.Lots.LotTDate AS CHAR(12))) AS LotTDate, dbo.Lots.LotQty,
                      dbo.Lots.LotPrice, dbo.Lots.LotGrossAmount, dbo.Broker.BrokerCode, dbo.Broker.BrokerName, dbo.Contracts.Contract_DPA_,
                      dbo.Status.StatusDescription, Lot_DPA_ = CASE ISNULL(dbo.Lots.Lot_DPA_, - 1) WHEN - 1 THEN 0 ELSE dbo.Lots.Lot_DPA_ END,
                      dbo.tbOrder.OrderDate, dbo.Lots.ContractNumber, dbo.OrdDetailList.OrdDetailSecType, dbo.OrdDetailList.OrderTypeSale,
                      dbo.OrdDetailList.Security_DPA_, dbo.OrdDetailList.CommissionRate, dbo.OrdDetailList.Client_DPA_, dbo.OrdDetailList.AgentCommission,
                      dbo.OrdDetailList.StaffCommission, dbo.OrdDetailList.SecurityCode, dbo.OrdDetailList.VolumeBoundary, dbo.OrdDetailList.VolumeRate,
                      dbo.OrdDetailList.MinimumCommission, dbo.OrdDetailList.CMARegulated, dbo.OrdDetailList.PostImmobilisedRate,
                      dbo.OrdDetailList.SecurityImmobilised, dbo.OrdDetailList.BondDescription, dbo.Lots.CDSTransaction, Contracts.ContractSettlementDate,
                      OrdDetailList.ClientCDSNo, OrdDetailList.InterBank, Contracts.IsInterBank, OrdDetailList.IsCustodian
FROM         dbo.Status RIGHT OUTER JOIN
                      dbo.Contracts RIGHT OUTER JOIN
                      dbo.Broker RIGHT OUTER JOIN
                      dbo.Lots RIGHT OUTER JOIN
                      dbo.OrdDetailList ON dbo.Lots.OrdDetail_DPA_ = dbo.OrdDetailList.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetailList.Order_DPA_ = dbo.tbOrder.Order_DPA_ ON dbo.Broker.Broker_DPA_ = dbo.Lots.Broker_DPA_ ON
                      dbo.Contracts.Contract_DPA_ = dbo.Lots.Contract_DPA_ ON dbo.Status.Status_DPA_ = dbo.Contracts.Status_DPA_
WHERE     (dbo.OrdDetailList.OrderHold = 0)
ORDER BY dbo.Contracts.Contract_DPA_ ASC



GO



CREATE VIEW dbo.IPODeletedOfferingList
AS
SELECT     TOP 100 PERCENT dbo.Offerings.Offering_DPA_, dbo.Offerings.PAL_No, dbo.Offerings.Client_DPA_, dbo.Offerings.Offering_Price, dbo.Offerings.ID_No,
                      dbo.Offerings.Alloted_Rights, dbo.Offerings.Accepted_Rights, dbo.Offerings.Renouncee, dbo.Offerings.Receipt, dbo.Offerings.Submitted,
                      dbo.Offerings.Submission_Date, dbo.Offerings.Ref_No, dbo.Offerings.Batch_No, CASE isnull(dbo.Offerings.Forward, 0)
                      WHEN 1 THEN 'Forward' ELSE 'Application' END AS APPStatus, dbo.Offerings.Offerings_Date, dbo.Offerings.OfferCheque, dbo.Offerings.OfferBank,
                      dbo.Offerings.TimeChanged, dbo.Offerings.Deleted, dbo.Offerings.Paid, dbo.Offerings.AppType, dbo.Offerings.shNAME2, dbo.Offerings.BatchSeq,
                      dbo.Offerings.Download_DPA_, dbo.Offerings.Alloted_Rights * dbo.Offerings.Offering_Price AS Amount_Payable,
                      dbo.Users.Surname + ' ' + dbo.Users.OtherNames AS ModifiedBy, dbo.Offerings.Download_Date, dbo.Offerings.ClientCellTel, dbo.Offerings.sms,
                      dbo.Offerings.PaymentMode, dbo.Client.ClientCDSNo, dbo.Offerings.smsNotes, dbo.Offerings.Extra, dbo.Offerings.AcceptanceType,
                      dbo.Offerings.CreatedBy, dbo.Offerings.DateCreated, dbo.Offerings.Payment_DPA_, dbo.Offerings.Downloaded, dbo.Offerings.Branch_DPA_,
                      dbo.Offerings.LastDownLoaded, dbo.Offerings.BatchPaymentMode, dbo.Offerings.BatchClosed, dbo.Offerings.CDSNumeric,
                      dbo.Offerings.BatchFileName, dbo.Offerings.Additional, dbo.Offerings.RefundMethod, dbo.Offerings.TaxExempt, dbo.Offerings.DividendMethod,
                      dbo.Offerings.EFTBankRef, dbo.Offerings.EFTBranchRef, dbo.Offerings.EFTSortCode, dbo.Offerings.EFTAccountNo, dbo.Offerings.Category,
                      dbo.Offerings.PaymentType, dbo.Offerings.PaymentRef, dbo.Offerings.PaymentBankRef, dbo.Offerings.PaymentBranchRef,
                      dbo.Offerings.PaymentSortCode, dbo.Offerings.PaymentAccountNo, dbo.Offerings.Updated, dbo.Offerings.OfferingType, dbo.Offerings.Certificate,
                      dbo.Offerings.Agent_DPA_, dbo.Offerings.CDSCharge, dbo.Offerings.CDACode, dbo.Offerings.AgentCode, dbo.Security.SecurityName,
                       dbo.Security.SecurityCode, dbo.Client.ClientName, dbo.Offerings.Status
FROM         dbo.Offerings INNER JOIN
                      dbo.Client ON dbo.Offerings.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Security ON dbo.Offerings.Offering = dbo.Security.Security_DPA_ LEFT OUTER JOIN
                      dbo.Users ON dbo.Offerings.ChangedBy = dbo.Users.UserID
WHERE     (ISNULL(dbo.Offerings.Deleted, 0) = 1) AND (dbo.Offerings.ChangedBy <> 10)
ORDER BY dbo.Offerings.Offering_DPA_ DESC




GO
CREATE VIEW dbo.IssueList
AS
SELECT     TOP 100 PERCENT dbo.Bond.Bond_DPA_, dbo.Security.SecurityCode, dbo.Security.SecurityName, dbo.Bond.BondIssue, dbo.Bond.BondIDate,
                      dbo.Bond.BondMDate, dbo.Bond.BondLife, dbo.Bond.BondPayment, dbo.Bond.BondRate, dbo.Security.Security_DPA_, dbo.Bond.Determination,
                      dbo.Bond.FaceValue, dbo.Security.OrderSecType_DPA_, dbo.Bond.BondRate AS Rate, dbo.Bond.ModifiedBy, dbo.Bond.DateModified,
                      dbo.Users.UserName AS [User], a.ForwardRate
FROM         dbo.Bond INNER JOIN
                      dbo.Security ON dbo.Bond.Security_DPA_ = dbo.Security.Security_DPA_ LEFT OUTER JOIN
                      dbo.Users ON dbo.Bond.ModifiedBy = dbo.Users.UserID LEFT OUTER JOIN
                          (SELECT     TOP 100 PERCENT FRate.ForwardRate_DPA_, FRate.ForwardRate, FRate.ActivationDate, FRate.Bond_DPA_, dbo.Users.UserName,
                                                   FRate.DateModified, FRate.ModifiedBy
                            FROM          dbo.ForwardRate FRate LEFT OUTER JOIN
                                                   dbo.Users ON FRate.ModifiedBy = dbo.Users.UserID
                            WHERE      (FLOOR(CAST(FRate.ActivationDate AS Float)) =
                                                       (SELECT     MAX(Floor(cast(ActivationDate AS Float))) AS UniqueDate
                                                         FROM          dbo.ForwardRate
                                                         WHERE      ActivationDate IS NOT NULL AND Frate.Bond_DPA_ = Bond_DPA_))) a ON dbo.Bond.Bond_DPA_ = a.Bond_DPA_
WHERE     (dbo.Security.OrderSecType_DPA_ = 1)
ORDER BY dbo.Security.SecurityName

GO


CREATE VIEW dbo.JournalEntriesList
AS
SELECT     TOP 100 PERCENT CASE Entry WHEN '' THEN Journal_DPA_ ELSE Entry END AS Entry, Narrative, Debit, Credit, AccountName, ControlAccount,
                      TransDate, Entity_DPA_
FROM         (SELECT     CAST(COUNT(*) AS NVARCHAR(500)) AS Entry, a.JournalNarrative AS Narrative, CAST(a.JournalEntryDebit AS NVARCHAR(500)) AS Debit,
                                              CAST(a.JournalEntryCredit AS NVARCHAR(500)) AS Credit, a.JournalEntryAccount AS AccountName, a.ControlAccount,
                                              a.JournalDate AS TransDate, a.Journal_DPA_, Entity_DPA_
                       FROM          (SELECT     JournalEntry_DPA_, JournalNarrative, JournalEntryDebit, JournalEntryCredit, JournalEntryAccount,
                                                                      '(' + EntityTypeName + ')' AS ControlAccount, Journal_DPA_, JournalDate, Entity_DPA_
                                               FROM          dbo.JournalList INNER JOIN
                                                                      dbo.FullEntityTypeList ON dbo.JournalList.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_) a CROSS JOIN
                                                  (SELECT     JournalEntry_DPA_, JournalNarrative, JournalEntryDebit, JournalEntryCredit, JournalEntryAccount,
                                                                           '(' + EntityTypeName + ')' AS ControlAccount, Journal_DPA_, JournalDate
                                                    FROM          dbo.JournalList INNER JOIN
                                                                           dbo.FullEntityTypeList ON dbo.JournalList.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_) b
                       WHERE      a.Journal_DPA_ = b.Journal_DPA_ AND a.JournalEntry_DPA_ >= b.JournalEntry_DPA_
                       GROUP BY a.Journal_DPA_, a.JournalNarrative, a.JournalEntryDebit, a.JournalEntryCredit, a.JournalEntryAccount, a.ControlAccount, a.JournalDate,
                                              A.Entity_DPA_
                       UNION ALL
                       SELECT     '' AS Entry, CAST(JournalDate AS NVARCHAR(500)) AS Narrative, '' AS Debit, '' AS Credit, '' AS AccountName, '' AS ControlAccount,
                                             JournalDate AS TransDate, Journal_DPA_, Entity_DPA_
                       FROM         dbo.JournalList
                       GROUP BY JournalDate, Journal_DPA_, Entity_DPA_) allData
WHERE     (Debit > N'0') AND (Credit > N'0')



GO
CREATE VIEW dbo.JournalFullList
AS
SELECT     TOP 100 PERCENT dbo.Journal.Journal_DPA_, dbo.JournalEntry.JournalEntry_DPA_, dbo.Journal.JournalDate,
                      dbo.Users.UserID AS JournalEntryUserID, dbo.Users.UserName AS JournalEntryUser, dbo.CompleteEntityList.EntityName AS JournalEntryAccount,
                      dbo.CompleteEntityList.EntityType AS JournalEntryEntity, dbo.JournalEntry.JournalEntryDebit, dbo.JournalEntry.JournalEntryCredit,
                      dbo.Journal.JournalNarrative, dbo.Journal.JournalCommitted, dbo.JournalEntry.Entity_DPA_, dbo.JournalEntry.EntityType_DPA_,
                      dbo.Journal.Released, dbo.Journal.ReleaseDate, UserList_1.[USER] AS ChangedBy, dbo.Journal.TimeChanged, UserList_1.[USER] AS ReleasedBy,
                      dbo.Journal.Deleted, dbo.JournalEntry.ReconcileDate, dbo.JournalEntry.Narrative AS JournalEntryNarrative
FROM         dbo.Journal INNER JOIN
                      dbo.JournalEntry ON dbo.Journal.Journal_DPA_ = dbo.JournalEntry.Journal_DPA_ INNER JOIN
                      dbo.Users ON dbo.Journal.UserID = dbo.Users.UserID INNER JOIN
                      dbo.CompleteEntityList ON dbo.JournalEntry.EntityType_DPA_ = dbo.CompleteEntityList.EntityType_DPA_ AND
                      dbo.JournalEntry.Entity_DPA_ = dbo.CompleteEntityList.Entity_DPA_ INNER JOIN
                      dbo.UserList UserList_1 ON dbo.Journal.UserID = UserList_1.UserID
WHERE     (dbo.JournalEntry.Deleted = 0) AND (dbo.Journal.Deleted = 0)
ORDER BY dbo.Journal.Journal_DPA_ DESC

GO

CREATE VIEW dbo.JournalList
AS
SELECT     TOP 100 PERCENT *
FROM         dbo.JournalFullList
WHERE     (JournalCommitted = 1)
ORDER BY Journal_DPA_ DESC


GO

CREATE VIEW dbo.JournalRelease
AS
SELECT     dbo.JournalList.Journal_DPA_, dbo.JournalList.JournalDate, dbo.JournalList.JournalEntryUserID, dbo.JournalList.JournalEntryUser,
                      dbo.JournalList.JournalNarrative, dbo.JournalList.JournalCommitted, dbo.Journal.Released, dbo.Journal.ReleaseDate,
                      SUM(dbo.JournalList.JournalEntryDebit) AS DebitAmount, SUM(dbo.JournalList.JournalEntryCredit) AS CreditAmount
FROM         dbo.Journal INNER JOIN
                      dbo.JournalList ON dbo.Journal.Journal_DPA_ = dbo.JournalList.Journal_DPA_
WHERE     (dbo.Journal.Released = 0) AND (dbo.Journal.Deleted = 0)
GROUP BY dbo.JournalList.Journal_DPA_, dbo.JournalList.JournalDate, dbo.JournalList.JournalEntryUserID, dbo.JournalList.JournalEntryUser,
                      dbo.JournalList.JournalNarrative, dbo.JournalList.JournalCommitted, dbo.Journal.Released, dbo.Journal.ReleaseDate


GO
CREATE VIEW dbo.JournalTotalList
AS
SELECT     Journal_DPA_, SUM(JournalEntryDebit) AS JournalDebitTotal, SUM(JournalEntryCredit) AS JournalCreditTotal
FROM         dbo.JournalFullList
GROUP BY Journal_DPA_

GO

CREATE VIEW dbo.kcbSchedules
AS
SELECT DISTINCT ScdID, COUNT(PayRecID) AS Total, SUM(AmtPayed) AS AmtPaid, SUM(NoShares) AS TotalShares
FROM         tblApplDetails
GROUP BY ScdID


GO

CREATE VIEW dbo.LatestPortfolios
AS
SELECT     dbo.PortfoliosQuantities.*, dbo.Client.ClientName AS ClientName
FROM         dbo.PortfoliosQuantities INNER JOIN
                      dbo.Client ON dbo.PortfoliosQuantities.Client_DPA_ = dbo.Client.Client_DPA_
WHERE     (dbo.Client.Deleted = 0)


GO


CREATE VIEW [dbo].[LeviesReport]
AS
/*SELECT     TOP 100 PERCENT dbo.Contract.Contract_DPA_, dbo.Contract.ContractTransferNo, dbo.Contract.ContractDeliveryDate, dbo.Lot.Lot_DPA_,
                      dbo.Lot.LotSlipNo, CAST(FLOOR(CAST(dbo.Lot.LotTDate AS float)) AS datetime) AS TransDate, dbo.Lot.LotQty, dbo.Lot.LotGrossAmount AS LotPrice,
                      dbo.Lot.ContractNumber, dbo.Contract.ContractDelivered, dbo.Contract.ContractNCertificate, dbo.Contract.ContractNCDate,
                      dbo.Contract.ContractNCDelivered, dbo.Contract.Voucher_DPA_, dbo.Contract.ContractVouchered, dbo.LevyContract.LevyAmount,
                      dbo.LevyContract.LevyName, dbo.tbOrder.OrderType_DPA_ AS OrdDetailType, dbo.Broker.BrokerCode, dbo.Client.ClientName AS OrdDetailClient,
                      dbo.Security.SecurityName AS OrdDetailSecurity, dbo.Security.OrderSecType_DPA_ AS OrdDetailSecType, dbo.Security.SecurityCode,
                      dbo.LevyContract.SystemMaintained, dbo.LevyContract.LevyShortName
FROM         dbo.Contract INNER JOIN
                      dbo.Lot ON dbo.Contract.Contract_DPA_ = dbo.Lot.Contract_DPA_ INNER JOIN
                      dbo.LevyContract ON dbo.Contract.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Broker ON dbo.Lot.Broker_DPA_ = dbo.Broker.Broker_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_
WHERE     (dbo.Client.Deleted = 0) AND (dbo.tbOrder.Deleted = 0) AND (dbo.LevyContract.Deleted = 0) AND (dbo.Contract.Deleted = 0) AND (dbo.Lot.Deleted = 0)
                      AND (dbo.OrdDetail.Deleted = 0)
ORDER BY dbo.tbOrder.OrderType_DPA_


*/



SELECT
	TOP 100 PERCENT Contract.Contract_DPA_, Contract.ContractTransferNo,
	Contract.ContractDeliveryDate, Lot.Lot_DPA_, Lot.LotSlipNo,
    CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime) AS TransDate,
	Lot.LotQty, Lot.LotGrossAmount AS LotPrice, Lot.ContractNumber,
	Contract.ContractDelivered, Contract.ContractNCertificate,
	Contract.ContractNCDate, Contract.ContractNCDelivered, Contract.Voucher_DPA_,
    Contract.ContractVouchered,
    CASE
		WHEN SystemMaintained = 11
		THEN LevyContract.LevyAmount - (SELECT levyamount FROM	LevyContract
        WHERE      systemMaintained = 25 AND deleted = 0 AND Contract_DPA_ = Contract.Contract_DPA_)
          ELSE dbo.LevyContract.LevyAmount END AS LevyAmount, LevyContract.LevyName, tbOrder.OrderType_DPA_ AS OrdDetailType, Broker.BrokerCode,
          Client.ClientName AS OrdDetailClient, Security.SecurityName AS OrdDetailSecurity, Security.OrderSecType_DPA_ AS OrdDetailSecType,
          Security.SecurityCode, LevyContract.SystemMaintained, LevyContract.LevyShortName
FROM         Contract INNER JOIN
          Lot ON Contract.Contract_DPA_ = Lot.Contract_DPA_ INNER JOIN
          LevyContract ON Contract.Contract_DPA_ = LevyContract.Contract_DPA_ INNER JOIN
          OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
          tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN
          Broker ON Lot.Broker_DPA_ = Broker.Broker_DPA_ INNER JOIN
          Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
          Security ON OrdDetail.Security_DPA_ = Security.Security_DPA_
WHERE     (Client.Deleted = 0) AND (tbOrder.Deleted = 0) AND (LevyContract.Deleted = 0) AND (Contract.Deleted = 0) AND (Lot.Deleted = 0) AND
          (OrdDetail.Deleted = 0)
ORDER BY tbOrder.OrderType_DPA_


GO

CREATE VIEW dbo.LeviesReportByDate
AS
SELECT     TOP 100 PERCENT CAST(FLOOR(CAST(TransDate AS float)) AS datetime) AS TransDate, SUM(LotPrice) AS Gross, SUM(LevyAmount) AS Commission,
                      SystemMaintained, LevyShortName
FROM         dbo.LeviesReport
GROUP BY CAST(FLOOR(CAST(TransDate AS float)) AS datetime), SystemMaintained, LevyShortName
ORDER BY CAST(FLOOR(CAST(TransDate AS float)) AS datetime)


GO
CREATE VIEW dbo.LevyContractList
AS
SELECT     dbo.Contract.Contract_DPA_, dbo.LevyContract.LevyContract_DPA_, dbo.LevyContract.LevyName, dbo.LevyContract.LevyAmount,
                      dbo.LevyContract.LevyRate, dbo.LevyContract.LevyBlock, dbo.Lot.Lot_DPA_, dbo.LevyContract.SystemMaintained, dbo.Lot.ContractNumber,
                      CONVERT(SMALLDATETIME, CAST(dbo.Lot.LotTDate AS CHAR(12))) AS LotTDate, dbo.LevyContract.LevyShortName
FROM         dbo.Contract INNER JOIN
                      dbo.LevyContract ON dbo.Contract.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.Lot ON dbo.Contract.Contract_DPA_ = dbo.Lot.Contract_DPA_

GO
CREATE VIEW dbo.LevyContracts
AS
SELECT     TOP 100 PERCENT Contract_DPA_, LevyAmount, LevyName, LevyContract_DPA_, LevyContract_EIT_, LevyRate, LevyBlock, LevyRatePercentage,
                      SystemMaintained, LevyShortName, LevyVATAmount
FROM         dbo.LevyContract
WHERE     (Deleted = 0)
ORDER BY LevyShortName

GO
CREATE VIEW dbo.LevyContracts_v2 AS SELECT TOP 100 PERCENT Contract_DPA_, LevyAmount, LevyName, LevyContract_DPA_, LevyContract_EIT_, LevyRate, LevyBlock, LevyRatePercentage, SystemMaintained, LevyShortName FROM dbo.LevyContract WHERE (Deleted = 0) AND (LevyShortName <> 'Agent') AND (LevyShortName <> 'Transfer') ORDER BY LevyShortName
GO
CREATE VIEW dbo.LevyContractSysMaintainedList
AS
SELECT     SystemMaintained
FROM         dbo.Levy
UNION
SELECT DISTINCT SystemMaintained
FROM         LevyContract



GO
CREATE VIEW dbo.LevyList
AS
SELECT     TOP 100 PERCENT LevyDescription, Levy_DPA_, LevyAmount, LevyType, LevyAppBond, LevyAppSecurity, LevyBlock, LevyActive, SystemMaintained,
                      LevyShortName, Vatable
FROM         dbo.Levy
ORDER BY LevyDescription

GO

CREATE VIEW dbo.LevyOrderList
AS
SELECT     TOP 100 PERCENT dbo.LevyReportOrder.*
FROM         dbo.LevyReportOrder
ORDER BY LevyOrder


GO
CREATE VIEW dbo.LevyOrderListSource
AS
SELECT     innerTBL.*
FROM         (SELECT DISTINCT LevyShortName As LevyName
                       FROM          dbo.LevyContract
                       UNION
                       SELECT     LevyShortName AS LevyName
                       FROM         dbo.Levy
                       WHERE     (LevyActive = 1)) innerTBL
WHERE     (LevyName NOT IN
                          (SELECT DISTINCT LevyName
                            FROM          LevyReportOrder))
GO
CREATE VIEW dbo.LevySecurityList
AS
SELECT     LevySecurity_DPA_, Levy_DPA_, Security_DPA_
FROM         dbo.LevySecurity

GO
CREATE VIEW dbo.LevyTransactionHistoryList
AS

SELECT     TOP 100 PERCENT '------' + CAST(Entity_DPA_ AS NVARCHAR(500)) Entity_DPA_, TransDate, REF
	,
	CASE
		WHEN ISDATE(Particulars) = 0 THEN Particulars
		ELSE	(SELECT EntityName FROM dbo.Entity dd WHERE dd.Entity_DPA_ = LevyTransactionHistory.Entity_DPA_) + ' - ' +  Particulars
	END AS Particulars
	,  Debit, Credit, IsOpeningBalance, Credit - Debit AS Balance
	, (SELECT EntityName FROM dbo.Entity dd WHERE dd.Entity_DPA_ = LevyTransactionHistory.Entity_DPA_) As [EntityName]

FROM
(SELECT     Entity_DPA_, EntityRegDate AS TransDate, '*******' AS REF, ' Opening Balance' AS Particulars, case when EntityOpeningBal < 0 then EntityOpeningBal else 0 end AS Debit, case when EntityOpeningBal >= 0 then EntityOpeningBal else 0 end AS Credit,
                                              1 AS IsOpeningBalance
                       FROM          dbo.Entity WHERE SystemMaintained = 1
UNION ALL

SELECT     dbo.Payment.Entity_DPA_ AS Entity_DPA_, dbo.Payment.PaymentPDate AS TransDate, CAST(dbo.Payment.PaymentReference AS VARCHAR(500)) AS Ref,
                     dbo.Payment.PaymentNarrative AS Particulars, CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                     CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit, 0 AS IsOpeningBalance
FROM         dbo.Payment INNER JOIN
                     dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
WHERE     (EntityType_DPA_ IN (SELECT EntityType_DPA_ FROM Entity WHERE SystemMaintained = 1))
	AND (Entity_DPA_ IN (SELECT Entity_DPA_ FROM Entity WHERE SystemMaintained = 1))
UNION ALL

SELECT     Entity_DPA_, CONVERT(datetime, CONVERT(nvarchar(12), LotTDate)) AS TransDate, '' AS Ref, CONVERT(nvarchar(12),
                      LotTDate) AS Particulars, 0 AS Debit, SUM(Commission) AS Credit, 0 AS IsOpeningBalance
FROM         (SELECT     CAST(SUM(dbo.Lot.LotQty) AS nvarchar) + ' ' + dbo.Security.SecurityName + ' @ ' + CAST(SUM(dbo.Lot.LotPrice) AS nvarchar)
                                              AS SecurityName, dbo.Lot.ContractNumber, dbo.Lot.LotTDate,
                                                  (SELECT     Entity_DPA_
                                                    FROM          Entity
                                                    WHERE      LevySystemMaintained = dbo.LevyContract.SystemMaintained) AS Entity_DPA_,
                                                  (SELECT     SUM(LevyAmount) AS dd
                                                    FROM          dbo.LevyContract AS LContracts
                                                    WHERE      (dbo.Lot.Contract_DPA_ = LContracts.Contract_DPA_) AND
                                                                           (LContracts.SystemMaintained = dbo.LevyContract.SystemMaintained)) AS Commission
                       FROM          dbo.Lot INNER JOIN
                                              dbo.OrdDetail ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                                              dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                                              dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                                              dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                                              dbo.LevyContract ON dbo.Lot.Contract_DPA_ = dbo.LevyContract.Contract_DPA_
                       WHERE     (dbo.LevyContract.SystemMaintained <> 12) AND (dbo.LevyContract.SystemMaintained <> 8)
                       GROUP BY dbo.Lot.LotTDate, dbo.Lot.ContractNumber, dbo.Lot.Contract_DPA_, dbo.Security.SecurityName, dbo.LevyContract.SystemMaintained)
                      ClientsStatement
GROUP BY Entity_DPA_, CONVERT(datetime, CONVERT(nvarchar(12), LotTDate)), CONVERT(nvarchar(12), LotTDate)

UNION ALL

SELECT     dbo.JournalList.Entity_DPA_, dbo.JournalList.JournalDate AS TransDate, CAST(dbo.JournalList.JournalEntry_DPA_ AS VARCHAR(500)) AS Ref,
                     dbo.JournalList.JournalNarrative AS Particulars, JournalList.JournalEntryDebit AS Debit,
                     JournalList.JournalEntryCredit AS Credit, 0 AS IsOpeningBalance
FROM         dbo.JournalList
WHERE      (EntityType_DPA_ IN (SELECT EntityType_DPA_ FROM Entity WHERE SystemMaintained = 1))
	AND (Entity_DPA_ IN (SELECT Entity_DPA_ FROM Entity WHERE SystemMaintained = 1))) LevyTransactionHistory

ORDER BY CONVERT(INT, Entity_DPA_), IsOpeningBalance DESC, TransDate

GO
CREATE VIEW dbo.LevyTransactionList
AS


(SELECT     Entity_DPA_, EntityRegDate AS TransDate, '' AS REF, ' Opening Balance' AS Particulars, case when EntityOpeningBal < 0 then EntityOpeningBal else 0 end AS Debit, case when EntityOpeningBal >= 0 then EntityOpeningBal else 0 end AS Credit,
                                              1 AS IsOpeningBalance,case when EntityOpeningBal < 0 then 0 - EntityOpeningBal else EntityOpeningBal end AS Balance
                       FROM          dbo.Entity WHERE SystemMaintained = 1)


union all

(SELECT     dbo.Payment.Entity_DPA_ AS Entity_DPA_, dbo.Payment.PaymentPDate AS TransDate, CAST(dbo.Payment.PaymentReference AS VARCHAR(500))
                      AS Ref, dbo.Payment.PaymentNarrative AS Particulars, CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                      CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit, 0 AS IsOpeningBalance,
                      CASE PayType.PayTypeIn WHEN 0 THEN 0 - Payment.PaymentAmount ELSE Payment.PaymentAmount END AS Balance
FROM         dbo.Payment INNER JOIN
                      dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_ INNER JOIN
                      dbo.Entity ON dbo.Entity.Entity_DPA_ = dbo.Payment.Entity_DPA_ AND dbo.Entity.EntityType_DPA_ = dbo.Payment.EntityType_DPA_
WHERE     (dbo.Entity.SystemMaintained = 1))

UNION ALL
SELECT LevyTransactions.*, Credit - Debit AS Balance FROM (
(SELECT     dbo.Entity.Entity_DPA_, dbo.Lot.LotTDate AS TransDate, dbo.Lot.ContractNumber AS Ref, CAST(SUM(dbo.Lot.LotQty) AS nvarchar)
                      + ' ' + dbo.Security.SecurityName + ' @ ' + CAST(SUM(dbo.Lot.LotPrice) AS nvarchar) AS Particulars, 0 AS Debit,
                          (SELECT     SUM(LevyAmount) AS dd
                            FROM          dbo.LevyContract AS LContracts
                            WHERE      (dbo.Lot.Contract_DPA_ = LContracts.Contract_DPA_) AND (LContracts.SystemMaintained = dbo.LevyContract.SystemMaintained))
                      AS Credit, 0 AS IsOpeningBalance
FROM         dbo.Lot INNER JOIN
                      dbo.OrdDetail ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.LevyContract ON dbo.Lot.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.Entity ON dbo.Entity.LevySystemMaintained = dbo.LevyContract.SystemMaintained
WHERE     (dbo.LevyContract.SystemMaintained <> 12) AND  (dbo.LevyContract.SystemMaintained <> 8)
GROUP BY dbo.Lot.LotTDate, dbo.Lot.ContractNumber, dbo.Lot.Contract_DPA_, dbo.Security.SecurityName, dbo.LevyContract.SystemMaintained,
                      dbo.Entity.Entity_DPA_)) LevyTransactions

UNION ALL

(SELECT     dbo.JournalList.Entity_DPA_, dbo.JournalList.JournalDate AS TransDate, CAST(dbo.JournalList.JournalEntry_DPA_ AS VARCHAR(500)) AS Ref,
                     dbo.JournalList.JournalNarrative AS Particulars, JournalList.JournalEntryDebit AS Debit,
                     JournalList.JournalEntryCredit AS Credit, 0 AS IsOpeningBalance,JournalList.JournalEntryCredit - JournalList.JournalEntryDebit AS Balance
FROM         dbo.JournalList INNER JOIN
                      dbo.Entity ON dbo.Entity.Entity_DPA_ = dbo.JournalList.Entity_DPA_ AND dbo.Entity.EntityType_DPA_ = dbo.JournalList.EntityType_DPA_
WHERE     (dbo.Entity.SystemMaintained = 1))










GO
CREATE VIEW dbo.live_ManualOrderReleaseList
AS
SELECT     dbo.tbOrder.Order_DPA_, dbo.OrdDetail.OrdDetail_DPA_, dbo.tbOrder.OrderDate, dbo.tbOrder.OrderHold, dbo.Security.SecurityName,
                      dbo.Client.ClientName, dbo.OrdDetail.OrdDetailPrice, dbo.OrdDetail.OrdDetailQty, dbo.OrdDetail.OrdDetailCertNo,
                      dbo.OrderType.OrderTypeDescription, dbo.tbOrder.OrderCanceled, dbo.tbOrder.OrderAutoReleaseDate
FROM         dbo.tbOrder INNER JOIN
                      dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_
WHERE     (dbo.tbOrder.OrderHold = 1) AND (dbo.tbOrder.Deleted <> 1) AND (dbo.OrdDetail.Deleted <> 1) AND (dbo.Client.IsFrozen <> 1) AND
                      (dbo.Client.Deleted <> 1)

GO
CREATE VIEW dbo.LotList
AS
SELECT     TOP 100 PERCENT dbo.OrdDetailList.Order_DPA_, dbo.OrdDetailList.OrdDetail_DPA_, dbo.OrdDetailList.OrdDetailType,
                      dbo.OrdDetailList.OrdDetailClient, dbo.OrdDetailList.OrdDetailSecurity, dbo.OrdDetailList.OrdDetailQty, dbo.OrdDetailList.OrdDetailPrice,
                      dbo.OrdDetailList.BalanceQty, dbo.Lots.LotSlipNo, dbo.Lots.LotTDate AS LotTDate, dbo.Lots.LotQty, dbo.Lots.LotPrice, dbo.Lots.LotGrossAmount,
                      dbo.Broker.BrokerCode, dbo.Broker.BrokerName, dbo.Contracts.Contract_DPA_, dbo.Status.StatusDescription, CASE ISNULL(dbo.Lots.Lot_DPA_, - 1)
                      WHEN - 1 THEN 0 ELSE dbo.Lots.Lot_DPA_ END AS Lot_DPA_, dbo.tbOrder.OrderDate, dbo.Lots.ContractNumber, dbo.OrdDetailList.OrdDetailSecType,
                      dbo.OrdDetailList.OrderTypeSale, dbo.OrdDetailList.Security_DPA_, dbo.OrdDetailList.CommissionRate, dbo.OrdDetailList.Client_DPA_,
                      dbo.OrdDetailList.AgentCommission, dbo.OrdDetailList.StaffCommission, dbo.OrdDetailList.SecurityCode, dbo.OrdDetailList.VolumeBoundary,
                      dbo.OrdDetailList.VolumeRate, dbo.OrdDetailList.MinimumCommission, dbo.OrdDetailList.CMARegulated, dbo.OrdDetailList.PostImmobilisedRate,
                      dbo.OrdDetailList.SecurityImmobilised, dbo.OrdDetailList.BondDescription, dbo.Lots.CDSTransaction, dbo.Contracts.ContractSettlementDate,
                      dbo.OrdDetailList.ClientCDSNo, dbo.OrdDetailList.InterBank, dbo.UserList.UserName AS ChangedBy, dbo.Lots.TimeChanged,
                      dbo.OrdDetailList.Best
FROM         dbo.Status RIGHT OUTER JOIN
                      dbo.Contracts RIGHT OUTER JOIN
                      dbo.Broker RIGHT OUTER JOIN
                      dbo.Lots RIGHT OUTER JOIN
                      dbo.OrdDetailList ON dbo.Lots.OrdDetail_DPA_ = dbo.OrdDetailList.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetailList.Order_DPA_ = dbo.tbOrder.Order_DPA_ ON dbo.Broker.Broker_DPA_ = dbo.Lots.Broker_DPA_ ON
                      dbo.Contracts.Contract_DPA_ = dbo.Lots.Contract_DPA_ ON dbo.Status.Status_DPA_ = dbo.Contracts.Status_DPA_ LEFT OUTER JOIN
                      dbo.UserList ON dbo.Lots.ChangedBy = dbo.UserList.UserID
WHERE     (dbo.OrdDetailList.OrderHold = 0)
ORDER BY dbo.OrdDetailList.OrdDetail_DPA_ DESC

GO

CREATE VIEW dbo.LotList AS
SELECT     TOP 100 PERCENT dbo.OrdDetailList.Order_DPA_, dbo.OrdDetailList.OrdDetail_DPA_, dbo.OrdDetailList.OrdDetailType,
                      dbo.OrdDetailList.OrdDetailClient, dbo.OrdDetailList.OrdDetailSecurity, dbo.OrdDetailList.OrdDetailQty, dbo.OrdDetailList.OrdDetailPrice,
                      dbo.OrdDetailList.BalanceQty, dbo.Lots.LotSlipNo, CONVERT(SMALLDATETIME, CAST(dbo.Lots.LotTDate AS CHAR(12))) AS LotTDate, dbo.Lots.LotQty,
                      dbo.Lots.LotPrice, dbo.Lots.LotGrossAmount, dbo.Broker.BrokerCode, dbo.Broker.BrokerName, dbo.Contracts.Contract_DPA_,
                      dbo.Status.StatusDescription, Lot_DPA_ = CASE ISNULL(dbo.Lots.Lot_DPA_, - 1) WHEN - 1 THEN 0 ELSE dbo.Lots.Lot_DPA_ END,
                      dbo.tbOrder.OrderDate, dbo.Lots.ContractNumber, dbo.OrdDetailList.OrdDetailSecType, dbo.OrdDetailList.OrderTypeSale,
                      dbo.OrdDetailList.Security_DPA_, dbo.OrdDetailList.CommissionRate, dbo.OrdDetailList.Client_DPA_, dbo.OrdDetailList.AgentCommission,
                      dbo.OrdDetailList.StaffCommission, dbo.OrdDetailList.SecurityCode, dbo.OrdDetailList.VolumeBoundary, dbo.OrdDetailList.VolumeRate,
                      dbo.OrdDetailList.MinimumCommission, dbo.OrdDetailList.CMARegulated, dbo.OrdDetailList.PostImmobilisedRate,
                      dbo.OrdDetailList.SecurityImmobilised, dbo.OrdDetailList.BondDescription, dbo.Lots.CDSTransaction, Lots.ContractSettlementDate,
                      OrdDetailList.ClientCDSNo, OrdDetailList.InterBank, UserList.UserName AS ChangedBy, Lots.timeChanged,OrdDetaillist.Best
FROM         dbo.Status RIGHT OUTER JOIN
                      dbo.Contracts RIGHT OUTER JOIN
                      dbo.Broker RIGHT OUTER JOIN
                      dbo.Lots RIGHT OUTER JOIN
                      dbo.OrdDetailList ON dbo.Lots.OrdDetail_DPA_ = dbo.OrdDetailList.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetailList.Order_DPA_ = dbo.tbOrder.Order_DPA_ ON dbo.Broker.Broker_DPA_ = dbo.Lots.Broker_DPA_ ON
                      dbo.Contracts.Contract_DPA_ = dbo.Lots.Contract_DPA_ ON dbo.Status.Status_DPA_ = dbo.Contracts.Status_DPA_ LEFT OUTER JOIN
                      UserList ON dbo.lots.Changedby = Userlist.UserId
WHERE     (dbo.OrdDetailList.OrderHold = 0)
ORDER BY dbo.OrdDetailList.OrdDetail_DPA_ DESC



GO


CREATE VIEW dbo.LotListBond
AS
SELECT     TOP 100 PERCENT dbo.OrdDetailList.Order_DPA_, dbo.OrdDetailList.OrdDetail_DPA_, dbo.OrdDetailList.OrdDetailType,
                      dbo.OrdDetailList.OrdDetailClient, dbo.OrdDetailList.OrdDetailSecurity, dbo.OrdDetailList.OrdDetailQty, dbo.OrdDetailList.OrdDetailPrice,
                      dbo.OrdDetailList.BalanceQty, dbo.Lots.LotSlipNo, CONVERT(SMALLDATETIME, CAST(dbo.Lots.LotTDate AS CHAR(12))) AS LotTDate, dbo.Lots.LotQty,
                      dbo.Lots.LotPrice, dbo.Lots.LotGrossAmount, dbo.Broker.BrokerCode, dbo.Broker.BrokerName, dbo.Contracts.Contract_DPA_,
                      dbo.Status.StatusDescription, Lot_DPA_ = CASE ISNULL(dbo.Lots.Lot_DPA_, - 1) WHEN - 1 THEN 0 ELSE dbo.Lots.Lot_DPA_ END,
                      dbo.tbOrder.OrderDate, dbo.Lots.ContractNumber, dbo.OrdDetailList.OrdDetailSecType, dbo.OrdDetailList.OrderTypeSale,
                      dbo.OrdDetailList.Security_DPA_, dbo.OrdDetailList.CommissionRate, dbo.OrdDetailList.Client_DPA_, dbo.OrdDetailList.AgentCommission,
                      dbo.OrdDetailList.StaffCommission, dbo.OrdDetailList.SecurityCode, dbo.OrdDetailList.VolumeBoundary, dbo.OrdDetailList.VolumeRate,
                      dbo.OrdDetailList.MinimumCommission, dbo.OrdDetailList.CMARegulated, dbo.OrdDetailList.PostImmobilisedRate,
                      dbo.OrdDetailList.SecurityImmobilised, dbo.OrdDetailList.BondDescription, dbo.Lots.CDSTransaction, Contracts.ContractSettlementDate,
                      OrdDetailList.ClientCDSNo, OrdDetailList.InterBank, UserList.UserName AS ChangedBy, Lots.timeChanged, OrdDetaillist.Best
FROM         dbo.Status RIGHT OUTER JOIN
                      dbo.Contracts RIGHT OUTER JOIN
                      dbo.Broker RIGHT OUTER JOIN
                      dbo.Lots RIGHT OUTER JOIN
                      dbo.OrdDetailList ON dbo.Lots.OrdDetail_DPA_ = dbo.OrdDetailList.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetailList.Order_DPA_ = dbo.tbOrder.Order_DPA_ ON dbo.Broker.Broker_DPA_ = dbo.Lots.Broker_DPA_ ON
                      dbo.Contracts.Contract_DPA_ = dbo.Lots.Contract_DPA_ ON dbo.Status.Status_DPA_ = dbo.Contracts.Status_DPA_ LEFT OUTER JOIN
                      UserList ON dbo.lots.Changedby = Userlist.UserId
WHERE     (dbo.OrdDetailList.OrderHold = 0 AND NOT (tborder.Proposal_DPA_ IS NULL))
ORDER BY dbo.OrdDetailList.OrdDetail_DPA_ DESC


GO
CREATE VIEW dbo.LotPayment
AS
SELECT     dbo.Lot.ContractNumber, dbo.Lot.LotGrossAmount, Payment_1.PaymentAmount AS PaymentAmount, Payment_1.PaymentPDate AS PaymentDate,
                      Payment_1.Payment_DPA_ AS Payment_DPA_, Payment_1.Voucher_DPA_ AS PaymentVoucher_DPA_, Payment_2.PaymentAmount AS ReceiptAmount,
                      Payment_2.PaymentPDate AS ReceiptDate, Payment_2.Payment_DPA_ AS Receipt_DPA_,
                      Payment_2.BrokerReceiptVoucher_DPA_ AS ReceiptVoucher_DPA_, dbo.Lot.Lot_DPA_, dbo.Contract.Contract_DPA_
FROM         dbo.Payment Payment_1 RIGHT OUTER JOIN
                      dbo.Voucher ON Payment_1.Voucher_DPA_ = dbo.Voucher.Voucher_DPA_ RIGHT OUTER JOIN
                      dbo.Lot INNER JOIN
                      dbo.Contract ON dbo.Lot.Contract_DPA_ = dbo.Contract.Contract_DPA_ LEFT OUTER JOIN
                      dbo.Payment Payment_2 RIGHT OUTER JOIN
                      dbo.BrokerReceiptVoucher ON Payment_2.BrokerReceiptVoucher_DPA_ = dbo.BrokerReceiptVoucher.BrokerReceiptVoucher_DPA_ ON
                      dbo.Contract.BrokerReceiptVoucher_DPA_ = dbo.BrokerReceiptVoucher.BrokerReceiptVoucher_DPA_ ON
                      dbo.Voucher.Voucher_DPA_ = dbo.Contract.Voucher_DPA_
WHERE     (Payment_1.Voucher_DPA_ IS NOT NULL) OR
                      (Payment_2.BrokerReceiptVoucher_DPA_ IS NOT NULL)

GO

CREATE   VIEW app.Lots AS
SELECT  Lot_DPA_                AS LotId,
        Contract_DPA_           AS ContractId,
        OrdDetail_DPA_          AS OrderDetailId,
        Broker_DPA_             AS BrokerId,
        ContractNumber          AS ContractNumber,
        LotPrice                AS Price,
        LotQty                  AS Quantity,
        LotGrossAmount          AS GrossAmount,
        LotTDate                AS TradeDate,
        ContractSettlementDate  AS SettlementDate
FROM dbo.Lot
WHERE Deleted = 0 OR Deleted IS NULL;

GO
CREATE VIEW dbo.Lots
AS
SELECT     Lot_DPA_, Contract_DPA_, OrdDetail_DPA_, LotPrice, LotQty, UniqueSlip, LotSlipNo, LotTDate, Broker_DPA_, ContractNumber, LotGrossAmount,
                      CDSImport_DPA_, CDSTransaction, ChangedBy, ContractSettlementDate, TimeChanged, Deleted
FROM         dbo.Lot
WHERE     (Deleted = 0)

GO
CREATE VIEW dbo.LotView
AS
SELECT     Lot_DPA_, Contract_DPA_, OrdDetail_DPA_, LotPrice, LotQty, LotSlipNo, CONVERT(SMALLDATETIME, CAST(LotTDate AS CHAR(12))) AS LotTDate,
                      Broker_DPA_, ContractNumber, LotGrossAmount, Deleted
FROM         dbo.Lot
WHERE     (Deleted = 0)

GO
CREATE VIEW dbo.MailConfigList
AS
SELECT     *
FROM         dbo.MailConfiguration

GO

CREATE VIEW dbo.MainMenuList
AS
SELECT     dbo.Menus.MenuID, dbo.Menus.mnuCaption, dbo.Menus.mnuDescription, dbo.Menus.mnuAction, dbo.Menus.Submenu, dbo.Menus.ismainMenu,
                      dbo.Menus.MainMenuID, dbo.Menus.IsReport, dbo.Menus.Image, dbo.Menus.mnuType, ISNULL(Menus_1.MenuID, dbo.Menus.MenuID)
                      AS DefaultChildID, ISNULL(Menus_1.mnuAction, dbo.Menus.mnuAction) AS DefaultChildAction, ISNULL(Menus_1.mnuDescription,
                      dbo.Menus.mnuDescription) AS DefaultChildDescription
FROM         dbo.Menus LEFT OUTER JOIN
                      dbo.Menus Menus_1 ON dbo.Menus.DefaultChildID = Menus_1.MenuID
WHERE     (dbo.Menus.IsReport <> 1) AND (dbo.Menus.ismainMenu = 1)


GO

CREATE VIEW dbo.ManualReleaseOrderList
AS
SELECT     OrderDate, OrderTypeName, OrderRef, ClientName, OrderHold, Order_DPA_, Client_DPA_, OrderHoldType_DPA_, ISNULL(CONVERT(nvarchar,
                      OrderAutoReleaseDate), '') AS OrderAutoReleaseDate, ClientCDSNo, SecurityName, OrdDetailQty, Quantity, CurrentBal, BalanceFree, OrdDetail_DPA_,
                      SecurityType, BondIssue
FROM         dbo.OrderListPlain
WHERE     (Order_DPA_ NOT IN
                          (SELECT     Order_DPA_
                            FROM          LotList))


GO

CREATE VIEW dbo.MarketSectorList
AS
SELECT     dbo.MarketSector.*
FROM         dbo.MarketSector


GO

CREATE VIEW dbo.MarketView
AS
SELECT DISTINCT TOP 100 PERCENT MAX(FLOOR(CAST(MktDate AS float))) AS MktUnique, LTRIM(MktCode) AS Code
FROM         dbo.datastream_Market
WHERE     (NOT (MktClose IS NULL)) AND (MktClose <> 0)
GROUP BY MktCode
ORDER BY LTRIM(MktCode)


GO
CREATE VIEW dbo.NominalTransactionList
AS

(SELECT     Account_DPA_, AccountOpeningDate AS TransDate, '' AS REF, ' Opening Balance' AS Particulars,
		case when AccountOpeningBal < 0 then AccountOpeningBal else 0 end AS Debit,
		case when AccountOpeningBal >= 0 then AccountOpeningBal else 0 end AS Credit,
                                              1 AS IsOpeningBalance,case when AccountOpeningBal < 0 then 0 - AccountOpeningBal else AccountOpeningBal end AS Balance
                       FROM          dbo.Account
		WHERE AccountTypeLevel1  <> 7)

union all

(SELECT     dbo.Payment.Entity_DPA_ AS Entity_DPA_, dbo.Payment.PaymentPDate AS TransDate, dbo.Payment.PaymentReference AS Ref,
                      dbo.Payment.PaymentNarrative AS Particulars, CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                      CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit, 0 AS IsOpeningBalance,
                      CASE PayType.PayTypeIn WHEN 0 THEN 0 - Payment.PaymentAmount ELSE Payment.PaymentAmount END AS Balance
FROM         dbo.Payment INNER JOIN
                      dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_ INNER JOIN
                      dbo.Account ON dbo.Account.Account_DPA_ = dbo.Payment.Entity_DPA_
WHERE     (dbo.Payment.EntityType_DPA_ = 5) AND (dbo.Account.AccountTypeLevel1 <> 7))

UNION ALL

(SELECT     dbo.JournalList.Entity_DPA_, dbo.JournalList.JournalDate AS TransDate, dbo.JournalList.JournalEntry_DPA_ AS Ref,
                     dbo.JournalList.JournalNarrative AS Particulars, JournalList.JournalEntryDebit AS Debit,
                     JournalList.JournalEntryCredit AS Credit, 0 AS IsOpeningBalance,JournalList.JournalEntryCredit - JournalList.JournalEntryDebit AS Balance
FROM         dbo.JournalList INNER JOIN
                      dbo.Account ON dbo.Account.Account_DPA_ = dbo.JournalList.Entity_DPA_
WHERE     (dbo.JournalList.EntityType_DPA_ = 5) AND (dbo.Account.AccountTypeLevel1 <> 7))
GO

CREATE VIEW dbo.NotMatched
AS
SELECT     TOP 100 PERCENT CDSImport_DPA_
FROM         dbo._CDS_Imported_Trades_
WHERE     (CDSImport_DPA_ NOT IN
                          (SELECT     CDSMatchedTradesList.CDSImport_DPA_
                            FROM          CDSMatchedTradesList))


GO
CREATE VIEW dbo.OfferingForm
AS
SELECT     dbo.Client.Client_DPA_, dbo.Security.SecurityCode, dbo.Security.SecurityName, dbo.Client.ClientName, dbo.Client.ClientCDSNo,
                      dbo.Offerings.Offering_DPA_, dbo.Offerings.Offering_Price, dbo.Offerings.Alloted_Rights, dbo.Agent.Agent_DPA_, dbo.Agent.AgentName,
                      dbo.Client.ClientCellTel, dbo.Client.ClientFax, dbo.Client.ClientOfficeTel, dbo.Client.ClientAddr
FROM         dbo.Offerings INNER JOIN
                      dbo.Client ON dbo.Offerings.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Security ON dbo.Offerings.Offering = dbo.Security.Security_DPA_ LEFT OUTER JOIN
                      dbo.Agent ON dbo.Client.Agent_DPA_ = dbo.Agent.Agent_DPA_

GO

CREATE VIEW dbo.OfferingList
AS
SELECT     TOP 100 PERCENT dbo.Offerings.PAL_No, dbo.Offerings.Client_DPA_, dbo.Offerings.Offering_DPA_, dbo.Offerings.Offering,
                      dbo.Offerings.Offering_Price, dbo.Offerings.ID_No, dbo.Offerings.Alloted_Rights, dbo.Offerings.Accepted_Rights, dbo.Offerings.Renouncee,
                      dbo.Offerings.Receipt, dbo.Offerings.Submitted, dbo.Offerings.Submission_Date, dbo.Offerings.Ref_No, dbo.Client.ClientName,
                      dbo.Offerings.Alloted_Rights * dbo.Offerings.Offering_Price AS Amount_Payable, dbo.Security.SecurityName, dbo.Security.SecurityCode,
                      dbo.Offerings.Batch_No, dbo.Offerings.TimeChanged, dbo.Users.Surname + ' ' + dbo.Users.OtherNames AS ModifiedBy, dbo.Offerings.Offerings_Date,
                      dbo.Offerings.Paid, dbo.Offerings.AppType, dbo.Offerings.shNAME2, dbo.Offerings.BatchSeq, dbo.Client.ClientIDPass, dbo.Offerings.Deleted,
                      dbo.Offerings.OfferCheque, dbo.Offerings.OfferBank, dbo.Offerings.BankCode, dbo.Offerings.AccountNo, dbo.Offerings.ChangedBy,
                      dbo.Offerings.Offering_EIT_, dbo.Offerings.Payment_DPA_, dbo.Offerings.Download_DPA_, dbo.Offerings.Downloaded, dbo.Offerings.BatchFileName,
                      dbo.Offerings.LastDownLoaded, dbo.Offerings.Additional, dbo.Offerings.AcceptanceType, dbo.Offerings.CreatedBy, dbo.Offerings.DateCreated,
                      dbo.Offerings.Download_Date, dbo.Offerings.Branch_DPA_, dbo.Offerings.Agent_DPA_, dbo.Offerings.PaymentMode, dbo.Offerings.BatchClosed,
                      dbo.Offerings.BatchPaymentMode, dbo.Offerings.ClientCellTel, dbo.Offerings.RefundMethod, dbo.Offerings.CDSNumeric, dbo.Offerings.TaxExempt,
                      dbo.Offerings.DividendMethod, dbo.Offerings.EFTBankRef, dbo.Offerings.EFTBranchRef, dbo.Offerings.EFTSortCode, dbo.Offerings.EFTAccountNo,
                      dbo.Offerings.Category, dbo.Offerings.PaymentType, dbo.Offerings.PaymentRef, dbo.Offerings.PaymentBankRef, dbo.Offerings.PaymentBranchRef,
                      dbo.Offerings.PaymentSortCode, dbo.Offerings.PaymentAccountNo, dbo.Offerings.OfferingType, dbo.Offerings.CDACode, dbo.Offerings.AgentCode,
                      dbo.Offerings.CDSPaid, dbo.Offerings.Certificate, dbo.Offerings.ReceivingBroker, dbo.Offerings.Status
FROM         dbo.Offerings INNER JOIN
                      dbo.Client ON dbo.Offerings.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Security ON dbo.Offerings.Offering = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.Users ON dbo.Offerings.ChangedBy = dbo.Users.UserID
WHERE     (dbo.Offerings.Deleted = 0) AND (dbo.Offerings.Forward <> 1)
ORDER BY dbo.Offerings.BatchSeq


GO

CREATE VIEW dbo.OfferingListFilter
AS
SELECT     TOP 100 PERCENT dbo.SecurityOffering.SecurityCode + SPACE
                          ((SELECT     ISNULL(MAX(LEN(SecurityCode)), 0) AS MAXLEN
                              FROM         dbo.SecurityOffering) - LEN(dbo.SecurityOffering.SecurityCode)) + ' : ' + dbo.SecurityOffering.SecurityName AS SecurityNameEx,
                      dbo.SecurityOffering.SecurityName, dbo.SecurityOffering.SecurityCode, dbo.SecurityOffering.SecurityMktPrice, dbo.SecurityOffering.SecurityAddr,
                      '0' AS OrderSecTypeDescription, MAX(dbo.SecTransFee.SecTransFeeADate) AS MaxOfSecTransFeeADate, dbo.SecurityOffering.Security_DPA_,
                      'Offering' AS OrderSecTypeDisplayName, dbo.SecurityOffering.OrderSecType_DPA_, dbo.SecurityOffering.Immobilised,
                      dbo.SecurityOffering.CanTrade
FROM         dbo.SecurityOffering INNER JOIN
                      dbo.SecTransFee ON dbo.SecurityOffering.Security_DPA_ = dbo.SecTransFee.Security_DPA_
GROUP BY dbo.SecurityOffering.SecurityName, dbo.SecurityOffering.SecurityCode, dbo.SecurityOffering.SecurityMktPrice, dbo.SecurityOffering.SecurityAddr,
                      dbo.SecurityOffering.Security_DPA_, dbo.SecurityOffering.OrderSecType_DPA_, dbo.SecurityOffering.Immobilised,
                      dbo.SecurityOffering.CanTrade
ORDER BY dbo.SecurityOffering.SecurityName


GO

CREATE VIEW dbo.OfferingsList
AS
SELECT     TOP 100 PERCENT dbo.Offerings.PAL_No, dbo.Offerings.Client_DPA_, dbo.Offerings.Offering_DPA_, dbo.Offerings.Offering,
                      dbo.Offerings.Offering_Price, dbo.Offerings.ID_No, dbo.Offerings.Alloted_Rights, dbo.Offerings.Accepted_Rights, dbo.Offerings.Renouncee,
                      dbo.Offerings.Receipt, dbo.Offerings.Submitted, dbo.Offerings.Submission_Date, dbo.Offerings.Ref_No, dbo.Client.ClientName,
                      dbo.Offerings.Alloted_Rights * dbo.Offerings.Offering_Price AS Amount_Payable, dbo.Security.SecurityName, dbo.Security.SecurityCode,
                      dbo.Offerings.Batch_No, dbo.Offerings.TimeChanged, dbo.Users.Surname + ' ' + dbo.Users.OtherNames AS ModifiedBy, dbo.Offerings.Offerings_Date,
                      dbo.Offerings.Paid, dbo.Offerings.AppType, dbo.Offerings.shNAME2, dbo.Offerings.BatchSeq, ISNULL(dbo.Security.BatchSize, 0) AS BatchSize,
                      CAST(FLOOR(CAST(dbo.Security.ClosingDate AS Float)) AS datetime) AS ClosingDate, dbo.Offerings.Forward, dbo.Broker.BrokerName,
                      dbo.Offerings.Status, dbo.Offerings.ReceivingBroker, dbo.Offerings.PaymentBankRef, dbo.Offerings.PaymentBranchRef,
                      dbo.Offerings.PaymentSortCode, dbo.Offerings.PaymentAccountNo, dbo.Offerings.PaymentRef, dbo.Offerings.PaymentType, dbo.Offerings.DateCreated,
                      Users_1.Surname + ' ' + Users_1.OtherNames AS CreatedBy
FROM         dbo.Offerings INNER JOIN
                      dbo.Client ON dbo.Offerings.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Security ON dbo.Offerings.Offering = dbo.Security.Security_DPA_ LEFT OUTER JOIN
                      dbo.Users Users_1 ON dbo.Offerings.CreatedBy = Users_1.UserID LEFT OUTER JOIN
                      dbo.Broker ON dbo.Offerings.ReceivingBroker = dbo.Broker.Broker_DPA_ LEFT OUTER JOIN
                      dbo.Users ON dbo.Offerings.ChangedBy = dbo.Users.UserID
WHERE     (dbo.Offerings.Deleted = 0) AND (dbo.Offerings.Forward = 0)
ORDER BY dbo.Offerings.Offering_DPA_ DESC


GO
CREATE VIEW dbo.OfferTypeList
AS
SELECT     OfferType_DPA_, Description
FROM         dbo.OfferType
WHERE     (Deleted = 0)

GO

CREATE VIEW dbo.OnlineClients
AS
SELECT     TOP 100 PERCENT CAST(dbo.OnlineClient.Client_DPA_ AS NVARCHAR(4000)) + SPACE
                          ((SELECT     ISNULL(MAX(LEN(CAST(dbo.OnlineClient.Client_DPA_ AS NVARCHAR(4000)))), 0) AS MAXLEN
                              FROM         dbo.OnlineClient) - LEN(CAST(dbo.OnlineClient.Client_DPA_ AS NVARCHAR(4000)))) + ' : ' + dbo.OnlineClient.ClientName AS ClientNameEx,
                       LTRIM(RTRIM(dbo.OnlineClient.ClientName)) AS ClientName, dbo.OwnerList.OwnerName, dbo.OnlineClient.ClientContact,
                      dbo.OnlineClient.ClientOfficeTel, dbo.OnlineClient.ClientCellTel, dbo.OnlineClient.ClientEmail, dbo.OnlineClient.ClientAddr, dbo.OnlineClient.ClientVIP,
                      dbo.OnlineClient.ClientCDSNo, dbo.OnlineClient.OnlineRegistration, ISNULL(dbo.OnlineClient.CreditLimit, 0) AS CreditLimit, dbo.OnlineClient.ClientFax,
                      ISNULL(dbo.OnlineClient.Owner_DPA_, 0) AS Owner_DPA_, ISNULL(dbo.Agent.Agent_DPA_, 0) AS Agent_DPA_, dbo.Agent.AgentName,
                      dbo.OnlineClient.TimeChanged, dbo.UserList.[USER] AS ChangedBy, dbo.ClientStatementBalances.Balance, dbo.OnlineClient.IsCustodian,
                      dbo.OnlineClient.Commission_DPA_, dbo.OnlineClient.ClientIDPass, dbo.OnlineClient.Branch_DPA_, dbo.OnlineClient.Class_DPA_,
                      dbo.OnlineClient.Client_DPA_, dbo.OnlineClient.Client_EIT_, dbo.OnlineClient.ClientBDate, dbo.OnlineClient.ClientHomeTel,
                      dbo.OnlineClient.ClientPhoto, dbo.OnlineClient.ClientSignature, dbo.OnlineClient.Gender_DPA_, dbo.OnlineClient.Residency_DPA_,
                      dbo.OnlineClient.ClientPAddr, dbo.OnlineClient.ClientComment, dbo.OnlineClient.EntityType_DPA_, ISNULL(dbo.OnlineClient.ClientOpeningBal, 0)
                      AS ClientOpeningBal, dbo.OnlineClient.ClientRegDate, dbo.OnlineClient.GenericSetting_DPA_, dbo.OnlineClient.GenericSetting_DPA_2,
                      dbo.OnlineClient.GenericSetting_DPA_3, dbo.OnlineClient.OldAccountNo, '' AS BankAccName, 0 AS BnkBranch_DPA_, 0 AS BankAccNumber,
                      0 AS BankAcc_DPA_
FROM         dbo.OnlineClient LEFT OUTER JOIN
                      dbo.ClientStatementBalances ON dbo.OnlineClient.Client_DPA_ = dbo.ClientStatementBalances.Client_DPA_ LEFT OUTER JOIN
                      dbo.UserList ON dbo.OnlineClient.ChangedBy = dbo.UserList.UserID LEFT OUTER JOIN
                      dbo.Agent ON dbo.OnlineClient.Agent_DPA_ = dbo.Agent.Agent_DPA_ LEFT OUTER JOIN
                      dbo.OwnerList ON dbo.OnlineClient.Owner_DPA_ = dbo.OwnerList.Owner_DPA_
WHERE     (dbo.OnlineClient.Deleted = 0)
ORDER BY dbo.OnlineClient.ClientName


GO
CREATE VIEW dbo.OnlineClients2
AS
SELECT     TOP 100 PERCENT CAST(dbo.OnlineClient.Client_DPA_ AS NVARCHAR(4000)) + SPACE
                          ((SELECT     ISNULL(MAX(LEN(CAST(dbo.OnlineClient.Client_DPA_ AS NVARCHAR(4000)))), 0) AS MAXLEN
                              FROM         dbo.OnlineClient) - LEN(CAST(dbo.OnlineClient.Client_DPA_ AS NVARCHAR(4000)))) + ' : ' + dbo.OnlineClient.ClientName AS ClientNameEx,
                       LTRIM(RTRIM(dbo.OnlineClient.ClientName)) AS ClientName, dbo.OwnerList.OwnerName, dbo.OnlineClient.ClientContact,
                      dbo.OnlineClient.ClientOfficeTel, dbo.OnlineClient.ClientCellTel, dbo.OnlineClient.ClientEmail, dbo.OnlineClient.ClientAddr, dbo.OnlineClient.ClientVIP,
                      dbo.OnlineClient.ClientCDSNo, dbo.OnlineClient.OnlineRegistration, ISNULL(dbo.OnlineClient.CreditLimit, 0) AS CreditLimit, dbo.OnlineClient.ClientFax,
                      ISNULL(dbo.OnlineClient.Owner_DPA_, 0) AS Owner_DPA_, ISNULL(dbo.Agent.Agent_DPA_, 0) AS Agent_DPA_, dbo.Agent.AgentName,
                      dbo.OnlineClient.TimeChanged, dbo.UserList.[USER] AS ChangedBy, dbo.OnlineClient.IsCustodian, dbo.OnlineClient.Commission_DPA_,
                      dbo.OnlineClient.ClientIDPass, dbo.OnlineClient.Branch_DPA_, dbo.OnlineClient.Class_DPA_, dbo.OnlineClient.Client_DPA_,
                      dbo.OnlineClient.Client_EIT_, dbo.OnlineClient.ClientBDate, dbo.OnlineClient.ClientHomeTel, dbo.OnlineClient.ClientPhoto,
                      dbo.OnlineClient.ClientSignature, dbo.OnlineClient.Gender_DPA_, dbo.OnlineClient.Residency_DPA_, dbo.OnlineClient.ClientPAddr,
                      dbo.OnlineClient.ClientComment, dbo.OnlineClient.EntityType_DPA_, ISNULL(dbo.OnlineClient.ClientOpeningBal, 0) AS ClientOpeningBal,
                      dbo.OnlineClient.ClientRegDate, dbo.OnlineClient.GenericSetting_DPA_, dbo.OnlineClient.GenericSetting_DPA_2,
                      dbo.OnlineClient.GenericSetting_DPA_3, dbo.OnlineClient.OldAccountNo, '' AS BankAccName, 0 AS BnkBranch_DPA_, 0 AS BankAccNumber,
                      0 AS BankAcc_DPA_
FROM         dbo.OnlineClient LEFT OUTER JOIN
                      dbo.UserList ON dbo.OnlineClient.ChangedBy = dbo.UserList.UserID LEFT OUTER JOIN
                      dbo.Agent ON dbo.OnlineClient.Agent_DPA_ = dbo.Agent.Agent_DPA_ LEFT OUTER JOIN
                      dbo.OwnerList ON dbo.OnlineClient.Owner_DPA_ = dbo.OwnerList.Owner_DPA_
WHERE     (dbo.OnlineClient.Deleted = 0)
ORDER BY dbo.OnlineClient.ClientName

GO

CREATE VIEW dbo.OnlineUserList
AS
SELECT     dbo.Users.UserID, dbo.Users.OtherNames + '  ' + dbo.Users.Surname AS [USER], dbo.Users.UserName, dbo.Users.Description,
                      dbo.Users.SecretQuestion, dbo.Users.SecretAnswer, dbo.Users.RequiresSecretQuestion, dbo.Users.Client_DPA_, dbo.Client.ClientName,
                      dbo.Users.Password, dbo.Users.Surname, dbo.Users.OtherNames, dbo.Users.StaffID, dbo.Users.Removed, dbo.Users.FirstTime, dbo.Users.Expires,
                      dbo.Users.Enabled, dbo.Users.RemoteUser, dbo.Client.ClientEmail, dbo.Users.Accepted, dbo.Users.Email
FROM         dbo.Users INNER JOIN
                      dbo.Client ON dbo.Users.Client_DPA_ = dbo.Client.Client_DPA_
WHERE     (dbo.Users.RemoteUser = 1) AND (dbo.Users.Accepted = 0)



GO

CREATE VIEW dbo.OpeningBalances
AS
Select Top 100 Percent * from
(
SELECT    Top 100 Percent  LTRIM(RTRIM(Client.ClientName)) AS EntityName, EntityType.EntityTypeName AS EntityType, Client.Client_DPA_ AS Entity_DPA_,
                      Client.EntityType_DPA_, CONVERT(char(10), Client.Client_DPA_) AS EntityCode, Client.ClientOpeningBal as OpeningBalance, Cast(Floor(Cast(Client.ClientRegDate AS Float))
                       AS DateTime) AS RegistrationDate, 'Client' AS TableName
FROM         Client INNER JOIN
                      EntityType ON Client.EntityType_DPA_ = EntityType.EntityType_DPA_
WHERE     (Client.Deleted = 0)
UNION ALL
SELECT   Top 100 Percent  LTRIM(RTRIM(EntityName)) AS EntityName, EntityType, Entity_DPA_, EntityType_DPA_, EntityCode, EntityOpeningBal, RegistrationDate,
                      'Entity' AS TABLEName
FROM         EntityList
UNION ALL
SELECT   Top 100 Percent  dbo.Broker.BrokerName AS EntityName, dbo.EntityType.EntityTypeName AS EntityType, dbo.Broker.Broker_DPA_ AS Entity_DPA_,
                      dbo.EntityType.EntityType_DPA_, dbo.Broker.BrokerCode AS EntityCode, dbo.Broker.BrokerOpeningBal, dbo.Broker.BrokerRegDate,
                      'Broker' AS TABLEName
FROM         dbo.Broker INNER JOIN
                      dbo.EntityType ON dbo.Broker.EntityType_DPA_ = dbo.EntityType.EntityType_DPA_
UNION ALL
SELECT  Top 100 Percent   dbo.Agent.AgentName AS EntityName, dbo.EntityType.EntityTypeName AS EntityType, dbo.Agent.Agent_DPA_ AS Entity_DPA_,
                      dbo.EntityType.EntityType_DPA_, CAST(dbo.Agent.Agent_DPA_ AS NVARCHAR(4000)) AS EntityCode, dbo.Agent.AgentOpeningBal,
                      dbo.Agent.AgentRegDate, 'Agent' AS TableName
FROM         dbo.Agent INNER JOIN
                      dbo.EntityType ON dbo.Agent.EntityType_DPA_ = dbo.EntityType.EntityType_DPA_
WHERE     (dbo.Agent.AgentName <> N'_none_') AND (dbo.Agent.Deleted = 0)
UNION ALL
SELECT  Top 100 Percent   dbo.OwnerList.OwnerName AS EntityName, dbo.EntityType.EntityTypeName AS EntityType, dbo.OwnerList.Owner_DPA_ AS Entity_DPA_,
                      dbo.EntityType.EntityType_DPA_, CAST(dbo.OwnerList.Owner_DPA_ AS NVARCHAR(4000)) AS EntityCode, dbo.OwnerList.OwnerOpeningBal,
                      dbo.OwnerList.OwnerRegDate, 'Owner' AS TableName
FROM         dbo.OwnerList INNER JOIN
                      dbo.EntityType ON dbo.OwnerList.EntityType_DPA_ = dbo.EntityType.EntityType_DPA_
UNION ALL
SELECT  Top 100 Percent   dbo.Account.AccountName AS EntityName, dbo.EntityType.EntityTypeName AS EntityType, dbo.Account.Account_DPA_ AS Entity_DPA_,
                      dbo.EntityType.EntityType_DPA_, dbo.Account.AccountCode AS EntityCode, dbo.Account.AccountOpeningBal,
                      CAST(FLOOR(CAST(dbo.Account.AccountOpeningDate AS Float)) AS DateTime) AS RegistrationDate, 'Account' AS TableName
FROM         dbo.Account INNER JOIN
                      dbo.EntityType ON dbo.Account.EntityType_DPA_ = dbo.EntityType.EntityType_DPA_

) a
order by a.EntityType,a.EntityName








GO

CREATE VIEW dbo.OrdDetailContractedQtyList
AS
SELECT     dbo.OrdDetail.OrdDetail_DPA_, SUM(dbo.Lot.LotQty) AS ContractQty
FROM         dbo.OrdDetail INNER JOIN
                      dbo.Lot ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.Lot.OrdDetail_DPA_
WHERE     (dbo.Lot.Deleted = 0) AND (dbo.OrdDetail.Deleted = 0)
GROUP BY dbo.OrdDetail.OrdDetail_DPA_


GO
/*
SELECT     dbo.OrdDetail.Amount, dbo.Client.ClientName AS OrdDetailClient, dbo.Security.SecurityCode + ' : ' + CONVERT(NVARCHAR(1000),
                      dbo.OrdDetail.OrdDetailQty) AS OrdDetailItem, dbo.OrdDetail.OrdDetailPrice, dbo.OrdDetail.OrdDetailQty,
                      dbo.OrderSecType.OrderSecTypeDisplayName AS OrdDetailSecType, CASE ISNULL(dbo.OrdDetailContractedQtyList.ContractQty, 1)
                      WHEN 1 THEN dbo.OrdDetail.OrdDetailQty ELSE dbo.OrdDetail.OrdDetailQty - dbo.OrdDetailContractedQtyList.ContractQty END AS BalanceQty,
                      CASE dbo.OrderSecType.OrderSecType_DPA_ WHEN 1 THEN dbo.Commission.BondCommission ELSE dbo.Commission.CommissionRate END AS CommissionRate,
                       CASE dbo.OrderSecType.OrderSecType_DPA_ WHEN 1 THEN dbo.Commission.BondBoundary ELSE dbo.Commission.SecurityBoundary END AS VolumeBoundary,
                       CASE dbo.OrderSecType.OrderSecType_DPA_ WHEN 1 THEN dbo.Commission.UpperBondCommission ELSE dbo.Commission.UpperSecurityCommission
                       END AS VolumeRate,
                      CASE dbo.OrderSecType.OrderSecType_DPA_ WHEN 1 THEN dbo.Commission.MinimumBondCommission ELSE dbo.Commission.MinimumSecurityCommission
                       END AS MinimumCommission, dbo.OrderSecType.OrderSecType_DPA_, dbo.Security.SecurityName + ' ' + ISNULL(dbo.Bond.BondIssue, '')
                      AS OrdDetailSecurity, dbo.OrderType.OrderTypeDescription AS OrdDetailType, dbo.OrdDetail.OrdDetail_DPA_, dbo.tbOrder.Order_DPA_,
                      dbo.tbOrder.OrderDate, dbo.tbOrder.OrderCanceled, dbo.tbOrder.OrderHold, dbo.tbOrder.OrderRef, dbo.Security.Security_DPA_,
                      dbo.OrderType.OrderTypeSale, CASE dbo.OrderType.OrderTypeSale WHEN 0 THEN 'B' ELSE 'S' END AS CDSOrderTypeSale,
                      dbo.tbOrder.OrderCompounded AS OrdDetailCompound, CONVERT(DATETIME, dbo.OrdDetail.OrdDetailValidity, 108) AS OrdDetailValidity,
                      dbo.OrdDetail.OrdDetailCertNo, dbo.Security.SecurityName, dbo.Client.Client_DPA_, dbo.Client.ClientCDSNo, ISNULL(dbo.AgentList.AgentCommission,
                      0) AS AgentCommission, ISNULL(dbo.OwnerList.OwnerName, 'No Manager') AS AccountManager, ISNULL(dbo.OwnerList.CommissionRate, 0)
                      AS StaffCommission, RTRIM(LTRIM(dbo.Security.SecurityCode + ' ' + ISNULL(dbo.Bond.BondIssue, ''))) AS SecurityCode,
                      dbo.Commission.CMARegulated, dbo.Commission.Immobilised AS PostImmobilisedRate, dbo.Security.Immobilised AS SecurityImmobilised,
                      dbo.OrdDetail.Best, dbo.tbOrder.InterBank, ISNULL(dbo.Bond.BondIssue, '') AS BondDescription, dbo.Client.IsCustodian, dbo.Client.EntityType_DPA_,
                      dbo.Client.Class_DPA_ AS Class, dbo.OrdDetail.Limit, dbo.AgentList.AgentName
FROM         dbo.tbOrder INNER JOIN
                      dbo.OrderSecType ON dbo.tbOrder.OrderSecType_DPA_ = dbo.OrderSecType.OrderSecType_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ LEFT OUTER JOIN
                      dbo.Bond ON dbo.OrdDetail.Bond_DPA_ = dbo.Bond.Bond_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.Commission ON dbo.Client.Commission_DPA_ = dbo.Commission.Commission_DPA_ LEFT OUTER JOIN
                      dbo.OwnerList ON dbo.Client.Owner_DPA_ = dbo.OwnerList.Owner_DPA_ LEFT OUTER JOIN
                      dbo.AgentList ON dbo.Client.Agent_DPA_ = dbo.AgentList.Agent_DPA_ LEFT OUTER JOIN
                      dbo.OrdDetailContractedQtyList ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.OrdDetailContractedQtyList.OrdDetail_DPA_
WHERE     (dbo.tbOrder.OrderCanceled = 0) AND (dbo.tbOrder.Deleted = 0) AND (dbo.OrdDetail.Deleted = 0) AND (dbo.Client.Deleted = 0)
*/
CREATE VIEW dbo.OrdDetailList
AS
SELECT     dbo.OrdDetail.Amount, dbo.Client.ClientName AS OrdDetailClient, dbo.Security.SecurityCode + ' : ' + CONVERT(NVARCHAR(1000),
                      dbo.OrdDetail.OrdDetailQty) AS OrdDetailItem, dbo.OrdDetail.OrdDetailPrice, dbo.OrdDetail.OrdDetailQty,
                      dbo.OrderSecType.OrderSecTypeDisplayName AS OrdDetailSecType, CASE ISNULL(dbo.OrdDetailContractedQtyList.ContractQty, 1)
                      WHEN 1 THEN dbo.OrdDetail.OrdDetailQty ELSE dbo.OrdDetail.OrdDetailQty - dbo.OrdDetailContractedQtyList.ContractQty END AS BalanceQty,
                      CASE dbo.OrderSecType.OrderSecType_DPA_ WHEN 1 THEN dbo.Commission.BondCommission ELSE dbo.Commission.CommissionRate END AS CommissionRate,
                       CASE dbo.OrderSecType.OrderSecType_DPA_ WHEN 1 THEN dbo.Commission.BondBoundary ELSE dbo.Commission.SecurityBoundary END AS VolumeBoundary,
                       CASE dbo.OrderSecType.OrderSecType_DPA_ WHEN 1 THEN dbo.Commission.UpperBondCommission ELSE dbo.Commission.UpperSecurityCommission
                       END AS VolumeRate,
                      CASE dbo.OrderSecType.OrderSecType_DPA_ WHEN 1 THEN dbo.Commission.MinimumBondCommission ELSE dbo.Commission.MinimumSecurityCommission
                       END AS MinimumCommission, dbo.OrderSecType.OrderSecType_DPA_, dbo.Security.SecurityName + ' ' + ISNULL(dbo.Bond.BondIssue, '')
                      AS OrdDetailSecurity, dbo.OrderType.OrderTypeDescription AS OrdDetailType, dbo.OrdDetail.OrdDetail_DPA_, dbo.tbOrder.Order_DPA_,
                      dbo.tbOrder.OrderDate, dbo.tbOrder.OrderCanceled, dbo.tbOrder.OrderHold, dbo.tbOrder.OrderRef, dbo.Security.Security_DPA_,
                      dbo.OrderType.OrderTypeSale, CASE dbo.OrderType.OrderTypeSale WHEN 0 THEN 'B' ELSE 'S' END AS CDSOrderTypeSale,
                      dbo.tbOrder.OrderCompounded AS OrdDetailCompound, CONVERT(DATETIME, dbo.OrdDetail.OrdDetailValidity, 108) AS OrdDetailValidity,
                      dbo.OrdDetail.OrdDetailCertNo, dbo.Security.SecurityName, dbo.Client.Client_DPA_, dbo.Client.ClientCDSNo, ISNULL(dbo.AgentList.AgentCommission,
                      0) AS AgentCommission, ISNULL(dbo.OwnerList.OwnerName, 'No Manager') AS AccountManager, ISNULL(dbo.OwnerList.CommissionRate, 0)
                      AS StaffCommission, RTRIM(LTRIM(dbo.Security.SecurityCode + ' ' + ISNULL(dbo.Bond.BondIssue, ''))) AS SecurityCode,
                      dbo.Commission.CMARegulated, dbo.Commission.Immobilised AS PostImmobilisedRate, dbo.Security.Immobilised AS SecurityImmobilised,
                      dbo.OrdDetail.Best, dbo.tbOrder.InterBank, ISNULL(dbo.Bond.BondIssue, '') AS BondDescription, dbo.Client.IsCustodian, dbo.Client.EntityType_DPA_,
                      dbo.Client.Class_DPA_ AS Class, dbo.OrdDetail.Limit, dbo.AgentList.AgentName, dbo.Client.Agent_DPA_
FROM         dbo.tbOrder INNER JOIN
                      dbo.OrderSecType ON dbo.tbOrder.OrderSecType_DPA_ = dbo.OrderSecType.OrderSecType_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ LEFT OUTER JOIN
                      dbo.Bond ON dbo.OrdDetail.Bond_DPA_ = dbo.Bond.Bond_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.Commission ON dbo.Client.Commission_DPA_ = dbo.Commission.Commission_DPA_ LEFT OUTER JOIN
                      dbo.OwnerList ON dbo.Client.Owner_DPA_ = dbo.OwnerList.Owner_DPA_ LEFT OUTER JOIN
                      dbo.AgentList ON dbo.Client.Agent_DPA_ = dbo.AgentList.Agent_DPA_ LEFT OUTER JOIN
                      dbo.OrdDetailContractedQtyList ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.OrdDetailContractedQtyList.OrdDetail_DPA_
WHERE     (dbo.tbOrder.OrderCanceled = 0) AND (dbo.tbOrder.Deleted = 0) AND (dbo.OrdDetail.Deleted = 0) AND (dbo.Client.Deleted = 0)

GO
CREATE VIEW dbo.Order_Book
AS
SELECT     TOP (100) PERCENT Order_DPA_, Client_DPA_, ClientName, SecurityCode, OrdDetailQty, BalanceQty, OrderTypeName, OrdDetailPrice, OrderDate, OrdDetailValidity,
                      OrderCanceled, ChangedBy, OrderReleasedBy
FROM         dbo.FullOrderList
WHERE     (OrderHold = 0)

GO

CREATE VIEW dbo.OrderContractCompounded
AS
SELECT     TOP 100 PERCENT dbo.Contract.Contract_DPA_, dbo.LotList.Lot_DPA_, dbo.LotList.BrokerCode, dbo.LotList.OrdDetailType,
                      dbo.LotList.OrdDetailSecType, dbo.LotList.OrderTypeSale, dbo.LotList.LotSlipNo, dbo.LotList.LotTDate, dbo.LotList.LotQty, dbo.LotList.LotPrice,
                      dbo.LotList.ContractNumber, dbo.LotList.OrdDetailClient, dbo.LotList.OrdDetailSecurity, dbo.LevyContract.LevyAmount, dbo.LevyContract.LevyName,
                      dbo.tbOrder.OrderRef, dbo.Client.ClientAddr, dbo.Client.Client_DPA_, dbo.LotList.LotGrossAmount AS GrossAmount, dbo.tbOrder.Order_DPA_,
                      dbo.LevyContract.SystemMaintained, CASE WHEN ISNULL(CHARINDEX('%', dbo.LevyContract.LevyRatePercentage), 0)
                      > 0 THEN '%' ELSE '' END AS LevyRatePercentage, dbo.LevyContract.LevyShortName, LotList.Security_DPA_, LotList.OrdDetail_DPA_
FROM         dbo.Contract INNER JOIN
                      dbo.LotList ON dbo.Contract.Contract_DPA_ = dbo.LotList.Contract_DPA_ INNER JOIN
                      dbo.LevyContract ON dbo.Contract.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.LotList.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.LotList.Client_DPA_ = dbo.Client.Client_DPA_
WHERE     (dbo.tbOrder.OrderCompounded = 1) AND tbOrder.Deleted = 0
ORDER BY dbo.LotList.LotTDate, dbo.LotList.OrdDetailSecurity, dbo.LotList.LotSlipNo, dbo.tbOrder.Order_DPA_



GO

CREATE   VIEW app.OrderDetails AS
SELECT  OrdDetail_DPA_     AS OrderDetailId,
        Order_DPA_         AS OrderId,
        Security_DPA_      AS SecurityId,
        OrdDetailPrice     AS Price,
        OrdDetailQty       AS Quantity,
        Amount             AS Amount,
        Best               AS IsBest
FROM dbo.OrdDetail
WHERE Deleted = 0 OR Deleted IS NULL;

GO
CREATE VIEW dbo.OrderForm
AS
SELECT     dbo.tbOrder.Order_DPA_, dbo.tbOrder.OrderDate, dbo.Client.ClientName, dbo.Client.ClientCDSNo, dbo.Client.Client_DPA_, dbo.Client.ClientAddr,
                      dbo.Client.ClientOfficeTel, dbo.Agent.AgentName, dbo.Agent.Agent_DPA_, dbo.Security.SecurityName, dbo.OrdDetail.OrdDetailQty,
                      dbo.OrdDetail.OrdDetailPrice, dbo.OrdDetail.OrdDetailValidity, dbo.OrdDetail.OrdDetailCertNo, dbo.OrderType.OrderTypeDescription,
                      dbo.OrdDetail.Amount, dbo.Security.OrderSecType_DPA_, dbo.OrdDetail.BondDescription, dbo.Security.SecurityCode, dbo.OrdDetail.Best,
                      dbo.Client.ClientFax, dbo.Client.ClientCellTel
FROM         dbo.tbOrder INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ LEFT OUTER JOIN
                      dbo.Agent ON dbo.Client.Agent_DPA_ = dbo.Agent.Agent_DPA_
WHERE     (dbo.tbOrder.Deleted = 0) AND (dbo.Client.Deleted = 0) AND (dbo.OrdDetail.Deleted = 0)

GO
CREATE VIEW dbo.OrderHoldTypeList
AS
SELECT     TOP 100 PERCENT OrderHoldType_DPA_, OrderHoldTypeName, DefaultSelection
FROM         dbo.OrderHoldType
WHERE     (OrderHoldType_DPA_ <> 1)
ORDER BY OrderHoldTypeName

GO
CREATE VIEW dbo.OrderList
AS
SELECT     TOP 100 PERCENT dbo.FullOrderList.*
FROM         dbo.FullOrderList
WHERE     (OrderHold = 0)

GO

CREATE VIEW dbo.OrderListPlain
AS
SELECT DISTINCT
                      TOP 100 PERCENT CONVERT(DATETIME, dbo.tbOrder.OrderDate, 108) AS OrderDate, dbo.OrderTypeList.OrderTypeName, dbo.tbOrder.OrderRef,
                      dbo.tbOrder.OrderHold, dbo.tbOrder.Order_DPA_, dbo.tbOrder.Client_DPA_, dbo.tbOrder.OrderHoldType_DPA_, dbo.tbOrder.OrderAutoReleaseDate,
                      dbo.Client.ClientName, dbo.Client.ClientCDSNo, dbo.Security.SecurityCode AS SecurityName, dbo.OrdDetail.OrdDetailQty, dbo.Holdings.Quantity,
                      dbo.ClientBalances.CurrentBal, dbo.Holdings.BalanceFree, dbo.tbOrder.OrderDateReleased, dbo.OrdDetail.OrdDetail_DPA_,
                      dbo.OrderSecType.OrderSecTypeDisplayName AS SecurityType, dbo.Bond.BondIssue
FROM         dbo.tbOrder INNER JOIN
                      dbo.OrderTypeList ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderTypeList.OrderType_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.OrderSecType ON dbo.tbOrder.OrderSecType_DPA_ = dbo.OrderSecType.OrderSecType_DPA_ LEFT OUTER JOIN
                      dbo.Bond ON dbo.OrdDetail.Bond_DPA_ = dbo.Bond.Bond_DPA_ LEFT OUTER JOIN
                      dbo.ClientBalances ON dbo.Client.Client_DPA_ = dbo.ClientBalances.client_DPA_ LEFT OUTER JOIN
                      dbo.Holdings ON dbo.Security.Security_DPA_ = dbo.Holdings.Security_DPA_ AND dbo.Client.Client_DPA_ = dbo.Holdings.Client_DPA_
WHERE     (dbo.tbOrder.OrderCanceled = 0) AND (dbo.Client.Deleted = 0) AND (dbo.tbOrder.Deleted = 0)
ORDER BY dbo.tbOrder.Order_DPA_ DESC


GO

CREATE   VIEW app.Orders AS
SELECT  Order_DPA_         AS OrderId,
        Client_DPA_        AS ClientId,
        Branch_DPA_        AS BranchId,
        Agent_DPA_         AS AgentId,
        OrderType_DPA_     AS OrderTypeId,
        OrderSecType_DPA_  AS SecTypeId,
        OrderHoldType_DPA_ AS HoldTypeId,
        OrderDate          AS OrderDate,
        OrderRef           AS Reference,
        OrderHold          AS IsOnHold,
        OrderCanceled      AS IsCancelled,
        OrderCompounded    AS IsCompounded,
        Remarks            AS Remarks,
        TimeCreated        AS CreatedOn
FROM dbo.tbOrder
WHERE Deleted = 0 OR Deleted IS NULL;

GO
CREATE VIEW dbo.OrderSecTypeList
AS
SELECT     OrderSecTypeDescription, OrderSecType_DPA_, OrderSecTypeDisplayName, DefaultSelection
FROM         dbo.OrderSecType

GO
CREATE VIEW dbo.OrderTypeList
AS
SELECT     OrderTypeDescription AS OrderTypeName, OrderType_DPA_, OrderTypeSale, DefaultSelection, RequireCertificate, HandlingFee
FROM         dbo.OrderType

GO
CREATE VIEW dbo.OutstandingOrderList
AS
SELECT     Order_DPA_ AS [Order No], OrderDate AS [Order Date], OrdDetailClient + '	[' + CONVERT(nvarchar(4000), Client_DPA_) + ']' AS Client,
                      OrdDetailSecurity AS Security, OrdDetailQty AS [Order Qty], OrdDetailPrice AS [Order Price], LotSlipNo AS [Slip No], LotTDate AS [Trade Date],
                      LotQty AS Qty, LotPrice AS Price, BrokerCode AS Broker, ContractNumber AS Contract
FROM         dbo.LotList
WHERE     (Order_DPA_ IN
                          (SELECT DISTINCT Order_DPA_
                            FROM          dbo.LotList
                            WHERE      (BalanceQty > 0)))



GO
CREATE VIEW dbo.OwnerCommissionList
AS

SELECT CAST(LotQty AS nvarchar) + ' ' + SecurityName + ' @ ' + CAST(LotPrice AS nvarchar) AS SecurityName, ContractNumber
	, LotTDate, Owner_DPA_, LotGrossAmount As Gross, LevyAmount As OwnerCommission

FROM
(SELECT     dbo.Lot.ContractNumber, dbo.Lot.LotTDate, dbo.OwnerList.Owner_DPA_, dbo.Lot.LotGrossAmount, SUM(dbo.LevyContract.LevyAmount)
                      AS LevyAmount, dbo.Lot.LotQty, dbo.Lot.LotPrice, dbo.Security.SecurityName
FROM         dbo.Lot INNER JOIN
                      dbo.LevyContract ON dbo.Lot.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.OwnerList ON dbo.Client.Owner_DPA_ = dbo.OwnerList.Owner_DPA_
WHERE     (dbo.LevyContract.SystemMaintained = 8)
GROUP BY dbo.Lot.ContractNumber, dbo.Lot.LotTDate, dbo.OwnerList.Owner_DPA_, dbo.Lot.LotGrossAmount, dbo.Lot.LotQty, dbo.Lot.LotPrice,
                      dbo.Security.SecurityName) innerTable


GO

CREATE VIEW dbo.OwnerCommissions
AS
SELECT DISTINCT
                      TOP 100 PERCENT dbo.Lots.ContractNumber, cast(Floor(cast(dbo.Lots.LotTDate AS Float)) AS DateTime) AS LotTDate, dbo.tbOrder.Client_DPA_,
                      dbo.OrderType.OrderTypeSale, dbo.LevyContract.LevyAmount, CAST(dbo.Lots.LotQty AS nvarchar)
                      + ' ' + dbo.Security.SecurityCode + ' @ ' + CAST(dbo.Lots.LotPrice AS nvarchar) AS SecurityName, dbo.Payments.PaymentReceiptNo,
                      dbo.LevyContract.Contract_DPA_, dbo.OwnerList.OwnerName, CASE WHEN ((dbo.OwnerList.Owner_DPA_ IS NULL) AND
                      (dbo.Client.Agent_DPA_ IS NULL) AND (Security.OrderSecType_DPA_ = 2) AND Client.IsCustodian = 0)
                      THEN 0 WHEN ((dbo.OwnerList.Owner_DPA_ IS NULL) AND (dbo.Client.Agent_DPA_ IS NULL) AND (Security.OrderSecType_DPA_ = 2) AND
                      Client.IsCustodian = 1) THEN - 1 ELSE dbo.OwnerList.Owner_DPA_ END AS Owner_DPA_, dbo.Client.ClientName, dbo.Lots.LotGrossAmount,
                      dbo.Client.Agent_DPA_, Client.IsCustodian
FROM         dbo.Lots INNER JOIN
                      dbo.LevyContract ON dbo.Lots.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.Lots.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ LEFT OUTER JOIN
                      dbo.OwnerList ON dbo.Client.Owner_DPA_ = dbo.OwnerList.Owner_DPA_ LEFT OUTER JOIN
                      dbo.Payments ON Lots.Contract_DPA_ = dbo.Payments.Contract_DPA_
WHERE     (dbo.tbOrder.OrderCanceled = 0) AND (dbo.LevyContract.SystemMaintained <> 8) AND (dbo.LevyContract.SystemMaintained <> 12) AND
                      (dbo.LevyContract.LevyShortName LIKE N'%Commission%') AND (dbo.Client.Deleted = 0) AND (dbo.OrdDetail.Deleted = 0) AND
                      (dbo.tbOrder.Deleted = 0) AND (dbo.LevyContract.Deleted = 0)
ORDER BY dbo.OwnerList.Owner_DPA_, Client.IsCustodian, cast(Floor(cast(dbo.Lots.LotTDate AS Float)) AS DateTime)


GO
CREATE VIEW dbo.OwnerEntityList
AS
SELECT     dbo.OwnerList.OwnerName AS EntityName, dbo.EntityType.EntityTypeName AS EntityType, dbo.OwnerList.Owner_DPA_ AS Entity_DPA_,
                      dbo.EntityType.EntityType_DPA_, CAST(dbo.OwnerList.Owner_DPA_ AS NVARCHAR(4000)) AS EntityCode, CAST(dbo.OwnerList.Owner_DPA_ AS NVARCHAR(4000))
                      + SPACE
                          ((SELECT     ISNULL(MAX(LEN(CAST(dbo.OwnerList.Owner_DPA_ AS NVARCHAR(4000)))), 0) AS MAXLEN
                              FROM         dbo.OwnerList) - LEN(CAST(dbo.OwnerList.Owner_DPA_ AS NVARCHAR(4000)))) + ' : ' + dbo.OwnerList.OwnerName AS EntityNameEx
FROM         dbo.OwnerList INNER JOIN
                      dbo.EntityType ON dbo.OwnerList.EntityType_DPA_ = dbo.EntityType.EntityType_DPA_




GO
CREATE VIEW dbo.OwnerList
AS
SELECT     TOP 100 PERCENT OwnerFname + ' ' + OwnerLName AS OwnerName, Owner_DPA_, CommissionRate, EntityType_DPA_, OwnerRegDate,
                      OwnerOpeningBal, IdNo, MobileNo, SendSMS
FROM         dbo.Owner
ORDER BY OwnerFname + ' ' + OwnerLName

GO
CREATE VIEW dbo.OwnerReport
AS

SELECT     TOP 100 PERCENT OwnerName AS [Account Manager], '' AS [Office Phone], '' AS [Fax], REPLACE(REPLACE('', CHAR(13), ','), CHAR(10), '') AS Address,
                       CONVERT(NVARCHAR(500),Owner_DPA_) AS [Code]
FROM         dbo.OwnerList
ORDER BY OwnerName




GO

CREATE VIEW dbo.OwnerStatement
AS

SELECT     TOP 100 PERCENT COUNT(*) AS ClientTransaction_DPA_, a.Owner_DPA_, Cast(floor(cast(a.TransDate AS float)) AS DateTime) AS TransDate, a.Ref,
                      a.Particulars, a.Debit, a.Credit, CASE WHEN (SUM(b.Balance) >= 0) THEN CONVERT(NVARCHAR, SUM(b.Balance)) + ' Cr' ELSE CONVERT(NVARCHAR,
                      ABS(SUM(b.Balance))) + ' Dr' END AS Balance, a.IsOpeningBalance
FROM         (SELECT     *
                       FROM          dbo.OwnerTransactionList) a CROSS JOIN
                          (SELECT     *
                            FROM          dbo.OwnerTransactionList) b
WHERE     a.TransDate >= b.TransDate AND a.Owner_DPA_ = b.Owner_DPA_
GROUP BY a.Owner_DPA_, Cast(floor(cast(a.TransDate AS float)) AS DateTime), a.Ref, a.Particulars, a.Debit, a.Credit, a.Balance, a.IsOpeningBalance
ORDER BY Cast(floor(cast(a.TransDate AS float)) AS DateTime), a.Owner_DPA_, a.Particulars, a.Ref, a.Debit, a.Credit, a.Balance, a.IsOpeningBalance




GO

CREATE VIEW dbo.OwnerTransactionList
AS
SELECT     OwnerTransactions.*, Credit - Debit AS Balance
	FROM         (SELECT     Owner_DPA_, OwnerRegDate AS TransDate, '' AS REF, ' Opening Balance' AS Particulars,
                              CASE WHEN OwnerOpeningBal < 0 THEN (0 - OwnerOpeningBal) ELSE 0 END AS Debit,
                              CASE WHEN OwnerOpeningBal >= 0 THEN OwnerOpeningBal ELSE 0 END AS Credit, 1 AS IsOpeningBalance
       FROM          dbo.Owner
       UNION ALL
       SELECT     dbo.Payment.Entity_DPA_ AS Owner_DPA_, dbo.Payment.PaymentPDate AS TransDate, dbo.Payment.PaymentReference AS Ref,
                             dbo.Payment.PaymentNarrative AS Particulars, CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.Payment INNER JOIN
                             dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
       WHERE     EntityType_DPA_ = 7
       UNION ALL
       SELECT     Owner_DPA_, LotTDate AS TransDate, ContractNumber AS Ref, SecurityName AS Particulars, 0 AS Debit, isnull(OwnerCommission, 0)
                             AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.OwnerCommissionList
       UNION ALL
	 SELECT     dbo.JournalList.Entity_DPA_ AS Owner_DPA_, dbo.JournalList.JournalDate AS TransDate, CAST(dbo.JournalList.JournalEntry_DPA_ AS NVARCHAR(500)) AS Ref,
                             dbo.JournalList.JournalNarrative AS Particulars, JournalList.JournalEntryDebit AS Debit,
                             JournalList.JournalEntryCredit AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.JournalList
       WHERE     EntityType_DPA_ = 7) OwnerTransactions





GO

CREATE VIEW dbo.PaymentClientStatementNos
AS
SELECT     Payment_DPA_, CONVERT(char(5), PaymentReceiptNo) AS ReceiptNo, 1 AS type
FROM         dbo.Payment
WHERE     (NOT (PaymentReceiptNo IS NULL)) AND (EntityType_DPA_ <> 8) AND Deleted = 0
UNION
SELECT     dbo.Payment.Payment_DPA_, dbo.LotList.ContractNumber, 2 AS Type
FROM         dbo.Payment INNER JOIN
                      dbo.Voucher ON dbo.Payment.Voucher_DPA_ = dbo.Voucher.Voucher_DPA_ INNER JOIN
                      dbo.Contract ON dbo.Voucher.Voucher_DPA_ = dbo.Contract.Voucher_DPA_ INNER JOIN
                      dbo.LotList ON dbo.Contract.Contract_DPA_ = dbo.LotList.Contract_DPA_
WHERE     lotlist.interbank <> 1 AND Payment.Deleted = 0
UNION
SELECT     dbo.Payment.Payment_DPA_, dbo.LotList.ContractNumber, 2 AS Type
FROM         dbo.BrokerReceiptVoucher INNER JOIN
                      dbo.Payment ON dbo.BrokerReceiptVoucher.BrokerReceiptVoucher_DPA_ = dbo.Payment.BrokerReceiptVoucher_DPA_ INNER JOIN
                      dbo.LotList INNER JOIN
                      dbo.Contract ON dbo.LotList.Contract_DPA_ = dbo.Contract.Contract_DPA_ ON
                      dbo.BrokerReceiptVoucher.BrokerReceiptVoucher_DPA_ = dbo.Contract.BrokerReceiptVoucher_DPA_
WHERE     lotlist.interbank <> 1 AND Payment.Deleted = 0 AND Contract.Deleted = 0
UNION
SELECT     dbo.Payment.Payment_DPA_, dbo.Payment.PaymentReference, 3 AS Type
FROM         dbo.Payment INNER JOIN
                      dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
WHERE     (dbo.Payment.PayType_DPA_ = 2) AND (dbo.Payment.Voucher_DPA_ IS NULL) AND Payment.Deleted = 0

GO
CREATE VIEW dbo.PaymentList
AS
SELECT     TOP 100 PERCENT dbo.Payments.Payment_DPA_, dbo.Payments.PaymentPDate, dbo.CompleteEntityList.EntityType,
                      dbo.CompleteEntityList.EntityName, dbo.BankAccountList.AccountName AS BankAccountName, CAST(dbo.Payments.PaymentAmount AS Float)
                      AS PaymentAmount, dbo.Payments.PaymentReference, dbo.Payments.TimeChanged, dbo.UserList.[USER] AS ChangedBy
FROM         dbo.Payments INNER JOIN
                      dbo.BankAccountList ON dbo.Payments.BankAccount_DPA_ = dbo.BankAccountList.Account_DPA_ INNER JOIN
                      dbo.PayType ON dbo.Payments.PayType_DPA_ = dbo.PayType.PayType_DPA_ INNER JOIN
                      dbo.CompleteEntityList ON dbo.Payments.EntityType_DPA_ = dbo.CompleteEntityList.EntityType_DPA_ AND
                      dbo.Payments.Entity_DPA_ = dbo.CompleteEntityList.Entity_DPA_ LEFT OUTER JOIN
                      dbo.UserList ON dbo.Payments.ChangedBy = dbo.UserList.UserID
WHERE     (dbo.PayType.PayTypeIn = 0) AND (dbo.Payments.Deleted = 0)
ORDER BY dbo.Payments.Payment_DPA_ DESC

GO
CREATE VIEW dbo.PaymentListWithoutCDSPrepaid
AS
SELECT     TOP 100 PERCENT dbo.Payment.Payment_DPA_, dbo.Payment.PaymentPDate, dbo.CompleteEntityList.EntityType,
                      dbo.CompleteEntityList.EntityName, dbo.BankAccountList.AccountName AS BankAccountName, CAST(dbo.Payment.PaymentAmount AS Float)
                      AS PaymentAmount, dbo.Payment.PaymentReference, dbo.Payment.TimeChanged, dbo.UserList.[USER] AS ChangedBy
FROM         dbo.Payment INNER JOIN
                      dbo.BankAccountList ON dbo.Payment.BankAccount_DPA_ = dbo.BankAccountList.Account_DPA_ INNER JOIN
                      dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_ INNER JOIN
                      dbo.CompleteEntityList ON dbo.Payment.EntityType_DPA_ = dbo.CompleteEntityList.EntityType_DPA_ AND
                      dbo.Payment.Entity_DPA_ = dbo.CompleteEntityList.Entity_DPA_ LEFT OUTER JOIN
                      dbo.UserList ON dbo.Payment.ChangedBy = dbo.UserList.UserID
WHERE     (dbo.PayType.PayTypeIn = 0) AND (dbo.CompleteEntityList.EntityType_DPA_ <> 8) AND (dbo.Payment.Deleted = 0)
ORDER BY dbo.Payment.Payment_DPA_ DESC

GO
CREATE VIEW dbo.PaymentRequestCombinedList AS SELECT TOP 100 PERCENT dbo.PaymentRequest.PaymentRequest_DPA_, dbo.PaymentRequest.PaymentPDate, dbo.BankAccountList.AccountName AS BankAccountName, CAST(dbo.PaymentRequest.PaymentAmount AS Float) AS PaymentAmount, dbo.PaymentRequest.PaymentReference, dbo.PaymentRequest.TimeChanged, dbo.UserList.[USER] AS ChangedBy, dbo.Client.Client_DPA_, dbo.Client.ClientName, dbo.PaymentRequest.AccountToUse, dbo.PaymentRequest.Status FROM dbo.PaymentRequest INNER JOIN dbo.BankAccountList ON dbo.PaymentRequest.BankAccount_DPA_ = dbo.BankAccountList.Account_DPA_ INNER JOIN dbo.PayType ON dbo.PaymentRequest.PayType_DPA_ = dbo.PayType.PayType_DPA_ INNER JOIN dbo.UserList ON dbo.PaymentRequest.ChangedBy = dbo.UserList.UserID INNER JOIN dbo.Client ON dbo.PaymentRequest.Entity_DPA_ = dbo.Client.Client_DPA_ WHERE (dbo.PayType.PayTypeIn = 0) ORDER BY dbo.PaymentRequest.PaymentRequest_DPA_ DESC
GO
CREATE VIEW dbo.PaymentRequestConfirmedList AS SELECT TOP 100 PERCENT dbo.PaymentRequest.PaymentRequest_DPA_, dbo.PaymentRequest.PaymentPDate, dbo.BankAccountList.AccountName AS BankAccountName, CAST(dbo.PaymentRequest.PaymentAmount AS Float) AS PaymentAmount, dbo.PaymentRequest.PaymentReference, dbo.PaymentRequest.TimeChanged, dbo.UserList.[USER] AS ChangedBy, dbo.Client.Client_DPA_, dbo.Client.ClientName, dbo.PaymentRequest.AccountToUse, dbo.PaymentRequest.Status FROM dbo.PaymentRequest INNER JOIN dbo.BankAccountList ON dbo.PaymentRequest.BankAccount_DPA_ = dbo.BankAccountList.Account_DPA_ INNER JOIN dbo.PayType ON dbo.PaymentRequest.PayType_DPA_ = dbo.PayType.PayType_DPA_ INNER JOIN dbo.UserList ON dbo.PaymentRequest.ChangedBy = dbo.UserList.UserID INNER JOIN dbo.Client ON dbo.PaymentRequest.Entity_DPA_ = dbo.Client.Client_DPA_ WHERE (dbo.PayType.PayTypeIn = 0) AND (dbo.PaymentRequest.Status <> N'Pending') ORDER BY dbo.PaymentRequest.PaymentRequest_DPA_ DESC
GO
CREATE VIEW dbo.PaymentRequestList AS SELECT TOP 100 PERCENT dbo.PaymentRequest.PaymentRequest_DPA_, dbo.PaymentRequest.PaymentPDate, dbo.BankAccountList.AccountName AS BankAccountName, CAST(dbo.PaymentRequest.PaymentAmount AS Float) AS PaymentAmount, dbo.PaymentRequest.PaymentReference, dbo.PaymentRequest.TimeChanged, dbo.UserList.[USER] AS ChangedBy, dbo.Client.Client_DPA_, dbo.Client.ClientName, dbo.PaymentRequest.AccountToUse, dbo.PaymentRequest.Status FROM dbo.PaymentRequest INNER JOIN dbo.BankAccountList ON dbo.PaymentRequest.BankAccount_DPA_ = dbo.BankAccountList.Account_DPA_ INNER JOIN dbo.PayType ON dbo.PaymentRequest.PayType_DPA_ = dbo.PayType.PayType_DPA_ INNER JOIN dbo.UserList ON dbo.PaymentRequest.ChangedBy = dbo.UserList.UserID INNER JOIN dbo.Client ON dbo.PaymentRequest.Entity_DPA_ = dbo.Client.Client_DPA_ WHERE (dbo.PayType.PayTypeIn = 0) AND (dbo.PaymentRequest.Status = N'Pending') ORDER BY dbo.PaymentRequest.PaymentRequest_DPA_ DESC
GO

CREATE   VIEW app.Payments AS
SELECT  Payment_DPA_       AS PaymentId,
        EntityType_DPA_    AS EntityTypeId,
        Entity_DPA_        AS EntityId,
        PayType_DPA_       AS PayTypeId,
        PaymentAmount      AS Amount,
        PaymentReceiptNo   AS ReceiptNo,
        PaymentPDate       AS PaymentDate,
        PaymentReference   AS Reference,
        PaymentNarrative   AS Narrative,
        BankAccount_DPA_   AS BankAccountId,
        Contract_DPA_      AS ContractId,
        Order_DPA_         AS OrderId
FROM dbo.Payment
WHERE Deleted = 0 OR Deleted IS NULL;

GO

CREATE VIEW dbo.Payments
AS
SELECT     dbo.Payment.*
FROM         dbo.Payment
WHERE     (Deleted = 0)


GO
CREATE VIEW dbo.PaymentSchedule
AS
SELECT     TOP 100 PERCENT dbo.CompleteEntityList.EntityType AS Entity, dbo.CompleteEntityList.EntityName AS Party,
                      dbo.BankAccountList.AccountName AS Bank_Account, dbo.Payment.PaymentReceiptNo AS Receipt_No, dbo.Payment.PaymentReference AS Reference,
                      CAST(dbo.Payment.PaymentPDate AS NVARCHAR(12)) AS Payment_Date, dbo.Payment.PaymentAmount AS Amount,
                      CASE (dbo.CompleteEntityList.EntityType_DPA_) WHEN 8 THEN 1 ELSE 0 END AS CdsGroup, dbo.BankAccountList.AccountCode
FROM         dbo.Payment INNER JOIN
                      dbo.BankAccountList ON dbo.Payment.BankAccount_DPA_ = dbo.BankAccountList.Account_DPA_ INNER JOIN
                      dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_ INNER JOIN
                      dbo.CompleteEntityList ON dbo.Payment.EntityType_DPA_ = dbo.CompleteEntityList.EntityType_DPA_ AND
                      dbo.Payment.Entity_DPA_ = dbo.CompleteEntityList.Entity_DPA_
WHERE     (dbo.PayType.PayTypeIn = 0) AND Payment.Deleted = 0
ORDER BY CdsGroup, Receipt_No

GO
CREATE VIEW dbo.PaymentsReport
AS
SELECT      TOP 100 PERCENT
	 dbo.CompleteEntityList.EntityType [Entity],
	dbo.CompleteEntityList.EntityName [Recipient],
                     dbo.BankAccountList.AccountName AS [Bank Account],
	dbo.Payment.PaymentReference [Reference],
		CAST(CAST(dbo.Payment.PaymentPDate AS NVARCHAR(12)) AS DATETIME) As [Payment Date],
				dbo.Payment.PaymentAmount [Amount]
FROM         dbo.Payment INNER JOIN
                      dbo.BankAccountList ON dbo.Payment.BankAccount_DPA_ = dbo.BankAccountList.Account_DPA_ INNER JOIN
                      dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_ INNER JOIN
                      dbo.CompleteEntityList ON dbo.Payment.EntityType_DPA_ = dbo.CompleteEntityList.EntityType_DPA_ AND
                      dbo.Payment.Entity_DPA_ = dbo.CompleteEntityList.Entity_DPA_
WHERE     (dbo.PayType.PayTypeIn = 0)
ORDER BY dbo.Payment.PaymentPDate

GO

CREATE VIEW dbo.PaymentTypesList
AS
SELECT     dbo.PaymentTypes.*
FROM         dbo.PaymentTypes


GO

CREATE VIEW dbo.PaymentVoucher
AS
SELECT     dbo.OrderType.OrderTypeDescription AS OrdDetailType, dbo.Security.SecurityName AS OrdDetailSecurity, dbo.Client.Client_DPA_,
                      dbo.Client.ClientName, dbo.Lot.LotSlipNo, dbo.Lot.LotTDate, dbo.Lot.ContractNumber, dbo.Lot.Contract_DPA_, dbo.Lot.LotTDate AS SettlementDate,
                      SUM(dbo.Lot.LotGrossAmount) -
                          (SELECT     SUM(LevyAmount) AS dd
                            FROM          LevyContract
                            WHERE      Contract_DPA_ = dbo.Lot.Contract_DPA_ AND SystemMaintained <> 12) AS NetAmount, dbo.Lot.LotGrossAmount, dbo.Lot.LotQty,
                      dbo.Lot.LotPrice, dbo.Security.SecurityCode
FROM         dbo.tbOrder INNER JOIN
                      dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.Lot ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.Lot.OrdDetail_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_
WHERE     (dbo.OrderType.OrderTypeSale = 1) AND (dbo.Client.Deleted = 0) AND (dbo.Lot.Deleted = 0) AND (dbo.tbOrder.Deleted = 0) AND
                      (dbo.OrdDetail.Deleted = 0) AND (dbo.Lot.Contract_DPA_ NOT IN
                          (SELECT     Contract_DPA_
                            FROM          intertransfer
                            WHERE      InterTransferType_DPA_ = 2))
GROUP BY dbo.OrderType.OrderTypeDescription, dbo.Security.SecurityName, dbo.Security.SecurityCode, dbo.Client.Client_DPA_, dbo.Client.ClientName,
                      dbo.Lot.LotSlipNo, dbo.Lot.LotTDate, dbo.Lot.ContractNumber, dbo.Lot.Contract_DPA_, dbo.Lot.LotGrossAmount, dbo.Lot.LotQty, dbo.Lot.LotPrice


GO

CREATE VIEW dbo.PaymentVoucherDates
AS
SELECT     dbo.Client.Client_DPA_, dbo.Lot.Contract_DPA_, dbo.Lot.LotTDate AS SettlementDate, dbo.OrderSecType.OrderSecType_DPA_,
                      dbo.OrderSecType.OrderSecTypeDisplayName, dbo.Class.ClassDescription, dbo.Class.Class_DPA_
FROM         dbo.tbOrder INNER JOIN
                      dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.Lot ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.Lot.OrdDetail_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.OrderSecType ON dbo.tbOrder.OrderSecType_DPA_ = dbo.OrderSecType.OrderSecType_DPA_ INNER JOIN
                      dbo.Class ON dbo.Client.Class_DPA_ = dbo.Class.Class_DPA_
WHERE     (dbo.OrderType.OrderTypeSale = 1) AND (dbo.Lot.Deleted = 0) AND (dbo.OrdDetail.Deleted = 0) AND (dbo.tbOrder.Deleted = 0)
GROUP BY dbo.OrderType.OrderTypeDescription, dbo.Security.SecurityName, dbo.Client.Client_DPA_, dbo.Client.ClientName, dbo.Lot.LotSlipNo,
                      dbo.Lot.LotTDate, dbo.Lot.ContractNumber, dbo.Lot.Contract_DPA_, dbo.Lot.LotGrossAmount, dbo.Lot.LotQty, dbo.Lot.LotPrice,
                      dbo.OrderSecType.OrderSecType_DPA_, dbo.OrderSecType.OrderSecTypeDisplayName, dbo.Class.ClassDescription, dbo.Class.Class_DPA_


GO

CREATE VIEW dbo.PaymentVoucherInExcel
AS
SELECT     dbo.OrderType.OrderTypeDescription AS OrdDetailType, dbo.Security.SecurityName AS OrdDetailSecurity, dbo.Client.Client_DPA_,
                      dbo.Client.ClientName, dbo.Lot.LotSlipNo, dbo.Lot.LotTDate, dbo.Lot.ContractNumber, dbo.Lot.Contract_DPA_, dbo.Lot.LotTDate AS SettlementDate,
                      SUM(dbo.Lot.LotGrossAmount) -
                          (SELECT     SUM(LevyAmount) AS dd
                            FROM          LevyContract
                            WHERE      Contract_DPA_ = dbo.Lot.Contract_DPA_ AND SystemMaintained <> 12) AS NetAmount, dbo.Lot.LotGrossAmount, dbo.Lot.LotQty,
                      dbo.Lot.LotPrice, dbo.Agent.Agent_DPA_ AS AgentCode, dbo.Agent.AgentName AS Agent, dbo.Security.SecurityCode
FROM         dbo.tbOrder INNER JOIN
                      dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.Lot ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.Lot.OrdDetail_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ LEFT OUTER JOIN
                      dbo.Agent ON dbo.Client.Agent_DPA_ = dbo.Agent.Agent_DPA_
WHERE     (dbo.OrderType.OrderTypeSale = 1) AND (dbo.Client.Deleted = 0) AND (dbo.Lot.Deleted = 0) AND (dbo.tbOrder.Deleted = 0) AND
                      (dbo.OrdDetail.Deleted = 0) AND (dbo.Lot.Contract_DPA_ NOT IN
                          (SELECT     Contract_DPA_
                            FROM          intertransfer
                            WHERE      InterTransferType_DPA_ = 2))
GROUP BY dbo.OrderType.OrderTypeDescription, dbo.Security.SecurityName, dbo.Client.Client_DPA_, dbo.Client.ClientName, dbo.Lot.LotSlipNo,
                      dbo.Lot.LotTDate, dbo.Lot.ContractNumber, dbo.Lot.Contract_DPA_, dbo.Lot.LotGrossAmount, dbo.Lot.LotQty, dbo.Lot.LotPrice, dbo.Agent.Agent_DPA_,
                      dbo.Agent.AgentName, dbo.Security.SecurityCode


GO
CREATE VIEW dbo.PayTypeList AS SELECT PayType.PayTypeDescription AS PayTypeDescription, PayType.PayType_DPA_ AS PayType_DPA_
FROM PayType

GO

CREATE VIEW dbo.PBs
AS
SELECT TOP 100 PERCENT dbo.PortfolioQuantities.*, dbo.CurrentBalances.CurrentBal, dbo.Security.SecurityName
FROM  dbo.PortfolioQuantities INNER JOIN
               dbo.CurrentBalances ON dbo.PortfolioQuantities.Client_DPA_ = dbo.CurrentBalances.Client_DPA_ INNER JOIN
               dbo.Security ON dbo.PortfolioQuantities.Security_DPA_ = dbo.Security.Security_DPA_
WHERE (dbo.PortfolioQuantities.OrderTypeSale = 0)
ORDER BY dbo.PortfolioQuantities.Client_DPA_, dbo.PortfolioQuantities.Security_DPA_


GO
CREATE VIEW dbo.PendingCheques
AS
SELECT     dbo.Client.ClientName, dbo.Payment.PaymentReference AS ChequeNo, dbo.Payment.PaymentAmount, dbo.Payment.Payment_DPA_,
                      dbo.Payment.ChequeCollection, dbo.Payment.ChequeCollectionDate, dbo.Client.Client_DPA_, dbo.Payment.PaymentNarrative,
                      dbo.Payment.ChequeCollectionModifyUser, dbo.Payment.ChequeCollectionModifyDate, dbo.Payment.ChequeCollectionUser,
                      dbo.Payment.ChequeCollectionNarrative
FROM         dbo.Payment INNER JOIN
                      dbo.Client ON dbo.Payment.Entity_DPA_ = dbo.Client.Client_DPA_
WHERE     (dbo.Payment.ChequeCollection = N'AWAITING COLLECTION') AND (dbo.Payment.EntityType_DPA_ = 1) AND (dbo.Client.Deleted = 0) AND
                      (dbo.Payment.Deleted = 0) OR
                      (dbo.Client.Deleted IS NULL) AND (dbo.Payment.Deleted IS NULL)

GO

CREATE VIEW dbo.PendingJournals
AS
SELECT     dbo.JournalList.Journal_DPA_, dbo.JournalList.JournalDate, dbo.JournalList.JournalEntryUser, dbo.JournalList.JournalEntryAccount,
                      dbo.JournalList.JournalEntryEntity, dbo.JournalList.JournalEntryDebit, dbo.JournalList.JournalEntryCredit
FROM         dbo.Journal INNER JOIN
                      dbo.JournalList ON dbo.Journal.Journal_DPA_ = dbo.JournalList.Journal_DPA_
WHERE     (dbo.Journal.Released = 0) AND (dbo.Journal.Deleted = 0)


GO

CREATE VIEW dbo.performanceBonds
AS
SELECT     SUM(dbo.LevyContract.LevyAmount) AS Commission, SUM(dbo.Lots.LotGrossAmount) AS Gross, CAST(FLOOR(CAST(dbo.Lots.LotTDate AS Float))
                      AS DateTime) AS TransDate, dbo.Client.Owner_DPA_, dbo.Client.Agent_DPA_, dbo.Client.IsCustodian
FROM         dbo.LevyContract INNER JOIN
                      dbo.Lots ON dbo.LevyContract.Contract_DPA_ = dbo.Lots.Contract_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.Lots.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_
WHERE     (dbo.LevyContract.SystemMaintained = 11) AND (dbo.tbOrder.OrderSecType_DPA_ = 1)
GROUP BY CAST(FLOOR(CAST(dbo.Lots.LotTDate AS Float)) AS DateTime), dbo.Client.Agent_DPA_, dbo.Client.Owner_DPA_, dbo.Client.IsCustodian


GO

CREATE VIEW dbo.performanceShares
AS
SELECT     SUM(dbo.LevyContract.LevyAmount) AS Commission, SUM(dbo.Lots.LotGrossAmount) AS Gross, CAST(FLOOR(CAST(dbo.Lots.LotTDate AS Float))
                      AS DateTime) AS TransDate, dbo.Client.Owner_DPA_, dbo.Client.Agent_DPA_, dbo.Lots.Contract_DPA_
FROM         dbo.LevyContract INNER JOIN
                      dbo.Lots ON dbo.LevyContract.Contract_DPA_ = dbo.Lots.Contract_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.Lots.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_
WHERE     (dbo.LevyContract.SystemMaintained = 11) AND (dbo.tbOrder.OrderSecType_DPA_ = 2)
GROUP BY CAST(FLOOR(CAST(dbo.Lots.LotTDate AS Float)) AS DateTime), dbo.Client.Agent_DPA_, dbo.Client.Owner_DPA_, dbo.Lots.Contract_DPA_


GO
CREATE VIEW dbo.PnLStep1
AS

SELECT
	dbo.EntityList.AccountType_DPA_ AS AccountType_DPA_,
	dbo.EntityList.AccountTypeName AS AccountType,
	'------' + CAST(dbo.EntityList.Entity_DPA_ AS VARCHAR(500)) AS [Account Code],
	dbo.EntityList.EntityName AS AccountName,
	CASE
		WHEN SUM(dbo.LevyTransactionList.Balance) < 0 THEN 0 - SUM(dbo.LevyTransactionList.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.LevyTransactionList.Balance) >= 0 THEN SUM(dbo.LevyTransactionList.Balance)
		ELSE 0 END AS Credit
FROM         dbo.EntityList INNER JOIN
                      dbo.LevyTransactionList ON dbo.EntityList.Entity_DPA_ = dbo.LevyTransactionList.Entity_DPA_
WHERE 		dbo.EntityList.AccountType_DPA_ IN (1,2)
GROUP BY dbo.EntityList.AccountType_DPA_,dbo.EntityList.AccountTypeName,dbo.EntityList.EntityName, dbo.EntityList.Entity_DPA_

UNION ALL

SELECT
	dbo.EntityList.AccountType_DPA_ AS AccountType_DPA_,
	dbo.EntityList.AccountTypeName AS AccountType,
	'------' + CAST(dbo.EntityList.Entity_DPA_ AS VARCHAR(500)) AS [Account Code],
	dbo.EntityList.EntityName AS AccountName,
	CASE
		WHEN SUM(dbo.LevyTransactionList.Balance) < 0 THEN 0 - SUM(dbo.LevyTransactionList.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.LevyTransactionList.Balance) >= 0 THEN SUM(dbo.LevyTransactionList.Balance)
		ELSE 0 END AS Credit
FROM         dbo.EntityList INNER JOIN
                      dbo.LevyTransactionList ON dbo.EntityList.Entity_DPA_ = dbo.LevyTransactionList.Entity_DPA_
WHERE 		dbo.EntityList.AccountType_DPA_ IN (1,2)
GROUP BY dbo.EntityList.AccountType_DPA_,dbo.EntityList.AccountTypeName,dbo.EntityList.EntityName, dbo.EntityList.Entity_DPA_

UNION ALL


(SELECT
	dbo.AccountList.AccountTypeLevel1_DPA_ AS AccountType_DPA_,
	dbo.AccountList.AccountTypeLevel1 AS AccountType,
	CAST(dbo.AccountList.Account_DPA_ AS SQL_VARIANT) AS [Account Code],
	dbo.AccountList.AccountName,
	CASE
		WHEN SUM(dbo.NominalTransactionList.Balance) < 0 THEN 0 - SUM(dbo.NominalTransactionList.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.NominalTransactionList.Balance) >= 0 THEN SUM(dbo.NominalTransactionList.Balance)
		ELSE 0 END AS Credit
FROM         dbo.AccountList INNER JOIN
                      dbo.NominalTransactionList ON dbo.AccountList.Account_DPA_ = dbo.NominalTransactionList.Account_DPA_
WHERE 		dbo.AccountList.AccountTypeLevel1_DPA_ IN (1,2)
GROUP BY dbo.AccountList.AccountTypeLevel1_DPA_,dbo.AccountList.AccountTypeLevel1,dbo.AccountList.AccountName, dbo.AccountList.Account_DPA_)






GO

CREATE VIEW dbo.Portfoliobs
AS
SELECT DISTINCT
               dbo.SecurityListSecurity.SecurityCode, dbo.LatestPortfolios.Client_DPA_, dbo.SecurityListSecurity.Sector, dbo.SecurityListSecurity.SecurityName,
               dbo.LatestPortfolios.LastQuantity, dbo.LatestPortfolios.LastPrice, dbo.LatestPortfolios.MarketPrice, dbo.LatestPortfolios.SecurityBookValue,
               dbo.LatestPortfolios.SecurityCurrentValue, dbo.LatestPortfolios.LastPL, dbo.LatestPortfolios.LastPL_, dbo.LatestPortfolios.ClientCurrentValue,
               dbo.LatestPortfolios.ClientBookValue, dbo.LatestPortfolios.TotalPortfolio, dbo.LatestPortfolios.ClientName, dbo.LatestPortfolios.Cash
FROM  dbo.PortfoliosQuantities INNER JOIN
               dbo.LatestPortfolios ON dbo.PortfoliosQuantities.Portfolio_DPA_ = dbo.LatestPortfolios.Portfolio_DPA_ LEFT OUTER JOIN
               dbo.SecurityListSecurity ON dbo.LatestPortfolios.Security_DPA_ = dbo.SecurityListSecurity.Security_DPA_ AND
               dbo.PortfoliosQuantities.Security_DPA_ = dbo.SecurityListSecurity.Security_DPA_


GO

CREATE VIEW dbo.PortfolioQuantities
AS
SELECT TOP 100 PERCENT Order_DPA_, SecurityCode, Client_DPA_, Security_DPA_, OrderTypeSale, ContractNumber, LotSlipNo, OrdDetailPrice,
               OrdDetailQty, OrdDetailClient, OrdDetail_DPA_, Amount, SecurityMktPrice, Lot_DPA_, LotQty, LotPrice, LotTDate, LotGrossAmount
FROM  (SELECT TOP 100 PERCENT dbo.LotList.Order_DPA_, dbo.LotList.SecurityCode, dbo.LotList.Client_DPA_, dbo.LotList.Security_DPA_,
                              dbo.LotList.OrderTypeSale, dbo.LotList.ContractNumber, dbo.LotList.LotSlipNo, dbo.LotList.OrdDetailPrice, dbo.LotList.OrdDetailQty,
                              dbo.LotList.OrdDetailClient, dbo.LotList.OrdDetail_DPA_, dbo.OrdDetail.Amount, dbo.datastream_SecurityPriceList.Price AS SecurityMktPrice,
                              dbo.LotList.Lot_DPA_, dbo.LotList.LotQty, dbo.LotList.LotPrice, dbo.LotList.LotTDate, dbo.LotList.LotGrossAmount
               FROM   dbo.LotList INNER JOIN
                              dbo.OrdDetail ON dbo.LotList.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                              dbo.datastream_SecurityPriceList ON dbo.LotList.Security_DPA_ = dbo.datastream_SecurityPriceList.Security_DPA_
               WHERE (dbo.LotList.OrdDetailSecType = N'Security') AND (dbo.LotList.ContractNumber IS NOT NULL)
               UNION
               SELECT TOP 100 PERCENT ClientPortfolioQuantities.*
               FROM  ClientPortfolioQuantities) a
ORDER BY OrdDetailClient, SecurityCode, LotTDate, ContractNumber

GO


CREATE VIEW dbo.PrimaryIssuesList
AS
SELECT     dbo.Bond.BondIssue, dbo.Bond.Security_DPA_, dbo.Security.SecurityCode AS BondType, dbo.Client.ClientName, dbo.PrimaryIssues.*,
                      dbo.PaymentTypes.Description AS PaymentType, dbo.Users.UserName AS [User]
FROM         dbo.PrimaryIssues INNER JOIN
                      dbo.Client ON dbo.PrimaryIssues.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Bond ON dbo.PrimaryIssues.Bond_DPA_ = dbo.Bond.Bond_DPA_ INNER JOIN
                      dbo.Security ON dbo.Bond.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.PaymentTypes ON dbo.PrimaryIssues.PaymentTypes_DPA_ = dbo.PaymentTypes.PaymentTypes_DPA_ LEFT OUTER JOIN
                      dbo.Users ON dbo.PrimaryIssues.ModifiedBy = dbo.Users.UserID
WHERE     (dbo.PrimaryIssues.Deleted <> 1) OR
                      (dbo.PrimaryIssues.Deleted IS NULL)



GO
CREATE VIEW dbo.PrinterList
AS
SELECT     Printer_DPA_, PrinterActualName, PrinterName, Description, Active
FROM         dbo.Printers

GO

CREATE VIEW dbo.ProfitNLossAccount
AS
SELECT     AccountsStatement.*, CreditBal - Debit AS Balance
FROM         (SELECT     dbo.Payment.BankAccount_DPA_, cast(floor(cast(dbo.Payment.PaymentPDate AS float)) AS datetime) AS TransDate,
                                              ISNULL(Lots.ContractNumber, '') + ' ' + Payment.PaymentReference AS Ref,
                                              ISNULL(Security.SecurityCode + ' @ ' + CAST(Lots.LotPrice AS nvarchar), '') + ' ' + Payment.PaymentNarrative AS Particulars,
                                              CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                                              CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Credit,
                                              CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS CreditBal, 0 AS OpeningBalance,
                                              PaymentReceiptNo AS ReceiptNo
                       FROM          OrdDetail INNER JOIN
                                              Lots ON OrdDetail.OrdDetail_DPA_ = Lots.OrdDetail_DPA_ INNER JOIN
                                              Security ON OrdDetail.Security_DPA_ = Security.Security_DPA_ RIGHT OUTER JOIN
                                              Payment INNER JOIN
                                              PayType ON Payment.PayType_DPA_ = PayType.PayType_DPA_ ON Lots.Contract_DPA_ = Payment.Contract_DPA_
                       UNION ALL
                       SELECT     TOP 100 PERCENT 4 AS Bankcode, CAST(FLOOR(CAST(InterTransfer.TransferDate AS float)) AS DateTime) AS Date,
                                             Lot.contractNumber AS Ref, isnull(InterTransfer.TransferNarrative, '') + CASE (OrderType.OrderTypeSale)
                                             WHEN 1 THEN 'SALE ' ELSE ' PURCHASE ' END + InterTransferType.TypeDescription AS Particulars, CASE (OrderType.OrderTypeSale)
                                             WHEN 0 THEN intertransfer.Transferamount ELSE 0 END AS Debit, CASE (OrderType.OrderTypeSale)
                                             WHEN 1 THEN intertransfer.Transferamount ELSE 0 END AS Credit, CASE (OrderType.OrderTypeSale)
                                             WHEN 1 THEN intertransfer.Transferamount ELSE 0 END AS CreditBal, 0 AS IsOpeningBal, isnull(intertransfer.transferreference, '')
                                             AS Receiptno
                       FROM         OrdDetail INNER JOIN
                                             Lot ON OrdDetail.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN
                                             tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN
                                             OrderType ON tbOrder.OrderType_DPA_ = OrderType.OrderType_DPA_ INNER JOIN
                                             InterTransfer INNER JOIN
                                             InterTransferType ON InterTransfer.InterTransferType_DPA_ = InterTransferType.InterTransferType_DPA_ ON
                                             Lot.Contract_DPA_ = InterTransfer.Contract_DPA_
                       UNION ALL
                       SELECT     dbo.Payment.Entity_DPA_ AS Entity_DPA_, cast(floor(cast(dbo.Payment.PaymentPDate AS float)) AS DateTime) AS TransDate,
                                             dbo.Payment.PaymentReference AS Ref, dbo.Payment.PaymentNarrative AS Particulars,
                                             CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS CreditBal, 0 AS IsOpeningBalance,
                                             PaymentReceiptNo AS ReceiptNo
                       FROM         dbo.Payment INNER JOIN
                                             dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
                       WHERE     (EntityType_DPA_ = 5)
                       UNION ALL
                       SELECT     dbo.JournalList.Entity_DPA_, Cast(floor(Cast(dbo.JournalList.JournalDate AS float)) AS DateTime) AS TransDate,
                                             CAST(dbo.JournalList.JournalEntry_DPA_ AS SQL_VARIANT) AS Ref, dbo.JournalList.JournalNarrative AS Particulars,
                                             JournalList.JournalEntryDebit AS Debit, JournalList.JournalEntryCredit AS Credit, JournalList.JournalEntryCredit AS CreditBal,
                                             0 AS IsOpeningBalance, '' AS ReceiptNo
                       FROM         dbo.JournalList
                       WHERE     (EntityType_DPA_ = 5) AND (Entity_DPA_ IN
                                                 (SELECT     Account_DPA_
                                                   FROM          ACCOUNT
                                                   WHERE      (Account_DPA_ IN
                                                                              (SELECT     Entity_DPA_
                                                                                FROM          Payment
                                                                                WHERE      EntityType_DPA_ = 5))))) AccountsStatement

GO
CREATE VIEW dbo.PurchaseContractList
AS
SELECT     dbo.Contract.Contract_DPA_, dbo.Contract.ContractTransferNo, dbo.Contract.ContractDeliveryDate, dbo.LotList.Lot_DPA_, dbo.LotList.BrokerCode,
                      dbo.LotList.OrdDetailType, dbo.LotList.OrderTypeSale, dbo.LotList.LotSlipNo, dbo.LotList.LotTDate, dbo.LotList.LotQty, dbo.LotList.LotPrice,
                      dbo.LotList.ContractNumber, dbo.LotList.StatusDescription, dbo.LotList.OrdDetailClient, dbo.Contract.ContractDelivered, dbo.LotList.OrdDetailSecurity,
                      dbo.Contract.ContractNCertificate, dbo.Contract.ContractNCDate, dbo.Contract.ContractNCDelivered, dbo.Contract.Voucher_DPA_,
                      dbo.Contract.ContractVouchered, DATEADD([day], 7, dbo.LotList.LotTDate) AS Settlementdate
FROM         dbo.Contract INNER JOIN
                      dbo.LotList ON dbo.Contract.Contract_DPA_ = dbo.LotList.Contract_DPA_
WHERE     (dbo.Contract.Deleted = 0) AND (dbo.LotList.ContractNumber LIKE 'P%') OR
                      (dbo.Contract.Deleted IS NULL)

GO
CREATE VIEW dbo.PurchaseOrderList
AS
SELECT     Order_DPA_ AS [Order No], OrderDate AS [Order Date], OrdDetailClient + '	[' + CONVERT(nvarchar(4000), Client_DPA_) + ']' AS Client,
                      OrdDetailSecurity AS Security, OrdDetailQty AS [Order Qty], OrdDetailPrice AS [Order Price], LotSlipNo AS [Slip No], LotTDate AS [Trade Date],
                      LotQty AS Qty, LotPrice AS Price, BrokerCode AS Broker, ContractNumber AS Contract
FROM         dbo.LotList
WHERE     (ContractNumber LIKE 'P%')

GO

CREATE VIEW dbo.PurchaseOrders
AS
SELECT DISTINCT
               TOP 100 PERCENT dbo.LotList.Order_DPA_, dbo.LotList.OrderDate, dbo.LotList.OrderTypeSale, dbo.LotList.Client_DPA_, dbo.LotList.OrdDetailClient,
               dbo.LotList.OrdDetailSecType, dbo.LotList.SecurityCode, dbo.LotList.OrdDetailQty, dbo.LotList.BalanceQty, dbo.LotList.OrdDetailPrice,
               dbo.LotList.OrdDetail_DPA_, dbo.OrdDetail.Amount, dbo.datastream_SecurityPriceList.Price, dbo.OrdDetail.Best, dbo.ClientTotals.Total,
               dbo.CurrentBalances.CurrentBal, dbo.Client.CreditLimit, - dbo.Client.CreditLimit - (dbo.CurrentBalances.CurrentBal - dbo.ClientTotals.Total)
               AS Excess, dbo.LotList.OrdDetailType
FROM  dbo.LotList INNER JOIN
               dbo.OrdDetail ON dbo.LotList.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
               dbo.Security ON dbo.LotList.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
               dbo.datastream_SecurityPriceList ON dbo.Security.Security_DPA_ = dbo.datastream_SecurityPriceList.Security_DPA_ INNER JOIN
               dbo.Client ON dbo.LotList.Client_DPA_ = dbo.Client.Client_DPA_ LEFT OUTER JOIN
               dbo.CurrentBalances ON dbo.LotList.Client_DPA_ = dbo.CurrentBalances.Client_DPA_ LEFT OUTER JOIN
               dbo.ClientTotals ON dbo.LotList.Client_DPA_ = dbo.ClientTotals.Client_DPA_
WHERE (RTRIM(dbo.LotList.OrdDetailType) LIKE '%Purchase%')
ORDER BY dbo.LotList.Client_DPA_


GO
CREATE VIEW dbo.PurchaseOrdersList
AS
SELECT DISTINCT
                      TOP 100 PERCENT dbo.LotList.Order_DPA_, dbo.LotList.OrderDate, dbo.LotList.OrderTypeSale, dbo.LotList.Client_DPA_, dbo.LotList.OrdDetailClient,
                      dbo.LotList.OrdDetailSecType, dbo.LotList.SecurityCode, ISNULL(dbo.LotList.OrdDetailQty, 0) AS OrdDetailQty, ISNULL(dbo.LotList.BalanceQty, 0)
                      AS BalanceQty, dbo.LotList.OrdDetailPrice, dbo.LotList.OrdDetail_DPA_, ISNULL(dbo.OrdDetail.Amount, 0) AS Amount,
                      ISNULL(dbo.datastream_SecurityPriceList.Price, 0) AS Price, ISNULL(dbo.LotList.Best, 0) AS Best
FROM         dbo.LotList INNER JOIN
                      dbo.OrdDetail ON dbo.LotList.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.Security ON dbo.LotList.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.datastream_SecurityPriceList ON dbo.Security.Security_DPA_ = dbo.datastream_SecurityPriceList.Security_DPA_
WHERE     (RTRIM(dbo.LotList.OrdDetailType) LIKE '%Purchase%')
ORDER BY dbo.LotList.Client_DPA_

GO

CREATE VIEW dbo.RealUnMatched
AS
SELECT dbo._CDS_Imported_Trades_.CDSImport_DPA_, dbo.OrdDetailList.OrdDetail_DPA_, dbo.OrdDetailList.Order_DPA_, dbo.OrdDetailList.OrderDate,
               dbo.OrdDetailList.OrderTypeSale, dbo.OrdDetailList.CDSOrderTypeSale, dbo.OrdDetailList.OrdDetailClient, dbo.OrdDetailList.SecurityCode,
               dbo.OrdDetailList.BalanceQty, dbo._CDS_Imported_Trades_.CDSRef, dbo._CDS_Imported_Trades_.Quantity, dbo._CDS_Imported_Trades_.Price,
               REPLACE(dbo._CDS_Imported_Trades_.ContraBrokerID, 'B', '') AS BrokerCode, dbo._CDS_Imported_Trades_.ContraBrokerID AS CDSBrokerCode,
               dbo.OrdDetailList.OrdDetailType, dbo.OrdDetailList.OrdDetailSecType, dbo.OrdDetailList.Security_DPA_, dbo.OrdDetailList.CommissionRate,
               dbo.OrdDetailList.OrdDetailQty, dbo.OrdDetailList.VolumeRate, dbo.OrdDetailList.VolumeBoundary, dbo.OrdDetailList.MinimumCommission,
               dbo.OrdDetailList.CMARegulated, dbo.OrdDetailList.PostImmobilisedRate, dbo.OrdDetailList.SecurityImmobilised, dbo.Broker.Broker_DPA_,
               dbo._CDS_Imported_Trades_.TradeTime, dbo.OrdDetailList.AgentCommission, dbo.OrdDetailList.StaffCommission, CONVERT(smalldatetime,
               SUBSTRING(dbo._CDS_Imported_Trades_.TradeDate, 3, 2) + '/' + SUBSTRING(dbo._CDS_Imported_Trades_.TradeDate, 1, 2)
               + '/' + SUBSTRING(dbo._CDS_Imported_Trades_.TradeDate, 5, 4)) AS TradeDate, CONVERT(smalldatetime,
               SUBSTRING(dbo._CDS_Imported_Trades_.SettlementDate, 3, 2) + '/' + SUBSTRING(dbo._CDS_Imported_Trades_.SettlementDate, 1, 2)
               + '/' + SUBSTRING(dbo._CDS_Imported_Trades_.SettlementDate, 5, 4)) AS SettlementDate, dbo._CDS_Imported_Trades_.SettlementAmount
FROM  dbo.OrdDetailList RIGHT OUTER JOIN
               dbo._CDS_Imported_Trades_ ON
               dbo.OrdDetailList.ClientCDSNo = dbo._CDS_Imported_Trades_.ClientPrefix + dbo._CDS_Imported_Trades_.ClientSuffix AND
               dbo.OrdDetailList.SecurityCode = dbo._CDS_Imported_Trades_.SecurityDescription AND
               dbo.OrdDetailList.CDSOrderTypeSale = dbo._CDS_Imported_Trades_.BuySell INNER JOIN
               dbo.Broker ON LTRIM(RTRIM(REPLACE(dbo._CDS_Imported_Trades_.ContraBrokerID, 'B', ''))) = dbo.Broker.BrokerCode
WHERE (dbo._CDS_Imported_Trades_.Processed = 0) AND (dbo.OrdDetailList.Order_DPA_ IS NULL)


GO
CREATE VIEW dbo.ReceiptList
AS
SELECT     TOP 100 PERCENT dbo.Payments.Payment_DPA_, dbo.Payments.PaymentPDate, dbo.CompleteEntityList.EntityType,
                      dbo.CompleteEntityList.EntityName, dbo.BankAccountList.AccountName AS BankAccountName, CAST(dbo.Payments.PaymentAmount AS Float)
                      AS PaymentAmount, dbo.Payments.PaymentReceiptNo, dbo.Payments.Order_DPA_, dbo.PaymentTypes.Description AS PaymentType,
                      dbo.Payments.PaymentReference, dbo.CompleteEntityList.EntityCode, dbo.Payments.EntityType_DPA_, dbo.Payments.PaymentNarrative,
                      dbo.Payments.TimeChanged, dbo.UserList.[USER] AS ChangedBy, dbo.tbOrder.OrderSecType_DPA_
FROM         dbo.UserList RIGHT OUTER JOIN
                      dbo.Payments INNER JOIN
                      dbo.BankAccountList ON dbo.Payments.BankAccount_DPA_ = dbo.BankAccountList.Account_DPA_ INNER JOIN
                      dbo.PayType ON dbo.Payments.PayType_DPA_ = dbo.PayType.PayType_DPA_ INNER JOIN
                      dbo.CompleteEntityList ON dbo.Payments.EntityType_DPA_ = dbo.CompleteEntityList.EntityType_DPA_ AND
                      dbo.Payments.Entity_DPA_ = dbo.CompleteEntityList.Entity_DPA_ LEFT OUTER JOIN
                      dbo.tbOrder ON dbo.Payments.Order_DPA_ = dbo.tbOrder.Order_DPA_ ON dbo.UserList.UserID = dbo.Payments.ChangedBy LEFT OUTER JOIN
                      dbo.PaymentTypes ON dbo.Payments.PaymentTypes_DPA_ = dbo.PaymentTypes.PaymentTypes_DPA_
WHERE     (dbo.PayType.PayTypeIn = 1) AND (NOT (dbo.Payments.PaymentReceiptNo IS NULL))
ORDER BY dbo.Payments.Payment_DPA_ DESC

GO
CREATE VIEW dbo.ReceiptSchedule
AS
SELECT     TOP 100 PERCENT dbo.CompleteEntityList.EntityType AS Entity, dbo.CompleteEntityList.EntityName AS Party,
                      dbo.BankAccountList.AccountName AS Bank_Account, dbo.Payment.PaymentReceiptNo AS Receipt_No, dbo.Payment.PaymentReference AS Reference,
                      CAST(dbo.Payment.PaymentPDate AS NVARCHAR(12)) AS Payment_Date, dbo.Payment.PaymentAmount AS Amount,
                      CASE (dbo.CompleteEntityList.EntityType_DPA_) WHEN 8 THEN 1 ELSE 0 END AS CdsGroup, dbo.BankAccountList.AccountCode,
                      dbo.Payment.PaymentNarrative
FROM         dbo.Payment INNER JOIN
                      dbo.BankAccountList ON dbo.Payment.BankAccount_DPA_ = dbo.BankAccountList.Account_DPA_ INNER JOIN
                      dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_ INNER JOIN
                      dbo.CompleteEntityList ON dbo.Payment.EntityType_DPA_ = dbo.CompleteEntityList.EntityType_DPA_ AND
                      dbo.Payment.Entity_DPA_ = dbo.CompleteEntityList.Entity_DPA_
WHERE     (dbo.PayType.PayTypeIn = 1) AND (dbo.Payment.Deleted = 0)
ORDER BY CdsGroup, Receipt_No

GO
CREATE VIEW dbo.ReceiptsReport
AS
SELECT      TOP 100 PERCENT
	 dbo.CompleteEntityList.EntityType [Entity],
	dbo.CompleteEntityList.EntityName [Recipient],
                     dbo.BankAccountList.AccountName AS [Bank Account],
	dbo.Payment.PaymentReference [Reference],
		CAST(dbo.Payment.PaymentPDate AS NVARCHAR(12)) As [Payment Date],
				dbo.Payment.PaymentAmount [Amount]
FROM         dbo.Payment INNER JOIN
                      dbo.BankAccountList ON dbo.Payment.BankAccount_DPA_ = dbo.BankAccountList.Account_DPA_ INNER JOIN
                      dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_ INNER JOIN
                      dbo.CompleteEntityList ON dbo.Payment.EntityType_DPA_ = dbo.CompleteEntityList.EntityType_DPA_ AND
                      dbo.Payment.Entity_DPA_ = dbo.CompleteEntityList.Entity_DPA_
WHERE     (dbo.PayType.PayTypeIn = 1)
ORDER BY dbo.Payment.PaymentPDate


GO

CREATE VIEW dbo.ReconciliationList
AS
SELECT     TOP 100 PERCENT EntryID, BankCode, AccountName, Date, Reference, Particulars, Debit, Credit, ReconcileDate, Type, TimeChanged,
                      Users.UserName AS ChangedBy
FROM         (SELECT     EntryID, BankCode, AccountName, Date, Reference, Particulars, CASE WHEN SUM(Credit - Debit) < 0 THEN 0 - SUM(Credit - Debit)
                                              ELSE 0 END AS Debit, CASE WHEN SUM(Credit - Debit) >= 0 THEN SUM(Credit - Debit) ELSE 0 END AS Credit, ReconcileDate, Type,
                                              MAX(b.TimeChanged) AS TimeChanged, MAX(b.ChangedBy) AS ChangedBy
                       FROM          (SELECT DISTINCT
                                                                      TOP 100 MIN(dbo.Payments.Payment_DPA_) AS EntryID, dbo.Account.AccountCode AS BankCode, dbo.Account.AccountName,
                                                                      Cast(Floor(Cast(dbo.Payments.PaymentPDate AS Float)) AS DateTime) AS Date, '' AS Reference, '' AS Particulars,
                                                                      SUM(CASE PayType.PayTypeIn WHEN 0 THEN Payments.PaymentAmount ELSE 0 END) AS Debit,
                                                                      SUM(CASE PayType.PayTypeIn WHEN 1 THEN Payments.PaymentAmount ELSE 0 END) AS Credit, dbo.Payments.ReconcileDate,
                                                                      3 AS type, MAX(Payments.TimeChanged) AS TimeChanged, MAX(Payments.ChangedBy) AS ChangedBy
                                               FROM          dbo.Payments INNER JOIN
                                                                      dbo.Account ON dbo.Payments.BankAccount_DPA_ = dbo.Account.Account_DPA_ INNER JOIN
                                                                      dbo.PayType ON dbo.Payments.PayType_DPA_ = dbo.PayType.PayType_DPA_
                                               WHERE      dbo.Account.Account_DPA_ = 4
                                               GROUP BY dbo.Account.AccountCode, dbo.Account.AccountName, dbo.Payments.ReconcileDate,
                                                                      Cast(Floor(Cast(dbo.Payments.PaymentPDate AS Float)) AS DateTime)) b
                       GROUP BY EntryID, BankCode, AccountName, Date, Reference, Particulars,ReconcileDate, Type
                       UNION
                       SELECT     TOP 100 PERCENT journalEntry.JournalEntry_DPA_, CONVERT(Char(50), JournalEntry.Entity_DPA_), CASE (JournalEntry.EntityType_DPA_)
                                             WHEN 1 THEN
                                                 (SELECT     Client.ClientName
                                                   FROM          Client
                                                   WHERE      Client.Client_DPA_ = JournalEntry.Entity_DPA_) WHEN 5 THEN
                                                 (SELECT     AccountName
                                                   FROM          Account
                                                   WHERE      Account.Account_DPA_ = JournalEntry.Entity_DPA_) END AS AccountName, Cast(Floor(Cast(Journal.JournalDate AS Float))
                                             AS DateTime) AS Date, '' AS Ref, 'Journal  ' + Journal.JournalNarrative AS Narrative, JournalEntry.JournalEntryDebit AS Debit,
                                             JournalEntry.JournalEntryCredit AS Credit, JournalEntry.ReconcileDate, 2 AS type, JournalEntry.TimeChanged AS TimeChanged,
                                             JournalEntry.ChangedBy AS ChangedBy
                       FROM         Journal INNER JOIN
                                             JournalEntry ON Journal.Journal_DPA_ = JournalEntry.Journal_DPA_
                       UNION
                       SELECT DISTINCT
                      TOP 100 PERCENT dbo.Payments.Payment_DPA_ AS EntryID, dbo.Account.AccountCode AS BankCode, dbo.Account.AccountName,
                      Cast(Floor(Cast(dbo.Payments.PaymentPDate AS Float)) AS DateTime) AS Date, dbo.Payments.PaymentReference AS Reference,
                      CASE PayType.PayTypeIn WHEN 0 THEN 'Payment  ' + Payments.PaymentNarrative ELSE 'Receipt  ' + Payments.PaymentNarrative END AS Particulars,
                       CASE PayType.PayTypeIn WHEN 0 THEN Payments.PaymentAmount ELSE 0 END AS Debit,
                      CASE PayType.PayTypeIn WHEN 1 THEN Payments.PaymentAmount ELSE 0 END AS Credit, dbo.Payments.ReconcileDate, 1 AS type,
                      Payments.TimeChanged AS TimeChanged, Payments.ChangedBy AS ChangedBy
FROM         dbo.Payments INNER JOIN
                      dbo.Account ON dbo.Payments.BankAccount_DPA_ = dbo.Account.Account_DPA_ INNER JOIN
                      dbo.PayType ON dbo.Payments.PayType_DPA_ = dbo.PayType.PayType_DPA_
WHERE    dbo.Account.Account_DPA_ <> 4
UNION
SELECT DISTINCT
                      TOP 100 PERCENT dbo.Payments.Payment_DPA_ AS EntryID, dbo.Account.AccountCode AS BankCode, dbo.Account.AccountName,
                      Cast(Floor(Cast(dbo.Payments.PaymentPDate AS Float)) AS DateTime) AS Date, dbo.Payments.PaymentReference AS Reference,
                      CASE PayType.PayTypeIn WHEN 1 THEN 'Payment  ' + Payments.PaymentNarrative ELSE 'Receipt  ' + Payments.PaymentNarrative END AS Particulars,
                       CASE PayType.PayTypeIn WHEN 1 THEN Payments.PaymentAmount ELSE 0 END AS Debit,
                      CASE PayType.PayTypeIn WHEN 0 THEN Payments.PaymentAmount ELSE 0 END AS Credit, dbo.Payments.ReconcileDate, 1 AS type,
                      Payments.TimeChanged AS TimeChanged, Payments.ChangedBy AS ChangedBy
FROM         dbo.Payments INNER JOIN
                      dbo.Account ON dbo.Payments.Entity_DPA_ = dbo.Account.Account_DPA_ INNER JOIN
                      dbo.PayType ON dbo.Payments.PayType_DPA_ = dbo.PayType.PayType_DPA_
WHERE     Payments.EntityType_DPA_ = 5
ORDER BY BankCode) a LEFT OUTER JOIN
                      users ON a.ChangedBy = Users.Userid
ORDER BY a.Date DESC










GO

CREATE VIEW dbo.ResidencyList
AS
SELECT     TOP 100 PERCENT ResidencyDescription AS ResidencyName, Residency_DPA_, DefaultSelection
FROM         dbo.Residency
ORDER BY ResidencyDescription


GO
CREATE VIEW dbo.SaleContractList
AS
SELECT     dbo.Contract.Contract_DPA_, dbo.Contract.ContractTransferNo, dbo.Contract.ContractDeliveryDate, dbo.LotList.Lot_DPA_, dbo.LotList.BrokerCode,
                      dbo.LotList.OrdDetailType, dbo.LotList.OrderTypeSale, dbo.LotList.LotSlipNo, dbo.LotList.LotTDate, dbo.LotList.LotQty, dbo.LotList.LotPrice,
                      dbo.LotList.ContractNumber, dbo.LotList.StatusDescription, dbo.LotList.OrdDetailClient, dbo.Contract.ContractDelivered, dbo.LotList.OrdDetailSecurity,
                      dbo.Contract.ContractNCertificate, dbo.Contract.ContractNCDate, dbo.Contract.ContractNCDelivered, dbo.Contract.Voucher_DPA_,
                      dbo.Contract.ContractVouchered, dbo.LotList.ContractNumber AS Expr1, dbo.Contract.ContractCertNo AS ContractCertificate, DATEADD([day], 7,
                      dbo.LotList.LotTDate) AS Settlementdate
FROM         dbo.Contract INNER JOIN
                      dbo.LotList ON dbo.Contract.Contract_DPA_ = dbo.LotList.Contract_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.LotList.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_
WHERE     (dbo.OrdDetail.Deleted = 0 OR
                      dbo.OrdDetail.Deleted IS NULL) AND (dbo.Contract.Deleted = 0 OR
                      dbo.Contract.Deleted IS NULL) AND (dbo.LotList.ContractNumber LIKE 'S%')

GO
CREATE VIEW dbo.SaleContracts
AS
SELECT     dbo.Contract.Contract_DPA_, dbo.Contract.ContractTransferNo, dbo.Contract.ContractDeliveryDate, dbo.LotList.Lot_DPA_, dbo.LotList.BrokerCode,
                      dbo.LotList.OrdDetailSecType, dbo.LotList.OrdDetailType, dbo.LotList.OrderTypeSale, dbo.LotList.LotSlipNo, dbo.LotList.LotTDate, dbo.LotList.LotQty,
                      dbo.LotList.LotPrice, dbo.LotList.LotGrossAmount, dbo.LotList.ContractNumber, dbo.LotList.StatusDescription, dbo.LotList.OrdDetailClient,
                      dbo.Contract.ContractDelivered, dbo.LotList.OrdDetailSecurity, dbo.Contract.ContractNCertificate, dbo.Contract.ContractNCDate,
                      dbo.Contract.ContractNCDelivered, dbo.Contract.Voucher_DPA_, dbo.Contract.ContractVouchered,
                      CASE WHEN LotList.OrdDetailSecType = 'Fixed' THEN LevyContract.LevyRate * LotList.LotGrossAmount / 100 ELSE LevyContract.LevyAmount END AS LevyAmount,
                       dbo.LevyContract.LevyName, dbo.LevyContract.SystemMaintained, CASE WHEN ISNULL(CHARINDEX('%', dbo.LevyContract.LevyRatePercentage), 0)
                      > 0 THEN dbo.LevyContract.LevyRatePercentage ELSE '' END AS LevyRatePercentage, dbo.LevyContract.LevyShortName, dbo.tbOrder.Client_DPA_,
                      dbo.LevyContract.LevyAmount AS Expr1, dbo.LevyContract.LevyVATAmount, dbo.LotList.ContractSettlementDate
FROM         dbo.Contract INNER JOIN
                      dbo.LotList ON dbo.Contract.Contract_DPA_ = dbo.LotList.Contract_DPA_ INNER JOIN
                      dbo.LevyContract ON dbo.Contract.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.LotList.Order_DPA_ = dbo.tbOrder.Order_DPA_
WHERE     (dbo.tbOrder.OrderCompounded = 0)

GO
CREATE VIEW dbo.SaleContracts
AS


SELECT     dbo.Contract.Contract_DPA_, dbo.Contract.ContractTransferNo, dbo.Contract.ContractDeliveryDate, dbo.LotList.Lot_DPA_, dbo.LotList.BrokerCode,
                      dbo.LotList.OrdDetailSecType, dbo.LotList.OrdDetailType, dbo.LotList.OrderTypeSale, dbo.LotList.LotSlipNo, dbo.LotList.LotTDate, dbo.LotList.LotQty,
                      dbo.LotList.LotPrice, dbo.LotList.LotGrossAmount, dbo.LotList.ContractNumber, dbo.LotList.StatusDescription, dbo.LotList.OrdDetailClient,
                      dbo.Contract.ContractDelivered, dbo.LotList.OrdDetailSecurity, dbo.Contract.ContractNCertificate, dbo.Contract.ContractNCDate,
                      dbo.Contract.ContractNCDelivered, dbo.Contract.Voucher_DPA_, dbo.Contract.ContractVouchered, dbo.LevyContract.LevyAmount,
                      dbo.LevyContract.LevyName, dbo.LevyContract.SystemMaintained, CASE WHEN ISNULL(CHARINDEX('%', dbo.LevyContract.LevyRatePercentage), 0)
                      > 0 THEN dbo.LevyContract.LevyRatePercentage ELSE '' END AS LevyRatePercentage, dbo.LevyContract.LevyShortName
FROM         dbo.Contract INNER JOIN
                      dbo.LotList ON dbo.Contract.Contract_DPA_ = dbo.LotList.Contract_DPA_ INNER JOIN
                      dbo.LevyContract ON dbo.Contract.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.LotList.Order_DPA_ = dbo.tbOrder.Order_DPA_
WHERE     (dbo.tbOrder.OrderCompounded = 0)
GO

CREATE VIEW dbo.SaleContractsInExcel
AS


SELECT     SaleContracts.*, dbo.Agent.Agent_DPA_ AS AgentCode, dbo.Agent.AgentName AS Agent
FROM         dbo.Lot INNER JOIN
                      dbo.OrdDetail ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client INNER JOIN
                      dbo.Agent ON dbo.Client.Agent_DPA_ = dbo.Agent.Agent_DPA_ ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ RIGHT OUTER JOIN

(SELECT     dbo.tbOrder.Order_DPA_, dbo.Contract.Contract_DPA_, dbo.Contract.ContractTransferNo, dbo.Contract.ContractDeliveryDate, dbo.LotList.Lot_DPA_, dbo.LotList.BrokerCode,
                      dbo.LotList.OrdDetailSecType, dbo.LotList.OrdDetailType, dbo.LotList.OrderTypeSale, dbo.LotList.LotSlipNo, dbo.LotList.LotTDate, dbo.LotList.LotQty,
                      dbo.LotList.LotPrice, dbo.LotList.LotGrossAmount, dbo.LotList.ContractNumber, dbo.LotList.StatusDescription, dbo.LotList.OrdDetailClient,
                      dbo.Contract.ContractDelivered, dbo.LotList.OrdDetailSecurity, dbo.Contract.ContractNCertificate, dbo.Contract.ContractNCDate,
                      dbo.Contract.ContractNCDelivered, dbo.Contract.Voucher_DPA_, dbo.Contract.ContractVouchered, dbo.LevyContract.LevyAmount,
                      dbo.LevyContract.LevyName, dbo.LevyContract.SystemMaintained, CASE WHEN ISNULL(CHARINDEX('%', dbo.LevyContract.LevyRatePercentage), 0)
                      > 0 THEN dbo.LevyContract.LevyRatePercentage ELSE NULL END AS LevyRatePercentage, dbo.LevyContract.LevyShortName,
                      Lotlist.MinimumCommission
FROM         dbo.Contract INNER JOIN
                      dbo.LotList ON dbo.Contract.Contract_DPA_ = dbo.LotList.Contract_DPA_ INNER JOIN
                      dbo.LevyContract ON dbo.Contract.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.LotList.Order_DPA_ = dbo.tbOrder.Order_DPA_
WHERE     (dbo.tbOrder.OrderCompounded = 0) AND LevyContract.Deleted = 0 AND tbOrder.Deleted = 0 AND Contract.Deleted = 0) as SaleContracts

ON dbo.Lot.Lot_DPA_ = SaleContracts.Lot_DPA_





GO
CREATE VIEW dbo.SaleOrderList
AS
SELECT     Order_DPA_ AS [Order No], OrderDate AS [Order Date], OrdDetailClient + '	[' + CONVERT(nvarchar(4000), Client_DPA_) + ']' AS Client,
                      OrdDetailSecurity AS Security, OrdDetailQty AS [Order Qty], OrdDetailPrice AS [Order Price], LotSlipNo AS [Slip No], LotTDate AS [Trade Date],
                      LotQty AS Qty, LotPrice AS Price, BrokerCode AS Broker, ContractNumber AS Contract
FROM         dbo.LotList
WHERE     (ContractNumber LIKE 'S%')

GO


CREATE VIEW dbo.SalesOrdersTotals
AS
SELECT     dbo.tbOrder.Client_DPA_, dbo.OrdDetail.Security_DPA_, SUM(CASE ISNULL(dbo.OrdDetailContractedQtyList.ContractQty, 1)
                      WHEN 1 THEN dbo.OrdDetail.OrdDetailQty ELSE dbo.OrdDetail.OrdDetailQty - dbo.OrdDetailContractedQtyList.ContractQty END) AS Total,
                      OrdDetail.Order_DPA_
FROM         dbo.OrdDetail INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ LEFT OUTER JOIN
                      dbo.OrdDetailContractedQtyList ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.OrdDetailContractedQtyList.OrdDetail_DPA_
WHERE     OrdDetail.Deleted = 0
GROUP BY dbo.tbOrder.Client_DPA_, dbo.OrdDetail.Security_DPA_, OrdDetail.Order_DPA_






GO
CREATE VIEW dbo.SecTransFeeList
AS
SELECT     dbo.SecTransFee.SecTransFeeADate AS SecTransFeeADate, dbo.SecTransFee.SecTransFeeActive AS SecTransFeeActive,
                      dbo.Security.SecurityName + ', ' + CONVERT(nvarchar(1000), dbo.SecTransFee.SecTransFeeFee) AS SecTransFeeSummary,
                      dbo.SecTransFee.SecTransFee_DPA_ AS SecTransFee_DPA_, dbo.Security.Security_DPA_, dbo.SecTransFee.SecTransFeeFee
FROM         dbo.Security INNER JOIN
                      dbo.SecTransFee ON dbo.Security.Security_DPA_ = dbo.SecTransFee.Security_DPA_

GO
CREATE VIEW SecTransFeeListLatest AS SELECT SecTransFeeList.Security_DPA_, SecTransFeeList.SecTransFeeFee AS Fee, Max(SecTransFeeList.SecTransFeeADate) AS SecTransFeeADate
FROM SecTransFeeList
GROUP BY SecTransFeeList.Security_DPA_, SecTransFeeList.SecTransFeeFee

GO

CREATE   VIEW app.Securities AS
SELECT  Security_DPA_      AS SecurityId,
        SecurityCode       AS Code,
        SecurityName       AS Name,
        SecurityMktPrice   AS MarketPrice,
        OrderSecType_DPA_  AS SecTypeId,
        Sector_DPA_        AS SectorId,
        CanTrade           AS CanTrade,
        Immobilised        AS IsImmobilised
FROM dbo.Security;   -- no Deleted column

GO
CREATE VIEW dbo.SecurityList
AS
SELECT     TOP 100 PERCENT dbo.SecurityListFilter.SecurityName, dbo.SecurityListFilter.SecurityCode, dbo.SecurityListFilter.SecurityMktPrice,
                      REPLACE(REPLACE(REPLACE(dbo.SecurityListFilter.SecurityAddr, CHAR(13), ','), CHAR(10), ''), CHAR(34), ' ') AS SecurityAddr,
                      dbo.SecurityListFilter.OrderSecTypeDescription, dbo.SecurityListFilter.Security_DPA_, dbo.SecurityListFilter.OrderSecTypeDisplayName,
                      dbo.SecTransFee.SecTransFeeFee, dbo.SecurityListFilter.OrderSecType_DPA_, dbo.SecurityListFilter.SecurityNameEx,
                      dbo.SecurityListFilter.Immobilised, dbo.SecurityListFilter.ExpiryDate, dbo.SecurityListFilter.OfferType
FROM         dbo.SecurityListFilter INNER JOIN
                      dbo.SecTransFee ON dbo.SecurityListFilter.Security_DPA_ = dbo.SecTransFee.Security_DPA_ AND
                      dbo.SecurityListFilter.MaxOfSecTransFeeADate = dbo.SecTransFee.SecTransFeeADate
GROUP BY dbo.SecurityListFilter.SecurityName, dbo.SecurityListFilter.SecurityCode, dbo.SecurityListFilter.SecurityMktPrice, dbo.SecurityListFilter.SecurityAddr,
                      dbo.SecurityListFilter.OrderSecTypeDescription, dbo.SecurityListFilter.Security_DPA_, dbo.SecurityListFilter.OrderSecTypeDisplayName,
                      dbo.SecTransFee.SecTransFeeFee, dbo.SecurityListFilter.OrderSecType_DPA_, dbo.SecurityListFilter.SecurityNameEx,
                      dbo.SecurityListFilter.Immobilised, dbo.SecurityListFilter.ExpiryDate, dbo.SecurityListFilter.OfferType
ORDER BY dbo.SecurityListFilter.SecurityName

GO
CREATE VIEW dbo.SecurityListFilter
AS
SELECT     TOP 100 PERCENT dbo.Security.SecurityCode + SPACE
                          ((SELECT     ISNULL(MAX(LEN(SecurityCode)), 0) AS MAXLEN
                              FROM         dbo.Security) - LEN(dbo.Security.SecurityCode)) + ' : ' + dbo.Security.SecurityName AS SecurityNameEx, dbo.Security.SecurityName,
                      dbo.Security.SecurityCode, dbo.Security.SecurityMktPrice, dbo.Security.SecurityAddr, dbo.OrderSecType.OrderSecTypeDescription,
                      MAX(dbo.SecTransFee.SecTransFeeADate) AS MaxOfSecTransFeeADate, dbo.Security.Security_DPA_, dbo.OrderSecType.OrderSecTypeDisplayName,
                      dbo.Security.OrderSecType_DPA_, dbo.Security.Immobilised, dbo.Security.Sector_DPA_, dbo.Security.ImportCode, dbo.Security.CanTrade,
                      dbo.Security.Offerings, dbo.Security.ExpiryDate, dbo.Security.BatchSize, dbo.Security.ClosingDate,
                      dbo.Users.Surname + N' ' + dbo.Users.OtherNames AS ModifiedBy, dbo.Security.TimeModified, dbo.OfferTypeList.Description AS OfferType,
                      dbo.Security.OfferType_DPA_, dbo.Security.ParentSecurity_DPA_, dbo.Security.RequiresExtra, dbo.Security.DefaultSelection, dbo.Security.Ratio,
                      dbo.Security.MinimumQty, dbo.Security.StepQty, dbo.Security.RequiresHoldings
FROM         dbo.OrderSecType INNER JOIN
                      dbo.Security ON dbo.OrderSecType.OrderSecType_DPA_ = dbo.Security.OrderSecType_DPA_ LEFT OUTER JOIN
                      dbo.SecTransFee ON dbo.Security.Security_DPA_ = dbo.SecTransFee.Security_DPA_ LEFT OUTER JOIN
                      dbo.OfferTypeList ON dbo.Security.OfferType_DPA_ = dbo.OfferTypeList.OfferType_DPA_ LEFT OUTER JOIN
                      dbo.Users ON dbo.Security.ModifiedBy = dbo.Users.UserID
GROUP BY dbo.Security.SecurityName, dbo.Security.SecurityCode, dbo.Security.SecurityMktPrice, dbo.Security.SecurityAddr,
                      dbo.OrderSecType.OrderSecTypeDescription, dbo.Security.Security_DPA_, dbo.OrderSecType.OrderSecTypeDisplayName,
                      dbo.Security.OrderSecType_DPA_, dbo.Security.Immobilised, dbo.Security.Sector_DPA_, dbo.Security.ImportCode, dbo.Security.CanTrade,
                      dbo.Security.Offerings, dbo.Security.ExpiryDate, dbo.Security.BatchSize, dbo.Security.ClosingDate,
                      dbo.Users.Surname + N' ' + dbo.Users.OtherNames, dbo.Security.TimeModified, dbo.OfferTypeList.Description, dbo.Security.OfferType_DPA_,
                      dbo.Security.ParentSecurity_DPA_, dbo.Security.RequiresExtra, dbo.Security.DefaultSelection, dbo.Security.Ratio, dbo.Security.MinimumQty,
                      dbo.Security.StepQty, dbo.Security.RequiresHoldings
ORDER BY dbo.Security.SecurityName

GO
CREATE VIEW dbo.SecurityListOfferings
AS
SELECT     TOP 100 PERCENT dbo.SecurityListFilter.SecurityName, dbo.SecurityListFilter.SecurityCode, dbo.SecurityListFilter.SecurityMktPrice,
                      REPLACE(REPLACE(REPLACE(dbo.SecurityListFilter.SecurityAddr, CHAR(13), ','), CHAR(10), ''), CHAR(34), ' ') AS SecurityAddr,
                      dbo.SecurityListFilter.OrderSecTypeDescription, dbo.SecurityListFilter.Security_DPA_, dbo.SecurityListFilter.OrderSecTypeDisplayName,
                      dbo.SecTransFee.SecTransFeeFee, dbo.SecurityListFilter.OrderSecType_DPA_, dbo.SecurityListFilter.SecurityNameEx,
                      dbo.SecurityListFilter.Immobilised, dbo.MarketSector.ShortDescription AS Sector, dbo.SecurityListFilter.ImportCode, dbo.SecurityListFilter.CanTrade,
                      dbo.SecurityListFilter.BatchSize, dbo.SecurityListFilter.ClosingDate, dbo.SecurityListFilter.ModifiedBy, dbo.SecurityListFilter.TimeModified,
                      dbo.SecurityListFilter.OfferType, dbo.SecurityListFilter.ParentSecurity_DPA_, dbo.SecurityListFilter.OfferType_DPA_,
                      dbo.SecurityListFilter.DefaultSelection, dbo.SecurityListFilter.Ratio, dbo.SecurityListFilter.MinimumQty, dbo.SecurityListFilter.StepQty,
                      dbo.SecurityListFilter.RequiresHoldings
FROM         dbo.SecurityListFilter INNER JOIN
                      dbo.OrderSecType ON dbo.SecurityListFilter.OrderSecType_DPA_ = dbo.OrderSecType.OrderSecType_DPA_ LEFT OUTER JOIN
                      dbo.SecTransFee ON dbo.SecurityListFilter.Security_DPA_ = dbo.SecTransFee.Security_DPA_ AND
                      dbo.SecurityListFilter.MaxOfSecTransFeeADate = dbo.SecTransFee.SecTransFeeADate LEFT OUTER JOIN
                      dbo.MarketSector ON dbo.SecurityListFilter.Sector_DPA_ = dbo.MarketSector.Sector_DPA_
WHERE     (dbo.OrderSecType.OrderSecType_DPA_ = 2) AND (dbo.SecurityListFilter.Offerings = 1)
GROUP BY dbo.SecurityListFilter.SecurityName, dbo.SecurityListFilter.SecurityCode, dbo.SecurityListFilter.SecurityMktPrice, dbo.SecurityListFilter.SecurityAddr,
                      dbo.SecurityListFilter.OrderSecTypeDescription, dbo.SecurityListFilter.Security_DPA_, dbo.SecurityListFilter.OrderSecTypeDisplayName,
                      dbo.SecTransFee.SecTransFeeFee, dbo.SecurityListFilter.OrderSecType_DPA_, dbo.SecurityListFilter.SecurityNameEx,
                      dbo.SecurityListFilter.Immobilised, dbo.MarketSector.ShortDescription, dbo.SecurityListFilter.ImportCode, dbo.SecurityListFilter.CanTrade,
                      dbo.SecurityListFilter.BatchSize, dbo.SecurityListFilter.ClosingDate, dbo.SecurityListFilter.ModifiedBy, dbo.SecurityListFilter.TimeModified,
                      dbo.SecurityListFilter.OfferType, dbo.SecurityListFilter.ParentSecurity_DPA_, dbo.SecurityListFilter.OfferType_DPA_,
                      dbo.SecurityListFilter.DefaultSelection, dbo.SecurityListFilter.Ratio, dbo.SecurityListFilter.MinimumQty, dbo.SecurityListFilter.StepQty,
                      dbo.SecurityListFilter.RequiresHoldings
ORDER BY dbo.SecurityListFilter.SecurityName

GO
CREATE VIEW dbo.SecurityListSecurity
AS
SELECT     TOP 100 PERCENT dbo.SecurityListFilter.SecurityName, dbo.SecurityListFilter.SecurityCode, dbo.SecurityListFilter.SecurityMktPrice,
                      REPLACE(REPLACE(REPLACE(dbo.SecurityListFilter.SecurityAddr, CHAR(13), ','), CHAR(10), ''), CHAR(34), ' ') AS SecurityAddr,
                      dbo.SecurityListFilter.OrderSecTypeDescription, dbo.SecurityListFilter.Security_DPA_, dbo.SecurityListFilter.OrderSecTypeDisplayName,
                      dbo.SecTransFee.SecTransFeeFee, dbo.SecurityListFilter.OrderSecType_DPA_, dbo.SecurityListFilter.SecurityNameEx,
                      dbo.SecurityListFilter.Immobilised, dbo.MarketSector.ShortDescription AS Sector, dbo.SecurityListFilter.ImportCode,
                      dbo.SecurityListFilter.ExpiryDate
FROM         dbo.SecurityListFilter INNER JOIN
                      dbo.SecTransFee ON dbo.SecurityListFilter.Security_DPA_ = dbo.SecTransFee.Security_DPA_ AND
                      dbo.SecurityListFilter.MaxOfSecTransFeeADate = dbo.SecTransFee.SecTransFeeADate INNER JOIN
                      dbo.OrderSecType ON dbo.SecurityListFilter.OrderSecType_DPA_ = dbo.OrderSecType.OrderSecType_DPA_ LEFT OUTER JOIN
                      dbo.MarketSector ON dbo.SecurityListFilter.Sector_DPA_ = dbo.MarketSector.Sector_DPA_
WHERE     (dbo.OrderSecType.OrderSecType_DPA_ = 2)
GROUP BY dbo.SecurityListFilter.SecurityName, dbo.SecurityListFilter.SecurityCode, dbo.SecurityListFilter.SecurityMktPrice, dbo.SecurityListFilter.SecurityAddr,
                      dbo.SecurityListFilter.OrderSecTypeDescription, dbo.SecurityListFilter.Security_DPA_, dbo.SecurityListFilter.OrderSecTypeDisplayName,
                      dbo.SecTransFee.SecTransFeeFee, dbo.SecurityListFilter.OrderSecType_DPA_, dbo.SecurityListFilter.SecurityNameEx,
                      dbo.SecurityListFilter.Immobilised, dbo.MarketSector.ShortDescription, dbo.SecurityListFilter.ImportCode, dbo.SecurityListFilter.ExpiryDate
ORDER BY dbo.SecurityListFilter.SecurityName

GO
CREATE VIEW dbo.SecurityReport
AS
SELECT     dbo.Security.SecurityCode AS Code, dbo.Security.SecurityName AS Security, dbo.OrderSecType.OrderSecTypeDisplayName AS Type,
                      dbo.Security.SecurityMktPrice AS [Market Price]
FROM         dbo.Security INNER JOIN
                      dbo.OrderSecType ON dbo.Security.OrderSecType_DPA_ = dbo.OrderSecType.OrderSecType_DPA_

GO
CREATE VIEW dbo.SettlementReminder
AS

SELECT     dbo.LotView.LotTDate As Traded, dbo.LotView.Broker_DPA_, dbo.Broker.BrokerCode + ' ' + dbo.Broker.BrokerName AS Broker, SUM(dbo.LotView.LotGrossAmount)
                      AS GrossAmt, dbo.OrderType.OrderTypeSale,
                          (SELECT     SUM(LevyAmount) AS dd
                            FROM          LevyContract
                            WHERE      Contract_DPA_ IN
                                                       (SELECT     Contract_DPA_
                                                         FROM          Lot
                                                         WHERE      LotTDate = dbo.LotView.LotTDate AND Broker_DPA_ = dbo.LotView.Broker_DPA_)) AS TotalLevies,
                      CASE dbo.OrderType.OrderTypeSale WHEN 1 THEN SUM(dbo.LotView.LotGrossAmount) -
                          (SELECT     SUM(LevyAmount) AS dd
                            FROM          LevyContract
                            WHERE      Contract_DPA_ IN
                                                       (SELECT     Contract_DPA_
                                                         FROM          Lot
                                                         WHERE      LotTDate = dbo.LotView.LotTDate AND Broker_DPA_ = dbo.LotView.Broker_DPA_)) ELSE SUM(dbo.LotView.LotGrossAmount)
                      +
                          (SELECT     SUM(LevyAmount) AS dd
                            FROM          LevyContract
                            WHERE      Contract_DPA_ IN
                                                       (SELECT     Contract_DPA_
                                                         FROM          Lot
                                                         WHERE      LotTDate = dbo.LotView.LotTDate AND Broker_DPA_ = dbo.LotView.Broker_DPA_)) END AS NetAmt
FROM         dbo.LotView INNER JOIN
                      dbo.Broker ON dbo.LotView.Broker_DPA_ = dbo.Broker.Broker_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.LotView.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_
WHERE dbo.Broker.Broker_DPA_ NOT IN (SELECT Broker_DPA_ FROM CompanyInfoView)
GROUP BY dbo.LotView.LotTDate, dbo.Broker.BrokerCode + ' ' + dbo.Broker.BrokerName, dbo.OrderType.OrderTypeSale, dbo.LotView.Broker_DPA_


GO

CREATE VIEW dbo.SettlementSchedule
AS
SELECT DISTINCT
                      TOP 100 PERCENT dbo.Lots.LotGrossAmount, dbo.Lots.LotSlipNo, dbo.OrderType.OrderTypeDescription, dbo.Client.Client_DPA_,
                      dbo.Client.ClientName, dbo.tbOrder.OrderType_DPA_, dbo.Security.SecurityCode, dbo.Lots.LotPrice, dbo.Lots.LotQty,
                      CAST(FLOOR(CAST(dbo.Lots.LotTDate AS Float)) AS DateTime) AS LotTDate, dbo.Lots.ContractNumber, dbo.Client.IsCustodian,
                      dbo.Security.OrderSecType_DPA_, CAST(FLOOR(CAST(dbo.Contract.ContractSettlementDate AS Float)) AS DateTime) AS SettlementDate,
                      a.Contract_DPA_, dbo.Client.ClientCDSNo
FROM         dbo.tbOrder INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                      dbo.Contract INNER JOIN
                      dbo.Lots ON dbo.Contract.Contract_DPA_ = dbo.Lots.Contract_DPA_ ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.Lots.OrdDetail_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ LEFT OUTER JOIN
                          (SELECT     *
                            FROM          Intertransfer
                            WHERE      InterTransferType_DPA_ = 1) a ON dbo.Contract.Contract_DPA_ = a.Contract_DPA_
ORDER BY dbo.tbOrder.OrderType_DPA_ DESC, dbo.Lots.LotSlipNo


GO

CREATE VIEW dbo.SettlementSchedule1
AS
SELECT TOP 100 PERCENT dbo.LotList.SecurityCode, dbo.LotList.LotSlipNo, dbo.LotList.ContractNumber, dbo.Payment.PaymentAmount,
               dbo.Payment.PaymentPDate, dbo.Payment.PayType_DPA_, dbo.Client.ClientName, dbo.Client.Client_DPA_,
               dbo.LotList.BrokerCode AS OwnerName
FROM  dbo.BrokerReceiptVoucher INNER JOIN
               dbo.Payment ON dbo.BrokerReceiptVoucher.BrokerReceiptVoucher_DPA_ = dbo.Payment.BrokerReceiptVoucher_DPA_ INNER JOIN
               dbo.LotList INNER JOIN
               dbo.Contract ON dbo.LotList.Contract_DPA_ = dbo.Contract.Contract_DPA_ ON
               dbo.BrokerReceiptVoucher.BrokerReceiptVoucher_DPA_ = dbo.Contract.BrokerReceiptVoucher_DPA_ INNER JOIN
               dbo.Client ON dbo.LotList.Client_DPA_ = dbo.Client.Client_DPA_
ORDER BY dbo.Payment.PayType_DPA_


GO

CREATE VIEW dbo.SettlementSchedule2
AS
SELECT TOP 100 PERCENT dbo.LotList.SecurityCode, dbo.LotList.LotSlipNo, dbo.LotList.ContractNumber, dbo.Payment.PaymentAmount,
               dbo.Payment.PaymentPDate, dbo.Payment.PayType_DPA_, dbo.Client.ClientName, dbo.Client.Client_DPA_,
               dbo.LotList.BrokerCode AS OwnerName
FROM  dbo.Payment INNER JOIN
               dbo.Voucher ON dbo.Payment.Voucher_DPA_ = dbo.Voucher.Voucher_DPA_ INNER JOIN
               dbo.Contract ON dbo.Voucher.Voucher_DPA_ = dbo.Contract.Voucher_DPA_ INNER JOIN
               dbo.LotList ON dbo.Contract.Contract_DPA_ = dbo.LotList.Contract_DPA_ INNER JOIN
               dbo.Client ON dbo.LotList.Client_DPA_ = dbo.Client.Client_DPA_
ORDER BY dbo.Payment.PayType_DPA_


GO
CREATE VIEW dbo.SettlementSlips
AS
SELECT     dbo.LotView.ContractNumber, dbo.LotView.LotTDate, dbo.OrderType.OrderTypeDescription AS OrdDetailType, dbo.Security.SecurityName AS OrdDetailSecurity,
                      SUM(dbo.LotView.LotPrice) AS LotPrice, SUM(dbo.LotView.LotQty) AS LotQty, dbo.LotView.LotSlipNo, dbo.Broker.BrokerCode, dbo.Broker.BrokerName,
                      dbo.LotView.LotTDate AS SettlementDate, SUM(dbo.LotView.LotGrossAmount) AS LevyAmount
FROM         dbo.Broker INNER JOIN
                      dbo.LotView ON dbo.Broker.Broker_DPA_ = dbo.LotView.Broker_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.LotView.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_
WHERE     (dbo.OrderType.OrderTypeSale = 0)
GROUP BY dbo.LotView.ContractNumber, dbo.LotView.LotTDate, dbo.OrderType.OrderTypeDescription, dbo.Security.SecurityName, dbo.LotView.LotSlipNo,
                      dbo.Broker.BrokerCode, dbo.Broker.BrokerName



GO

CREATE VIEW dbo.SettlePaymentVoucher
AS
SELECT     dbo.OrderType.OrderTypeDescription AS OrdDetailType, dbo.Security.SecurityName AS OrdDetailSecurity, dbo.Client.Client_DPA_,
                      dbo.Client.ClientName, dbo.Lot.LotSlipNo, dbo.Lot.LotTDate, dbo.Lot.ContractNumber, dbo.Lot.Contract_DPA_, SUM(dbo.Lot.LotGrossAmount) -
                          (SELECT     SUM(LevyAmount) AS dd
                            FROM          LevyContract
                            WHERE      Contract_DPA_ = dbo.Lot.Contract_DPA_ AND SystemMaintained <> 12) AS NetAmount, dbo.Lot.LotGrossAmount, dbo.Lot.LotQty,
                      dbo.Lot.LotPrice, dbo.PaymentVoucherSettlementDates.SettlementDate
FROM         dbo.tbOrder INNER JOIN
                      dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.Lot ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.Lot.OrdDetail_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.PaymentVoucherSettlementDates ON dbo.Lot.Contract_DPA_ = dbo.PaymentVoucherSettlementDates.Contract_DPA_
WHERE     (dbo.OrderType.OrderTypeSale = 1) AND (dbo.Lot.Deleted = 0) AND (dbo.OrdDetail.Deleted = 0) AND (dbo.tbOrder.Deleted = 0) AND
                      (dbo.Lot.Contract_DPA_ NOT IN
                          (SELECT     Contract_DPA_
                            FROM          intertransfer
                            WHERE      InterTransferType_DPA_ = 2))
GROUP BY dbo.OrderType.OrderTypeDescription, dbo.Security.SecurityName, dbo.Client.Client_DPA_, dbo.Client.ClientName, dbo.Lot.LotSlipNo,
                      dbo.Lot.LotTDate, dbo.Lot.ContractNumber, dbo.Lot.Contract_DPA_, dbo.Lot.LotGrossAmount, dbo.Lot.LotQty, dbo.Lot.LotPrice,
                      dbo.PaymentVoucherSettlementDates.SettlementDate


GO
CREATE VIEW dbo.ShareAnnouncementTypeList
AS
SELECT     TOP 100 PERCENT dbo.ShareAnnouncementType.*
FROM         dbo.ShareAnnouncementType
ORDER BY ShareAnnouncementTypeName

GO

CREATE VIEW dbo.ShareList
AS
SELECT     TOP 100 PERCENT dbo.Security.SecurityName AS Security, dbo.Share.Share_DPA_, CONVERT(DATETIME, dbo.Share.SharePDate, 108) AS SharePDate,
                       CONVERT(DATETIME, dbo.Share.ShareClosing, 108) AS ShareClosing, CONVERT(DATETIME, dbo.Share.ShareYEnd, 108) AS ShareYEnd,
                      dbo.Share.ShareAnnouncement AS AnnouncementDate, dbo.ShareAnnouncementType.ShareAnnouncementTypeName AS Announcement
FROM         dbo.Security INNER JOIN
                      dbo.Share ON dbo.Security.Security_DPA_ = dbo.Share.Security_DPA_ INNER JOIN
                      dbo.ShareAnnouncementType ON dbo.Share.ShareAnnouncementType_DPA_ = dbo.ShareAnnouncementType.ShareAnnouncementType_DPA_
ORDER BY dbo.Security.SecurityName + ', ' + CONVERT(NVARCHAR(1000), dbo.Share.ShareAnnouncement)


GO
CREATE VIEW dbo.ShowAllSmsContracts
AS
SELECT  top 100 percent   CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime) AS LotTDate, tbOrder.Order_DPA_, OrdDetail.OrdDetail_DPA_, Lot.Contract_DPA_,
                      Lot.LotGrossAmount, SUM(LevyContract.LevyAmount) AS LevyAmount, Client.Client_DPA_, Client.ClientName, Security.SecurityCode, Lot.LotQty,
                      ROUND(SUM(Lot.LotGrossAmount) / SUM(Lot.LotQty), 2) AS LotPrice, Client.updateOnContract, Client.ClientCDSNo, Client.ClientCellTel,
                      CASE (tbOrder.OrderType_DPA_) WHEN 1 THEN Lot.LotGrossAmount + SUM(LevyContract.LevyAmount)
                      ELSE Lot.LotGrossAmount - SUM(LevyContract.LevyAmount) END AS contractAmount, tbOrder.OrderType_DPA_,
                      CAST(FLOOR(CAST(Contract.ContractSettlementDate AS float)) AS datetime) AS ContractSettlementDate
FROM         Lot INNER JOIN
                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_ INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN
                      Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
                      Commission ON Client.Commission_DPA_ = Commission.Commission_DPA_ INNER JOIN
                      Security ON OrdDetail.Security_DPA_ = Security.Security_DPA_ INNER JOIN
                      Contract ON Lot.Contract_DPA_ = Contract.Contract_DPA_
WHERE     (Lot.Deleted <> 1) AND (LevyContract.Deleted <> 1) AND (CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime) = DATEADD(d, - 3,
                      CAST(FLOOR(CAST(GETDATE() AS float)) AS smalldatetime)))
GROUP BY CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime), tbOrder.Order_DPA_, OrdDetail.OrdDetail_DPA_, Client.Client_DPA_, Client.ClientName,
                      Security.SecurityCode, Client.updateOnContract, Client.ClientCDSNo, Client.ClientCellTel, tbOrder.OrderType_DPA_,
                      CAST(FLOOR(CAST(Contract.ContractSettlementDate AS float)) AS datetime), Lot.Contract_DPA_, Lot.LotQty, Lot.LotGrossAmount
ORDER BY Client.ClientName, Security.SecurityCode
GO
CREATE VIEW dbo.ShowCompoundedContracts
AS
SELECT     dbo.Lot.LotTDate, dbo.tbOrder.Order_DPA_, dbo.OrdDetail.OrdDetail_DPA_, MIN(dbo.Lot.Contract_DPA_) AS Contract_DPA_,
                      SUM(dbo.Lot.LotGrossAmount) AS Gross, SUM(dbo.LevyContract.LevyAmount) AS Commission, dbo.Commission.MinimumSecurityCommission,
                      dbo.Commission.CommissionRate, dbo.Commission.UpperSecurityCommission, dbo.Commission.SecurityBoundary,
                      dbo.Commission.MinimumSecurityCommission AS MinCommission, ROUND(dbo.Commission.CommissionRate / 100 * SUM(dbo.Lot.LotGrossAmount),
                      2) AS LowerCommission, ROUND(dbo.Commission.UpperSecurityCommission / 100 * SUM(dbo.Lot.LotGrossAmount), 2) AS UpperCommission,
                      dbo.Client.Client_DPA_, dbo.Client.ClientName
FROM         dbo.Lot INNER JOIN
                      dbo.LevyContract ON dbo.Lot.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Commission ON dbo.Client.Commission_DPA_ = dbo.Commission.Commission_DPA_
WHERE     (dbo.Lot.Deleted <> 1) AND (dbo.LevyContract.Deleted <> 1) AND (CAST(FLOOR(CAST(dbo.Lot.LotTDate AS float)) AS datetime) = DATEADD(d, -3, CAST(FLOOR(CAST(GETDATE()
                      AS float)) AS smalldatetime))) AND (dbo.tbOrder.OrderCompounded = 1) AND (dbo.LevyContract.SystemMaintained = 11)
GROUP BY dbo.Lot.LotTDate, dbo.OrdDetail.OrdDetail_DPA_, dbo.LevyContract.LevyShortName, dbo.LevyContract.SystemMaintained,
                      dbo.Commission.MinimumSecurityCommission, dbo.Commission.CommissionRate, dbo.Commission.UpperSecurityCommission,
                      dbo.Commission.SecurityBoundary, dbo.tbOrder.Order_DPA_, dbo.Client.Client_DPA_, dbo.Client.ClientName


GO
CREATE VIEW dbo.Temp2
AS

SELECT     CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime) AS LotTDate, LevyContract.LevyContract_DPA_,
                      ShowProperCommissionsForAllCompoundedContracts.ProperCommissionRate,
                      ROUND(ShowProperCommissionsForAllCompoundedContracts.ProperCommissionRate / 100 * Lot.LotGrossAmount, 1) AS ProperCommission3
FROM         Lot INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_ INNER JOIN
                          (SELECT     cast(floor(cast(Lot.LotTDate AS float)) AS datetime) AS LotTDate, OrdDetail.OrdDetail_DPA_,
                                                   Round(CASE WHEN Commission.SecurityBoundary = 0 THEN ROUND(Commission.CommissionRate / 100 * SUM(Lot.LotGrossAmount), 2)
                                                   ELSE CASE WHEN SUM(Lot.LotGrossAmount)
                                                   <= Commission.SecurityBoundary THEN ROUND(Commission.CommissionRate / 100 * SUM(Lot.LotGrossAmount), 2)
                                                   ELSE ROUND(Commission.UpperSecurityCommission / 100 * SUM(Lot.LotGrossAmount), 2) END END / SUM(Lot.LotGrossAmount) * 100,
                                                   2) AS ProperCommissionRate
                            FROM          Lot INNER JOIN
                                                   LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_ INNER JOIN
                                                   OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                                                   tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN
                                                   Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
                                                   Commission ON Client.Commission_DPA_ = Commission.Commission_DPA_
                            WHERE      (Lot.Deleted <> 1) AND (LevyContract.Deleted <> 1) AND (tbOrder.OrderCompounded = 1) AND (LevyContract.SystemMaintained = 11)
                            GROUP BY cast(floor(cast(Lot.LotTDate AS float)) AS datetime), OrdDetail.OrdDetail_DPA_, LevyContract.LevyShortName,
                                                   LevyContract.SystemMaintained, Commission.MinimumSecurityCommission, Commission.CommissionRate,
                                                   Commission.UpperSecurityCommission, Commission.SecurityBoundary, tbOrder.Order_DPA_)
                      ShowProperCommissionsForAllCompoundedContracts ON
                      OrdDetail.OrdDetail_DPA_ = ShowProperCommissionsForAllCompoundedContracts.OrdDetail_DPA_
WHERE     (LevyContract.Deleted <> 1) AND (Lot.Deleted <> 1) AND (LevyContract.SystemMaintained = 11)
GO
CREATE VIEW dbo.ShowProperCommissionForEachLevyEntry
AS
SELECT     LevyContract.LevyContract_DPA_, ShowProperCommissionsForAllCompoundedContracts.ProperCommissionRate,
                      ROUND(ShowProperCommissionsForAllCompoundedContracts.ProperCommissionRate / 100 * Lot.LotGrossAmount, 1) AS ProperCommission3
FROM         Lot INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_ INNER JOIN
                          (SELECT     cast(floor(cast(Lot.LotTDate AS float)) AS datetime) AS LotTDate, OrdDetail.OrdDetail_DPA_,
                                                   Round(CASE WHEN Commission.SecurityBoundary = 0 THEN ROUND(Commission.CommissionRate / 100 * SUM(Lot.LotGrossAmount), 2)
                                                   ELSE CASE WHEN SUM(Lot.LotGrossAmount)
                                                   <= Commission.SecurityBoundary THEN ROUND(Commission.CommissionRate / 100 * SUM(Lot.LotGrossAmount), 2)
                                                   ELSE ROUND(Commission.UpperSecurityCommission / 100 * SUM(Lot.LotGrossAmount), 2) END END / SUM(Lot.LotGrossAmount) * 100,
                                                   2) AS ProperCommissionRate
                            FROM          Lot INNER JOIN
                                                   LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_ INNER JOIN
                                                   OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                                                   tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN
                                                   Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
                                                   Commission ON Client.Commission_DPA_ = Commission.Commission_DPA_
                            WHERE      (Lot.Deleted <> 1) AND (LevyContract.Deleted <> 1) AND (tbOrder.OrderCompounded = 1) AND (LevyContract.SystemMaintained = 11)
                            GROUP BY cast(floor(cast(Lot.LotTDate AS float)) AS datetime), OrdDetail.OrdDetail_DPA_, LevyContract.LevyShortName,
                                                   LevyContract.SystemMaintained, Commission.MinimumSecurityCommission, Commission.CommissionRate,
                                                   Commission.UpperSecurityCommission, Commission.SecurityBoundary, tbOrder.Order_DPA_)
                      ShowProperCommissionsForAllCompoundedContracts ON
                      OrdDetail.OrdDetail_DPA_ = ShowProperCommissionsForAllCompoundedContracts.OrdDetail_DPA_
WHERE     (LevyContract.Deleted <> 1) AND (Lot.Deleted <> 1) AND (LevyContract.SystemMaintained = 11)
GO
CREATE VIEW dbo.Show_ProperCommissionsForAllCompoundedContracts
AS

SELECT     cast(floor(cast(Lot.LotTDate as float)) as datetime) as LotTDate, OrdDetail.OrdDetail_DPA_,
                      Round(CASE WHEN Commission.SecurityBoundary = 0 THEN ROUND(Commission.CommissionRate / 100 * SUM(Lot.LotGrossAmount), 2)
                      ELSE CASE WHEN SUM(Lot.LotGrossAmount)
                      <= Commission.SecurityBoundary THEN ROUND(Commission.CommissionRate / 100 * SUM(Lot.LotGrossAmount), 2)
                      ELSE ROUND(Commission.UpperSecurityCommission / 100 * SUM(Lot.LotGrossAmount), 2) END END / SUM(Lot.LotGrossAmount) * 100, 2)
                      AS ProperCommissionRate
FROM         Lot INNER JOIN
                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_ INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN
                      Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
                      Commission ON Client.Commission_DPA_ = Commission.Commission_DPA_
WHERE     (Lot.Deleted <> 1) AND (LevyContract.Deleted <> 1) AND (tbOrder.OrderCompounded = 1) AND (LevyContract.SystemMaintained = 11)
GROUP BY cast(floor(cast(Lot.LotTDate AS float)) AS datetime), OrdDetail.OrdDetail_DPA_, LevyContract.LevyShortName, LevyContract.SystemMaintained, Commission.MinimumSecurityCommission,
                      Commission.CommissionRate, Commission.UpperSecurityCommission, Commission.SecurityBoundary, tbOrder.Order_DPA_




GO
CREATE VIEW dbo.ShowSMSAllContractsCompounded
AS

SELECT     TOP 100 PERCENT LotTDate, Order_DPA_, OrdDetail_DPA_, MIN(Contract_DPA_) AS Contract_DPA_, SUM(LotGrossAmount) AS LotGrossAmount,
                      SUM(LevyAmount) AS LevyAmount, Client_DPA_, ClientName, SecurityCode, SUM(LotQty) AS LotQty, ROUND(SUM(LotGrossAmount) / SUM(LotQty), 2)
                      AS LotPrice, updateOnContract, ClientCDSNo, ClientCellTel, SUM(contractAmount) AS contractAmount, OrderType_DPA_, ContractSettlementDate
FROM         (SELECT     CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime) AS LotTDate, tbOrder.Order_DPA_, OrdDetail.OrdDetail_DPA_, Lot.Contract_DPA_,
                                              Lot.LotGrossAmount, SUM(LevyContract.LevyAmount) AS LevyAmount, Client.Client_DPA_, Client.ClientName, Security.SecurityCode,
                                              Lot.LotQty, ROUND(SUM(Lot.LotGrossAmount) / SUM(Lot.LotQty), 2) AS LotPrice, Client.updateOnContract, Client.ClientCDSNo,
                                              Client.ClientCellTel, CASE (tbOrder.OrderType_DPA_) WHEN 1 THEN Lot.LotGrossAmount + SUM(LevyContract.LevyAmount)
                                              ELSE Lot.LotGrossAmount - SUM(LevyContract.LevyAmount) END AS contractAmount, tbOrder.OrderType_DPA_,
                                              CAST(FLOOR(CAST(Contract.ContractSettlementDate AS float)) AS datetime) AS ContractSettlementDate
                       FROM          dbo.Lot INNER JOIN
                                              dbo.LevyContract ON dbo.Lot.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                                              dbo.OrdDetail ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                                              dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                                              dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                                              dbo.Commission ON dbo.Client.Commission_DPA_ = dbo.Commission.Commission_DPA_ INNER JOIN
                                              dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                                              dbo.Contract ON dbo.Lot.Contract_DPA_ = dbo.Contract.Contract_DPA_ INNER JOIN
                                                  (SELECT     MAX(LotTDate) AS LastDate
                                                    FROM          Lot
                                                    WHERE      (Deleted <> 1)) LastTransactionDate ON CAST(FLOOR(CAST(dbo.Lot.LotTDate AS float)) AS datetime)
                                              = CAST(FLOOR(CAST(LastTransactionDate.LastDate AS float)) AS datetime)
                       WHERE      (dbo.Lot.Deleted <> 1) AND (dbo.LevyContract.Deleted <> 1)
                       GROUP BY CAST(FLOOR(CAST(dbo.Lot.LotTDate AS float)) AS datetime), dbo.tbOrder.Order_DPA_, dbo.OrdDetail.OrdDetail_DPA_,
                                              dbo.Client.Client_DPA_, dbo.Client.ClientName, dbo.Security.SecurityCode, dbo.Client.updateOnContract, dbo.Client.ClientCDSNo,
                                              dbo.Client.ClientCellTel, dbo.tbOrder.OrderType_DPA_, CAST(FLOOR(CAST(dbo.Contract.ContractSettlementDate AS float)) AS datetime),
                                              dbo.Lot.Contract_DPA_, dbo.Lot.LotQty,  dbo.Lot.LotGrossAmount) ShowAllSmsContracts
GROUP BY LotTDate, Order_DPA_, OrdDetail_DPA_, Client_DPA_, ClientName, SecurityCode, updateOnContract, ClientCDSNo, ClientCellTel, OrderType_DPA_,
                      ContractSettlementDate
ORDER BY ClientName, SecurityCode


GO

CREATE VIEW dbo.SilasTrades
AS
SELECT DISTINCT
               TOP 100 PERCENT dbo.Security.SecurityCode, dbo.Client.Client_DPA_ AS Code, SUM(dbo.DB_OrdDetailContractedQtyList.BalanceQty) AS Quantity,
               ISNULL(- dbo.Client.CreditLimit - (dbo.ClientBalances.CurrentBal - dbo.ClientTotal.Total), 0) AS Excess, dbo.ClientBalances.CurrentBal,
               ISNULL(dbo.ClientTotal.Total, 0) AS Total, dbo.Client.CreditLimit, dbo.Security.Security_DPA_
FROM  dbo.OrdDetail INNER JOIN
               dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
               dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
               dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
               dbo.OrderSecType ON dbo.tbOrder.OrderSecType_DPA_ = dbo.OrderSecType.OrderSecType_DPA_ INNER JOIN
               dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
               dbo.DB_OrdDetailContractedQtyList ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.DB_OrdDetailContractedQtyList.OrdDetail_DPA_ LEFT OUTER JOIN
               dbo.ClientBalances ON dbo.Client.Client_DPA_ = dbo.ClientBalances.client_DPA_ LEFT OUTER JOIN
               dbo.ClientTotal ON dbo.Client.Client_DPA_ = dbo.ClientTotal.Client_DPA_ LEFT OUTER JOIN
               dbo.DB_DataStreamPriceList ON dbo.Security.Security_DPA_ = dbo.DB_DataStreamPriceList.SecKnow_DPA
WHERE (dbo.DB_OrdDetailContractedQtyList.BalanceQty > 0) AND (dbo.tbOrder.OrderHold = 0) AND (dbo.tbOrder.OrderCanceled = 0)
GROUP BY dbo.Client.Client_DPA_, dbo.ClientBalances.CurrentBal, dbo.ClientTotal.Total, dbo.Security.SecurityCode, dbo.Client.CreditLimit,
               dbo.Security.Security_DPA_
ORDER BY dbo.Client.Client_DPA_, dbo.Security.Security_DPA_


GO
CREATE VIEW dbo.smscontract
AS
SELECT DISTINCT
                      dbo.Lot.Contract_DPA_, dbo.Lot.LotGrossAmount, dbo.Client.ClientName, dbo.Client.Client_DPA_, dbo.Client.ClientCellTel, dbo.Client.ClientCDSNo,
                      SUM(dbo.LevyContract.LevyAmount) AS LevyAmount, dbo.Lot.ContractNumber, dbo.Lot.LotTDate, dbo.Client.updateOnContract,
                      dbo.Security.SecurityCode, dbo.tbOrder.OrderType_DPA_, dbo.Contract.ContractSettlementDate, dbo.Lot.LotPrice, dbo.Lot.LotQty,
                      ISNULL(dbo.DB_DataStreamPriceList.MktClose, 0) AS SecurityMktPrice, CASE (tbOrder.OrderType_DPA_)
                      WHEN 1 THEN Lot.LotGrossAmount + SUM(LevyContract.LevyAmount) ELSE Lot.LotGrossAmount - SUM(LevyContract.LevyAmount)
                      END AS contractAmount
FROM         dbo.Lot INNER JOIN
                      dbo.OrdDetail ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.LevyContract ON dbo.Lot.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.Contract ON dbo.Lot.Contract_DPA_ = dbo.Contract.Contract_DPA_ INNER JOIN
                          (SELECT     MAX(LotTDate) AS LastDate
                            FROM          Lot
                            WHERE      (Deleted <> 1)) LastTransactionDate ON CAST(FLOOR(CAST(dbo.Lot.LotTDate AS float)) AS datetime)
                      = CAST(FLOOR(CAST(LastTransactionDate.LastDate AS float)) AS datetime) LEFT OUTER JOIN
                      dbo.DB_DataStreamPriceList ON dbo.Security.Security_DPA_ = dbo.DB_DataStreamPriceList.SecKnow_DPA
WHERE     (dbo.LevyContract.Deleted <> 1) AND (dbo.Lot.Deleted <> 1) AND (dbo.OrdDetail.Deleted <> 1) AND (dbo.tbOrder.Deleted <> 1) AND
                      (dbo.Client.Deleted <> 1)
GROUP BY dbo.Lot.Contract_DPA_, dbo.Lot.LotGrossAmount, dbo.Client.ClientName, dbo.Client.Client_DPA_, dbo.Client.ClientCellTel, dbo.Client.ClientCDSNo,
                      dbo.tbOrder.OrderType_DPA_, dbo.Lot.ContractNumber, dbo.Lot.LotTDate, dbo.Client.updateOnContract, dbo.Security.SecurityCode,
                      dbo.Contract.ContractSettlementDate, dbo.Lot.LotPrice, dbo.Lot.LotQty, ISNULL(dbo.DB_DataStreamPriceList.MktClose, 0)

GO
CREATE VIEW dbo.SmsContract1
AS
SELECT DISTINCT
                      Lot.LotGrossAmount, Client.ClientName, Client.Client_DPA_, Client.ClientCellTel, Client.ClientCDSNo, SUM(LevyContract.LevyAmount) AS LevyAmount,
                      CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime) AS LotTDate, Client.updateOnContract, Security.SecurityCode, tbOrder.OrderType_DPA_,
                      tbOrder.OrderCompounded, CASE (tbOrder.OrderType_DPA_) WHEN 1 THEN Lot.LotGrossAmount + SUM(LevyContract.LevyAmount)
                      ELSE Lot.LotGrossAmount - SUM(LevyContract.LevyAmount) END AS contractAmount
FROM         Lot INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN
                      Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_ INNER JOIN
                      Security ON OrdDetail.Security_DPA_ = Security.Security_DPA_
WHERE     (LevyContract.Deleted <> 1) AND (Lot.Deleted <> 1) AND (OrdDetail.Deleted <> 1) AND (tbOrder.Deleted <> 1) AND (Client.Deleted <> 1) AND
                      (CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS smalldatetime) = DATEADD(d, - 1, CAST(FLOOR(CAST(GETDATE() AS float)) AS smalldatetime)))
GROUP BY Lot.LotGrossAmount, Client.ClientName, Client.Client_DPA_, Client.ClientCellTel, Client.ClientCDSNo, tbOrder.OrderType_DPA_, Lot.LotTDate,
                      Client.updateOnContract, Security.SecurityCode, tbOrder.OrderCompounded
HAVING      (tbOrder.OrderCompounded <> 1)

GO
CREATE VIEW dbo.smsContract
AS
SELECT DISTINCT
                      dbo.Lot.Contract_DPA_, dbo.Lot.LotGrossAmount, dbo.Client.ClientName, dbo.Client.Client_DPA_, dbo.Client.ClientCellTel, dbo.Client.ClientCDSNo,
                      SUM(dbo.LevyContract.LevyAmount) AS LevyAmount, dbo.Lot.ContractNumber, dbo.Lot.LotTDate, dbo.Client.updateOnContract,
                      dbo.Security.SecurityCode, dbo.tbOrder.OrderType_DPA_, dbo.Contract.ContractSettlementDate, dbo.Lot.LotPrice, dbo.Lot.LotQty,
                      CASE (tbOrder.OrderType_DPA_) WHEN 1 THEN Lot.LotGrossAmount + SUM(LevyContract.LevyAmount)
                      ELSE Lot.LotGrossAmount - SUM(LevyContract.LevyAmount) END AS contractAmount
FROM         Lot INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN
                      Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_ INNER JOIN
                      Security ON OrdDetail.Security_DPA_ = Security.Security_DPA_ INNER JOIN
                      Contract ON Lot.Contract_DPA_ = Contract.Contract_DPA_ INNER JOIN
                          (SELECT     MAX(LotTDate) AS LastDate
                            FROM          Lot
                            WHERE      (Deleted <> 1)) LastTransactionDate ON CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
                      = CAST(FLOOR(CAST(LastTransactionDate.LastDate AS float)) AS datetime)
WHERE     (LevyContract.Deleted <> 1) AND (Lot.Deleted <> 1) AND (OrdDetail.Deleted <> 1) AND (tbOrder.Deleted <> 1) AND (Client.Deleted <> 1)
GROUP BY Lot.Contract_DPA_, Lot.LotGrossAmount, Client.ClientName, Client.Client_DPA_, Client.ClientCellTel, Client.ClientCDSNo, tbOrder.OrderType_DPA_,
                      Lot.ContractNumber, Lot.LotTDate, Client.updateOnContract, Security.SecurityCode, Contract.ContractSettlementDate, Lot.LotPrice, Lot.LotQty
GO
CREATE VIEW dbo.SmsDebtors
AS
SELECT     TOP 100 PERCENT Client.Client_DPA_, Client.ClientCDSNo, Client.ClientName, CurrentBalances.CurrentBal AS Balance, Client.ClientCellTel,
                      Client.updateOnDebt
FROM         Client INNER JOIN
                          (SELECT     SUM(ISNULL(StatementList.Credit - StatementList.Debit, 0)) + Client.ClientOpeningBal AS CurrentBal, StatementList.Client_DPA_
                            FROM          StatementList INNER JOIN
                                                   Client ON StatementList.Client_DPA_ = Client.Client_DPA_
                            WHERE      (Client.Deleted = 0) AND (CAST(FLOOR(CAST(StatementList.TransDate AS float)) AS datetime) <= DATEADD(d, - 7,
                                                   CAST(FLOOR(CAST(GETDATE() AS float)) AS smalldatetime)))
                            GROUP BY StatementList.Client_DPA_, Client.ClientOpeningBal) CurrentBalances ON Client.Client_DPA_ = CurrentBalances.Client_DPA_
WHERE     (Client.Deleted = 0) AND (CurrentBalances.CurrentBal < 0)
ORDER BY CurrentBalances.CurrentBal







GO
CREATE VIEW dbo.smsDebtors1
AS
SELECT     dbo.Client.Client_DPA_, dbo.Client.ClientCDSNo, dbo.Client.ClientName, dbo.Client.ClientCellTel, dbo.CurrentBalances.CurrentBal,
                      dbo.Client.updateOnDebt
FROM         dbo.CurrentBalances INNER JOIN
                      dbo.Client ON dbo.CurrentBalances.Client_DPA_ = dbo.Client.Client_DPA_
WHERE     (dbo.CurrentBalances.CurrentBal < 0) AND (dbo.Client.Deleted <> 1)

GO

CREATE VIEW dbo.StaffStatement
AS
SELECT     Entity_DPA_, CAST(FLOOR(CAST(JournalDate AS float)) AS DateTime) AS TransDate, CAST(JournalEntry_DPA_ AS Nvarchar(500)) AS Ref,
                      JournalNarrative AS Particulars, JournalEntryDebit AS Debit, JournalEntryCredit AS Credit, JournalEntryCredit AS CreditBal, 0 AS IsOpeningBalance,
                      '' AS PaymentReceiptNo, 4 AS receipttype
FROM         dbo.JournalList
WHERE     (EntityType_DPA_ = 7)




GO

CREATE VIEW dbo.StatementList
AS
SELECT     TOP 100 PERCENT COUNT(*) AS ClientTransaction_DPA_, Client_DPA_, CONVERT(SmallDateTime, LEFT(TransDate, 11)) AS TransDate, REF,
                      Particulars, Debit, Credit, CreditBal, IsOpeningBalance, PaymentReceiptNo, ReceiptType, Balance
FROM         dbo.ClientStatementTransactionList
GROUP BY Client_DPA_, CAST(TransDate AS float), TransDate, REF, Particulars, Debit, Credit, CreditBal, IsOpeningBalance, PaymentReceiptNo, ReceiptType,
                      Balance
ORDER BY IsOpeningBalance DESC, CONVERT(SmallDateTime, LEFT(TransDate, 11)), ReceiptType


GO

CREATE VIEW dbo.StatementList1
AS
SELECT     TOP 100 PERCENT COUNT(*) AS ClientTransaction_DPA_, Client_DPA_, CONVERT(SmallDateTime, LEFT(TransDate, 11)) AS TransDate, REF,
                      Particulars, Debit, Credit, CreditBal, IsOpeningBalance, PaymentReceiptNo, ReceiptType, Balance
FROM         dbo.ClientStatementTransactionList1
GROUP BY Client_DPA_, CAST(TransDate AS float), TransDate, REF, Particulars, Debit, Credit, CreditBal, IsOpeningBalance, PaymentReceiptNo, ReceiptType,
                      Balance
ORDER BY IsOpeningBalance DESC, CONVERT(SmallDateTime, LEFT(TransDate, 11)), ReceiptType




GO
CREATE VIEW dbo.StatusList AS SELECT Status.StatusDescription AS StatusDescription, Status.Status_DPA_ AS Status_DPA_
FROM Status

GO
CREATE VIEW dbo.StockWatchDataStreamPriceList
AS
SELECT     TOP 100 PERCENT LastSecurityDates.Security_DPA_, LastSecurityDates.SecurityCode, LastSecurityDates.LastDate, dbo.datastream_Market.MktDate,
                      ISNULL(dbo.datastream_Market.MktClose, 0) AS Price
FROM         (SELECT     Security.Security_DPA_, Security.SecurityCode, MAX(datastream_Market.MktDate) AS LastDate
                       FROM          Security LEFT OUTER JOIN
                                              datastream_Market ON Security.SecurityCode = datastream_Market.MktCode
                       GROUP BY Security.Security_DPA_, Security.SecurityCode) LastSecurityDates LEFT OUTER JOIN
                      dbo.datastream_Market ON LastSecurityDates.SecurityCode = dbo.datastream_Market.MktCode AND
                      LastSecurityDates.LastDate = dbo.datastream_Market.MktDate
ORDER BY LastSecurityDates.SecurityCode

GO
CREATE VIEW dbo.StockWatchList
AS
SELECT     TOP 100 PERCENT dbo.StockWatch.StockWatch_DPA_, dbo.Client.Client_DPA_, dbo.Client.ClientCDSNo, dbo.Client.ClientName,
                      dbo.Users.OtherNames + N'  ' + dbo.Users.Surname AS userName, dbo.StockWatch.Security_DPA_, datastream_SecurityPriceList.SecurityCode,
                      ISNULL(datastream_SecurityPriceList.Price, 0) AS Price, dbo.StockWatch.TimeChanged, dbo.Client.ClientCellTel
FROM         dbo.StockWatch INNER JOIN
                      dbo.Client ON dbo.StockWatch.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Users ON dbo.StockWatch.Changedby = dbo.Users.UserID INNER JOIN
                          (SELECT     LastSecurityDates.Security_DPA_, LastSecurityDates.SecurityCode, LastSecurityDates.LastDate, datastream_Market.MktDate,
                                                   ISNULL(datastream_Market.MktClose, 0) AS Price
                            FROM          (SELECT     Security.Security_DPA_, Security.SecurityCode, MAX(datastream_Market.MktDate) AS LastDate
                                                    FROM          Security LEFT OUTER JOIN
                                                                           datastream_Market ON Security.SecurityCode = datastream_Market.MktCode
                                                    GROUP BY Security.Security_DPA_, Security.SecurityCode) LastSecurityDates LEFT OUTER JOIN
                                                   datastream_Market ON LastSecurityDates.SecurityCode = datastream_Market.MktCode AND
                                                   LastSecurityDates.LastDate = datastream_Market.MktDate) datastream_SecurityPriceList ON
                      dbo.StockWatch.Security_DPA_ = datastream_SecurityPriceList.Security_DPA_
WHERE     (dbo.StockWatch.Deleted <> 1)
ORDER BY dbo.Client.Client_DPA_

GO
CREATE VIEW dbo.SubmitOfferingList
AS
SELECT     PAL_No, Client_DPA_, ClientName, SecurityName, Offering_DPA_, Offering_Price, ID_No, Alloted_Rights, Accepted_Rights, Renouncee, Submitted,
                      Submission_Date, PaymentReceiptNo, PaymentReference, Amount_Payable
FROM         dbo.OfferingsList
WHERE     (Submitted = 0)

GO
CREATE VIEW dbo.SystemEntityList
AS
SELECT     EntityName AS AccountName, Entity_DPA_ AS Account_DPA_
FROM         dbo.EntityList
WHERE     (SystemMaintained = 1)

GO

CREATE VIEW dbo.tempStockWatch
AS
SELECT     TOP 100 PERCENT dbo.StockWatch.Client_DPA_, COUNT(dbo.StockWatch.Client_DPA_) AS StoctCount, dbo.StockWatchList.Client_DPA_ AS ClientCode,
                      dbo.StockWatchList.SecurityCode, dbo.StockWatchList.Price, dbo.StockWatchList.ClientCellTel
FROM         dbo.StockWatch INNER JOIN
                      dbo.StockWatchList ON dbo.StockWatch.Client_DPA_ = dbo.StockWatchList.Client_DPA_
WHERE     (dbo.StockWatch.Deleted <> 1)
GROUP BY dbo.StockWatch.Client_DPA_, dbo.StockWatchList.Client_DPA_, dbo.StockWatchList.SecurityCode, dbo.StockWatchList.Price,
                      dbo.StockWatchList.ClientCellTel
ORDER BY dbo.StockWatch.Client_DPA_


GO
CREATE VIEW dbo.Testing_Users_MW
AS
SELECT     dbo.UserGroups.UserID, dbo.Users.OtherNames, dbo.Users.Surname, dbo.Users.Description, dbo.Users.Enabled, dbo.Groups.GroupName AS [Group Name],
                      dbo.Groups.Description AS 'Group Description', dbo.Menus.mnuCaption AS 'Screen Caption', dbo.Menus.IsReport AS 'Is Report?', dbo.MenuGroups.CanAdd,
                      dbo.MenuGroups.CanEdit, dbo.MenuGroups.CanDelete, dbo.MenuGroups.CanSort, dbo.MenuGroups.CanFilter, dbo.MenuGroups.CanSearch
FROM         dbo.UserGroups INNER JOIN
                      dbo.Users ON dbo.UserGroups.UserID = dbo.Users.UserID INNER JOIN
                      dbo.Groups ON dbo.UserGroups.GroupID = dbo.Groups.GroupID INNER JOIN
                      dbo.MenuGroups ON dbo.Groups.GroupID = dbo.MenuGroups.GroupID INNER JOIN
                      dbo.Menus ON dbo.MenuGroups.MenuID = dbo.Menus.MenuID

GO

CREATE VIEW dbo.TimeLimitList AS SELECT TimeLimit.TimeLimitAction AS TimeLimitAction, TimeLimit.TimeLimitLimDaysInt AS TimeLimitInternal, TimeLimit.TimeLimitLimDaysNSE AS TimeLimitNSE, TimeLimit.TimeLimit_DPA_ AS TimeLimit_DPA_
FROM TimeLimit


GO
CREATE VIEW dbo.TotalSalesOrderHoldings
AS
SELECT DISTINCT
                      a.TotalQty AS TotalBalance, a.Client AS ClientName, dbo.Holdings.TradeDate, dbo.Holdings.Client_DPA_, dbo.Holdings.Security_DPA_,
                      dbo.Holdings.Quantity, dbo.Holdings.AccountStatus, dbo.Holdings.BalanceFree, CASE WHEN BalanceFree = 'N' THEN 0 ELSE CONVERT(INT,
                      dbo.Holdings.Quantity) END AS QuantityExcess
FROM         dbo.Holdings INNER JOIN
                          (SELECT DISTINCT
                                                   TOP 100 PERCENT dbo.Security.SecurityCode, dbo.Client.Client_DPA_ AS Code, dbo.Client.ClientName AS Client,
                                                   SUM(dbo.DB_OrdDetailContractedQtyList.BalanceQty) AS TotalQty, dbo.Security.SecurityName AS ordDetailSecurity,
                                                   dbo.OrdDetail.Security_DPA_
                            FROM          dbo.OrdDetail INNER JOIN
                                                   dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                                                   dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                                                   dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                                                   dbo.OrderSecType ON dbo.tbOrder.OrderSecType_DPA_ = dbo.OrderSecType.OrderSecType_DPA_ INNER JOIN
                                                   dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                                                   dbo.DB_OrdDetailContractedQtyList ON
                                                   dbo.OrdDetail.OrdDetail_DPA_ = dbo.DB_OrdDetailContractedQtyList.OrdDetail_DPA_ LEFT OUTER JOIN
                                                   dbo.Agent ON dbo.Client.Agent_DPA_ = dbo.Agent.Agent_DPA_ LEFT OUTER JOIN
                                                   dbo.ClientBalances ON dbo.Client.Client_DPA_ = dbo.ClientBalances.client_DPA_ LEFT OUTER JOIN
                                                   dbo.ClientTotal ON dbo.Client.Client_DPA_ = dbo.ClientTotal.Client_DPA_ LEFT OUTER JOIN
                                                   dbo.DB_DataStreamPriceList ON dbo.Security.Security_DPA_ = dbo.DB_DataStreamPriceList.SecKnow_DPA
                            WHERE      (dbo.DB_OrdDetailContractedQtyList.BalanceQty > 0) AND (dbo.tbOrder.OrderHold = 0) AND (dbo.tbOrder.OrderCanceled = 0) AND
                                                   (dbo.tbOrder.Deleted = 0) AND (dbo.OrdDetail.Deleted = 0) AND (dbo.OrderType.OrderTypeSale = 1) AND
                                                   (dbo.OrderSecType.OrderSecTypeDisplayName = N'Security')
                            GROUP BY dbo.Client.Client_DPA_, dbo.Client.ClientName, dbo.Security.SecurityCode, dbo.OrdDetail.Security_DPA_,
                                                   dbo.Security.SecurityName
                            ORDER BY dbo.Client.Client_DPA_) a ON dbo.Holdings.Client_DPA_ = a.Code AND dbo.Holdings.Security_DPA_ = a.Security_DPA_

GO

CREATE VIEW dbo.TotalSalesOrderholdings
AS
SELECT DISTINCT
                      a.TotalQty AS TotalBalance, a.Client AS ClientName, dbo.Holdings.TradeDate, dbo.Holdings.Client_DPA_, dbo.Holdings.Security_DPA_,
                      dbo.Holdings.Quantity, dbo.Holdings.AccountStatus, dbo.Holdings.BalanceFree,
                      CASE WHEN BalanceFree = 'N' THEN 0 ELSE dbo.Holdings.Quantity END AS QuantityExcess
FROM         dbo.Holdings INNER JOIN
                          (SELECT DISTINCT
                                                   TOP 100 PERCENT dbo.Security.SecurityCode, dbo.Client.Client_DPA_ AS Code, dbo.Client.ClientName AS Client,
                                                   SUM(dbo.DB_OrdDetailContractedQtyList.BalanceQty) AS TotalQty, dbo.Security.SecurityName AS ordDetailSecurity,
                                                   dbo.OrdDetail.Security_DPA_
                            FROM          dbo.OrdDetail INNER JOIN
                                                   dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                                                   dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                                                   dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                                                   dbo.OrderSecType ON dbo.tbOrder.OrderSecType_DPA_ = dbo.OrderSecType.OrderSecType_DPA_ INNER JOIN
                                                   dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                                                   dbo.DB_OrdDetailContractedQtyList ON
                                                   dbo.OrdDetail.OrdDetail_DPA_ = dbo.DB_OrdDetailContractedQtyList.OrdDetail_DPA_ LEFT OUTER JOIN
                                                   dbo.Agent ON dbo.Client.Agent_DPA_ = dbo.Agent.Agent_DPA_ LEFT OUTER JOIN
                                                   dbo.ClientBalances ON dbo.Client.Client_DPA_ = dbo.ClientBalances.client_DPA_ LEFT OUTER JOIN
                                                   dbo.ClientTotal ON dbo.Client.Client_DPA_ = dbo.ClientTotal.Client_DPA_ LEFT OUTER JOIN
                                                   dbo.DB_DataStreamPriceList ON dbo.Security.Security_DPA_ = dbo.DB_DataStreamPriceList.SecKnow_DPA
                            WHERE      (dbo.DB_OrdDetailContractedQtyList.BalanceQty > 0) AND (dbo.tbOrder.OrderHold = 0) AND (dbo.tbOrder.OrderCanceled = 0) AND
                                                   (dbo.tbOrder.Deleted = 0) AND (dbo.OrdDetail.Deleted = 0) AND (dbo.OrderType.OrderTypeSale = 1) AND
                                                   (dbo.OrderSecType.OrderSecTypeDisplayName = N'Security')
                            GROUP BY dbo.Client.Client_DPA_, dbo.Client.ClientName, dbo.Security.SecurityCode, dbo.OrdDetail.Security_DPA_,
                                                   dbo.Security.SecurityName
                            ORDER BY dbo.Client.Client_DPA_) a ON dbo.Holdings.Client_DPA_ = a.Code AND dbo.Holdings.Security_DPA_ = a.Security_DPA_

GO

CREATE VIEW dbo.TotalTradingSchedule
AS
SELECT DISTINCT
                      TOP 100 PERCENT dbo.OrderType.OrderTypeDescription AS OrdDetailType, dbo.OrdDetail.OrdDetailPrice, dbo.OrdDetail.Order_DPA_,
                      dbo.tbOrder.OrderRef, dbo.tbOrder.OrderDate, dbo.Security.SecurityCode, dbo.Client.Client_DPA_ AS Code, dbo.Client.ClientName AS Client,
                      dbo.OrdDetail.Best, dbo.OrdDetail.OrdDetailValidity AS Validity, dbo.OrdDetail.Amount, dbo.DB_DataStreamPriceList.MktClose AS Price,
                      dbo.DB_OrdDetailContractedQtyList.BalanceQty, CASE (OrderTypeSale) WHEN 0 THEN ISNULL(dbo.ClientTotal.Total, 0)
                      ELSE Isnull(SalesOrdersTotals.Total, 0) END AS Total, Isnull(dbo.TotalSalesOrderholdings.Quantity, 0) AS Balance, CASE (OrderTypeSale)
                      WHEN 0 THEN (- isnull(dbo.Client.CreditLimit, 0) - (isnull(dbo.ClientBalances.CurrentBal, 0) - isnull(dbo.ClientTotal.Total, 0)))
                      ELSE (isnull(SalesOrdersTotals.Total, 0) - isnull(TotalSalesOrderholdings.QuantityExcess, 0)) END AS Excess, dbo.ClientBalances.CurrentBal,
                      dbo.OrderSecType.OrderSecTypeDisplayName AS OrdDetailSecType, dbo.Security.SecurityName AS ordDetailSecurity, dbo.Client.CreditLimit,
                      dbo.OrderType.OrderTypeSale, dbo.OrdDetail.Security_DPA_, dbo.Client.ClientCDSNo, dbo.Agent.AgentName, dbo.Agent.Agent_DPA_ AS AgentCode,
                      isnull(TotalSalesOrderholdings.TotalBalance, 0) AS TotalBalance, TotalSalesOrderholdings.BalanceFree, dbo.Bond.BondIssue,
                      dbo.Client.Iscustodian
FROM         dbo.OrdDetail LEFT OUTER JOIN
                      dbo.Bond ON dbo.OrdDetail.Bond_DPA_ = dbo.Bond.Bond_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                      dbo.OrderSecType ON dbo.tbOrder.OrderSecType_DPA_ = dbo.OrderSecType.OrderSecType_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.DB_OrdDetailContractedQtyList ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.DB_OrdDetailContractedQtyList.OrdDetail_DPA_ INNER JOIN
                      dbo.SalesOrdersTotals ON dbo.OrdDetail.Order_DPA_ = dbo.SalesOrdersTotals.Order_DPA_ AND
                      dbo.Security.Security_DPA_ = dbo.SalesOrdersTotals.Security_DPA_ LEFT OUTER JOIN
                      dbo.TotalSalesOrderholdings ON dbo.Client.Client_DPA_ = dbo.TotalSalesOrderholdings.Client_DPA_ AND
                      dbo.Security.Security_DPA_ = dbo.TotalSalesOrderholdings.Security_DPA_ LEFT OUTER JOIN
                      dbo.Agent ON dbo.Client.Agent_DPA_ = dbo.Agent.Agent_DPA_ LEFT OUTER JOIN
                      dbo.ClientBalances ON dbo.Client.Client_DPA_ = dbo.ClientBalances.client_DPA_ LEFT OUTER JOIN
                      dbo.ClientTotal ON dbo.Client.Client_DPA_ = dbo.ClientTotal.Client_DPA_ LEFT OUTER JOIN
                      dbo.DB_DataStreamPriceList ON dbo.Security.Security_DPA_ = dbo.DB_DataStreamPriceList.SecKnow_DPA
WHERE     (dbo.DB_OrdDetailContractedQtyList.BalanceQty > 0) AND (dbo.tbOrder.OrderHold = 0) AND (dbo.tbOrder.OrderCanceled = 0) AND
                      (dbo.tbOrder.Deleted = 0) AND (dbo.OrdDetail.Deleted = 0)


GO
CREATE VIEW dbo.TradeAffirmation
AS
SELECT     dbo.LotView.LotSlipNo, dbo.LotView.ContractNumber, dbo.LotView.LotQty, dbo.LotView.LotPrice, dbo.tbOrder.OrderRef, dbo.tbOrder.Order_DPA_,
                      dbo.Client.ClientAddr AS AccountAddress, dbo.Client.ClientName AS Account, dbo.LotView.LotTDate,
                      dbo.LotView.LotGrossAmount AS SettlementGrossAmount, ISNULL(dbo.Client.ClientContact, '') AS Owner, dbo.Client.ClientFax,
                      dbo.Security.SecurityName,
                      CASE dbo.OrderType.OrderTypeSale WHEN 0 THEN 'PURCHASE OF ' + dbo.OrdDetailList.OrdDetailSecurity ELSE 'SALE OF ' + dbo.OrdDetailList.OrdDetailSecurity
                       END AS ReferenceHeader, CASE dbo.OrderType.OrderTypeSale WHEN 1 THEN (dbo.LotView.LotGrossAmount) -
                          (SELECT     SUM((CASE SystemMaintained WHEN 12 THEN 0 - LevyAmount ELSE LevyAmount END)) AS dd
                            FROM          LevyContractList
                            WHERE      LotTDate = dbo.LotView.LotTDate AND Contract_DPA_ = dbo.LotView.Contract_DPA_ AND SystemMaintained <> 13) -
                          ((SELECT     ISNULL(MAX(LevyAmount), 0)
                              FROM         LevyContractList
                              WHERE     LotTDate = dbo.LotView.LotTDate AND Contract_DPA_ = dbo.LotView.Contract_DPA_ AND SystemMaintained = 13) /
                          (SELECT     COUNT(Contract_DPA_)
                            FROM          LotView
                            WHERE      OrdDetail_DPA_ = dbo.LotView.OrdDetail_DPA_)) ELSE (dbo.LotView.LotGrossAmount) +
                          (SELECT     SUM(LevyAmount) AS dd
                            FROM          LevyContractList
                            WHERE      LotTDate = dbo.LotView.LotTDate AND Contract_DPA_ = dbo.LotView.Contract_DPA_ AND SystemMaintained <> 13 AND
                                                   SystemMaintained <> 12) +
                          ((SELECT     MAX(LevyAmount)
                              FROM         LevyContractList
                              WHERE     LotTDate = dbo.LotView.LotTDate AND Contract_DPA_ = dbo.LotView.Contract_DPA_ AND SystemMaintained = 13) /
                          (SELECT     COUNT(Contract_DPA_)
                            FROM          LotView
                            WHERE      OrdDetail_DPA_ = dbo.LotView.OrdDetail_DPA_)) -
                          (SELECT     MAX(LevyAmount)
                            FROM          LevyContractList
                            WHERE      LotTDate = dbo.LotView.LotTDate AND Contract_DPA_ = dbo.LotView.Contract_DPA_ AND SystemMaintained = 12)
                      END AS SettlementAmount
FROM         dbo.LotView INNER JOIN
                      dbo.OrdDetail ON dbo.LotView.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                      dbo.OrdDetailList ON dbo.OrdDetailList.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_
WHERE     dbo.Client.Class_DPA_ = 3 AND dbo.OrdDetail.Deleted = 0 AND dbo.tbOrder.Deleted = 0

GO

CREATE VIEW dbo.TradeAffirmationBond
AS
SELECT     dbo.LotView.LotQty, dbo.LotView.LotPrice, dbo.tbOrder.OrderRef, dbo.tbOrder.Order_DPA_, dbo.LotView.LotTDate,
                      dbo.LotView.LotGrossAmount AS SettlementGrossAmount, dbo.OrderType.OrderTypeSale, dbo.Security.SecurityName,
                      dbo.LevyContract.LevyAmount AS Commission, dbo.Bond.BondIssue, dbo.Bond.BondIDate AS IssueDate, dbo.Bond.BondMDate AS MaturityDate,
                      dbo.Client.ClientAddr, dbo.Client.ClientFax, dbo.Client.ClientName AS Account, dbo.Broker.BrokerName, dbo.LotView.LotSlipNo,
                      dbo.Owner.OwnerFname + ' ' + dbo.Owner.OwnerLName AS AccountManager, dbo.Client.ClientContact AS Owner,
                      dbo.OrderType.OrderTypeDescription AS TradeAction, dbo.Client.Client_DPA_
FROM         dbo.LotView INNER JOIN
                      dbo.OrdDetail ON dbo.LotView.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                      dbo.LevyContract ON dbo.LotView.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.Bond ON dbo.OrdDetail.BondDescription = dbo.Bond.BondIssue INNER JOIN
                      dbo.Broker ON dbo.LotView.Broker_DPA_ = dbo.Broker.Broker_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Owner ON dbo.Client.Owner_DPA_ = dbo.Owner.Owner_DPA_
WHERE     (dbo.tbOrder.OrderCompounded = 1) AND (dbo.Security.OrderSecType_DPA_ = 1) AND (dbo.LevyContract.LevyAmount <> 0) AND (dbo.Client.Deleted = 0)
                       AND (dbo.LevyContract.Deleted = 0) AND (dbo.tbOrder.Deleted = 0) AND (dbo.OrdDetail.Deleted = 0)


GO
CREATE VIEW dbo.TradeAffirmationWithoutReturnable
AS
SELECT     dbo.LotView.LotSlipNo, dbo.LotView.ContractNumber, dbo.LotView.LotQty, dbo.LotView.LotPrice, dbo.tbOrder.OrderRef, dbo.tbOrder.Order_DPA_,
                      dbo.Client.ClientAddr AS AccountAddress, dbo.Client.ClientName AS Account, dbo.LotView.LotTDate,
                      dbo.LotView.LotGrossAmount AS SettlementGrossAmount, ISNULL(dbo.Client.ClientContact, '') AS Owner, dbo.Client.ClientFax,
                      dbo.Security.SecurityName,
                      CASE dbo.OrderType.OrderTypeSale WHEN 0 THEN 'PURCHASE OF ' + dbo.OrdDetailList.OrdDetailSecurity ELSE 'SALE OF ' + dbo.OrdDetailList.OrdDetailSecurity
                       END AS ReferenceHeader, CASE dbo.OrderType.OrderTypeSale WHEN 1 THEN (dbo.LotView.LotGrossAmount) -
                          (SELECT     SUM((CASE SystemMaintained WHEN 12 THEN 0 - LevyAmount ELSE LevyAmount END)) AS dd
                            FROM          LevyContractList
                            WHERE      LotTDate = dbo.LotView.LotTDate AND Contract_DPA_ = dbo.LotView.Contract_DPA_ AND SystemMaintained <> 13 AND
                                                   SystemMaintained <> 12) -
                          ((SELECT     ISNULL(MAX(LevyAmount), 0)
                              FROM         LevyContractList
                              WHERE     LotTDate = dbo.LotView.LotTDate AND Contract_DPA_ = dbo.LotView.Contract_DPA_ AND SystemMaintained = 13) /
                          (SELECT     COUNT(Contract_DPA_)
                            FROM          LotView
                            WHERE      OrdDetail_DPA_ = dbo.LotView.OrdDetail_DPA_)) ELSE (dbo.LotView.LotGrossAmount) +
                          (SELECT     SUM(LevyAmount) AS dd
                            FROM          LevyContractList
                            WHERE      LotTDate = dbo.LotView.LotTDate AND Contract_DPA_ = dbo.LotView.Contract_DPA_ AND SystemMaintained <> 13 AND
                                                   SystemMaintained <> 12) +
                          ((SELECT     MAX(LevyAmount)
                              FROM         LevyContractList
                              WHERE     LotTDate = dbo.LotView.LotTDate AND Contract_DPA_ = dbo.LotView.Contract_DPA_ AND SystemMaintained = 13) /
                          (SELECT     COUNT(Contract_DPA_)
                            FROM          LotView
                            WHERE      OrdDetail_DPA_ = dbo.LotView.OrdDetail_DPA_)) END AS SettlementAmount
FROM         dbo.LotView INNER JOIN
                      dbo.OrdDetail ON dbo.LotView.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                      dbo.OrdDetailList ON dbo.OrdDetailList.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_
WHERE     (dbo.Client.Class_DPA_ IN (2, 3)) AND (dbo.OrdDetail.Deleted = 0) AND (dbo.tbOrder.Deleted = 0)

GO

CREATE VIEW dbo.TradesArrears
AS
SELECT DISTINCT
                      TOP 100 PERCENT dbo.OrderType.OrderTypeDescription AS OrdDetailType, dbo.OrdDetail.OrdDetailPrice, dbo.OrdDetail.Order_DPA_,
                      dbo.tbOrder.OrderRef, dbo.tbOrder.OrderDate, dbo.Security.SecurityCode, dbo.Client.Client_DPA_ AS Code, dbo.Client.ClientName AS Client,
                      dbo.OrdDetail.Best, dbo.OrdDetail.OrdDetailValidity AS Validity, dbo.OrdDetail.Amount, dbo.DB_DataStreamPriceList.MktClose AS Price,
                      dbo.DB_OrdDetailContractedQtyList.BalanceQty, ISNULL(- dbo.Client.CreditLimit - (dbo.ClientBalances.CurrentBal - dbo.ClientTotal.Total), 0) AS Excess,
                      dbo.ClientBalances.CurrentBal, ISNULL(dbo.ClientTotal.Total, 0) AS Total, dbo.OrderSecType.OrderSecTypeDisplayName AS OrdDetailSecType,
                      dbo.Security.SecurityName AS ordDetailSecurity, dbo.Client.CreditLimit, dbo.OrderType.OrderTypeSale, dbo.OrdDetail.Security_DPA_,
                      dbo.Client.ClientCDSNo, CASE WHEN len(dbo.Agent.AgentName) > 10 THEN LEFT(dbo.Agent.AgentName, 10)
                      + '...' ELSE dbo.Agent.AgentName END AS AgentName, dbo.Agent.Agent_DPA_ AS AgentCode, Lots.ContractNumber
FROM         dbo.OrdDetail INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                      dbo.OrderSecType ON dbo.tbOrder.OrderSecType_DPA_ = dbo.OrderSecType.OrderSecType_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.DB_OrdDetailContractedQtyList ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.DB_OrdDetailContractedQtyList.OrdDetail_DPA_ LEFT OUTER JOIN
                      dbo.Agent ON dbo.Client.Agent_DPA_ = dbo.Agent.Agent_DPA_ LEFT OUTER JOIN
                      dbo.ClientBalances ON dbo.Client.Client_DPA_ = dbo.ClientBalances.client_DPA_ LEFT OUTER JOIN
                      dbo.ClientTotal ON dbo.Client.Client_DPA_ = dbo.ClientTotal.Client_DPA_ LEFT OUTER JOIN
                      dbo.DB_DataStreamPriceList ON dbo.Security.Security_DPA_ = dbo.DB_DataStreamPriceList.SecKnow_DPA INNER JOIN
                      lots ON ordDetail.ordDetail_DPA_ = lots.OrdDetail_DPA_
WHERE     (dbo.DB_OrdDetailContractedQtyList.BalanceQty > 0) AND (dbo.tbOrder.OrderHold = 0) AND (dbo.tbOrder.OrderCanceled = 0) AND
                      (dbo.tbOrder.Deleted = 0) AND (dbo.OrdDetail.Deleted = 0)

GO
CREATE VIEW dbo.TradesImports
AS
SELECT     TOP 100 PERCENT dbo._CDS_Imported_Trades_.CDSImport_DPA_, dbo.OrdDetailList.OrdDetail_DPA_, dbo.OrdDetailList.Order_DPA_,
                      dbo.OrdDetailList.OrderDate, dbo.OrdDetailList.OrderTypeSale, dbo.OrdDetailList.CDSOrderTypeSale, dbo.OrdDetailList.OrdDetailClient,
                      dbo.OrdDetailList.SecurityCode, dbo.OrdDetailList.BalanceQty, dbo._CDS_Imported_Trades_.CDSRef, dbo._CDS_Imported_Trades_.Quantity,
                      dbo._CDS_Imported_Trades_.Price, REPLACE(dbo._CDS_Imported_Trades_.ContraBrokerID, 'B', '') AS BrokerCode,
                      dbo._CDS_Imported_Trades_.ContraBrokerID AS CDSBrokerCode, dbo.OrdDetailList.OrdDetailType, dbo.OrdDetailList.OrdDetailSecType,
                      dbo.OrdDetailList.Security_DPA_, dbo.OrdDetailList.CommissionRate, dbo.OrdDetailList.OrdDetailQty, dbo.OrdDetailList.VolumeRate,
                      dbo.OrdDetailList.VolumeBoundary, dbo.OrdDetailList.MinimumCommission, dbo.OrdDetailList.CMARegulated,
                      dbo.OrdDetailList.PostImmobilisedRate, dbo.OrdDetailList.SecurityImmobilised, dbo.Broker.Broker_DPA_, dbo._CDS_Imported_Trades_.TradeTime,
                      dbo.OrdDetailList.AgentCommission, dbo.OrdDetailList.StaffCommission, CONVERT(smalldatetime,
                      SUBSTRING(dbo._CDS_Imported_Trades_.TradeDate, 3, 2) + '/' + SUBSTRING(dbo._CDS_Imported_Trades_.TradeDate, 1, 2)
                      + '/' + SUBSTRING(dbo._CDS_Imported_Trades_.TradeDate, 5, 4)) AS TradeDate, CONVERT(smalldatetime,
                      SUBSTRING(dbo._CDS_Imported_Trades_.SettlementDate, 3, 2) + '/' + SUBSTRING(dbo._CDS_Imported_Trades_.SettlementDate, 1, 2)
                      + '/' + SUBSTRING(dbo._CDS_Imported_Trades_.SettlementDate, 5, 4)) AS SettlementDate, dbo._CDS_Imported_Trades_.SettlementAmount,
                      dbo.OrdDetailList.EntityType_DPA_, dbo.OrdDetailList.IsCustodian, dbo.OrdDetailList.InterBank, dbo.OrdDetailList.Client_DPA_,
                      dbo.OrdDetailList.Class, dbo._CDS_Imported_Trades_.CommissionRate AS Commission
FROM         dbo.OrdDetailList RIGHT OUTER JOIN
                      dbo._CDS_Imported_Trades_ ON dbo.OrdDetailList.BalanceQty >= dbo._CDS_Imported_Trades_.Quantity AND
                      dbo.OrdDetailList.ClientCDSNo = dbo._CDS_Imported_Trades_.ClientPrefix + dbo._CDS_Imported_Trades_.ClientSuffix AND
                      dbo.OrdDetailList.SecurityCode = dbo._CDS_Imported_Trades_.SecurityDescription AND
                      dbo.OrdDetailList.CDSOrderTypeSale = dbo._CDS_Imported_Trades_.BuySell INNER JOIN
                      dbo.Broker ON LTRIM(RTRIM(REPLACE(dbo._CDS_Imported_Trades_.ContraBrokerID, 'B', ''))) = dbo.Broker.BrokerCode
WHERE     (dbo.OrdDetailList.OrderHold = 0)
ORDER BY dbo.OrdDetailList.SecurityCode, dbo._CDS_Imported_Trades_.CDSRef, dbo.OrdDetailList.OrdDetailClient

GO
CREATE VIEW dbo.TradingSchedule
AS
SELECT     TOP 100 PERCENT OrdDetailSecurity, OrdDetailType, OrdDetailPrice, Order_DPA_, BalanceQty, OrderRef, OrderDate, SecurityCode,
                      OrdDetailSecType,Limit
FROM         (SELECT     dbo.Client.ClientName AS OrdDetailClient, dbo.Security.SecurityCode + ' : ' + CONVERT(NVARCHAR(1000), dbo.OrdDetail.OrdDetailQty)
                                              AS OrdDetailItem, dbo.OrdDetail.OrdDetailPrice AS OrdDetailPrice, dbo.OrdDetail.OrdDetailQty AS OrdDetailQty,
                                              BalanceQty = CASE ISNULL(dbo.OrdDetailContractedQtyList.ContractQty, 1)
                                              WHEN 1 THEN dbo.OrdDetail.OrdDetailQty ELSE dbo.OrdDetail.OrdDetailQty - dbo.OrdDetailContractedQtyList.ContractQty END,
                                              dbo.OrderSecType.OrderSecTypeDisplayName AS OrdDetailSecType, dbo.OrderSecType.OrderSecType_DPA_ AS OrderSecType_DPA_,
                                              dbo.Security.SecurityName AS OrdDetailSecurity, dbo.OrderType.OrderTypeDescription AS OrdDetailType,
                                              dbo.OrdDetail.OrdDetail_DPA_ AS OrdDetail_DPA_, dbo.[OrderList].Order_DPA_, dbo.[OrderList].OrderDate, dbo.[OrderList].OrderCanceled,
                                              dbo.[OrderList].OrderHold, dbo.[OrderList].OrderRef, dbo.Security.Security_DPA_, dbo.Commission.CommissionRate,
                                              dbo.OrderType.OrderTypeSale, dbo.OrdDetail.OrdDetailCompound, CONVERT(DATETIME, dbo.OrdDetail.OrdDetailValidity, 108)
                                              AS OrdDetailValidity, dbo.OrdDetail.OrdDetailCertNo, dbo.Security.SecurityName, dbo.Client.Client_DPA_, dbo.Security.SecurityCode,dbo.OrdDetail.Limit
                       FROM          dbo.[OrderList] INNER JOIN
                                              dbo.OrderSecType ON dbo.[OrderList].OrderSecType_DPA_ = dbo.OrderSecType.OrderSecType_DPA_ INNER JOIN
                                              dbo.OrderType ON dbo.[OrderList].OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                                              dbo.Client ON dbo.[OrderList].Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                                              dbo.OrdDetail ON dbo.[OrderList].Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                                              dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                                              dbo.Commission ON dbo.Client.Commission_DPA_ = dbo.Commission.Commission_DPA_ LEFT OUTER JOIN
                                              dbo.OrdDetailContractedQtyList ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.OrdDetailContractedQtyList.OrdDetail_DPA_) innerTBL
WHERE     (innerTBL.OrderHold = 0 AND innerTBL.BalanceQty > 0)
GROUP BY OrdDetailSecurity, OrdDetailType, OrdDetailPrice, Order_DPA_, OrdDetail_DPA_, BalanceQty, OrderRef, OrderDate, SecurityCode,
                      OrdDetailSecType,Limit
ORDER BY OrdDetail_DPA_


GO

CREATE VIEW dbo.TransactionsList_Journal
AS
SELECT     LEFT(dbo.Users.OtherNames, 1) + '. ' + LEFT(dbo.Users.Surname, 3) + '. ' AS [User], dbo.Journal.TimeChanged AS [Mod Date],
                      dbo.Journal.JournalDate AS [Journal Date], dbo.FullEntityTypeList.EntityTypeName AS Entity, dbo.tblCompleteEntityList.EntityName AS Account,
                      dbo.JournalEntry.JournalEntry_DPA_ AS [Entry No], dbo.JournalEntry.JournalEntryDebit AS Debit, dbo.JournalEntry.JournalEntryCredit AS Credit,
                      CASE WHEN dbo.JournalEntry.Deleted = 1 THEN 'Y' ELSE 'N' END AS Del
FROM         dbo.JournalEntry INNER JOIN
                      dbo.tblCompleteEntityList ON dbo.JournalEntry.EntityType_DPA_ = dbo.tblCompleteEntityList.EntityType_DPA_ AND
                      dbo.JournalEntry.Entity_DPA_ = dbo.tblCompleteEntityList.Entity_DPA_ INNER JOIN
                      dbo.FullEntityTypeList ON dbo.tblCompleteEntityList.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_ INNER JOIN
                      dbo.Journal ON dbo.JournalEntry.Journal_DPA_ = dbo.Journal.Journal_DPA_ INNER JOIN
                      dbo.Users ON dbo.Journal.ChangedBy = dbo.Users.UserID
WHERE     (dbo.Journal.JournalCommitted = 1) AND (dbo.Journal.Released = 1)


GO

CREATE VIEW dbo.TransactionsList_Payment
AS
SELECT     TOP 100 PERCENT LEFT(dbo.Users.OtherNames, 1) + '. ' + LEFT(dbo.Users.Surname, 3) + '.' AS [User], dbo.Payment.TimeChanged AS [Mod Date],
                      dbo.Payment.PaymentPDate AS [Payment Date], dbo.Payment.Payment_DPA_ AS [Payment No],
                      'Ref: ' + dbo.Payment.PaymentReference + ' Narr: ' + dbo.Payment.PaymentNarrative AS Narrative, dbo.Payment.PaymentAmount AS Amount,
                      dbo.PaymentTypes.Description AS Mode, '... ' + RIGHT(tblCompleteEntityList_1.EntityName, 23) AS Bank,
                      CASE WHEN dbo.Payment.Deleted = 1 THEN 'Y' ELSE 'N' END AS Del
FROM         dbo.Payment INNER JOIN
                      dbo.tblCompleteEntityList ON dbo.Payment.EntityType_DPA_ = dbo.tblCompleteEntityList.EntityType_DPA_ AND
                      dbo.Payment.Entity_DPA_ = dbo.tblCompleteEntityList.Entity_DPA_ INNER JOIN
                      dbo.tblCompleteEntityList tblCompleteEntityList_1 ON dbo.Payment.BankAccount_DPA_ = tblCompleteEntityList_1.Entity_DPA_ INNER JOIN
                      dbo.PaymentTypes ON dbo.Payment.PaymentTypes_DPA_ = dbo.PaymentTypes.PaymentTypes_DPA_ INNER JOIN
                      dbo.Users ON dbo.Payment.ChangedBy = dbo.Users.UserID
WHERE     (tblCompleteEntityList_1.EntityType_DPA_ = 5) AND (dbo.Payment.PayType_DPA_ = 2)
ORDER BY dbo.Payment.Payment_DPA_


GO

CREATE VIEW dbo.TransactionsList_Receipt
AS
SELECT     TOP 100 PERCENT LEFT(dbo.Users.OtherNames, 1) + '. ' + LEFT(dbo.Users.Surname, 3) + '.' AS [User], dbo.Payment.TimeChanged AS [Mod Date],
                      dbo.Payment.PaymentPDate AS [Payment Date], dbo.Payment.PaymentReceiptNo AS [Receipt No],
                      'Ref: ' + dbo.Payment.PaymentReference + ' Narr: ' + dbo.Payment.PaymentNarrative AS Narrative, dbo.Payment.PaymentAmount AS Amount,
                      dbo.PaymentTypes.Description AS Mode, '... ' + RIGHT(tblCompleteEntityList_1.EntityName, 23) AS Bank,
                      CASE WHEN dbo.Payment.Deleted = 1 THEN 'Y' ELSE 'N' END AS Del
FROM         dbo.Payment INNER JOIN
                      dbo.tblCompleteEntityList ON dbo.Payment.EntityType_DPA_ = dbo.tblCompleteEntityList.EntityType_DPA_ AND
                      dbo.Payment.Entity_DPA_ = dbo.tblCompleteEntityList.Entity_DPA_ INNER JOIN
                      dbo.tblCompleteEntityList tblCompleteEntityList_1 ON dbo.Payment.BankAccount_DPA_ = tblCompleteEntityList_1.Entity_DPA_ INNER JOIN
                      dbo.PaymentTypes ON dbo.Payment.PaymentTypes_DPA_ = dbo.PaymentTypes.PaymentTypes_DPA_ INNER JOIN
                      dbo.Users ON dbo.Payment.ChangedBy = dbo.Users.UserID
WHERE     (tblCompleteEntityList_1.EntityType_DPA_ = 5) AND (NOT (dbo.Payment.PaymentReceiptNo IS NULL))


GO

CREATE VIEW dbo.TransactionsList_Trades
AS
SELECT     TOP 100 PERCENT dbo.Lot.TimeChanged AS [Mod.], dbo.Lot.LotTDate AS Traded, dbo.Lot.ContractSettlementDate AS [Settle.],
                      dbo.Client.Client_DPA_ AS Code, LEFT(dbo.Client.ClientName, 6) + ' ... ' + RIGHT(dbo.Client.ClientName, 10) AS Client,
                      ISNULL(LTRIM(dbo.Agent.Agent_DPA_), '') + ': ' + LEFT(dbo.Agent.AgentName, 7) AS Agent, dbo.Lot.Contract_DPA_ AS [Cont.],
                      CASE WHEN tbOrder.OrderSecType_DPA_ = 1 THEN 'B: ' + isnull(OrdDetail.BondDescription, '') ELSE Security.SecurityCode END AS Sec,
                      dbo.Lot.LotQty AS Qty, dbo.Lot.LotPrice AS Price,
                      CASE WHEN tbOrder.OrderSecType_DPA_ = 1 THEN Lot.LotQty * Lot.LotPrice / 100 ELSE Lot.LotQty * Lot.LotPrice END AS [Cons.],
                      Returnable.LevyAmount AS [Ret.], CASE WHEN Client.IsCustodian = 1 THEN 'Y' ELSE 'N' END AS [Cust.], tblContractCharges.Charges AS [Charg.],
                      CASE WHEN dbo.Lot.Deleted = 1 THEN 'Y' ELSE 'N' END AS [Del.]
FROM         dbo.Lot INNER JOIN
                      dbo.OrdDetail ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                          (SELECT     Contract_DPA_, SUM(LevyAmount) AS Charges
                            FROM          LevyContract
                            WHERE      (Deleted <> 1) AND (SystemMaintained <> 12)
                            GROUP BY Contract_DPA_) tblContractCharges ON dbo.Lot.Contract_DPA_ = tblContractCharges.Contract_DPA_ LEFT OUTER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ LEFT OUTER JOIN
                      dbo.Agent ON dbo.Client.Agent_DPA_ = dbo.Agent.Agent_DPA_ LEFT OUTER JOIN
                          (SELECT     Contract_DPA_, LevyContract_DPA_, LevyAmount, SystemMaintained
                            FROM          LevyContract
                            WHERE      (Deleted <> 1) AND (SystemMaintained = 12) AND (LevyAmount <> 0)) Returnable ON
                      dbo.Lot.Contract_DPA_ = Returnable.Contract_DPA_
ORDER BY dbo.Lot.Contract_DPA_


GO

CREATE VIEW dbo.TrialLeviesReport
AS
SELECT DISTINCT
                      TOP 100 PERCENT dbo.Contract.Contract_DPA_, dbo.Contract.ContractTransferNo, dbo.Contract.ContractDeliveryDate, dbo.Lot.Lot_DPA_,
                      dbo.Lot.LotSlipNo, CAST(FLOOR(CAST(dbo.Lot.LotTDate AS float)) AS datetime) AS TransDate, dbo.Lot.LotQty, dbo.Lot.LotGrossAmount AS LotPrice,
                      dbo.Lot.ContractNumber, dbo.Contract.ContractDelivered, dbo.Contract.ContractNCertificate, dbo.Contract.ContractNCDate,
                      dbo.Contract.ContractNCDelivered, dbo.Contract.Voucher_DPA_, dbo.Contract.ContractVouchered, dbo.LevyContract.LevyAmount,
                      dbo.LevyContract.LevyName, dbo.tbOrder.OrderType_DPA_ AS OrdDetailType, dbo.Broker.BrokerCode, dbo.Client.ClientName AS OrdDetailClient,
                      dbo.Security.SecurityName AS OrdDetailSecurity, dbo.Security.OrderSecType_DPA_ AS OrdDetailSecType, dbo.Security.SecurityCode,
                      dbo.LevyContract.SystemMaintained, dbo.LevyContract.LevyShortName
FROM         dbo.Contract INNER JOIN
                      dbo.Lot ON dbo.Contract.Contract_DPA_ = dbo.Lot.Contract_DPA_ INNER JOIN
                      dbo.LevyContract ON dbo.Contract.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Broker ON dbo.Lot.Broker_DPA_ = dbo.Broker.Broker_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_
WHERE     (dbo.Client.Deleted = 0) AND (dbo.tbOrder.Deleted = 0) AND (dbo.LevyContract.Deleted = 0) AND (dbo.Contract.Deleted = 0) AND (dbo.Lot.Deleted = 0)
                      AND (dbo.OrdDetail.Deleted = 0) AND (dbo.Client.IsCustodian = 0)
ORDER BY CAST(FLOOR(CAST(dbo.Lot.LotTDate AS float)) AS datetime), dbo.Security.SecurityCode


GO

CREATE VIEW dbo.TrialLeviesReportByDate
AS
SELECT DISTINCT
                      TOP 100 PERCENT CAST(FLOOR(CAST(TransDate AS float)) AS datetime) AS TransDate, SUM(LotPrice) AS Gross, SUM(LevyAmount) AS Commission,
                      SystemMaintained, LevyShortName
FROM         dbo.TrialLeviesReport
GROUP BY CAST(FLOOR(CAST(TransDate AS float)) AS datetime), SystemMaintained, LevyShortName
ORDER BY CAST(FLOOR(CAST(TransDate AS float)) AS datetime)


GO


CREATE VIEW [dbo].[TurnOverDetails]
AS
SELECT     dbo.tblTurnOver.Turnover_DPA_, dbo.tblTurnOver.TradeDate, dbo.tblTurnOver.MarketTurnOver,
                      dbo.Users.OtherNames + ' ' + dbo.Users.Surname AS CreatedBy, dbo.tblTurnOver.TimeCreated,
                      Users_1.OtherNames + ' ' + Users_1.Surname AS ChangedBy, dbo.tblTurnOver.TimeChanged
FROM         dbo.tblTurnOver INNER JOIN
                      dbo.Users ON dbo.tblTurnOver.CreatedBy = dbo.Users.UserID LEFT OUTER JOIN
                      dbo.Users AS Users_1 ON dbo.tblTurnOver.ChangedBy = Users_1.UserID




GO
CREATE VIEW dbo.UnactivatedClientList
AS
SELECT     dbo.FullClientList.*
FROM         dbo.FullClientList
WHERE     (ClientCDSNo IS NULL)

GO

CREATE VIEW dbo.UnSubmitOfferingList
AS
SELECT     PAL_No, Client_DPA_, ClientName, SecurityName, Offering_DPA_, Offering_Price, ID_No, Alloted_Rights, Accepted_Rights, Renouncee, Submitted,
                      Submission_Date, PaymentReceiptNo, PaymentReference, Amount_Payable
FROM         dbo.OfferingsList
WHERE     (Submitted = 1)


GO
CREATE VIEW dbo.UserEntityList
AS
SELECT     dbo.Entity.*
FROM         dbo.Entity
WHERE     (SystemMaintained = 0)

GO
CREATE VIEW dbo.UserList
AS
SELECT     UserID, OtherNames + '  ' + Surname AS [USER], UserName, Description
FROM         dbo.Users

GO

CREATE VIEW dbo.UsersGroups
AS
SELECT     dbo.UserGroups.UserID, dbo.Groups.GroupName
FROM         dbo.UserGroups INNER JOIN
                      dbo.Groups ON dbo.UserGroups.GroupID = dbo.Groups.GroupID


GO
CREATE VIEW dbo.VoucherList
AS
SELECT     dbo.Voucher.Voucher_DPA_, SUM(dbo.ContractList.ContractAmount) AS VoucherAmount, dbo.Voucher.VoucherDate AS VoucherDate,
                      dbo.Voucher.VoucherPaid, dbo.ContractList.BrokerCode, dbo.BrokerList.BrokerName, dbo.BrokerList.Broker_DPA_
FROM         dbo.ContractList INNER JOIN
                      dbo.BrokerList ON dbo.ContractList.BrokerCode = dbo.BrokerList.BrokerCode INNER JOIN
                      dbo.Voucher ON dbo.ContractList.Voucher_DPA_ = dbo.Voucher.Voucher_DPA_
GROUP BY dbo.Voucher.Voucher_DPA_, dbo.Voucher.VoucherDate, dbo.ContractList.BrokerCode, dbo.BrokerList.BrokerName, dbo.BrokerList.Broker_DPA_,
                      dbo.Voucher.VoucherPaid

GO

CREATE VIEW dbo.Wanjau_ContractSchedule
AS
SELECT     dbo.Lot.Contract_DPA_, dbo.Lot.LotPrice, dbo.Lot.LotQty, dbo.Lot.LotGrossAmount, dbo.Client.Client_DPA_, dbo.Client.ClientName,
                      dbo.OrderType.OrderTypeDescription, dbo.Lot.LotTDate, dbo.Lot.LotSlipNo, dbo.Security.Security_DPA_, dbo.Security.SecurityCode
FROM         dbo.Lot INNER JOIN
                      dbo.Contract ON dbo.Lot.Contract_DPA_ = dbo.Contract.Contract_DPA_ INNER JOIN
                      dbo.OrdDetail ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_
WHERE     (dbo.Lot.LotTDate > CONVERT(DATETIME, '2005-03-17 00:00:00', 102)) AND (dbo.Lot.LotTDate < CONVERT(DATETIME, '2005-03-18 00:00:00', 102))


GO

CREATE VIEW dbo.Wanjau_ContractsDescending
AS
SELECT     TOP 100 PERCENT dbo.Client.ClientCDSNo, dbo.Client.ClientName, dbo.tbOrder.Order_DPA_, dbo.Lot.Contract_DPA_, dbo.LevyContract.LevyName,
                      dbo.LevyContract.LevyAmount, dbo.LevyContract.LevyRate, dbo.Lot.LotTDate, dbo.Security.SecurityCode, dbo.Lot.LotPrice, dbo.Lot.LotQty,
                      dbo.LevyContract.LevyRatePercentage
FROM         dbo.OrdDetail INNER JOIN
                      dbo.Lot ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.Lot.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.LevyContract ON dbo.Lot.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_
WHERE     (dbo.Lot.LotTDate > CONVERT(DATETIME, '2005-03-08 00:00:00', 102)) AND (dbo.LevyContract.LevyName LIKE N'%agent%' OR
                      dbo.LevyContract.LevyName LIKE N'%brok%') AND (dbo.Lot.Contract_DPA_ = 1476)
ORDER BY dbo.Lot.Contract_DPA_ DESC, dbo.LevyContract.LevyName


GO

CREATE VIEW dbo.Wanjau_DataStreamPriceQueries
AS
SELECT     TOP 100 PERCENT dbo.datastream_Securities.SecNameShort, dbo.datastream_Market.MktDate, dbo.datastream_Market.MktClose,
                      dbo.datastream_Market.MktVolume
FROM         dbo.datastream_Securities INNER JOIN
                      dbo.datastream_Market ON dbo.datastream_Securities.SecCode = dbo.datastream_Market.MktCode
WHERE     (NOT (dbo.datastream_Market.MktClose IS NULL))
ORDER BY dbo.datastream_Securities.SecNameShort, dbo.datastream_Market.MktDate DESC


GO

CREATE VIEW dbo.Wanjau_Misc1
AS
SELECT     TOP 100 PERCENT dbo.Client.ClientCDSNo, dbo.Client.ClientName, dbo.tbOrder.Order_DPA_, dbo.Lot.Contract_DPA_, dbo.LevyContract.LevyName,
                      dbo.LevyContract.LevyAmount, dbo.LevyContract.LevyRate, dbo.Lot.LotTDate, dbo.Security.SecurityCode, dbo.Lot.LotPrice, dbo.Lot.LotQty,
                      dbo.Client.Client_DPA_
FROM         dbo.OrdDetail INNER JOIN
                      dbo.Lot ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.Lot.OrdDetail_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.LevyContract ON dbo.Lot.Contract_DPA_ = dbo.LevyContract.Contract_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_
WHERE     (dbo.LevyContract.LevyName LIKE N'%br%' OR
                      dbo.LevyContract.LevyName LIKE N'%ag%') AND (dbo.tbOrder.Order_DPA_ = 1002)
ORDER BY dbo.Lot.Contract_DPA_ DESC, dbo.LevyContract.LevyName


GO

CREATE VIEW dbo.Wanjau_Misc2
AS
SELECT TOP 100 PERCENT dbo.MenuGroups.ID, dbo.MenuGroups.GroupID, dbo.Menus.MenuID, dbo.Menus.mnuCaption, dbo.MenuGroups.CanAdd,
               dbo.MenuGroups.CanEdit, dbo.MenuGroups.CanDelete, dbo.MenuGroups.CanSort, dbo.MenuGroups.CanFilter, dbo.MenuGroups.CanSearch,
               dbo.Groups.Description, dbo.Menus.IsReport
FROM  dbo.MenuGroups INNER JOIN
               dbo.Menus ON dbo.MenuGroups.MenuID = dbo.Menus.MenuID INNER JOIN
               dbo.Groups ON dbo.MenuGroups.GroupID = dbo.Groups.GroupID
WHERE (dbo.MenuGroups.GroupID = 11) AND (dbo.Menus.IsReport = 0)
ORDER BY dbo.Menus.MenuID


GO

CREATE VIEW dbo.Wanjau_Misc3
AS
SELECT     dbo.Users.UserName, dbo.AuditTrail.UserID, dbo.AuditTrail.AuditTrail_DPA_, dbo.AuditTrail.AuditTrailAction, dbo.AuditTrail.AuditTrailMoment,
                      dbo.AuditTrailItem.AuditTrailItemField, dbo.AuditTrailItem.AuditTrailItemValue
FROM         dbo.Users INNER JOIN
                      dbo.AuditTrail ON dbo.Users.UserID = dbo.AuditTrail.UserID RIGHT OUTER JOIN
                      dbo.AuditTrailItem ON dbo.AuditTrail.AuditTrail_DPA_ = dbo.AuditTrailItem.AuditTrail_DPA_
WHERE     (dbo.AuditTrailItem.AuditTrailItemField = N'Order_DPA_') AND (dbo.AuditTrailItem.AuditTrailItemValue LIKE N'%1279%')


GO

CREATE VIEW dbo.Wanjau_Misc4
AS
SELECT     TOP 100 PERCENT dbo.Lot.Contract_DPA_, dbo.Payment.Contract_DPA_ AS Payment_Contract_DPA_, dbo.Lot.LotPrice, dbo.Lot.LotQty,
                      dbo.Lot.LotGrossAmount, dbo.Lot.LotTDate, dbo.Payment.PaymentPDate, dbo.Payment.Payment_DPA_, dbo.Payment.BankAccount_DPA_,
                      dbo.Payment.Entity_DPA_, dbo.Payment.EntityType_DPA_
FROM         dbo.Lot INNER JOIN
                      dbo.OrdDetail ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.Payment ON dbo.Lot.LotGrossAmount = dbo.Payment.PaymentAmount
WHERE     (dbo.Payment.Contract_DPA_ IS NOT NULL) AND (dbo.Payment.Contract_DPA_ <> 0)
ORDER BY dbo.Lot.Contract_DPA_


GO

CREATE VIEW dbo.Wanjau_Misc6
AS
SELECT     TOP 100 PERCENT dbo.Lot.Contract_DPA_, dbo.Payment.Contract_DPA_ AS Payment_Contract_DPA_, dbo.Lot.LotPrice, dbo.Lot.LotQty,
                      dbo.Lot.LotGrossAmount, dbo.Lot.LotTDate, dbo.Payment.PaymentPDate, dbo.Payment.Payment_DPA_, dbo.Payment.BankAccount_DPA_,
                      dbo.Payment.Entity_DPA_, dbo.Payment.EntityType_DPA_
FROM         dbo.Lot INNER JOIN
                      dbo.OrdDetail ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.Payment ON dbo.Lot.LotGrossAmount = dbo.Payment.PaymentAmount
WHERE     (dbo.Payment.Contract_DPA_ IS NOT NULL) AND (dbo.Payment.Contract_DPA_ <> 0)
ORDER BY dbo.Lot.Contract_DPA_


GO

CREATE VIEW dbo.Wanjau_Misc7
AS
SELECT     dbo.InterTransfer.Contract_DPA_, dbo.Lot.LotTDate, dbo.InterTransfer.InterTransfer_DPA_, dbo.InterTransfer.TransferDate,
                      dbo.InterTransfer.InterTransferType_DPA_
FROM         dbo.InterTransfer INNER JOIN
                      dbo.Lot ON dbo.InterTransfer.Contract_DPA_ = dbo.Lot.Contract_DPA_ AND CONVERT(int, dbo.InterTransfer.TransferDate) = CONVERT(int,
                      dbo.Lot.LotTDate)
WHERE     (dbo.InterTransfer.InterTransferType_DPA_ = 2)


GO

CREATE VIEW dbo.Wanjau_Misc8
AS
SELECT     dbo.Client.Client_DPA_, dbo.Client.ClientName, dbo.tbOrder.Order_DPA_, dbo.Security.SecurityName, dbo.OrdDetail.OrdDetailQty,
                      dbo.OrdDetail.OrdDetailPrice
FROM         dbo.tbOrder INNER JOIN
                      dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_
WHERE     (dbo.Client.Client_DPA_ = 100552)


GO

CREATE VIEW dbo.Wanjau_PendingOrders
AS
SELECT DISTINCT
                      TOP 100 PERCENT dbo.LotList.Client_DPA_, dbo.LotList.SecurityCode, dbo.LotList.BalanceQty, dbo.LotList.Order_DPA_, dbo.Client.ClientName,
                      dbo.Client.ClientCDSNo, dbo.LotList.OrderTypeSale, dbo.Agent.AgentName, dbo.LotList.CommissionRate
FROM         dbo.Client INNER JOIN
                      dbo.LotList ON dbo.Client.Client_DPA_ = dbo.LotList.Client_DPA_ INNER JOIN
                      dbo.Agent ON dbo.Client.Agent_DPA_ = dbo.Agent.Agent_DPA_
WHERE     (dbo.LotList.BalanceQty > 0)
ORDER BY dbo.LotList.Client_DPA_, dbo.LotList.SecurityCode


GO

CREATE VIEW dbo.Wanjau_UnimportedNominalAccount
AS
SELECT     dbo.AccountImport.ID, dbo.AccountImport.NAME, dbo.AccountImport.Balance, dbo.Account.Account_DPA_
FROM         dbo.AccountImport LEFT OUTER JOIN
                      dbo.Account ON dbo.AccountImport.ID = dbo.Account.AccountCode
WHERE     (dbo.Account.Account_DPA_ IS NULL)


GO

CREATE VIEW dbo.Wanjaus_TradingSchedule
AS
SELECT DISTINCT
                      TOP 100 PERCENT dbo.OrderType.OrderTypeDescription AS OrdDetailType, dbo.OrdDetail.OrdDetailPrice, dbo.OrdDetail.Order_DPA_,
                      dbo.tbOrder.OrderRef, dbo.tbOrder.OrderDate, dbo.Security.SecurityCode, dbo.Client.Client_DPA_ AS Code, dbo.Client.ClientName AS Client,
                      dbo.OrdDetail.Best, dbo.OrdDetail.OrdDetailValidity AS Validity, dbo.OrdDetail.Amount, dbo.DB_DataStreamPriceList.MktClose AS Price,
                      dbo.DB_OrdDetailContractedQtyList.BalanceQty, ISNULL(- dbo.Client.CreditLimit - dbo.ClientBalances.CurrentBal - dbo.ClientTotal.Total, 0) AS Excess,
                      dbo.ClientBalances.CurrentBal, ISNULL(dbo.ClientTotal.Total, 0) AS Total, dbo.OrderSecType.OrderSecTypeDisplayName AS OrdDetailSecType,
                      dbo.Security.SecurityName AS ordDetailSecurity, dbo.Client.CreditLimit
FROM         dbo.OrdDetail INNER JOIN
                      dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN
                      dbo.OrderSecType ON dbo.tbOrder.OrderSecType_DPA_ = dbo.OrderSecType.OrderSecType_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.DB_OrdDetailContractedQtyList ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.DB_OrdDetailContractedQtyList.OrdDetail_DPA_ LEFT OUTER JOIN
                      dbo.ClientBalances ON dbo.Client.Client_DPA_ = dbo.ClientBalances.client_DPA_ LEFT OUTER JOIN
                      dbo.ClientTotal ON dbo.Client.Client_DPA_ = dbo.ClientTotal.Client_DPA_ LEFT OUTER JOIN
                      dbo.DB_DataStreamPriceList ON dbo.Security.Security_DPA_ = dbo.DB_DataStreamPriceList.SecKnow_DPA
WHERE     (dbo.DB_OrdDetailContractedQtyList.BalanceQty > 0) AND (dbo.tbOrder.OrderHold = 0) AND (dbo.tbOrder.OrderCanceled = 0)


GO

CREATE VIEW dbo.WebtbOrderList
AS
SELECT     dbo.WebtbOrder.Order_DPA_, dbo.WebtbOrder.OrderDate, dbo.OrderTypeList.OrderTypeName, dbo.WebtbOrder.OrderRef, dbo.Client.ClientName,
                      dbo.SecurityList.SecurityCode, dbo.WebtbOrder.OrdDetailQty, dbo.WebtbOrder.OrdDetailPrice, dbo.WebtbOrder.OrdDetailValidity,
                      dbo.WebtbOrder.Accepted, dbo.WebtbOrder.ApprovalAction, dbo.WebtbOrder.Reason, dbo.WebtbOrder.ActionDate, dbo.WebtbOrder.ActionTime,
                      dbo.WebtbOrder.UserName, dbo.Client.Client_DPA_, dbo.Client.ClientEmail
FROM         dbo.WebtbOrder INNER JOIN
                      dbo.OrderTypeList ON dbo.WebtbOrder.OrderType_DPA_ = dbo.OrderTypeList.OrderType_DPA_ INNER JOIN
                      dbo.Client ON dbo.WebtbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN
                      dbo.SecurityList ON dbo.WebtbOrder.Security_DPA_ = dbo.SecurityList.Security_DPA_
WHERE     (dbo.WebtbOrder.Accepted = 0)


GO
CREATE VIEW dbo.WWCDSSettlements
AS

SELECT  TOP 100 PERCENT    Contract.ContractSettlementDate, Lot.LotTDate, Client.Client_DPA_, Client.ClientName, Lot.LotSlipNo, Lot.ContractNumber, Security.SecurityCode,
                      Lot.LotQty, Lot.LotPrice, CASE WHEN tbOrder.OrderType_DPA_ = 2 THEN Lot.LotGrossAmount ELSE - Lot.LotGrossAmount END AS Gross
FROM         Contract INNER JOIN
                      Lot ON Contract.Contract_DPA_ = Lot.Contract_DPA_ INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN
                      Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
                      Security ON OrdDetail.Security_DPA_ = Security.Security_DPA_
WHERE     (CAST(FLOOR(CAST(Contract.ContractSettlementDate AS float)) AS datetime) = CONVERT(DATETIME, '2005-12-30 00:00:00', 102)) AND
                      (tbOrder.IsCustodian <> 1)
ORDER BY Lot.LotSlipNo
GO

CREATE VIEW dbo.ZeroHoldings
AS
SELECT DISTINCT dbo.Client.ClientName AS ClientName, dbo.Client.Client_DPA_ AS Client_DPA_, dbo.Client.ClientCDSNo
FROM         dbo.Client LEFT OUTER JOIN
                      dbo.Holdings ON dbo.Client.Client_DPA_ = dbo.Holdings.Client_DPA_
WHERE     (dbo.Holdings.Client_DPA_ IS NULL)


GO

CREATE FUNCTION cont_Round05 (@Amount float)
RETURNS money
AS
BEGIN

Declare @Round1 float
Declare @Round2 float
Declare @Diff1 float
Declare @Check1 float
Declare @Check2 float
Declare @Check3 float
Declare @CheckSum float
Declare @Result money

set @Round1 = ROUND(@Amount*100,0)
set @Round2 = FLOOR(@Amount*10)*10
set @Diff1 = @Round1-@Round2
set @Check1 = (case when @Diff1<=2 then 0 else 0 end)
set @Check2 = (case when @Diff1>=3 and @Diff1<=7  then 5 else 0 end)
set @Check3 = (case when @Diff1>=8 then 10 else 0 end)
set @CheckSum = @Check1+@Check2+@Check3
set @Result = round((@Round2+@CheckSum)/100,2)

	RETURN(@Result)
END


GO
	CREATE Procedure [dbo].[AA_Comm_By_Institution_Daily]
		(@Date as Datetime)
	AS
	--Declare @Date AS DATETIME
	--Set @Date = '21 Dec 2011'
	Select
			InstitutionName, ResidencyDescription, BrokerComm, AgentComm, NetComm, LotTDate
		from
			[AA_View_Comm_By_Institution]
		Where
			LotTDate = @Date
GO
	CREATE Procedure [dbo].[AA_Comm_By_Institution_Monthly]
	(@MonthStartDate as Datetime, @MonthToDate as Datetime)
	AS
	--Declare @MonthStartDate AS DATETIME
	--Declare @MonthToDate AS DATETIME
	--Set @MonthStartDate = '01 Feb 2012'
	--Set @MonthToDate = '22 Feb 2012'

	Select
			InstitutionName, ResidencyDescription, Sum(BrokerComm)BrokerComm, Sum(AgentComm)AgentComm, Sum(NetComm) NetComm
		from
			[AA_View_Comm_By_Institution]
		Where
			LotTDate BETWEEN @MonthStartDate And @MonthToDate
		Group By
			InstitutionName, ResidencyDescription
GO
	Create Procedure [dbo].[AA_Comm_By_Institution_YearToDate]
		(@FinYearBegin as Datetime, @MonthToDate as Datetime, @MonthStartDate as Datetime)
	AS
	--Declare @FinYearBegin AS DATETIME
	--Declare @MonthToDate AS DATETIME
	--Declare @MonthStartDate AS DATETIME
	--Set @MonthStartDate = '01 March 2013'
	--Set @FinYearBegin = '01 Jan 2013'
	--Set @MonthToDate = '03 March 2013'

	Select
		InstitutionName,
		ResidencyDescription,
		SUM(DayBrokerComm) DayBrokerComm,
		SUM(DayAgentComm) DayAgentComm,
		SUM(DayNetComm) DayNetComm,
		SUM(MonthBrokerComm) MonthBrokerComm,
		SUM(MonthAgentComm) MonthAgentComm,
		SUM(MonthNetComm) MonthNetComm,
		SUM(YearBrokerComm) YearBrokerComm,
		SUM(YearAgentComm) YearAgentComm,
		SUM(YearNetComm) YearNetComm
	From
		(Select
			ISNULL(ISNULL(D.InstitutionName, M.InstitutionName), Y.InstitutionName) InstitutionName,
			ISNULL(ISNULL(D.ResidencyDescription, M.ResidencyDescription), Y.ResidencyDescription)ResidencyDescription,
			ISNULL(D.BrokerComm, 0) DayBrokerComm,
			ISNULL(D.AgentComm, 0)DayAgentComm,
			ISNULL(D.NetComm, 0) DayNetComm,
			ISNULL(M.BrokerComm, 0)MonthBrokerComm,
			ISNULL(M.AgentComm, 0)MonthAgentComm,
			ISNULL(M.NetComm, 0)MonthNetComm,
			ISNULL(Y.BrokerComm, 0)YearBrokerComm,
			ISNULL(Y.AgentComm, 0)YearAgentComm,
			ISNULL(Y.NetComm, 0)YearNetComm
		From
			(Select
					InstitutionName
					, ResidencyDescription
					, Sum(BrokerComm)BrokerComm
					, Sum(AgentComm)AgentComm
					, Sum(NetComm) NetComm
				from
					[AA_View_Comm_By_Institution]
				Where
					LotTDate = @MonthToDate
				Group By
					InstitutionName, ResidencyDescription) D
			FULL OUTER JOIN
				(Select
					InstitutionName
					, ResidencyDescription
					, Sum(BrokerComm) BrokerComm
					, Sum(AgentComm) AgentComm
					, Sum(NetComm) NetComm
				from
					[AA_View_Comm_By_Institution]
				Where
					LotTDate BETWEEN @MonthStartDate And @MonthToDate
				Group By
					InstitutionName, ResidencyDescription) M
			ON
				D.InstitutionName = M.InstitutionName
			FULL OUTER JOIN
				(Select
					InstitutionName
					, ResidencyDescription
					, Sum(BrokerComm) BrokerComm
					, Sum(AgentComm) AgentComm
					, Sum(NetComm)  NetComm
				from
					[AA_View_Comm_By_Institution]
				Where
					LotTDate BETWEEN @FinYearBegin And @MonthToDate
				Group By
					InstitutionName, ResidencyDescription) Y
				ON
					D.InstitutionName = Y.InstitutionName) C
		Group By
			InstitutionName,
			ResidencyDescription
		Order By
			InstitutionName
GO
	CREATE Procedure [dbo].[AA_Comm_By_Institution_YearToDate]
		(@FinYearBegin as Datetime, @MonthToDate as Datetime)
	AS
	--Declare @FinYearBegin AS DATETIME
	--Declare @MonthToDate AS DATETIME
	--Set @FinYearBegin = '01 Jul 2011'
	--Set @MonthToDate = '22 Dec 2011'

	Select
			InstitutionName, ResidencyDescription, Sum(BrokerComm)BrokerComm, Sum(AgentComm)AgentComm, Sum(NetComm) NetComm
		from
			[AA_View_Comm_By_Institution]
		Where
			LotTDate BETWEEN @FinYearBegin And @MonthToDate
		Group By
			InstitutionName, ResidencyDescription
GO
CREATE Procedure [dbo].[AA_CommSplit_By_Institution_Daily]
	(@Date as Datetime)
AS

--Declare @Date AS DATETIME
--Set @Date = '29 Dec 2011'
Select
	'Local' As 'Residency', ROUND(ISNULL(Q1.NetComm, 0.0000000001)/(ISNULL(Q1.NetComm, 0.0000000001) + ISNULL(Q2.NetComm, 0.0000000001)),4) CommSplit
From
	(Select
		ResidencyDescription, ISNULL(SUM(NetComm),0.0000000001) NetComm, LotTDate
	from
		[AA_View_Comm_By_Institution]
	Where
		LotTDate = @Date
	And
		ResidencyDescription = 'Local'
	Group By ResidencyDescription, LotTDate)Q1

Full Outer Join
	(Select
		ResidencyDescription, ISNULL(SUM(NetComm),0.0000000001) NetComm, LotTDate
	from
		[AA_View_Comm_By_Institution]
	Where
		LotTDate = @Date
	And
		ResidencyDescription = 'Foreign'
	Group By
		ResidencyDescription, LotTDate)Q2
	ON
		Q1.LotTdate = Q2.LotTdate
UNION ALL
	Select
		'Foreign',  ROUND(ISNULL(Q2.NetComm, 0.0000000001)/(ISNULL(Q1.NetComm, 0.0000000001) + ISNULL(Q2.NetComm, 0.0000000001)),4)  CommSplit
	From
		(Select
			ResidencyDescription, ISNULL(SUM(NetComm),0.0000000001) NetComm, LotTDate
		from
			[AA_View_Comm_By_Institution]
		Where
			LotTDate = @Date
		And
			ResidencyDescription = 'Local'
		Group By ResidencyDescription, LotTDate)Q1
	Full Outer Join
		(Select
			ResidencyDescription, ISNULL(SUM(NetComm),0.0000000001) NetComm, LotTDate
		from
			[AA_View_Comm_By_Institution]
		Where
			LotTDate = @Date
		And
			ResidencyDescription = 'Foreign'
		Group By
			ResidencyDescription, LotTDate)Q2
		ON
			Q1.LotTdate = Q2.LotTdate

GO
CREATE Procedure [dbo].[AA_CommSplit_By_Institution_Monthly](@MonthStartDate as Datetime, @MonthToDate as Datetime)
AS
--Declare @MonthStartDate AS DATETIME
--Declare @MonthToDate AS DATETIME
--Set @MonthStartDate = '01 Dec 2011'
--Set @MonthToDate = '22 Dec 2011'
Select
	'Local' As 'Residency', Avg(ISNULL(Q1.NetComm, 0.0000000001)/(ISNULL(Q1.NetComm, 0.0000000001) + ISNULL(Q2.NetComm, 0.0000000001))) CommSplit
From
	(Select
		ResidencyDescription, ISNULL(SUM(NetComm),0.0000000001) NetComm, LotTDate
	from
		[AA_View_Comm_By_Institution]
	Where
		LotTDate BETWEEN @MonthStartDate And @MonthToDate
	And
		ResidencyDescription = 'Local'
	Group By ResidencyDescription, LotTDate)Q1

Full Outer Join
	(Select
		ResidencyDescription, ISNULL(SUM(NetComm),0.0000000001) NetComm, LotTDate
	from
		[AA_View_Comm_By_Institution]
	Where
		LotTDate BETWEEN @MonthStartDate And @MonthToDate
	And
		ResidencyDescription = 'Foreign'
	Group By
		ResidencyDescription, LotTDate)Q2
	ON
		Q1.LotTdate = Q2.LotTdate
UNION ALL
	Select
		'Foreign', Avg(ISNULL(Q2.NetComm, 0.0000000001)/(ISNULL(Q1.NetComm, 0.0000000001) + ISNULL(Q2.NetComm, 0.0000000001)))  CommSplit
	From
		(Select
			ResidencyDescription, ISNULL(SUM(NetComm),0.0000000001) NetComm, LotTDate
		from
			[AA_View_Comm_By_Institution]
		Where
			LotTDate BETWEEN @MonthStartDate And @MonthToDate
		And
			ResidencyDescription = 'Local'
		Group By ResidencyDescription, LotTDate)Q1
	Full Outer Join
		(Select
			ResidencyDescription, ISNULL(SUM(NetComm),0.0000000001) NetComm, LotTDate
		from
			[AA_View_Comm_By_Institution]
		Where
			LotTDate BETWEEN @MonthStartDate And @MonthToDate
		And
			ResidencyDescription = 'Foreign'
		Group By
			ResidencyDescription, LotTDate)Q2
		ON
			Q1.LotTdate = Q2.LotTdate

GO
CREATE Procedure [dbo].[AA_CommSplit_By_Institution_YearToDate](@MonthToDate as Datetime, @FinYearBegin as Datetime)
AS
--Declare @FinYearBegin AS DATETIME
--Declare @MonthToDate AS DATETIME
--Set @FinYearBegin = '01 Jul 2011'
--Set @MonthToDate = '22 Dec 2011'
Select
	'Local' As 'Residency', ROUND(Avg(ISNULL(Q1.NetComm, 0.0000000001)/(ISNULL(Q1.NetComm, 0.0000000001) + ISNULL(Q2.NetComm, 0.0000000001))), 4) CommSplit
From
	(Select
		ResidencyDescription, ISNULL(SUM(NetComm),0.0000000001) NetComm, LotTDate
	from
		[AA_View_Comm_By_Institution]
	Where
		LotTDate BETWEEN @FinYearBegin And @MonthToDate
	And
		ResidencyDescription = 'Local'
	Group By ResidencyDescription, LotTDate)Q1

Full Outer Join
	(Select
		ResidencyDescription, ISNULL(SUM(NetComm),0.0000000001) NetComm, LotTDate
	from
		[AA_View_Comm_By_Institution]
	Where
		LotTDate BETWEEN @FinYearBegin And @MonthToDate
	And
		ResidencyDescription = 'Foreign'
	Group By
		ResidencyDescription, LotTDate)Q2
	ON
		Q1.LotTdate = Q2.LotTdate
UNION ALL
	Select
		'Foreign', ROUND(Avg(ISNULL(Q2.NetComm, 0.0000000001)/(ISNULL(Q1.NetComm, 0.0000000001) + ISNULL(Q2.NetComm, 0.0000000001))), 4)  CommSplit
	From
		(Select
			ResidencyDescription, ISNULL(SUM(NetComm),0.0000000001) NetComm, LotTDate
		from
			[AA_View_Comm_By_Institution]
		Where
			LotTDate BETWEEN @FinYearBegin And @MonthToDate
		And
			ResidencyDescription = 'Local'
		Group By ResidencyDescription, LotTDate)Q1
	Full Outer Join
		(Select
			ResidencyDescription, ISNULL(SUM(NetComm),0.0000000001) NetComm, LotTDate
		from
			[AA_View_Comm_By_Institution]
		Where
			LotTDate BETWEEN @FinYearBegin And @MonthToDate
		And
			ResidencyDescription = 'Foreign'
		Group By
			ResidencyDescription, LotTDate)Q2
		ON
			Q1.LotTdate = Q2.LotTdate

GO

CREATE PROCEDURE [dbo].[AccountHistoryProc] AS


SELECT  dbo.HistoryTransactionsList.* INTO #tableA FROM dbo.HistoryTransactionsList


SELECT * FROM

(SELECT  TOP 100 PERCENT a.Entity_DPA_,CAST(FLOOR(CAST(a.TransDate AS float)) AS datetime) as TransDate,a.Ref,a.Particulars,a.Debit, a.Credit,
	CASE WHEN (SUM(b.Balance) >= 0) Then CONVERT(NVARCHAR, SUM(b.Balance)) + ' Cr' ELSE CONVERT(NVARCHAR,  ABS(SUM(b.Balance))) + ' Dr' END   AS Balance,
	a.IsOpeningBalance As IsOpeningBalance, a.EntityName

FROM #tableA  a
CROSS JOIN #tableA  b
WHERE (b.TransDate <= a.TransDate) AND (a.Entity_DPA_ = b.Entity_DPA_)
GROUP BY a.Entity_DPA_,a.TransDate, a.Ref,a.Particulars,a.Debit,a.Credit, a.IsOpeningBalance, a.EntityName
ORDER BY a.Entity_DPA_, a.TransDate) TBLTRANS

UNION ALL

SELECT CAST(BankAccount_DPA_ AS NVARCHAR(500)) AS BankAccount_DPA_,
	CAST(FLOOR(CAST(TransDate AS float)) AS datetime) AS TransDate,
	Ref, Particulars,
	Debit, Credit, Balance,
	IsOpeningBalance,
	AccountName As EntityName FROM dbo.BankAccountStatement
INNER JOIN dbo.Account ON  dbo.BankAccountStatement.BankAccount_DPA_ =  dbo.Account.Account_DPA_

UNION ALL


SELECT
	'******1' AS [Entity_DPA_],
	CAST(FLOOR(CAST(GetDate() AS float)) AS datetime) AS TransDate,'' AS Ref, '' AS Particulars,
	CASE
		WHEN SUM(Type1.Balance) < 0 THEN 0 - SUM(Type1.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(Type1.Balance) >= 0 THEN SUM(Type1.Balance)
		ELSE 0 END AS Credit,
		CASE
			WHEN (SUM(Type1.Balance) >= 0) THEN CONVERT(NVARCHAR, SUM(Type1.Balance)) + ' Cr'
			ELSE CONVERT(NVARCHAR,  ABS(SUM(Type1.Balance))) + ' Dr' END   AS Balance,
	0 AS IsOpeningBalance,
	(SELECT EntityTypeName FROM dbo.FullEntityTypeList WHERE EntityType_DPA_ = 1) AS EntityName

	 FROM


(SELECT     ClientsStatement.*, Credit - Debit AS Balance
                       FROM          (SELECT     Client_DPA_, CAST(FLOOR(CAST(ClientRegDate AS float)) AS datetime) AS TransDate, '' AS REF, ' Opening Balance' AS Particulars, 0 AS Debit, ClientOpeningBal AS Credit,
                              1 AS IsOpeningBalance
       FROM          dbo.Client
       UNION ALL
       SELECT     dbo.Payment.Entity_DPA_ AS Client_DPA_, CAST(FLOOR(CAST(dbo.Payment.PaymentPDate AS float)) AS datetime) AS TransDate, dbo.Payment.PaymentReference AS Ref,
                             dbo.Payment.PaymentNarrative AS Particulars,
                             CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.Payment INNER JOIN
                             dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
       WHERE     EntityType_DPA_ = 1
       UNION ALL
       SELECT     Client_DPA_, CAST(FLOOR(CAST(LotTDate AS float)) AS datetime) AS TransDate, ContractNumber AS Ref, SecurityName AS Particulars,
                             CASE OrderTypeSale WHEN 0 THEN NetAmt ELSE 0 END AS Debit,
                             CASE OrderTypeSale WHEN 1 THEN NetAmt ELSE 0 END AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.ClientTransactionsSubList
       UNION ALL
       SELECT     dbo.JournalList.Entity_DPA_ AS Client_DPA_, CAST(FLOOR(CAST(dbo.JournalList.JournalDate AS float)) AS datetime) AS TransDate, dbo.JournalList.JournalEntry_DPA_ AS Ref,
                             dbo.JournalList.JournalNarrative AS Particulars, JournalList.JournalEntryDebit AS Debit,
                             JournalList.JournalEntryCredit AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.JournalList
       WHERE     EntityType_DPA_ = 1) ClientsStatement) Type1

UNION ALL

SELECT
	'******2' AS [Account Code],
	CAST(FLOOR(CAST(GetDate() AS float)) AS datetime) AS TransDate,'' AS Ref, '' AS Particulars,
	CASE
		WHEN SUM(Type2.Balance) < 0 THEN 0 - SUM(Type2.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(Type2.Balance) >= 0 THEN SUM(Type2.Balance)
		ELSE 0 END AS Credit,
		CASE
			WHEN (SUM(Type2.Balance) >= 0) THEN CONVERT(NVARCHAR, SUM(Type2.Balance)) + ' Cr'
			ELSE CONVERT(NVARCHAR,  ABS(SUM(Type2.Balance))) + ' Dr' END   AS Balance,
	0 AS IsOpeningBalance,
	(SELECT EntityTypeName FROM dbo.FullEntityTypeList WHERE EntityType_DPA_ = 2) AS AccountName
FROM
	(SELECT     AgentTransactions.*, Credit - Debit AS Balance
	FROM         (SELECT     Agent_DPA_, CAST(FLOOR(CAST(AgentRegDate AS float)) AS datetime) AS TransDate, '' AS REF, ' Opening Balance' AS Particulars,
                              CASE WHEN AgentOpeningBal < 0 THEN (0 - AgentOpeningBal) ELSE 0 END AS Debit,
                              CASE WHEN AgentOpeningBal >= 0 THEN AgentOpeningBal ELSE 0 END AS Credit, 1 AS IsOpeningBalance
       FROM          dbo.Agent
       UNION ALL
       SELECT     dbo.Payment.Entity_DPA_ AS Agent_DPA_, CAST(FLOOR(CAST(dbo.Payment.PaymentPDate AS float)) AS datetime) AS TransDate, dbo.Payment.PaymentReference AS Ref,
                             dbo.Payment.PaymentNarrative AS Particulars, CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.Payment INNER JOIN
                             dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
       WHERE     EntityType_DPA_ = 2
       UNION ALL
       SELECT     Agent_DPA_, CAST(FLOOR(CAST(LotTDate AS float)) AS datetime) AS TransDate, ContractNumber AS Ref, SecurityName AS Particulars, 0 AS Debit, isnull(AgentCommission, 0)
                             AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.AgentCommissionList
       UNION ALL
	 SELECT     dbo.JournalList.Entity_DPA_ AS Agent_DPA_, CAST(FLOOR(CAST(dbo.JournalList.JournalDate AS float)) AS datetime) AS TransDate, dbo.JournalList.JournalEntry_DPA_ AS Ref,
                             dbo.JournalList.JournalNarrative AS Particulars, JournalList.JournalEntryDebit AS Debit,
                             JournalList.JournalEntryCredit AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.JournalList
       WHERE     EntityType_DPA_ = 2) AgentTransactions) Type2

UNION ALL

SELECT
	'******7' AS [Account Code],
	CAST(FLOOR(CAST(GetDate() AS float)) AS datetime) AS TransDate,'' AS Ref, '' AS Particulars,
	CASE
		WHEN SUM(Type7.Balance) < 0 THEN 0 - SUM(Type7.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(Type7.Balance) >= 0 THEN SUM(Type7.Balance)
		ELSE 0 END AS Credit,
		CASE
			WHEN (SUM(Type7.Balance) >= 0) THEN CONVERT(NVARCHAR, SUM(Type7.Balance)) + ' Cr'
			ELSE CONVERT(NVARCHAR,  ABS(SUM(Type7.Balance))) + ' Dr' END   AS Balance,
	0 AS IsOpeningBalance,
	(SELECT EntityTypeName FROM dbo.FullEntityTypeList WHERE EntityType_DPA_ = 7) AS AccountName
FROM
	(SELECT     OwnerTransactions.*, Credit - Debit AS Balance
	FROM         (SELECT     Owner_DPA_, CAST(FLOOR(CAST(OwnerRegDate AS float)) AS datetime) AS TransDate, '' AS REF, ' Opening Balance' AS Particulars,
                              CASE WHEN OwnerOpeningBal < 0 THEN (0 - OwnerOpeningBal) ELSE 0 END AS Debit,
                              CASE WHEN OwnerOpeningBal >= 0 THEN OwnerOpeningBal ELSE 0 END AS Credit, 1 AS IsOpeningBalance
       FROM          dbo.Owner
       UNION ALL
       SELECT     dbo.Payment.Entity_DPA_ AS Owner_DPA_,CAST(FLOOR(CAST(dbo.Payment.PaymentPDate AS float)) AS datetime) AS TransDate, dbo.Payment.PaymentReference AS Ref,
                             dbo.Payment.PaymentNarrative AS Particulars, CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.Payment INNER JOIN
                             dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
       WHERE     EntityType_DPA_ = 7
       UNION ALL
       SELECT     Owner_DPA_, CAST(FLOOR(CAST(LotTDate AS float)) AS datetime) AS TransDate, ContractNumber AS Ref, SecurityName AS Particulars, 0 AS Debit, isnull(OwnerCommission, 0)
                             AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.OwnerCommissionList
       UNION ALL
	 SELECT     dbo.JournalList.Entity_DPA_ AS Owner_DPA_, CAST(FLOOR(CAST(dbo.JournalList.JournalDate AS float)) AS datetime) AS TransDate, dbo.JournalList.JournalEntry_DPA_ AS Ref,
                             dbo.JournalList.JournalNarrative AS Particulars, JournalList.JournalEntryDebit AS Debit,
                             JournalList.JournalEntryCredit AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.JournalList
       WHERE     EntityType_DPA_ = 7) OwnerTransactions) Type7



union all



SELECT
	'******3' AS [Account Code],
	CAST(FLOOR(CAST(GetDate() AS float)) AS datetime) AS TransDate,'' AS Ref, '' AS Particulars,
	CASE
		WHEN SUM(Type3.Balance) < 0 THEN 0 - SUM(Type3.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(Type3.Balance) >= 0 THEN SUM(Type3.Balance)
		ELSE 0 END AS Credit,
	CASE
		WHEN (SUM(Type3.Balance) >= 0) THEN CONVERT(NVARCHAR, SUM(Type3.Balance)) + ' Cr'
		ELSE CONVERT(NVARCHAR,  ABS(SUM(Type3.Balance))) + ' Dr' END   AS Balance,
	0 AS IsOpeningBalance,
	(SELECT EntityTypeName FROM dbo.FullEntityTypeList WHERE EntityType_DPA_ = 3) AS AccountName

	FROM
(SELECT     BrokerTransactions.*, Credit - Debit AS Balance
FROM         (SELECT     Broker_DPA_, CAST(FLOOR(CAST(BrokerRegDate AS float)) AS datetime) AS TransDate, '' AS REF, ' Opening Balance' AS Particulars, case when BrokerOpeningBal < 0 then BrokerOpeningBal else 0 end AS Debit, case when BrokerOpeningBal >= 0 then BrokerOpeningBal else 0 end AS Credit,
                                              1 AS IsOpeningBalance
                       FROM          dbo.Broker
                       UNION ALL
                       SELECT     dbo.Payment.Entity_DPA_ AS Broker_DPA_, CAST(FLOOR(CAST(dbo.Payment.PaymentPDate AS float)) AS datetime) AS TransDate, dbo.Payment.PaymentReference AS Ref,
                                             dbo.Payment.PaymentNarrative AS Particulars, CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit, 0 AS IsOpeningBalance
                       FROM         dbo.Payment INNER JOIN
                                             dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
                       WHERE     EntityType_DPA_ = 3
                       UNION ALL
                       SELECT     Broker_DPA_, CAST(FLOOR(CAST(LotTDate AS float)) AS datetime) AS TransDate, ContractNumber AS Ref, SecurityName AS Particulars,
                                             CASE OrderTypeSale WHEN 1 THEN Gross ELSE 0 END AS Debit, CASE OrderTypeSale WHEN 0 THEN Gross ELSE 0 END AS Credit,
                                             0 AS IsOpeningBalance
                       FROM         dbo.BrokerTransactionsSubList
			UNION ALL
		       SELECT     dbo.JournalList.Entity_DPA_ AS Broker_DPA_, CAST(FLOOR(CAST(dbo.JournalList.JournalDate AS float)) AS datetime) AS TransDate, dbo.JournalList.JournalEntry_DPA_ AS Ref,
		                             dbo.JournalList.JournalNarrative AS Particulars, JournalList.JournalEntryDebit AS Debit,
		                             JournalList.JournalEntryCredit AS Credit, 0 AS IsOpeningBalance
		       FROM         dbo.JournalList
		       WHERE     EntityType_DPA_ = 3) BrokerTransactions) Type3

UNION ALL
SELECT
	'******' + CAST(dbo.FullEntityTypeList.EntityType_DPA_ AS VARCHAR(500)) AS [Account Code],
	CAST(FLOOR(CAST(GetDate() AS float)) AS datetime) AS TransDate,'' AS Ref, '' AS Particulars,
	CASE
		WHEN SUM(dbo.EntityTransactionList.Balance) < 0 THEN 0 - SUM(dbo.EntityTransactionList.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.EntityTransactionList.Balance) >= 0 THEN SUM(dbo.EntityTransactionList.Balance)
		ELSE 0 END AS Credit,
	CASE
		WHEN (SUM(EntityTransactionList.Balance) >= 0) THEN CONVERT(NVARCHAR, SUM(EntityTransactionList.Balance)) + ' Cr'
		ELSE CONVERT(NVARCHAR,  ABS(SUM(EntityTransactionList.Balance))) + ' Dr' END   AS Balance,
	0 AS IsOpeningBalance,
	dbo.FullEntityTypeList.EntityTypeName AS AccountName

FROM         dbo.Entity INNER JOIN
                      dbo.FullEntityTypeList ON dbo.Entity.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_ INNER JOIN
                      dbo.EntityTransactionList ON dbo.Entity.Entity_DPA_ = dbo.EntityTransactionList.Entity_DPA_
GROUP BY dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName, dbo.FullEntityTypeList.EntityType_DPA_




DROP TABLE #tableA

GO

CREATE PROCEDURE BestBalanceQtys AS UPDATE OrdDetail
SET       OrdDetail.OrdDetailQty = Case dbo.datastream_SecurityPriceList.Price
when 0 Then 0 Else CONVERT(Numeric, dbo.OrdDetail.Amount / (dbo.datastream_SecurityPriceList.Price * 1.020825))End
FROM         OrdDetail INNER JOIN
                      datastream_SecurityPriceList ON OrdDetail.Security_DPA_ = datastream_SecurityPriceList.Security_DPA_ INNER JOIN
                      tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_
WHERE     (OrdDetail.Best = 1) AND (OrdDetail.Amount <> 0) AND (tbOrder.OrderType_DPA_ = 1)
GO

CREATE PROCEDURE [dbo].[ChartofAccountsProc] AS


SELECT
	(SELECT EntityTypeAccountType FROM dbo.FullEntityTypeList WHERE EntityType_DPA_ = 1) AS AccountType,
	(SELECT EntityTypeName FROM dbo.FullEntityTypeList WHERE EntityType_DPA_ = 1) AS AccountName,
	'******1' AS [Account Code],
	SUM(Type1.Balance) AS Balance FROM
(SELECT     ClientsStatement.*, Credit - Debit AS Balance
                       FROM          (SELECT     Client_DPA_, ClientRegDate AS TransDate, '' AS REF, ' Opening Balance' AS Particulars, 0 AS Debit, ClientOpeningBal AS Credit,
                              1 AS IsOpeningBalance
       FROM          dbo.Client
       UNION ALL
       SELECT     dbo.Payment.Entity_DPA_ AS Client_DPA_, dbo.Payment.PaymentPDate AS TransDate, dbo.Payment.PaymentReference AS Ref,
                             dbo.Payment.PaymentNarrative AS Particulars,
                             CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.Payment INNER JOIN
                             dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
       WHERE     EntityType_DPA_ = 1
       UNION ALL
       SELECT     Client_DPA_, LotTDate AS TransDate, ContractNumber AS Ref, SecurityName AS Particulars,
                             CASE OrderTypeSale WHEN 0 THEN NetAmt ELSE 0 END AS Debit,
                             CASE OrderTypeSale WHEN 1 THEN NetAmt ELSE 0 END AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.ClientTransactionsSubList
       UNION ALL
       SELECT     dbo.JournalList.Entity_DPA_ AS Client_DPA_, dbo.JournalList.JournalDate AS TransDate, dbo.JournalList.JournalEntry_DPA_ AS Ref,
                             dbo.JournalList.JournalNarrative AS Particulars, JournalList.JournalEntryDebit AS Debit,
                             JournalList.JournalEntryCredit AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.JournalList
       WHERE     EntityType_DPA_ = 1) ClientsStatement) Type1

UNION ALL

SELECT
	(SELECT EntityTypeAccountType FROM dbo.FullEntityTypeList WHERE EntityType_DPA_ = 2) AS AccountType,
	(SELECT EntityTypeName FROM dbo.FullEntityTypeList WHERE EntityType_DPA_ = 2) AS AccountName,
	'******2' AS [Account Code],
	SUM(Type2.Balance) AS Balance FROM
(SELECT     AgentTransactions.*, Credit - Debit AS Balance
FROM         (SELECT     Agent_DPA_, AgentRegDate AS TransDate, '' AS REF, ' Opening Balance' AS Particulars,
                              CASE WHEN AgentOpeningBal < 0 THEN (0 - AgentOpeningBal) ELSE 0 END AS Debit,
                              CASE WHEN AgentOpeningBal >= 0 THEN AgentOpeningBal ELSE 0 END AS Credit, 1 AS IsOpeningBalance
       FROM          dbo.Agent
       UNION ALL
       SELECT     dbo.Payment.Entity_DPA_ AS Agent_DPA_, dbo.Payment.PaymentPDate AS TransDate, dbo.Payment.PaymentReference AS Ref,
                             dbo.Payment.PaymentNarrative AS Particulars, CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.Payment INNER JOIN
                             dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
       WHERE     EntityType_DPA_ = 2
       UNION ALL
       SELECT     Agent_DPA_, LotTDate AS TransDate, ContractNumber AS Ref, SecurityName AS Particulars, 0 AS Debit, isnull(AgentCommission, 0)
                             AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.AgentCommissionList
       UNION ALL
	 SELECT     dbo.JournalList.Entity_DPA_ AS Agent_DPA_, dbo.JournalList.JournalDate AS TransDate, dbo.JournalList.JournalEntry_DPA_ AS Ref,
                             dbo.JournalList.JournalNarrative AS Particulars, JournalList.JournalEntryDebit AS Debit,
                             JournalList.JournalEntryCredit AS Credit, 0 AS IsOpeningBalance
       FROM         dbo.JournalList
       WHERE     EntityType_DPA_ = 2) AgentTransactions) Type2


union all

SELECT
	(SELECT EntityTypeAccountType FROM dbo.FullEntityTypeList WHERE EntityType_DPA_ = 3) AS AccountType,
	(SELECT EntityTypeName FROM dbo.FullEntityTypeList WHERE EntityType_DPA_ = 3) AS AccountName,
	'******3' AS [Account Code],
	SUM(Type3.BALANCE) AS Balance FROM
(SELECT     BrokerTransactions.*, Credit - Debit AS Balance
FROM         (SELECT     Broker_DPA_, BrokerRegDate AS TransDate, '' AS REF, ' Opening Balance' AS Particulars, case when BrokerOpeningBal < 0 then BrokerOpeningBal else 0 end AS Debit, case when BrokerOpeningBal >= 0 then BrokerOpeningBal else 0 end AS Credit,
                                              1 AS IsOpeningBalance
                       FROM          dbo.Broker
                       UNION ALL
                       SELECT     dbo.Payment.Entity_DPA_ AS Broker_DPA_, dbo.Payment.PaymentPDate AS TransDate, dbo.Payment.PaymentReference AS Ref,
                                             dbo.Payment.PaymentNarrative AS Particulars, CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit, 0 AS IsOpeningBalance
                       FROM         dbo.Payment INNER JOIN
                                             dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
                       WHERE     EntityType_DPA_ = 3
                       UNION ALL
                       SELECT     Broker_DPA_, LotTDate AS TransDate, ContractNumber AS Ref, SecurityName AS Particulars,
                                             CASE OrderTypeSale WHEN 1 THEN Gross ELSE 0 END AS Debit, CASE OrderTypeSale WHEN 0 THEN Gross ELSE 0 END AS Credit,
                                             0 AS IsOpeningBalance
                       FROM         dbo.BrokerTransactionsSubList
			UNION ALL
		       SELECT     dbo.JournalList.Entity_DPA_ AS Broker_DPA_, dbo.JournalList.JournalDate AS TransDate, dbo.JournalList.JournalEntry_DPA_ AS Ref,
		                             dbo.JournalList.JournalNarrative AS Particulars, JournalList.JournalEntryDebit AS Debit,
		                             JournalList.JournalEntryCredit AS Credit, 0 AS IsOpeningBalance
		       FROM         dbo.JournalList
		       WHERE     EntityType_DPA_ = 3) BrokerTransactions) Type3

union all

SELECT
	(SELECT EntityTypeAccountType FROM dbo.FullEntityTypeList WHERE EntityType_DPA_ = 7) AS AccountType,
	(SELECT EntityTypeName FROM dbo.FullEntityTypeList WHERE EntityType_DPA_ = 7) AS AccountName,
	'******7' AS [Account Code],
	SUM(Type7.BALANCE) AS Balance FROM
(SELECT     OwnerTransactions.*, Credit - Debit AS Balance
FROM         (SELECT     Owner_DPA_, OwnerRegDate AS TransDate, '' AS REF, ' Opening Balance' AS Particulars, case when OwnerOpeningBal < 0 then OwnerOpeningBal else 0 end AS Debit, case when OwnerOpeningBal >= 0 then OwnerOpeningBal else 0 end AS Credit,
                                              1 AS IsOpeningBalance
                       FROM          dbo.Owner
                       UNION ALL
                       SELECT     dbo.Payment.Entity_DPA_ AS Owner_DPA_, dbo.Payment.PaymentPDate AS TransDate, dbo.Payment.PaymentReference AS Ref,
                                             dbo.Payment.PaymentNarrative AS Particulars, CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit, 0 AS IsOpeningBalance
                       FROM         dbo.Payment INNER JOIN
                                             dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
                       WHERE     EntityType_DPA_ = 7
                       UNION ALL
                       SELECT     Owner_DPA_, LotTDate AS TransDate, ContractNumber AS Ref, SecurityName AS Particulars, 0 AS Debit, isnull(OwnerCommission, 0)
		                             AS Credit, 0 AS IsOpeningBalance
		       FROM         dbo.OwnerCommissionList
			UNION ALL
		       SELECT     dbo.JournalList.Entity_DPA_ AS Owner_DPA_, dbo.JournalList.JournalDate AS TransDate, dbo.JournalList.JournalEntry_DPA_ AS Ref,
		                             dbo.JournalList.JournalNarrative AS Particulars, JournalList.JournalEntryDebit AS Debit,
		                             JournalList.JournalEntryCredit AS Credit, 0 AS IsOpeningBalance
		       FROM         dbo.JournalList
		       WHERE     EntityType_DPA_ = 7) OwnerTransactions) Type7


UNION ALL
SELECT     dbo.FullEntityTypeList.EntityTypeAccountType AS AccountType,dbo.FullEntityTypeList.EntityTypeName AS AccountName, '******' + CAST(dbo.FullEntityTypeList.EntityType_DPA_ AS VARCHAR(500)) AS [Account Code],  SUM(dbo.EntityTransactionList.Balance)
                      AS Balance
FROM         dbo.Entity INNER JOIN
                      dbo.FullEntityTypeList ON dbo.Entity.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_ INNER JOIN
                      dbo.EntityTransactionList ON dbo.Entity.Entity_DPA_ = dbo.EntityTransactionList.Entity_DPA_
GROUP BY dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName, dbo.FullEntityTypeList.EntityType_DPA_

UNION ALL

SELECT     dbo.EntityList.AccountTypeName AS AccountType,dbo.EntityList.EntityName AS AccountName, '------' + CAST(dbo.EntityList.Entity_DPA_ AS VARCHAR(500)) AS [Account Code],  SUM(dbo.LevyTransactionList.Balance)
                      AS Balance
FROM         dbo.EntityList INNER JOIN
                      dbo.LevyTransactionList ON dbo.EntityList.Entity_DPA_ = dbo.LevyTransactionList.Entity_DPA_
GROUP BY dbo.EntityList.AccountTypeName,dbo.EntityList.EntityName, dbo.EntityList.Entity_DPA_

UNION ALL


SELECT     dbo.AccountList.AccountTypeLevel1 AS AccountType,dbo.AccountList.AccountName, CAST(dbo.AccountList.Account_DPA_ AS SQL_VARIANT) AS [Account Code],  SUM(dbo.NominalTransactionList.Balance)
                      AS Balance
FROM         dbo.AccountList INNER JOIN
                      dbo.NominalTransactionList ON dbo.AccountList.Account_DPA_ = dbo.NominalTransactionList.Account_DPA_
GROUP BY dbo.AccountList.AccountTypeLevel1,dbo.AccountList.AccountName, dbo.AccountList.Account_DPA_

UNION ALL

SELECT    dbo.AccountList.AccountTypeLevel1 AS AccountType,dbo.AccountList.AccountName, CAST(dbo.AccountList.Account_DPA_ AS SQL_VARIANT) AS [Account Code],  SUM(dbo.BankTransactionList.Balance) AS Balance
FROM         dbo.BankTransactionList INNER JOIN
                      dbo.AccountList ON dbo.BankTransactionList.BankAccount_DPA_ = dbo.AccountList.Account_DPA_
GROUP BY dbo.AccountList.AccountTypeLevel1,dbo.AccountList.AccountName, dbo.AccountList.Account_DPA_

GO


CREATE PROCEDURE ClientBalanceProcedure @clientDPA int
AS
SET NOCOUNT ON
--UPDATE ClientBalances SET CurrentBal=CurrentBalances.CurrentBal from ClientBalances inner join CurrentBalances on ClientBalances.Client_DPA_= CurrentBalances.Client_DPA_ Where CurrentBalances.Client_DPA_=@clientDPA
delete from ClientBalances where Client_DPA_=@clientDPA
Insert Into ClientBalances(Client_DPA_,CurrentBal)
SELECT  ClientsStatement.Client_DPA_,
SUM(ISNULL(ClientsStatement.Credit-ClientsStatement.Debit, 0)) + dbo.Client.ClientOpeningBal AS CurrentBal
FROM (
Select * from ClientStatement WHERE Client_DPA_ = @clientDPA
) ClientsStatement
INNER JOIN
                      dbo.Client ON ClientsStatement.Client_DPA_ = dbo.Client.Client_DPA_
WHERE     (dbo.Client.Deleted = 0)  AND dbo.Client.Client_DPA_= @clientDPA
GROUP BY ClientsStatement.Client_DPA_, dbo.Client.ClientOpeningBal


SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE ClientBalanceProcedure1
AS
SET NOCOUNT ON
UPDATE ClientBalances SET CurrentBal=CurrentBalances.CurrentBal from ClientBalances inner join CurrentBalances on ClientBalances.Client_DPA_=CurrentBalances.Client_DPA_
GO

CREATE PROCEDURE ClientBalancesDelete
AS
Delete From ClientBalances

GO

CREATE PROCEDURE [dbo].[ClientBalancesProcedure]
AS

SET NOCOUNT ON
delete from ClientBalances
Insert Into ClientBalances(Client_DPA_,CurrentBal) Select Client_DPA_,CurrentBal From CurrentBalances
GO

CREATE PROCEDURE ClientBalancesUpdate
AS
UPDATE ClientBalances
SET       CurrentBal =
                   (SELECT CurrentBal
                    FROM   CurrentBalances
                    WHERE CurrentBalances.Client_DPA_ = ClientBalances.Client_DPA_)

GO
CREATE      PROCEDURE ClientStatementProcAll
AS

SET NOCOUNT ON/*
delete from tblStatementList where Year(TransDate)>2006 --where Client_DPA_ = @ClientID
insert tblStatementList
SELECT     ClientsStatement.*, CreditBal - Debit AS Balance
FROM         (
                       SELECT     dbo.Payment.Entity_DPA_ AS Client_DPA_, dbo.Payment.PaymentPDate AS TransDate, dbo.Payment.PaymentReference AS Ref,
                                             dbo.Payment.PaymentNarrative AS Particulars, CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS CreditBal, 0 AS IsOpeningBalance,
                                             PaymentClientStatementNos.ReceiptNo, PaymentClientStatementNos.type AS receipttype
                       FROM         dbo.Payment INNER JOIN
                                             dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_ INNER JOIN
                                             (SELECT     Payment_DPA_, CONVERT(char(5), PaymentReceiptNo) AS ReceiptNo, 1 AS type
FROM         dbo.Payment
WHERE     (NOT (PaymentReceiptNo IS NULL)) AND (EntityType_DPA_ <> 8) AND Deleted = 0 and Year(dbo.Payment.PaymentPDate)>2006 --and Entity_DPA_=@ClientID
UNION
SELECT     dbo.Payment.Payment_DPA_, LotList.ContractNumber, 2 AS Type
FROM         dbo.Payment INNER JOIN
                      dbo.Voucher ON dbo.Payment.Voucher_DPA_ = dbo.Voucher.Voucher_DPA_ INNER JOIN
                      dbo.Contract ON dbo.Voucher.Voucher_DPA_ = dbo.Contract.Voucher_DPA_ INNER JOIN
                      (SELECT     Contract.Voucher_DPA_, Lot.ContractNumber, tbOrder.InterBank,Lot.Contract_DPA_
FROM         tbOrder INNER JOIN
                      OrdDetail ON tbOrder.Order_DPA_ = OrdDetail.Order_DPA_ INNER JOIN
                      Lot ON OrdDetail.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN
                      Contract ON Lot.Contract_DPA_ = Contract.Contract_DPA_
WHERE     (tbOrder.OrderCanceled = 0) AND (tbOrder.OrderHold = 0) AND (tbOrder.Deleted = 0) AND (OrdDetail.Deleted = 0) AND
                      (Lot.Deleted = 0) AND (Contract.Deleted = 0) AND (NOT (Contract.Voucher_DPA_ IS NULL))) LotList ON dbo.Contract.Contract_DPA_ = LotList.Contract_DPA_
WHERE     lotlist.interbank <> 1 AND Payment.Deleted = 0 and Year(dbo.Payment.PaymentPDate)>2006 --and payment.entity_DPA_=@ClientID
UNION
SELECT     dbo.Payment.Payment_DPA_, LotList.ContractNumber, 2 AS Type
FROM         dbo.BrokerReceiptVoucher INNER JOIN
                      dbo.Payment ON dbo.BrokerReceiptVoucher.BrokerReceiptVoucher_DPA_ = dbo.Payment.BrokerReceiptVoucher_DPA_ INNER JOIN
                      (SELECT     Contract.BrokerReceiptVoucher_DPA_, Lot.ContractNumber, tbOrder.InterBank,Lot.Contract_DPA_
FROM         tbOrder INNER JOIN
                      OrdDetail ON tbOrder.Order_DPA_ = OrdDetail.Order_DPA_ INNER JOIN
                      Lot ON OrdDetail.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN
                      Contract ON Lot.Contract_DPA_ = Contract.Contract_DPA_
WHERE     (tbOrder.OrderCanceled = 0) AND (tbOrder.OrderHold = 0) AND (tbOrder.Deleted = 0)  AND (OrdDetail.Deleted = 0) AND
                      (Lot.Deleted = 0) AND (Contract.Deleted = 0) AND (NOT (Contract.BrokerReceiptVoucher_DPA_ IS NULL))) LotList INNER JOIN
                      dbo.Contract ON LotList.Contract_DPA_ = dbo.Contract.Contract_DPA_ ON
                      dbo.BrokerReceiptVoucher.BrokerReceiptVoucher_DPA_ = dbo.Contract.BrokerReceiptVoucher_DPA_
WHERE     lotlist.interbank <> 1 AND Payment.Deleted = 0 AND Contract.Deleted = 0 and Year(dbo.Payment.PaymentPDate)>2006 --and payment.entity_DPA_=@ClientID
UNION
SELECT     dbo.Payment.Payment_DPA_, dbo.Payment.PaymentReference, 3 AS Type
FROM         dbo.Payment INNER JOIN
                      dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
WHERE     (dbo.Payment.PayType_DPA_ = 2) AND (dbo.Payment.Voucher_DPA_ IS NULL) AND Payment.Deleted = 0 and Year(dbo.Payment.PaymentPDate)>2006) PaymentClientStatementNos ON Payment.Payment_DPA_ = PaymentClientStatementNos.Payment_DPA_
                       WHERE     EntityType_DPA_ = 1 AND Payment.deleted = 0 and Year(dbo.Payment.PaymentPDate)>2006 --and Payment.Entity_DPA_=@ClientID
                       UNION ALL
                       SELECT     Client_DPA_, LotTDate AS TransDate, ContractNumber AS Ref, SecurityName AS Particulars,
                                             CASE OrderTypeSale WHEN 0 THEN NetAmt ELSE 0 END AS Debit, CASE OrderTypeSale WHEN 1 THEN NetAmt ELSE 0 END AS Credit,
                                             CASE OrderTypeSale WHEN 1 THEN NetAmt ELSE 0 END AS CreditBal, 0 AS IsOpeningBalance, ContractNumber AS PaymentReceiptNo,
                                             2 AS receipttype
                       FROM         dbo.ClientTransactionsSubList
					where Year(LotTDate)>2006
                       UNION  ALL
                       SELECT     dbo.JournalList.Entity_DPA_ AS Client_DPA_, dbo.JournalList.JournalDate AS TransDate,
                                             CAST(dbo.JournalList.JournalEntry_DPA_ AS Nvarchar(500)) AS Ref, dbo.JournalList.JournalNarrative AS Particulars,
                                             JournalList.JournalEntryDebit AS Debit, JournalList.JournalEntryCredit AS Credit, JournalList.JournalEntryCredit AS CreditBal,
                                             0 AS IsOpeningBalance, '' AS PaymentReceiptNo, 4 AS receipttype
                       FROM         dbo.JournalList
                       WHERE     EntityType_DPA_ = 1 and Year(dbo.JournalList.JournalDate) >2006 --and dbo.JournalList.Entity_DPA_ = @ClientID
                       UNION ALL
                       SELECT     dbo.InterTransfers.Client_DPA_, dbo.InterTransfer.TransferDate, dbo.InterTransfer.TransferReference AS REF,
                                             isnull(dbo.InterTransfer.TransferNarrative, '') + '  ' + InterTransferType.TypeDescription AS Particulars, 0 AS Debit,
                                             dbo.InterTransfer.TransferAmount AS Credit, dbo.InterTransfer.TransferAmount AS CreditBal, 0 AS IsOpeningBalance,
                                             dbo.InterTransfers.ContractNumber AS PaymentReceiptNo, 5 AS receipttype
                       FROM         dbo.Contract INNER JOIN
                                             dbo.InterTransfer ON dbo.Contract.Contract_DPA_ = dbo.InterTransfer.Contract_DPA_ INNER JOIN
                                             dbo.InterTransfers ON dbo.Contract.Contract_DPA_ = dbo.InterTransfers.Contract_DPA_ INNER JOIN
                                             dbo.InterTransferType ON dbo.InterTransfer.InterTransferType_DPA_ = dbo.InterTransferType.InterTransferType_DPA_
                       WHERE     (dbo.InterTransfers.OrderTypeSale = 0) AND intertransfer.deleted = 0 and Year(dbo.InterTransfer.TransferDate)>2006 --and dbo.InterTransfers.Client_DPA_=@ClientID
                       UNION ALL
                       SELECT     dbo.InterTransfers.Client_DPA_, dbo.InterTransfer.TransferDate, dbo.InterTransfer.TransferReference AS REF,
                                             isnull(dbo.InterTransfer.TransferNarrative, '') + '  ' + InterTransferType.TypeDescription AS Particulars,
                                             dbo.InterTransfer.TransferAmount AS Debit, 0 AS Credit, 0 AS CreditBal, 0 AS IsOpeningBalance,
                                             dbo.InterTransfers.ContractNumber AS PaymentReceiptNo, 5 AS receipttype
                       FROM         dbo.Contract INNER JOIN
                                             dbo.InterTransfer ON dbo.Contract.Contract_DPA_ = dbo.InterTransfer.Contract_DPA_ INNER JOIN
                                             dbo.InterTransfers ON dbo.Contract.Contract_DPA_ = dbo.InterTransfers.Contract_DPA_ INNER JOIN
                                             dbo.InterTransferType ON dbo.InterTransfer.InterTransferType_DPA_ = dbo.InterTransferType.InterTransferType_DPA_
                       WHERE     (dbo.InterTransfers.OrderTypeSale = 1) AND intertransfer.deleted = 0 and Year(dbo.InterTransfer.TransferDate)>2006 --and dbo.InterTransfers.Client_DPA_ = @ClientID
                       UNION ALL
                       -- Do not show  forwards---
		SELECT     Offerings.Client_DPA_, CAST(FLOOR(CAST(Offerings.Offerings_Date AS float)) AS DateTime) AS TransDate, Offerings.PAL_No AS REF,
		                     RTRIM(CAST(CAST(Offerings.Alloted_Rights AS int) AS char)) + ' ' + CASE WHEN len(Security.SecurityName)
		                     > 20 THEN LEFT(Security.SecurityName, 20) + '...' ELSE Security.SecurityName END + ' @ ' + CAST(Offerings.Offering_Price AS char(10))
		                     AS Particulars, Offerings.Alloted_Rights * Offerings.Offering_Price AS Debit, 0 AS Credit, 0 AS CreditBal, 0 AS IsOpeningbal,
		                     '' AS PaymentReceiptNo, - 1 AS ReceiptType
		FROM         Offerings INNER JOIN
		                     Security ON Offerings.Offering = Security.Security_DPA_
		WHERE     (Offerings.Deleted = 0)  AND (Security.Security_DPA_ = 99) and Year(Offerings.Offerings_Date)>2006
		OR (Offerings.Deleted = 0)  AND  (Offerings.Forward =0) AND (Security.Security_DPA_ <> 99) and Year(Offerings.Offerings_Date)>2006 --and Offerings.Client_DPA_ = @ClientID
                                              --AND CAST(FLOOR(CAST(Security.ClosingDate AS float)) AS DateTime) <= '13-Apr-2006'
                       UNION ALL
			SELECT     Offerings.Client_DPA_,
                      CAST(FLOOR(CAST(Offerings.Offerings_Date AS float)) AS Datetime) AS TransDate, cast(Offering_DPA_ as nvarchar) + '  ' + Offerings.PAL_No AS Ref,
                      'CDS CHARGE ' + CASE WHEN len(ltrim(rtrim(SecurityName))) > 20 THEN LEFT(ltrim(rtrim(SecurityName)), 20)
                      + '... ' ELSE ltrim(rtrim(SecurityName)) END + '@' + CAST(Offerings.Offering_Price AS char(10)) AS Particulars,  Offerings.CDSCharge AS Debit, 0 AS Credit,
                      0 AS CreditBal, 0 AS IsOpeningBal, Offerings.PAL_No AS PaymentReceiptNo, - 1 AS ReceiptType
		FROM         Offerings INNER JOIN
                      Client ON Offerings.Client_DPA_ = Client.Client_DPA_ INNER JOIN
                      Security ON Offerings.Offering = Security.Security_DPA_
		WHERE     Offerings.Deleted = 0 and isnull(Certificate,0)=0 and isnull(Offerings.CDSCharge,0)<>0

		       UNION ALL
                       SELECT     TOP 100 PERCENT dbo.Client.Client_DPA_, dbo.CPortfolio.CPortfolioPDate AS TransDate, dbo.CPortfolio.Reference AS ref,
                       ISNULL(dbo.CPortfolio.narrative, CAST(dbo.CPortfolio.CPortfolioQty AS nvarchar) + ' ' + dbo.Security.SecurityName) AS Particulars, 0 AS Debit,
                       0 AS Credit, 0 AS CreditBal, 0 AS IsOpeningBal, '' AS PaymentReceiptNo, 0 AS ReceiptType
		       FROM         dbo.Security INNER JOIN
			                      dbo.Client INNER JOIN
			                      dbo.CPortfolio ON dbo.Client.Client_DPA_ = dbo.CPortfolio.Client_DPA_ ON dbo.Security.Security_DPA_ = dbo.CPortfolio.Security_DPA_
		       WHERE     (dbo.Client.Deleted <> 1) AND (dbo.CPortfolio.Deleted =0) and Year(dbo.CPortfolio.CPortfolioPDate)>2006 --and dbo.Client.Client_DPA_ = @ClientID
) ClientsStatement

*/


SET NOCOUNT ON
delete from tblStatementList where Year(TransDate)>2006 --where Client_DPA_ = @ClientID
insert tblStatementList
SELECT     ClientsStatement.*, CreditBal - Debit AS Balance
FROM         (
                       SELECT     dbo.Payment.Entity_DPA_ AS Client_DPA_, dbo.Payment.PaymentPDate AS TransDate, dbo.Payment.PaymentReference AS Ref,
                                             dbo.Payment.PaymentNarrative AS Particulars, CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS CreditBal, 0 AS IsOpeningBalance,
                                             PaymentClientStatementNos.ReceiptNo, PaymentClientStatementNos.type AS receipttype
                       FROM         dbo.Payment INNER JOIN
                                             dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_ INNER JOIN
                                             (SELECT     Payment_DPA_, CONVERT(char(5), PaymentReceiptNo) AS ReceiptNo, 1 AS type
FROM         dbo.Payment
WHERE     (NOT (PaymentReceiptNo IS NULL)) AND (EntityType_DPA_ <> 8) AND Deleted = 0 and Year(dbo.Payment.PaymentPDate)>2006 --and Entity_DPA_=@ClientID
UNION
SELECT     dbo.Payment.Payment_DPA_, LotList.ContractNumber, 2 AS Type
FROM         dbo.Payment INNER JOIN
                      dbo.Voucher ON dbo.Payment.Voucher_DPA_ = dbo.Voucher.Voucher_DPA_ INNER JOIN
                      dbo.Contract ON dbo.Voucher.Voucher_DPA_ = dbo.Contract.Voucher_DPA_ INNER JOIN
                      (SELECT     Contract.Voucher_DPA_, Lot.ContractNumber, tbOrder.InterBank,Lot.Contract_DPA_
FROM         tbOrder INNER JOIN
                      OrdDetail ON tbOrder.Order_DPA_ = OrdDetail.Order_DPA_ INNER JOIN
                      Lot ON OrdDetail.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN
                      Contract ON Lot.Contract_DPA_ = Contract.Contract_DPA_
WHERE     (tbOrder.OrderCanceled = 0) AND (tbOrder.OrderHold = 0) AND (tbOrder.Deleted = 0) AND (OrdDetail.Deleted = 0) AND
                      (Lot.Deleted = 0) AND (Contract.Deleted = 0) AND (NOT (Contract.Voucher_DPA_ IS NULL))) LotList ON dbo.Contract.Contract_DPA_ = LotList.Contract_DPA_
WHERE     lotlist.interbank <> 1 AND Payment.Deleted = 0 and Year(dbo.Payment.PaymentPDate)>2006 --and payment.entity_DPA_=@ClientID
UNION
SELECT     dbo.Payment.Payment_DPA_, LotList.ContractNumber, 2 AS Type
FROM         dbo.BrokerReceiptVoucher INNER JOIN
                      dbo.Payment ON dbo.BrokerReceiptVoucher.BrokerReceiptVoucher_DPA_ = dbo.Payment.BrokerReceiptVoucher_DPA_ INNER JOIN
                      (SELECT     Contract.BrokerReceiptVoucher_DPA_, Lot.ContractNumber, tbOrder.InterBank,Lot.Contract_DPA_
FROM         tbOrder INNER JOIN
                      OrdDetail ON tbOrder.Order_DPA_ = OrdDetail.Order_DPA_ INNER JOIN
                      Lot ON OrdDetail.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN
                      Contract ON Lot.Contract_DPA_ = Contract.Contract_DPA_
WHERE     (tbOrder.OrderCanceled = 0) AND (tbOrder.OrderHold = 0) AND (tbOrder.Deleted = 0)  AND (OrdDetail.Deleted = 0) AND
                      (Lot.Deleted = 0) AND (Contract.Deleted = 0) AND (NOT (Contract.BrokerReceiptVoucher_DPA_ IS NULL))) LotList INNER JOIN
                      dbo.Contract ON LotList.Contract_DPA_ = dbo.Contract.Contract_DPA_ ON
                      dbo.BrokerReceiptVoucher.BrokerReceiptVoucher_DPA_ = dbo.Contract.BrokerReceiptVoucher_DPA_
WHERE     lotlist.interbank <> 1 AND Payment.Deleted = 0 AND Contract.Deleted = 0 and Year(dbo.Payment.PaymentPDate)>2006 --and payment.entity_DPA_=@ClientID
UNION
SELECT     dbo.Payment.Payment_DPA_, dbo.Payment.PaymentReference, 3 AS Type
FROM         dbo.Payment INNER JOIN
                      dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
WHERE     (dbo.Payment.PayType_DPA_ = 2) AND (dbo.Payment.Voucher_DPA_ IS NULL) AND Payment.Deleted = 0 and Year(dbo.Payment.PaymentPDate)>2006) PaymentClientStatementNos ON Payment.Payment_DPA_ = PaymentClientStatementNos.Payment_DPA_
                       WHERE     EntityType_DPA_ = 1 AND Payment.deleted = 0 and Year(dbo.Payment.PaymentPDate)>2006 --and Payment.Entity_DPA_=@ClientID
                       UNION ALL
                       SELECT     Client_DPA_, LotTDate AS TransDate, ContractNumber AS Ref, SecurityName AS Particulars,
                                             CASE OrderTypeSale WHEN 0 THEN NetAmt ELSE 0 END AS Debit, CASE OrderTypeSale WHEN 1 THEN NetAmt ELSE 0 END AS Credit,
                                             CASE OrderTypeSale WHEN 1 THEN NetAmt ELSE 0 END AS CreditBal, 0 AS IsOpeningBalance, ContractNumber AS PaymentReceiptNo,
                                             2 AS receipttype
                       FROM         dbo.ClientTransactionsSubList
					where Year(LotTDate)>2006
                       UNION  ALL
                       SELECT     dbo.JournalList.Entity_DPA_ AS Client_DPA_, dbo.JournalList.JournalDate AS TransDate,
                                             CAST(dbo.JournalList.JournalEntry_DPA_ AS Nvarchar(500)) AS Ref, dbo.JournalList.JournalNarrative AS Particulars,
                                             JournalList.JournalEntryDebit AS Debit, JournalList.JournalEntryCredit AS Credit, JournalList.JournalEntryCredit AS CreditBal,
                                             0 AS IsOpeningBalance, '' AS PaymentReceiptNo, 4 AS receipttype
                       FROM         dbo.JournalList
                       WHERE     EntityType_DPA_ = 1 and Year(dbo.JournalList.JournalDate) >2006 --and dbo.JournalList.Entity_DPA_ = @ClientID
                       UNION ALL
                       SELECT     dbo.InterTransfers.Client_DPA_, dbo.InterTransfer.TransferDate, dbo.InterTransfer.TransferReference AS REF,
                                             isnull(dbo.InterTransfer.TransferNarrative, '') + '  ' + InterTransferType.TypeDescription AS Particulars, 0 AS Debit,
                                             dbo.InterTransfer.TransferAmount AS Credit, dbo.InterTransfer.TransferAmount AS CreditBal, 0 AS IsOpeningBalance,
                                             dbo.InterTransfers.ContractNumber AS PaymentReceiptNo, 5 AS receipttype
                       FROM         dbo.Contract INNER JOIN
                                             dbo.InterTransfer ON dbo.Contract.Contract_DPA_ = dbo.InterTransfer.Contract_DPA_ INNER JOIN
                                             dbo.InterTransfers ON dbo.Contract.Contract_DPA_ = dbo.InterTransfers.Contract_DPA_ INNER JOIN
                                             dbo.InterTransferType ON dbo.InterTransfer.InterTransferType_DPA_ = dbo.InterTransferType.InterTransferType_DPA_
                       WHERE     (dbo.InterTransfers.OrderTypeSale = 0) AND intertransfer.deleted = 0 and Year(dbo.InterTransfer.TransferDate)>2006 --and dbo.InterTransfers.Client_DPA_=@ClientID
                       UNION ALL
                       SELECT     dbo.InterTransfers.Client_DPA_, dbo.InterTransfer.TransferDate, dbo.InterTransfer.TransferReference AS REF,
                                             isnull(dbo.InterTransfer.TransferNarrative, '') + '  ' + InterTransferType.TypeDescription AS Particulars,
                                             dbo.InterTransfer.TransferAmount AS Debit, 0 AS Credit, 0 AS CreditBal, 0 AS IsOpeningBalance,
                                             dbo.InterTransfers.ContractNumber AS PaymentReceiptNo, 5 AS receipttype
                       FROM         dbo.Contract INNER JOIN
                                             dbo.InterTransfer ON dbo.Contract.Contract_DPA_ = dbo.InterTransfer.Contract_DPA_ INNER JOIN
                                             dbo.InterTransfers ON dbo.Contract.Contract_DPA_ = dbo.InterTransfers.Contract_DPA_ INNER JOIN
                                             dbo.InterTransferType ON dbo.InterTransfer.InterTransferType_DPA_ = dbo.InterTransferType.InterTransferType_DPA_
                       WHERE     (dbo.InterTransfers.OrderTypeSale = 1) AND intertransfer.deleted = 0 and Year(dbo.InterTransfer.TransferDate)>2006 --and dbo.InterTransfers.Client_DPA_ = @ClientID
                       UNION ALL
                       -- Do not show  forwards---
		/*SELECT     Offerings.Client_DPA_, CAST(FLOOR(CAST(Offerings.Offerings_Date AS float)) AS DateTime) AS TransDate, Offerings.PAL_No AS REF,
                      RTRIM(CAST(CAST(isnull(Offerings.Additional,0) + Offerings.Alloted_Rights AS int) AS char)) + ' ' + CASE WHEN len(Security.SecurityName)
                      > 20 THEN LEFT(Security.SecurityName, 20) + '...' ELSE Security.SecurityName END + ' @ ' + CAST(Offerings.Offering_Price AS char(10)) AS Particulars,
                       (isnull(Offerings.Additional,0) + Offerings.Alloted_Rights) * Offerings.Offering_Price AS Debit, 0 AS Credit, 0 AS CreditBal, 0 AS IsOpeningbal,
                      '' AS PaymentReceiptNo, - 1 AS ReceiptType
		FROM         Offerings INNER JOIN*/
		SELECT     Offerings.Client_DPA_, CAST(FLOOR(CAST(Offerings.Offerings_Date AS float)) AS DateTime) AS TransDate, Offerings.PAL_No AS REF,
		                     RTRIM(CAST(CAST(Offerings.Alloted_Rights AS int) AS char)) + ' ' + CASE WHEN len(Security.SecurityName)
		                     > 20 THEN LEFT(Security.SecurityName, 20) + '...' ELSE Security.SecurityName END + ' @ ' + CAST(Offerings.Offering_Price AS char(10))
		                     AS Particulars, Offerings.Alloted_Rights * Offerings.Offering_Price AS Debit, 0 AS Credit, 0 AS CreditBal, 0 AS IsOpeningbal,
		                     '' AS PaymentReceiptNo, - 1 AS ReceiptType
		FROM         Offerings INNER JOIN
		                     Security ON Offerings.Offering = Security.Security_DPA_
		WHERE     (Offerings.Deleted = 0)  AND (Security.Security_DPA_ = 99) and Year(Offerings.Offerings_Date)>2006
		OR (Offerings.Deleted = 0)  AND  (Offerings.Forward =0) AND (Security.Security_DPA_ <> 99) and Year(Offerings.Offerings_Date)>2006 --and Offerings.Client_DPA_ = @ClientID
                                              --AND CAST(FLOOR(CAST(Security.ClosingDate AS float)) AS DateTime) <= '13-Apr-2006'
                       UNION ALL
			SELECT     Offerings.Client_DPA_,
                      CAST(FLOOR(CAST(Offerings.Offerings_Date AS float)) AS Datetime) AS TransDate, cast(Offering_DPA_ as nvarchar) + '  ' + Offerings.PAL_No AS Ref,
                      'CDS CHARGE ' + CASE WHEN len(ltrim(rtrim(SecurityName))) > 20 THEN LEFT(ltrim(rtrim(SecurityName)), 20)
                      + '... ' ELSE ltrim(rtrim(SecurityName)) END + '@' + CAST(Offerings.Offering_Price AS char(10)) AS Particulars,  Offerings.CDSCharge AS Debit, 0 AS Credit,
                      0 AS CreditBal, 0 AS IsOpeningBal, Offerings.PAL_No AS PaymentReceiptNo, - 1 AS ReceiptType
		FROM         Offerings INNER JOIN
                      Client ON Offerings.Client_DPA_ = Client.Client_DPA_ INNER JOIN
                      Security ON Offerings.Offering = Security.Security_DPA_
		WHERE     Offerings.Deleted = 0 and isnull(Certificate,0)=0 and isnull(Offerings.CDSCharge,0)<>0

		       UNION ALL
                       SELECT     TOP 100 PERCENT dbo.Client.Client_DPA_, dbo.CPortfolio.CPortfolioPDate AS TransDate, dbo.CPortfolio.Reference AS ref,
                       ISNULL(dbo.CPortfolio.narrative, CAST(dbo.CPortfolio.CPortfolioQty AS nvarchar) + ' ' + dbo.Security.SecurityName) AS Particulars, 0 AS Debit,
                       0 AS Credit, 0 AS CreditBal, 0 AS IsOpeningBal, '' AS PaymentReceiptNo, 0 AS ReceiptType
		       FROM         dbo.Security INNER JOIN
			                      dbo.Client INNER JOIN
			                      dbo.CPortfolio ON dbo.Client.Client_DPA_ = dbo.CPortfolio.Client_DPA_ ON dbo.Security.Security_DPA_ = dbo.CPortfolio.Security_DPA_
		       WHERE     (dbo.Client.Deleted <> 1) AND (dbo.CPortfolio.Deleted =0) and Year(dbo.CPortfolio.CPortfolioPDate)>2006 --and dbo.Client.Client_DPA_ = @ClientID
) ClientsStatement
GO


--ClientStatementProcBrief 13
CREATE      PROCEDURE ClientStatementProcBrief @ClientID int
AS


delete from tblStatementList where cast(floor(cast(TransDate as float)) as Datetime)=cast(floor(cast(getdate()-30 as float )) as Datetime) and Client_DPA_ = @ClientID
insert tblStatementList

SELECT     ClientsStatement.*, CreditBal - Debit AS Balance
FROM         (
 SELECT     dbo.Payment.Entity_DPA_ AS Client_DPA_, dbo.Payment.PaymentPDate AS TransDate, dbo.Payment.PaymentReference AS Ref,
                                             dbo.Payment.PaymentNarrative AS Particulars, CASE PayType.PayTypeIn WHEN 0 THEN Payment.PaymentAmount ELSE 0 END AS Debit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS Credit,
                                             CASE PayType.PayTypeIn WHEN 1 THEN Payment.PaymentAmount ELSE 0 END AS CreditBal, 0 AS IsOpeningBalance,
                                             PaymentClientStatementNos.ReceiptNo, PaymentClientStatementNos.type AS receipttype
                       FROM         dbo.Payment INNER JOIN
                                             dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_ INNER JOIN
                                             (SELECT     Payment_DPA_, CONVERT(char(5), PaymentReceiptNo) AS ReceiptNo, 1 AS type
FROM         dbo.Payment
WHERE     (NOT (PaymentReceiptNo IS NULL)) AND (EntityType_DPA_ <> 8) AND Deleted = 0 and cast(floor(cast(dbo.Payment.PaymentPDate as float)) as Datetime)=cast(floor(cast(getdate()-30 as float )) as Datetime)and Entity_DPA_=@ClientID
UNION
SELECT     dbo.Payment.Payment_DPA_, LotList.ContractNumber, 2 AS Type
FROM         dbo.Payment INNER JOIN
                      dbo.Voucher ON dbo.Payment.Voucher_DPA_ = dbo.Voucher.Voucher_DPA_ INNER JOIN
                      dbo.Contract ON dbo.Voucher.Voucher_DPA_ = dbo.Contract.Voucher_DPA_ INNER JOIN
                      (SELECT     Contract.Voucher_DPA_, Lot.ContractNumber, tbOrder.InterBank,Lot.Contract_DPA_
FROM         tbOrder INNER JOIN
                      OrdDetail ON tbOrder.Order_DPA_ = OrdDetail.Order_DPA_ INNER JOIN
                      Lot ON OrdDetail.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN
                      Contract ON Lot.Contract_DPA_ = Contract.Contract_DPA_
WHERE     (tbOrder.OrderCanceled = 0) AND (tbOrder.OrderHold = 0) AND (tbOrder.Deleted = 0) AND (OrdDetail.Deleted = 0) AND
                      (Lot.Deleted = 0) AND (Contract.Deleted = 0) AND (NOT (Contract.Voucher_DPA_ IS NULL))) LotList ON dbo.Contract.Contract_DPA_ = LotList.Contract_DPA_
WHERE     lotlist.interbank <> 1 AND Payment.Deleted = 0 and cast(floor(cast(dbo.Payment.PaymentPDate as float)) as Datetime)=cast(floor(cast(getdate()-30 as float )) as Datetime)and payment.entity_DPA_=@ClientID
UNION
SELECT     dbo.Payment.Payment_DPA_, LotList.ContractNumber, 2 AS Type
FROM         dbo.BrokerReceiptVoucher INNER JOIN
                      dbo.Payment ON dbo.BrokerReceiptVoucher.BrokerReceiptVoucher_DPA_ = dbo.Payment.BrokerReceiptVoucher_DPA_ INNER JOIN
                      (SELECT     Contract.BrokerReceiptVoucher_DPA_, Lot.ContractNumber, tbOrder.InterBank,Lot.Contract_DPA_
FROM         tbOrder INNER JOIN
                      OrdDetail ON tbOrder.Order_DPA_ = OrdDetail.Order_DPA_ INNER JOIN
                      Lot ON OrdDetail.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN
                      Contract ON Lot.Contract_DPA_ = Contract.Contract_DPA_
WHERE     (tbOrder.OrderCanceled = 0) AND (tbOrder.OrderHold = 0) AND (tbOrder.Deleted = 0)  AND (OrdDetail.Deleted = 0) AND
                      (Lot.Deleted = 0) AND (Contract.Deleted = 0) AND (NOT (Contract.BrokerReceiptVoucher_DPA_ IS NULL))) LotList INNER JOIN
                      dbo.Contract ON LotList.Contract_DPA_ = dbo.Contract.Contract_DPA_ ON
                      dbo.BrokerReceiptVoucher.BrokerReceiptVoucher_DPA_ = dbo.Contract.BrokerReceiptVoucher_DPA_
WHERE     lotlist.interbank <> 1 AND Payment.Deleted = 0 AND Contract.Deleted = 0 and cast(floor(cast(dbo.Payment.PaymentPDate as float)) as Datetime)=cast(floor(cast(getdate()-30 as float )) as Datetime)and payment.entity_DPA_=@ClientID
UNION
SELECT     dbo.Payment.Payment_DPA_, dbo.Payment.PaymentReference, 3 AS Type
FROM         dbo.Payment INNER JOIN
                      dbo.PayType ON dbo.Payment.PayType_DPA_ = dbo.PayType.PayType_DPA_
WHERE     (dbo.Payment.PayType_DPA_ = 2) AND (dbo.Payment.Voucher_DPA_ IS NULL) AND Payment.Deleted = 0 and Year(dbo.Payment.PaymentPDate)>2006) PaymentClientStatementNos ON Payment.Payment_DPA_ = PaymentClientStatementNos.Payment_DPA_
                       WHERE     EntityType_DPA_ = 1 AND Payment.deleted = 0 and cast(floor(cast(dbo.Payment.PaymentPDate as float)) as Datetime)=cast(floor(cast(getdate()-30 as float )) as Datetime) and Payment.Entity_DPA_=@ClientID
                       UNION ALL
                       SELECT     Client_DPA_, LotTDate AS TransDate, ContractNumber AS Ref, SecurityName AS Particulars,
                                             CASE OrderTypeSale WHEN 0 THEN NetAmt ELSE 0 END AS Debit, CASE OrderTypeSale WHEN 1 THEN NetAmt ELSE 0 END AS Credit,
                                             CASE OrderTypeSale WHEN 1 THEN NetAmt ELSE 0 END AS CreditBal, 0 AS IsOpeningBalance, ContractNumber AS PaymentReceiptNo,
                                             2 AS receipttype
                       FROM         dbo.ClientTransactionsSubList
					where cast(floor(cast(LotTDate as float)) as Datetime)=cast(floor(cast(getdate()-30 as float )) as Datetime)
                       UNION  ALL
                       SELECT     dbo.JournalList.Entity_DPA_ AS Client_DPA_, dbo.JournalList.JournalDate AS TransDate,
                                             CAST(dbo.JournalList.JournalEntry_DPA_ AS Nvarchar(500)) AS Ref, dbo.JournalList.JournalNarrative AS Particulars,
                                             JournalList.JournalEntryDebit AS Debit, JournalList.JournalEntryCredit AS Credit, JournalList.JournalEntryCredit AS CreditBal,
                                             0 AS IsOpeningBalance, '' AS PaymentReceiptNo, 4 AS receipttype
                       FROM         dbo.JournalList
                       WHERE     EntityType_DPA_ = 1 and cast(floor(cast(dbo.JournalList.JournalDate as float)) as Datetime)=cast(floor(cast(getdate()-30 as float )) as Datetime) and dbo.JournalList.Entity_DPA_ = @ClientID
                       UNION ALL
                       SELECT     dbo.InterTransfers.Client_DPA_, dbo.InterTransfer.TransferDate, dbo.InterTransfer.TransferReference AS REF,
                                             isnull(dbo.InterTransfer.TransferNarrative, '') + '  ' + InterTransferType.TypeDescription AS Particulars, 0 AS Debit,
                                             dbo.InterTransfer.TransferAmount AS Credit, dbo.InterTransfer.TransferAmount AS CreditBal, 0 AS IsOpeningBalance,
                                             dbo.InterTransfers.ContractNumber AS PaymentReceiptNo, 5 AS receipttype
                       FROM         dbo.Contract INNER JOIN
                                             dbo.InterTransfer ON dbo.Contract.Contract_DPA_ = dbo.InterTransfer.Contract_DPA_ INNER JOIN
                                             dbo.InterTransfers ON dbo.Contract.Contract_DPA_ = dbo.InterTransfers.Contract_DPA_ INNER JOIN
                                             dbo.InterTransferType ON dbo.InterTransfer.InterTransferType_DPA_ = dbo.InterTransferType.InterTransferType_DPA_
                       WHERE     (dbo.InterTransfers.OrderTypeSale = 0) AND intertransfer.deleted = 0 and cast(floor(cast(dbo.InterTransfer.TransferDate as float)) as Datetime)=cast(floor(cast(getdate()-30 as float )) as Datetime) and dbo.InterTransfers.Client_DPA_=@ClientID
                       UNION ALL
                       SELECT     dbo.InterTransfers.Client_DPA_, dbo.InterTransfer.TransferDate, dbo.InterTransfer.TransferReference AS REF,
                                             isnull(dbo.InterTransfer.TransferNarrative, '') + '  ' + InterTransferType.TypeDescription AS Particulars,
                                             dbo.InterTransfer.TransferAmount AS Debit, 0 AS Credit, 0 AS CreditBal, 0 AS IsOpeningBalance,
                                             dbo.InterTransfers.ContractNumber AS PaymentReceiptNo, 5 AS receipttype
                       FROM         dbo.Contract INNER JOIN
                                             dbo.InterTransfer ON dbo.Contract.Contract_DPA_ = dbo.InterTransfer.Contract_DPA_ INNER JOIN
                                             dbo.InterTransfers ON dbo.Contract.Contract_DPA_ = dbo.InterTransfers.Contract_DPA_ INNER JOIN
                                             dbo.InterTransferType ON dbo.InterTransfer.InterTransferType_DPA_ = dbo.InterTransferType.InterTransferType_DPA_
                       WHERE     (dbo.InterTransfers.OrderTypeSale = 1) AND intertransfer.deleted = 0 and cast(floor(cast(dbo.InterTransfer.TransferDate as float)) as Datetime)=cast(floor(cast(getdate()-30 as float )) as Datetime) and dbo.InterTransfers.Client_DPA_ = @ClientID
                       UNION ALL
                       -- Do not show  forwards---
		/*SELECT     Offerings.Client_DPA_, CAST(FLOOR(CAST(Offerings.Offerings_Date AS float)) AS DateTime) AS TransDate, Offerings.PAL_No AS REF,
                      RTRIM(CAST(CAST(isnull(Offerings.Additional,0) + Offerings.Alloted_Rights AS int) AS char)) + ' ' + CASE WHEN len(Security.SecurityName)
                      > 20 THEN LEFT(Security.SecurityName, 20) + '...' ELSE Security.SecurityName END + ' @ ' + CAST(Offerings.Offering_Price AS char(10)) AS Particulars,
                       (isnull(Offerings.Additional,0) + Offerings.Alloted_Rights) * Offerings.Offering_Price AS Debit, 0 AS Credit, 0 AS CreditBal, 0 AS IsOpeningbal,
                      '' AS PaymentReceiptNo, - 1 AS ReceiptType
		FROM         Offerings INNER JOIN*/
		SELECT     Offerings.Client_DPA_, CAST(FLOOR(CAST(Offerings.Offerings_Date AS float)) AS DateTime) AS TransDate, Offerings.PAL_No AS REF,
		                     RTRIM(CAST(CAST(Offerings.Alloted_Rights AS int) AS char)) + ' ' + CASE WHEN len(Security.SecurityName)
		                     > 20 THEN LEFT(Security.SecurityName, 20) + '...' ELSE Security.SecurityName END + ' @ ' + CAST(Offerings.Offering_Price AS char(10))
		                     AS Particulars, Offerings.Alloted_Rights * Offerings.Offering_Price AS Debit, 0 AS Credit, 0 AS CreditBal, 0 AS IsOpeningbal,
		                     '' AS PaymentReceiptNo, - 1 AS ReceiptType
		FROM         Offerings INNER JOIN
		                     Security ON Offerings.Offering = Security.Security_DPA_
		WHERE     (Offerings.Deleted = 0)  AND (Security.Security_DPA_ = 99) and cast(floor(cast(Offerings.Offerings_Date as float)) as Datetime)=cast(floor(cast(getdate()-30 as float )) as Datetime)
		OR (Offerings.Deleted = 0)  AND  (Offerings.Forward =0) AND (Security.Security_DPA_ <> 99) and cast(floor(cast(Offerings.Offerings_Date as float)) as Datetime)=cast(floor(cast(getdate()-30 as float )) as Datetime) and Offerings.Client_DPA_ = @ClientID
                                              --AND CAST(FLOOR(CAST(Security.ClosingDate AS float)) AS DateTime) <= '13-Apr-2006'
                       UNION ALL
			SELECT     Offerings.Client_DPA_,
                      CAST(FLOOR(CAST(Offerings.Offerings_Date AS float)) AS Datetime) AS TransDate, cast(Offering_DPA_ as nvarchar) + '  ' + Offerings.PAL_No AS Ref,
                      'CDS CHARGE ' + CASE WHEN len(ltrim(rtrim(SecurityName))) > 20 THEN LEFT(ltrim(rtrim(SecurityName)), 20)
                      + '... ' ELSE ltrim(rtrim(SecurityName)) END + '@' + CAST(Offerings.Offering_Price AS char(10)) AS Particulars,  Offerings.CSDCharge AS Debit, 0 AS Credit,
                      0 AS CreditBal, 0 AS IsOpeningBal, Offerings.PAL_No AS PaymentReceiptNo, - 1 AS ReceiptType
		FROM         Offerings INNER JOIN
                      Client ON Offerings.Client_DPA_ = Client.Client_DPA_ INNER JOIN
                      Security ON Offerings.Offering = Security.Security_DPA_
		WHERE     Offerings.Deleted = 0 and isnull(Certificate,0)=0 and isnull(Offerings.CSDCharge,0)<>0

		       UNION ALL
                       SELECT     TOP 100 PERCENT dbo.Client.Client_DPA_, dbo.CPortfolio.CPortfolioPDate AS TransDate, dbo.CPortfolio.Reference AS ref,
                       ISNULL(dbo.CPortfolio.narrative, CAST(dbo.CPortfolio.CPortfolioQty AS nvarchar) + ' ' + dbo.Security.SecurityName) AS Particulars, 0 AS Debit,
                       0 AS Credit, 0 AS CreditBal, 0 AS IsOpeningBal, '' AS PaymentReceiptNo, 0 AS ReceiptType
		       FROM         dbo.Security INNER JOIN
			                      dbo.Client INNER JOIN
			                      dbo.CPortfolio ON dbo.Client.Client_DPA_ = dbo.CPortfolio.Client_DPA_ ON dbo.Security.Security_DPA_ = dbo.CPortfolio.Security_DPA_
		       WHERE     (dbo.Client.Deleted <> 1) AND (dbo.CPortfolio.Deleted =0) and cast(floor(cast(dbo.CPortfolio.CPortfolioPDate as float)) as Datetime)=cast(floor(cast(getdate()-30 as float )) as Datetime) and dbo.Client.Client_DPA_ = @ClientID
) ClientsStatement



GO


CREATE PROCEDURE ClientTotalProcedure @clientDPA int
AS
--Update  ClientTotal SET Total = ClientTotals.Total from ClientTotal inner join ClientTotals on ClientTotal.client_DPA_=ClientTotals.Client_DPA_ where ClientTotals.Client_DPA_=@clientDPA

Delete From ClientTotal where Client_DPA_=@clientDPA
Insert Into ClientTotal(Client_DPA_,Total)

Select Client_DPA_,Sum(Total) as Total
from
(
SELECT     Client_DPA_, SUM(ISNULL(BalanceQty * OrdDetailPrice, 0)) AS Total
FROM         (SELECT     ordtbl.BalanceQty, CASE (dbo.OrdDetail.Best) WHEN 1 THEN dbo.datastream_SecurityPriceList.Price * 1.020825 ELSE CONVERT(float,
                      dbo.OrdDetail.OrdDetailPrice) END AS OrdDetailPrice, dbo.tbOrder.Order_DPA_, dbo.tbOrder.Client_DPA_, dbo.tbOrder.OrderCanceled
FROM         dbo.OrdDetail INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                          (SELECT     CASE ISNULL(OrdDetailContractedQtyList.ContractQty, 1)
                                                   WHEN 1 THEN dbo.OrdDetail.OrdDetailQty ELSE dbo.OrdDetail.OrdDetailQty - OrdDetailContractedQtyList.ContractQty END AS BalanceQty,
                                                    dbo.OrdDetail.OrdDetail_DPA_
                            FROM          dbo.OrdDetail LEFT OUTER JOIN
                                                   (SELECT     dbo.OrdDetail.OrdDetail_DPA_, SUM(dbo.Lot.LotQty) AS ContractQty
FROM         dbo.OrdDetail INNER JOIN
                      dbo.Lot ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.Lot.OrdDetail_DPA_
inner join tbOrder on OrdDetail.Order_DPA_=tbOrder.Order_DPA_
WHERE     (dbo.Lot.Deleted = 0) AND (dbo.OrdDetail.Deleted = 0) and  tbOrder.Client_DPA_=@clientDPA
GROUP BY dbo.OrdDetail.OrdDetail_DPA_) OrdDetailContractedQtyList ON dbo.OrdDetail.OrdDetail_DPA_ = OrdDetailContractedQtyList.OrdDetail_DPA_
                            WHERE      (dbo.OrdDetail.Deleted = 0)) ordtbl ON ordtbl.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN
                      dbo.datastream_SecurityPriceList ON dbo.OrdDetail.Security_DPA_ = dbo.datastream_SecurityPriceList.Security_DPA_
WHERE     (dbo.OrdDetail.Deleted = 0) AND (dbo.tbOrder.Deleted = 0) AND (dbo.tbOrder.OrderType_DPA_ = 1) AND (ordtbl.BalanceQty > 0) AND
                   (dbo.tbOrder.OrderCanceled = 0)  AND dbo.tbOrder.Client_DPA_=@clientDPA) a
GROUP BY Client_DPA_
union all
SELECT     Client_DPA_, SUM(PaymentAmount) AS Clienttotal
FROM         PaymentRequests
WHERE     (Processed_DPA_ IS NULL) AND (Deleted = 0) and Client_DPA_=@clientDPA
GROUP BY Client_DPA_) a
group by Client_DPA_


GO

CREATE PROCEDURE ClientTotalsDelete
AS
Delete From ClientTotal

GO

CREATE PROCEDURE ClientTotalsProcedure
AS
Insert Into ClientTotal(Client_DPA_,Total) Select Client_DPA_,Total From ClientTotals


GO




CREATE proc [dbo].[cont_CreateContract]

@OrdDetail_DPA_ int,
@Broker_DPA_ int,
@LotPrice money,
@LotQty money,
@LotSlipNo varchar(20),
@LotTDateDesc varchar(20),
@ContractSettlementDateDesc varchar(20),
@ChangedBy int

as

/*
Declare @ChangedBy int
set @ChangedBy = 12

Declare @OrdDetail_DPA_ int
set @OrdDetail_DPA_ = 12

Declare @Broker_DPA_ int
set @Broker_DPA_ = 3

Declare @LotPrice money
set @LotPrice = 100

Declare @LotQty money
set @LotQty = 100

Declare @LotSlipNo varchar(20)
set @LotSlipNo = 'Slip'

Declare @LotTDateDesc varchar(20)
set @LotTDateDesc = '7-July-2008'

Declare @ContractSettlementDateDesc varchar(20)
set @ContractSettlementDateDesc = '13-July-2008'

-- End Variables
*/

Declare @LotTDate datetime
set @LotTDate = convert(datetime, @LotTDateDesc)

Declare @ContractSettlementDate datetime
set @ContractSettlementDate = convert(datetime, @ContractSettlementDateDesc)

Declare @UniqueID uniqueidentifier
set @UniqueID = (SELECT     NEWID() AS uniqueID)



/*
Default values to be changed on the contract table
Status_DPA_ 1
*/

INSERT INTO Contract
                      (Contract_EIT_, ContractSettlementDate)
SELECT
@UniqueID AS Contract_EIT_,
@ContractSettlementDate AS ContractSettlementDate

Declare @Contract_DPA_ int
set @Contract_DPA_ = (SELECT Contract_DPA_ FROM  Contract WHERE (Contract_EIT_ = @UniqueID))

Declare @Side as char
set @Side = (SELECT     LEFT(OrderType.OrderTypeDescription, 1) AS Side FROM OrdDetail INNER JOIN tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN OrderType ON tbOrder.OrderType_DPA_ = OrderType.OrderType_DPA_ WHERE (OrdDetail.OrdDetail_DPA_ = @OrdDetail_DPA_))

Declare @ContractNumber as varchar(30)
set @ContractNumber = @Side + convert(varchar(20), @Contract_DPA_)

Declare @LotGross as money
set @LotGross = round(@LotQty * @LotPrice,2)

INSERT INTO Lot
(Contract_DPA_, OrdDetail_DPA_, LotPrice, LotQty, LotSlipNo, LotTDate, Broker_DPA_, ContractNumber, LotGrossAmount, ChangedBy,
ContractSettlementDate, TimeChanged)
SELECT
@Contract_DPA_ as Contract_DPA_,
@OrdDetail_DPA_ as OrdDetail_DPA_,
@LotPrice as LotPrice,
@LotQty as LotQty,
@LotSlipNo as LotSlipNo,
@LotTDate as LotTDate,
@Broker_DPA_ as Broker_DPA_,
@ContractNumber as ContractNumber,
@LotGross as LotGrossAmount,
@ChangedBy as ChangedBy,
@ContractSettlementDate as ContractSettlementDate,
getdate() as TimeChanged

-- Insert to Levy Contract
Declare @SystemMaintained int
Declare @LevyDescription varchar(30)
Declare @LevyRate float
Declare @BrokerAmount float
Declare @LevyShortName varchar(30)
Declare @LevyAmount money
Declare @VatableAmount money
Declare @LowerRate float
Declare @UpperRate float
Declare @GrossBoundary float
Declare @BrokerCommission float
Declare @LevyVATAmount float

set @VatableAmount = 0



-- Insert Broker Commission
set @SystemMaintained = 11
set @LevyDescription = 'Broker Commission'
set @LevyShortName = 'Commission'
set @LowerRate = 0
set @UpperRate = 0
set @GrossBoundary = 0

INSERT INTO LevyContract
(Contract_DPA_, LevyAmount, LevyName, LevyRate, LevyBlock, LevyShortName, LevyRatePercentage, SystemMaintained,
ChangedBy, TimeChanged,LevyVATAmount)
SELECT
@Contract_DPA_ as Contract_DPA_,
0 as LevyAmount,
@LevyDescription as LevyName,
0 as LevyRate,
0 as LevyBlock,
@LevyShortName as LevyShortName,
convert(varchar(10),@LevyRate) + '%' as LevyRatePercentage,
@SystemMaintained as SystemMaintained,
@ChangedBy as ChangedBy,
getdate() as TimeChanged,
0 as LevyVATAmount




-- Insert MSE Commission
set @SystemMaintained = 25
set @LevyDescription = ltrim(rtrim((SELECT TOP 1 CommissionDescription FROM Commission WHERE (SystemMaintained = @SystemMaintained))))
set @LevyShortName = ltrim(rtrim((SELECT TOP 1 'MSEComm' AS LevyShortName FROM Commission WHERE (SystemMaintained = @SystemMaintained))))
set @LowerRate = 0
set @UpperRate = 0
set @GrossBoundary = 0
set @LevyRate = 0
set @LevyAmount = 0
set @VatableAmount = 0
set @LevyVATAmount = 0


INSERT INTO LevyContract
(Contract_DPA_, LevyAmount, LevyName, LevyRate, LevyBlock, LevyShortName, LevyRatePercentage, SystemMaintained,
ChangedBy, TimeChanged,LevyVATAmount)
SELECT
@Contract_DPA_ as Contract_DPA_,
@LevyAmount as LevyAmount,
@LevyDescription as LevyName,
@LevyRate as LevyRate,
0 as LevyBlock,
@LevyShortName as LevyShortName,
convert(varchar(10),@LevyRate) + '%' as LevyRatePercentage,
@SystemMaintained as SystemMaintained,
@ChangedBy as ChangedBy,
getdate() as TimeChanged,
@LevyVATAmount as LevyVATAmount




-- Insert Agent Commission
set @SystemMaintained = 12
set @LevyDescription = 'Agent Commission'
set @LevyShortName = 'Agent'
set @LevyRate = 0
set @LevyAmount = 0

INSERT INTO LevyContract
(Contract_DPA_, LevyAmount, LevyName, LevyRate, LevyBlock, LevyShortName, LevyRatePercentage, SystemMaintained,
ChangedBy, TimeChanged)
SELECT
@Contract_DPA_ as Contract_DPA_,
@LevyAmount as LevyAmount,
@LevyDescription as LevyName,
@LevyRate as LevyRate,
0 as LevyBlock,
@LevyShortName as LevyShortName,
convert(varchar(10),@LevyRate) + '%' as LevyRatePercentage,
@SystemMaintained as SystemMaintained,
@ChangedBy as ChangedBy,
getdate() as TimeChanged



-- Insert Handling Fee
set @SystemMaintained = 100
set @LevyDescription = ltrim(rtrim((SELECT TOP 1 LevyDescription FROM Levy WHERE (SystemMaintained = 100))))
set @LevyShortName = ltrim(rtrim((SELECT TOP 1  LevyShortName FROM Levy WHERE (SystemMaintained = 100))))
set @LowerRate = 0
set @LevyRate = 0
set @LevyAmount = 0
set @VatableAmount = 0
set @LevyVATAmount = 0

INSERT INTO LevyContract
(Contract_DPA_, LevyAmount, LevyName, LevyRate, LevyBlock, LevyShortName, LevyRatePercentage, SystemMaintained,
ChangedBy, TimeChanged,LevyVATAmount)
SELECT
@Contract_DPA_ as Contract_DPA_,
@LevyAmount as LevyAmount,
@LevyDescription as LevyName,
@LevyRate as LevyRate,
0 as LevyBlock,
@LevyShortName as LevyShortName,
'0'as LevyRatePercentage,
@SystemMaintained as SystemMaintained,
@ChangedBy as ChangedBy,
getdate() as TimeChanged,
@LevyVATAmount as LevyVATAmount


-- Insert VAT
--vat id 16.5 % of broker commission
set @SystemMaintained = 99
set @LevyDescription = LTRIM(RTRIM((SELECT TOP 1 LevyDescription FROM Levy WHERE (SystemMaintained = @SystemMaintained))))
set @LevyShortName = LTRIM(RTRIM(UPPER((SELECT TOP 1 LevyShortName FROM Levy WHERE (SystemMaintained = @SystemMaintained)))))
set @LevyRate = round(isnull(
	(
		SELECT   TOP 1  (SELECT     LevyAmount  FROM Levy WHERE  (SystemMaintained = @SystemMaintained)) AS CDSLevyRate
	),0),2)
set @LevyAmount = 0

INSERT INTO LevyContract
(Contract_DPA_, LevyAmount, LevyName, LevyRate, LevyBlock, LevyShortName, LevyRatePercentage, SystemMaintained,
ChangedBy, TimeChanged)
SELECT
@Contract_DPA_ as Contract_DPA_,
@LevyAmount as LevyAmount,
@LevyDescription as LevyName,
@LevyRate as LevyRate,
0 as LevyBlock,
@LevyShortName as LevyShortName,
convert(varchar(10),@LevyRate) + '%' as LevyRatePercentage,
@SystemMaintained as SystemMaintained,
@ChangedBy as ChangedBy,
getdate() as TimeChanged


------------------------------------------------------------------
-- CGT (Capital Gains Tax) — SystemMaintained = 101, SALE side only
-------------------------------------------------------------------
SET @SystemMaintained = 101
SET @LevyDescription  = LTRIM(RTRIM((SELECT TOP 1 LevyDescription
                                       FROM dbo.Levy
                                      WHERE SystemMaintained = @SystemMaintained)))
SET @LevyShortName    = LTRIM(RTRIM((SELECT TOP 1 LevyShortName
                                       FROM dbo.Levy
                                      WHERE SystemMaintained = @SystemMaintained)))
SET @LevyRate         = ISNULL((SELECT TOP 1 LevyAmount
                                  FROM dbo.Levy
                                 WHERE SystemMaintained = @SystemMaintained), 0)
SET @LevyAmount       = ROUND(@LotGross * @LevyRate / 100.0, 2)
SET @LevyVATAmount    = 0

IF @Side = 'S'
BEGIN
    INSERT INTO dbo.LevyContract
        (Contract_DPA_, LevyAmount, LevyName, LevyRate, LevyBlock, LevyShortName,
         LevyRatePercentage, SystemMaintained, ChangedBy, TimeChanged, LevyVATAmount)
    SELECT
         @Contract_DPA_, @LevyAmount, @LevyDescription, @LevyRate, 0, @LevyShortName,
         CONVERT(varchar(10), @LevyRate) + '%',
         @SystemMaintained, @ChangedBy, GETDATE(), @LevyVATAmount
END



--Redo Broker Commissions so that the broker commission is re-calculated when there is more than
--one contract for the same security for the same day
exec cont_RedoBrokerCommissions @OrdDetail_DPA_, @LotTDate







GO








CREATE proc cont_CreateContract

@OrdDetail_DPA_ int,
@Broker_DPA_ int,
@LotPrice money,
@LotQty money,
@LotSlipNo varchar(20),
@LotTDateDesc varchar(20),
@ContractSettlementDateDesc varchar(20),
@ChangedBy int

as

/*
Declare @ChangedBy int
set @ChangedBy = 12

Declare @OrdDetail_DPA_ int
set @OrdDetail_DPA_ = 12

Declare @Broker_DPA_ int
set @Broker_DPA_ = 3

Declare @LotPrice money
set @LotPrice = 100

Declare @LotQty money
set @LotQty = 100

Declare @LotSlipNo varchar(20)
set @LotSlipNo = 'Slip'

Declare @LotTDateDesc varchar(20)
set @LotTDateDesc = '7-July-2008'

Declare @ContractSettlementDateDesc varchar(20)
set @ContractSettlementDateDesc = '13-July-2008'

-- End Variables
*/

Declare @LotTDate datetime
set @LotTDate = convert(datetime, @LotTDateDesc)

Declare @ContractSettlementDate datetime
set @ContractSettlementDate = convert(datetime, @ContractSettlementDateDesc)

Declare @UniqueID uniqueidentifier
set @UniqueID = (SELECT     NEWID() AS uniqueID)



/*
Default values to be changed on the contract table
Status_DPA_ 1
*/

INSERT INTO Contract
                      (Contract_EIT_, ContractSettlementDate)
SELECT
@UniqueID AS Contract_EIT_,
@ContractSettlementDate AS ContractSettlementDate

Declare @Contract_DPA_ int
set @Contract_DPA_ = (SELECT Contract_DPA_ FROM  Contract WHERE (Contract_EIT_ = @UniqueID))

Declare @Side as char
set @Side = (SELECT     LEFT(OrderType.OrderTypeDescription, 1) AS Side FROM OrdDetail INNER JOIN tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN OrderType ON tbOrder.OrderType_DPA_ = OrderType.OrderType_DPA_ WHERE (OrdDetail.OrdDetail_DPA_ = @OrdDetail_DPA_))

Declare @ContractNumber as varchar(30)
set @ContractNumber = @Side + convert(varchar(20), @Contract_DPA_)

Declare @LotGross as money
set @LotGross = round(@LotQty * @LotPrice,2)

INSERT INTO Lot
(Contract_DPA_, OrdDetail_DPA_, LotPrice, LotQty, LotSlipNo, LotTDate, Broker_DPA_, ContractNumber, LotGrossAmount, ChangedBy,
ContractSettlementDate, TimeChanged)
SELECT
@Contract_DPA_ as Contract_DPA_,
@OrdDetail_DPA_ as OrdDetail_DPA_,
@LotPrice as LotPrice,
@LotQty as LotQty,
@LotSlipNo as LotSlipNo,
@LotTDate as LotTDate,
@Broker_DPA_ as Broker_DPA_,
@ContractNumber as ContractNumber,
@LotGross as LotGrossAmount,
@ChangedBy as ChangedBy,
@ContractSettlementDate as ContractSettlementDate,
getdate() as TimeChanged

-- Insert to Levy Contract
Declare @SystemMaintained int
Declare @LevyDescription varchar(30)
Declare @LevyRate float
Declare @BrokerAmount float
Declare @LevyShortName varchar(30)
Declare @LevyAmount money
Declare @VatableAmount money
Declare @LowerRate float
Declare @UpperRate float
Declare @GrossBoundary float
Declare @BrokerCommission float
Declare @LevyVATAmount float

set @VatableAmount = 0

/*
-- Insert CSD Levy
set @SystemMaintained = 98
set @LevyDescription = LTRIM(RTRIM((SELECT TOP 1 LevyDescription FROM Levy WHERE (SystemMaintained = @SystemMaintained))))
set @LevyShortName = LTRIM(RTRIM(UPPER((SELECT TOP 1 LevyShortName FROM Levy WHERE (SystemMaintained = @SystemMaintained)))))
set @LevyRate = round(isnull(
	(
		SELECT   TOP 1  (SELECT     LevyAmount  FROM Levy WHERE  (SystemMaintained = @SystemMaintained)) AS CDSLevyRate
		FROM         OrdDetail INNER JOIN  Security ON OrdDetail.Security_DPA_ = Security.Security_DPA_
		WHERE     (OrdDetail.OrdDetail_DPA_ = @OrdDetail_DPA_) AND (Security.IsOnCSD = 1)
	),0),2)
set @LevyAmount = dbo.cont_Round05 (@LevyRate * @LotGross / 100)
set @VatableAmount = @VatableAmount + @LevyAmount

INSERT INTO LevyContract
(Contract_DPA_, LevyAmount, LevyName, LevyRate, LevyBlock, LevyShortName, LevyRatePercentage, SystemMaintained,
ChangedBy, TimeChanged)
SELECT
@Contract_DPA_ as Contract_DPA_,
@LevyAmount as LevyAmount,
@LevyDescription as LevyName,
@LevyRate as LevyRate,
0 as LevyBlock,
@LevyShortName as LevyShortName,
convert(varchar(10),@LevyRate) + '%' as LevyRatePercentage,
@SystemMaintained as SystemMaintained,
@ChangedBy as ChangedBy,
getdate() as TimeChanged
FROM         OrdDetail INNER JOIN  Security ON OrdDetail.Security_DPA_ = Security.Security_DPA_
WHERE     (OrdDetail.OrdDetail_DPA_ = @OrdDetail_DPA_) AND (Security.IsOnCSD = 1)

*/



-- Insert Broker Commission
set @SystemMaintained = 11
set @LevyDescription = 'Broker Commission'
set @LevyShortName = 'Commission'
set @LowerRate = (
SELECT     Commission.CommissionRate
FROM         tbOrder INNER JOIN
                      OrdDetail ON tbOrder.Order_DPA_ = OrdDetail.Order_DPA_ INNER JOIN
                      Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
                      Commission ON Client.Commission_DPA_ = Commission.Commission_DPA_
WHERE     (OrdDetail.OrdDetail_DPA_ = @OrdDetail_DPA_)
)
set @UpperRate = (
SELECT     Commission.UpperSecurityCommission
FROM         tbOrder INNER JOIN
                      OrdDetail ON tbOrder.Order_DPA_ = OrdDetail.Order_DPA_ INNER JOIN
                      Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
                      Commission ON Client.Commission_DPA_ = Commission.Commission_DPA_
WHERE     (OrdDetail.OrdDetail_DPA_ = @OrdDetail_DPA_)
)
set @GrossBoundary = (
SELECT     Commission.SecurityBoundary
FROM         tbOrder INNER JOIN
                      OrdDetail ON tbOrder.Order_DPA_ = OrdDetail.Order_DPA_ INNER JOIN
                      Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
                      Commission ON Client.Commission_DPA_ = Commission.Commission_DPA_
WHERE     (OrdDetail.OrdDetail_DPA_ = @OrdDetail_DPA_)
)


--Apply appropriate Levy percentage
--2% on the first MWK 50 000
--1,5% on the next MWK 50 000
--1% over MWK 100 000
Declare @RunningBrokerComm money
Declare @Band1 money
Declare @Band2 money
Declare @Band3 money




--Get the first Band
if @LotGross > 50000
begin
	set @Band1 = 50000
	set @LevyRate =2
end
else
begin
	set @Band1 = @LotGross
	set @LevyRate =2
end


--set @LevyRate =2
set @RunningBrokerComm = dbo.cont_Round05(@Band1 * (@LevyRate/100))
/*
print cast(@LotGross as varchar)
print 'Band1 - '+ cast(@Band1 as varchar)
print '@RunningBrokerComm1  - ' + cast(@RunningBrokerComm as varchar)
print ''*/
--Get the second Band
if  @LotGross > 100000
begin
	set @Band2 = 50000
	set @LevyRate =1.5
end
else
begin
	if @LotGross > 50000
	begin
		set @Band2 = @LotGross - @Band1
		set @LevyRate =1.5
	end
	else
	begin
		set @Band2 = 0
	end
end
--set @LevyRate =1.5
set @RunningBrokerComm =@RunningBrokerComm + dbo.cont_Round05( @Band2 * (@LevyRate/100))

/*
print 'Band2 - '+ cast(@Band2 as nvarchar)
print '@RunningBrokerComm2  - ' + cast(@RunningBrokerComm as varchar)
print ''
*/


if @LotGross > 100000
begin
	set @Band3 = @LotGross -(@Band1 + @Band2)
	set @LevyRate =1
end
else
begin
	set @Band3 = 0
	set @LevyRate =1
end

--set @LevyRate =1
set @RunningBrokerComm = @RunningBrokerComm + dbo.cont_Round05( @Band3 * (@LevyRate/100))
--set @LevyRate =(@RunningBrokerComm/@LotGross) * 100



--set @LevyRate = (case when @LotGross < @GrossBoundary then @LowerRate else @UpperRate end)
set @LevyAmount = @RunningBrokerComm

--recalculate Broker Commission
set @levyRate=(@RunningBrokerComm/@LotGross)*100

--Set Minimum amount to 50 MWK if levy amount is Less than 50
if @LevyAmount <50
begin
	set @LevyAmount=50
	--set @LevyRate='minimum'
end


set @BrokerAmount = @LevyAmount --Get  Broker Amount so as to calaculate MSE
set @VatableAmount = @VatableAmount + @LevyAmount
set @BrokerCommission = @LevyAmount
set @LevyVATAmount = @LevyAmount*(round(isnull(
	(
		SELECT   TOP 1  (SELECT     LevyAmount  FROM Levy WHERE  (SystemMaintained = 99)) AS CDSLevyRate
	),0),2)/100)

INSERT INTO LevyContract
(Contract_DPA_, LevyAmount, LevyName, LevyRate, LevyBlock, LevyShortName, LevyRatePercentage, SystemMaintained,
ChangedBy, TimeChanged,LevyVATAmount)
SELECT
@Contract_DPA_ as Contract_DPA_,
@LevyAmount as LevyAmount,
@LevyDescription as LevyName,
@LevyRate as LevyRate,
0 as LevyBlock,
@LevyShortName as LevyShortName,
convert(varchar(10),@LevyRate) + '%' as LevyRatePercentage,
@SystemMaintained as SystemMaintained,
@ChangedBy as ChangedBy,
getdate() as TimeChanged,
@LevyVATAmount as LevyVATAmount




-- Insert MSE Commission
set @SystemMaintained = 25
set @LevyDescription = ltrim(rtrim((SELECT TOP 1 CommissionDescription FROM Commission WHERE (SystemMaintained = @SystemMaintained))))
set @LevyShortName = ltrim(rtrim((SELECT TOP 1 'MSEComm' AS LevyShortName FROM Commission WHERE (SystemMaintained = @SystemMaintained))))
set @LowerRate = (SELECT TOP 1 CommissionRate FROM Commission WHERE (SystemMaintained = @SystemMaintained))
set @UpperRate = (SELECT TOP 1 UpperSecurityCommission FROM Commission WHERE (SystemMaintained = @SystemMaintained))
set @GrossBoundary = isnull((SELECT TOP 1 SecurityBoundary FROM Commission WHERE (SystemMaintained = @SystemMaintained)),0)
--set @LevyRate = (case when @LotGross < @GrossBoundary then @LowerRate else @UpperRate end)
set @LevyRate = (SELECT TOP 1 CommissionRate FROM Commission WHERE (SystemMaintained = @SystemMaintained))

set @LevyAmount = dbo.cont_Round05 (@LevyRate * @BrokerAmount / 100)
set @VatableAmount = @VatableAmount + @LevyAmount
set @LevyVATAmount = @LevyAmount*(round(isnull(
	(
		SELECT   TOP 1  (SELECT     LevyAmount  FROM Levy WHERE  (SystemMaintained = 99)) AS CDSLevyRate
	),0),2)/100)

INSERT INTO LevyContract
(Contract_DPA_, LevyAmount, LevyName, LevyRate, LevyBlock, LevyShortName, LevyRatePercentage, SystemMaintained,
ChangedBy, TimeChanged,LevyVATAmount)
SELECT
@Contract_DPA_ as Contract_DPA_,
@LevyAmount as LevyAmount,
@LevyDescription as LevyName,
@LevyRate as LevyRate,
0 as LevyBlock,
@LevyShortName as LevyShortName,
convert(varchar(10),@LevyRate) + '%' as LevyRatePercentage,
@SystemMaintained as SystemMaintained,
@ChangedBy as ChangedBy,
getdate() as TimeChanged,
@LevyVATAmount as LevyVATAmount




-- Insert Agent Commission
set @SystemMaintained = 12
set @LevyDescription = 'Agent Commission'
set @LevyShortName = 'Agent'
set @LevyRate = isnull((
SELECT     Commission.CommissionRate
FROM         OrdDetail INNER JOIN
                      tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN
                      Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
                      Agent ON Client.Agent_DPA_ = Agent.Agent_DPA_ INNER JOIN
                      Commission ON Agent.Commission_DPA_ = Commission.Commission_DPA_
WHERE     (OrdDetail.OrdDetail_DPA_ = @OrdDetail_DPA_)
),0)
set @LevyAmount = dbo.cont_Round05 (@LevyRate * @BrokerCommission / 100)

INSERT INTO LevyContract
(Contract_DPA_, LevyAmount, LevyName, LevyRate, LevyBlock, LevyShortName, LevyRatePercentage, SystemMaintained,
ChangedBy, TimeChanged)
SELECT
@Contract_DPA_ as Contract_DPA_,
@LevyAmount as LevyAmount,
@LevyDescription as LevyName,
@LevyRate as LevyRate,
0 as LevyBlock,
@LevyShortName as LevyShortName,
convert(varchar(10),@LevyRate) + '%' as LevyRatePercentage,
@SystemMaintained as SystemMaintained,
@ChangedBy as ChangedBy,
getdate() as TimeChanged



-- Insert Handling Fee
set @SystemMaintained = 100
set @LevyDescription = ltrim(rtrim((SELECT TOP 1 LevyDescription FROM Levy WHERE (SystemMaintained = 100))))
set @LevyShortName = ltrim(rtrim((SELECT TOP 1  LevyShortName FROM Levy WHERE (SystemMaintained = 100))))
set @LowerRate = (SELECT TOP 1 LevyAmount FROM Levy WHERE (SystemMaintained = 100))
--set @UpperRate = (SELECT TOP 1 UpperSecurityCommission FROM Commission WHERE (SystemMaintained = @SystemMaintained))
--set @GrossBoundary = isnull((SELECT TOP 1 SecurityBoundary FROM Commission WHERE (SystemMaintained = @SystemMaintained)),0)
--set @LevyRate = (case when @LotGross < @GrossBoundary then @LowerRate else @UpperRate end)
set @LevyRate = (SELECT TOP 1 LevyAmount FROM levy WHERE (SystemMaintained = @SystemMaintained))

set @LevyAmount = dbo.cont_Round05 (@LevyRate )
set @VatableAmount = @VatableAmount + @LevyAmount
set @LevyVATAmount = @LevyAmount*(round(isnull(
	(
		SELECT   TOP 1  (SELECT     LevyAmount  FROM Levy WHERE  (SystemMaintained = 99)) AS CDSLevyRate
	),0),2)/100)

INSERT INTO LevyContract
(Contract_DPA_, LevyAmount, LevyName, LevyRate, LevyBlock, LevyShortName, LevyRatePercentage, SystemMaintained,
ChangedBy, TimeChanged,LevyVATAmount)
SELECT
@Contract_DPA_ as Contract_DPA_,
@LevyAmount as LevyAmount,
@LevyDescription as LevyName,
@LevyRate as LevyRate,
0 as LevyBlock,
@LevyShortName as LevyShortName,
'0'as LevyRatePercentage,
@SystemMaintained as SystemMaintained,
@ChangedBy as ChangedBy,
getdate() as TimeChanged,
@LevyVATAmount as LevyVATAmount


-- Insert VAT
--vat id 16.5 % of broker commission
set @SystemMaintained = 99
set @LevyDescription = LTRIM(RTRIM((SELECT TOP 1 LevyDescription FROM Levy WHERE (SystemMaintained = @SystemMaintained))))
set @LevyShortName = LTRIM(RTRIM(UPPER((SELECT TOP 1 LevyShortName FROM Levy WHERE (SystemMaintained = @SystemMaintained)))))
set @LevyRate = round(isnull(
	(
		SELECT   TOP 1  (SELECT     LevyAmount  FROM Levy WHERE  (SystemMaintained = @SystemMaintained)) AS CDSLevyRate
	),0),2)
set @LevyAmount = dbo.cont_Round05 ((@LevyRate * @BrokerAmount) / 100)

INSERT INTO LevyContract
(Contract_DPA_, LevyAmount, LevyName, LevyRate, LevyBlock, LevyShortName, LevyRatePercentage, SystemMaintained,
ChangedBy, TimeChanged)
SELECT
@Contract_DPA_ as Contract_DPA_,
@LevyAmount as LevyAmount,
@LevyDescription as LevyName,
@LevyRate as LevyRate,
0 as LevyBlock,
@LevyShortName as LevyShortName,
convert(varchar(10),@LevyRate) + '%' as LevyRatePercentage,
@SystemMaintained as SystemMaintained,
@ChangedBy as ChangedBy,
getdate() as TimeChanged


--Redo Broker Commissions so that the broker commission is re-calculated when there is more than
--one contract for the same security for the same day
exec cont_RedoBrokerCommissions @OrdDetail_DPA_, @LotTDate
exec cont_RedoBasicFee @OrdDetail_DPA_, @LotTDate




GO

CREATE Proc cont_EditBrokerComm
@ContractID bigint,@Rate float, @ChangeType  int ,@ChangedBy int

as
/*
set @ContractID = 33
set @Rate = 1.3
set @ChangeType = 1
set @ChangedBy = 10*/

--cont_EditBrokerComm 54 , 0.5,1,10

/*
*Please note:

Editing Broker Commission affects thing Broker Commission, AgentCommission, MSE Commission and the VAT
*/
--Redo BrokerComm


--Get the total gross for contracts generated for the specific counter for that day per order item
Declare @LotGross money
set @LotGross = (
		SELECT     SUM(Lot.LotGrossAmount) AS LotGrossAmount
		FROM         Lot INNER JOIN
		                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_
		WHERE     (Lot.Deleted <> 1)
		GROUP BY Lot.Contract_DPA_
		HAVING      (Lot.Contract_DPA_ = @ContractID)
)
--set @LotGross = null
if @LotGross is null
begin
	return
end


-- System maintained value for broker commission
Declare @SystemMaintained tinyint
Declare @LevyRate float
Declare @LevyRatePercentage varchar(50)
Declare @LevyAmount money
Declare @LevyVATAmount money
Declare @MSEComm money
declare @BrokerAmount money

set @SystemMaintained = 11

/*
@ChangeType = 1 then rate
@ChangeType = 1 then Broker Amount
*/

if @ChangeType = 1 --Rate
begin
	set @LevyRatePercentage = convert(varchar(20),@Rate) + ' % '
	set @LevyAmount = @LotGross * (@Rate/100)
	set @LevyRate = @Rate
	set @LevyVATAmount = round(@LevyAmount*(round(isnull(
		(
			SELECT   TOP 1  (SELECT     LevyAmount  FROM Levy WHERE  (SystemMaintained = 99)) AS CDSLevyRate
		),0),2)/100), 2)
	select @LotGross, @LevyAmount, @LevyVATAmount, @LevyVATAmount/@LevyAmount * 100
end
else
begin
	set @LevyRate = (@Rate/@LotGross)*100
	set @LevyRatePercentage = convert(varchar(20),round(@LevyRate,2)) + ' % '
	set @LevyAmount = @Rate
	set @LevyVATAmount = round(@LevyAmount*(round(isnull(
		(
			SELECT   TOP 1  (SELECT     LevyAmount  FROM Levy WHERE  (SystemMaintained = 99)) AS CDSLevyRate
		),0),2)/100), 2)
	select @LevyAmount, @LevyVATAmount, @LevyVATAmount/@LevyAmount * 100
end

-- Update LevyContract Items
set @BrokerAmount = @LevyAmount
UPDATE    LevyContract
SET             LevyAmount = dbo.cont_Round05(@LevyAmount),
		LevyVATAmount = dbo.cont_Round05(@LevyVATAmount),
		LevyRate =@LevyRate,
		LevyRatePercentage = @LevyRatePercentage,
		ChangedBy = @ChangedBy,
		TimeChanged = getdate()

FROM         Lot INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_
WHERE     (Lot.Deleted <> 1) AND (Lot.Contract_DPA_ = @ContractID) AND (LevyContract.SystemMaintained = @SystemMaintained) AND (LevyContract.Deleted <> 1)



--Redo MSEComm

set @SystemMaintained = 25
--Claculate MSEComm

set @Levyrate = (SELECT TOP 1 isnull(CommissionRate,0) FROM Commission WHERE (SystemMaintained = @SystemMaintained))
Set @MSEComm = @LevyAmount * (@Levyrate/100)
Set @LevyVATAmount = round(@MSEComm *(round(isnull(
		(
			SELECT   TOP 1  (SELECT     LevyAmount  FROM Levy WHERE  (SystemMaintained = 99)) AS CDSLevyRate
		),0),2)/100), 2)

UPDATE    LevyContract
SET             LevyAmount = dbo.cont_Round05(@MSEComm),
		LevyVATAmount = round(@LevyVATAmount, 2),
		ChangedBy = @ChangedBy,
		TimeChanged = getdate()
FROM         Lot INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_
WHERE     (Lot.Deleted <> 1) AND (Lot.Contract_DPA_ = @ContractID) AND (LevyContract.SystemMaintained = @SystemMaintained) AND (LevyContract.Deleted <> 1)
select @LevyAmount, @LevyVATAmount, @LevyVATAmount/@LevyAmount * 100


--Redo Agent Comm
set @SystemMaintained = 12

set @LevyRate = isnull((
SELECT     isnull(Commission.CommissionRate,0)
FROM         OrdDetail INNER JOIN
                      tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN
                      Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
                      Agent ON Client.Agent_DPA_ = Agent.Agent_DPA_ INNER JOIN
                      Commission ON Agent.Commission_DPA_ = Commission.Commission_DPA_ INNER JOIN
                      Lot ON OrdDetail.OrdDetail_DPA_ = Lot.OrdDetail_DPA_
WHERE     (Lot.Contract_DPA_ = @ContractID)),0)

set @LevyAmount = dbo.cont_Round05 (isnull(@LevyRate * (@BrokerAmount-@MSEComm) / 100,0))
Set @LevyVATAmount = round(@LevyAmount *(round(isnull(
		(
			SELECT   TOP 1  (SELECT     LevyAmount  FROM Levy WHERE  (SystemMaintained = 99)) AS CDSLevyRate
		),0),2)/100), 2)

UPDATE    LevyContract
SET             LevyAmount = dbo.cont_Round05(@LevyAmount),
		LevyVATAmount = dbo.cont_Round05(@LevyVATAmount),
		ChangedBy = @ChangedBy,
		TimeChanged = getdate()
FROM         Lot INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_
WHERE     (Lot.Deleted <> 1) AND (Lot.Contract_DPA_ = @ContractID) AND (LevyContract.SystemMaintained = @SystemMaintained) AND (LevyContract.Deleted <> 1)
select @LevyAmount, @LevyVATAmount




--Redo VAT
set @SystemMaintained = 99

set   @LevyAmount  = isnull((SELECT     SUM(LevyVATAmount) AS VAT
				FROM         LevyContract
				WHERE     (Contract_DPA_ = @ContractID) AND  (SystemMaintained = 11 OR
                      SystemMaintained = 100) AND (Deleted = 0)),0)
UPDATE    LevyContract
SET             LevyAmount = @LevyAmount,
		LevyVATAmount = 0,
		ChangedBy = @ChangedBy,
		TimeChanged = getdate()
FROM         Lot INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_
WHERE     (Lot.Deleted <> 1) AND (Lot.Contract_DPA_ = @ContractID) AND (LevyContract.SystemMaintained = @SystemMaintained) AND (LevyContract.Deleted <> 1)
select @LevyAmount As Vat



GO




--cont_RedoBasicFee 77, '3-Dec-2008'

--cont_RedoBasicFee 66, '27-Nov-2008'
CREATE proc cont_RedoBasicFee

@OrdDetail_DPA_ int,
@TradeDate smalldatetime

as

Declare @SystemMaintained tinyint
Declare @LevyRate float
Declare @LevyAmount money
Declare @LevyVATAmount money

/*
--Samples for passed variables
Declare @OrdDetail_DPA_ int
set @OrdDetail_DPA_ = 55

Declare @TradeDate smalldatetime
set @TradeDate = convert(smalldatetime, '5-Nov-2008')
*/


-- Recalculate Basic Fee
set @SystemMaintained = 100
set @LevyRate = (SELECT TOP 1 LevyAmount FROM levy WHERE (SystemMaintained = @SystemMaintained))

set @LevyAmount = @LevyRate
set @LevyVATAmount = @LevyAmount*(round(isnull(
	(
		SELECT   TOP 1  (SELECT     LevyAmount  FROM Levy WHERE  (SystemMaintained = 99)) AS CDSLevyRate
	),0),2)/100)


UPDATE    LevyContract
SET             LevyAmount = 0,
		LevyVATAmount = 0,
		LevyRate = 0,
		LevyRatePercentage = '0'
FROM         Lot INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_
WHERE
	(Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
	(Lot.Deleted <> 1) AND
	(CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
                      = CAST(FLOOR(CAST(@TradeDate AS float)) AS datetime)) AND
	(LevyContract.SystemMaintained = @SystemMaintained)
	AND (LevyContract.Deleted <> 1)


UPDATE    LevyContract
SET             LevyAmount = @LevyRate,
		LevyVATAmount = round(@LevyVATAmount,2),
		LevyRate = @LevyRate,
		LevyRatePercentage = '0'
Where LevyContract_DPA_ =
(
	Select Max (LevyContract_DPA_)
	FROM         Lot INNER JOIN
	                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
	                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_
	WHERE
		(Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
		(Lot.Deleted <> 1) AND
		(CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
	                      = CAST(FLOOR(CAST(@TradeDate AS float)) AS datetime)) AND
		(LevyContract.SystemMaintained = @SystemMaintained)
		AND (LevyContract.Deleted <> 1)
)

--Update the VAT Totals
UPDATE    LevyContract
SET              LevyAmount = round(VATTotals.LevyVATAmount,2)
FROM         (

	SELECT     Lot.Contract_DPA_, LevyContract.LevyContract_DPA_
	FROM          Lot INNER JOIN
	                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
	                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_
	WHERE
		(OrdDetail.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
		(cast(floor(cast(Lot.LotTDate AS float)) AS datetime) =
			cast(floor(cast(@TradeDate AS float)) AS datetime)) AND
		(LevyContract.SystemMaintained = 99) AND
		(LevyContract.Deleted <> 1) AND
		(Lot.Deleted <> 1)

)
                      VATEntries INNER JOIN
                      LevyContract ON VATEntries.LevyContract_DPA_ = LevyContract.LevyContract_DPA_ INNER JOIN
                          (

	SELECT     SUM(LevyContract.LevyVATAmount) AS LevyVATAmount, LevyContract.Contract_DPA_
	FROM         LevyContract INNER JOIN
	                      Lot ON LevyContract.Contract_DPA_ = Lot.Contract_DPA_
	WHERE
		(Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
		(CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
			= cast(floor(cast(@TradeDate AS float)) AS datetime)) AND
		(LevyContract.SystemMaintained = 11) AND
		(LevyContract.Deleted <> 1)
		OR
	        (Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
		(CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
			= cast(floor(cast(@TradeDate AS float)) AS datetime)) AND
		(LevyContract.SystemMaintained = 100) AND
		(LevyContract.Deleted <> 1)
	GROUP BY LevyContract.Contract_DPA_



			) VATTotals ON VATEntries.Contract_DPA_ = VATTotals.Contract_DPA_

GO




CREATE proc cont_RedoBasicFee

@OrdDetail_DPA_ int,
@TradeDate smalldatetime

as

Declare @SystemMaintained tinyint
Declare @LevyRate float
Declare @LevyAmount money
Declare @LevyVATAmount money

/*
--Samples for passed variables
Declare @OrdDetail_DPA_ int
set @OrdDetail_DPA_ = 55

Declare @TradeDate smalldatetime
set @TradeDate = convert(smalldatetime, '5-Nov-2008')
*/


-- Recalculate Basic Fee
set @SystemMaintained = 100
set @LevyRate = (SELECT TOP 1 LevyAmount FROM levy WHERE (SystemMaintained = @SystemMaintained))

set @LevyAmount = dbo.cont_Round05 (@LevyRate )
set @LevyVATAmount = @LevyAmount*(round(isnull(
	(
		SELECT   TOP 1  (SELECT     LevyAmount  FROM Levy WHERE  (SystemMaintained = 99)) AS CDSLevyRate
	),0),2)/100)


UPDATE    LevyContract
SET             LevyAmount = 0,
		LevyVATAmount = 0,
		LevyRate = 0,
		LevyRatePercentage = '0'
FROM         Lot INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_
WHERE
	(Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
	(Lot.Deleted <> 1) AND
	(CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
                      = CAST(FLOOR(CAST(@TradeDate AS float)) AS datetime)) AND
	(LevyContract.SystemMaintained = @SystemMaintained)
	AND (LevyContract.Deleted <> 1)


UPDATE    LevyContract
SET             LevyAmount = @LevyRate,
		LevyVATAmount = @LevyVATAmount,
		LevyRate = @LevyRate,
		LevyRatePercentage = '0'
Where LevyContract_DPA_ =
(
	Select Max (LevyContract_DPA_)
	FROM         Lot INNER JOIN
	                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
	                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_
	WHERE
		(Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
		(Lot.Deleted <> 1) AND
		(CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
	                      = CAST(FLOOR(CAST(@TradeDate AS float)) AS datetime)) AND
		(LevyContract.SystemMaintained = @SystemMaintained)
		AND (LevyContract.Deleted <> 1)
)


--Update the VAT Totals
UPDATE    LevyContract
SET              LevyAmount = VATTotals.LevyVATAmount
FROM         (

	SELECT     Lot.Contract_DPA_, LevyContract.LevyContract_DPA_
	FROM          Lot INNER JOIN
	                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
	                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_
	WHERE
		(OrdDetail.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
		(cast(floor(cast(Lot.LotTDate AS float)) AS datetime) =
			cast(floor(cast(@TradeDate AS float)) AS datetime)) AND
		(LevyContract.SystemMaintained = 99) AND
		(LevyContract.Deleted <> 1) AND
		(Lot.Deleted <> 1)

)
                      VATEntries INNER JOIN
                      LevyContract ON VATEntries.LevyContract_DPA_ = LevyContract.LevyContract_DPA_ INNER JOIN
                          (

	SELECT     SUM(LevyContract.LevyVATAmount) AS LevyVATAmount, LevyContract.Contract_DPA_
	FROM         LevyContract INNER JOIN
	                      Lot ON LevyContract.Contract_DPA_ = Lot.Contract_DPA_
	WHERE
		(Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
		(CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
			= cast(floor(cast(@TradeDate AS float)) AS datetime)) AND
		(LevyContract.SystemMaintained = 11) AND
		(LevyContract.Deleted <> 1)
		OR
	        (Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
		(CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
			= cast(floor(cast(@TradeDate AS float)) AS datetime)) AND
		(LevyContract.SystemMaintained = 100) AND
		(LevyContract.Deleted <> 1)
	GROUP BY LevyContract.Contract_DPA_



			) VATTotals ON VATEntries.Contract_DPA_ = VATTotals.Contract_DPA_





GO


--cont_RedoBrokerCommissions 77, '03-Dec-2008'


CREATE proc cont_RedoBrokerCommissions

@OrdDetail_DPA_ int,
@TradeDate smalldatetime

as


/*
--Samples for passed variables
Declare @OrdDetail_DPA_ int
set @OrdDetail_DPA_ = 77

Declare @TradeDate smalldatetime
set @TradeDate = convert(smalldatetime, '3-Dec-2008')
*/

--Get the total gross for contracts generated for the specific counter for that day per order item
Declare @TotalLotGross money
set @TotalLotGross = (
SELECT     SUM(Lot.LotGrossAmount) AS LotGrossAmount
FROM         Lot INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_
WHERE     (Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND (Lot.Deleted <> 1) AND (CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
                      = cast(floor(cast(@TradeDate AS float)) AS datetime))
)

-- System maintained value for broker commission
Declare @SystemMaintained tinyint
set @SystemMaintained = 11

--Apply appropriate Levy percentage
--2% on the first MWK 50 000
--1.5% on the next MWK 50 000
--1% over MWK 100 000
Declare @RunningBrokerCommAccurate money
Declare @Band1 money
Declare @Band2 money
Declare @Band3 money
Declare @LevyRate float
Declare @LevyRatePercentage varchar(50)
Declare @LevyAmount money
Declare @LevyVATAmount money
Declare @BrokerCommVATAmountAccurate money
Declare @MSEComm money

Declare @CommLowerRate float
set @CommLowerRate = isnull((
	SELECT  top 1     Commission.CommissionRate AS CommLowerRate
	FROM         tbOrder INNER JOIN
	                      OrdDetail ON tbOrder.Order_DPA_ = OrdDetail.Order_DPA_ INNER JOIN
	                      Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
	                      Commission ON Client.Commission_DPA_ = Commission.Commission_DPA_
	WHERE     (OrdDetail.OrdDetail_DPA_ = @OrdDetail_DPA_)
),0)

Declare @CommMiddleRate float
set @CommMiddleRate = isnull((
	SELECT  top 1     Commission.MedianSecurityCommission AS CommMiddleRate
	FROM         tbOrder INNER JOIN
	                      OrdDetail ON tbOrder.Order_DPA_ = OrdDetail.Order_DPA_ INNER JOIN
	                      Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
	                      Commission ON Client.Commission_DPA_ = Commission.Commission_DPA_
	WHERE     (OrdDetail.OrdDetail_DPA_ = @OrdDetail_DPA_)
),0)

Declare @CommUpperRate float
set @CommUpperRate = isnull((
	SELECT  top 1     Commission.UpperSecurityCommission AS CommUpperRate
	FROM         tbOrder INNER JOIN
	                      OrdDetail ON tbOrder.Order_DPA_ = OrdDetail.Order_DPA_ INNER JOIN
	                      Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
	                      Commission ON Client.Commission_DPA_ = Commission.Commission_DPA_
	WHERE     (OrdDetail.OrdDetail_DPA_ = @OrdDetail_DPA_)
),0)

Declare @EquityGrossAmountBoundaryLowerMiddle money
set @EquityGrossAmountBoundaryLowerMiddle = isnull((
	SELECT  top 1     Commission.SecurityBoundary
	FROM         tbOrder INNER JOIN
	                      OrdDetail ON tbOrder.Order_DPA_ = OrdDetail.Order_DPA_ INNER JOIN
	                      Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
	                      Commission ON Client.Commission_DPA_ = Commission.Commission_DPA_
	WHERE     (OrdDetail.OrdDetail_DPA_ = @OrdDetail_DPA_)
),0)

Declare @EquityGrossAmountBoundaryMiddleUpper money
set @EquityGrossAmountBoundaryMiddleUpper = isnull((
	SELECT  top 1  Commission.SecondSecurityBoundary
	FROM         tbOrder INNER JOIN
	                      OrdDetail ON tbOrder.Order_DPA_ = OrdDetail.Order_DPA_ INNER JOIN
	                      Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
	                      Commission ON Client.Commission_DPA_ = Commission.Commission_DPA_
	WHERE     (OrdDetail.OrdDetail_DPA_ = @OrdDetail_DPA_)
),0)

Declare @EquityCommMinimum money
set @EquityCommMinimum = isnull((
	SELECT  top 1  Commission.MinimumSecurityCommission
	FROM         tbOrder INNER JOIN
	                      OrdDetail ON tbOrder.Order_DPA_ = OrdDetail.Order_DPA_ INNER JOIN
	                      Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
	                      Commission ON Client.Commission_DPA_ = Commission.Commission_DPA_
	WHERE     (OrdDetail.OrdDetail_DPA_ = @OrdDetail_DPA_)
),0)




--Get the first Band
--2% on the first MWK 50 000
set @LevyRate = @CommLowerRate
set @LevyRatePercentage = convert(varchar(20), round(@EquityGrossAmountBoundaryLowerMiddle / 1000, 0)) + 'k: ' + convert(varchar(10),@LevyRate) + '%'
if @TotalLotGross > @EquityGrossAmountBoundaryLowerMiddle
begin
	set @Band1 = @EquityGrossAmountBoundaryLowerMiddle
end
else
begin
	set @Band1 = @TotalLotGross
end
set @RunningBrokerCommAccurate = round(@Band1 * (@LevyRate/100), 2)


--Get the second Band
--1.5% on the next MWK 50 000
set @LevyRate = @CommMiddleRate
set @LevyRatePercentage = @LevyRatePercentage + ', ' + convert(varchar(20), round(@EquityGrossAmountBoundaryMiddleUpper / 1000, 0)) + 'k: ' + convert(varchar(10),@LevyRate) + '%'
if  @TotalLotGross > @EquityGrossAmountBoundaryMiddleUpper
begin
	set @Band2 = @EquityGrossAmountBoundaryMiddleUpper - @EquityGrossAmountBoundaryLowerMiddle
end
else
begin
	if @TotalLotGross > @EquityGrossAmountBoundaryLowerMiddle
	begin
		set @Band2 = @TotalLotGross - @EquityGrossAmountBoundaryLowerMiddle
	end
	else
	begin
		set @Band2 = 0
	end
end
set @RunningBrokerCommAccurate = @RunningBrokerCommAccurate + round( @Band2 * (@LevyRate/100), 2)



--Get the third Band
--1% over MWK 100 000
set @LevyRate = @CommUpperRate
set @LevyRatePercentage = @LevyRatePercentage + ', ' + convert(varchar(10),@LevyRate) + '%'
if @TotalLotGross > @EquityGrossAmountBoundaryMiddleUpper
begin
	set @Band3 = @TotalLotGross - @EquityGrossAmountBoundaryMiddleUpper
end
else
begin
	set @Band3 = 0
end
set @RunningBrokerCommAccurate = @RunningBrokerCommAccurate + round( @Band3 * (@LevyRate/100), 2)

--Set Minimum amount to 50 MWK if levy amount is Less than 50
if @RunningBrokerCommAccurate < @EquityCommMinimum
begin
	set @RunningBrokerCommAccurate = @EquityCommMinimum
	set @LevyRate = 0
	set @LevyRatePercentage = 'Minimum'
end

set @LevyAmount = @RunningBrokerCommAccurate
set @LevyVATAmount = round(@LevyAmount*(round(isnull(
	(
		SELECT   TOP 1  (SELECT     LevyAmount  FROM Levy WHERE  (SystemMaintained = 99)) AS CDSLevyRate
	),0),2)/100), 2)
set @BrokerCommVATAmountAccurate = @LevyVATAmount

select @TotalLotGross, @LevyAmount, @LevyVATAmount, @LevyVATAmount/@LevyAmount * 100

-- Update LevyContract Items
UPDATE    LevyContract
SET             LevyAmount = round(round(Lot.LotGrossAmount / @TotalLotGross,2) * @LevyAmount,2),
		LevyVATAmount = round(round(Lot.LotGrossAmount / @TotalLotGross,2) * @LevyVATAmount,2),
		LevyRate = @LevyRate,
		LevyRatePercentage = @LevyRatePercentage
FROM         Lot INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_
WHERE
	(Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
	(Lot.Deleted <> 1) AND
	(CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
                      = CAST(FLOOR(CAST(@TradeDate AS float)) AS datetime)) AND
	(LevyContract.SystemMaintained = @SystemMaintained)
	AND (LevyContract.Deleted <> 1)


Declare @RunningBrokerComSumIndividual decimal(18,4)
set @RunningBrokerComSumIndividual = (
SELECT     SUM(LevyContract.LevyAmount) AS LevyAmount
FROM         Lot INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_
WHERE
	(Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
	(Lot.Deleted <> 1) AND
	(CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
                      = CAST(FLOOR(CAST(@TradeDate AS float)) AS datetime)) AND
	(LevyContract.SystemMaintained = @SystemMaintained)
	AND (LevyContract.Deleted <> 1)

)


select @RunningBrokerComSumIndividual as VV



Declare @BrokerComVATAmountSumIndividual money
set @BrokerComVATAmountSumIndividual = (
SELECT     SUM(LevyContract.LevyVATAmount) AS LevyVATAmount
FROM         Lot INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_
WHERE
	(Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
	(Lot.Deleted <> 1) AND
	(CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
                      = CAST(FLOOR(CAST(@TradeDate AS float)) AS datetime)) AND
	(LevyContract.SystemMaintained = @SystemMaintained)
	AND (LevyContract.Deleted <> 1)

)

select @RunningBrokerComSumIndividual, @BrokerComVATAmountSumIndividual

--print convert(varchar(20),@RunningBrokerComSumIndividual)
se
UPDATE    LevyContract
SET
	LevyAmount = round(LevyAmount + @RunningBrokerCommAccurate - @RunningBrokerComSumIndividual,2),
	LevyVATAmount = round(LevyVATAmount + @BrokerCommVATAmountAccurate - @BrokerComVATAmountSumIndividual,2 )
where Levycontract_dpa_ = (
select max (levycontract.levycontract_dpa_)
FROM         Lot INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_
WHERE
	(Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
	(Lot.Deleted <> 1) AND
	(CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
                      = CAST(FLOOR(CAST(@TradeDate AS float)) AS datetime)) AND
	(LevyContract.SystemMaintained = @SystemMaintained)
	AND (LevyContract.Deleted <> 1)
)




--Update MSE Commissions
set @SystemMaintained = 25
set @LevyRate = (SELECT TOP 1 isnull(LevyAmount,0) FROM Levy WHERE (SystemMaintained = @SystemMaintained))
set @LevyRatePercentage = convert(varchar(10),@LevyRate) + '%'
set @LevyAmount = round(@LevyRate / 100 * @RunningBrokerCommAccurate, 2)
set @LevyVATAmount = round(@LevyAmount*(round(isnull(
	(
		SELECT   TOP 1  (SELECT     LevyAmount  FROM Levy WHERE  (SystemMaintained = 99)) AS CDSLevyRate
	),0),2)/100), 2)


-- Update LevyContract for MSE Commissions
UPDATE    LevyContract
SET             LevyAmount = round(round(Lot.LotGrossAmount / @TotalLotGross,2) * @LevyAmount,2),
		LevyVATAmount = round(round(Lot.LotGrossAmount / @TotalLotGross,2) * @LevyVATAmount,2),
		LevyRate = @LevyRate,
		LevyRatePercentage = @LevyRatePercentage
FROM         Lot INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_
WHERE
	(Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
	(Lot.Deleted <> 1) AND
	(CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
                      = CAST(FLOOR(CAST(@TradeDate AS float)) AS datetime)) AND
	(LevyContract.SystemMaintained = @SystemMaintained)
	AND (LevyContract.Deleted <> 1)



--Update Agent Commissions
set @SystemMaintained = 12
set @LevyRate = isnull((
SELECT     Commission.CommissionRate
FROM         OrdDetail INNER JOIN
                      tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN
                      Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
                      Agent ON Client.Agent_DPA_ = Agent.Agent_DPA_ INNER JOIN
                      Commission ON Agent.Commission_DPA_ = Commission.Commission_DPA_
WHERE     (OrdDetail.OrdDetail_DPA_ = @OrdDetail_DPA_)
),0)
set @MSEComm = round(isnull((
	SELECT     sum(LevyContract.LevyAmount) as MSEComm
	FROM         LevyContract INNER JOIN
	                      Lot ON LevyContract.Contract_DPA_ = Lot.Contract_DPA_
	WHERE     (Lot.OrdDetail_DPA_ = @OrdDetail_DPA_)
		AND (CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
			= CAST(FLOOR(CAST(@TradeDate AS float)) AS datetime))
		AND (LevyContract.SystemMaintained = 25)
		AND (LevyContract.Deleted <> 1)
), 0),2)
set @LevyRatePercentage = convert(varchar(10),@LevyRate) + '%'
set @LevyAmount = round(@LevyRate / 100 * (@RunningBrokerCommAccurate - @MSEComm), 2)

-- Update LevyContract for Agent Commissions
UPDATE    LevyContract
SET             LevyAmount = round(round(Lot.LotGrossAmount / @TotalLotGross ,2)* @LevyAmount,2),
		LevyRate = @LevyRate,
		LevyRatePercentage = @LevyRatePercentage
FROM         Lot INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_
WHERE
	(Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
	(Lot.Deleted <> 1) AND
	(CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
                      = CAST(FLOOR(CAST(@TradeDate AS float)) AS datetime)) AND
	(LevyContract.SystemMaintained = @SystemMaintained)
	AND (LevyContract.Deleted <> 1)




-- Recalculate Basic Fee
set @SystemMaintained = 100
set @LevyRate = (SELECT TOP 1 LevyAmount FROM levy WHERE (SystemMaintained = @SystemMaintained))
set @LevyAmount = dbo.cont_Round05 (@LevyRate )
set @LevyVATAmount = @LevyAmount*(round(isnull(
	(
		SELECT   TOP 1  (SELECT     LevyAmount  FROM Levy WHERE  (SystemMaintained = 99)) AS CDSLevyRate
	),0),2)/100)


UPDATE    LevyContract
SET             LevyAmount = 0,
		LevyVATAmount = 0,
		LevyRate = 0,
		LevyRatePercentage = '0'
FROM         Lot INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_
WHERE
	(Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
	(Lot.Deleted <> 1) AND
	(CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
                      = CAST(FLOOR(CAST(@TradeDate AS float)) AS datetime)) AND
	(LevyContract.SystemMaintained = @SystemMaintained)
	AND (LevyContract.Deleted <> 1)


UPDATE    LevyContract
SET             LevyAmount = @LevyRate,
		LevyVATAmount = dbo.cont_Round05(@LevyVATAmount),
		LevyRate = @LevyRate,
		LevyRatePercentage = '0'
Where LevyContract_DPA_ =
(
	Select Max (LevyContract_DPA_)
	FROM         Lot INNER JOIN
	                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
	                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_
	WHERE
		(Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
		(Lot.Deleted <> 1) AND
		(CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
	                      = CAST(FLOOR(CAST(@TradeDate AS float)) AS datetime)) AND
		(LevyContract.SystemMaintained = @SystemMaintained)
		AND (LevyContract.Deleted <> 1)
)


--Update the VAT Totals
UPDATE    LevyContract
SET              LevyAmount = round(VATTotals.LevyVATAmount,2)
FROM         (

	SELECT     Lot.Contract_DPA_, LevyContract.LevyContract_DPA_
	FROM          Lot INNER JOIN
	                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
	                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_
	WHERE
		(OrdDetail.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
		(cast(floor(cast(Lot.LotTDate AS float)) AS datetime) =
			cast(floor(cast(@TradeDate AS float)) AS datetime)) AND
		(LevyContract.SystemMaintained = 99) AND
		(LevyContract.Deleted <> 1) AND
		(Lot.Deleted <> 1)

)
                      VATEntries INNER JOIN
                      LevyContract ON VATEntries.LevyContract_DPA_ = LevyContract.LevyContract_DPA_ INNER JOIN
                          (

	SELECT     SUM(LevyContract.LevyVATAmount) AS LevyVATAmount, LevyContract.Contract_DPA_
	FROM         LevyContract INNER JOIN
	                      Lot ON LevyContract.Contract_DPA_ = Lot.Contract_DPA_
	WHERE
		(Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
		(CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
			= cast(floor(cast(@TradeDate AS float)) AS datetime)) AND
		(LevyContract.SystemMaintained = 11) AND
		(LevyContract.Deleted <> 1)
		OR
	        (Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
		(CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
			= cast(floor(cast(@TradeDate AS float)) AS datetime)) AND
		(LevyContract.SystemMaintained = 100) AND
		(LevyContract.Deleted <> 1)
	GROUP BY LevyContract.Contract_DPA_



			) VATTotals ON VATEntries.Contract_DPA_ = VATTotals.Contract_DPA_



-- Make sure order has been marked as compounded
UPDATE    tbOrder
SET              OrderCompounded = 1
WHERE     (Order_DPA_ =
                          (SELECT     Order_DPA_
                            FROM          OrdDetail
                            WHERE      (OrdDetail_DPA_ = @OrdDetail_DPA_)))


GO

CREATE proc cont_RedoBrokerCommissions

@OrdDetail_DPA_ int,
@TradeDate smalldatetime

as

/*
--Samples for passed variables
Declare @OrdDetail_DPA_ int
set @OrdDetail_DPA_ = 13

Declare @TradeDate smalldatetime
set @TradeDate = convert(smalldatetime, '11-Nov-2008')
*/

--Get the total gross for contracts generated for the specific counter for that day per order item
Declare @LotGross money
set @LotGross = (
SELECT     SUM(Lot.LotGrossAmount) AS LotGrossAmount
FROM         Lot INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_
WHERE     (Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND (Lot.Deleted <> 1) AND (CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
                      = cast(floor(cast(@TradeDate AS float)) AS datetime))
)

-- System maintained value for broker commission
Declare @SystemMaintained tinyint
set @SystemMaintained = 11

--Apply appropriate Levy percentage
--2% on the first MWK 50 000
--1.5% on the next MWK 50 000
--1% over MWK 100 000
Declare @RunningBrokerComm money
Declare @Band1 money
Declare @Band2 money
Declare @Band3 money
Declare @LevyRate float
Declare @LevyRatePercentage varchar(50)
Declare @LevyAmount money
Declare @LevyVATAmount money
Declare @MSEComm money


--Get the first Band
--2% on the first MWK 50 000
set @LevyRate = 2
set @LevyRatePercentage = '50k: ' + convert(varchar(10),@LevyRate) + '%'
if @LotGross > 50000
begin
	set @Band1 = 50000
end
else
begin
	set @Band1 = @LotGross
end
set @RunningBrokerComm = round(@Band1 * (@LevyRate/100), 2)


--Get the second Band
--1.5% on the next MWK 50 000
set @LevyRate = 1.5
set @LevyRatePercentage = @LevyRatePercentage + ', 100k: ' + convert(varchar(10),@LevyRate) + '%'
if  @LotGross > 100000
begin
	set @Band2 = 50000
end
else
begin
	if @LotGross > 50000
	begin
		set @Band2 = @LotGross - @Band1
	end
	else
	begin
		set @Band2 = 0
	end
end
set @RunningBrokerComm = @RunningBrokerComm + round( @Band2 * (@LevyRate/100), 2)



--Get the third Band
--1% over MWK 100 000
set @LevyRate = 1
set @LevyRatePercentage = @LevyRatePercentage + ', ' + convert(varchar(10),@LevyRate) + '%'
if @LotGross > 100000
begin
	set @Band3 = @LotGross - (@Band1 + @Band2)
end
else
begin
	set @Band3 = 0
end
set @RunningBrokerComm = @RunningBrokerComm + round( @Band3 * (@LevyRate/100), 2)

--Set Minimum amount to 50 MWK if levy amount is Less than 50
if @LevyAmount < 50
begin
	set @LevyAmount = 50
	set @LevyRate = 0
	set @LevyRatePercentage = 'Minimum'
end

set @LevyAmount = @RunningBrokerComm
set @LevyVATAmount = round(@LevyAmount*(round(isnull(
	(
		SELECT   TOP 1  (SELECT     LevyAmount  FROM Levy WHERE  (SystemMaintained = 99)) AS CDSLevyRate
	),0),2)/100), 2)
set @LevyAmount = @RunningBrokerComm

select @LotGross, @LevyAmount, @LevyVATAmount, @LevyVATAmount/@LevyAmount * 100

-- Update LevyContract Items
UPDATE    LevyContract
SET             LevyAmount = round(Lot.LotGrossAmount / @LotGross * @LevyAmount, 2),
		LevyVATAmount = round(Lot.LotGrossAmount / @LotGross * @LevyVATAmount, 2),
		LevyRate = @LevyRate,
		LevyRatePercentage = @LevyRatePercentage
FROM         Lot INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_
WHERE
	(Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
	(Lot.Deleted <> 1) AND
	(CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
                      = CAST(FLOOR(CAST(@TradeDate AS float)) AS datetime)) AND
	(LevyContract.SystemMaintained = @SystemMaintained)
	AND (LevyContract.Deleted <> 1)


Declare @LevyAmountNew money
set @LevyAmountNew = (
SELECT     SUM(LevyContract.LevyAmount) AS LevyAmount
FROM         Lot INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_
WHERE
	(Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
	(Lot.Deleted <> 1) AND
	(CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
                      = CAST(FLOOR(CAST(@TradeDate AS float)) AS datetime)) AND
	(LevyContract.SystemMaintained = @SystemMaintained)
	AND (LevyContract.Deleted <> 1)

)

Declare @LevyVATAmountNew money
set @LevyVATAmountNew = (
SELECT     SUM(LevyContract.LevyVATAmount) AS LevyVATAmount
FROM         Lot INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_
WHERE
	(Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
	(Lot.Deleted <> 1) AND
	(CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
                      = CAST(FLOOR(CAST(@TradeDate AS float)) AS datetime)) AND
	(LevyContract.SystemMaintained = @SystemMaintained)
	AND (LevyContract.Deleted <> 1)

)

select @LevyAmountNew, @LevyVATAmountNew



UPDATE    LevyContract
SET
	LevyAmount = LevyAmount + dbo.cont_Round05(@LevyAmount) - @LevyAmountNew,
	LevyVATAmount = LevyVATAmount + dbo.cont_Round05(@LevyVATAmount) - @LevyVATAmountNew
where Levycontract_dpa_ = (
select max (levycontract.levycontract_dpa_)
FROM         Lot INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_
WHERE
	(Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
	(Lot.Deleted <> 1) AND
	(CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
                      = CAST(FLOOR(CAST(@TradeDate AS float)) AS datetime)) AND
	(LevyContract.SystemMaintained = @SystemMaintained)
	AND (LevyContract.Deleted <> 1)
)




--Update Agent Commissions
set @SystemMaintained = 12
set @LevyRate = isnull((
SELECT     Commission.CommissionRate
FROM         OrdDetail INNER JOIN
                      tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN
                      Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
                      Agent ON Client.Agent_DPA_ = Agent.Agent_DPA_ INNER JOIN
                      Commission ON Agent.Commission_DPA_ = Commission.Commission_DPA_
WHERE     (OrdDetail.OrdDetail_DPA_ = @OrdDetail_DPA_)
),0)
set @MSEComm = isnull((
	SELECT     sum(LevyContract.LevyAmount) as MSEComm
	FROM         LevyContract INNER JOIN
	                      Lot ON LevyContract.Contract_DPA_ = Lot.Contract_DPA_
	WHERE     (Lot.OrdDetail_DPA_ = @OrdDetail_DPA_)
		AND (CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
			= CAST(FLOOR(CAST(@TradeDate AS float)) AS datetime))
		AND (LevyContract.SystemMaintained = 25)
		AND (LevyContract.Deleted <> 1)
), 0)
set @LevyRatePercentage = convert(varchar(10),@LevyRate) + '%'
set @LevyAmount = round(@LevyRate / 100 * (@LevyAmount - @MSEComm), 2)

-- Update LevyContract for Agent Commissions
UPDATE    LevyContract
SET             LevyAmount = dbo.cont_Round05(Lot.LotGrossAmount / @LotGross * @LevyAmount),
		LevyRate = @LevyRate,
		LevyRatePercentage = @LevyRatePercentage
FROM         Lot INNER JOIN
                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_
WHERE
	(Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
	(Lot.Deleted <> 1) AND
	(CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
                      = CAST(FLOOR(CAST(@TradeDate AS float)) AS datetime)) AND
	(LevyContract.SystemMaintained = @SystemMaintained)
	AND (LevyContract.Deleted <> 1)

-- Make sure order has been marked as compounded
UPDATE    tbOrder
SET              OrderCompounded = 1
WHERE     (Order_DPA_ =
                          (SELECT     Order_DPA_
                            FROM          OrdDetail
                            WHERE      (OrdDetail_DPA_ = @OrdDetail_DPA_)))




--Update the VAT Totals
UPDATE    LevyContract
SET              LevyAmount = VATTotals.LevyVATAmount
FROM         (

	SELECT     Lot.Contract_DPA_, LevyContract.LevyContract_DPA_
	FROM          Lot INNER JOIN
	                      OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN
	                      LevyContract ON Lot.Contract_DPA_ = LevyContract.Contract_DPA_
	WHERE
		(OrdDetail.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
		(cast(floor(cast(Lot.LotTDate AS float)) AS datetime) =
			cast(floor(cast(@TradeDate AS float)) AS datetime)) AND
		(LevyContract.SystemMaintained = 99) AND
		(LevyContract.Deleted <> 1) AND
		(Lot.Deleted <> 1)

)
                      VATEntries INNER JOIN
                      LevyContract ON VATEntries.LevyContract_DPA_ = LevyContract.LevyContract_DPA_ INNER JOIN
                          (

	SELECT     SUM(LevyContract.LevyVATAmount) AS LevyVATAmount, LevyContract.Contract_DPA_
	FROM         LevyContract INNER JOIN
	                      Lot ON LevyContract.Contract_DPA_ = Lot.Contract_DPA_
	WHERE
		(Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
		(CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
			= cast(floor(cast(@TradeDate AS float)) AS datetime)) AND
		(LevyContract.SystemMaintained = 11) AND
		(LevyContract.Deleted <> 1)
		OR
	        (Lot.OrdDetail_DPA_ = @OrdDetail_DPA_) AND
		(CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
			= cast(floor(cast(@TradeDate AS float)) AS datetime)) AND
		(LevyContract.SystemMaintained = 100) AND
		(LevyContract.Deleted <> 1)
	GROUP BY LevyContract.Contract_DPA_



			) VATTotals ON VATEntries.Contract_DPA_ = VATTotals.Contract_DPA_

GO


CREATE PROCEDURE [dbo].[ContractLeviesCrossTab]  AS
/*
DECLARE @SQLStatement nvarchar(4000)

EXEC sp_pivot 'MAX', 'LevyAmount', 'LevyContracts', 'Contract_DPA_', 'LevyShortName', @outSQL = @SQLStatement OUTPUT

IF @SQLStatement = NULL
	BEGIN
		 SELECT  @SQLStatement AS RESULT_FLD WHERE 0 = 1
	END

ELSE

	 BEGIN

		--Add LevyVAtAmount
		--set @SQLStatement = replace(@SQLStatement,'LevyAmount',' LevyAmount + LevyVATAmount ')
		EXEC ('SELECT   *  FROM (SELECT     CONVERT(SMALLDATETIME, CAST(dbo.Lot.LotTDate AS CHAR(12))) AS LotTDate,dbo.Lot.Contract_DPA_ as Contract, dbo.Lot.ContractNumber AS ContractNumber, dbo.OrdDetailList.OrderTypeSale, dbo.Security.SecurityCode,OrdDetailList.OrderSecType_DPA_, dbo.Broker.BrokerCode,
		                      dbo.Lot.LotSlipNo, dbo.Lot.LotPrice, dbo.Lot.LotQty, dbo.OrdDetailList.OrdDetailClient AS ClientName, dbo.OrdDetailList.Client_DPA_ AS Client_DPA_, dbo.Lot.Contract_DPA_, dbo.Lot.LotGrossAmount
				FROM  dbo.OrdDetailList INNER JOIN
		                      dbo.Lot ON dbo.OrdDetailList.OrdDetail_DPA_ = dbo.Lot.OrdDetail_DPA_ INNER JOIN
		                      dbo.Broker ON dbo.Lot.Broker_DPA_ = dbo.Broker.Broker_DPA_ LEFT OUTER JOIN
		                      dbo.Security ON dbo.OrdDetailList.Security_DPA_ = dbo.Security.Security_DPA_ where lot.deleted<>1) AS INNERTBL
			INNER JOIN (' + @SQLStatement + ') CROSSTABTBL ON CONVERT(NVARCHAR(1000), INNERTBL.Contract_DPA_) = CONVERT(NVARCHAR(1000), CROSSTABTBL.Contract_DPA_  )')

	END


DECLARE @SQLStatement nvarchar(4000)

EXEC sp_pivot 'MAX', 'LevyAmount', 'LevyContracts', 'Contract_DPA_', 'LevyShortName', @outSQL = @SQLStatement OUTPUT

IF @SQLStatement = NULL
	BEGIN
		 SELECT  @SQLStatement AS RESULT_FLD WHERE 0 = 1
	END

ELSE

	 BEGIN

		--Add LevyVAtAmount
		set @SQLStatement = 'SELECT     CASE WHEN (GROUPING([Contract_DPA_]) = 1) THEN ''Total'' ELSE CAST(+ [Contract_DPA_] AS NVARCHAR(255)) END AS [Contract_DPA_],
                   MAX(CASE CAST([LevyShortName] AS VARCHAR(255)) WHEN ''Agent'' THEN LevyAmount + LevyVATAmount ELSE 0 END) AS [Agent],
                   MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''Basic'' THEN LevyAmount + LevyVATAmount ELSE 0 END) AS [Basic],
                   MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''Commission'' THEN LevyAmount + LevyVATAmount ELSE 0 END)
                   AS [Commission], MAX(CASE CAST([LevyShortName] AS nVARCHAR(255))
                   WHEN ''MSEComm'' THEN LevyAmount  ELSE 0 END) AS [MSEComm],
                   MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''VAT'' THEN LevyAmount + LevyVATAmount ELSE 0 END) AS [VAT],
                   MAX(LevyAmount + LevyVATAmount) AS Total
FROM          LevyContracts
GROUP BY Contract_DPA_ WITH CUBE'
		EXEC ('SELECT   *  FROM (SELECT     CONVERT(SMALLDATETIME, CAST(dbo.Lot.LotTDate AS CHAR(12))) AS LotTDate,dbo.Lot.Contract_DPA_ as Contract, dbo.Lot.ContractNumber AS ContractNumber, dbo.OrdDetailList.OrderTypeSale, dbo.Security.SecurityCode,OrdDetailList.OrderSecType_DPA_, dbo.Broker.BrokerCode,
		                      dbo.Lot.LotSlipNo, dbo.Lot.LotPrice, dbo.Lot.LotQty, dbo.OrdDetailList.OrdDetailClient AS ClientName, dbo.OrdDetailList.Client_DPA_ AS Client_DPA_, dbo.Lot.Contract_DPA_, dbo.Lot.LotGrossAmount
				FROM  dbo.OrdDetailList INNER JOIN
		                      dbo.Lot ON dbo.OrdDetailList.OrdDetail_DPA_ = dbo.Lot.OrdDetail_DPA_ INNER JOIN
		                      dbo.Broker ON dbo.Lot.Broker_DPA_ = dbo.Broker.Broker_DPA_ LEFT OUTER JOIN
		                      dbo.Security ON dbo.OrdDetailList.Security_DPA_ = dbo.Security.Security_DPA_ where lot.deleted<>1) AS INNERTBL
			INNER JOIN (' + @SQLStatement + ') CROSSTABTBL ON CONVERT(NVARCHAR(1000), INNERTBL.Contract_DPA_) = CONVERT(NVARCHAR(1000), CROSSTABTBL.Contract_DPA_  )')

	END
*/


DECLARE @SQLStatement nvarchar(4000)

EXEC sp_pivot 'MAX', 'LevyAmount', 'LevyContracts', 'Contract_DPA_', 'LevyShortName', @outSQL = @SQLStatement OUTPUT

IF @SQLStatement = NULL
	BEGIN
		 SELECT  @SQLStatement AS RESULT_FLD WHERE 0 = 1
	END

ELSE

	 BEGIN

		--Add LevyVAtAmount
		set @SQLStatement = 'SELECT     CASE WHEN (GROUPING([Contract_DPA_]) = 1) THEN ''Total'' ELSE CAST(+ [Contract_DPA_] AS NVARCHAR(255)) END AS [Contract_DPA_],
                                                   MAX(CASE CAST([LevyShortName] AS VARCHAR(255)) WHEN ''Agent'' THEN LevyAmount ELSE 0 END) AS [Agent],
                                                   MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''Basic'' THEN LevyAmount ELSE 0 END) AS [Basic],
                                                   MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''Commission'' THEN LevyAmount ELSE 0 END)
                                                   - MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''MSEComm'' THEN LevyAmount ELSE 0 END) AS [Commission],
                                                   MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''MSEComm'' THEN LevyAmount ELSE 0 END) AS [MSEComm],
                                                   MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''VAT'' THEN LevyAmount ELSE 0 END) AS [VAT],
                                                   MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''CGT'' THEN LevyAmount ELSE 0 END) AS [CGT],

 						   MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''Basic'' THEN LevyAmount ELSE 0 END) +
                                                   (MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''Commission'' THEN LevyAmount ELSE 0 END)
                                                   - MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''MSEComm'' THEN LevyAmount ELSE 0 END)) +
                                                   MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''MSEComm'' THEN LevyAmount ELSE 0 END) +
                                                   MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''VAT'' THEN LevyAmount ELSE 0 END) +
                                                   MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''CGT'' THEN LevyAmount ELSE 0 END) AS Total

                            FROM          LevyContracts
                            GROUP BY Contract_DPA_ WITH CUBE'
		exec( 'SELECT   *  FROM (SELECT     CONVERT(SMALLDATETIME, CAST(dbo.Lot.LotTDate AS CHAR(12))) AS LotTDate,dbo.Lot.Contract_DPA_ as Contract, dbo.Lot.ContractNumber AS ContractNumber, dbo.OrdDetailList.OrderTypeSale, dbo.Security.SecurityCode,OrdDetailList.OrderSecType_DPA_, dbo.Broker.BrokerCode,
		                      dbo.Lot.LotSlipNo, dbo.Lot.LotPrice, dbo.Lot.LotQty, dbo.OrdDetailList.OrdDetailClient AS ClientName, dbo.OrdDetailList.Client_DPA_ AS Client_DPA_, dbo.Lot.Contract_DPA_, dbo.Lot.LotGrossAmount
				FROM  dbo.OrdDetailList INNER JOIN
		                      dbo.Lot ON dbo.OrdDetailList.OrdDetail_DPA_ = dbo.Lot.OrdDetail_DPA_ INNER JOIN
		                      dbo.Broker ON dbo.Lot.Broker_DPA_ = dbo.Broker.Broker_DPA_ LEFT OUTER JOIN
		                      dbo.Security ON dbo.OrdDetailList.Security_DPA_ = dbo.Security.Security_DPA_ where lot.deleted<>1) AS INNERTBL
			INNER JOIN (' + @SQLStatement + ') CROSSTABTBL ON CONVERT(NVARCHAR(1000), INNERTBL.Contract_DPA_) = CONVERT(NVARCHAR(1000), CROSSTABTBL.Contract_DPA_  )' )

	END

GO






CREATE PROCEDURE [ContractLeviesCrossTabCommission]
@FromDate varchar(20),
@ToDate varchar(20)


 AS


DECLARE @SQLStatement nvarchar(4000)

EXEC sp_pivot 'MAX', 'LevyAmount', 'LevyContracts', 'Contract_DPA_', 'LevyShortName', @outSQL = @SQLStatement OUTPUT

IF @SQLStatement = NULL
	BEGIN
		 SELECT  @SQLStatement AS RESULT_FLD WHERE 0 = 1
	END

ELSE

	 BEGIN


		exec ('SELECT     *
FROM  ( SELECT     OrdDetailList.Agent_DPA_, Contract.ContractSettlementDate, CONVERT(SMALLDATETIME, CAST(Lot.LotTDate AS CHAR(12))) AS LotTDate,
                      Lot.Contract_DPA_ AS Contract, Lot.ContractNumber AS ContractNumber, Security.SecurityCode, Lot.LotSlipNo, Lot.LotPrice, Lot.LotQty,
                      Lot.Contract_DPA_, Lot.LotGrossAmount
FROM         OrdDetailList INNER JOIN
                      Lot ON OrdDetailList.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN
                      Broker ON Lot.Broker_DPA_ = Broker.Broker_DPA_ INNER JOIN
                      Contract ON Lot.Contract_DPA_ = Contract.Contract_DPA_ LEFT OUTER JOIN
                      Security ON OrdDetailList.Security_DPA_ = Security.Security_DPA_
WHERE     (Lot.Deleted <> 1) AND (OrdDetailList.Agent_DPA_ IS NOT NULL) AND (OrdDetailList.Agent_DPA_ IS NOT NULL) AND
                      (CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime) BETWEEN ''' + @FromDate +  ''' AND ''' +  @ToDate + ''') ) AS INNERTBL
			INNER JOIN (' + @SQLStatement + ') CROSSTABTBL ON CONVERT(NVARCHAR(1000), INNERTBL.Contract_DPA_) = CONVERT(NVARCHAR(1000), CROSSTABTBL.Contract_DPA_  )')

	END


GO




CREATE PROCEDURE [ContractLeviesCrossTabForCustodians]
@FromDate varchar(20),
@ToDate varchar(20)


 AS


DECLARE @SQLStatement nvarchar(4000)

EXEC sp_pivot 'MAX', 'LevyAmount', 'LevyContracts', 'Contract_DPA_', 'LevyShortName', @outSQL = @SQLStatement OUTPUT

IF @SQLStatement = NULL
	BEGIN
		 SELECT  @SQLStatement AS RESULT_FLD WHERE 0 = 1
	END

ELSE

	 BEGIN


		exec ('SELECT     *
FROM  ( SELECT     OrdDetailList.Agent_DPA_, Contract.ContractSettlementDate, CONVERT(SMALLDATETIME, CAST(Lot.LotTDate AS CHAR(12))) AS LotTDate,
                      Lot.Contract_DPA_ AS Contract, Lot.ContractNumber AS ContractNumber, Security.SecurityCode, Lot.LotSlipNo, Lot.LotPrice, Lot.LotQty,
                      Lot.Contract_DPA_, Lot.LotGrossAmount
FROM         OrdDetailList INNER JOIN
                      Lot ON OrdDetailList.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN
                      Broker ON Lot.Broker_DPA_ = Broker.Broker_DPA_ INNER JOIN
                      Contract ON Lot.Contract_DPA_ = Contract.Contract_DPA_ LEFT OUTER JOIN
                      Security ON OrdDetailList.Security_DPA_ = Security.Security_DPA_
WHERE     (Lot.Deleted <> 1) AND (OrdDetailList.Agent_DPA_ IS NOT NULL) AND (OrdDetailList.Agent_DPA_ IS NOT NULL) AND
                      (CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime) BETWEEN ''' + @FromDate +  ''' AND ''' +  @ToDate + ''') AND (OrdDetailList.IsCustodian = 1)) AS INNERTBL
			INNER JOIN (' + @SQLStatement + ') CROSSTABTBL ON CONVERT(NVARCHAR(1000), INNERTBL.Contract_DPA_) = CONVERT(NVARCHAR(1000), CROSSTABTBL.Contract_DPA_  )')

	END

GO





CREATE PROCEDURE [ContractLeviesCrossTabForNonCustodians]
@FromDate varchar(20),
@ToDate varchar(20)


 AS


DECLARE @SQLStatement nvarchar(4000)

EXEC sp_pivot 'MAX', 'LevyAmount', 'LevyContracts', 'Contract_DPA_', 'LevyShortName', @outSQL = @SQLStatement OUTPUT

IF @SQLStatement = NULL
	BEGIN
		 SELECT  @SQLStatement AS RESULT_FLD WHERE 0 = 1
	END

ELSE

	 BEGIN


		exec ('SELECT     *
FROM  ( SELECT     OrdDetailList.Agent_DPA_, Contract.ContractSettlementDate, CONVERT(SMALLDATETIME, CAST(Lot.LotTDate AS CHAR(12))) AS LotTDate,
                      Lot.Contract_DPA_ AS Contract, Lot.ContractNumber AS ContractNumber, Security.SecurityCode, Lot.LotSlipNo, Lot.LotPrice, Lot.LotQty,
                      Lot.Contract_DPA_, Lot.LotGrossAmount
FROM         OrdDetailList INNER JOIN
                      Lot ON OrdDetailList.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN
                      Broker ON Lot.Broker_DPA_ = Broker.Broker_DPA_ INNER JOIN
                      Contract ON Lot.Contract_DPA_ = Contract.Contract_DPA_ LEFT OUTER JOIN
                      Security ON OrdDetailList.Security_DPA_ = Security.Security_DPA_
WHERE     (Lot.Deleted <> 1) AND (OrdDetailList.Agent_DPA_ IS NOT NULL) AND (OrdDetailList.Agent_DPA_ IS NOT NULL) AND
                      (CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime) BETWEEN ''' + @FromDate +  ''' AND ''' +  @ToDate + ''') AND (OrdDetailList.IsCustodian = 0)) AS INNERTBL
			INNER JOIN (' + @SQLStatement + ') CROSSTABTBL ON CONVERT(NVARCHAR(1000), INNERTBL.Contract_DPA_) = CONVERT(NVARCHAR(1000), CROSSTABTBL.Contract_DPA_  )')

	END


GO


CREATE PROCEDURE [ContractLeviesCrossTabNew]

@Traded varchar(20) = '16-Jul-2008'

AS

DECLARE @SQLStatement nvarchar(4000)

EXEC sp_pivot 'MAX', 'LevyAmount', 'LevyContracts', 'Contract_DPA_', 'LevyShortName', @outSQL = @SQLStatement OUTPUT

IF @SQLStatement = NULL
	BEGIN
		 SELECT  @SQLStatement AS RESULT_FLD WHERE 0 = 1
	END

ELSE

	 BEGIN

		set @SQLStatement = 'SELECT   *  FROM (SELECT     CONVERT(SMALLDATETIME, CAST(Lot.LotTDate AS CHAR(12))) AS LotTDate, Lot.Contract_DPA_ AS Contract, Lot.ContractNumber AS ContractNumber,
			                      OrdDetailList.OrderTypeSale, Security.SecurityCode, OrdDetailList.OrderSecType_DPA_, Broker.BrokerCode, Lot.LotSlipNo, Lot.LotPrice, Lot.LotQty,
			                      OrdDetailList.OrdDetailClient AS ClientName, OrdDetailList.Client_DPA_ AS Client_DPA_, Lot.Contract_DPA_, Lot.LotGrossAmount,
			                      CAST(FLOOR(CAST(Contract.ContractSettlementDate AS float)) AS datetime) AS ContractSettlementDate
			FROM         OrdDetailList INNER JOIN
			                      Lot ON OrdDetailList.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN
			                      Broker ON Lot.Broker_DPA_ = Broker.Broker_DPA_ INNER JOIN
			                      Contract ON Lot.Contract_DPA_ = Contract.Contract_DPA_ LEFT OUTER JOIN
			                      Security ON OrdDetailList.Security_DPA_ = Security.Security_DPA_
			WHERE     (Lot.Deleted <> 1) AND (CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime) = CONVERT(DATETIME, ''' + @Traded + '''))) AS INNERTBL
			INNER JOIN (' + @SQLStatement + ') CROSSTABTBL ON CONVERT(NVARCHAR(1000), INNERTBL.Contract_DPA_) = CONVERT(NVARCHAR(1000), CROSSTABTBL.Contract_DPA_  )'


		EXEC (@SQLStatement)

	END

GO


create PROCEDURE [ContractLeviesCrossTabNewSettle]

@Traded varchar(20) = '16-Jul-2008'

AS

DECLARE @SQLStatement nvarchar(4000)

EXEC sp_pivot 'MAX', 'LevyAmount', 'LevyContracts', 'Contract_DPA_', 'LevyShortName', @outSQL = @SQLStatement OUTPUT

IF @SQLStatement = NULL
	BEGIN
		 SELECT  @SQLStatement AS RESULT_FLD WHERE 0 = 1
	END

ELSE

	 BEGIN

		set @SQLStatement = 'SELECT   *  FROM (SELECT     CONVERT(SMALLDATETIME, CAST(Lot.LotTDate AS CHAR(12))) AS LotTDate, Lot.Contract_DPA_ AS Contract, Lot.ContractNumber AS ContractNumber,
			                      OrdDetailList.OrderTypeSale, Security.SecurityCode, OrdDetailList.OrderSecType_DPA_, Broker.BrokerCode, Lot.LotSlipNo, Lot.LotPrice, Lot.LotQty,
			                      OrdDetailList.OrdDetailClient AS ClientName, OrdDetailList.Client_DPA_ AS Client_DPA_, Lot.Contract_DPA_, Lot.LotGrossAmount,
			                      CAST(FLOOR(CAST(Contract.ContractSettlementDate AS float)) AS datetime) AS ContractSettlementDate
			FROM         OrdDetailList INNER JOIN
			                      Lot ON OrdDetailList.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN
			                      Broker ON Lot.Broker_DPA_ = Broker.Broker_DPA_ INNER JOIN
			                      Contract ON Lot.Contract_DPA_ = Contract.Contract_DPA_ LEFT OUTER JOIN
			                      Security ON OrdDetailList.Security_DPA_ = Security.Security_DPA_
			WHERE     (Lot.Deleted <> 1) AND (CAST(FLOOR(CAST(Contract.ContractSettlementDate AS float)) AS datetime) = CONVERT(DATETIME, ''' + @Traded + '''))) AS INNERTBL
			INNER JOIN (' + @SQLStatement + ') CROSSTABTBL ON CONVERT(NVARCHAR(1000), INNERTBL.Contract_DPA_) = CONVERT(NVARCHAR(1000), CROSSTABTBL.Contract_DPA_  )'


		EXEC (@SQLStatement)

	END

GO



create PROCEDURE [ContractLeviesCrossTabNewSettleCSD]

@Traded varchar(20) = '16-Jul-2008'

AS

DECLARE @SQLStatement nvarchar(4000)

EXEC sp_pivot 'MAX', 'LevyAmount', 'LevyContracts', 'Contract_DPA_', 'LevyShortName', @outSQL = @SQLStatement OUTPUT

IF @SQLStatement = NULL
	BEGIN
		 SELECT  @SQLStatement AS RESULT_FLD WHERE 0 = 1
	END

ELSE

	 BEGIN

		set @SQLStatement = 'SELECT   *  FROM (SELECT     CONVERT(SMALLDATETIME, CAST(Lot.LotTDate AS CHAR(12))) AS LotTDate, Lot.Contract_DPA_ AS Contract, Lot.ContractNumber AS ContractNumber,
			                      OrdDetailList.OrderTypeSale, Security.SecurityCode, OrdDetailList.OrderSecType_DPA_, Broker.BrokerCode, Lot.LotSlipNo, Lot.LotPrice, Lot.LotQty,
			                      OrdDetailList.OrdDetailClient AS ClientName, OrdDetailList.Client_DPA_ AS Client_DPA_, Lot.Contract_DPA_, Lot.LotGrossAmount,
			                      CAST(FLOOR(CAST(Contract.ContractSettlementDate AS float)) AS datetime) AS ContractSettlementDate
			FROM         OrdDetailList INNER JOIN
			                      Lot ON OrdDetailList.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN
			                      Broker ON Lot.Broker_DPA_ = Broker.Broker_DPA_ INNER JOIN
			                      Contract ON Lot.Contract_DPA_ = Contract.Contract_DPA_ LEFT OUTER JOIN
			                      Security ON OrdDetailList.Security_DPA_ = Security.Security_DPA_
			WHERE     (Security.IsOnCSD = 1) and (Lot.Deleted <> 1) AND (CAST(FLOOR(CAST(Contract.ContractSettlementDate AS float)) AS datetime) = CONVERT(DATETIME, ''' + @Traded + '''))) AS INNERTBL
			INNER JOIN (' + @SQLStatement + ') CROSSTABTBL ON CONVERT(NVARCHAR(1000), INNERTBL.Contract_DPA_) = CONVERT(NVARCHAR(1000), CROSSTABTBL.Contract_DPA_  )'


		EXEC (@SQLStatement)

	END


GO





CREATE PROCEDURE [ContractLeviesCrossTabNewSettleNonCSD]

@Traded varchar(20) , @ToTrade varchar(20)

AS

DECLARE @SQLStatement nvarchar(4000)

EXEC sp_pivot 'MAX', 'LevyAmount', 'LevyContracts', 'Contract_DPA_', 'LevyShortName', @outSQL = @SQLStatement OUTPUT

IF @SQLStatement = NULL
	BEGIN
		 SELECT  @SQLStatement AS RESULT_FLD WHERE 0 = 1
	END

ELSE

	 BEGIN
/*
		set @SQLStatement = 'SELECT   *  FROM (SELECT     CONVERT(SMALLDATETIME, CAST(Lot.LotTDate AS CHAR(12))) AS LotTDate, Lot.Contract_DPA_ AS Contract, Lot.ContractNumber AS ContractNumber,
			                      OrdDetailList.OrderTypeSale, Security.SecurityCode, OrdDetailList.OrderSecType_DPA_, Broker.BrokerCode, Lot.LotSlipNo, Lot.LotPrice, Lot.LotQty,
			                      OrdDetailList.OrdDetailClient AS ClientName, OrdDetailList.Client_DPA_ AS Client_DPA_, Lot.Contract_DPA_, Lot.LotGrossAmount,
			                      CAST(FLOOR(CAST(Contract.ContractSettlementDate AS float)) AS datetime) AS ContractSettlementDate
			FROM         OrdDetailList INNER JOIN
			                      Lot ON OrdDetailList.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN
			                      Broker ON Lot.Broker_DPA_ = Broker.Broker_DPA_ INNER JOIN
			                      Contract ON Lot.Contract_DPA_ = Contract.Contract_DPA_ LEFT OUTER JOIN
			                      Security ON OrdDetailList.Security_DPA_ = Security.Security_DPA_
			WHERE     (Security.IsOnCSD <> 1) and (Lot.Deleted <> 1) AND (CAST(FLOOR(CAST(Contract.ContractSettlementDate AS float)) AS datetime) = CONVERT(DATETIME, ''' + @Traded + '''))) AS INNERTBL
			INNER JOIN (' + @SQLStatement + ') CROSSTABTBL ON CONVERT(NVARCHAR(1000), INNERTBL.Contract_DPA_) = CONVERT(NVARCHAR(1000), CROSSTABTBL.Contract_DPA_  )'

*/

		/*set @SQLStatement = 'SELECT     CASE WHEN (GROUPING([Contract_DPA_]) = 1) THEN ''Total'' ELSE CAST(+ [Contract_DPA_] AS NVARCHAR(255))
		   END AS [Contract_DPA_], MAX(CASE CAST([LevyShortName] AS VARCHAR(255))
		   WHEN ''Agent'' THEN LevyAmount + LevyVATAmount ELSE 0 END) AS [Agent], MAX(CASE CAST([LevyShortName] AS nVARCHAR(255))
		   WHEN ''Basic'' THEN LevyAmount + LevyVATAmount ELSE 0 END) AS [Basic], MAX(CASE CAST([LevyShortName] AS nVARCHAR(255))
		   WHEN ''Commission'' THEN LevyAmount + LevyVATAmount ELSE 0 END) AS [Commission],
		   MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''MSEComm'' THEN LevyAmount ELSE 0 END) AS [MSEComm],
		   MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''VAT'' THEN LevyAmount + LevyVATAmount ELSE 0 END) AS [VAT],
		   MAX(LevyAmount + LevyVATAmount) AS Total
		FROM          LevyContracts
		GROUP BY Contract_DPA_ WITH CUBE'	*/
		set @SQLStatement = 'SELECT     CASE WHEN (GROUPING([Contract_DPA_]) = 1) THEN ''Total'' ELSE CAST(+ [Contract_DPA_] AS NVARCHAR(255)) END AS [Contract_DPA_],
                                                   MAX(CASE CAST([LevyShortName] AS VARCHAR(255)) WHEN ''Agent'' THEN LevyAmount ELSE 0 END) AS [Agent],
                                                   MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''Basic'' THEN LevyAmount ELSE 0 END) AS [Basic],
                                                   MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''Commission'' THEN LevyAmount ELSE 0 END)
                                                   - MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''MSEComm'' THEN LevyAmount ELSE 0 END) AS [Commission],
                                                   MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''MSEComm'' THEN LevyAmount ELSE 0 END) AS [MSEComm],
                                                   MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''VAT'' THEN LevyAmount ELSE 0 END) AS [VAT],

 						   MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''Basic'' THEN LevyAmount ELSE 0 END) +
                                                   (MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''Commission'' THEN LevyAmount ELSE 0 END)
                                                   - MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''MSEComm'' THEN LevyAmount ELSE 0 END)) +
                                                   MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''MSEComm'' THEN LevyAmount ELSE 0 END) +
                                                   MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN ''VAT'' THEN LevyAmount ELSE 0 END) AS Total

                            FROM          LevyContracts
                            GROUP BY Contract_DPA_ WITH CUBE'

		set @SQLStatement = 'SELECT   *  FROM (SELECT     CONVERT(SMALLDATETIME, CAST(Lot.LotTDate AS CHAR(12))) AS LotTDate, Lot.Contract_DPA_ AS Contract, Lot.ContractNumber AS ContractNumber,
			                      OrdDetailList.OrderTypeSale, Security.SecurityCode, OrdDetailList.OrderSecType_DPA_, Broker.BrokerCode, Lot.LotSlipNo, Lot.LotPrice, Lot.LotQty,
			                      OrdDetailList.OrdDetailClient AS ClientName, OrdDetailList.Client_DPA_ AS Client_DPA_, Lot.Contract_DPA_, Lot.LotGrossAmount,
			                      CAST(FLOOR(CAST(Contract.ContractSettlementDate AS float)) AS datetime) AS ContractSettlementDate
			FROM         OrdDetailList INNER JOIN
			                      Lot ON OrdDetailList.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN
			                      Broker ON Lot.Broker_DPA_ = Broker.Broker_DPA_ INNER JOIN
			                      Contract ON Lot.Contract_DPA_ = Contract.Contract_DPA_ LEFT OUTER JOIN
			                      Security ON OrdDetailList.Security_DPA_ = Security.Security_DPA_
			WHERE    (Lot.Deleted <> 1) AND (CAST(FLOOR(CAST(Contract.ContractSettlementDate AS float)) AS datetime) between  CONVERT(DATETIME, ''' + @Traded + ''' )  and CONVERT(DATETIME, ''' + @ToTrade + ''' ) )) AS INNERTBL
			INNER JOIN ('+ @SQLStatement+') CROSSTABTBL ON CONVERT(NVARCHAR(1000), INNERTBL.Contract_DPA_) = CONVERT(NVARCHAR(1000), CROSSTABTBL.Contract_DPA_  ) ORDER BY INNERTBL.LotTDate'



		EXEC (@SQLStatement)

	END
GO

CREATE PROCEDURE ContractLeviesCrossTabToDate
@ToDate nvarchar(100)
AS
DECLARE @SQLStatement nvarchar(4000)
EXEC sp_pivot 'MAX', 'LevyAmount', 'LevyContracts', 'Contract_DPA_', 'LevyShortName', @outSQL = @SQLStatement OUTPUT
IF @SQLStatement = NULL
	BEGIN
		 SELECT  @SQLStatement AS RESULT_FLD WHERE 0 = 1
	END

ELSE
	 BEGIN
		EXEC ('SELECT   *  FROM (SELECT     dbo.OrdDetailList.AgentName, dbo.OrdDetailList.IsCustodian, dbo.Contract.ContractSettlementDate,
				CONVERT(SMALLDATETIME, CAST(dbo.Lot.LotTDate AS CHAR(12))) AS LotTDate,dbo.Lot.Contract_DPA_ as Contract, dbo.Lot.ContractNumber AS ContractNumber, dbo.OrdDetailList.OrderTypeSale, dbo.Security.SecurityCode,OrdDetailList.OrderSecType_DPA_, dbo.Broker.BrokerCode,
		                      dbo.Lot.LotSlipNo, dbo.Lot.LotPrice, dbo.Lot.LotQty, dbo.OrdDetailList.OrdDetailClient AS ClientName, dbo.OrdDetailList.Client_DPA_ AS Client_DPA_, dbo.Lot.Contract_DPA_, dbo.Lot.LotGrossAmount
				FROM  dbo.OrdDetailList INNER JOIN
		                      dbo.Lot ON dbo.OrdDetailList.OrdDetail_DPA_ = dbo.Lot.OrdDetail_DPA_ INNER JOIN
		                      dbo.Broker ON dbo.Lot.Broker_DPA_ = dbo.Broker.Broker_DPA_ INNER JOIN
                      Contract ON Lot.Contract_DPA_ = Contract.Contract_DPA_  LEFT OUTER JOIN
		                      dbo.Security ON dbo.OrdDetailList.Security_DPA_ = dbo.Security.Security_DPA_ where lot.deleted<>1
			and cast(floor(cast(Lot.lottdate as float)) as datetime) = cast(floor(cast(convert(datetime,''' + @ToDate + ''') as float)) as datetime)
			) AS INNERTBL
			INNER JOIN (' + @SQLStatement + ') CROSSTABTBL ON CONVERT(NVARCHAR(1000), INNERTBL.Contract_DPA_) = CONVERT(NVARCHAR(1000), CROSSTABTBL.Contract_DPA_  )

')
	END


GO

CREATE PROCEDURE [dbo].[CurrentAssetsProc] @StartDate DateTime,@EndDate DateTime AS

SET NOCOUNT ON
SELECT  PnL.* INTO #tableA FROM (
--General entities
SELECT
	dbo.FullEntityTypeList.AccountType_DPA_ AS AccountType_DPA_,
	dbo.FullEntityTypeList.EntityTypeAccountType AS AccountType,
	'******' + CONVERT(VARCHAR(500),dbo.FullEntityTypeList.EntityType_DPA_) AS [Account Code],
	dbo.FullEntityTypeList.EntityTypeName AS AccountName,
	CASE
		WHEN SUM(dbo.EntityTransactionList.Balance) < 0 THEN 0 - SUM(dbo.EntityTransactionList.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.EntityTransactionList.Balance) >= 0 THEN SUM(dbo.EntityTransactionList.Balance)
		ELSE 0 END AS Credit,Cast(Floor(Cast(Min(EntityTransactionList.TransDate) as Float)) as DateTime) as TransDate
FROM         dbo.Entity INNER JOIN
                      dbo.FullEntityTypeList ON dbo.Entity.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_ INNER JOIN
                      dbo.EntityTransactionList ON dbo.Entity.Entity_DPA_ = dbo.EntityTransactionList.Entity_DPA_
WHERE 		dbo.FullEntityTypeList.AccountType_DPA_ IN (4,7)
GROUP BY dbo.FullEntityTypeList.AccountType_DPA_,dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName, dbo.FullEntityTypeList.EntityType_DPA_
) PnL


--Clients
INSERT INTO #tableA
SELECT     dbo.FullEntityTypeList.AccountType_DPA_ AS AccountType_DPA_,
		dbo.FullEntityTypeList.EntityTypeAccountType AS AccountType,
		'******1' AS [Account Code],
		dbo.FullEntityTypeList.EntityTypeName AS AccountName,CASE WHEN SUM(StatementList.Balance) < 0
THEN 0 - SUM(StatementList.Balance) ELSE 0 END AS Debit, CASE WHEN SUM(StatementList.Balance) >= 0
THEN SUM(StatementList.Balance) ELSE 0 END AS Credit,TransDate From (Select StatementList.*,1 as EntityType_DPA_ from StatementList) StatementList INNER JOIN
		dbo.FullEntityTypeList ON StatementList.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_
WHERE 		dbo.FullEntityTypeList.AccountType_DPA_ IN (4,7)
GROUP BY dbo.FullEntityTypeList.AccountType_DPA_,dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName,TransDate

--Agents
INSERT INTO #tableA
SELECT     dbo.FullEntityTypeList.AccountType_DPA_ AS AccountType_DPA_,
		dbo.FullEntityTypeList.EntityTypeAccountType AS AccountType,
		'******1' AS [Account Code],
		dbo.FullEntityTypeList.EntityTypeName AS AccountName,CASE WHEN SUM(AgentStatement.Balance) < 0
THEN 0 - SUM(AgentStatement.Balance) ELSE 0 END AS Debit, CASE WHEN SUM(AgentStatement.Balance) >= 0
THEN SUM(AgentStatement.Balance) ELSE 0 END AS Credit,TransDate From (Select AgentStatement.*,2 as EntityType_DPA_ from AgentStatement) AgentStatement INNER JOIN
		dbo.FullEntityTypeList ON AgentStatement.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_
WHERE 		dbo.FullEntityTypeList.AccountType_DPA_ IN (4,7)
GROUP BY dbo.FullEntityTypeList.AccountType_DPA_,dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName,TransDate

--Brokers
INSERT INTO #tableA
SELECT     dbo.FullEntityTypeList.AccountType_DPA_ AS AccountType_DPA_,
		dbo.FullEntityTypeList.EntityTypeAccountType AS AccountType,
		'******1' AS [Account Code],
		dbo.FullEntityTypeList.EntityTypeName AS AccountName,CASE WHEN SUM(BrokerStatement.Credit-BrokerStatement.Debit) < 0
THEN 0 - SUM(BrokerStatement.Credit-BrokerStatement.Debit) ELSE 0 END AS Debit, CASE WHEN SUM(BrokerStatement.Credit-BrokerStatement.Debit) >= 0
THEN SUM(BrokerStatement.Credit-BrokerStatement.Debit) ELSE 0 END AS Credit,TransDate From (Select BrokerStatement.*,3 as EntityType_DPA_ from BrokerStatement) BrokerStatement INNER JOIN
		dbo.FullEntityTypeList ON BrokerStatement.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_
WHERE 		dbo.FullEntityTypeList.AccountType_DPA_ IN (4,7)
GROUP BY dbo.FullEntityTypeList.AccountType_DPA_,dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName,TransDate

--Account Managers
INSERT INTO #tableA
SELECT     dbo.FullEntityTypeList.AccountType_DPA_ AS AccountType_DPA_,
		dbo.FullEntityTypeList.EntityTypeAccountType AS AccountType,
		'******1' AS [Account Code],
		dbo.FullEntityTypeList.EntityTypeName AS AccountName,CASE WHEN SUM(OwnerStatement.Credit-OwnerStatement.Debit) < 0
THEN 0 - SUM(OwnerStatement.Credit-OwnerStatement.Debit) ELSE 0 END AS Debit, CASE WHEN SUM(OwnerStatement.Credit-OwnerStatement.Debit) >= 0
THEN SUM(OwnerStatement.Credit-OwnerStatement.Debit) ELSE 0 END AS Credit,TransDate From (Select OwnerStatement.*,7 as EntityType_DPA_ from OwnerStatement) OwnerStatement INNER JOIN
		dbo.FullEntityTypeList ON OwnerStatement.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_
WHERE 		dbo.FullEntityTypeList.AccountType_DPA_ IN (4,7)
GROUP BY dbo.FullEntityTypeList.AccountType_DPA_,dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName,TransDate


--Levies and Broker commission
INSERT INTO #tableA
SELECT
	dbo.EntityList.AccountType_DPA_ AS AccountType_DPA_,
	dbo.EntityList.AccountTypeName AS AccountType,
	'------' + CONVERT(VARCHAR(500),dbo.EntityList.Entity_DPA_) AS [Account Code],
	dbo.EntityList.EntityName AS AccountName,
	CASE
		WHEN SUM(dbo.BrokerCommissionStatement.Balance) < 0 THEN 0 - SUM(dbo.BrokerCommissionStatement.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.BrokerCommissionStatement.Balance) >= 0 THEN SUM(dbo.BrokerCommissionStatement.Balance)
		ELSE 0 END AS Credit,BrokerCommissionStatement.TransDate
FROM         dbo.EntityList INNER JOIN
                      dbo.BrokerCommissionStatement ON dbo.EntityList.Entity_DPA_ = dbo.BrokerCommissionStatement.Entity_DPA_
WHERE 		dbo.EntityList.AccountType_DPA_ IN (4,7)
GROUP BY dbo.EntityList.AccountType_DPA_,dbo.EntityList.AccountTypeName,dbo.EntityList.EntityName, dbo.EntityList.Entity_DPA_,BrokerCommissionStatement.TransDate

INSERT INTO #tableA
SELECT
	dbo.EntityList.AccountType_DPA_ AS AccountType_DPA_,
	dbo.EntityList.AccountTypeName AS AccountType,
	'------' + CONVERT(VARCHAR(500),dbo.EntityList.Entity_DPA_) AS [Account Code],
	dbo.EntityList.EntityName AS AccountName,
	CASE
		WHEN SUM(dbo.CommissionStatement.Balance) < 0 THEN 0 - SUM(dbo.CommissionStatement.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.CommissionStatement.Balance) >= 0 THEN SUM(dbo.CommissionStatement.Balance)
		ELSE 0 END AS Credit,CommissionStatement.TransDate
FROM         dbo.EntityList INNER JOIN
                      dbo.CommissionStatement ON dbo.EntityList.Entity_DPA_ = dbo.CommissionStatement.Entity_DPA_
WHERE 		dbo.EntityList.AccountType_DPA_ IN (4,7)
GROUP BY dbo.EntityList.AccountType_DPA_,dbo.EntityList.AccountTypeName,dbo.EntityList.EntityName, dbo.EntityList.Entity_DPA_,CommissionStatement.TransDate

--Trading Account
INSERT INTO #tableA
SELECT
	dbo.EntityList.AccountType_DPA_ AS AccountType_DPA_,
	dbo.EntityList.AccountTypeName AS AccountType,
	'------' + CONVERT(VARCHAR(500),dbo.EntityList.Entity_DPA_) AS [Account Code],
	dbo.EntityList.EntityName AS AccountName,
	CASE
		WHEN SUM(dbo.CDSControlStatement.Balance) < 0 THEN 0 - SUM(dbo.CDSControlStatement.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.CDSControlStatement.Balance) >= 0 THEN SUM(dbo.CDSControlStatement.Balance)
		ELSE 0 END AS Credit,CDSControlStatement.TransDate
FROM         dbo.EntityList INNER JOIN
                      dbo.CDSControlStatement ON dbo.EntityList.Entity_DPA_ = dbo.CDSControlStatement.Client_DPA_
WHERE 		dbo.EntityList.AccountType_DPA_ IN (4,7)
GROUP BY dbo.EntityList.AccountType_DPA_,dbo.EntityList.AccountTypeName,dbo.EntityList.EntityName, dbo.EntityList.Entity_DPA_,CDSControlStatement.TransDate

--Nominal accounts
INSERT INTO #tableA
SELECT
	dbo.AccountList.AccountTypeLevel1_DPA_ AS AccountType_DPA_,
	dbo.AccountList.AccountTypeLevel1 AS AccountType,
	CONVERT(VARCHAR(500),dbo.AccountList.Account_DPA_) AS [Account Code],
	dbo.AccountList.AccountName,
	CASE
		WHEN SUM(dbo.DB_BankAccountStatement.Balance) < 0 THEN 0 - SUM(dbo.DB_BankAccountStatement.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.DB_BankAccountStatement.Balance) >= 0 THEN SUM(dbo.DB_BankAccountStatement.Balance)
		ELSE 0 END AS Credit,DB_BankAccountStatement.TransDate
FROM         dbo.DB_BankAccountStatement INNER JOIN
                      dbo.AccountList ON dbo.DB_BankAccountStatement.BankAccount_DPA_ = dbo.AccountList.Account_DPA_
WHERE 		dbo.AccountList.AccountTypeLevel1_DPA_ IN (4,7)
GROUP BY dbo.AccountList.AccountTypeLevel1_DPA_,dbo.AccountList.AccountTypeLevel1,dbo.AccountList.AccountName, dbo.AccountList.Account_DPA_,DB_BankAccountStatement.TransDate



SELECT  IDENTITY(int, 1,1)  AS SequenceID, #tableA.AccountType_DPA_,#tableA.AccountType,#tableA.[Account Code],#tableA.AccountName,Sum(#tableA.Credit - #tableA.Debit) AS Balance
	 INTO #tableB FROM #tableA Where #tableA.TransDate between @StartDate and @enddate Group by #tableA.AccountType_DPA_,#tableA.AccountType,#tableA.[Account Code],#tableA.AccountName ORDER BY AccountType_DPA_

DROP TABLE #tableA

select T.SequenceID,T.AccountType_DPA_,T.AccountType,T.[Account Code],T.AccountName,T.Balance,
       case when SequenceID = (select top 1 SequenceID from #tableB
                           where AccountType_DPA_
                              = T.AccountType_DPA_
                          order by SequenceID desc)
            then (select CONVERT(CHAR(20),sum(Balance))
                     from #tableB
                     where SequenceID <= T.SequenceID
                        and AccountType_DPA_
                              = T.AccountType_DPA_)
            else ' ' end as 'SubTotal',
       case when SequenceID = (select top 1 SequenceID from #tableB
                           order by SequenceID desc)
            then (select CONVERT(CHAR(20),sum(Balance))
                      from #tableB)
             else ' ' end as 'GrandTotal'
from #tableB T
  order by SequenceID



DROP TABLE #tableB


GO

CREATE PROCEDURE [dbo].[CurrentLiabilitiesProc] @StartDate DateTime,@EndDate DateTime AS


SET NOCOUNT ON
SELECT  PnL.* INTO #tableA FROM (
--General entities
SELECT
	dbo.FullEntityTypeList.AccountType_DPA_ AS AccountType_DPA_,
	dbo.FullEntityTypeList.EntityTypeAccountType AS AccountType,
	'******' + CONVERT(VARCHAR(500),dbo.FullEntityTypeList.EntityType_DPA_) AS [Account Code],
	dbo.FullEntityTypeList.EntityTypeName AS AccountName,
	CASE
		WHEN SUM(dbo.EntityTransactionList.Balance) < 0 THEN 0 - SUM(dbo.EntityTransactionList.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.EntityTransactionList.Balance) >= 0 THEN SUM(dbo.EntityTransactionList.Balance)
		ELSE 0 END AS Credit,Cast(Floor(Cast(Min(EntityTransactionList.TransDate) as Float)) as DateTime) as TransDate
FROM         dbo.Entity INNER JOIN
                      dbo.FullEntityTypeList ON dbo.Entity.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_ INNER JOIN
                      dbo.EntityTransactionList ON dbo.Entity.Entity_DPA_ = dbo.EntityTransactionList.Entity_DPA_
WHERE 		dbo.FullEntityTypeList.AccountType_DPA_ =5
GROUP BY dbo.FullEntityTypeList.AccountType_DPA_,dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName, dbo.FullEntityTypeList.EntityType_DPA_
) PnL


--Clients
INSERT INTO #tableA
SELECT     dbo.FullEntityTypeList.AccountType_DPA_ AS AccountType_DPA_,
		dbo.FullEntityTypeList.EntityTypeAccountType AS AccountType,
		'******1' AS [Account Code],
		dbo.FullEntityTypeList.EntityTypeName AS AccountName,CASE WHEN SUM(StatementList.Balance) < 0
THEN 0 - SUM(StatementList.Balance) ELSE 0 END AS Debit, CASE WHEN SUM(StatementList.Balance) >= 0
THEN SUM(StatementList.Balance) ELSE 0 END AS Credit,TransDate From (Select StatementList.*,1 as EntityType_DPA_ from StatementList) StatementList INNER JOIN
		dbo.FullEntityTypeList ON StatementList.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_
WHERE 		dbo.FullEntityTypeList.AccountType_DPA_= 5
GROUP BY dbo.FullEntityTypeList.AccountType_DPA_,dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName,TransDate

--Agents
INSERT INTO #tableA
SELECT     dbo.FullEntityTypeList.AccountType_DPA_ AS AccountType_DPA_,
		dbo.FullEntityTypeList.EntityTypeAccountType AS AccountType,
		'******1' AS [Account Code],
		dbo.FullEntityTypeList.EntityTypeName AS AccountName,CASE WHEN SUM(AgentStatement.Balance) < 0
THEN 0 - SUM(AgentStatement.Balance) ELSE 0 END AS Debit, CASE WHEN SUM(AgentStatement.Balance) >= 0
THEN SUM(AgentStatement.Balance) ELSE 0 END AS Credit,TransDate From (Select AgentStatement.*,2 as EntityType_DPA_ from AgentStatement) AgentStatement INNER JOIN
		dbo.FullEntityTypeList ON AgentStatement.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_
WHERE 		dbo.FullEntityTypeList.AccountType_DPA_=5
GROUP BY dbo.FullEntityTypeList.AccountType_DPA_,dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName,TransDate

--Brokers
INSERT INTO #tableA
SELECT     dbo.FullEntityTypeList.AccountType_DPA_ AS AccountType_DPA_,
		dbo.FullEntityTypeList.EntityTypeAccountType AS AccountType,
		'******1' AS [Account Code],
		dbo.FullEntityTypeList.EntityTypeName AS AccountName,CASE WHEN SUM(BrokerStatement.Credit-BrokerStatement.Debit) < 0
THEN 0 - SUM(BrokerStatement.Credit-BrokerStatement.Debit) ELSE 0 END AS Debit, CASE WHEN SUM(BrokerStatement.Credit-BrokerStatement.Debit) >= 0
THEN SUM(BrokerStatement.Credit-BrokerStatement.Debit) ELSE 0 END AS Credit,TransDate From (Select BrokerStatement.*,3 as EntityType_DPA_ from BrokerStatement) BrokerStatement INNER JOIN
		dbo.FullEntityTypeList ON BrokerStatement.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_
WHERE 		dbo.FullEntityTypeList.AccountType_DPA_ =5
GROUP BY dbo.FullEntityTypeList.AccountType_DPA_,dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName,TransDate

--Account Managers
INSERT INTO #tableA
SELECT     dbo.FullEntityTypeList.AccountType_DPA_ AS AccountType_DPA_,
		dbo.FullEntityTypeList.EntityTypeAccountType AS AccountType,
		'******1' AS [Account Code],
		dbo.FullEntityTypeList.EntityTypeName AS AccountName,CASE WHEN SUM(OwnerStatement.Credit-OwnerStatement.Debit) < 0
THEN 0 - SUM(OwnerStatement.Credit-OwnerStatement.Debit) ELSE 0 END AS Debit, CASE WHEN SUM(OwnerStatement.Credit-OwnerStatement.Debit) >= 0
THEN SUM(OwnerStatement.Credit-OwnerStatement.Debit) ELSE 0 END AS Credit,TransDate From (Select OwnerStatement.*,7 as EntityType_DPA_ from OwnerStatement) OwnerStatement INNER JOIN
		dbo.FullEntityTypeList ON OwnerStatement.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_
WHERE 		dbo.FullEntityTypeList.AccountType_DPA_ =5
GROUP BY dbo.FullEntityTypeList.AccountType_DPA_,dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName,TransDate


--Levies and Broker commission
INSERT INTO #tableA
SELECT
	dbo.EntityList.AccountType_DPA_ AS AccountType_DPA_,
	dbo.EntityList.AccountTypeName AS AccountType,
	'------' + CONVERT(VARCHAR(500),dbo.EntityList.Entity_DPA_) AS [Account Code],
	dbo.EntityList.EntityName AS AccountName,
	CASE
		WHEN SUM(dbo.BrokerCommissionStatement.Balance) < 0 THEN 0 - SUM(dbo.BrokerCommissionStatement.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.BrokerCommissionStatement.Balance) >= 0 THEN SUM(dbo.BrokerCommissionStatement.Balance)
		ELSE 0 END AS Credit,BrokerCommissionStatement.TransDate
FROM         dbo.EntityList INNER JOIN
                      dbo.BrokerCommissionStatement ON dbo.EntityList.Entity_DPA_ = dbo.BrokerCommissionStatement.Entity_DPA_
WHERE 		dbo.EntityList.AccountType_DPA_ = 5
GROUP BY dbo.EntityList.AccountType_DPA_,dbo.EntityList.AccountTypeName,dbo.EntityList.EntityName, dbo.EntityList.Entity_DPA_,BrokerCommissionStatement.TransDate

INSERT INTO #tableA
SELECT
	dbo.EntityList.AccountType_DPA_ AS AccountType_DPA_,
	dbo.EntityList.AccountTypeName AS AccountType,
	'------' + CONVERT(VARCHAR(500),dbo.EntityList.Entity_DPA_) AS [Account Code],
	dbo.EntityList.EntityName AS AccountName,
	CASE
		WHEN SUM(dbo.CommissionStatement.Balance) < 0 THEN 0 - SUM(dbo.CommissionStatement.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.CommissionStatement.Balance) >= 0 THEN SUM(dbo.CommissionStatement.Balance)
		ELSE 0 END AS Credit,CommissionStatement.TransDate
FROM         dbo.EntityList INNER JOIN
                      dbo.CommissionStatement ON dbo.EntityList.Entity_DPA_ = dbo.CommissionStatement.Entity_DPA_
WHERE 		dbo.EntityList.AccountType_DPA_ = 5
GROUP BY dbo.EntityList.AccountType_DPA_,dbo.EntityList.AccountTypeName,dbo.EntityList.EntityName, dbo.EntityList.Entity_DPA_,CommissionStatement.TransDate

--Trading Account
INSERT INTO #tableA
SELECT
	dbo.EntityList.AccountType_DPA_ AS AccountType_DPA_,
	dbo.EntityList.AccountTypeName AS AccountType,
	'------' + CONVERT(VARCHAR(500),dbo.EntityList.Entity_DPA_) AS [Account Code],
	dbo.EntityList.EntityName AS AccountName,
	CASE
		WHEN SUM(dbo.CDSControlStatement.Balance) < 0 THEN 0 - SUM(dbo.CDSControlStatement.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.CDSControlStatement.Balance) >= 0 THEN SUM(dbo.CDSControlStatement.Balance)
		ELSE 0 END AS Credit,CDSControlStatement.TransDate
FROM         dbo.EntityList INNER JOIN
                      dbo.CDSControlStatement ON dbo.EntityList.Entity_DPA_ = dbo.CDSControlStatement.Client_DPA_
WHERE 		dbo.EntityList.AccountType_DPA_ = 5
GROUP BY dbo.EntityList.AccountType_DPA_,dbo.EntityList.AccountTypeName,dbo.EntityList.EntityName, dbo.EntityList.Entity_DPA_,CDSControlStatement.TransDate

--Nominal accounts
INSERT INTO #tableA
SELECT
	dbo.AccountList.AccountTypeLevel1_DPA_ AS AccountType_DPA_,
	dbo.AccountList.AccountTypeLevel1 AS AccountType,
	CONVERT(VARCHAR(500),dbo.AccountList.Account_DPA_) AS [Account Code],
	dbo.AccountList.AccountName,
	CASE
		WHEN SUM(dbo.DB_BankAccountStatement.Balance) < 0 THEN 0 - SUM(dbo.DB_BankAccountStatement.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.DB_BankAccountStatement.Balance) >= 0 THEN SUM(dbo.DB_BankAccountStatement.Balance)
		ELSE 0 END AS Credit,DB_BankAccountStatement.TransDate
FROM         dbo.DB_BankAccountStatement INNER JOIN
                      dbo.AccountList ON dbo.DB_BankAccountStatement.BankAccount_DPA_ = dbo.AccountList.Account_DPA_
WHERE 		dbo.AccountList.AccountTypeLevel1_DPA_ = 5
GROUP BY dbo.AccountList.AccountTypeLevel1_DPA_,dbo.AccountList.AccountTypeLevel1,dbo.AccountList.AccountName, dbo.AccountList.Account_DPA_,DB_BankAccountStatement.TransDate



SELECT  IDENTITY(int, 1,1)  AS SequenceID, #tableA.AccountType_DPA_,#tableA.AccountType,#tableA.[Account Code],#tableA.AccountName,Sum(#tableA.Credit - #tableA.Debit) AS Balance
	 INTO #tableB FROM #tableA Where #tableA.TransDate between @StartDate and @enddate Group by #tableA.AccountType_DPA_,#tableA.AccountType,#tableA.[Account Code],#tableA.AccountName ORDER BY AccountType_DPA_

DROP TABLE #tableA

select T.SequenceID,T.AccountType_DPA_,T.AccountType,T.[Account Code],T.AccountName,T.Balance,
       case when SequenceID = (select top 1 SequenceID from #tableB
                           where AccountType_DPA_
                              = T.AccountType_DPA_
                          order by SequenceID desc)
            then (select CONVERT(CHAR(20),sum(Balance))
                     from #tableB
                     where SequenceID <= T.SequenceID
                        and AccountType_DPA_
                              = T.AccountType_DPA_)
            else ' ' end as 'SubTotal',
       case when SequenceID = (select top 1 SequenceID from #tableB
                           order by SequenceID desc)
            then (select CONVERT(CHAR(20),sum(Balance))
                      from #tableB)
             else ' ' end as 'GrandTotal'
from #tableB T
  order by SequenceID



DROP TABLE #tableB


GO
CREATE PROCEDURE [CustodianContractLevies_v2]  @StartDate datetime, @EndDate datetime  AS
DECLARE @SQLStatement nvarchar(4000)
EXEC sp_pivot 'MAX', 'LevyAmount', 'LevyContracts_v2', 'Contract_DPA_', 'LevyShortName', @outSQL = @SQLStatement OUTPUT
IF @SQLStatement = NULL
	BEGIN
		 SELECT  @SQLStatement AS RESULT_FLD WHERE 0 = 1
	END

ELSE
	 BEGIN
--AND (dbo.Lot.LotTDate =  ' +  @StartDate + '   AND  ' +  @EndDate + '  )
EXEC ('SELECT   *  FROM (SELECT   TOP 100 PERCENT  CONVERT(SMALLDATETIME, CAST(dbo.Lot.LotTDate AS CHAR(12))) AS LotTDate, dbo.Lot.Contract_DPA_ AS Contract,
                      dbo.Lot.ContractNumber AS ContractNumber, dbo.OrdDetailList.OrderTypeSale, dbo.Security.SecurityCode, dbo.OrdDetailList.OrderSecType_DPA_,
                      dbo.Broker.BrokerCode, dbo.Lot.LotSlipNo, dbo.Lot.LotPrice, dbo.Lot.LotQty, dbo.OrdDetailList.OrdDetailClient AS ClientName,
                      dbo.OrdDetailList.Client_DPA_ AS Client_DPA_, dbo.Lot.Contract_DPA_, dbo.Lot.LotGrossAmount, dbo.tbOrder.IsCustodian,
                      dbo.Client.GenericSetting_DPA_
FROM         dbo.OrdDetailList INNER JOIN
                      dbo.Lot ON dbo.OrdDetailList.OrdDetail_DPA_ = dbo.Lot.OrdDetail_DPA_ INNER JOIN
                      dbo.Broker ON dbo.Lot.Broker_DPA_ = dbo.Broker.Broker_DPA_ INNER JOIN
                      dbo.tbOrder ON dbo.OrdDetailList.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN
                      dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ LEFT OUTER JOIN
                      dbo.Security ON dbo.OrdDetailList.Security_DPA_ = dbo.Security.Security_DPA_
WHERE     (dbo.Lot.Deleted <> 1)  AND (dbo.Client.Class_DPA_ = 3)    AND (dbo.Lot.LotTDate  between ''' +  @StartDate + '''   AND ''' +  @EndDate + ''')
Order by dbo.Lot.LotTDate,dbo.Security.SecurityCode,dbo.Lot.ContractNumber DESC) AS INNERTBL
INNER JOIN (' + @SQLStatement + ') CROSSTABTBL ON CONVERT(NVARCHAR(1000), INNERTBL.Contract_DPA_) = CONVERT(NVARCHAR(1000), CROSSTABTBL.Contract_DPA_  )')
END
GO

CREATE PROCEDURE DebtorsAndCreditors
	@TransDate DATETIME

AS
SELECT     	DISTINCT
		a.Client_DPA_,
		MAX(b.TransDate) AS LastDate,
		SUM(b.Balance) AS Balance

FROM         (SELECT * FROM dbo.ClientTransactionList) a CROSS JOIN
                          (SELECT * FROM dbo.ClientTransactionList) b
WHERE
	--a.TransDate >= b.TransDate
	a.Client_DPA_ = b.Client_DPA_
	AND a.TransDate <= @TransDate
	AND b.TransDate <= @TransDate
GROUP BY a.Client_DPA_, a.TransDate
ORDER BY MAX(b.TransDate)  DESC

GO

CREATE PROCEDURE DeleteApplication
@clientDPA int,
@offerBank int,
@jDate smalldatetime,
@offerQty money,
@offerPrice money,
@JournalNarrative varchar(500),
@userID int
,
@Certificate int
--@CDSCharge float
--@CDSBank int
--@TimeChanged smalldatetime
--@JournalEIT nvarchar(500)
AS
BEGIN
	---DECLARE JOURNAL SETTINGS---
	DECLARE @JournalDPA int
	DECLARE @JournalEntryDPA int
	--SET @JournalDPA = (SELECT MAX(ISNULL(Journal_DPA_,0))+1 FROM Journal)

	DECLARE @guid uniqueidentifier
	SET @guid = NEWID()

	--DECLARE @ReleasedBy int
	--SET @ReleasedBy = 10390 --system account at dyer

	DECLARE @ReleaseDate datetime
	SET @ReleaseDate = GetDate()

	DECLARE @TimeChanged smalldatetime
	SET @TimeChanged = GetDate()
	---DECLARE JOURNAL ENTRY SETTINGS---
	DECLARE @DebitAmount money
	DECLARE @CreditAmount money

	---CREATE JOURNAL---
	INSERT INTO Journal
	(
		--Journal_DPA_,
		Journal_EIT_,
		JournalDate,
		JournalNarrative,
		UserID,
		JournalCommitted,
		ChangedBy,
		TimeChanged,
		Released,
		ReleaseDate
	)
	VALUES
	(
		--@JournalDPA,
		@guid,
		@JDate,
		@JournalNarrative,
		@userID,
		1,
		@userID,
		@TimeChanged,
		1,
		@ReleaseDate
	)

	SET @JournalDPA = (SELECT @@IDENTITY)
	---CREATE JOURNAL ENTRIES---
	Begin
		Set @DebitAmount = @offerQty * @offerPrice
		Set @CreditAmount = 0

		--a) client account
		Begin
			Set @JournalEntryDPA = (Select Max(ISNULL(JournalEntry_DPA_,0))+1 From JournalEntry)

			Insert into JournalEntry
				(
					JournalEntry_DPA_,
					EntityType_DPA_,
					Entity_DPA_,
					JournalEntryDebit,
					JournalEntryCredit,
					Journal_DPA_,
					--Narrative,
					ChangedBy,
					TimeChanged
				)

				Values
				(
					@JournalEntryDPA,
					1,
					@clientDPA,
					@DebitAmount,
					@CreditAmount,
					@JournalDPA,
					--@JournalNarrative,
					@userID,
					@TimeChanged
				)

			/*IF @Certificate=0
				BEGIN
				Set @JournalEntryDPA = (Select Max(ISNULL(JournalEntry_DPA_,0))+1 From JournalEntry)
				Insert into JournalEntry
				(
					JournalEntry_DPA_,
					EntityType_DPA_,
					Entity_DPA_,
					JournalEntryDebit,
					JournalEntryCredit,
					Journal_DPA_,
					--Narrative,
					ChangedBy,
					TimeChanged
				)

				Values
				(
					@JournalEntryDPA,
					1,
					@clientDPA,
					@CDSCharge,
					0,
					@JournalDPA,
					--@JournalNarrative,
					@userID,
					@TimeChanged
				)
			END*/
		End
		--b) IPO account
		Begin
			Set @JournalEntryDPA = (Select Max(ISNULL(JournalEntry_DPA_,0))+1 From JournalEntry)

			Insert into JournalEntry
				(
					JournalEntry_DPA_,
					EntityType_DPA_,
					Entity_DPA_,
					JournalEntryDebit,
					JournalEntryCredit,
					Journal_DPA_,
					--Narrative,
					ChangedBy,
					TimeChanged
				)

				Values
				(
					@JournalEntryDPA,
					5,
					@offerBank,
					@CreditAmount,
					@DebitAmount,
					@JournalDPA,
					--@JournalNarrative,
					@userID,
					@TimeChanged
				)

				/*IF @Certificate=0
				BEGIN
				Set @JournalEntryDPA = (Select Max(ISNULL(JournalEntry_DPA_,0))+1 From JournalEntry)
				Insert into JournalEntry
				(
					JournalEntry_DPA_,
					EntityType_DPA_,
					Entity_DPA_,
					JournalEntryDebit,
					JournalEntryCredit,
					Journal_DPA_,
					--Narrative,
					ChangedBy,
					TimeChanged
				)

				Values
				(
					@JournalEntryDPA,
					5,
					@CDSBank,
					0,
					@CDSCharge,
					@JournalDPA,
					--@JournalNarrative,
					@userID,
					@TimeChanged
				)
			END*/
		End
	End
	--NOW CREATE REVERSE ENTRIES
	Begin
		---CREATE JOURNAL---
		SET @JournalNarrative = 'REVERSAL: ' + @JournalNarrative
		SET @guid = NEWID()
		--SET @JournalDPA = (SELECT MAX(ISNULL(Journal_DPA_,0))+1 FROM Journal)

		INSERT INTO Journal
		(
			--Journal_DPA_,
			Journal_EIT_,
			JournalDate,
			JournalNarrative,
			UserID,
			JournalCommitted,
			ChangedBy,
			TimeChanged,
			Released,
			ReleaseDate
		)
		VALUES
		(
			--@JournalDPA,
			@guid,
			@JDate,
			@JournalNarrative,
			@userID,
			1,
			@userID,
			@TimeChanged,
			1,
			@ReleaseDate
		)

		SET @JournalDPA = (SELECT @@IDENTITY)
		---CREATE JOURNAL ENTRIES---
		Begin
			--do the opposite
			Set @DebitAmount = 0
			Set @CreditAmount = @offerQty * @offerPrice

			--a) client account
			Begin
				Set @JournalEntryDPA = (Select Max(ISNULL(JournalEntry_DPA_,0))+1 From JournalEntry)

				Insert into JournalEntry
					(
						JournalEntry_DPA_,
						EntityType_DPA_,
						Entity_DPA_,
						JournalEntryDebit,
						JournalEntryCredit,
						Journal_DPA_,
						--Narrative,
						ChangedBy,
						TimeChanged
					)

					Values
					(
						@JournalEntryDPA,
						1,
						@clientDPA,
						@DebitAmount,
						@CreditAmount,
						@JournalDPA,
						--@JournalNarrative,
						@userID,
						@TimeChanged
					)

				/*IF @Certificate=0
				BEGIN
				Set @JournalEntryDPA = (Select Max(ISNULL(JournalEntry_DPA_,0))+1 From JournalEntry)
				Insert into JournalEntry
				(
					JournalEntry_DPA_,
					EntityType_DPA_,
					Entity_DPA_,
					JournalEntryDebit,
					JournalEntryCredit,
					Journal_DPA_,
					--Narrative,
					ChangedBy,
					TimeChanged
				)

				Values
				(
					@JournalEntryDPA,
					1,
					@ClientDPA,
					0,
					@CDSCharge,
					@JournalDPA,
					--@JournalNarrative,
					@userID,
					@TimeChanged
				)
			END*/
			End

			--b) IPO account
			Begin
				Set @JournalEntryDPA = (Select Max(ISNULL(JournalEntry_DPA_,0))+1 From JournalEntry)

				Insert into JournalEntry
					(
						JournalEntry_DPA_,
						EntityType_DPA_,
						Entity_DPA_,
						JournalEntryDebit,
						JournalEntryCredit,
						Journal_DPA_,
						--Narrative,
						ChangedBy,
						TimeChanged
					)

					Values
					(
						@JournalEntryDPA,
						5,
						@offerBank,
						@CreditAmount,
						@DebitAmount,
						@JournalDPA,
						--@JournalNarrative,
						@userID,
						@TimeChanged
					)

				/*IF @Certificate=0
				BEGIN
				Set @JournalEntryDPA = (Select Max(ISNULL(JournalEntry_DPA_,0))+1 From JournalEntry)
				Insert into JournalEntry
				(
					JournalEntry_DPA_,
					EntityType_DPA_,
					Entity_DPA_,
					JournalEntryDebit,
					JournalEntryCredit,
					Journal_DPA_,
					--Narrative,
					ChangedBy,
					TimeChanged
				)

				Values
				(
					@JournalEntryDPA,
					5,
					@CDSBank,
					@CDSCharge,
					0,
					@JournalDPA,
					--@JournalNarrative,
					@userID,
					@TimeChanged
				)
			END*/
			End
		End
	End


END




GO






CREATE  Proc DeleteForward
@clientDPA int,
@offerBank int,
@jDate smalldatetime,
@offerQty money,
@offerPrice money,
@JournalNarrative varchar(500),
@userID int,
@Certificate int,
@CDSCharge float,
@CDSBank int
--@TimeChanged smalldatetime
--@JournalEIT nvarchar(500)
as
BEGIN
	---DECLARE JOURNAL SETTINGS---
	if @Certificate = 1
		goto exit_Proc
	DECLARE @JournalDPA int
	DECLARE @JournalEntryDPA int
	SET @JournalDPA = (SELECT MAX(ISNULL(Journal_DPA_,0))+1 FROM Journal)

	DECLARE @guid uniqueidentifier
	SET @guid = NEWID()

	--DECLARE @ReleasedBy int
	--SET @ReleasedBy = 10390 --system account at dyer

	DECLARE @ReleaseDate datetime
	SET @ReleaseDate = GetDate()

	DECLARE @TimeChanged smalldatetime
	SET @TimeChanged = GetDate()
	---DECLARE JOURNAL ENTRY SETTINGS---
	DECLARE @DebitAmount money
	DECLARE @CreditAmount money

	---CREATE JOURNAL---
	INSERT INTO Journal
	(
		Journal_DPA_,
		Journal_EIT_,
		JournalDate,
		JournalNarrative,
		UserID,
		JournalCommitted,
		ChangedBy,
		TimeChanged,
		Released,
		ReleaseDate
	)
	VALUES
	(
		@JournalDPA,
		@guid,
		@jDate,
		@JournalNarrative,
		@userID,
		1,
		@userID,
		@TimeChanged,
		1,
		@ReleaseDate
	)
	---CREATE JOURNAL ENTRIES---
	Begin
		Set @DebitAmount = @cdsCharge
		Set @CreditAmount = 0

		--a) client account
		Begin
			Set @JournalEntryDPA = (Select Max(ISNULL(JournalEntry_DPA_,0))+1 From JournalEntry)

			IF @Certificate=0
				BEGIN
				--Set @JournalEntryDPA = (Select Max(ISNULL(JournalEntry_DPA_,0))+1 From JournalEntry)
				Insert into JournalEntry
				(
					JournalEntry_DPA_,
					EntityType_DPA_,
					Entity_DPA_,
					JournalEntryDebit,
					JournalEntryCredit,
					Journal_DPA_,
					Narrative,
					ChangedBy,
					TimeChanged
				)

				Values
				(
					@JournalEntryDPA,
					1,
					@clientDPA,
					@CDSCharge,
					0,
					@JournalDPA,
					@JournalNarrative,
					@userID,
					@TimeChanged
				)
			END
		End
		--b) IPO account
		Begin
			IF @Certificate=0
			BEGIN
			Set @JournalEntryDPA = (Select Max(ISNULL(JournalEntry_DPA_,0))+1 From JournalEntry)
			Insert into JournalEntry
			(
				JournalEntry_DPA_,
				EntityType_DPA_,
				Entity_DPA_,
				JournalEntryDebit,
				JournalEntryCredit,
				Journal_DPA_,
				Narrative,
				ChangedBy,
				TimeChanged
			)

			Values
			(
				@JournalEntryDPA,
				5,
				@CDSBank,
				0,
				@CDSCharge,
				@JournalDPA,
				@JournalNarrative,
				@userID,
				@TimeChanged
			)
		END
		End
	End

	--NOW CREATE REVERSE ENTRIES
	Begin
		---CREATE JOURNAL---
		SET @JournalNarrative = 'REVERSAL: ' + @JournalNarrative
		SET @guid = NEWID()
		SET @JournalDPA = (SELECT MAX(ISNULL(Journal_DPA_,0))+1 FROM Journal)

		INSERT INTO Journal
		(
			Journal_DPA_,
			Journal_EIT_,
			JournalDate,
			JournalNarrative,
			UserID,
			JournalCommitted,
			ChangedBy,
			TimeChanged,
			Released,
			ReleaseDate
		)
		VALUES
		(
			@JournalDPA,
			@guid,
			GetDate(),
			@JournalNarrative,
			@userID,
			1,
			@userID,
			@TimeChanged,
			1,
			@ReleaseDate
		)

		---CREATE JOURNAL ENTRIES---
		Begin
			--do the opposite
			Set @DebitAmount = 0
			Set @CreditAmount = @CDSCharge

			--a) client account
			Begin
				IF @Certificate=0
				BEGIN
				Set @JournalEntryDPA = (Select Max(ISNULL(JournalEntry_DPA_,0))+1 From JournalEntry)
				Insert into JournalEntry
				(
					JournalEntry_DPA_,
					EntityType_DPA_,
					Entity_DPA_,
					JournalEntryDebit,
					JournalEntryCredit,
					Journal_DPA_,
					Narrative,
					ChangedBy,
					TimeChanged
				)

				Values
				(
					@JournalEntryDPA,
					1,
					@ClientDPA,
					0,
					@CDSCharge,
					@JournalDPA,
					@JournalNarrative,
					@userID,
					@TimeChanged
				)
			END
			End

			--b) IPO account
			Begin
				IF @Certificate=0
				BEGIN
				Set @JournalEntryDPA = (Select Max(ISNULL(JournalEntry_DPA_,0))+1 From JournalEntry)
				Insert into JournalEntry
				(
					JournalEntry_DPA_,
					EntityType_DPA_,
					Entity_DPA_,
					JournalEntryDebit,
					JournalEntryCredit,
					Journal_DPA_,
					Narrative,
					ChangedBy,
					TimeChanged
				)

				Values
				(
					@JournalEntryDPA,
					5,
					@CDSBank,
					@CDSCharge,
					0,
					@JournalDPA,
					@JournalNarrative,
					@userID,
					@TimeChanged
				)
			END
			End
		End
	End





exit_Proc:


end







GO
create proc dbo.dt_addtosourcecontrol
    @vchSourceSafeINI varchar(255) = '',
    @vchProjectName   varchar(255) ='',
    @vchComment       varchar(255) ='',
    @vchLoginName     varchar(255) ='',
    @vchPassword      varchar(255) =''

as

set nocount on

declare @iReturn int
declare @iObjectId int
select @iObjectId = 0

declare @iStreamObjectId int
select @iStreamObjectId = 0

declare @VSSGUID varchar(100)
select @VSSGUID = 'SQLVersionControl.VCS_SQL'

declare @vchDatabaseName varchar(255)
select @vchDatabaseName = db_name()

declare @iReturnValue int
select @iReturnValue = 0

declare @iPropertyObjectId int
declare @vchParentId varchar(255)

declare @iObjectCount int
select @iObjectCount = 0

    exec @iReturn = master.dbo.sp_OACreate @VSSGUID, @iObjectId OUT
    if @iReturn <> 0 GOTO E_OAError


    /* Create Project in SS */
    exec @iReturn = master.dbo.sp_OAMethod @iObjectId,
											'AddProjectToSourceSafe',
											NULL,
											@vchSourceSafeINI,
											@vchProjectName output,
											@@SERVERNAME,
											@vchDatabaseName,
											@vchLoginName,
											@vchPassword,
											@vchComment


    if @iReturn <> 0 GOTO E_OAError

    /* Set Database Properties */

    begin tran SetProperties

    /* add high level object */

    exec @iPropertyObjectId = dbo.dt_adduserobject_vcs 'VCSProjectID'

    select @vchParentId = CONVERT(varchar(255),@iPropertyObjectId)

    exec dbo.dt_setpropertybyid @iPropertyObjectId, 'VCSProjectID', @vchParentId , NULL
    exec dbo.dt_setpropertybyid @iPropertyObjectId, 'VCSProject' , @vchProjectName , NULL
    exec dbo.dt_setpropertybyid @iPropertyObjectId, 'VCSSourceSafeINI' , @vchSourceSafeINI , NULL
    exec dbo.dt_setpropertybyid @iPropertyObjectId, 'VCSSQLServer', @@SERVERNAME, NULL
    exec dbo.dt_setpropertybyid @iPropertyObjectId, 'VCSSQLDatabase', @vchDatabaseName, NULL

    if @@error <> 0 GOTO E_General_Error

    commit tran SetProperties

    select @iObjectCount = 0;

CleanUp:
    select @vchProjectName
    select @iObjectCount
    return

E_General_Error:
    /* this is an all or nothing.  No specific error messages */
    goto CleanUp

E_OAError:
    exec dbo.dt_displayoaerror @iObjectId, @iReturn
    goto CleanUp



GO
create proc dbo.dt_addtosourcecontrol_u
    @vchSourceSafeINI nvarchar(255) = '',
    @vchProjectName   nvarchar(255) ='',
    @vchComment       nvarchar(255) ='',
    @vchLoginName     nvarchar(255) ='',
    @vchPassword      nvarchar(255) =''

as
	-- This procedure should no longer be called;  dt_addtosourcecontrol should be called instead.
	-- Calls are forwarded to dt_addtosourcecontrol to maintain backward compatibility
	set nocount on
	exec dbo.dt_addtosourcecontrol
		@vchSourceSafeINI,
		@vchProjectName,
		@vchComment,
		@vchLoginName,
		@vchPassword



GO
/*
**	Add an object to the dtproperties table
*/
create procedure dbo.dt_adduserobject
as
	set nocount on
	/*
	** Create the user object if it does not exist already
	*/
	begin transaction
		insert dbo.dtproperties (property) VALUES ('DtgSchemaOBJECT')
		update dbo.dtproperties set objectid=@@identity
			where id=@@identity and property='DtgSchemaOBJECT'
	commit
	return @@identity

GO
create procedure dbo.dt_adduserobject_vcs
    @vchProperty varchar(64)

as

set nocount on

declare @iReturn int
    /*
    ** Create the user object if it does not exist already
    */
    begin transaction
        select @iReturn = objectid from dbo.dtproperties where property = @vchProperty
        if @iReturn IS NULL
        begin
            insert dbo.dtproperties (property) VALUES (@vchProperty)
            update dbo.dtproperties set objectid=@@identity
                    where id=@@identity and property=@vchProperty
            select @iReturn = @@identity
        end
    commit
    return @iReturn



GO
create proc dbo.dt_checkinobject
    @chObjectType  char(4),
    @vchObjectName varchar(255),
    @vchComment    varchar(255)='',
    @vchLoginName  varchar(255),
    @vchPassword   varchar(255)='',
    @iVCSFlags     int = 0,
    @iActionFlag   int = 0,   /* 0 => AddFile, 1 => CheckIn */
    @txStream1     Text = '', /* drop stream   */ /* There is a bug that if items are NULL they do not pass to OLE servers */
    @txStream2     Text = '', /* create stream */
    @txStream3     Text = ''  /* grant stream  */


as

	set nocount on

	declare @iReturn int
	declare @iObjectId int
	select @iObjectId = 0
	declare @iStreamObjectId int

	declare @VSSGUID varchar(100)
	select @VSSGUID = 'SQLVersionControl.VCS_SQL'

	declare @iPropertyObjectId int
	select @iPropertyObjectId  = 0

    select @iPropertyObjectId = (select objectid from dbo.dtproperties where property = 'VCSProjectID')

    declare @vchProjectName   varchar(255)
    declare @vchSourceSafeINI varchar(255)
    declare @vchServerName    varchar(255)
    declare @vchDatabaseName  varchar(255)
    declare @iReturnValue	  int
    declare @pos			  int
    declare @vchProcLinePiece varchar(255)


    exec dbo.dt_getpropertiesbyid_vcs @iPropertyObjectId, 'VCSProject',       @vchProjectName   OUT
    exec dbo.dt_getpropertiesbyid_vcs @iPropertyObjectId, 'VCSSourceSafeINI', @vchSourceSafeINI OUT
    exec dbo.dt_getpropertiesbyid_vcs @iPropertyObjectId, 'VCSSQLServer',     @vchServerName    OUT
    exec dbo.dt_getpropertiesbyid_vcs @iPropertyObjectId, 'VCSSQLDatabase',   @vchDatabaseName  OUT

    if @chObjectType = 'PROC'
    begin
        if @iActionFlag = 1
        begin
            /* Procedure Can have up to three streams
            Drop Stream, Create Stream, GRANT stream */

            begin tran compile_all

            /* try to compile the streams */
            exec (@txStream1)
            if @@error <> 0 GOTO E_Compile_Fail

            exec (@txStream2)
            if @@error <> 0 GOTO E_Compile_Fail

            exec (@txStream3)
            if @@error <> 0 GOTO E_Compile_Fail
        end

        exec @iReturn = master.dbo.sp_OACreate @VSSGUID, @iObjectId OUT
        if @iReturn <> 0 GOTO E_OAError

        exec @iReturn = master.dbo.sp_OAGetProperty @iObjectId, 'GetStreamObject', @iStreamObjectId OUT
        if @iReturn <> 0 GOTO E_OAError

        if @iActionFlag = 1
        begin

            declare @iStreamLength int

			select @pos=1
			select @iStreamLength = datalength(@txStream2)

			if @iStreamLength > 0
			begin

				while @pos < @iStreamLength
				begin

					select @vchProcLinePiece = substring(@txStream2, @pos, 255)

					exec @iReturn = master.dbo.sp_OAMethod @iStreamObjectId, 'AddStream', @iReturnValue OUT, @vchProcLinePiece
            		if @iReturn <> 0 GOTO E_OAError

					select @pos = @pos + 255

				end

				exec @iReturn = master.dbo.sp_OAMethod @iObjectId,
														'CheckIn_StoredProcedure',
														NULL,
														@sProjectName = @vchProjectName,
														@sSourceSafeINI = @vchSourceSafeINI,
														@sServerName = @vchServerName,
														@sDatabaseName = @vchDatabaseName,
														@sObjectName = @vchObjectName,
														@sComment = @vchComment,
														@sLoginName = @vchLoginName,
														@sPassword = @vchPassword,
														@iVCSFlags = @iVCSFlags,
														@iActionFlag = @iActionFlag,
														@sStream = ''

			end
        end
        else
        begin

            select colid, text into #ProcLines
            from syscomments
            where id = object_id(@vchObjectName)
            order by colid

            declare @iCurProcLine int
            declare @iProcLines int
            select @iCurProcLine = 1
            select @iProcLines = (select count(*) from #ProcLines)
            while @iCurProcLine <= @iProcLines
            begin
                select @pos = 1
                declare @iCurLineSize int
                select @iCurLineSize = len((select text from #ProcLines where colid = @iCurProcLine))
                while @pos <= @iCurLineSize
                begin
                    select @vchProcLinePiece = convert(varchar(255),
                        substring((select text from #ProcLines where colid = @iCurProcLine),
                                  @pos, 255 ))
                    exec @iReturn = master.dbo.sp_OAMethod @iStreamObjectId, 'AddStream', @iReturnValue OUT, @vchProcLinePiece
                    if @iReturn <> 0 GOTO E_OAError
                    select @pos = @pos + 255
                end
                select @iCurProcLine = @iCurProcLine + 1
            end
            drop table #ProcLines

            exec @iReturn = master.dbo.sp_OAMethod @iObjectId,
													'CheckIn_StoredProcedure',
													NULL,
													@sProjectName = @vchProjectName,
													@sSourceSafeINI = @vchSourceSafeINI,
													@sServerName = @vchServerName,
													@sDatabaseName = @vchDatabaseName,
													@sObjectName = @vchObjectName,
													@sComment = @vchComment,
													@sLoginName = @vchLoginName,
													@sPassword = @vchPassword,
													@iVCSFlags = @iVCSFlags,
													@iActionFlag = @iActionFlag,
													@sStream = ''
        end

        if @iReturn <> 0 GOTO E_OAError

        if @iActionFlag = 1
        begin
            commit tran compile_all
            if @@error <> 0 GOTO E_Compile_Fail
        end

    end

CleanUp:
	return

E_Compile_Fail:
	declare @lerror int
	select @lerror = @@error
	rollback tran compile_all
	RAISERROR (@lerror,16,-1)
	goto CleanUp

E_OAError:
	if @iActionFlag = 1 rollback tran compile_all
	exec dbo.dt_displayoaerror @iObjectId, @iReturn
	goto CleanUp



GO
create proc dbo.dt_checkinobject_u
    @chObjectType  char(4),
    @vchObjectName nvarchar(255),
    @vchComment    nvarchar(255)='',
    @vchLoginName  nvarchar(255),
    @vchPassword   nvarchar(255)='',
    @iVCSFlags     int = 0,
    @iActionFlag   int = 0,   /* 0 => AddFile, 1 => CheckIn */
    @txStream1     text = '',  /* drop stream   */ /* There is a bug that if items are NULL they do not pass to OLE servers */
    @txStream2     text = '',  /* create stream */
    @txStream3     text = ''   /* grant stream  */

as
	-- This procedure should no longer be called;  dt_checkinobject should be called instead.
	-- Calls are forwarded to dt_checkinobject to maintain backward compatibility.
	set nocount on
	exec dbo.dt_checkinobject
		@chObjectType,
		@vchObjectName,
		@vchComment,
		@vchLoginName,
		@vchPassword,
		@iVCSFlags,
		@iActionFlag,
		@txStream1,
		@txStream2,
		@txStream3



GO
create proc dbo.dt_checkoutobject
    @chObjectType  char(4),
    @vchObjectName varchar(255),
    @vchComment    varchar(255),
    @vchLoginName  varchar(255),
    @vchPassword   varchar(255),
    @iVCSFlags     int = 0,
    @iActionFlag   int = 0/* 0 => Checkout, 1 => GetLatest, 2 => UndoCheckOut */

as

	set nocount on

	declare @iReturn int
	declare @iObjectId int
	select @iObjectId =0

	declare @VSSGUID varchar(100)
	select @VSSGUID = 'SQLVersionControl.VCS_SQL'

	declare @iReturnValue int
	select @iReturnValue = 0

	declare @vchTempText varchar(255)

	/* this is for our strings */
	declare @iStreamObjectId int
	select @iStreamObjectId = 0

    declare @iPropertyObjectId int
    select @iPropertyObjectId = (select objectid from dbo.dtproperties where property = 'VCSProjectID')

    declare @vchProjectName   varchar(255)
    declare @vchSourceSafeINI varchar(255)
    declare @vchServerName    varchar(255)
    declare @vchDatabaseName  varchar(255)
    exec dbo.dt_getpropertiesbyid_vcs @iPropertyObjectId, 'VCSProject',       @vchProjectName   OUT
    exec dbo.dt_getpropertiesbyid_vcs @iPropertyObjectId, 'VCSSourceSafeINI', @vchSourceSafeINI OUT
    exec dbo.dt_getpropertiesbyid_vcs @iPropertyObjectId, 'VCSSQLServer',     @vchServerName    OUT
    exec dbo.dt_getpropertiesbyid_vcs @iPropertyObjectId, 'VCSSQLDatabase',   @vchDatabaseName  OUT

    if @chObjectType = 'PROC'
    begin
        /* Procedure Can have up to three streams
           Drop Stream, Create Stream, GRANT stream */

        exec @iReturn = master.dbo.sp_OACreate @VSSGUID, @iObjectId OUT

        if @iReturn <> 0 GOTO E_OAError

        exec @iReturn = master.dbo.sp_OAMethod @iObjectId,
												'CheckOut_StoredProcedure',
												NULL,
												@sProjectName = @vchProjectName,
												@sSourceSafeINI = @vchSourceSafeINI,
												@sObjectName = @vchObjectName,
												@sServerName = @vchServerName,
												@sDatabaseName = @vchDatabaseName,
												@sComment = @vchComment,
												@sLoginName = @vchLoginName,
												@sPassword = @vchPassword,
												@iVCSFlags = @iVCSFlags,
												@iActionFlag = @iActionFlag

        if @iReturn <> 0 GOTO E_OAError


        exec @iReturn = master.dbo.sp_OAGetProperty @iObjectId, 'GetStreamObject', @iStreamObjectId OUT

        if @iReturn <> 0 GOTO E_OAError

        create table #commenttext (id int identity, sourcecode varchar(255))


        select @vchTempText = 'STUB'
        while @vchTempText is not null
        begin
            exec @iReturn = master.dbo.sp_OAMethod @iStreamObjectId, 'GetStream', @iReturnValue OUT, @vchTempText OUT
            if @iReturn <> 0 GOTO E_OAError

            if (@vchTempText = '') set @vchTempText = null
            if (@vchTempText is not null) insert into #commenttext (sourcecode) select @vchTempText
        end

        select 'VCS'=sourcecode from #commenttext order by id
        select 'SQL'=text from syscomments where id = object_id(@vchObjectName) order by colid

    end

CleanUp:
    return

E_OAError:
    exec dbo.dt_displayoaerror @iObjectId, @iReturn
    GOTO CleanUp



GO
create proc dbo.dt_checkoutobject_u
    @chObjectType  char(4),
    @vchObjectName nvarchar(255),
    @vchComment    nvarchar(255),
    @vchLoginName  nvarchar(255),
    @vchPassword   nvarchar(255),
    @iVCSFlags     int = 0,
    @iActionFlag   int = 0/* 0 => Checkout, 1 => GetLatest, 2 => UndoCheckOut */

as

	-- This procedure should no longer be called;  dt_checkoutobject should be called instead.
	-- Calls are forwarded to dt_checkoutobject to maintain backward compatibility.
	set nocount on
	exec dbo.dt_checkoutobject
		@chObjectType,
		@vchObjectName,
		@vchComment,
		@vchLoginName,
		@vchPassword,
		@iVCSFlags,
		@iActionFlag



GO
CREATE PROCEDURE dbo.dt_displayoaerror
    @iObject int,
    @iresult int
as

set nocount on

declare @vchOutput      varchar(255)
declare @hr             int
declare @vchSource      varchar(255)
declare @vchDescription varchar(255)

    exec @hr = master.dbo.sp_OAGetErrorInfo @iObject, @vchSource OUT, @vchDescription OUT

    select @vchOutput = @vchSource + ': ' + @vchDescription
    raiserror (@vchOutput,16,-1)

    return


GO
CREATE PROCEDURE dbo.dt_displayoaerror_u
    @iObject int,
    @iresult int
as
	-- This procedure should no longer be called;  dt_displayoaerror should be called instead.
	-- Calls are forwarded to dt_displayoaerror to maintain backward compatibility.
	set nocount on
	exec dbo.dt_displayoaerror
		@iObject,
		@iresult



GO
/*
**	Drop one or all the associated properties of an object or an attribute
**
**	dt_dropproperties objid, null or '' -- drop all properties of the object itself
**	dt_dropproperties objid, property -- drop the property
*/
create procedure dbo.dt_droppropertiesbyid
	@id int,
	@property varchar(64)
as
	set nocount on

	if (@property is null) or (@property = '')
		delete from dbo.dtproperties where objectid=@id
	else
		delete from dbo.dtproperties
			where objectid=@id and property=@property


GO
/*
**	Drop an object from the dbo.dtproperties table
*/
create procedure dbo.dt_dropuserobjectbyid
	@id int
as
	set nocount on
	delete from dbo.dtproperties where objectid=@id

GO
/*
**	Generate an ansi name that is unique in the dtproperties.value column
*/
create procedure dbo.dt_generateansiname(@name varchar(255) output)
as
	declare @prologue varchar(20)
	declare @indexstring varchar(20)
	declare @index integer

	set @prologue = 'MSDT-A-'
	set @index = 1

	while 1 = 1
	begin
		set @indexstring = cast(@index as varchar(20))
		set @name = @prologue + @indexstring
		if not exists (select value from dtproperties where value = @name)
			break

		set @index = @index + 1

		if (@index = 10000)
			goto TooMany
	end

Leave:

	return

TooMany:

	set @name = 'DIAGRAM'
	goto Leave

GO
/*
**	Retrieve the owner object(s) of a given property
*/
create procedure dbo.dt_getobjwithprop
	@property varchar(30),
	@value varchar(255)
as
	set nocount on

	if (@property is null) or (@property = '')
	begin
		raiserror('Must specify a property name.',-1,-1)
		return (1)
	end

	if (@value is null)
		select objectid id from dbo.dtproperties
			where property=@property

	else
		select objectid id from dbo.dtproperties
			where property=@property and value=@value

GO
/*
**	Retrieve the owner object(s) of a given property
*/
create procedure dbo.dt_getobjwithprop_u
	@property varchar(30),
	@uvalue nvarchar(255)
as
	set nocount on

	if (@property is null) or (@property = '')
	begin
		raiserror('Must specify a property name.',-1,-1)
		return (1)
	end

	if (@uvalue is null)
		select objectid id from dbo.dtproperties
			where property=@property

	else
		select objectid id from dbo.dtproperties
			where property=@property and uvalue=@uvalue

GO
/*
**	Retrieve properties by id's
**
**	dt_getproperties objid, null or '' -- retrieve all properties of the object itself
**	dt_getproperties objid, property -- retrieve the property specified
*/
create procedure dbo.dt_getpropertiesbyid
	@id int,
	@property varchar(64)
as
	set nocount on

	if (@property is null) or (@property = '')
		select property, version, value, lvalue
			from dbo.dtproperties
			where  @id=objectid
	else
		select property, version, value, lvalue
			from dbo.dtproperties
			where  @id=objectid and @property=property

GO
/*
**	Retrieve properties by id's
**
**	dt_getproperties objid, null or '' -- retrieve all properties of the object itself
**	dt_getproperties objid, property -- retrieve the property specified
*/
create procedure dbo.dt_getpropertiesbyid_u
	@id int,
	@property varchar(64)
as
	set nocount on

	if (@property is null) or (@property = '')
		select property, version, uvalue, lvalue
			from dbo.dtproperties
			where  @id=objectid
	else
		select property, version, uvalue, lvalue
			from dbo.dtproperties
			where  @id=objectid and @property=property

GO
create procedure dbo.dt_getpropertiesbyid_vcs
    @id       int,
    @property varchar(64),
    @value    varchar(255) = NULL OUT

as

    set nocount on

    select @value = (
        select value
                from dbo.dtproperties
                where @id=objectid and @property=property
                )


GO
create procedure dbo.dt_getpropertiesbyid_vcs_u
    @id       int,
    @property varchar(64),
    @value    nvarchar(255) = NULL OUT

as

    -- This procedure should no longer be called;  dt_getpropertiesbyid_vcsshould be called instead.
	-- Calls are forwarded to dt_getpropertiesbyid_vcs to maintain backward compatibility.
	set nocount on
    exec dbo.dt_getpropertiesbyid_vcs
		@id,
		@property,
		@value output


GO
create proc dbo.dt_isundersourcecontrol
    @vchLoginName varchar(255) = '',
    @vchPassword  varchar(255) = '',
    @iWhoToo      int = 0 /* 0 => Just check project; 1 => get list of objs */

as

	set nocount on

	declare @iReturn int
	declare @iObjectId int
	select @iObjectId = 0

	declare @VSSGUID varchar(100)
	select @VSSGUID = 'SQLVersionControl.VCS_SQL'

	declare @iReturnValue int
	select @iReturnValue = 0

	declare @iStreamObjectId int
	select @iStreamObjectId   = 0

	declare @vchTempText varchar(255)

    declare @iPropertyObjectId int
    select @iPropertyObjectId = (select objectid from dbo.dtproperties where property = 'VCSProjectID')

    declare @vchProjectName   varchar(255)
    declare @vchSourceSafeINI varchar(255)
    declare @vchServerName    varchar(255)
    declare @vchDatabaseName  varchar(255)
    exec dbo.dt_getpropertiesbyid_vcs @iPropertyObjectId, 'VCSProject',       @vchProjectName   OUT
    exec dbo.dt_getpropertiesbyid_vcs @iPropertyObjectId, 'VCSSourceSafeINI', @vchSourceSafeINI OUT
    exec dbo.dt_getpropertiesbyid_vcs @iPropertyObjectId, 'VCSSQLServer',     @vchServerName    OUT
    exec dbo.dt_getpropertiesbyid_vcs @iPropertyObjectId, 'VCSSQLDatabase',   @vchDatabaseName  OUT

    if (@vchProjectName = '')	set @vchProjectName		= null
    if (@vchSourceSafeINI = '') set @vchSourceSafeINI	= null
    if (@vchServerName = '')	set @vchServerName		= null
    if (@vchDatabaseName = '')	set @vchDatabaseName	= null

    if (@vchProjectName is null) or (@vchSourceSafeINI is null) or (@vchServerName is null) or (@vchDatabaseName is null)
    begin
        RAISERROR('Not Under Source Control',16,-1)
        return
    end

    if @iWhoToo = 1
    begin

        /* Get List of Procs in the project */
        exec @iReturn = master.dbo.sp_OACreate @VSSGUID, @iObjectId OUT
        if @iReturn <> 0 GOTO E_OAError

        exec @iReturn = master.dbo.sp_OAMethod @iObjectId,
												'GetListOfObjects',
												NULL,
												@vchProjectName,
												@vchSourceSafeINI,
												@vchServerName,
												@vchDatabaseName,
												@vchLoginName,
												@vchPassword

        if @iReturn <> 0 GOTO E_OAError

        exec @iReturn = master.dbo.sp_OAGetProperty @iObjectId, 'GetStreamObject', @iStreamObjectId OUT

        if @iReturn <> 0 GOTO E_OAError

        create table #ObjectList (id int identity, vchObjectlist varchar(255))

        select @vchTempText = 'STUB'
        while @vchTempText is not null
        begin
            exec @iReturn = master.dbo.sp_OAMethod @iStreamObjectId, 'GetStream', @iReturnValue OUT, @vchTempText OUT
            if @iReturn <> 0 GOTO E_OAError

            if (@vchTempText = '') set @vchTempText = null
            if (@vchTempText is not null) insert into #ObjectList (vchObjectlist ) select @vchTempText
        end

        select vchObjectlist from #ObjectList order by id
    end

CleanUp:
    return

E_OAError:
    exec dbo.dt_displayoaerror @iObjectId, @iReturn
    goto CleanUp



GO
create proc dbo.dt_isundersourcecontrol_u
    @vchLoginName nvarchar(255) = '',
    @vchPassword  nvarchar(255) = '',
    @iWhoToo      int = 0 /* 0 => Just check project; 1 => get list of objs */

as
	-- This procedure should no longer be called;  dt_isundersourcecontrol should be called instead.
	-- Calls are forwarded to dt_isundersourcecontrol to maintain backward compatibility.
	set nocount on
	exec dbo.dt_isundersourcecontrol
		@vchLoginName,
		@vchPassword,
		@iWhoToo



GO
create procedure dbo.dt_removefromsourcecontrol

as

    set nocount on

    declare @iPropertyObjectId int
    select @iPropertyObjectId = (select objectid from dbo.dtproperties where property = 'VCSProjectID')

    exec dbo.dt_droppropertiesbyid @iPropertyObjectId, null

    /* -1 is returned by dt_droppopertiesbyid */
    if @@error <> 0 and @@error <> -1 return 1

    return 0



GO
/*
**	If the property already exists, reset the value; otherwise add property
**		id -- the id in sysobjects of the object
**		property -- the name of the property
**		value -- the text value of the property
**		lvalue -- the binary value of the property (image)
*/
create procedure dbo.dt_setpropertybyid
	@id int,
	@property varchar(64),
	@value varchar(255),
	@lvalue image
as
	set nocount on
	declare @uvalue nvarchar(255)
	set @uvalue = convert(nvarchar(255), @value)
	if exists (select * from dbo.dtproperties
			where objectid=@id and property=@property)
	begin
		--
		-- bump the version count for this row as we update it
		--
		update dbo.dtproperties set value=@value, uvalue=@uvalue, lvalue=@lvalue, version=version+1
			where objectid=@id and property=@property
	end
	else
	begin
		--
		-- version count is auto-set to 0 on initial insert
		--
		insert dbo.dtproperties (property, objectid, value, uvalue, lvalue)
			values (@property, @id, @value, @uvalue, @lvalue)
	end


GO
/*
**	If the property already exists, reset the value; otherwise add property
**		id -- the id in sysobjects of the object
**		property -- the name of the property
**		uvalue -- the text value of the property
**		lvalue -- the binary value of the property (image)
*/
create procedure dbo.dt_setpropertybyid_u
	@id int,
	@property varchar(64),
	@uvalue nvarchar(255),
	@lvalue image
as
	set nocount on
	--
	-- If we are writing the name property, find the ansi equivalent.
	-- If there is no lossless translation, generate an ansi name.
	--
	declare @avalue varchar(255)
	set @avalue = null
	if (@uvalue is not null)
	begin
		if (convert(nvarchar(255), convert(varchar(255), @uvalue)) = @uvalue)
		begin
			set @avalue = convert(varchar(255), @uvalue)
		end
		else
		begin
			if 'DtgSchemaNAME' = @property
			begin
				exec dbo.dt_generateansiname @avalue output
			end
		end
	end
	if exists (select * from dbo.dtproperties
			where objectid=@id and property=@property)
	begin
		--
		-- bump the version count for this row as we update it
		--
		update dbo.dtproperties set value=@avalue, uvalue=@uvalue, lvalue=@lvalue, version=version+1
			where objectid=@id and property=@property
	end
	else
	begin
		--
		-- version count is auto-set to 0 on initial insert
		--
		insert dbo.dtproperties (property, objectid, value, uvalue, lvalue)
			values (@property, @id, @avalue, @uvalue, @lvalue)
	end

GO
create proc dbo.dt_validateloginparams
    @vchLoginName  varchar(255),
    @vchPassword   varchar(255)
as

set nocount on

declare @iReturn int
declare @iObjectId int
select @iObjectId =0

declare @VSSGUID varchar(100)
select @VSSGUID = 'SQLVersionControl.VCS_SQL'

    declare @iPropertyObjectId int
    select @iPropertyObjectId = (select objectid from dbo.dtproperties where property = 'VCSProjectID')

    declare @vchSourceSafeINI varchar(255)
    exec dbo.dt_getpropertiesbyid_vcs @iPropertyObjectId, 'VCSSourceSafeINI', @vchSourceSafeINI OUT

    exec @iReturn = master.dbo.sp_OACreate @VSSGUID, @iObjectId OUT
    if @iReturn <> 0 GOTO E_OAError

    exec @iReturn = master.dbo.sp_OAMethod @iObjectId,
											'ValidateLoginParams',
											NULL,
											@sSourceSafeINI = @vchSourceSafeINI,
											@sLoginName = @vchLoginName,
											@sPassword = @vchPassword
    if @iReturn <> 0 GOTO E_OAError

CleanUp:
    return

E_OAError:
    exec dbo.dt_displayoaerror @iObjectId, @iReturn
    GOTO CleanUp



GO
create proc dbo.dt_validateloginparams_u
    @vchLoginName  nvarchar(255),
    @vchPassword   nvarchar(255)
as

	-- This procedure should no longer be called;  dt_validateloginparams should be called instead.
	-- Calls are forwarded to dt_validateloginparams to maintain backward compatibility.
	set nocount on
	exec dbo.dt_validateloginparams
		@vchLoginName,
		@vchPassword



GO
create proc dbo.dt_vcsenabled

as

set nocount on

declare @iObjectId int
select @iObjectId = 0

declare @VSSGUID varchar(100)
select @VSSGUID = 'SQLVersionControl.VCS_SQL'

    declare @iReturn int
    exec @iReturn = master.dbo.sp_OACreate @VSSGUID, @iObjectId OUT
    if @iReturn <> 0 raiserror('', 16, -1) /* Can't Load Helper DLLC */



GO
/*
**	This procedure returns the version number of the stored
**    procedures used by legacy versions of the Microsoft
**	Visual Database Tools.  Version is 7.0.00.
*/
create procedure dbo.dt_verstamp006
as
	select 7000

GO
/*
**	This procedure returns the version number of the stored
**    procedures used by the the Microsoft Visual Database Tools.
**	Version is 7.0.05.
*/
create procedure dbo.dt_verstamp007
as
	select 7005

GO
create proc dbo.dt_whocheckedout
        @chObjectType  char(4),
        @vchObjectName varchar(255),
        @vchLoginName  varchar(255),
        @vchPassword   varchar(255)

as

set nocount on

declare @iReturn int
declare @iObjectId int
select @iObjectId =0

declare @VSSGUID varchar(100)
select @VSSGUID = 'SQLVersionControl.VCS_SQL'

    declare @iPropertyObjectId int

    select @iPropertyObjectId = (select objectid from dbo.dtproperties where property = 'VCSProjectID')

    declare @vchProjectName   varchar(255)
    declare @vchSourceSafeINI varchar(255)
    declare @vchServerName    varchar(255)
    declare @vchDatabaseName  varchar(255)
    exec dbo.dt_getpropertiesbyid_vcs @iPropertyObjectId, 'VCSProject',       @vchProjectName   OUT
    exec dbo.dt_getpropertiesbyid_vcs @iPropertyObjectId, 'VCSSourceSafeINI', @vchSourceSafeINI OUT
    exec dbo.dt_getpropertiesbyid_vcs @iPropertyObjectId, 'VCSSQLServer',     @vchServerName    OUT
    exec dbo.dt_getpropertiesbyid_vcs @iPropertyObjectId, 'VCSSQLDatabase',   @vchDatabaseName  OUT

    if @chObjectType = 'PROC'
    begin
        exec @iReturn = master.dbo.sp_OACreate @VSSGUID, @iObjectId OUT

        if @iReturn <> 0 GOTO E_OAError

        declare @vchReturnValue varchar(255)
        select @vchReturnValue = ''

        exec @iReturn = master.dbo.sp_OAMethod @iObjectId,
												'WhoCheckedOut',
												@vchReturnValue OUT,
												@sProjectName = @vchProjectName,
												@sSourceSafeINI = @vchSourceSafeINI,
												@sObjectName = @vchObjectName,
												@sServerName = @vchServerName,
												@sDatabaseName = @vchDatabaseName,
												@sLoginName = @vchLoginName,
												@sPassword = @vchPassword

        if @iReturn <> 0 GOTO E_OAError

        select @vchReturnValue

    end

CleanUp:
    return

E_OAError:
    exec dbo.dt_displayoaerror @iObjectId, @iReturn
    GOTO CleanUp



GO
create proc dbo.dt_whocheckedout_u
        @chObjectType  char(4),
        @vchObjectName nvarchar(255),
        @vchLoginName  nvarchar(255),
        @vchPassword   nvarchar(255)

as

	-- This procedure should no longer be called;  dt_whocheckedout should be called instead.
	-- Calls are forwarded to dt_whocheckedout to maintain backward compatibility.
	set nocount on
	exec dbo.dt_whocheckedout
		@chObjectType,
		@vchObjectName,
		@vchLoginName,
		@vchPassword



GO

CREATE PROCEDURE [dbo].[FixedAssetsProc] @StartDate DateTime,@EndDate DateTime AS


SET NOCOUNT ON
SELECT  PnL.* INTO #tableA FROM (
--General entities
SELECT
	dbo.FullEntityTypeList.AccountType_DPA_ AS AccountType_DPA_,
	dbo.FullEntityTypeList.EntityTypeAccountType AS AccountType,
	'******' + CONVERT(VARCHAR(500),dbo.FullEntityTypeList.EntityType_DPA_) AS [Account Code],
	dbo.FullEntityTypeList.EntityTypeName AS AccountName,
	CASE
		WHEN SUM(dbo.EntityTransactionList.Balance) < 0 THEN 0 - SUM(dbo.EntityTransactionList.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.EntityTransactionList.Balance) >= 0 THEN SUM(dbo.EntityTransactionList.Balance)
		ELSE 0 END AS Credit,Cast(Floor(Cast(Min(EntityTransactionList.TransDate) as Float)) as DateTime) as TransDate
FROM         dbo.Entity INNER JOIN
                      dbo.FullEntityTypeList ON dbo.Entity.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_ INNER JOIN
                      dbo.EntityTransactionList ON dbo.Entity.Entity_DPA_ = dbo.EntityTransactionList.Entity_DPA_
WHERE 		dbo.FullEntityTypeList.AccountType_DPA_ =3
GROUP BY dbo.FullEntityTypeList.AccountType_DPA_,dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName, dbo.FullEntityTypeList.EntityType_DPA_
) PnL


--Clients
INSERT INTO #tableA
SELECT     dbo.FullEntityTypeList.AccountType_DPA_ AS AccountType_DPA_,
		dbo.FullEntityTypeList.EntityTypeAccountType AS AccountType,
		'******1' AS [Account Code],
		dbo.FullEntityTypeList.EntityTypeName AS AccountName,CASE WHEN SUM(StatementList.Balance) < 0
THEN 0 - SUM(StatementList.Balance) ELSE 0 END AS Debit, CASE WHEN SUM(StatementList.Balance) >= 0
THEN SUM(StatementList.Balance) ELSE 0 END AS Credit,TransDate From (Select StatementList.*,1 as EntityType_DPA_ from StatementList) StatementList INNER JOIN
		dbo.FullEntityTypeList ON StatementList.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_
WHERE 		dbo.FullEntityTypeList.AccountType_DPA_= 3
GROUP BY dbo.FullEntityTypeList.AccountType_DPA_,dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName,TransDate

--Agents
INSERT INTO #tableA
SELECT     dbo.FullEntityTypeList.AccountType_DPA_ AS AccountType_DPA_,
		dbo.FullEntityTypeList.EntityTypeAccountType AS AccountType,
		'******1' AS [Account Code],
		dbo.FullEntityTypeList.EntityTypeName AS AccountName,CASE WHEN SUM(AgentStatement.Balance) < 0
THEN 0 - SUM(AgentStatement.Balance) ELSE 0 END AS Debit, CASE WHEN SUM(AgentStatement.Balance) >= 0
THEN SUM(AgentStatement.Balance) ELSE 0 END AS Credit,TransDate From (Select AgentStatement.*,2 as EntityType_DPA_ from AgentStatement) AgentStatement INNER JOIN
		dbo.FullEntityTypeList ON AgentStatement.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_
WHERE 		dbo.FullEntityTypeList.AccountType_DPA_=3
GROUP BY dbo.FullEntityTypeList.AccountType_DPA_,dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName,TransDate

--Brokers
INSERT INTO #tableA
SELECT     dbo.FullEntityTypeList.AccountType_DPA_ AS AccountType_DPA_,
		dbo.FullEntityTypeList.EntityTypeAccountType AS AccountType,
		'******1' AS [Account Code],
		dbo.FullEntityTypeList.EntityTypeName AS AccountName,CASE WHEN SUM(BrokerStatement.Credit-BrokerStatement.Debit) < 0
THEN 0 - SUM(BrokerStatement.Credit-BrokerStatement.Debit) ELSE 0 END AS Debit, CASE WHEN SUM(BrokerStatement.Credit-BrokerStatement.Debit) >= 0
THEN SUM(BrokerStatement.Credit-BrokerStatement.Debit) ELSE 0 END AS Credit,TransDate From (Select BrokerStatement.*,3 as EntityType_DPA_ from BrokerStatement) BrokerStatement INNER JOIN
		dbo.FullEntityTypeList ON BrokerStatement.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_
WHERE 		dbo.FullEntityTypeList.AccountType_DPA_ =3
GROUP BY dbo.FullEntityTypeList.AccountType_DPA_,dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName,TransDate

--Account Managers
INSERT INTO #tableA
SELECT     dbo.FullEntityTypeList.AccountType_DPA_ AS AccountType_DPA_,
		dbo.FullEntityTypeList.EntityTypeAccountType AS AccountType,
		'******1' AS [Account Code],
		dbo.FullEntityTypeList.EntityTypeName AS AccountName,CASE WHEN SUM(OwnerStatement.Credit-OwnerStatement.Debit) < 0
THEN 0 - SUM(OwnerStatement.Credit-OwnerStatement.Debit) ELSE 0 END AS Debit, CASE WHEN SUM(OwnerStatement.Credit-OwnerStatement.Debit) >= 0
THEN SUM(OwnerStatement.Credit-OwnerStatement.Debit) ELSE 0 END AS Credit,TransDate From (Select OwnerStatement.*,7 as EntityType_DPA_ from OwnerStatement) OwnerStatement INNER JOIN
		dbo.FullEntityTypeList ON OwnerStatement.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_
WHERE 		dbo.FullEntityTypeList.AccountType_DPA_ =3
GROUP BY dbo.FullEntityTypeList.AccountType_DPA_,dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName,TransDate


--Levies and Broker commission
INSERT INTO #tableA
SELECT
	dbo.EntityList.AccountType_DPA_ AS AccountType_DPA_,
	dbo.EntityList.AccountTypeName AS AccountType,
	'------' + CONVERT(VARCHAR(500),dbo.EntityList.Entity_DPA_) AS [Account Code],
	dbo.EntityList.EntityName AS AccountName,
	CASE
		WHEN SUM(dbo.BrokerCommissionStatement.Balance) < 0 THEN 0 - SUM(dbo.BrokerCommissionStatement.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.BrokerCommissionStatement.Balance) >= 0 THEN SUM(dbo.BrokerCommissionStatement.Balance)
		ELSE 0 END AS Credit,BrokerCommissionStatement.TransDate
FROM         dbo.EntityList INNER JOIN
                      dbo.BrokerCommissionStatement ON dbo.EntityList.Entity_DPA_ = dbo.BrokerCommissionStatement.Entity_DPA_
WHERE 		dbo.EntityList.AccountType_DPA_ = 3
GROUP BY dbo.EntityList.AccountType_DPA_,dbo.EntityList.AccountTypeName,dbo.EntityList.EntityName, dbo.EntityList.Entity_DPA_,BrokerCommissionStatement.TransDate

INSERT INTO #tableA
SELECT
	dbo.EntityList.AccountType_DPA_ AS AccountType_DPA_,
	dbo.EntityList.AccountTypeName AS AccountType,
	'------' + CONVERT(VARCHAR(500),dbo.EntityList.Entity_DPA_) AS [Account Code],
	dbo.EntityList.EntityName AS AccountName,
	CASE
		WHEN SUM(dbo.CommissionStatement.Balance) < 0 THEN 0 - SUM(dbo.CommissionStatement.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.CommissionStatement.Balance) >= 0 THEN SUM(dbo.CommissionStatement.Balance)
		ELSE 0 END AS Credit,CommissionStatement.TransDate
FROM         dbo.EntityList INNER JOIN
                      dbo.CommissionStatement ON dbo.EntityList.Entity_DPA_ = dbo.CommissionStatement.Entity_DPA_
WHERE 		dbo.EntityList.AccountType_DPA_ = 3
GROUP BY dbo.EntityList.AccountType_DPA_,dbo.EntityList.AccountTypeName,dbo.EntityList.EntityName, dbo.EntityList.Entity_DPA_,CommissionStatement.TransDate

--Trading Account
INSERT INTO #tableA
SELECT
	dbo.EntityList.AccountType_DPA_ AS AccountType_DPA_,
	dbo.EntityList.AccountTypeName AS AccountType,
	'------' + CONVERT(VARCHAR(500),dbo.EntityList.Entity_DPA_) AS [Account Code],
	dbo.EntityList.EntityName AS AccountName,
	CASE
		WHEN SUM(dbo.CDSControlStatement.Balance) < 0 THEN 0 - SUM(dbo.CDSControlStatement.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.CDSControlStatement.Balance) >= 0 THEN SUM(dbo.CDSControlStatement.Balance)
		ELSE 0 END AS Credit,CDSControlStatement.TransDate
FROM         dbo.EntityList INNER JOIN
                      dbo.CDSControlStatement ON dbo.EntityList.Entity_DPA_ = dbo.CDSControlStatement.Client_DPA_
WHERE 		dbo.EntityList.AccountType_DPA_ = 3
GROUP BY dbo.EntityList.AccountType_DPA_,dbo.EntityList.AccountTypeName,dbo.EntityList.EntityName, dbo.EntityList.Entity_DPA_,CDSControlStatement.TransDate

--Nominal accounts
INSERT INTO #tableA
SELECT
	dbo.AccountList.AccountTypeLevel1_DPA_ AS AccountType_DPA_,
	dbo.AccountList.AccountTypeLevel1 AS AccountType,
	CONVERT(VARCHAR(500),dbo.AccountList.Account_DPA_) AS [Account Code],
	dbo.AccountList.AccountName,
	CASE
		WHEN SUM(dbo.DB_BankAccountStatement.Balance) < 0 THEN 0 - SUM(dbo.DB_BankAccountStatement.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.DB_BankAccountStatement.Balance) >= 0 THEN SUM(dbo.DB_BankAccountStatement.Balance)
		ELSE 0 END AS Credit,DB_BankAccountStatement.TransDate
FROM         dbo.DB_BankAccountStatement INNER JOIN
                      dbo.AccountList ON dbo.DB_BankAccountStatement.BankAccount_DPA_ = dbo.AccountList.Account_DPA_
WHERE 		dbo.AccountList.AccountTypeLevel1_DPA_ = 3
GROUP BY dbo.AccountList.AccountTypeLevel1_DPA_,dbo.AccountList.AccountTypeLevel1,dbo.AccountList.AccountName, dbo.AccountList.Account_DPA_,DB_BankAccountStatement.TransDate



SELECT  IDENTITY(int, 1,1)  AS SequenceID, #tableA.AccountType_DPA_,#tableA.AccountType,#tableA.[Account Code],#tableA.AccountName,Sum(#tableA.Credit - #tableA.Debit) AS Balance
	 INTO #tableB FROM #tableA Where #tableA.TransDate between @StartDate and @enddate Group by #tableA.AccountType_DPA_,#tableA.AccountType,#tableA.[Account Code],#tableA.AccountName ORDER BY AccountType_DPA_

DROP TABLE #tableA

select T.SequenceID,T.AccountType_DPA_,T.AccountType,T.[Account Code],T.AccountName,T.Balance,
       case when SequenceID = (select top 1 SequenceID from #tableB
                           where AccountType_DPA_
                              = T.AccountType_DPA_
                          order by SequenceID desc)
            then (select CONVERT(CHAR(20),sum(Balance))
                     from #tableB
                     where SequenceID <= T.SequenceID
                        and AccountType_DPA_
                              = T.AccountType_DPA_)
            else ' ' end as 'SubTotal',
       case when SequenceID = (select top 1 SequenceID from #tableB
                           order by SequenceID desc)
            then (select CONVERT(CHAR(20),sum(Balance))
                      from #tableB)
             else ' ' end as 'GrandTotal'
from #tableB T
  order by SequenceID



DROP TABLE #tableB


GO


CREATE  Proc [dbo].[GenerateClientBalanceToDate] @date Varchar(20)
as
Declare @TransDate datetime


set @TransDate = convert(datetime,@date)

Begin
truncate table ClientBalancesTEMP
Insert Into ClientBalancesTEMP(Client_DPA_,CurrentBal)
SELECT  ClientsStatement.Client_DPA_,
SUM(ISNULL(ClientsStatement.Credit-ClientsStatement.Debit, 0)) AS CurrentBal
FROM                  (
SELECT     *
FROM         statementlist
WHERE     (CAST(FLOOR(CAST(TransDate AS float)) AS datetime) <= CAST(FLOOR(CAST(@TransDate AS float)) AS datetime) )
) ClientsStatement
GROUP BY ClientsStatement.Client_DPA_
End



--This is the query that will pick this query

SELECT     Client.Client_DPA_ AS Code, Client.ClientName AS Name,
                      Client.ClientIDPass AS [ID/Passport], CASE Iscustodian WHEN 1 THEN 'Custodian' ELSE 'Non - Custodian' END AS Status,
                      ROUND(ISNULL(ClientBalancesTEMP.CurrentBal, 0), 2) AS Balance,@TransDate as BalanceUPTO
FROM         Client LEFT OUTER JOIN
                      ClientBalancesTEMP ON Client.Client_DPA_ = ClientBalancesTEMP.client_DPA_
WHERE     (Client.Deleted = 0) and client.iscustodian =0











--GenerateClientBalanceToDate '31-OCT-2009'
GO

CREATE PROCEDURE [dbo].[JournaEntriesProc] AS




SELECT CASE Entry
	WHEN '' THEN Journal_DPA_
	ELSE Entry END AS Entry, Narrative, Debit, Credit, AccountName, ControlAccount FROM

(SELECT CAST(Count(*) AS NVARCHAR(500)) AS Entry,
	a.JournalNarrative AS Narrative, CAST(a.JournalEntryDebit AS NVARCHAR(500)) AS Debit,
	CAST(a.JournalEntryCredit AS NVARCHAR(500)) AS Credit, a.JournalEntryAccount AS AccountName,
	a.ControlAccount, a.JournalDate AS TransDate, a.Journal_DPA_
	FROM

		(SELECT JournalEntry_DPA_, JournalNarrative, JournalEntryDebit, JournalEntryCredit, JournalEntryAccount,
			'(' + EntityTypeName + ')' As ControlAccount, Journal_DPA_, JournalDate
			 FROM dbo.JournalList
				INNER JOIN dbo.FullEntityTypeList ON
				dbo.JournalList.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_) a
		CROSS JOIN

			(SELECT JournalEntry_DPA_, JournalNarrative, JournalEntryDebit, JournalEntryCredit, JournalEntryAccount,
			'(' + EntityTypeName + ')' As ControlAccount, Journal_DPA_, JournalDate
			 FROM dbo.JournalList
				INNER JOIN dbo.FullEntityTypeList ON
				dbo.JournalList.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_) b
		WHERE a.Journal_DPA_ = b.Journal_DPA_ AND a.JournalEntry_DPA_ >= b.JournalEntry_DPA_
	GROUP BY a.Journal_DPA_, a.JournalNarrative, a.JournalEntryDebit,
		a.JournalEntryCredit, a.JournalEntryAccount, a.ControlAccount, a.JournalDate


UNION ALL

SELECT '' AS Entry, CAST(JournalDate AS NVARCHAR(500)) AS Narrative,'' AS Debit,'' AS Credit,'' As AccountName,'' As ControlAccount,
			JournalDate AS TransDate, Journal_DPA_

			FROM dbo.JournalList
			GROUP BY JournalDate, Journal_DPA_) allData
ORDER BY Journal_DPA_, TransDate, CONVERT(int, Entry), ControlAccount


GO

CREATE PROCEDURE [dbo].[LongTermLiabilitiesProc] @StartDate DateTime,@EndDate DateTime AS


SET NOCOUNT ON
SELECT  PnL.* INTO #tableA FROM (
--General entities
SELECT
	dbo.FullEntityTypeList.AccountType_DPA_ AS AccountType_DPA_,
	dbo.FullEntityTypeList.EntityTypeAccountType AS AccountType,
	'******' + CONVERT(VARCHAR(500),dbo.FullEntityTypeList.EntityType_DPA_) AS [Account Code],
	dbo.FullEntityTypeList.EntityTypeName AS AccountName,
	CASE
		WHEN SUM(dbo.EntityTransactionList.Balance) < 0 THEN 0 - SUM(dbo.EntityTransactionList.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.EntityTransactionList.Balance) >= 0 THEN SUM(dbo.EntityTransactionList.Balance)
		ELSE 0 END AS Credit,Cast(Floor(Cast(Min(EntityTransactionList.TransDate) as Float)) as DateTime) as TransDate
FROM         dbo.Entity INNER JOIN
                      dbo.FullEntityTypeList ON dbo.Entity.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_ INNER JOIN
                      dbo.EntityTransactionList ON dbo.Entity.Entity_DPA_ = dbo.EntityTransactionList.Entity_DPA_
WHERE 		dbo.FullEntityTypeList.AccountType_DPA_ =6
GROUP BY dbo.FullEntityTypeList.AccountType_DPA_,dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName, dbo.FullEntityTypeList.EntityType_DPA_
) PnL


--Clients
INSERT INTO #tableA
SELECT     dbo.FullEntityTypeList.AccountType_DPA_ AS AccountType_DPA_,
		dbo.FullEntityTypeList.EntityTypeAccountType AS AccountType,
		'******1' AS [Account Code],
		dbo.FullEntityTypeList.EntityTypeName AS AccountName,CASE WHEN SUM(StatementList.Balance) < 0
THEN 0 - SUM(StatementList.Balance) ELSE 0 END AS Debit, CASE WHEN SUM(StatementList.Balance) >= 0
THEN SUM(StatementList.Balance) ELSE 0 END AS Credit,TransDate From (Select StatementList.*,1 as EntityType_DPA_ from StatementList) StatementList INNER JOIN
		dbo.FullEntityTypeList ON StatementList.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_
WHERE 		dbo.FullEntityTypeList.AccountType_DPA_= 6
GROUP BY dbo.FullEntityTypeList.AccountType_DPA_,dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName,TransDate

--Agents
INSERT INTO #tableA
SELECT     dbo.FullEntityTypeList.AccountType_DPA_ AS AccountType_DPA_,
		dbo.FullEntityTypeList.EntityTypeAccountType AS AccountType,
		'******1' AS [Account Code],
		dbo.FullEntityTypeList.EntityTypeName AS AccountName,CASE WHEN SUM(AgentStatement.Balance) < 0
THEN 0 - SUM(AgentStatement.Balance) ELSE 0 END AS Debit, CASE WHEN SUM(AgentStatement.Balance) >= 0
THEN SUM(AgentStatement.Balance) ELSE 0 END AS Credit,TransDate From (Select AgentStatement.*,2 as EntityType_DPA_ from AgentStatement) AgentStatement INNER JOIN
		dbo.FullEntityTypeList ON AgentStatement.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_
WHERE 		dbo.FullEntityTypeList.AccountType_DPA_=6
GROUP BY dbo.FullEntityTypeList.AccountType_DPA_,dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName,TransDate

--Brokers
INSERT INTO #tableA
SELECT     dbo.FullEntityTypeList.AccountType_DPA_ AS AccountType_DPA_,
		dbo.FullEntityTypeList.EntityTypeAccountType AS AccountType,
		'******1' AS [Account Code],
		dbo.FullEntityTypeList.EntityTypeName AS AccountName,CASE WHEN SUM(BrokerStatement.Credit-BrokerStatement.Debit) < 0
THEN 0 - SUM(BrokerStatement.Credit-BrokerStatement.Debit) ELSE 0 END AS Debit, CASE WHEN SUM(BrokerStatement.Credit-BrokerStatement.Debit) >= 0
THEN SUM(BrokerStatement.Credit-BrokerStatement.Debit) ELSE 0 END AS Credit,TransDate From (Select BrokerStatement.*,3 as EntityType_DPA_ from BrokerStatement) BrokerStatement INNER JOIN
		dbo.FullEntityTypeList ON BrokerStatement.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_
WHERE 		dbo.FullEntityTypeList.AccountType_DPA_ =6
GROUP BY dbo.FullEntityTypeList.AccountType_DPA_,dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName,TransDate

--Account Managers
INSERT INTO #tableA
SELECT     dbo.FullEntityTypeList.AccountType_DPA_ AS AccountType_DPA_,
		dbo.FullEntityTypeList.EntityTypeAccountType AS AccountType,
		'******1' AS [Account Code],
		dbo.FullEntityTypeList.EntityTypeName AS AccountName,CASE WHEN SUM(OwnerStatement.Credit-OwnerStatement.Debit) < 0
THEN 0 - SUM(OwnerStatement.Credit-OwnerStatement.Debit) ELSE 0 END AS Debit, CASE WHEN SUM(OwnerStatement.Credit-OwnerStatement.Debit) >= 0
THEN SUM(OwnerStatement.Credit-OwnerStatement.Debit) ELSE 0 END AS Credit,TransDate From (Select OwnerStatement.*,7 as EntityType_DPA_ from OwnerStatement) OwnerStatement INNER JOIN
		dbo.FullEntityTypeList ON OwnerStatement.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_
WHERE 		dbo.FullEntityTypeList.AccountType_DPA_ =6
GROUP BY dbo.FullEntityTypeList.AccountType_DPA_,dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName,TransDate


--Levies and Broker commission
INSERT INTO #tableA
SELECT
	dbo.EntityList.AccountType_DPA_ AS AccountType_DPA_,
	dbo.EntityList.AccountTypeName AS AccountType,
	'------' + CONVERT(VARCHAR(500),dbo.EntityList.Entity_DPA_) AS [Account Code],
	dbo.EntityList.EntityName AS AccountName,
	CASE
		WHEN SUM(dbo.BrokerCommissionStatement.Balance) < 0 THEN 0 - SUM(dbo.BrokerCommissionStatement.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.BrokerCommissionStatement.Balance) >= 0 THEN SUM(dbo.BrokerCommissionStatement.Balance)
		ELSE 0 END AS Credit,BrokerCommissionStatement.TransDate
FROM         dbo.EntityList INNER JOIN
                      dbo.BrokerCommissionStatement ON dbo.EntityList.Entity_DPA_ = dbo.BrokerCommissionStatement.Entity_DPA_
WHERE 		dbo.EntityList.AccountType_DPA_ = 6
GROUP BY dbo.EntityList.AccountType_DPA_,dbo.EntityList.AccountTypeName,dbo.EntityList.EntityName, dbo.EntityList.Entity_DPA_,BrokerCommissionStatement.TransDate

INSERT INTO #tableA
SELECT
	dbo.EntityList.AccountType_DPA_ AS AccountType_DPA_,
	dbo.EntityList.AccountTypeName AS AccountType,
	'------' + CONVERT(VARCHAR(500),dbo.EntityList.Entity_DPA_) AS [Account Code],
	dbo.EntityList.EntityName AS AccountName,
	CASE
		WHEN SUM(dbo.CommissionStatement.Balance) < 0 THEN 0 - SUM(dbo.CommissionStatement.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.CommissionStatement.Balance) >= 0 THEN SUM(dbo.CommissionStatement.Balance)
		ELSE 0 END AS Credit,CommissionStatement.TransDate
FROM         dbo.EntityList INNER JOIN
                      dbo.CommissionStatement ON dbo.EntityList.Entity_DPA_ = dbo.CommissionStatement.Entity_DPA_
WHERE 		dbo.EntityList.AccountType_DPA_ = 6
GROUP BY dbo.EntityList.AccountType_DPA_,dbo.EntityList.AccountTypeName,dbo.EntityList.EntityName, dbo.EntityList.Entity_DPA_,CommissionStatement.TransDate

--Trading Account
INSERT INTO #tableA
SELECT
	dbo.EntityList.AccountType_DPA_ AS AccountType_DPA_,
	dbo.EntityList.AccountTypeName AS AccountType,
	'------' + CONVERT(VARCHAR(500),dbo.EntityList.Entity_DPA_) AS [Account Code],
	dbo.EntityList.EntityName AS AccountName,
	CASE
		WHEN SUM(dbo.CDSControlStatement.Balance) < 0 THEN 0 - SUM(dbo.CDSControlStatement.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.CDSControlStatement.Balance) >= 0 THEN SUM(dbo.CDSControlStatement.Balance)
		ELSE 0 END AS Credit,CDSControlStatement.TransDate
FROM         dbo.EntityList INNER JOIN
                      dbo.CDSControlStatement ON dbo.EntityList.Entity_DPA_ = dbo.CDSControlStatement.Client_DPA_
WHERE 		dbo.EntityList.AccountType_DPA_ = 6
GROUP BY dbo.EntityList.AccountType_DPA_,dbo.EntityList.AccountTypeName,dbo.EntityList.EntityName, dbo.EntityList.Entity_DPA_,CDSControlStatement.TransDate

--Nominal accounts
INSERT INTO #tableA
SELECT
	dbo.AccountList.AccountTypeLevel1_DPA_ AS AccountType_DPA_,
	dbo.AccountList.AccountTypeLevel1 AS AccountType,
	CONVERT(VARCHAR(500),dbo.AccountList.Account_DPA_) AS [Account Code],
	dbo.AccountList.AccountName,
	CASE
		WHEN SUM(dbo.DB_BankAccountStatement.Balance) < 0 THEN 0 - SUM(dbo.DB_BankAccountStatement.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.DB_BankAccountStatement.Balance) >= 0 THEN SUM(dbo.DB_BankAccountStatement.Balance)
		ELSE 0 END AS Credit,DB_BankAccountStatement.TransDate
FROM         dbo.DB_BankAccountStatement INNER JOIN
                      dbo.AccountList ON dbo.DB_BankAccountStatement.BankAccount_DPA_ = dbo.AccountList.Account_DPA_
WHERE 		dbo.AccountList.AccountTypeLevel1_DPA_ = 6
GROUP BY dbo.AccountList.AccountTypeLevel1_DPA_,dbo.AccountList.AccountTypeLevel1,dbo.AccountList.AccountName, dbo.AccountList.Account_DPA_,DB_BankAccountStatement.TransDate



SELECT  IDENTITY(int, 1,1)  AS SequenceID, #tableA.AccountType_DPA_,#tableA.AccountType,#tableA.[Account Code],#tableA.AccountName,Sum(#tableA.Credit - #tableA.Debit) AS Balance
	 INTO #tableB FROM #tableA Where #tableA.TransDate between @StartDate and @enddate Group by #tableA.AccountType_DPA_,#tableA.AccountType,#tableA.[Account Code],#tableA.AccountName ORDER BY AccountType_DPA_

DROP TABLE #tableA

select T.SequenceID,T.AccountType_DPA_,T.AccountType,T.[Account Code],T.AccountName,T.Balance,
       case when SequenceID = (select top 1 SequenceID from #tableB
                           where AccountType_DPA_
                              = T.AccountType_DPA_
                          order by SequenceID desc)
            then (select CONVERT(CHAR(20),sum(Balance))
                     from #tableB
                     where SequenceID <= T.SequenceID
                        and AccountType_DPA_
                              = T.AccountType_DPA_)
            else ' ' end as 'SubTotal',
       case when SequenceID = (select top 1 SequenceID from #tableB
                           order by SequenceID desc)
            then (select CONVERT(CHAR(20),sum(Balance))
                      from #tableB)
             else ' ' end as 'GrandTotal'
from #tableB T
  order by SequenceID



DROP TABLE #tableB


GO

CREATE PROCEDURE [dbo].[PnLProc] @StartDate DateTime,@EndDate DateTime AS

SET NOCOUNT ON
SELECT  PnL.* INTO #tableA FROM (
--General entities
SELECT
	dbo.FullEntityTypeList.AccountType_DPA_ AS AccountType_DPA_,
	dbo.FullEntityTypeList.EntityTypeAccountType AS AccountType,
	'******' + CONVERT(VARCHAR(500),dbo.FullEntityTypeList.EntityType_DPA_) AS [Account Code],
	dbo.FullEntityTypeList.EntityTypeName AS AccountName,
	CASE
		WHEN SUM(dbo.EntityTransactionList.Balance) < 0 THEN 0 - SUM(dbo.EntityTransactionList.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.EntityTransactionList.Balance) >= 0 THEN SUM(dbo.EntityTransactionList.Balance)
		ELSE 0 END AS Credit,Cast(Floor(Cast(Min(EntityTransactionList.TransDate) as Float)) as DateTime) as TransDate
FROM         dbo.Entity INNER JOIN
                      dbo.FullEntityTypeList ON dbo.Entity.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_ INNER JOIN
                      dbo.EntityTransactionList ON dbo.Entity.Entity_DPA_ = dbo.EntityTransactionList.Entity_DPA_
WHERE 		dbo.FullEntityTypeList.AccountType_DPA_ IN (1,2) and IsOpeningBalance <> 1
GROUP BY dbo.FullEntityTypeList.AccountType_DPA_,dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName, dbo.FullEntityTypeList.EntityType_DPA_
) PnL


--Clients
INSERT INTO #tableA
SELECT     dbo.FullEntityTypeList.AccountType_DPA_ AS AccountType_DPA_,
		dbo.FullEntityTypeList.EntityTypeAccountType AS AccountType,
		'******1' AS [Account Code],
		dbo.FullEntityTypeList.EntityTypeName AS AccountName,CASE WHEN SUM(StatementList.Balance) < 0
THEN 0 - SUM(StatementList.Balance) ELSE 0 END AS Debit, CASE WHEN SUM(StatementList.Balance) >= 0
THEN SUM(StatementList.Balance) ELSE 0 END AS Credit,TransDate From (Select StatementList.*,1 as EntityType_DPA_ from StatementList) StatementList INNER JOIN
		dbo.FullEntityTypeList ON StatementList.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_
WHERE 		dbo.FullEntityTypeList.AccountType_DPA_ IN (1,2) and IsOpeningBalance <> 1
GROUP BY dbo.FullEntityTypeList.AccountType_DPA_,dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName,TransDate

--Agents
INSERT INTO #tableA
SELECT     dbo.FullEntityTypeList.AccountType_DPA_ AS AccountType_DPA_,
		dbo.FullEntityTypeList.EntityTypeAccountType AS AccountType,
		'******1' AS [Account Code],
		dbo.FullEntityTypeList.EntityTypeName AS AccountName,CASE WHEN SUM(AgentStatement.Balance) < 0
THEN 0 - SUM(AgentStatement.Balance) ELSE 0 END AS Debit, CASE WHEN SUM(AgentStatement.Balance) >= 0
THEN SUM(AgentStatement.Balance) ELSE 0 END AS Credit,TransDate From (Select AgentStatement.*,2 as EntityType_DPA_ from AgentStatement) AgentStatement INNER JOIN
		dbo.FullEntityTypeList ON AgentStatement.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_
WHERE 		dbo.FullEntityTypeList.AccountType_DPA_ IN (1,2) and IsOpeningBalance <> 1
GROUP BY dbo.FullEntityTypeList.AccountType_DPA_,dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName,TransDate

--Brokers
INSERT INTO #tableA
SELECT     dbo.FullEntityTypeList.AccountType_DPA_ AS AccountType_DPA_,
		dbo.FullEntityTypeList.EntityTypeAccountType AS AccountType,
		'******1' AS [Account Code],
		dbo.FullEntityTypeList.EntityTypeName AS AccountName,CASE WHEN SUM(BrokerStatement.Credit-BrokerStatement.Debit) < 0
THEN 0 - SUM(BrokerStatement.Credit-BrokerStatement.Debit) ELSE 0 END AS Debit, CASE WHEN SUM(BrokerStatement.Credit-BrokerStatement.Debit) >= 0
THEN SUM(BrokerStatement.Credit-BrokerStatement.Debit) ELSE 0 END AS Credit,TransDate From (Select BrokerStatement.*,3 as EntityType_DPA_ from BrokerStatement) BrokerStatement INNER JOIN
		dbo.FullEntityTypeList ON BrokerStatement.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_
WHERE 		dbo.FullEntityTypeList.AccountType_DPA_ IN (1,2) and IsOpeningBalance <> 1
GROUP BY dbo.FullEntityTypeList.AccountType_DPA_,dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName,TransDate

--Account Managers
INSERT INTO #tableA
SELECT     dbo.FullEntityTypeList.AccountType_DPA_ AS AccountType_DPA_,
		dbo.FullEntityTypeList.EntityTypeAccountType AS AccountType,
		'******1' AS [Account Code],
		dbo.FullEntityTypeList.EntityTypeName AS AccountName,CASE WHEN SUM(OwnerStatement.Credit-OwnerStatement.Debit) < 0
THEN 0 - SUM(OwnerStatement.Credit-OwnerStatement.Debit) ELSE 0 END AS Debit, CASE WHEN SUM(OwnerStatement.Credit-OwnerStatement.Debit) >= 0
THEN SUM(OwnerStatement.Credit-OwnerStatement.Debit) ELSE 0 END AS Credit,TransDate From (Select OwnerStatement.*,7 as EntityType_DPA_ from OwnerStatement) OwnerStatement INNER JOIN
		dbo.FullEntityTypeList ON OwnerStatement.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_
WHERE 		dbo.FullEntityTypeList.AccountType_DPA_ IN (1,2) and IsOpeningBalance <> 1
GROUP BY dbo.FullEntityTypeList.AccountType_DPA_,dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName,TransDate


--Levies and Broker commission
INSERT INTO #tableA
SELECT
	dbo.EntityList.AccountType_DPA_ AS AccountType_DPA_,
	dbo.EntityList.AccountTypeName AS AccountType,
	'------' + CONVERT(VARCHAR(500),dbo.EntityList.Entity_DPA_) AS [Account Code],
	dbo.EntityList.EntityName AS AccountName,
	CASE
		WHEN SUM(dbo.BrokerCommissionStatement.Balance) < 0 THEN 0 - SUM(dbo.BrokerCommissionStatement.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.BrokerCommissionStatement.Balance) >= 0 THEN SUM(dbo.BrokerCommissionStatement.Balance)
		ELSE 0 END AS Credit,BrokerCommissionStatement.TransDate
FROM         dbo.EntityList INNER JOIN
                      dbo.BrokerCommissionStatement ON dbo.EntityList.Entity_DPA_ = dbo.BrokerCommissionStatement.Entity_DPA_
WHERE 		dbo.EntityList.AccountType_DPA_ IN (1,2) and IsOpeningBalance <> 1
GROUP BY dbo.EntityList.AccountType_DPA_,dbo.EntityList.AccountTypeName,dbo.EntityList.EntityName, dbo.EntityList.Entity_DPA_,BrokerCommissionStatement.TransDate

INSERT INTO #tableA
SELECT
	dbo.EntityList.AccountType_DPA_ AS AccountType_DPA_,
	dbo.EntityList.AccountTypeName AS AccountType,
	'------' + CONVERT(VARCHAR(500),dbo.EntityList.Entity_DPA_) AS [Account Code],
	dbo.EntityList.EntityName AS AccountName,
	CASE
		WHEN SUM(dbo.CommissionStatement.Balance) < 0 THEN 0 - SUM(dbo.CommissionStatement.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.CommissionStatement.Balance) >= 0 THEN SUM(dbo.CommissionStatement.Balance)
		ELSE 0 END AS Credit,CommissionStatement.TransDate
FROM         dbo.EntityList INNER JOIN
                      dbo.CommissionStatement ON dbo.EntityList.Entity_DPA_ = dbo.CommissionStatement.Entity_DPA_
WHERE 		dbo.EntityList.AccountType_DPA_ IN (1,2) and IsOpeningBalance <> 1
GROUP BY dbo.EntityList.AccountType_DPA_,dbo.EntityList.AccountTypeName,dbo.EntityList.EntityName, dbo.EntityList.Entity_DPA_,CommissionStatement.TransDate

--Trading Account
INSERT INTO #tableA
SELECT
	dbo.EntityList.AccountType_DPA_ AS AccountType_DPA_,
	dbo.EntityList.AccountTypeName AS AccountType,
	'------' + CONVERT(VARCHAR(500),dbo.EntityList.Entity_DPA_) AS [Account Code],
	dbo.EntityList.EntityName AS AccountName,
	CASE
		WHEN SUM(dbo.CDSControlStatement.Balance) < 0 THEN 0 - SUM(dbo.CDSControlStatement.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.CDSControlStatement.Balance) >= 0 THEN SUM(dbo.CDSControlStatement.Balance)
		ELSE 0 END AS Credit,CDSControlStatement.TransDate
FROM         dbo.EntityList INNER JOIN
                      dbo.CDSControlStatement ON dbo.EntityList.Entity_DPA_ = dbo.CDSControlStatement.Client_DPA_
WHERE 		dbo.EntityList.AccountType_DPA_ IN (1,2) and IsOpeningBalance <> 1
GROUP BY dbo.EntityList.AccountType_DPA_,dbo.EntityList.AccountTypeName,dbo.EntityList.EntityName, dbo.EntityList.Entity_DPA_,CDSControlStatement.TransDate

--Nominal accounts
INSERT INTO #tableA
SELECT
	dbo.AccountList.AccountTypeLevel1_DPA_ AS AccountType_DPA_,
	dbo.AccountList.AccountTypeLevel1 AS AccountType,
	CONVERT(VARCHAR(500),dbo.AccountList.Account_DPA_) AS [Account Code],
	dbo.AccountList.AccountName,
	CASE
		WHEN SUM(dbo.DB_BankAccountStatement.Balance) < 0 THEN 0 - SUM(dbo.DB_BankAccountStatement.Balance)
		ELSE 0 END AS Debit,
	CASE
		WHEN SUM(dbo.DB_BankAccountStatement.Balance) >= 0 THEN SUM(dbo.DB_BankAccountStatement.Balance)
		ELSE 0 END AS Credit,DB_BankAccountStatement.TransDate
FROM         dbo.DB_BankAccountStatement INNER JOIN
                      dbo.AccountList ON dbo.DB_BankAccountStatement.BankAccount_DPA_ = dbo.AccountList.Account_DPA_
WHERE 		dbo.AccountList.AccountTypeLevel1_DPA_ IN (1,2) and IsOpeningBalance <> 1
GROUP BY dbo.AccountList.AccountTypeLevel1_DPA_,dbo.AccountList.AccountTypeLevel1,dbo.AccountList.AccountName, dbo.AccountList.Account_DPA_,DB_BankAccountStatement.TransDate



SELECT  IDENTITY(int, 1,1)  AS SequenceID, #tableA.AccountType_DPA_,#tableA.AccountType,#tableA.[Account Code],#tableA.AccountName,Sum(#tableA.Credit - #tableA.Debit) AS Balance
	 INTO #tableB FROM #tableA Where #tableA.TransDate between @StartDate and @enddate Group by #tableA.AccountType_DPA_,#tableA.AccountType,#tableA.[Account Code],#tableA.AccountName ORDER BY AccountType_DPA_

DROP TABLE #tableA

select T.SequenceID,T.AccountType_DPA_,T.AccountType,T.[Account Code],T.AccountName,T.Balance,
       case when SequenceID = (select top 1 SequenceID from #tableB
                           where AccountType_DPA_
                              = T.AccountType_DPA_
                          order by SequenceID desc)
            then (select CONVERT(CHAR(20),sum(Balance))
                     from #tableB
                     where SequenceID <= T.SequenceID
                        and AccountType_DPA_
                              = T.AccountType_DPA_)
            else ' ' end as 'SubTotal',
       case when SequenceID = (select top 1 SequenceID from #tableB
                           order by SequenceID desc)
            then (select CONVERT(CHAR(20),sum(Balance))
                      from #tableB)
             else ' ' end as 'GrandTotal'
from #tableB T
  order by SequenceID



DROP TABLE #tableB


GO
CREATE proc qry_GetLevies @startDate smalldatetime, @enddate smalldatetime

as

--qry_GetLevies '1-May-2009','31-May-2009'

drop table temp_MarchContractSchedule
SELECT     CROSSTABTBL.*, INNERTBL.AgentName, INNERTBL.IsCustodian, INNERTBL.ContractSettlementDate, INNERTBL.LotTDate, INNERTBL.Contract,
                      INNERTBL.ContractNumber, INNERTBL.OrderTypeSale, INNERTBL.SecurityCode, INNERTBL.OrderSecType_DPA_, INNERTBL.BrokerCode,
                      INNERTBL.LotSlipNo, INNERTBL.LotPrice, INNERTBL.LotQty, INNERTBL.ClientName, INNERTBL.Client_DPA_, INNERTBL.LotGrossAmount
INTO            temp_MarchContractSchedule
FROM         (SELECT     dbo.OrdDetailList.AgentName, dbo.OrdDetailList.IsCustodian, dbo.Lot.ContractSettlementDate, CONVERT(SMALLDATETIME,
                                              CAST(dbo.Lot.LotTDate AS CHAR(12))) AS LotTDate, dbo.Lot.Contract_DPA_ AS Contract, dbo.Lot.ContractNumber AS ContractNumber,
                                              dbo.OrdDetailList.OrderTypeSale, dbo.Security.SecurityCode, OrdDetailList.OrderSecType_DPA_, dbo.Broker.BrokerCode,
                                              dbo.Lot.LotSlipNo, dbo.Lot.LotPrice, dbo.Lot.LotQty, dbo.OrdDetailList.OrdDetailClient AS ClientName,
                                              dbo.OrdDetailList.Client_DPA_ AS Client_DPA_, dbo.Lot.Contract_DPA_, dbo.Lot.LotGrossAmount
                       FROM          dbo.OrdDetailList INNER JOIN
                                              dbo.Lot ON dbo.OrdDetailList.OrdDetail_DPA_ = dbo.Lot.OrdDetail_DPA_ INNER JOIN
                                              dbo.Broker ON dbo.Lot.Broker_DPA_ = dbo.Broker.Broker_DPA_ LEFT OUTER JOIN
                                              dbo.Security ON dbo.OrdDetailList.Security_DPA_ = dbo.Security.Security_DPA_
                       WHERE      lot.deleted <> 1 and cast(floor(cast(Lot.lottdate as float)) as datetime) Between cast(floor(cast(@startDate as float)) as datetime)
			and cast(floor(cast(@endDate as float)) as datetime) ) INNERTBL INNER JOIN
                          (SELECT     CASE WHEN (GROUPING([Contract_DPA_]) = 1) THEN 'Total' ELSE CAST(+ [Contract_DPA_] AS NVARCHAR(255)) END AS [Contract_DPA_],
                                                   MAX(CASE CAST([LevyShortName] AS VARCHAR(255)) WHEN 'Agent' THEN LevyAmount ELSE 0 END) AS [Agent],
                                                   MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN 'Basic' THEN LevyAmount ELSE 0 END) AS [Basic],
                                                   MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN 'Commission' THEN LevyAmount ELSE 0 END) AS [Commission],
                                                   MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN 'MSEComm' THEN LevyAmount ELSE 0 END) AS [MSEComm],
                                                   MAX(CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN 'VAT' THEN LevyAmount ELSE 0 END) AS [VAT], MAX(LevyAmount)
                                                   AS Total
                            FROM          LevyContracts
                            GROUP BY Contract_DPA_ WITH CUBE) CROSSTABTBL ON CONVERT(NVARCHAR(1000), INNERTBL.Contract_DPA_) = CONVERT(NVARCHAR(1000),
                      CROSSTABTBL.Contract_DPA_)



SELECT     Client_DPA_ AS Code, ClientName AS Name, AgentName, CASE IsCustodian WHEN 1 THEN 'Y' ELSE 'N' END AS Custodian,
                      CAST(LotTDate AS datetime) AS Traded, CAST(ContractSettlementDate AS datetime) AS Settle,
                      CASE OrderTypeSale WHEN 1 THEN 'Sale' ELSE 'Purchase' END AS Side, SecurityCode AS Security, BrokerCode AS Broker, ContractNumber,
                      LotSlipNo AS Slip, LotPrice AS Price, LotQty AS Qty, LotGrossAmount AS Gross, Agent, Basic, Commission, MSEComm, VAT, Total
FROM         temp_MarchContractSchedule
WHERE     (CAST(FLOOR(CAST(LotTDate AS float)) AS datetime) BETWEEN CAST(FLOOR(CAST(CAST(@startDate AS datetime) AS float)) AS datetime) AND
                      CAST(FLOOR(CAST(CAST(@endDate AS datetime) AS float)) AS datetime))

GO

CREATE PROCEDURE ShortName
AS
UPDATE LevyContract
SET       ShortName = CASE WHEN LevyShortName LIKE '%Trans%' THEN 'a' + LevyShortName WHEN LevyShortName LIKE '%Stamp%' THEN 'b' + LevyShortName
                WHEN LevyShortName LIKE '%CMA%' THEN 'c' + LevyShortName WHEN LevyShortName LIKE '%Comp%' THEN 'd' + LevyShortName WHEN LevyShortName
                LIKE '%NSE%' THEN 'e' + LevyShortName WHEN LevyShortName LIKE '%CDS%' THEN 'f' + LevyShortName WHEN LevyShortName LIKE '%Comm%' THEN
                'g' + LevyShortName WHEN LevyShortName LIKE '%Agent%' THEN 'h' + LevyShortName ELSE LevyShortName END
WHERE ShortName IS NULL

GO


CREATE PROCEDURE sp_pivot
	/*
	   Purpose:     Creates a Pivot(tm) table for the specified table,
	                view or select statement
	   Author:      Knowing Ltd
	   Version:     1.1
	   History:     May 2004  version 1.1

	   Input parameters:
	  	@Aggregate_Function (optional)
	  		the aggregate function to use for the pivot
	                default function is SUM
	        @Aggregate_Column
	                name of column for aggregate
	        @TableOrView_Name
	                name of table or view to use
	                if name contains spaces or other special
	                characters [] should be used
	                Can also be a valid SELECT statement
	        @Select_Column
	                Column for first column in result table
	                for this column row values are displayed
	        @Pivot_Column
	                Column that is transformed into columns
	                for this column column values are displayed
		@DEBUG
			Set this flag to 1 to get debug-information

	      Example usage:
	        Table given   aTable
	        content:      Product    Salesman    Sales
	                      P1         Sa          12
	                      P2         Sb          10
	                      P2         Sb          3
	                      P3         Sa          12
	                      P1         Sc          8
	                      P3         Sa          1
	                      P2         Sa          NULL
		CALL
		EXEC sp_Transform 'SUM', 'Sales', 'aTable', 'Product', 'Salesman'
	or      EXEC sp_Transform @Aggregate_Column='Sales', @TableOrViewName='aTable',
	                            @Select_Column='Product', @Pivot_Column='Salesman'

	Result:
	        Product| Sa       | Sb       | Sc      | Total
	        -------+----------+----------+---------+---------
	        P1     | 12,00    |  0,00    |  8,00   |  20,00
	        P2     |  0,00    | 13,00    |  0,00   |  13,00
	        P3     | 13,00    |  0,00    |  0,00   |  13,00
	        -------+----------+----------+---------+---------
	        Total  | 25,00    | 13,00    |  8,00   |  46,00


	*/
	@Aggregate_Function nvarchar(30) = 'SUM',
	@Aggregate_Column   nvarchar(255),
	@TableOrView_Name   nvarchar(255),
	@Select_Column	    nvarchar(255),
	@Pivot_Column       nvarchar(255),
	@DEBUG		    bit = 0,
	@OutSQL nvarchar(4000) OUTPUT
AS


	SET NOCOUNT ON

	DECLARE @TransformPart   nvarchar(4000)
	DECLARE @SQLColRetrieval nvarchar(4000)
	DECLARE @SQLSelectIntro  nvarchar(4000)
	DECLARE @SQLSelectFinal  nvarchar(4000)

	IF @Aggregate_Function NOT IN ('SUM', 'COUNT', 'MAX', 'MIN', 'AVG', 'STDEV', 'VAR', 'VARP', 'STDEVP')
		BEGIN RAISERROR ('Invalid aggregate function: %s', 10, 1, @Aggregate_Function) END
	ELSE

	BEGIN

		SELECT @SQLSelectIntro = 'SELECT CASE WHEN (GROUPING('  +
				     	QUOTENAME(@Select_Column)       +
					') = 1) THEN ''Total'' ELSE '   +
					'CAST( + '                      +
		                        QUOTENAME(@Select_Column)       +
					' AS NVARCHAR(255)) END As '    +
					QUOTENAME(@Select_Column)       +
					', '
		SET @SQLColRetrieval =
		N'SELECT @TransformPart = CASE WHEN @TransformPart IS NULL THEN ' +
				N'''' + @Aggregate_Function + N'(CASE CAST(' +
				QUOTENAME(CAST(@Pivot_Column AS VARCHAR(255))) +
				N' AS VARCHAR(255)) WHEN '''''' + CAST('  +
				QUOTENAME(@Pivot_Column) +
				N' AS NVarchar(255)) + '''''' THEN ' + @Aggregate_Column +
				N' ELSE 0 END) AS '' + QUOTENAME(' +
				QUOTENAME(CAST(@Pivot_Column AS VARCHAR(255))) +
				N') ELSE  @TransformPart + '', ' + @Aggregate_Function +
				N' (CASE CAST(' + QUOTENAME(@Pivot_Column) +
				N' AS nVARCHAR(255)) WHEN '''''' + CAST(' +
				QUOTENAME(CAST(@Pivot_Column As VarChar(255))) +
				N' AS nVARCHAR(255)) + '''''' THEN ' +
				@Aggregate_Column +
				N' ELSE 0 END) AS '' + QUOTENAME(' +
				QUOTENAME(CAST(@Pivot_Column AS VARCHAR(255))) +
				N') END FROM (SELECT DISTINCT ' +
				QUOTENAME(CAST(@Pivot_Column AS VARCHAR(255))) +
				N' FROM ' + @TableOrView_Name + ') SelInner'
		EXEC sp_executesql @SQLColRetrieval,
		                   N'@TransformPart nvarchar(4000) OUTPUT',
		                   @TransformPart OUTPUT
		SET @SQLSelectFinal =
		                   N', ' + @Aggregate_Function + N'(' +
		                   CAST(@Aggregate_Column As Varchar(255)) +
		                   N') As Total FROM ' + @TableOrView_Name + N' GROUP BY ' +
		                   @Select_Column + N' WITH CUBE'


		SELECT @OutSQL = @SQLSelectIntro + @TransformPart + @SQLSelectFinal
	END

GO

CREATE PROCEDURE sp_transform
/*
   Purpose:     Creates a Pivot(tm) table for the specified table,
                view or select statement
   Author:      svenh@itrain.de
   Version:     1.1
   History:     march 2000 version 1.0
                july 2002  version 1.1

   Input parameters:
  	@Aggregate_Function (optional)
  		the aggregate function to use for the pivot
                default function is SUM
        @Aggregate_Column
                name of column for aggregate
        @TableOrView_Name
                name of table or view to use
                if name contains spaces or other special
                characters [] should be used
                Can also be a valid SELECT statement
        @Select_Column
                Column for first column in result table
                for this column row values are displayed
        @Pivot_Column
                Column that is transformed into columns
                for this column column values are displayed
	@DEBUG
		Set this flag to 1 to get debug-information

      Example usage:
        Table given   aTable
        content:      Product    Salesman    Sales
                      P1         Sa          12
                      P2         Sb          10
                      P2         Sb          3
                      P3         Sa          12
                      P1         Sc          8
                      P3         Sa          1
                      P2         Sa          NULL
	CALL
	EXEC sp_Transform 'SUM', 'Sales', 'aTable', 'Product', 'Salesman'
or      EXEC sp_Transform @Aggregate_Column='Sales', @TableOrViewName='aTable',
                            @Select_Column='Product', @Pivot_Column='Salesman'

Result:
        Product| Sa       | Sb       | Sc      | Total
        -------+----------+----------+---------+---------
        P1     | 12,00    |  0,00    |  8,00   |  20,00
        P2     |  0,00    | 13,00    |  0,00   |  13,00
        P3     | 13,00    |  0,00    |  0,00   |  13,00
        -------+----------+----------+---------+---------
        Total  | 25,00    | 13,00    |  8,00   |  46,00


*/
	@Aggregate_Function nvarchar(30) = 'SUM',
	@Aggregate_Column   nvarchar(255),
	@TableOrView_Name   nvarchar(255),
	@Select_Column	    nvarchar(255),
	@Pivot_Column       nvarchar(255),
	@DEBUG		    bit = 0
AS
SET NOCOUNT ON
DECLARE @TransformPart   nvarchar(4000)
DECLARE @SQLColRetrieval nvarchar(4000)
DECLARE @SQLSelectIntro  nvarchar(4000)
DECLARE @SQLSelectFinal  nvarchar(4000)

IF @Aggregate_Function NOT IN ('SUM', 'COUNT', 'MAX', 'MIN', 'AVG', 'STDEV', 'VAR', 'VARP', 'STDEVP')
	BEGIN RAISERROR ('Invalid aggregate function: %s', 10, 1, @Aggregate_Function) END
ELSE
BEGIN
	SELECT @SQLSelectIntro = 'SELECT CASE WHEN (GROUPING('  +
			     	QUOTENAME(@Select_Column)       +
				') = 1) THEN ''Total'' ELSE '   +
				'CAST( + '                      +
                                QUOTENAME(@Select_Column)       +
				' AS NVARCHAR(255)) END As '    +
				QUOTENAME(@Select_Column)       +
				', '
	IF @DEBUG = 1 PRINT @sqlselectintro
	SET @SQLColRetrieval =
	N'SELECT @TransformPart = CASE WHEN @TransformPart IS NULL THEN ' +
			N'''' + @Aggregate_Function + N'(CASE CAST(' +
			QUOTENAME(CAST(@Pivot_Column AS VARCHAR(255))) +
			N' AS VARCHAR(255)) WHEN '''''' + CAST('  +
			QUOTENAME(@Pivot_Column) +
			N' AS NVarchar(255)) + '''''' THEN ' + @Aggregate_Column +
			N' ELSE 0 END) AS '' + QUOTENAME(' +
			QUOTENAME(CAST(@Pivot_Column AS VARCHAR(255))) +
			N') ELSE  @TransformPart + '', ' + @Aggregate_Function +
			N' (CASE CAST(' + QUOTENAME(@Pivot_Column) +
			N' AS nVARCHAR(255)) WHEN '''''' + CAST(' +
			QUOTENAME(CAST(@Pivot_Column As VarChar(255))) +
			N' AS nVARCHAR(255)) + '''''' THEN ' +
			@Aggregate_Column +
			N' ELSE 0 END) AS '' + QUOTENAME(' +
			QUOTENAME(CAST(@Pivot_Column AS VARCHAR(255))) +
			N') END FROM (SELECT DISTINCT ' +
			QUOTENAME(CAST(@Pivot_Column AS VARCHAR(255))) +
			N' FROM ' + @TableOrView_Name + ') SelInner'
	IF @DEBUG = 1 PRINT @SQLColRetrieval
	EXEC sp_executesql @SQLColRetrieval,
                           N'@TransformPart nvarchar(4000) OUTPUT',
                           @TransformPart OUTPUT
	IF @DEBUG = 1 PRINT @TransformPart
	SET @SQLSelectFinal =
                           N', ' + @Aggregate_Function + N'(' +
                           CAST(@Aggregate_Column As Varchar(255)) +
                           N') As Total FROM ' + @TableOrView_Name + N' GROUP BY ' +
                           @Select_Column + N' WITH CUBE'
	IF @DEBUG = 1 PRINT @SQLSelectFinal
	EXEC (@SQLSelectIntro + @TransformPart + @SQLSelectFinal)
END

GO
create proc temp

as


Declare @Amount float
set @Amount = 12.775


GO
CREATE proc Temp02

as


Declare @Amount float
set @Amount = 12.2245

Declare @Round1 float
Declare @Round2 float
Declare @Diff1 float
Declare @Check1 float
Declare @Check2 float
Declare @Check3 float
Declare @CheckSum float
Declare @Result float

set @Round1 = ROUND(@Amount*100,0)
set @Round2 = FLOOR(@Amount*10)*10
set @Diff1 = @Round1-@Round2
set @Check1 = (case when @Diff1<=2 then 0 else 0 end)
set @Check2 = (case when @Diff1>=3 and @Diff1<=7  then 5 else 0 end)
set @Check3 = (case when @Diff1>=8 then 10 else 0 end)
set @CheckSum = @Check1+@Check2+@Check3
set @Result = round((@Round2+@CheckSum)/100,2)


select @Result


GO
CREATE procedure Temp03 AS  SELECT   *  FROM (SELECT     CONVERT(SMALLDATETIME, CAST(Lot.LotTDate AS CHAR(12))) AS LotTDate, Lot.Contract_DPA_ AS Contract, Lot.ContractNumber AS ContractNumber,
                      OrdDetailList.OrderTypeSale, Security.SecurityCode, OrdDetailList.OrderSecType_DPA_, Broker.BrokerCode, Lot.LotSlipNo, Lot.LotPrice, Lot.LotQty,
                      OrdDetailList.OrdDetailClient AS ClientName, OrdDetailList.Client_DPA_ AS Client_DPA_, Lot.Contract_DPA_, Lot.LotGrossAmount,
                      CAST(FLOOR(CAST(Lot.ContractSettlementDate AS float)) AS datetime) AS ContractSettlementDate
FROM         OrdDetailList INNER JOIN
                      Lot ON OrdDetailList.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN
                      Broker ON Lot.Broker_DPA_ = Broker.Broker_DPA_ LEFT OUTER JOIN
                      Security ON OrdDetailList.Security_DPA_ = Security.Security_DPA_
WHERE     (Lot.Deleted <> 1) AND (CAST(FLOOR(CAST(Lot.ContractSettlementDate AS float)) AS datetime) = CONVERT(DATETIME, '16-Jul-2008'))) AS INNERTBL
			INNER JOIN (SELECT CASE WHEN (GROUPING([Contract_DPA_]) = 1) THEN 'Total' ELSE CAST( + [Contract_DPA_] AS NVARCHAR(255)) END As [Contract_DPA_], MAX(CASE CAST([LevyShortName] AS VARCHAR(255)) WHEN 'Agent' THEN LevyAmount ELSE 0 END) AS [Agent], MAX (CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN 'BSEComm' THEN LevyAmount ELSE 0 END) AS [BSEComm], MAX (CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN 'Commission' THEN LevyAmount ELSE 0 END) AS [Commission], MAX (CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN 'CSD' THEN LevyAmount ELSE 0 END) AS [CSD], MAX (CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN 'Handling' THEN LevyAmount ELSE 0 END) AS [Handling], MAX (CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN 'Transfer' THEN LevyAmount ELSE 0 END) AS [Transfer], MAX (CASE CAST([LevyShortName] AS nVARCHAR(255)) WHEN 'VAT' THEN LevyAmount ELSE 0 END) AS [VAT], MAX(LevyAmount) As Total FROM LevyContracts GROUP BY Contract_DPA_ WITH CUBE) CROSSTABTBL ON CONVERT(NVARCHAR(1000), INNERTBL.Contract_DPA_) = CONVERT(NVARCHAR(1000), CROSSTABTBL.Contract_DPA_  )
GO

CREATE PROCEDURE dbo.TestClientStatement AS

SELECT     TOP 100 PERCENT COUNT(*) AS ClientTransaction_DPA_,
		a.Client_DPA_,
		a.TransDate, a.Ref,
		a.Particulars,
		a.Debit,
		a.Credit,
		CASE WHEN (SUM(b.Balance) >= 0) Then CONVERT(NVARCHAR, SUM(b.Balance)) + ' Cr' ELSE CONVERT(NVARCHAR,  ABS(SUM(b.Balance))) + ' Dr' END   AS Balance,
                a.IsOpeningBalance
FROM         (SELECT * FROM dbo.ClientTransactionList) a CROSS JOIN
                          (SELECT * FROM dbo.ClientTransactionList) b
WHERE a.TransDate >= b.TransDate AND a.Client_DPA_ = b.Client_DPA_
GROUP BY a.Client_DPA_, a.TransDate, a.Ref, a.Particulars, a.Debit, a.Credit, a.Balance, a.IsOpeningBalance
ORDER BY a.Client_DPA_, a.ClientTransaction_DPA_, a.Particulars, a.Ref, a.Debit, a.Credit, a.Balance, a.IsOpeningBalance

GO

CREATE PROCEDURE [dbo].[TrialBalanceProc] @StartDate DateTime,@EndDate DateTime AS
SET NOCOUNT ON
SELECT TBal.*
INTO #tableA
FROM ( SELECT '******' + CONVERT(VARCHAR(500),dbo.FullEntityTypeList.EntityType_DPA_) AS [Account Code],
 dbo.FullEntityTypeList.EntityTypeName AS AccountName, CASE WHEN SUM(dbo.EntityTransactionList.Balance) < 0
THEN 0 - SUM(dbo.EntityTransactionList.Balance) ELSE 0 END AS Debit, CASE
WHEN SUM(dbo.EntityTransactionList.Balance) >= 0 THEN SUM(dbo.EntityTransactionList.Balance) ELSE 0 END AS Credit,
Cast(Floor(Cast(Min(EntityTransactionList.TransDate) as Float)) as DateTime) as TransDate FROM dbo.Entity
INNER JOIN dbo.FullEntityTypeList ON dbo.Entity.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_
INNER JOIN dbo.EntityTransactionList ON dbo.Entity.Entity_DPA_ = dbo.EntityTransactionList.Entity_DPA_ GROUP BY
dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName,
dbo.FullEntityTypeList.EntityType_DPA_ ) TBal

INSERT INTO #tableA
SELECT '******2' AS [Account Code],
 (SELECT EntityTypeName FROM dbo.FullEntityTypeList
WHERE EntityType_DPA_ = 2) AS AccountName,Case When Sum(AgentStatement.Balance) <0
then 0-Sum(AgentStatement.Balance) else 0 end as Debit,Case When Sum(AgentStatement.Balance)>=0
then Sum(AgentStatement.Balance) else 0 end as Credit,TransDate from AgentStatement
Group by TransDate

INSERT INTO #tableA
SELECT '******3' AS [Account Code],
(SELECT EntityTypeName FROM dbo.FullEntityTypeList WHERE EntityType_DPA_ = 3)
AS AccountName, CASE WHEN SUM(Credit-Debit) < 0 THEN 0 - SUM(Credit-Debit) ELSE 0 END AS Debit,
CASE WHEN SUM(Credit-Debit) >= 0 THEN SUM(Credit-Debit) ELSE 0 END AS Credit,
Cast(Floor(Cast(TransDate as Float)) as Datetime) as TransDate FROM BrokerStatement
Group by Cast(Floor(Cast(TransDate as Float)) as Datetime)

INSERT INTO #tableA
SELECT '******1' AS [Account Code], (SELECT EntityTypeName FROM dbo.FullEntityTypeList
WHERE EntityType_DPA_ = 1) AS AccountName,CASE WHEN SUM(StatementList.Balance) < 0
THEN 0 - SUM(StatementList.Balance) ELSE 0 END AS Debit, CASE WHEN SUM(StatementList.Balance) >= 0
THEN SUM(StatementList.Balance) ELSE 0 END AS Credit,TransDate from StatementList Group by TransDate

INSERT INTO #tableA
SELECT '******7' AS [Account Code], (SELECT EntityTypeName FROM dbo.FullEntityTypeList
WHERE EntityType_DPA_ = 7) AS AccountName, CASE WHEN SUM(Credit-Debit) < 0 THEN 0 - SUM(Credit-Debit) ELSE 0
END AS Debit, CASE WHEN SUM(Credit-Debit) >= 0 THEN SUM(Credit-Debit) ELSE 0 END
AS Credit,Cast(Floor(Cast(TransDate as Float)) as DateTime) as TransDate FROM OwnerStatement
Group by Cast(Floor(Cast(TransDate as Float)) as DateTime)

INSERT INTO #tableA
SELECT Distinct '------' + CONVERT(VARCHAR(500),Entity.Entity_DPA_) AS [Account Code], Entity.EntityName AS
AccountName, CASE WHEN SUM(Balance) < 0 THEN 0 - SUM(Balance) ELSE 0 END AS Debit, CASE WHEN SUM(Balance) >= 0
THEN SUM(Balance) ELSE 0 END AS Credit,Cast(Floor(Cast(TransDate as Float)) as DateTime) as
TransDate FROM BrokerCommissionStatement inner join Entity on
BrokerCommissionStatement.Entity_DPA_=Entity.Entity_DPA_ Group by Entity.Entity_DPA_,Entity.EntityName,
Cast(Floor(Cast(TransDate as Float)) as DateTime)

INSERT INTO #tableA
SELECT Distinct '------' + CONVERT(VARCHAR(500),Entity.Entity_DPA_) AS [Account code],Entity.EntityName AS
AccountName, CASE WHEN SUM(Balance) < 0 THEN 0- SUM(Balance) ELSE 0 END as Debit, CASE WHEN SUM(Balance) >= 0
THEN SUM(Balance) ELSE 0 END AS Credit,TransDate From CommissionStatement inner join Entity on
CommissionStatement.Entity_DPA_ =Entity.Entity_DPA_ Group By Entity.Entity_DPA_,Entity.EntityName,TransDate

INSERT INTO #tableA
SELECT '------' + CONVERT(VARCHAR(500),Entity.Entity_DPA_) AS [Account Code],Entity.EntityName AS AccountName,
CASE WHEN SUM(Balance) < 0 THEN 0 - SUM(Balance) else 0 END AS Debit, CASE WHEN SUM(Balance) >= 0
THEN SUM(Balance) ELSE 0 End AS Credit,TransDate From CDSControlStatement inner join Entity
on CDSControlStatement.Client_DPA_=Entity.Entity_DPA_ Group by Entity_DPA_,Entity.EntityName,TransDate

INSERT INTO #tableA
SELECT CONVERT(VARCHAR(500),dbo.AccountList.AccountCode) AS [Account Code], dbo.AccountList.AccountName,
CASE WHEN SUM(dbo.DB_BankAccountStatement.Balance) < 0 THEN 0 - SUM(DB_BankAccountStatement.Balance)
ELSE 0 END AS Debit, CASE WHEN SUM(dbo.DB_BankAccountStatement.Balance) >= 0
THEN SUM(dbo.DB_BankAccountStatement.Balance) ELSE 0 END AS Credit,Cast(Floor(Cast(Db_BankAccountStatement.TransDate
as Float)) as DateTime) as TransDate FROM dbo.DB_BankAccountStatement INNER JOIN dbo.AccountList
ON dbo.DB_BankAccountStatement.BankAccount_DPA_ = dbo.AccountList.Account_DPA_
GROUP BY dbo.AccountList.AccountTypeLevel1,dbo.AccountList.AccountName, dbo.AccountList.AccountCode,
Cast(Floor(Cast(Db_BankAccountStatement.TransDate as Float)) as DateTime)

SELECT IDENTITY(int, 1,1) AS SequenceID, #tableA.* INTO #tableB FROM #tableA
ORDER BY [Account Code] select T.[Account Code], T.AccountName,CASE WHEN
Sum(T.Credit-T.Debit) <0 then 0-Sum(T.Credit-T.Debit) else 0 end as Debit,
CASE WHEN Sum(T.Credit-T.Debit)>=0 then Sum(T.Credit-T.Debit) else 0 end as Credit
from #tableB T Where T.TransDate between @StartDate and @enddate
Group by T.[Account Code], T.AccountName order by T.[Account Code]

Drop table #tableA
Drop table #tableB




GO

create proc UpdateClientBalanceFromJournal
@Journal_DPA_ int
as
INSERT INTO ClientBalances
                      (client_DPA_, CurrentBal)
SELECT     JournalEntry.Entity_DPA_, 0 AS CurrentBal
FROM         JournalEntry LEFT OUTER JOIN
                      ClientBalances ON JournalEntry.Entity_DPA_ = ClientBalances.client_DPA_
WHERE     (JournalEntry.EntityType_DPA_ = 1) AND (JournalEntry.Journal_DPA_ = @Journal_DPA_) AND (ClientBalances.client_DPA_ IS NULL)
UPDATE    ClientBalances
SET              CurrentBal = ClientBalances.CurrentBal - JournalEntry.JournalEntryDebit + JournalEntry.JournalEntryCredit
FROM         Journal INNER JOIN
                      JournalEntry ON Journal.Journal_DPA_ = JournalEntry.Journal_DPA_ LEFT OUTER JOIN
                      ClientBalances ON JournalEntry.Entity_DPA_ = ClientBalances.client_DPA_
WHERE     (Journal.Journal_DPA_ = @Journal_DPA_) AND (JournalEntry.EntityType_DPA_ = 1) AND (JournalEntry.Deleted <> 1) AND (Journal.Deleted <> 1) AND
                      (Journal.JournalCommitted = 1)


GO
CREATE PROCEDURE [dbo].[UpdateClientBalances] AS
SET NOCOUNT ON

exec ClientBalancesDelete
exec ClientBalancesProcedure
exec ClientTotalsDelete
exec ClientTotalsProcedure
GO
CREATE PROCEDURE UpdateCompoundedContractCommissions as
UPDATE    LevyContract
SET              LevyAmount = ProperCommission3
FROM         (SELECT     ShowProperCommissionForCompoundedContractsForEachLevyEntry.LotTDate,
                                              ShowProperCommissionForCompoundedContractsForEachLevyEntry.LevyContract_DPA_,
                                              ShowProperCommissionForCompoundedContractsForEachLevyEntry.ProperCommissionRate,
                                              ShowProperCommissionForCompoundedContractsForEachLevyEntry.ProperCommission3
                       FROM          ShowProperCommissionForCompoundedContractsForEachLevyEntry INNER JOIN
                                                  (SELECT     CAST(FLOOR(CAST(MAX(LotTDate) AS float)) AS datetime) AS LastDate
                                                    FROM          Lot
                                                    WHERE      (Deleted <> 1)) LastDateTransactions ON
                                              ShowProperCommissionForCompoundedContractsForEachLevyEntry.LotTDate = LastDateTransactions.LastDate) A
WHERE     LevyContract.LevyContract_DPA_ = A.LevyContract_DPA_

UPDATE    LevyContract
SET              LevyAmount = ProperAgentAmount
FROM         (SELECT     LevyContract.Contract_DPA_, LevyContract_1.LevyContract_DPA_ AS AgentLevyContract_DPA_,
                                              ShowProperCommissionForCompoundedContractsForEachLevyEntry.ProperCommission3 AS BrokerAmount,
                                              LevyContract_1.LevyAmount AS AgentAmount, Commission.CommissionRate,
                                              ROUND(Commission.CommissionRate / 100 * ShowProperCommissionForCompoundedContractsForEachLevyEntry.ProperCommission3, 2)
                                              AS ProperAgentAmount
                       FROM          tbOrder INNER JOIN
                                              OrdDetail ON tbOrder.Order_DPA_ = OrdDetail.Order_DPA_ INNER JOIN
                                              Lot ON OrdDetail.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN
                                              ShowProperCommissionForCompoundedContractsForEachLevyEntry INNER JOIN
                                              LevyContract ON
                                              ShowProperCommissionForCompoundedContractsForEachLevyEntry.LevyContract_DPA_ = LevyContract.LevyContract_DPA_ INNER JOIN
                                              LevyContract LevyContract_1 ON LevyContract.Contract_DPA_ = LevyContract_1.Contract_DPA_ ON
                                              Lot.Contract_DPA_ = LevyContract.Contract_DPA_ INNER JOIN
                                              Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN
                                              Agent ON Client.Agent_DPA_ = Agent.Agent_DPA_ INNER JOIN
                                              Commission ON Agent.Commission_DPA_ = Commission.Commission_DPA_ INNER JOIN
                                                  (SELECT     CAST(FLOOR(CAST(MAX(LotTDate) AS float)) AS datetime) AS LastDate
                                                    FROM          Lot
                                                    WHERE      (Deleted <> 1)) LastDateTransactions ON CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime)
                                              = CAST(FLOOR(CAST(LastDateTransactions.LastDate AS float)) AS datetime)
                       WHERE      (LevyContract_1.SystemMaintained = 12) AND (Lot.Deleted <> 1)) A
WHERE     LevyContract.LevyContract_DPA_ = A.AgentLevyContract_DPA_
GO

CREATE proc UpdateIndividualClientBalance
@Client_DPA_ int,
@RecAmount money
as
Declare @ClientExists tinyint
set @ClientExists = isnull ((
SELECT     count(*) as No
FROM         ClientBalances
WHERE     (client_DPA_ = @Client_dpa_)
),0)
Declare @CurrentBalance money
set @CurrentBalance = isnull ((
SELECT     CurrentBal
FROM         ClientBalances
WHERE     (client_DPA_ = @Client_dpa_)
),0)
if @ClientExists = 0
Begin
INSERT INTO ClientBalances
                      (CurrentBal, client_DPA_)
VALUES     (@RecAmount, @Client_dpa_)
End
Else
Begin
UPDATE    ClientBalances
SET              CurrentBal = @CurrentBalance + @RecAmount
WHERE     (client_DPA_ = @Client_dpa_)
End
GO
