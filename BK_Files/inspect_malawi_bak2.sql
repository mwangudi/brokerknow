SET NOCOUNT ON;
CREATE TABLE #h (
  BackupName nvarchar(128) NULL, BackupDescription nvarchar(255) NULL, BackupType smallint NULL,
  ExpirationDate datetime NULL, Compressed bit NULL, Position smallint NULL, DeviceType tinyint NULL,
  UserName nvarchar(128) NULL, ServerName nvarchar(128) NULL, DatabaseName nvarchar(128) NULL,
  DatabaseVersion int NULL, DatabaseCreationDate datetime NULL, BackupSize numeric(20,0) NULL,
  FirstLSN numeric(25,0) NULL, LastLSN numeric(25,0) NULL, CheckpointLSN numeric(25,0) NULL,
  DatabaseBackupLSN numeric(25,0) NULL, BackupStartDate datetime NULL, BackupFinishDate datetime NULL,
  SortOrder smallint NULL, CodePage smallint NULL, UnicodeLocaleId int NULL, UnicodeComparisonStyle int NULL,
  CompatibilityLevel tinyint NULL, SoftwareVendorId int NULL, SoftwareVersionMajor int NULL,
  SoftwareVersionMinor int NULL, SoftwareVersionBuild int NULL, MachineName nvarchar(128) NULL,
  Flags int NULL, BindingID uniqueidentifier NULL, RecoveryForkID uniqueidentifier NULL,
  Collation nvarchar(128) NULL, FamilyGUID uniqueidentifier NULL, HasBulkLoggedData bit NULL,
  IsSnapshot bit NULL, IsReadOnly bit NULL, IsSingleUser bit NULL, HasBackupChecksums bit NULL,
  IsDamaged bit NULL, BeginsLogChain bit NULL, HasIncompleteMetaData bit NULL, IsForceOffline bit NULL,
  IsCopyOnly bit NULL, FirstRecoveryForkID uniqueidentifier NULL, ForkPointLSN numeric(25,0) NULL,
  RecoveryModel nvarchar(60) NULL, DifferentialBaseLSN numeric(25,0) NULL,
  DifferentialBaseGUID uniqueidentifier NULL, BackupTypeDescription nvarchar(60) NULL,
  BackupSetGUID uniqueidentifier NULL, CompressedBackupSize numeric(20,0) NULL, Containment tinyint NULL,
  KeyAlgorithm nvarchar(32) NULL, EncryptorThumbprint varbinary(20) NULL, EncryptorType nvarchar(32) NULL,
  LastValidRestoreTime datetime NULL, TimeZone nvarchar(32) NULL, CompressionAlgorithm nvarchar(32) NULL
);
INSERT INTO #h EXEC('RESTORE HEADERONLY FROM DISK = N''/var/opt/mssql/restore/Malawi120626_02.bak''');
SELECT 'HEADER' AS section, DatabaseName, BackupTypeDescription, CONVERT(varchar,BackupFinishDate,120) AS finished,
       RecoveryModel, CompatibilityLevel, ServerName
FROM #h;
GO
SET NOCOUNT ON;
CREATE TABLE #f (
  LogicalName nvarchar(128), PhysicalName nvarchar(260), Type char(1), FileGroupName nvarchar(128) NULL,
  Size numeric(20,0), MaxSize numeric(20,0), FileId bigint, CreateLSN numeric(25,0), DropLSN numeric(25,0) NULL,
  UniqueId uniqueidentifier, ReadOnlyLSN numeric(25,0) NULL, ReadWriteLSN numeric(25,0) NULL,
  BackupSizeInBytes bigint, SourceBlockSize int, FileGroupId int, LogGroupGUID uniqueidentifier NULL,
  DifferentialBaseLSN numeric(25,0) NULL, DifferentialBaseGUID uniqueidentifier NULL, IsReadOnly bit,
  IsPresent bit, TDEThumbprint varbinary(32) NULL, SnapshotUrl nvarchar(360) NULL
);
INSERT INTO #f EXEC('RESTORE FILELISTONLY FROM DISK = N''/var/opt/mssql/restore/Malawi120626_02.bak''');
SELECT 'FILE' AS section, LogicalName, Type, PhysicalName FROM #f;
GO
