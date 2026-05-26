



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

set @Side = (SELECT     LEFT(OrderType.OrderTypeDescription, 1) AS Side FROM OrdDetail INNER JOIN tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN OrderType ON tbOrder.OrderType_DPA_ = OrderType.OrderType_DPA_ WHERE (OrdDetail.OrdDetail_DPA
_ = @OrdDetail_DPA_))



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













