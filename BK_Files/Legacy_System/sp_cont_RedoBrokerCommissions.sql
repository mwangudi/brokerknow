

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



