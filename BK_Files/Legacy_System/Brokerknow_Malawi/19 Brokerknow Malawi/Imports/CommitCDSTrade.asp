<!--#include file="../libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "CommitCDSTrade"
	const DataEntity = "CDSTrade"
	const DataEntityPlural = "CDSTrades"
	const ActionFolder = "Import"
	
	Dim action
	Dim conn 
   	Dim sqlStr
   	Dim rs
   	Dim ID
   	Dim rsEdit
   	Dim IDHolder
	Dim IDArray
	Dim ItemID
	Dim OrderDetail_DPA_
	
	UserId=Session("UserID")			

	action = ucase(Request.Form("delAction"))
	Params = Request.Form("CommitParams")
	
	
	OrderDetail_DPA_ = Request("OrderDetail_DPA_")
		
	ID = Request("ID")			

	itemids=Split(Params,",")	
	
	
		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No trade specified for committing"
                		
                </script>
                <% response.end
        End If                

	select case action 
		case "EXECUTE"

			For i=0 to UBound(itemids)			

			Ids=Split(itemids(i),"<->")
			
			OrderDetail_DPA_=Ids(1)
			ID=Ids(0)
			'PDate=itemids(2)

			Dim CDSTradRS
			
			
			sqlStr = "SELECT * FROM CDSMatchedTradesList WHERE CDSImport_DPA_= " & ID & " and (OrdDetail_DPA_=" & OrderDetail_DPA_ & ")"			  
			
   			Set conn = GetActiveConnection("KBroker")
   			set CDSTradRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
   			If CDSTradRS.EOF Or CDSTradRS.BOF Then
			'Do nothing coz its already processed
			else
			'save data
			Dim slip
			Dim broker
			Dim tDate
			Dim qty
			Dim price
			Dim orderType
			Dim orderIsSaleType
			Dim securityID
			Dim commission
			Dim agentCommission
			Dim staffCommission
			Dim volComm
			Dim volBound
			Dim minComm
			Dim cma
			Dim imobRate
			Dim secImob
			Dim regularComm
			Dim ContractDPA
			Dim interbank
			Dim custodian
			Dim cliententity
			Dim client
			Dim Commission1
			Dim Commission2
			Dim SettlementDate
			Dim ClientDpa

			volComm = CDSTradRS.Fields("VolumeRate")
			volBound = CDSTradRS.Fields("VolumeBoundary")
			minComm = CDSTradRS.Fields("MinimumCommission")
			cma = CDSTradRS.Fields("CMARegulated")
			imobRate = CDSTradRS.Fields("PostImmobilisedRate")
			secImob = CDSTradRS.Fields("SecurityImmobilised")
			broker = CDSTradRS.Fields("Broker_DPA_")
			tDate = CDSTradRS.Fields("TradeDate")
			tDate = tDate & " " & formatdatetime(CDSTradRS.Fields("TradeTime"),vbShortTime)
			slip = trim(CDSTradRS.Fields("CDSRef"))
			qty = clng(CDSTradRS.Fields("Quantity"))
			price = ccur(CDSTradRS.Fields("Price"))
			orderType = CDSTradRS.Fields("OrdDetailType")
			orderSecType = CDSTradRS.Fields("OrdDetailSecType")
			orderIsSaleType = cbool(CDSTradRS.Fields("OrderTypeSale"))
			securityID = CDSTradRS.Fields("Security_DPA_")
			agentCommission = CDSTradRS.Fields("AgentCommission")
			staffCommission = CDSTradRS.Fields("StaffCommission")
			regularComm = CDSTradRS.Fields("CommissionRate")
			SettlementDate=CDSTradRS.Fields("SettlementDate")
			clientClass =Cint(Request.Form("Class"))
			'EnteredCommission=CDSTradRS.Fields("Commission")
		    EnteredCommission = ""
			ClientDpa=CDSTradRS.Fields("Client_DPA_")

			if isnull(EnteredCommission) or EnteredCommission ="" then
			else
			regularComm=EnteredCommission
			end if

			if(CDSTradRS.Fields("InterBank")=true) then	
			interbank=1
			else
			interbank=0
			end if
			
			if(CDSTradRS.Fields("IsCustodian")=true) then	
			custodian=1
			else
			custodian=0
			end if

  			Set TimeLimitRs = CreateObject("ADODB.Recordset")   						        
   			TimeLimitRs.CursorLocation = adUseClient

			client=CDSTradRS.Fields("Client_DPA_")
			cliententity =CDSTradRS.Fields("EntityType_DPA_")
			
			if(ClientClass =6)	 then
				if orderIsSaleType then
         				sqlStr="SELECT TimeLimitLimDaysNSE From TimeLimit where TimeLimit_DPA_=7"   
				else
					sqlStr="SELECT TimeLimitLimDaysNSE From TimeLimit where TimeLimit_DPA_=5"   
				end if
        		else
				if orderSecType = "Fixed" then
				sqlStr="SELECT TimeLimitLimDaysNSE From TimeLimit where TimeLimit_DPA_=6"   
				else
   				sqlStr="SELECT TimeLimitLimDaysNSE From TimeLimit where TimeLimit_DPA_=1"   
				end if
   			end if
   
   				set TimeLimitRs=Conn.Execute(sqlStr) 
   
				if not(TimeLimitRS.eof and TimeLimitRs.bof) then
	 			NoOfDays=TimeLimitRS("TimeLimitLimDaysNSE")
				end if
	
				'settlementDate=LTdate(CDate(tDate),NoOfDays)				

			sqlStr="SELECT Lots.LotSlipNo, tbOrder.Client_DPA_ FROM Lots INNER JOIN " & _
                   " OrdDetail ON Lots.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN " & _
                   " tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_"

			set TimeLimitRs=Conn.Execute(sqlStr) 

			if not (TimeLimitRs.eof or TimeLimitRs.bof) then
				Do while TimeLimitRs.eof=false
					if(TimeLimitRs("LotSlipNo")=slip and TimeLimitRs("Client_DPA_")=ClientDpa) then
					%>
			        <script language = 'vbscript'>
			            	window.self.ShowMessage "The Slip No Already Exists"            	
							window.history.back
			        </script>
			        <%
					response.end
					end if
				TimeLimitRs.MoveNext
				loop
			end if

			'save contract
			 set guid = server.createobject("NDUtils.CGUID")
			 guidStr = guid.GenerateGUID
			
			'Response.write(custodian)
			'Response.end 	 

			 sqlStr = "INSERT INTO [Contract] (Contract_DPA_, Contract_EIT_, Status_DPA_,IsInterBank,ContractSettlementDate) " & _
			         "SELECT " & " " & "iif(isnull(max([Contract_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Contract'),max([Contract_DPA_]) + 1)" & " " & " as Contract_DPA_" & _
			         "," & "'" & guidStr & "'" & " as Contract_EIT_" & _
			         "," & " " & 1 & " " & " as Status_DPA_" & _
			         "," & " " & interbank & " " & " as IsInterBank" & _
				   "," & "#" & FormatDate(SettlementDate) & "#" & " as ContractSettlementDate" & _
			         " FROM [Contract]"
			         
			 Set conn = GetActiveConnection("KBroker")
			 conn.BeginTrans
					conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
				     
					'obtain contract number
					sqlStr = "SELECT [Contract.Contract_DPA_] FROM [Contract] WHERE [Contract.Contract_EIT_] = " & "'" & guidStr & "'"
				     
					Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					If (rs.EOF Or rs.BOF) Then%>
			         			<script language = 'vbscript'>
			         					ShowMessage "A serious error has been encountered while saving the data. Try saving again"
				         					
			         			</script>
			         			<% response.end
					End If
				     
				    ContractDPA=rs("Contract_DPA_") 
					'calculate amounts
					Dim grossAmount  'this is the amount before application of Levies
					    
						
					if orderSecType = "Fixed" then ' "F" is FIXED security
						grossAmount = (price * qty) / 100
					else
						grossAmount = price * qty
					end if
						
						
					'save lot
					sqlStr = "INSERT INTO [Lot] (Lot_DPA_,Contract_DPA_,OrdDetail_DPA_,LotPrice" & _
							",LotQty,LotSlipNo,LotTDate,Broker_DPA_,ContractNumber, LotGrossAmount,CDSTransaction,CDSImport_DPA_,ChangedBy)" & _
							" SELECT " & " " & "iif(isnull(max([Lot_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Lot'),max([Lot_DPA_]) + 1)" & " " & " as Lot_DPA_" & _
							"," & " " & rs.fields("Contract_DPA_") & " " & " as Contract_DPA_" & _
							"," & " " & CDSTradRS.Fields("OrdDetail_DPA_") & " " & " as OrdDetail_DPA_" & _
							"," & " " & price & " " & " as LotPrice" & _
							"," & " " & qty & " " & " as LotQty" & _
							"," & " " & slip & " " & " as LotSlipNo" & _
							"," & "#" & FormatDate(tDate) & "#" & " as LotTDate" & _
							"," & " " & broker & " " & " as Broker_DPA_" & _
							"," & "'" & left(orderType,1) & rs.fields("Contract_DPA_") & "'" & " as ContractNumber" & _								
							"," & " " & grossAmount & " " & " as LotGrossAmount" & _
							"," & " " & 1 & " " & " as CDSTransaction" & _
							"," & " " & ID & " " & " as CDSImport_DPA_" & _
							"," & " " & UserId & " " & " as ChangedBy" & _
							" FROM [Lot]"
				     
					sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
					conn.Execute sqlStr
						
										
					'apply relevant levies
					Dim levyRS
					Dim cond
					if orderSecType = "Fixed" then 
						cond = "WHERE LevyAppBond = 1 AND LevyActive = 1"
					else
						cond = "WHERE (LevyAppSecurity = 1) AND (LevyActive = 1) AND (" & _
								" Levy_DPA_ IN (SELECT Levy_DPA_ FROM LevySecurity " & _
								" WHERE Security_DPA_ = " & securityID & "))"
					end if
						
					sqlStr = "SELECT * FROM [LevyList] " & cond
					Set levyRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
						
					Dim levyAmount
					Dim vatAmount
					vatAmount=0
					If Not (levyRS.EOF Or levyRS.BOF) Then
							

							levyRS.MoveFirst
							Do Until levyRS.EOF
									if levyRS.Fields("LevyType") = "P" Then ' "P" is for PERCENTAGE ie Levies that are a percentage of Gross
											levyAmount = CCur((levyRS.Fields("LevyAmount")/100.00) * grossAmount)
											LevyRatePercentage = levyRS.Fields("LevyAmount") & "%"
									else
											Dim blocks
											blocks = Int(grossAmount / levyRS.Fields("LevyBlock"))
											if (grossAmount mod levyRS.Fields("LevyBlock")) <> 0 then
													blocks = blocks + 1
											end if
											levyAmount = CCur(blocks * levyRS.Fields("LevyAmount"))
											LevyRatePercentage = levyRS.Fields("LevyAmount") & " for every " & levyRS.Fields("LevyBlock")
									end if

									'VAT
									if abs(levyRS.Fields("Vatable")) = 1 then 
										vatAmount= vatAmount + (levyAmount*)
									end if 
									'Calculate 
									sqlStr = "INSERT INTO [LevyContract] (LevyContract_DPA_,Contract_DPA_,LevyName,LevyAmount,LevyRate,LevyBlock,SystemMaintained,LevyShortName,LevyRatePercentage) SELECT " & " " & "iif(isnull(max([LevyContract_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'LevyContract'),max([LevyContract_DPA_]) + 1)" & " " & " as LevyContract_DPA_" & _
											"       ," & " " & rs.Fields("Contract_DPA_") & " " & " as Contract_DPA_" & _
											"       ," & "'" & levyRS.Fields("LevyDescription") & "'" & " as LevyName" & _
											"       ," & " " & RoundPoint05(levyAmount) & " " & " as LevyAmount" & _
											"       ," & " " & levyRS.Fields("LevyAmount") & " " & " as LevyRate" & _
											"       ," & " " & levyRS.Fields("LevyBlock") & " " & " as LevyBlock" & _
											"       ," & " " & levyRS.Fields("SystemMaintained") & " " & " as SystemMaintained" & _	
											"       ," & " '" & levyRS.Fields("LevyShortName") & "' " & " as LevyShortName" & _												
											"       ," & " '" & LevyRatePercentage  & "' " & " as LevyRatePercentage" & _												
											"        FROM [LevyContract]"
									conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
									levyRS.MoveNext
							Loop

							'Enter VAT
							sqlStr = "SELECT     * "
									"FROM         LevyList "
									"WHERE     (Vatable = 0) AND (LevyShortName LIKE N'%VAT%')"
							Set levyVATRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
							LevyRatePercentage = levyVATRS.Fields("LevyAmount") & "%"
							sqlStr = "INSERT INTO [LevyContract] (LevyContract_DPA_,Contract_DPA_,LevyName,LevyAmount,LevyRate,LevyBlock,SystemMaintained,LevyShortName,LevyRatePercentage) SELECT " & " " & "iif(isnull(max([LevyContract_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'LevyContract'),max([LevyContract_DPA_]) + 1)" & " " & " as LevyContract_DPA_" & _
											"       ," & " " & rs.Fields("Contract_DPA_") & " " & " as Contract_DPA_" & _
											"       ," & "'" & levyVATRS.Fields("LevyDescription") & "'" & " as LevyName" & _
											"       ," & " " & RoundPoint05(VatAmount) & " " & " as LevyAmount" & _
											"       ," & " " & levyVATRS.Fields("LevyAmount") & " " & " as LevyRate" & _
											"       ," & " " & levyVATRS.Fields("LevyBlock") & " " & " as LevyBlock" & _
											"       ," & " " & levyVATRS.Fields("SystemMaintained") & " " & " as SystemMaintained" & _	
											"       ," & " '" & levyVATRS.Fields("LevyShortName") & "' " & " as LevyShortName" & _												
											"       ," & " '" & LevyRatePercentage  & "' " & " as LevyRatePercentage" & _												
											"        FROM [LevyContract]"
							Set levyVATRS = nothing
									conn.Execute SQLServerFormatWithCustomMax(sqlStr)
					end if
						
					'consider the transfer fee
					if Not(orderIsSaleType) then
						Dim transRS
						sqlStr = "SELECT * FROM [SecTransFeeListLatest] WHERE Security_DPA_ = " & securityID
						Set transRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
						sqlStr = "INSERT INTO [LevyContract] (LevyContract_DPA_,Contract_DPA_,LevyName,LevyAmount,LevyRate,LevyBlock,LevyShortName,SystemMaintained) SELECT " & " " & "iif(isnull(max([LevyContract_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'LevyContract'),max([LevyContract_DPA_]) + 1)" & " " & " as LevyContract_DPA_" & _
								"       ," & " " & rs.Fields("Contract_DPA_") & " " & " as Contract_DPA_" & _
								"       ," & "'Transfer Fee'" & " as LevyName" & _
								"       ," & " " & transRS.Fields("Fee") & " " & " as LevyAmount" & _
								"       ," & " " & transRS.Fields("Fee") & " " & " as LevyRate" & _
								"       ," & " " & 0 & " " & " as LevyBlock" & _
								"       ," & " 'Transfer' " & " as LevyShortName" & _
								"       ,13 " & " as SystemMaintained" & _	
								"        FROM [LevyContract]"
						conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
					end if
						
					'Apply broker commission
						
					Dim commissionRS
					if cma then
							if secImob <> imobRate then
									if secImob then
											'fetch post-immobilisation commission
											sqlStr = "SELECT * FROM CommissionList WHERE (CMARegulated = 1) AND (Immobilised = 1)"
									else
											'fetch pre-immobilisation commission
											sqlStr = "SELECT * FROM CommissionList WHERE (CMARegulated = 1) AND (Immobilised = 0)"
									end if
										
									Set commissionRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
									If (commissionRS.EOF Or commissionRS.BOF) Then%>
									 			<script language = 'vbscript'>
									 					ShowMessage "The assigned commission rate is inviolation of CMA regulations"
										 					
									 			</script>
									 			<% response.end
									End If
										
									'use obtained rates
									if orderSecType = "Fixed" then 
												regularComm = commissionRS.Fields("BondCommission")
												volComm = commissionRS.Fields("UpperBondCommission")
												volBound = commissionRS.Fields("BondBoundary")
												minComm = commissionRS.Fields(" MinimumBondCommission")
												
												if volBound > 0 then
													if grossAmount >= volBound then
														Commission1=minComm*volBound
														Commission2=volComm*(grossAmount-volBound)
														commission = Commission1 + Commission2
													else
														commission = minComm
													end if
												else
													commission = regularComm
												end if
						

										else
												regularComm = commissionRS.Fields("CommissionRate")
												volComm = commissionRS.Fields("UpperSecurityCommission")
												volBound = commissionRS.Fields("SecurityBoundary")
												minComm = commissionRS.Fields("MinimumSecurityCommission")
												if volBound > 0 then
													if grossAmount >= volBound then
														commission = volComm
													else
														commission = regularComm
													end if
												else
													commission = regularComm
												end if
						
										end if
								end if
						end if						
						
					if volBound > 0 then
							if grossAmount >= volBound then
									commission = volComm
							else
									commission = regularComm
							end if
					else
							commission = regularComm
					end if
						
					if commission = "" then%>
								<script language = 'vbscript'>
										ShowMessage "The commission rate for this trade could not be determined"
										 					
								</script>
								<% response.end
					end if
						
					Dim commissionAmount
						
					commissionAmount = CCur((commission/100.00) * grossAmount)
												
					if commissionAmount < minComm then
							commissionAmount = minComm
					end if
						
					levyAmount = commissionAmount
						
					sqlStr = "INSERT INTO [LevyContract] (LevyContract_DPA_,Contract_DPA_,LevyName,LevyAmount,LevyRate,LevyBlock,LevyShortName,LevyRatePercentage,SystemMaintained) SELECT " & " " & "iif(isnull(max([LevyContract_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'LevyContract'),max([LevyContract_DPA_]) + 1)" & " " & " as LevyContract_DPA_" & _
							"       ," & " " & rs.Fields("Contract_DPA_") & " " & " as Contract_DPA_" & _
							"       ," & "'Broker Commission'" & " as LevyName" & _
							"       ," & " " & RoundPoint05(levyAmount) & " " & " as LevyAmount" & _
							"       ," & " " & commission & " " & " as LevyRate" & _
							"       ," & " " & 0 & " " & " as LevyBlock" & _
							"       ," & " 'Commission' " & " as LevyShortName" & _
							"       ," & " '" & commission & "%" & "' " & " as LevyRatePercentage" & _
							"       ,11 " & " as SystemMaintained" & _	
							"        FROM [LevyContract]"
					sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
					conn.Execute sqlStr
						
					'Apply agent commission
					Dim tmpLevy 'broker commission
						
					tmpLevy = commissionAmount
					levyAmount = CCur((agentCommission/100.00) * tmpLevy)
					sqlStr = "INSERT INTO [LevyContract] (LevyContract_DPA_,Contract_DPA_,LevyName,LevyAmount,LevyRate,LevyBlock,LevyShortName,LevyRatePercentage,SystemMaintained) SELECT " & " " & "iif(isnull(max([LevyContract_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'LevyContract'),max([LevyContract_DPA_]) + 1)" & " " & " as LevyContract_DPA_" & _
							"       ," & " " & rs.Fields("Contract_DPA_") & " " & " as Contract_DPA_" & _
							"       ," & "'Agent Commission'" & " as LevyName" & _
							"       ," & " " & RoundPoint05(levyAmount) & " " & " as LevyAmount" & _
							"       ," & " " & agentCommission * (commission/100.00) & " " & " as LevyRate" & _
							"       ," & " " & 0 & " " & " as LevyBlock" & _
							"       ," & " 'Agent' " & " as LevyShortName" & _
							"       ," & " '" & agentCommission & "%" & "' " & " as LevyRatePercentage" & _
							"       ,12 " & " as SystemMaintained" & _	
							"        FROM [LevyContract]"
					conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
					
					'Run stored Procedure to apply correct commissions
					'conn.execute("UpdateCompoundedContractCommissions")
					
					'process payment
					Dim pDate
					Dim payRS
					Dim ContractsSel
					Dim receiptVoucher
					Dim amount
					Dim bank
					Dim entity
					Dim account
					Dim custOrder
					Dim clientVoucher
					
					pDate = CDSTradRS.Fields("SettlementDate") & " " & Time
					ContractsSel = rs.Fields("Contract_DPA_")
					amount = ccur(CDSTradRS.Fields("SettlementAmount"))
					bank = 4 'CDS clearing account
					entity = 8 'Broker
					account = 14 'CDS system account
					custOrder = "Null"
					clientVoucher = "Null"
					
					if orderIsSaleType then														
							if(Cint(custodian)=1 or Cint(interbank) =1) then 'Custodian transaction done by stivo							
								
								if(Cint(custodian)=1 and Cint(interbank) <>1) then ''//Inter bank takes priority over Custodian
									set guid = server.createobject("NDUtils.CGUID")
									InterTransfer_EIT_ = guid.GenerateGUID											

									sqlStr = "INSERT INTO [InterTransfer] (SourceEntityType_DPA_,InterTransfer_EIT_,SourceEntity_DPA_" & _
											 "       ,TargetEntityType_DPA_,TargetEntity_DPA_,TransferAmount,TransferDate,ChangedBy,InterTransferType_DPA_,Contract_DPA_)" & _
											 " Values( " & " " & 8 & " " & " " & _
											 "       ," & "'" & InterTransfer_EIT_ & "'" & " " & _
											 "       ," & " " & 14 & " " & " ," & " " & cliententity & " " & " " & _
											 "		," & " " & client & " " & " " & _
											 "		," & " " & ccur(amount) & " " & " " & _
											 "		," & "#" & FormatDate(SettlementDate) & "#" & " " & _
											 "		," & " " & UserId & " " & " " & _												 
											 "		," & " " & 2 & " " & " " & _				
											 "       ," & " " & ContractDPA & " " & ")"
     
									sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
									conn.Execute sqlStr
								end if
							else
							
							'do a receipt (CDS system account (broker) -> CDS clearing account (bank)
							guidStr = guid.GenerateGUID
							sqlStr = "INSERT INTO [BrokerReceiptVoucher] (BrokerReceiptVoucher_DPA_, VoucherDate, BrokerReceiptVoucher_EIT_) SELECT " & " " & "iif(isnull(max([BrokerReceiptVoucher_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'BrokerReceiptVoucher'),max([BrokerReceiptVoucher_DPA_]) + 1)" & " " & " as BrokerReceiptVoucher_DPA_" & _
									"       ," & "#" & FormatDate(pDate) & "#" & " as VoucherDate" & _
									"       ," & "'" & guidStr & "'" & " as BrokerReceiptVoucher_EIT_" & _
									"        FROM [BrokerReceiptVoucher]"
							conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
											
							'obtain voucher number
							sqlStr = "SELECT [BrokerReceiptVoucher_DPA_] FROM [BrokerReceiptVoucher] WHERE [BrokerReceiptVoucher_EIT_] = " & "'" & guidStr & "'"
				
							Set payRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
							If (payRS.EOF Or payRS.BOF) Then%>
							         	<script language = 'vbscript'>
							         			ShowMessage "A serious error has been encountered while saving the data. Try saving again"
											         			
							         	</script>
							         	<% response.end
							End If
							receiptVoucher = payRS.fields("BrokerReceiptVoucher_DPA_").value
											
							'add contracts to voucher
							sqlStr = "UPDATE Contract SET BrokerReceiptVoucher_DPA_ = " & receiptVoucher & _
									", BrokerReceiptVouchered = 1 WHERE Contract_DPA_ IN (" & ContractsSel &  ")"
				
							conn.Execute SQLServerFormat(HandleQuote(sqlStr))
							
							guidStr = guid.GenerateGUID
							sqlStr = "INSERT INTO [Payment] ( PaymentAmount, PaymentPDate, BankAccount_DPA_, " & _
									"Payment_EIT_,PayType_DPA_, EntityType_DPA_, Payment_DPA_, PaymentReference, " & _
									"PaymentNarrative, Entity_DPA_, Order_DPA_,BrokerReceiptVoucher_DPA_,Contract_DPA_,ChangedBy) " & _
									"SELECT " & " " & ccur(amount) & " " & " as PaymentAmount" & _
									"," & "#" & FormatDate(settlementDate) & "#" & " as PaymentPDate" & _
									"," & " " & bank & " " & " as BankAccount_DPA_" & _
									"," & "'" & guidStr & "'" & " as Payment_EIT_" & "," & "1 as PayType_DPA_" & _
									"," & " " & entity & " " & " as EntityType_DPA_" & _
									"," & " " & "iif(isnull(max([Payment_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Payment'),max([Payment_DPA_]) + 1)" & " " & " as Payment_DPA_" & _
									"," & "'" & reference & "'" & " as PaymentReference" & _
									"," & "'" & narrative & "'" & " as PaymentNarrative" & _
									"," & " " & account & " " & " as Entity_DPA_" & _
									"," & " " & custOrder & " " & " as Order_DPA_" & _
									"," & " " & receiptVoucher & " " & " as BrokerReceiptVoucher_DPA_" & _
									"," & " " & ContractDPA & " " & " as Contract_DPA_" & _
									"," & " " & UserId & " " & " as ChangedBy " & _																		
									" FROM [Payment]"
							
							sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))

							conn.Execute sqlStr
							
							end if		
								
					else
									if(Cint(custodian)=1 OR Cint(interbank)=1) then 'Custodian transaction done by stivo
										
										if(Cint(custodian)=1 and Cint(interbank) <>1) then ''//Inter bank takes priority over Custodian

											set guid = server.createobject("NDUtils.CGUID")
											InterTransfer_EIT_ = guid.GenerateGUID
										
											sqlStr = "INSERT INTO [InterTransfer] (SourceEntityType_DPA_,InterTransfer_EIT_,SourceEntity_DPA_" & _
												 	",TargetEntityType_DPA_,TargetEntity_DPA_,TransferAmount,TransferDate,ChangedBy,InterTransferType_DPA_,Contract_DPA_)" & _
												 	" Values( " & " " & cliententity & " " & " " & _
												 	"," & "'" & InterTransfer_EIT_ & "'" & " " & _
												 	"," & " " & client & " " & " ," & " " & 8 & " " & " " & _
												 	"," & " " & 14 & " " & " " & _
												 	"," & " " & ccur(amount) & " " & " " & _
												 	"," & "#" & FormatDate(SettlementDate) & "#" & " " & _
												 	"," & " " & UserId & " " & " " & _												 
												 	"," & " " & 2 & " " & " " & _				
												 	"," & " " & ContractDPA & " " & ")"
      
											sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))

											conn.Execute sqlStr
										end if
									else
									'do a payment (CDS clearing account (bank) -> CDS system account (broker)
									guidStr = guid.GenerateGUID
									sqlStr = "INSERT INTO [Voucher] (Voucher_DPA_, VoucherDate, Voucher_EIT_) SELECT " & " " & "iif(isnull(max([Voucher_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Voucher'),max([Voucher_DPA_]) + 1)" & " " & " as Voucher_DPA_" & _
											"       ," & "#" & FormatDate(PDate) & "#" & " as VoucherDate" & _
											"       ," & "'" & guidStr & "'" & " as Voucher_EIT_" & _
											"        FROM [Voucher]"
									conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
											
									'obtain voucher number
									sqlStr = "SELECT [Voucher_DPA_] FROM [Voucher] WHERE [Voucher_EIT_] = " & "'" & guidStr & "'"
				
									Set payRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
									If (payRS.EOF Or payRS.BOF) Then%>
							         			<script language = 'vbscript'>
							         					ShowMessage "A serious error has been encountered while saving the data. Try saving again"
											         			
							         			</script>
							         			<% response.end
									End If
									brokerVoucher = payRS.fields("Voucher_DPA_").value
											
									'add contracts to voucher
									sqlStr = "UPDATE Contract SET Voucher_DPA_ = " & brokerVoucher & _
											", ContractVouchered = 1 WHERE Contract_DPA_ IN (" & ContractsSel &  ")"
				
									conn.Execute SQLServerFormat(HandleQuote(sqlStr))
							
									sqlStr = "INSERT INTO [Payment] ( PaymentAmount, PaymentPDate, BankAccount_DPA_, " & _
											"PayType_DPA_, EntityType_DPA_, Payment_DPA_, PaymentReference, PaymentNarrative, " & _
											"Entity_DPA_, Voucher_DPA_,ClientVoucher_DPA_,Contract_DPA_,ChangedBy) " & _
											"SELECT " & " " & CCur(amount) & " " & " as PaymentAmount" & _
											"," & "#" & FormatDate(SettlementDate) & "#" & " as PaymentPDate" & _
											"," & " " & bank & " " & " as BankAccount_DPA_" & _
											"," & "2 as PayType_DPA_" & _
											"," & " " & entity & " " & " as EntityType_DPA_" & _
											"," & " " & "iif(isnull(max([Payment_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Payment'),max([Payment_DPA_]) + 1)" & " " & " as Payment_DPA_" & _
											"," & "'" & reference & "'" & " as PaymentReference" & _
											"," & "'" & narrative & "'" & " as PaymentNarrative" & _
											"," & " " & account & " " & " as Entity_DPA_" & _
											"," & " " & brokerVoucher & " " & " as Voucher_DPA_" & _
											"," & " " & clientVoucher & " " & " as ClientVoucher_DPA_" & _
											"," & " " & ContractDPA & " " & " as Contract_DPA_" & _
											"," & " " & UserId & " " & " as ChangedBy " & _
											" FROM [Payment]"
					
									sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))

									conn.Execute sqlStr							

											sqlStr = "UPDATE [Voucher] SET VoucherPaid = 1 WHERE Voucher_DPA_ = "  & brokerVoucher
											conn.Execute SQLServerFormat(HandleQuote(sqlStr))
									end if			
							
							
					end if
					
					'mark CDS trade as processed
					sqlStr = "UPDATE _CDS_Imported_Trades_ SET Processed = 1" & _
							" WHERE CDSImport_DPA_=" & ID
					conn.Execute SQLServerFormat(HandleQuote(sqlStr))
			conn.CommitTrans
			end if
			Next
			case "UPDATE"
			
			Ids=Split(ID,"<->")			
			OrderDetail_DPA_=Ids(1)
			ID=Ids(0)
			
			Commission = Request.Form("txtCommission")			
				
        'validate Description
        If Trim(Commission) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Commission"
                		
                </script>
                <% response.end
        End If
        'validate size of Description
        If Len(Commission) > 100 Then%>
                <script language = 'vbscript'>
                ShowMessage "Commission can only be 100 characters in length"
                
                </script>
                <% response.end
        End If
		
		'ensure Commission is numeric
				If (Not IsNumeric(Commission)) Then%>
				    <script language = 'vbscript'>
						ShowMessage "Commission must be numeric"
						
				    </script>
				    <% response.end
				End If

		Set conn = GetActiveConnection("KBroker")
        
        'save data
        
        sqlStr = "UPDATE [_CDS_Imported_Trades_] SET CommissionRate = " & "" & Commission & "" & " WHERE CDSImport_DPA_  = " & ID					

	   conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
        
        conn.Close
        Set conn = Nothing        
   		'end If				
    end select
	response.redirect "CDSMatchedTradesList.asp"
    	
%>