<!--#include file="../libroutines.asp"-->

	<%
	const UDLName = "KBroker"
	const DataSource = "EditLot"
	const DataEntity = "Lot"
	const DataEntityPlural = "Lots"
	const ActionFolder = "Operations"
	
	const LinkedIndependent = 1
    const LinkedDependent = 2
	
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim guidStr 
	Dim guid
	Dim ID
	Dim ItemID
	Dim rsEdit
	Dim IDHolder
	Dim IDArray
	
	UserId=Session("UserID")
	
	action = ucase(Request.Form("action"))
	IDHolder = Request("ID")
	
	'action = "EXECUTE_DETAIL"
	'IDHolder = "3320<->7171"
	
	'IDHolder = "1<->2"
	
	If (Trim(IDHolder) = "") or (IDHolder = "0") Then%>
			<script language = 'vbscript'>
                	ShowMessage "Please select an Order item for Lot allocation"
                	window.self.close
			</script>
			<%response.end
	End If
	
	IDArray = split(IDHolder,"<->")
	ID = IDArray(lbound(IDArray))
	
	'ItemID = IDArray(ubound(IDArray))
	ItemID = Request.Form("ItemID")
	
	'ItemID = -1
	
	if action <> "" then
			If Trim(ItemID) = "" Then%>
					<script language = 'vbscript'>
            				ShowMessage "No item specified"                				
					</script>
					<%response.end
			End If
	end if
	
	select case action 
		case "EXECUTE_DETAIL"

				
			Dim slip
			Dim broker
			Dim tDate
			Dim qty
			Dim price
			Dim orderType
			Dim orderIsSaleType
			Dim securityID
			Dim commission	
			Dim orderSecType
			Dim agentCommission
			Dim staffCommission
			Dim volComm
			Dim volBound
			Dim minComm
			Dim cma
			Dim imobRate
			Dim secImob
			Dim regularComm			
			Dim interbank
			Dim ContractDPA
			Dim custodian
			Dim client
			Dim clientEntity
						
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
			Dim ClientClass  
			Dim SettlementDate

			if itemID = "-1" then
				broker = Request.Form("cboBroker")
				tDate = Trim(Request.Form("txtTDate"))
				tDate = tDate & " " & Time
				slip = Request.Form("txtSlip")
				qty = Request.Form("txtQty")
				price = Request.Form("txtPrice")
				SettlementDate = Trim(Request.Form("txtSDate"))
				SettlementDate = SettlementDate & " " & Time
			else
				broker = Request.Form("cboBrokerInPlace")
				tDate = Trim(Request.Form("Date"))
				tDate = tDate & " " & Time
				slip = Request.Form("Slip")
				qty = Request.Form("Quantity")
				price = Request.Form("Price")
				SettlementDate = Request.Form("SettlementDate")
				SettlementDate = SettlementDate & " " & Time
			end if
						
			sDate = FormatDate(SettlementDate)
						
			orderType = Request.Form("txtOrderType")
			orderIsSaleType = cbool(Request.Form("txtOrderIsSaleType"))
			securityID = Request.Form("txtSecurityID")
			regularComm = Request.Form("txtCommission")
			orderSecType = Request.Form("txtInstrument")
			agentCommission = Request.Form("txtAgentCommission")
			staffCommission = Request.Form("txtStaffCommission")
			volComm = Request.Form("txtVolumeCommission")
			volBound = ccur(Request.Form("txtVolumeBoundary"))
			minComm = ccur(Request.Form("txtMinimumCommission"))
			cma = Request.Form("txtCMA")
			imobRate = Request.Form("txtPostImmobilisedRate")
			secImob = Request.Form("txtSecurityImmobilised")
			interbank=Cint(Request.Form("txtinterbank"))       
			custodian=Cint(Request.Form("txtcustodian"))       
			client =Request.Form("txtClientDPA")       
			cliententity =Cint(Request.Form("txtEntityDPA"))       
			clientClass =Cint(Request.Form("txtClass"))
						
			Set TimeLimitRs = CreateObject("ADODB.Recordset")   						        
			TimeLimitRs.CursorLocation = adUseClient
			   				
			Set conn = GetActiveConnection("KBroker")
			             
			'validate Broker
			If Trim(Broker) = "" Then%>
				<script language = 'vbscript'>
					ShowMessage "Please specify the Broker"
				</script>
				<% 
				
				response.end
			End If
			'validate Slip
			If Trim(Slip) = "" Then%>
				<script language = 'vbscript'>
					ShowMessage "Please specify the Ref No."
				</script>
				<% response.end
			End If
			'ensure Slip is numeric
			'If (Not IsNumeric(Slip)) Then%>
				<script language = 'vbscript'>
					'ShowMessage "Slip No. must be numeric"
				</script>
				<% 'response.end
			'End If
			'validate Estimated Price
			If Trim(Price) = "" Then%>
				<script language = 'vbscript'>
					ShowMessage "Please specify the Price "
				</script>
			<% response.end
			End If
			'validate Estimated Quantity
			If Trim(qty)= "" Then%>
				<script language = 'vbscript'>
					ShowMessage "Please specify the Quantity "
				</script>
				<% response.end
			End If
			'ensure Order Detail Estimated Price is numeric
			If (Price <> "") And (Not IsNumeric(Price)) Then%>
				<script language = 'vbscript'>
					ShowMessage "Price must be a number"
				</script>
				<% response.end
			End If
			'ensure Order Detail Estimated Quantity is numeric
			If (Qty <> "") And (Not IsNumeric(Qty)) Then%>
				<script language = 'vbscript'>
					ShowMessage "Quantity must be a number"
				</script>
				<% response.end
			End If
			'ensure date is valid format
			If Not IsDate(tDate) Then%>
				<script language = 'vbscript'>
					ShowMessage "Date must be a valid date"					
				</script>
				<% response.end
			End If

			'check for necessary settings before proceeding
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
				
			settlementDate=LTdate(CDate(tDate),NoOfDays)
			  
			'calculate amounts
			Dim grossAmount  'this is the amount before application of Levies
			Dim levyRS
			Dim blocks
						
			'calculate amounts						
			if orderSecType = "Fixed" then ' "F" is FIXED security
				grossAmount = (price * qty) / 100
			else
				grossAmount = price * qty
			end if		
						 
			conn.BeginTrans
						
				if itemID = "-1" then
					'Add Operation: save contract
					set guid = server.createobject("NDUtils.CGUID")
					guidStr = guid.GenerateGUID

					
					sqlStr = "INSERT INTO [Contract] (Contract_DPA_, Contract_EIT_, Status_DPA_,IsInterBank,ContractSettlementDate) " & _
							 "SELECT " & " " & "iif(isnull(max([Contract_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Contract'),max([Contract_DPA_]) + 1)" & " " & " as Contract_DPA_" & _
							 "," & "'" & guidStr & "'" & " as Contract_EIT_" & _
							 "," & " " & 1 & " " & " as Status_DPA_" & _
							 "," & " " & interbank & " " & " as IsInterBank" & _
							 "," & "#" & FormatDate(sDate) & "#" & " as ContractSettlementDate" & _
							 " FROM [Contract]"

					conn.Execute("SET IDENTITY_INSERT [Contract] ON")
					conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
					conn.Execute("SET IDENTITY_INSERT [Contract] OFF")
					
					'obtain contract number
					sqlStr = "SELECT Contract_DPA_ FROM [Contract] WHERE [Contract.Contract_EIT_] = " & "'" & guidStr & "'"
					Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
										
					If (rs.EOF Or rs.BOF) Then
						%>
						<script language = 'vbscript'>
								ShowMessage "A serious error has been encountered while saving the data. Try saving again"
					    </script>
						<%
						response.end
					End If
													
					ContractDPA = rs("Contract_DPA_")
										
					CDSFlag = 0

					conn.Execute("SET IDENTITY_INSERT [Lot] ON")										
					sqlStr = "INSERT INTO [Lot] (Lot_DPA_,Contract_DPA_,OrdDetail_DPA_,LotPrice" & _
							",LotQty,LotSlipNo,LotTDate,Broker_DPA_,ContractNumber, LotGrossAmount,CDSTransaction,ChangedBy)" & _
							" SELECT " & " " & "iif(isnull(max([Lot_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Lot'),max([Lot_DPA_]) + 1)" & " " & " as Lot_DPA_" & _
							"," & " " & rs.fields("Contract_DPA_") & " " & " as Contract_DPA_" & _
							"," & " " & ID & " " & " as OrdDetail_DPA_" & _
							"," & " " & price & " " & " as LotPrice" & _
							"," & " " & qty & " " & " as LotQty" & _
							"," & " '" & slip & "' " & " as LotSlipNo" & _
							"," & "#" & FormatDate(tDate) & "#" & " as LotTDate" & _
							"," & " " & broker & " " & " as Broker_DPA_" & _
							"," & "'" & left(orderType,1) & rs.fields("Contract_DPA_") & "'" & " as ContractNumber" & _								
							"," & " " & grossAmount & " " & " as LotGrossAmount" & _
							"," & " " & CDSFlag & " " & " as CDSTransaction" & _
							"," & " " & UserId & " " & " as ChangedBy" & _
							" FROM [Lot]"
					sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))

					conn.Execute sqlStr
					conn.Execute("SET IDENTITY_INSERT [Lot] OFF")									
					'apply relevant levies
					Dim cond
									
					if orderSecType = "Fixed" then 
						cond = " WHERE LevyAppBond = 1 AND LevyActive = 1"
					else
						cond = " WHERE (LevyAppSecurity = 1) AND (LevyActive = 1) AND (" & _
							" Levy_DPA_ IN (SELECT Levy_DPA_ FROM LevySecurity " & _
							" WHERE Security_DPA_ = " & securityID & "))"
					end if
									
					sqlStr = "SELECT * FROM [LevyList] " & cond
					Set levyRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
										
					Dim levyAmount
										
					If Not (levyRS.EOF Or levyRS.BOF) Then
						levyRS.MoveFirst
						Do Until levyRS.EOF
							if levyRS.Fields("LevyType") = "P" Then
								levyAmount = CCur((levyRS.Fields("LevyAmount")/100.00) * grossAmount)
								LevyRatePercentage = levyRS.Fields("LevyAmount") & "%"
							else
								blocks = Int(grossAmount / levyRS.Fields("LevyBlock"))
								if (grossAmount mod levyRS.Fields("LevyBlock")) <> 0 then
									blocks = blocks + 1
								end if
								levyAmount = CCur(blocks * levyRS.Fields("LevyAmount"))
								LevyRatePercentage = levyRS.Fields("LevyAmount") & " for every " & levyRS.Fields("LevyBlock")
							end if
							
							conn.Execute("SET IDENTITY_INSERT [LevyContract] ON")

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

							conn.Execute("SET IDENTITY_INSERT [LevyContract] OFF")
							levyRS.MoveNext
						Loop
					end if
										
					'consider the transfer fee
					if Not(orderIsSaleType) then
						Dim transRS
						sqlStr = "SELECT * FROM [SecTransFeeListLatest] WHERE Security_DPA_ = " & securityID
												
												
						Set transRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
						conn.Execute("SET IDENTITY_INSERT [LevyContract] ON")

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
						conn.Execute("SET IDENTITY_INSERT [LevyContract] OFF")
					end if
					
					pDate = settlementDate
					ContractsSel = rs.Fields("Contract_DPA_")
					amount = grossAmount
					bank = 4 'CDS clearing account
					entity = 8 'Broker
					account = 14 'CDS system account
					custOrder = "Null"
					clientVoucher = "Null"
													
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
							If (commissionRS.EOF Or commissionRS.BOF) Then
								%>
							 	<script language = 'vbscript'>
							 			ShowMessage "The assigned commission rate is inviolation of CMA regulations"
							 	</script>
							 	<%
							 	response.end
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
										
					if commission = "" then
						%>
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
					
					conn.Execute("SET IDENTITY_INSERT [LevyContract] ON")
							
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
					conn.Execute("SET IDENTITY_INSERT [LevyContract] OFF")

					'GET Botswana Stock Exchange Fee
					Dim objBSE
					sqlStr = "SELECT CommissionRate, SecurityBoundary, UpperSecurityCommission, Vatable FROM Commission WHERE CommissionDescription = 'Botswana Stock Exchange Fee'"
					Set objBSE = conn.Execute(sqlStr)

					If Not objBSE.EOF then
						dblBSEComm = objBSE("CommissionRate")
						dblBSEBoundary = objBSE("SecurityBoundary")
						dblBSEUpperComm = objBSE("UpperSecurityCommission")
						bolVatable = objBSE("Vatable")

						If grossAmount < dblBSEBoundary then dblBSEComm = dblBSEUpperComm

						dblBSECommission = CCur((dblBSEComm/100.00) * grossAmount)

						If bolVatable then vatamount = vatamount + dblBSECommission
						UpdatedContractDPA = rs.Fields("Contract_DPA_")
						conn.Execute("SET IDENTITY_INSERT [LevyContract] ON")
						sqlStr = "INSERT INTO [LevyContract] (LevyContract_DPA_,Contract_DPA_,LevyName,LevyAmount,LevyRate,LevyBlock,LevyShortName,LevyRatePercentage,SystemMaintained) SELECT " & " " & "iif(isnull(max([LevyContract_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'LevyContract'),max([LevyContract_DPA_]) + 1)" & " " & " as LevyContract_DPA_" & _
								"       ," & " " & rs.Fields("Contract_DPA_") & " " & " as Contract_DPA_" & _
								"       ," & "'Botswana Stock Exchange Commission'" & " as LevyName" & _
								"       ," & " " & RoundPoint05(dblBSECommission) & " " & " as LevyAmount" & _
								"       ," & " " & dblBSEComm & " " & " as LevyRate" & _
								"       ," & " " & 0 & " " & " as LevyBlock" & _
								"       ," & " 'BSEComm' " & " as LevyShortName" & _
								"       ," & " '" & dblBSEComm & "%" & "' " & " as LevyRatePercentage" & _
								"       ,25 " & " as SystemMaintained" & _	
								"        FROM [LevyContract]"

						sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
						conn.Execute sqlStr
						conn.Execute("SET IDENTITY_INSERT [LevyContract] OFF")

					End If
					Set objBSE = Nothing

							
					'Apply agent commission
					Dim tmpLevy 'broker commission
									
					tmpLevy = commissionAmount
					levyAmount = CCur((agentCommission/100.00) * tmpLevy)
					conn.Execute("SET IDENTITY_INSERT [LevyContract] ON")
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
					conn.Execute("SET IDENTITY_INSERT [LevyContract] OFF")									
					''Security Is Not Immobilised
					''Do not post to CDS Account

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
								 	"," & "#" & FormatDate(settlementDate) & "#" & " " & _
								 	"," & " " & UserId & " " & " " & _												 
								 	"," & " " & 2 & " " & " " & _				
								 	"," & " " & ContractDPA & " " & ")"
			      
							sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))

							conn.Execute sqlStr
						end if
					else
						'do a payment (CDS clearing account (bank) -> CDS system account (broker)
						set guid = server.createobject("NDUtils.CGUID")	
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
							"," & "#" & FormatDate(settlementDate) & "#" & " as PaymentPDate" & _
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
							"," & " " & UserId & " " & " as ChangedBy" & _
							" FROM [Payment]"
								
						sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))

						conn.Execute sqlStr										
																
						sqlStr = "UPDATE [Voucher] SET VoucherPaid = 1 WHERE Voucher_DPA_ = "  & brokerVoucher
						conn.Execute SQLServerFormat(HandleQuote(sqlStr))
					end if		


					If Len(UpdatedContractDPA & "") > 0 then
						'Calculate and Enter VAT *****************************************
						sqlStr = "SELECT SUM(LevyAmount) FROM LevyContractList WHERE LevyShortName <> 'VAT' and      (SystemMaintained <> 12) AND Contract_DPA_= " & UpdatedContractDPA 
						dblVatableAmount = conn.execute(sqlStr)(0)

						sqlStr = "SELECT LevyRate FROM LevyContractList WHERE LevyShortName = 'VAT' AND Contract_DPA_= " & UpdatedContractDPA
						dblVATRate = conn.execute(sqlStr)(0)

						dblVATAmount=(dblVatableAmount*(dblVATRate/100))

						sqlStr = "UPDATE LevyContractList SET LevyAmount = " & dblVATAmount & " WHERE LevyShortName = 'VAT' AND Contract_DPA_= " & UpdatedContractDPA
						conn.execute(sqlStr)
					End If					
										
					''RECALCULATE THE COMMISSIONS
					'Client Contracts
					sqlStr = "UPDATE LevyContract" & _
					" SET LevyAmount = ProperCommission3" & _
					" FROM (SELECT ShowProperCommissionForCompoundedContractsForEachLevyEntry.LotTDate, " & _
					" ShowProperCommissionForCompoundedContractsForEachLevyEntry.ProperCommissionRate, " & _
					" ShowProperCommissionForCompoundedContractsForEachLevyEntry.ProperCommission3, LevyContract.LevyContract_DPA_" & _
					" FROM ShowProperCommissionForCompoundedContractsForEachLevyEntry INNER JOIN" & _
					" (SELECT cast('"& FormatDate(tDate) &"' AS datetime) AS LastDate) LastDateTransactions ON " & _
					" ShowProperCommissionForCompoundedContractsForEachLevyEntry.LotTDate = LastDateTransactions.LastDate INNER JOIN" & _
					" LevyContract ON " & _
					" ShowProperCommissionForCompoundedContractsForEachLevyEntry.LevyContract_DPA_ = LevyContract.LevyContract_DPA_" & _
					" WHERE (LevyContract.Contract_DPA_ = "& ContractDPA &")) A" & _
					" WHERE LevyContract.LevyContract_DPA_ = A.LevyContract_DPA_"
												
					conn.Execute(sqlStr)
				else
					'edit detail data
					sqlStr = "UPDATE Lot SET ContractSettlementDate = #" & FormatDate(sDate) & "#," & _
						" LotPrice = " & price & "," & _
						" LotQty = " & qty & ", LotSlipNo = '" & slip & "'," & _
						" LotTDate = #" & FormatDate(tDate) & "#," & _
						" Broker_DPA_ = " & broker & _
						", LotGrossAmount = " & grossAmount & _
						", ChangedBy = " & UserId & _
						", TimeChanged = GetDate()" & _
						" WHERE Lot_DPA_=" & itemID
					conn.Execute SQLServerFormat(HandleQuote(sqlStr))
										
					sqlStr1 = "SELECT * FROM OrdDetailList WHERE OrdDetailList.OrdDetail_DPA_= " & ID
										
					'retrieve contract number
					sqlStr = "SELECT Distinct tbOrder.OrderType_DPA_, Contract.IsInterBank, Lot.Lot_DPA_, Lot.Contract_DPA_, Lot.CDSTransaction " & _
							  " FROM Lot INNER JOIN Contract ON Lot.Contract_DPA_ = Contract.Contract_DPA_ INNER JOIN OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN " & _
			                  " tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ WHERE (Contract.Deleted = 0) AND (tbOrder.Deleted = 0) AND (OrdDetail.Deleted = 0) AND (Lot.Deleted = 0) AND (Lot.Lot_DPA_ = " & itemID & ")"
										
					Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))							
										
					If (rs.EOF Or rs.BOF) Then
						%>
						<script language = 'vbscript'>
							ShowMessage "A serious error has been encountered while saving the data. Try saving again"
						</script>
						<% response.end
					End If
										
					Dim transIsCDS
					Dim dblVATAmount
					
					dblVATAmount = 0
					transIsCDS = cbool(rs.fields("CDSTransaction").value)
					ContractDPA=rs("Contract_DPA_")
					orderTypeDPA=Cint(rs("OrderType_DPA_"))

					if(cbool(rs.fields("IsInterBank").value)) then
						interbanktransfer=1
					else
						interbanktransfer=0
					end if
										
					'update levies
					'Apply broker commission
					'Dim commissionRS
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
							If (commissionRS.EOF Or commissionRS.BOF) Then
								%>
							 	<script language = 'vbscript'>
							 			ShowMessage "The assigned commission rate is in violation of CMA regulations"
															 					
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
												
					commissionAmount = CCur((commission/100.00) * grossAmount)
																	
					if commissionAmount < minComm then
						commissionAmount = minComm
					end if
														
					levyAmount = commissionAmount
					
					sqlStr = "SELECT * FROM LevyContractList WHERE LevyShortName <> 'Handling' AND LevyShortName <> 'VAT' AND Contract_DPA_= " & rs.fields("Contract_DPA_")
							
					Set levyRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					If (levyRS.EOF Or levyRS.BOF) Then
						%>
						<script language = 'vbscript'>
								ShowMessage "The levies for the selected contract are missing. This is a serious database corruption."
						</script>
						<% response.end
					End If
										
					brokerAgentcommission = commissionAmount 'Set broker commission to be used in calculating levies

					dblBSECommission = 0
					'GET Botswana Stock Exchange Fee
					sqlStr = "SELECT CommissionRate, SecurityBoundary, UpperSecurityCommission, Vatable FROM Commission WHERE CommissionDescription = 'Botswana Stock Exchange Fee'"
					Set objBSE = conn.Execute(sqlStr)
					If Not objBSE.EOF then
						dblBSEComm = objBSE("CommissionRate")
						dblBSEBoundary = objBSE("SecurityBoundary")
						dblBSEUpperComm = objBSE("UpperSecurityCommission")
						bolVatable = objBSE("Vatable")

						If grossAmount < dblBSEBoundary then dblBSEComm = dblBSEUpperComm

						dblBSECommission = CCur((dblBSEComm/100.00) * grossAmount)
					
					End If
					Set objBSE = Nothing
									
					do until levyRS.eof
						if levyRS.Fields("LevyBlock") <> 0 then
								'block type levy
								blocks = Int(grossAmount / levyRS.Fields("LevyBlock"))
								if (grossAmount mod levyRS.Fields("LevyBlock")) <> 0 then
									blocks = blocks + 1
								end if
								levyAmount = CCur(blocks * levyRS.Fields("LevyRate"))
						elseif levyRS.Fields("LevyName") = "Transfer Fee" then
								'transfer fee
								levyAmount = levyRS.Fields("LevyRate")
						else
								'broker commission or other percentage-based commission
								levyAmount = CCur((levyRS.Fields("LevyRate")/100.00) * grossAmount)
						end if
							         				
						sqlStr = "UPDATE LevyContract SET LevyAmount = " & RoundPoint05(levyAmount) & _
								  " WHERE LevyContract_DPA_ = " & levyRS.Fields("LevyContract_DPA_")
												
						if (Cint(trim(levyRS("SystemMaintained")))= 11) then ' Broker Commission
							sqlStr = " UPDATE LevyContract SET LevyAmount = " & RoundPoint05(brokerAgentcommission) & _
								     " ,LevyRate=" & commission & ",LevyRatePercentage='" & commission & "%" & "'  WHERE LevyContract_DPA_ = " & levyRS.Fields("LevyContract_DPA_")						
						end if

						if (Cint(trim(levyRS("SystemMaintained")))= 25) then ' BSE Commission
							sqlStr = " UPDATE LevyContract SET LevyAmount = " & RoundPoint05(dblBSECommission) & _
								     " ,LevyRate=" & dblBSEComm & ",LevyRatePercentage='" & dblBSEComm & "%" & "'  WHERE LevyContract_DPA_ = " & levyRS.Fields("LevyContract_DPA_")						
						end if

												
						if(Cint(levyRS("SystemMaintained"))=12) then' Agent Commission								
							AgentcommissionAmount = CCur((agentCommission/100.00) * brokerAgentcommission)		
			                
							sqlStr = "UPDATE LevyContract SET LevyAmount = " & RoundPoint05(AgentcommissionAmount) & _
								   " ,LevyRate=" & agentCommission * (commission/100.00) & ",LevyRatePercentage='" & agentCommission & "%" & "' WHERE LevyContract_DPA_ = " & levyRS.Fields("LevyContract_DPA_")							
						end if	
																				
						conn.execute(sqlStr)
												
						levyRS.movenext
					loop

					UpdatedContractDPA = rs.fields("Contract_DPA_")
					'Calculate and Enter VAT *****************************************
					sqlStr = "SELECT SUM(LevyAmount) FROM LevyContractList WHERE LevyShortName <> 'VAT' AND Contract_DPA_= " & rs.fields("Contract_DPA_")
					dblVatableAmount = conn.execute(sqlStr)(0)

					sqlStr = "SELECT LevyRate FROM LevyContractList WHERE LevyShortName = 'VAT' AND Contract_DPA_= " & rs.fields("Contract_DPA_")
					dblVATRate = conn.execute(sqlStr)(0)

					dblVATAmount=(dblVatableAmount*(dblVATRate/100))

					
					sqlStr = "UPDATE LevyContractList SET LevyAmount = " & dblVATAmount & " WHERE LevyShortName = 'VAT' AND Contract_DPA_= " & rs.fields("Contract_DPA_")
					conn.execute(sqlStr)

					'do a payment (CDS clearing account (bank) -> CDS system account (broker)*******************
					set guid = server.createobject("NDUtils.CGUID")	
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
					
					pDate = settlementDate
					ContractsSel = rs.Fields("Contract_DPA_")
					amount = grossAmount
					bank = 4 'CDS clearing account
					entity = 8 'Broker
					account = 14 'CDS system account
					custOrder = "Null"
					clientVoucher = "Null"
					
					'add contracts to voucher
					sqlStr = "UPDATE Contract SET Voucher_DPA_ = " & brokerVoucher & _
							", ContractVouchered = 1 WHERE Contract_DPA_ IN (" & ContractsSel &  ")"
							
					conn.Execute SQLServerFormat(HandleQuote(sqlStr))
															
					sqlStr = "INSERT INTO [Payment] ( PaymentAmount, PaymentPDate, BankAccount_DPA_, " & _
							"PayType_DPA_, EntityType_DPA_, Payment_DPA_, PaymentReference, PaymentNarrative, " & _
							"Entity_DPA_, Voucher_DPA_,ClientVoucher_DPA_,Contract_DPA_,ChangedBy) " & _
							"SELECT " & " " & CCur(amount) & " " & " as PaymentAmount" & _
							"," & "#" & FormatDate(settlementDate) & "#" & " as PaymentPDate" & _
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
							"," & " " & UserId & " " & " as ChangedBy" & _
							" FROM [Payment]"
						'Response.Write sqlstr
						'Response.End 
								
					sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))

					conn.Execute sqlStr										
															
					sqlStr = "UPDATE [Voucher] SET VoucherPaid = 1 WHERE Voucher_DPA_ = "  & brokerVoucher
					conn.Execute SQLServerFormat(HandleQuote(sqlStr))
				
					If (custodian <>0 Or interbank=1) Then	' is interbank or Custodian															
						sqlStr="Update Payment Set Deleted=1,ChangedBy=" & UserId & ",TimeChanged=GetDate() where Contract_DPA_=" & ContractDPA 
						Conn.execute(sqlStr)
																										
						if(custodian <>0 and interbank=0) then
							sqlStr="Select * From InterTransfer where (Contract_DPA_=" & ContractDPA & ") and InterTransferType_DPA_=2"									

							Set commissionRS =Conn.Execute(sqlStr)										
														         	
							if Not(commissionRS.eof or commissionRS.Bof)	then
								Conn.Execute("Update InterTransfer Set Deleted=0 Where Contract_DPA_=" & ContractDPA & " and InterTransferType_DPA_=2")
							else
								if orderIsSaleType then												
									set guid = server.createobject("NDUtils.CGUID")	
									InterTransfer_EIT_ = guid.GenerateGUID
														
									sqlStr = "INSERT INTO [InterTransfer] (SourceEntityType_DPA_,InterTransfer_EIT_,SourceEntity_DPA_" & _
									 				 "       ,TargetEntityType_DPA_,TargetEntity_DPA_,TransferAmount,TransferDate,ChangedBy,InterTransferType_DPA_,Contract_DPA_)" & _
									 				 " Values( " & " " & 8 & " " & " " & _
									 				 "       ," & "'" & InterTransfer_EIT_ & "'" & " " & _
									 				 "       ," & " " & 14 & " " & " ," & " " & cliententity & " " & " " & _
									 				 "		," & " " & client & " " & " " & _
									 				 "		," & " " & ccur(grossAmount) & " " & " " & _
									 				 "		," & "#" & FormatDate(settlementDate) & "#" & " " & _
									 				 "		," & " " & UserId & " " & " " & _												 
									 				 "		," & " " & 2 & " " & " " & _				
									 			     "       ," & " " & ContractDPA & " " & ")"
			      
									sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))

									conn.Execute sqlStr												
								else
									set guid = server.createobject("NDUtils.CGUID")
									InterTransfer_EIT_ = guid.GenerateGUID

									sqlStr = "INSERT INTO [InterTransfer] (SourceEntityType_DPA_,InterTransfer_EIT_,SourceEntity_DPA_" & _
										 	",TargetEntityType_DPA_,TargetEntity_DPA_,TransferAmount,TransferDate,ChangedBy,InterTransferType_DPA_,Contract_DPA_)" & _
									 		" Values( " & " " & cliententity & " " & " " & _
									 		"," & "'" & InterTransfer_EIT_ & "'" & " " & _
									 		"," & " " & client & " " & " ," & " " & 8 & " " & " " & _
									 		"," & " " & 14 & " " & " " & _
									 		"," & " " & ccur(grossAmount) & " " & " " & _
									 		"," & "#" & FormatDate(settlementDate) & "#" & " " & _
									 		"," & " " & UserId & " " & " " & _												 
									 		"," & " " & 2 & " " & " " & _				
									 		"," & " " & ContractDPA & " " & ")"
			      
											sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))

											conn.Execute sqlStr
								end if
							end if

							sqlStr = "UPDATE InterTransfer SET " & _
									 "  TransferAmount = " & grossAmount & _
									 ", TransferDate = '" & FormatDate(settlementDate) & "'" & _
									 ", ChangedBy = " & UserId  & _
									 ", TimeChanged = GetDate() " & _
									 " WHERE Contract_DPA_= " & ContractDPA
											
							conn.Execute SQLServerFormat(HandleQuote(sqlStr))									
						end if
										
						Conn.Execute("Update Contract Set IsInterBank=" & interbank & " Where Contract_DPA_="& ContractDPA)
					end if
										
					UserId=Session("UserID")
					sqlStr = "Update Contracts Set ModifiedBy=" & UserId & ", DateModified=GetDate(), ContractSettlementDate=#" & FormatDate(sDate) & "# WHERE Contract_DPA_ = " & ContractDPA
					conn.Execute SQLServerFormat(HandleQuote(sqlStr))
														
					''RECALCULATE THE COMMISSIONS
					'Client Contracts
					sqlStr = "UPDATE LevyContract" & _
					" SET LevyAmount = ProperCommission3" & _
					" FROM (SELECT ShowProperCommissionForCompoundedContractsForEachLevyEntry.LotTDate, " & _
					" ShowProperCommissionForCompoundedContractsForEachLevyEntry.ProperCommissionRate, " & _
					" ShowProperCommissionForCompoundedContractsForEachLevyEntry.ProperCommission3, LevyContract.LevyContract_DPA_" & _
					" FROM ShowProperCommissionForCompoundedContractsForEachLevyEntry INNER JOIN" & _
					" (SELECT cast('"& FormatDate(tDate) &"' AS datetime) AS LastDate) LastDateTransactions ON " & _
					" ShowProperCommissionForCompoundedContractsForEachLevyEntry.LotTDate = LastDateTransactions.LastDate INNER JOIN" & _
					" LevyContract ON " & _
					" ShowProperCommissionForCompoundedContractsForEachLevyEntry.LevyContract_DPA_ = LevyContract.LevyContract_DPA_" & _
					" WHERE (LevyContract.Contract_DPA_ = "& ContractDPA &")) A" & _
					" WHERE LevyContract.LevyContract_DPA_ = A.LevyContract_DPA_"
															
					conn.Execute(sqlStr)
				end if 

			conn.CommitTrans    
			conn.Close
			Set conn = Nothing
			%>
		<SCRIPT LANGUAGE="JAVASCRIPT">					
				this.location="EditLotItem.asp?ID=" +<%= ID%>;
		</SCRIPT>
		<%
			'WriteDialogRelocateScript "EditLotItem.asp?ID=" & IDHolder
			Response.End
			
		case "EXECUTE_DELETE"
			Set conn = GetActiveConnection("KBroker")
			
			'obtain levies, contract and lot to be deleted
			sqlStr = "SELECT * FROM LevyContractList WHERE Lot_DPA_=" & ItemID
				
			Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			If (rs.EOF Or rs.BOF) Then%>
						<script language = 'vbscript'>
				         		ShowMessage "No Levies were found for this Lot. This is a serious database corruption."
								
						</script>
						<% response.end
			End If
			
			contractID = rs("Contract_DPA_")
			lotID = rs("Lot_DPA_")
					
			conn.BeginTrans
					
					'delete levies
					do until rs.eof
							DeleteItem "LevyContract","Contract_DPA_",contractID
							rs.movenext
					loop							

					sqlStr = "SELECT * FROM ContractList WHERE Contract_DPA_=" & contractID
   					set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
   					If not(rs.EOF Or rs.BOF) Then
   							if (rs.Fields("ContractDelivered") or rs.Fields("ContractNCDelivered")) then%>
									<script language = 'vbscript'>
                							window.self.ShowMessage "The selected Lot has been delivered and cannot be deleted"
						                	
									</script>
									<% response.end
   							end if
					End If
					
					transIsCDS = cbool(rs.fields("CDSTransaction"))					
									
					'delete payment
					if transIsCDS then
							if(rs("IsInterBank")=false) and IsNull(rs("Intertransfer")) then														
							sqlStr = "SELECT Payment_DPA_,BrokerReceiptVoucher_DPA_,Voucher_DPA_ FROM Payment WHERE Contract_DPA_=" & contractID
							Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
							If (rs.EOF Or rs.BOF) Then%>
					        			<script language = 'vbscript'>
					        					ShowMessage "A serious error has been encountered while saving the data. Try saving again"
												         			
					        			</script>
					        			<% response.end
							End If
							
							Dim paymentID
							Dim brokerVoucherID
							Dim VoucherID
							
							paymentID = rs.fields("Payment_DPA_")
							
							DeleteItem "Payment","Payment_DPA_",paymentID
			
							sqlStr = "UPDATE Contract SET BrokerReceiptVoucher_DPA_ = NULL" & _
									", BrokerReceiptVouchered = 0 WHERE Contract_DPA_ =" & contractID
							conn.Execute SQLServerFormat(HandleQuote(sqlStr))
							
							if(isnull(rs("Voucher_DPA_"))) then
							brokerVoucherID = rs.fields("BrokerReceiptVoucher_DPA_")							
							DeleteItem "BrokerReceiptVoucher","BrokerReceiptVoucher_DPA_",brokerVoucherID
							else
							VoucherID = rs.fields("Voucher_DPA_")							
							DeleteItem "Voucher","Voucher_DPA_",VoucherID							
							end if
						end if
					end if
					
					'delete lot
					DeleteItem "Lot","Lot_DPA_",lotID
					
					'delete contract
					DeleteItem "Contract","Contract_DPA_",contractID
			        
			        DeleteItem "InterTransfer","Contract_DPA_",contractID
			        			        
			conn.CommitTrans
			conn.Close
			Set conn = Nothing
			WriteDialogRelocateScript "EditLotItem.asp?ID=" & IDHolder
			Response.End
   	
   	case else
   			sqlStr = "SELECT * FROM OrdDetailList WHERE OrdDetailList.OrdDetail_DPA_= " & ID
		   	
   			Set conn = GetActiveConnection("KBroker")
   			set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
   			If rs.EOF Or rs.BOF Then%>
					<script language = 'vbscript'>
                			window.self.ShowMessage "The selected Order item cannot be retrieved for lot allocation"
		                	
					</script>
					<% response.end
			End If
   	end select

	If Len(UpdatedContractDPA & "") > 0 then
		On Error Resume Next
		'Calculate and Enter VAT *****************************************
		sqlStr = "SELECT SUM(LevyAmount) FROM LevyContractList WHERE LevyShortName <> 'VAT' AND Contract_DPA_= " & UpdatedContractDPA
		dblVatableAmount = conn.execute(sqlStr)(0)

		sqlStr = "SELECT LevyRate FROM LevyContractList WHERE LevyShortName = 'VAT' AND Contract_DPA_= " & UpdatedContractDPA
		dblVATRate = conn.execute(sqlStr)(0)

		dblVATAmount=(dblVatableAmount*(dblVATRate/100))

		
		sqlStr = "UPDATE LevyContractList SET LevyAmount = " & dblVATAmount & " WHERE LevyShortName = 'VAT' AND Contract_DPA_= " & UpdatedContractDPA
		conn.execute(sqlStr)
	End If

%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit <%=DataEntity%> Item</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
 <script language='vbscript'>
			function ItemSelected(itemID)
					
 					frm<%=DataSource%>Item.elements("ItemID").value = itemID
			end function
			
			
			function SaveInPlaceEdit()
				    Dim myOwnerFrame				
					'UpdateID
					'Set window.parent.dialogArguments.opener.parent.frames("footer").editDocOpener = window.self
					'frm<%=DataSource%>Item.target = "deleteFrame" 					
					frm<%=DataSource%>Item.submit
			end function
		</script>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
		
		<!-- ActiveUI stylesheet and scripts -->
		<link href="../runtime/classic/activeui.css" rel="stylesheet" type="text/css">
		<script src="../runtime/activeui.js"></script>
		<!-- Include patches here -->
		<script src="../runtime/paging1.js"></script>
		<!-- grid format -->
		<style> 
			.active-controls-grid {height: 100%; font: menu;}
			.active-row-highlight .active-row-cell {background-color: skyblue}
		    
		    
		     	
			.active-column-0 {width: 50px;}
			.active-column-1 {width: 70px;}
			.active-column-2 {width: 120px;}
			.active-column-3 {width: 80px;}
			.active-column-4 {width: 80px;}
			.active-column-5 {width: 200px;}
			.active-column-6 {width: 100px;}
			.active-column-7 {width: 100px;}
			
			
			.active-grid-row,
			.active-grid-row.active-list-item,
			.active-scroll-left .active-list-item {height: 22px;}
			
			
			.active-selection-true, .active-selection-true .active-row-cell {
				color: blue!important;
				background-color: bisque!important;
				}
		</style>
		

</head>

<body Class="Dialog" SCROLL="No">	
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<form name = 'frm<%=DataSource%>Item' id='frmMain' method = 'post' action = '<%=DataSource%>Item.asp' OnSubmit="UpdateID();">

<SCRIPT language="JavaScript">
	var calTDate;
	function changeDateInterface(selCol){
		try{
			calTDate = new ctlSpiffyCalendarBox('calTDate', 'frm<%=DataSource%>Item', 'txtTDate', 'cmdTDate','<%= FormatDate(Date) %>', 1); 
			calTDate.readonly = false;
			calTDate.returnOutStringOnWrite(); 
			//var parentDiv = document.all.item("txtTDate").parentNode;
			//parentDiv.innerHTML = calTDate.writeControl();
			//if (selCol==null || selCol == "undefined"){
				document.all.item("txtTDate").outerHTML = calTDate.writeControl();
		//	}
		//	else{
		//		document.all.item(selCol).outerHTML = calTDate.writeControl();
				
		//	}	
			//parentDiv.style.zIndex = 10;
			//parentDiv.childNodes(1).style.zIndex = 10	;
			
			//Contract Settlement Date
			calSDate = new ctlSpiffyCalendarBox('calSDate', 'frm<%=DataSource%>Item', 'txtSDate', 'cmdSDate','<%= FormatDate(Date) %>', 1); 
			calSDate.readonly = false;
			calSDate.returnOutStringOnWrite(); 

			document.all.item("txtSDate").outerHTML = calSDate.writeControl();
		
		}
		
		catch(e){}	
		
		document.all.item('txtSlip').focus();
	}
	
	document.body.onload = changeDateInterface;
	
	//function ShowGrid()
	//{
	///	ShowMessage(document.all.item("GridCell").innerText);
	//}
</SCRIPT>
 
  
 <%
 Set conn = GetActiveConnection("KBroker")
 
	Dim rowCount
	Dim brokerList
	Dim quote 
	
	quote =  chr(34)
	brokerList = GetBrokerList("cboBroker")
	       
    sqlStr = "SELECT * FROM LotList WHERE LotList.OrdDetail_DPA_= " & ID
    sqlStr1 = "SELECT * FROM OrdDetailList WHERE OrdDetailList.OrdDetail_DPA_= " & ID
       
    Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr1)))
    
    'Response.Write(sqlStr1)
    'Response.End
    
    If (rs.EOF Or rs.BOF) Then
    %>
					<script language = 'vbscript'>
                			window.self.close		                	
					</script>
					<% 
				  'WriteDialogRelocateScript "AddLot.asp?ID=" & DetailID					
				  Response.redirect "AddLot.asp?ID=" & ID					
				  response.end
			
    end if
    
    If Not(rs.EOF Or rs.BOF) Then
    clientdpa=rs.Fields("Client_DPA_")
   
    
		'store the Order Type%>
        <input type = 'hidden' name ='txtOrderType' id = 'txtOrderType' size='9' value = '<%=rs.Fields("OrdDetailType")%>'>
        <input type = 'hidden' name ='txtInstrument' id = "txtInstrument" size='9' value = '<%=rs.Fields("OrdDetailSecType")%>'>
        <input type = 'hidden' name ='txtOrderIsSaleType' id = "txtOrderIsSaleType" size="20" value = '<%=rs.Fields("OrderTypeSale")%>'>
        <input type = 'hidden' name ='txtSecurityID' id = "txtSecurityID" size="20" value = '<%=rs.Fields("Security_DPA_")%>'>
        <input type = 'hidden' name ='txtAgentCommission' id = "txtAgentCommission" size="20" value = '<%=rs.Fields("AgentCommission")%>'>
        <input type = 'hidden' name ='txtStaffCommission' id = "txtStaffCommission" size="20" value = '<%=rs.Fields("StaffCommission")%>'>
        <input type = 'hidden' name ='txtCommission' id = "txtCommission" size="20" value = '<%=rs.Fields("CommissionRate")%>'>
        <input type = 'hidden' name ='txtVolumeCommission' id = "txtVolumeCommission" size="20" value = '<%=rs.Fields("VolumeRate")%>'></td>
	  <input type = 'hidden' name ='txtVolumeBoundary' id = "txtVolumeBoundary" size="20" value = '<%=rs.Fields("VolumeBoundary")%>'></td>
	  <input type = 'hidden' name ='txtMinimumCommission' id = "txtMinimumCommission" size="20" value = '<%=rs.Fields("MinimumCommission")%>'></td>
	  <input type = 'hidden' name ='txtCMA' id = "txtCMA" size="20" value = '<%=rs.Fields("CMARegulated")%>'></td>
		<input type = 'hidden' name ='txtPostImmobilisedRate' id = "txtPostImmobilisedRate" size="20" value = '<%=rs.Fields("PostImmobilisedRate")%>'></td>
		<input type = 'hidden' name ='txtSecurityImmobilised' id = "txtSecurityImmobilised" size="20" value = '<%=rs.Fields("SecurityImmobilised")%>'></td>
		<input type = 'hidden' name ='txtClientDPA' id = "txtClientDPA" size="20" value = '<%=clientdpa%>'></td>
		<input type = 'hidden' name ='txtEntityDPA' id = "txtEntityDPA" size="20" value = '<%=rs.Fields("EntityType_DPA_")%>'></td>
            <input type = 'hidden' name ='txtClass' id = "txtClass" size="20" value = '<%=rs.Fields("Class")%>'>
            
            <input type = 'hidden' name ='txtSettlementDate' id = "txtSettlementDate" size="20" value = ''>
            </td>

		<% 
		if rs.Fields("InterBank")=true then
		%>
		<input type = 'hidden' name ='txtinterbank' id = "txtinterbank" size="20" value = '1'></td>
		<%
		else
		%>
		<input type = 'hidden' name ='txtinterbank' id = "txtinterbank" size="20" value = '0'></td>
		<%
		end if
		%>
	 
		<% 
		if rs.Fields("IsCustodian")=true then
		%>
		<input type = 'hidden' name ='txtcustodian' id = "txtcustodian" size="20" value = '1'></td>
		<%
		else
		%>
		<input type = 'hidden' name ='txtcustodian' id = "txtcustodian" size="20" value = '0'></td>
		<%
		end if
		%>

		<!-- grid data -->
		<% 'row data
	end if
	
	Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
	If Not(rs.EOF Or rs.BOF) Then
		
		rowCount = 0 
		rs.MoveFirst
		Do Until rs.EOF 
		
'======================= Begin_Alter_Across_Entities =================================
			'row ID 
			
			rowData = rowData & quote & rs.Fields("Lot_DPA_") & quote & " : " 
			
			'row data 
			rowData = rowData & "[" 
			rowData = rowData & quote & rs.Fields("Lot_DPA_") & quote & "," 
			rowData = rowData & quote & rs.Fields("LotSlipNo") & quote & ","
			rowData = rowData & quote & FormatDate(rs.Fields("LotTDate")) & quote & ","
			rowData = rowData & quote & rs.Fields("LotQty") & quote & ","
			rowData = rowData & quote & rs.Fields("LotPrice") & quote & ","
			rowData = rowData & quote & rs.Fields("BrokerCode") & " : " & rs.Fields("BrokerName") & quote & ","
			rowData = rowData & quote & FormatDate(rs.Fields("ContractSettlementDate")) & quote & ","
			rowData = rowData & quote & " " & quote  
			rowData = rowData & "]" 
			
			rowIDs = rowIDs & quote & rs.Fields("Lot_DPA_") & quote 
			rowCount = rowCount + 1
		
		
			rs.MoveNext 
			
			
				'build the row IDs array
				rowIDs = rowIDs & "," 
				rowData = rowData & ","	
			
'======================= End_Alter_Across_Entities =================================

			
		Loop

		
		rs.MoveFirst
		%><script language="javascript">
			//update the quantity balance
			try{
				window.parent.frames("header").document.frmMain.elements("txtBalance").value = '<%= FormatNum(rs.Fields("BalanceQty")) %>';
			}
			catch(e){}	
		 </script><%
	End if
	
		'row ID 	
		rowData = rowData & quote & -1 & quote & " : " 
				
		'row data 
		rowData = rowData & "[" 
		rowData = rowData & quote & "New Line" & quote & "," 
		'rowData = rowData & quote & "<input type = 'text' name ='txtSlip' id = 'txtSlip' size='9' onChange = 'AddRowInProgress();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>" & quote & ","
		'rowData = rowData & quote & "<input type='text' name='txtTDate' size=40 value='" & FormatDate(Date) & "' onChange = 'AddRowInProgress();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>" & quote & ","
		'rowData = rowData & quote & "<input type = 'text' name ='txtQty' id = 'txtQty' size='9' onChange = 'AddRowInProgress();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>" & quote & ","
		'rowData = rowData & quote & "<input type = 'text' name ='txtPrice' id = 'txtPrice' size='9' onChange = 'AddRowInProgress();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>" & quote & ","
		rowData = rowData & quote & "<input type = 'text' name ='txtSlip' id = 'txtSlip' size='7' OnClick='event.cancelBubble=true;'>" & quote & ","
		rowData = rowData & quote & "<input type='text' name='txtTDate' size=40 value='" & FormatDate(Date) & "' OnClick='event.cancelBubble=true;'>" & quote & ","
		rowData = rowData & quote & "<input type = 'text' name ='txtQty' id = 'txtQty' size='9' OnClick='event.cancelBubble=true;'>" & quote & ","
		rowData = rowData & quote & "<input type = 'text' name ='txtPrice' id = 'txtPrice' size='9' OnClick='event.cancelBubble=true;'>" & quote & ","
		rowData = rowData & quote & brokerList & quote  & ","
		rowData = rowData & quote & "<input type='text' name='txtSDate' size=40 value='" & FormatDate(Date) & "' OnClick='event.cancelBubble=true;'>" & quote  & ","
		rowData = rowData & quote & "<input type=button value='Add' Class=Buttons OnClick='JavaScript: AddRowInProgress();'>&nbsp;&nbsp;<input type='reset' value='Cancel' Class=Buttons>" & quote 
		rowData = rowData & "]" 
	

		'build the row IDs array 
		rowIDs = rowIDs & quote & -1 & quote 
		rowCount = rowCount + 1
		
'======================= Begin_Alter_Across_Entities =================================%> 
		<script language="javascript">
			//column titles 
			var colCount = 8;
			var colNames = ["", "Ref",  
					 "Date", "Quantity","Price","Broker","SettlementDate", ""];
			
			var myColumns = ["Lot No", "Ref",  
					 "Date", "Quantity","Price","Broker","SettlementDate" ,""];
		</script>
<%'======================= End_Alter_Across_Entities =================================%>			
		<script language="javascript">
		
			
			//data
			var myData = {<%=rowData%>}; 
			var myRowIDs = [<%=rowIDs%>]; 
			
			
			//editing
			var inPlaceEdit = false;
			var addInProgress = false;
			var clickedRowID = -1; 
			var dataChanged = false;
			var prevRow = -1;//the row currently under in-place edit
			
			function EditInPlaceDataChanged()
			{
				dataChanged = true;
			}
			
			function AddRowInProgress()
			{
				addInProgress = true;
			}
			
			var currentBrokerName = "";
			var RowEditFn = function(src)
			{
				var rowIndex = src.getProperty("row/index");
				var i;
				if (rowIndex==0) return;
				for(i = 0; i < colCount; i++)
				{
					if(colNames[i] != "")
					{
						if(prevRow >= 0)
						{
							if(colNames[i]=="Broker")
							{
								myData[prevRow][i] = currentBrokerName;
							}
							else
							{
								myData[prevRow][i] = document.frmMain.elements(colNames[i]).value;
							}
						}
						if(colNames[i]=="Broker")
						{	
							currentBrokerName =  myData[rowIndex][i];
							myData[rowIndex][i] = inPlaceList;
							
						}
						else
						{
							if(colNames[i]=="Date" || colNames[i]=="SettlementDate")
							{
								currentDate =  myData[rowIndex][i];								
								myData[rowIndex][i] = "<INPUT TYPE='text' NAME='" + colNames[i] + "' ID='" + colNames[i] + "' VALUE='" + currentDate + "' onChange = 'EditInPlaceDataChanged();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>";
								//show the calendar
								changeDateInterface(colNames[i]);
				
								
							}
							else {
								myData[rowIndex][i] = "<INPUT TYPE='text' NAME='" + colNames[i] + "' ID='" + colNames[i] + "' VALUE='" + myData[rowIndex][i] + "' onChange = 'EditInPlaceDataChanged();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>";
							}	
						}
					}
				}
				
				myData[rowIndex][colCount - 1] = "<INPUT TYPE='button' class='Buttons' VALUE='Save' onClick = 'SaveInPlaceEdit();event.cancelBubble=true;'>&nbsp;<INPUT TYPE='button' class='Buttons' VALUE='Cancel' onClick = 'cancelEditRow();event.cancelBubble=true;'>";
				inPlaceEdit = true;
				prevRow = rowIndex;
				grid.refresh();
				//select the appropriate item
				var secList = document.frmMain.elements("cboBrokerInPlace");
				for (i=0; i < secList.options.length; i++) {
					if(secList.options(i).text == currentBrokerName)
					{
							secList.options(i).selected = true;
					}
				}
				
				
			}
			
			function cancelEditRow(){
					var i;
							for(i = 0; i < colCount; i++)
							{
								if(colNames[i] != "")
								{
									if(colNames[i]=="Broker")
									{
										myData[prevRow][i] = currentBrokerName;
									}
									else
									{
										myData[prevRow][i] = document.frmMain.elements(colNames[i]).value;
									}
								}
							}
							myData[prevRow][colCount - 1] = "";						
							inPlaceEdit = false;
							prevRow = -1;
							grid.refresh();
			}
			
			var RowChangeFn = function(src)
			{
				if(inPlaceEdit || addInProgress)
				{
					if(dataChanged || addInProgress)
					{
						ItemSelected(prevRow);
						SaveInPlaceEdit();
					}
					else
					{
						if(prevRow != clickedRowID)
						{
							cancelEditRow();
						}
					}
				}
			}
			
			var HandleClick = function(src)
			{
				clickedRowID = src.getProperty("row/index");
				ItemSelected(clickedRowID);
			}
			
			var headerID;
			
			function UpdateID(){
				headerID = window.parent.frames["header"].document.all.item("ID").value;
				//headerID = window.document.all.item("ID").value;
				document.all.item("ID").value = headerID;
				
				theSDate = window.parent.frames["header"].document.all.item("txtSDate").value;
				//theSDate = window.document.all.item("txtSDate").value;
				document.all.item("txtSettlementDate").value = theSDate;
			}
			
			function restoreID(){
				document.all.item("ID").value = headerID;
			}
			
			function HandleDeleteAction()
			{
					document.frmMain.elements("action").value = "Execute_Delete"
					SaveInPlaceEdit();
					document.frmMain.elements("action").value = "Execute_Detail"
			}
			
			function HandleSaveAction()
			{
					document.frmMain.elements("action").value = "Execute_Save"
					SaveInPlaceEdit();
					document.frmMain.elements("action").value = "Execute_Detail"
			}
			//get ready for in-place edit
			var inPlaceList = "<%=GetBrokerList("cboBrokerInPlace")%>"
		</script> 
		
		<script language="javascript"> 

			// create ActiveUI Grid javascript object 
			var grid = new Active.Controls.Grid; 
			
			
			// set rows ids 
			grid.setRowValues(myRowIDs); 
			

			// set number of columns 
			grid.setColumnCount(colCount); 

			// provide cells and headers text 
			grid.setDataText(function(i, j){return myData[i][j]}); 
			grid.setColumnText(function(i){return myColumns[i]}); 

			// set click action handler 
			grid.setAction("click", HandleClick); 
			//grid.setAction("dblclick", RowEditFn); 
			grid.setAction("selectionChanged", RowChangeFn);
			
			//stripes 
			var alternate = function(){ return this.getProperty("row/order") % 2 ? "gainsboro" : "white";} 
			var row = new Active.Templates.Row; row.setStyle("background", alternate); 
			row.setEvent("onmouseover", "mouseover(this, 'active-row-highlight')"); 
			row.setEvent("onmouseout", "mouseout(this, 'active-row-highlight')"); 
			grid.setTemplate("row", row); 
			var column = new Active.Templates.Text; 
			column.setStyle("border-right", "1px solid white");  
			grid.setTemplate("column", column);  grid.setRowHeaderWidth("0px"); 
			
			//disable sort
			grid.getTemplate("top/item").setEvent("onmousedown", null);
			
			// write grid html to the page 
			document.write(grid); 
			
			//let grid be aware of composite layout
			grid.getLayoutTemplate().action("adjustSize");
		</script> 
        <%
       
  function GetBrokerList(listName)
		Dim secList
		
		secList = "<select name = '" & listName & "' id = '" & listName & "' size='1' "
		secList = secList & "OnClick='event.cancelBubble=true;' "  
		secList = secList & "onChange='event.cancelBubble=true;' " 
		secList = secList & "onKeypress='return (dodefaultaction()==\""\""); ' "  
		secList = secList & "onKeydown='return (dodefaultaction()==\""\"");event.cancelBubble=true;' "  
		secList = secList & "onKeyup='return (change(" & listName & "));' "  
		secList = secList & "onfocus='txtval = \""\"";inputIsItemCode = 1;' "  
		secList = secList & "onblur='txtval = \""\"";inputIsItemCode = 1;'>"
		
		secList = secList & "<option selected SearchCode = '0' SearchText = ''  value = ''></option>"
		
        sqlStr = "SELECT * FROM [BrokerList] Order By BrokerCode"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF
                        secList = secList & "<option SearchCode = '" & rs.Fields("BrokerCode") & "' SearchText = '" & rs.Fields("BrokerName") & "'  value = '" & rs.Fields("Broker_DPA_") & "'>" & rs.Fields("BrokerNameEx") & "</option>"
                        rs.MoveNext
                Loop
        End If
	    secList = secList & "</select>"
	    GetBrokerList = secList
  end function
  
  function DeleteItem(EntityName,KeyField,DelItemID)
		dim delRS
			'find out whether any child records exist
			sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = '" & EntityName & "') AND (ChildType = " & LinkedIndependent & ")"
			Set delRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			If Not (delRS.BOF Or delRS.EOF) Then
					Dim childRS
					Dim tableName
					
					delRS.MoveFirst
					Do Until delRS.EOF
                			tableName = delRS.Fields("Child")
							sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE " & KeyField & " = " & DelItemID & " and Deleted <>1"							
					
							Set childRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
							If Not (childRS.BOF Or childRS.EOF) Then%>
                					<script language = 'vbscript'>
                						ShowMessage "<%=delRS.Fields("DeletionMessage")%>"
	                					
                					</script>
                					<%response.end
							End If
							delRS.MoveNext
					Loop
			End If
			
			'delete from database
			if(ucase(Trim(EntityName))="LOT" OR ucase(Trim(EntityName))="PAYMENT") then
			sqlStr = "Update  [" & EntityName & "] Set Deleted = 1,ChangedBy=" & UserId & ",TimeChanged=GetDate() WHERE " & KeyField & " = " & DelItemID
			else
			sqlStr = "Update  [" & EntityName & "] Set Deleted = 1 WHERE " & KeyField & " = " & DelItemID
			end if
			
			conn.Execute SQLServerFormat(HandleQuote(sqlStr))
  end function
  

 %>
 <table border="0" width="100%" ID="Table1">
<tr><td>
<input type = 'hidden' name ='ItemID' id = 'ItemID'>
<input type = 'hidden' name ='ID' id = 'ID' value="<%= IDHolder %>">
<input type = 'hidden' name ='action' id = 'action' value="Execute_Detail">
</td>
</tr>
</table>
</form>
</body>

</html>