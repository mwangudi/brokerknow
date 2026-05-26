<!--#include file="../libroutines.asp"-->
<%
	'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "AddLot"
		const DataEntity = "Lot"
		const DataEntityPlural = "Lots"
		const ActionFolder = "Operations"
    '======================= End_Alter_Across_Entities =================================
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim guidStr 
   Dim guid 
	
	UserId=Session("UserID")
	
	action = ucase(Request.Form("action"))
	
	if action = "EXECUTE" then
		Dim buttonAction
		Dim reloadRequired
		
		reloadRequired = false
		buttonAction = Trim(Trim(Ucase(Request.Form("cmdAdd"))))
		if buttonAction = "SAVE" then
				ID = Request("ID")

				If Trim(ID) = "" Then%>
						<script language = 'vbscript'>
                				ShowMessage "No Order item was specified for Lot allocation"
                				window.self.close
						</script>
						<%response.end
				End If
				
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
			 	Dim interbank
				Dim ContractDPA
				Dim custodian
				Dim client
				Dim cliententity
				Dim ClientClass
				Dim ContractSettlementDate 

				volComm = Request.Form("txtVolumeCommission")
				volBound = ccur(Request.Form("txtVolumeBoundary"))
				minComm = ccur(Request.Form("txtMinimumCommission"))
				cma = Request.Form("txtCMA")
				imobRate = Request.Form("txtPostImmobilisedRate")
				secImob = Request.Form("txtSecurityImmobilised")
				broker = Request.Form("cboBroker")
				tDate = Trim(Request.Form("txtTDate"))
				tDate = tDate & " " & Time				
				slip = Request.Form("txtSlip")
				qty = Request.Form("txtQty")
				price = Request.Form("txtPrice")
				orderType = Request.Form("txtOrderType")
				orderSecType = Request.Form("txtInstrument")
				orderIsSaleType = cbool(Request.Form("txtOrderIsSaleType"))
				securityID = Request.Form("txtSecurityID")
				agentCommission = Request.Form("txtAgentCommission")
				staffCommission = Request.Form("txtStaffCommission")
				regularComm = Request.Form("txtCommission")
				interbank=Cint(Request.Form("txtinterbank"))       
       			custodian=Cint(Request.Form("txtcustodian"))       
       			client =Request.Form("txtClientDPA")       
       			cliententity =Cint(Request.Form("txtEntityDPA"))       
       			clientClass =Cint(Request.Form("txtClass"))
				ContractSettlementDate =trim(Request.Form("txtSettleDate"))

       			Set TimeLimitRs = CreateObject("ADODB.Recordset")   						        
   				TimeLimitRs.CursorLocation = adUseClient
   				
   				'Set CustodianRS	= CreateObject("ADODB.Recordset")   						        
   				'CustodianRS.CursorLocation = adUseClient
   				   				
				Set conn = GetActiveConnection("KBroker")
			     
				'validate Settlement Date
				 If Trim(ContractSettlementDate) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Settlement Date."
				         		
				         </script>
				         <% response.end
				 End If

				 'validate Broker
				 If Trim(Broker) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Broker"
				         		
				         </script>
				         <% response.end
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
						ShowMessage "Order Detail Estimated Price must be numeric"
						
				    </script>
				    <% response.end
				End If
				'ensure Order Detail Estimated Quantity is numeric
				If (Qty <> "") And (Not IsNumeric(Qty)) Then%>
				    <script language = 'vbscript'>
						ShowMessage "Order Detail Estimated Quantity must be numeric"
						
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
				
				

				if not(TimeLimitRS.eof or TimeLimitRs.bof) then
	 				NoOfDays=TimeLimitRS("TimeLimitLimDaysNSE")
				end if

			
				         			
	
				settlementDate=LTdate(CDate(tDate),NoOfDays)
				
				 'save contract
				 set guid = server.createobject("NDUtils.CGUID")
				 guidStr = guid.GenerateGUID				 
				
				 sqlStr = "SET IDENTITY_INSERT Contract ON INSERT INTO [Contract] (Contract_DPA_, Contract_EIT_, Status_DPA_,IsInterBank,ContractSettlementDate) " & _
				         "SELECT " & " " & "iif(isnull(max([Contract_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Contract'),max([Contract_DPA_]) + 1)" & " " & " as Contract_DPA_" & _
				         "," & "'" & guidStr & "'" & " as Contract_EIT_" & _
				         "," & " " & 1 & " " & " as Status_DPA_" & _
				         "," & " " & interbank & " " & " as IsInterBank" & _
						 "," & "#" & ContractSettlementDate & "#" & " as ContractSettlementDate" & _
				         " FROM [Contract]   SET IDENTITY_INSERT Contract OFF"
				 
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
						Dim CDSFlag
						
						if secImob then 
								CDSFlag = 1
						else
								CDSFlag = 0
						end if

						sqlStr = "SET IDENTITY_INSERT Lot ON INSERT INTO [Lot] (Lot_DPA_,Contract_DPA_,OrdDetail_DPA_,LotPrice" & _
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
								" FROM [Lot]  SET IDENTITY_INSERT Lot OFF"
				     
						sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
						conn.Execute sqlStr
						
						'apply relevant levies
						Dim levyRS
						Dim cond
						if orderSecType = "Fixed" then 
							cond = "WHERE LevyShortName <> 'VAT' AND LevyShortName <> 'CSD' AND LevyAppBond = 1 AND LevyActive = 1"
						else
							cond = "WHERE LevyShortName <> 'VAT' AND LevyShortName <> 'CSD' AND (LevyAppSecurity = 1) AND (LevyActive = 1) AND (" & _
									" Levy_DPA_ IN (SELECT Levy_DPA_ FROM LevySecurity " & _
									" WHERE Security_DPA_ = " & securityID & "))"
						end if
												
						sqlStr = "SELECT * FROM [LevyList] " & cond

						Set levyRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
							
						Dim levyAmount
						Dim vatAmount
						vatAmount=0
						if not(levyRs.Eof or LevyRs.bof) then
							levyRS.MoveFirst	
						end if

						sqlStr = "SELECT     * " & _
									"FROM         Levy "& _
									"WHERE     (Vatable = 0) AND (LevyShortName LIKE '%VAT%')"

						Set levyVATRS = conn.Execute(sqlStr)
						If Not (levyRS.EOF Or levyRS.BOF) Then													
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
										if levyRS.Fields("Vatable") = 1 or levyRS.Fields("Vatable") = "1" then 
											 				
										end if 

										vatAmount = vatAmount + levyAmount
										sqlStr = " INSERT INTO [LevyContract] (Contract_DPA_,LevyName,LevyAmount,LevyRate,LevyBlock,SystemMaintained,LevyShortName,LevyRatePercentage) values( " & _
												"       " & " " & rs.Fields("Contract_DPA_")  & _
												"       ," & "'" & levyRS.Fields("LevyDescription") & "'" & _
												"       ," & " " & RoundPoint05(levyAmount) & " "   & _
												"       ," & " " & levyRS.Fields("LevyAmount") & " "  & _
												"       ," & " " & levyRS.Fields("LevyBlock") & " "  & _
												"       ," & " " & levyRS.Fields("SystemMaintained") & " "  & _	
												"       ," & " '" & levyRS.Fields("LevyShortName") & "' " & _				
												"       ," & " '" & LevyRatePercentage  & "' "  & _							
												"        )"
																'response.write sqlStr & "<br><br>"					
												levyRS.MoveNext
										
										conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
										%>
										<script>
											//alert("<%= err.Description & "/////" & HandleQuote(sqlStr) %>")
										</script>
										<%
										'conn.Execute sqlStr

								Loop
								'response.write vatamount :response.end


							
							
						end if													
						
							'Enter CSD Fee
							sqlstr = "SELECT IsOnCSD FROM Security WHERE Security_DPA_ = " & securityID
							If conn.execute(sqlstr)(0) then

								sqlstr = "SELECT LevyAmount FROM Levy WHERE LevyShortName = 'CSD'"
								dblCSDPerc = conn.execute(sqlstr)(0)

								dblCSDFee = (dblCSDPerc * grossAmount) / 100
								vatAmount = vatAmount + dblCSDFee

								sqlStr = "INSERT INTO [LevyContract] (Contract_DPA_,LevyName,LevyAmount,LevyRate,LevyBlock,SystemMaintained,LevyShortName,LevyRatePercentage) values ("& _
												"       " & " " & rs.Fields("Contract_DPA_") & _
												"       ," & "'CSD Fee'"  & _
												"       ," & " " & RoundPoint05(dblCSDFee)  & _
												"       ," & " " & dblCSDPerc & " "& _
												"       ," & " 0 "  & _
												"       ," & " " & levyVATRS.Fields("SystemMaintained") & _	
												"       ," & " 'CSD' "  & _												
												"       ," & " '" & dblCSDPerc & "%' "  & _								
												"        ) "
								'response.write sqlstr
								conn.execute (sqlstr)


							End If


							'enter Handling fee
							sqlstr="SELECT     OrderType.HandlingFee " & _
							 " FROM         tbOrder INNER JOIN " & _
							 "                       OrdDetail ON tbOrder.Order_DPA_ = OrdDetail.Order_DPA_ INNER JOIN " & _
							 "                       OrderType ON tbOrder.OrderType_DPA_ = OrderType.OrderType_DPA_ " & _
							 " WHERE     (OrdDetail.OrdDetail_DPA_ = "&  ID  &")"

							set rsHadlingFee = conn.execute(sqlstr)
							
							handlingFee = rsHadlingFee("HandlingFee")
							set rsHadlingFee= nothing
							vatAmount = vatAmount + handlingFee
							sqlStr = "INSERT INTO [LevyContract] (Contract_DPA_,LevyName,LevyAmount,LevyRate,LevyBlock,SystemMaintained,LevyShortName,LevyRatePercentage) values ("& _
											"       " & " " & rs.Fields("Contract_DPA_") & _
											"       ," & "'Handling fee'"  & _
											"       ," & " " & RoundPoint05(handlingFee)  & _
											"       ," & " 0 "& _
											"       ," & " 0 "  & _
											"       ," & " 97 "  & _	
											"       ," & " 'Handling' "  & _												
											"       ," & " '0' "  & _								
											"        ) "
							'response.write sqlstr
							Set rsHandling = conn.execute (sqlstr)


						'consider the transfer fee
						if Not(orderIsSaleType) then
							Dim transRS
							sqlStr = "SELECT * FROM [SecTransFeeListLatest] WHERE Security_DPA_ = " & securityID
							Set transRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
							sqlStr = "INSERT INTO [LevyContract] (Contract_DPA_,LevyName,LevyAmount,LevyRate,LevyBlock,LevyShortName,SystemMaintained) VALUES( " & _
									"       " & " " & rs.Fields("Contract_DPA_") & " "  & _
									"       ," & "'Transfer Fee'"  & _
									"       ," & " " & transRS.Fields("Fee") & " "  & _
									"       ," & " " & transRS.Fields("Fee") & " " & _
									"       ," & " " & 0 & " "  & _
									"       ," & " 'Transfer' " & _
									"       ,13 " & _	
									"        ) "
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
										 					ShowMessage "The assigned commission rate is inviolation of BSEC regulations"
										 					
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
						vatamount = vatamount + levyAmount



						sqlStr = " INSERT INTO [LevyContract] (Contract_DPA_,LevyName,LevyAmount,LevyRate,LevyBlock,LevyShortName,LevyRatePercentage,SystemMaintained) Values ("  & _
								"       " & " " & rs.Fields("Contract_DPA_") & _
								"       ," & "'Broker Commission'"  & _
								"       ," & " " & RoundPoint05(levyAmount)  & _
								"       ," & " " & commission & " " & _
								"       ," & " " & 0 & " "  & _
								"       ," & " 'Commission' "  & _
								"       ," & " '" & commission & "%" & "' "  & _
								"       ,11 "  & _	
								"        )"
						sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
						conn.Execute sqlStr
				

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
						
							sqlStr = " INSERT INTO [LevyContract] (Contract_DPA_,LevyName,LevyAmount,LevyRate,LevyBlock,LevyShortName,LevyRatePercentage,SystemMaintained) Values ("  & _
									"       " & " " & rs.Fields("Contract_DPA_") & _
									"       ," & "'Botswana Stock Exchange Commission'"  & _
									"       ," & " " & RoundPoint05(dblBSECommission)  & _
									"       ," & " " & dblBSEComm & " " & _
									"       ," & " " & 0 & " "  & _
									"       ," & " 'BSEComm' "  & _
									"       ," & " '" & dblBSEComm & "%" & "' "  & _
									"       ,25 "  & _	
									"        )"
							sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
							conn.Execute sqlStr


						End If
						Set objBSE = Nothing


						'Apply agent commission
						Dim tmpLevy 'broker commission
						tmpLevy = commissionAmount
						levyAmount = CCur((agentCommission/100.00) * tmpLevy)
						vatamount = vatamount+levyAmount


						sqlStr = "INSERT INTO [LevyContract] (Contract_DPA_,LevyName,LevyAmount,LevyRate,LevyBlock,LevyShortName,LevyRatePercentage,SystemMaintained) values( " & " " & _
								"       " & " " & rs.Fields("Contract_DPA_") & " " & _
								"       ," & "'Agent Commission'"  & _
								"       ," & " " & RoundPoint05(levyAmount) & " " & _
								"       ," & " " & agentCommission * (commission/100.00) & " " & _
								"       ," & " " & 0 & " "  & _
								"       ," & " 'Agent' "  & _
								"       ," & " '" & agentCommission & "%" & "' " & _
								"       ,12 " & _	
								"       )"
						conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
						'Enter VAT
							LevyRatePercentage = levyVATRS.Fields("LevyAmount") & "%"
							
							vatamount=(vatAmount*(levyVATRS.Fields("LevyAmount")/100))


							'response.write vatamount :response.end
							sqlStr = "INSERT INTO [LevyContract] (Contract_DPA_,LevyName,LevyAmount,LevyRate,LevyBlock,SystemMaintained,LevyShortName,LevyRatePercentage) values (" & _
											"       " & " " & rs.Fields("Contract_DPA_") & _
											"       ," & "'" & levyVATRS.Fields("LevyDescription") & "'" & _
											"       ," & " " & RoundPoint05(VatAmount) & " "  & _
											"       ," & " " & levyVATRS.Fields("LevyAmount") & " "  & _
											"       ," & " " & levyVATRS.Fields("LevyBlock") & " "  & _
											"       ," & " " & levyVATRS.Fields("SystemMaintained") & " " & _	
											"       ," & " '" & levyVATRS.Fields("LevyShortName") & "' "  & _							
											"       ," & " '" & LevyRatePercentage  & "' " & _							
											"       )"


							conn.Execute sqlStr
						
							Set levyVATRS = nothing
						
						
						'post to accounts for CDS transactions
						if secImob then
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
					
								pDate = settlementDate
								ContractsSel = rs.Fields("Contract_DPA_")
								amount = grossAmount
								bank = 4 'CDS clearing account
								entity = 8 'Broker
								account = 14 'CDS system account
								custOrder = "Null"
								clientVoucher = "Null"
					
								if orderIsSaleType then
										if(Cint(custodian)=1 OR Cint(interbank)=1) then 'Custodian transaction done by stivo
											if(Cint(custodian)=1 and Cint(interbank) <>1) then
												set guid = server.createobject("NDUtils.CGUID")
												InterTransfer_EIT_ = guid.GenerateGUID
										
												sqlStr = "INSERT INTO [InterTransfer] (SourceEntityType_DPA_,InterTransfer_EIT_,SourceEntity_DPA_" & _
														 "       ,TargetEntityType_DPA_,TargetEntity_DPA_,TransferAmount,TransferDate,ChangedBy,InterTransferType_DPA_,Contract_DPA_)" & _
												 		 " Values( " & " " & 8 & " " & " " & _
												 		 "       ," & "'" & InterTransfer_EIT_ & "'" & " " & _
												 		 "       ," & " " & 14 & " " & " ," & " " & cliententity & " " & " " & _
												 		 "		," & " " & client & " " & " " & _
												 		 "		," & " " & ccur(amount) & " " & " " & _
												 		 "		," & "#" & FormatDate(settlementDate) & "#" & " " & _
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
												 		"," & "#" & FormatDate(settlementDate) & "#" & " " & _
												 		"," & " " & UserId & " " & " " & _												 
												 		"," & " " & 2 & " " & " " & _				
												 		"," & " " & ContractDPA & " " & ")"
      
												sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
										 		'Response.Write(sqlStr)
										 		'Response.End
										 
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
										
								end if
						end if
				conn.CommitTrans
				
				WriteDialogRelocateScript "EditLot.asp?ID=" & ID
				Response.end
		end if

	end if
	
	Dim IDHolder
	Dim IDArray
	Dim ItemID
	
	IDHolder = Request("ID")
	If (Trim(IDHolder) = "") or (IDHolder = "0") Then%>
			<script language = 'vbscript'>
                	ShowMessage "Please select an Order item for Lot allocation"
                	window.self.close
			</script>
			<%response.end
	End If
	
	IDArray = split(IDHolder,"<->")
	ID = IDArray(lbound(IDArray))
	ItemID = IDArray(ubound(IDArray))
	
	if (ItemID <> 0) then
			Response.redirect "EditLot.asp?ID=" & ID
			Response.end
	end if
   	sqlStr = "SELECT * FROM OrdDetailList WHERE OrdDetailList.OrdDetail_DPA_= " & ID
   	
   	Set conn = GetActiveConnection("KBroker")
   	set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
   	If rs.EOF Or rs.BOF Then%>
            <script language = 'vbscript'>
                	window.self.ShowMessage "The selected Order item cannot be retrieved for lot allocation"
                	
            </script>
            <% response.end
    End If
    
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add <%=DataEntity%></title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
<%
'set default settlement date
 DefaultDate = FormatDate(dateadd("d",5,Now()))
%>
<SCRIPT language="JavaScript">
	var calTDate = new ctlSpiffyCalendarBox("calTDate", "frm<%=DataSource%>", "txtTDate","cmdTDate","<%=FormatDate(Date)%>",1);
	var calDate = new ctlSpiffyCalendarBox("calDate", "frm<%=DataSource%>", "txtSettleDate","cmdSettleDate","<%=DefaultDate%>",1);
</SCRIPT>
<!--END CALENDAR -->

<script >
		var validNavigate = false;
		function ReleaseRecord()
		{
			if(!validNavigate)
			{
 				event.returnValue = "Please use the cancel button to close the dialog"
 			}
		}
		
		function AllowedNavigation()
		{
			validNavigate = true;
		}
</script>
</head>

<body Class="Dialog">
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>

<form name = 'frm<%=DataSource%>' method = 'post' action = '<%=DataSource%>.asp' target = 'deleteFrame' id='frmMain' OnSubmit="JavaScript: UpdateDialogHandle();">


<table border="0" width="100%" height="265">
  <tr>
    <td width="17%" height="25">Order No</td>
    <td width="83%" height="25"><input readonly = 'true' class=readonly   STYLE="width: 100px; text-align: right" type = 'text' name ='txtOrderNo' id = 'txtOrderNo' size="20" value = '<%=rs.Fields("Order_DPA_")%>'></td>
  </tr>
  <tr>
    <td width="17%" height="25">Item No</td>
    <td width="83%" height="25"><input readonly = 'true' class=readonly   STYLE="width: 100px; text-align: right" type = 'text' name ='txtItemNo' id = 'txtItemNo' size="20" value = '<%=rs.Fields("OrdDetail_DPA_")%>'></td>
  </tr>
  <tr>
    <td width="17%" height="25">Order Type</td>
    <td width="83%" height="25"><input readonly = 'true' class=readonly  STYLE="width: 150px"   type = 'text' name ='txtOrderType' id = 'txtOrderType' size="20" value = '<%=rs.Fields("OrdDetailType")%>'>
    <input readonly = 'true' class=readonly  type = 'hidden' name ='txtOrderIsSaleType' id = "txtOrderIsSaleType" size="20" value = '<%=rs.Fields("OrderTypeSale")%>'></td>
  </tr>
  <tr>
    <td width="17%" height="25">Instrument</td>
    <td width="83%" height="25"><input readonly = 'true' class=readonly STYLE="width: 150px"  type = 'text' name ='txtInstrument' id = 'txtInstrument' size="20" value = '<%=rs.Fields("OrdDetailSecType")%>'></td>
  </tr>
   <tr>
    <td width="17%" height="25">Client</td>
    <td width="83%" height="25">
		<input readonly = 'true' class=readonly  type = 'text' name ='txtClient'  STYLE="width: 300px" id = 'txtClient' size="20" value = '<%=rs.Fields("OrdDetailClient")%>'>
		<input type = 'hidden' name ='txtAgentCommission' id = "txtAgentCommission" size="20" value = '<%=rs.Fields("AgentCommission")%>'></td>
		<input type = 'hidden' name ='txtStaffCommission' id = "txtStaffCommission" size="20" value = '<%=rs.Fields("StaffCommission")%>'></td>
		<input type = 'hidden' name ='txtCommission' id = "txtCommission" size="20" value = '<%=rs.Fields("CommissionRate")%>'></td>
		<input type = 'hidden' name ='txtVolumeCommission' id = "txtVolumeCommission" size="20" value = '<%=rs.Fields("VolumeRate")%>'></td>
		<input type = 'hidden' name ='txtVolumeBoundary' id = "txtVolumeBoundary" size="20" value = '<%=rs.Fields("VolumeBoundary")%>'></td>
		<input type = 'hidden' name ='txtMinimumCommission' id = "txtMinimumCommission" size="20" value = '<%=rs.Fields("MinimumCommission")%>'></td>
		<input type = 'hidden' name ='txtCMA' id = "txtCMA" size="20" value = '<%=rs.Fields("CMARegulated")%>'></td>
		<input type = 'hidden' name ='txtPostImmobilisedRate' id = "txtPostImmobilisedRate" size="20" value = '<%=rs.Fields("PostImmobilisedRate")%>'></td>
		<input type = 'hidden' name ='txtSecurityImmobilised' id = "txtSecurityImmobilised" size="20" value = '<%=rs.Fields("SecurityImmobilised")%>'></td>
		<input type = 'hidden' name ='txtClientDPA' id = "txtClientDPA" size="20" value = '<%=rs.Fields("Client_DPA_")%>'></td>
		<input type = 'hidden' name ='txtEntityDPA' id = "txtEntityDPA" size="20" value = '<%=rs.Fields("EntityType_DPA_")%>'></td>
		<input type = 'hidden' name ='txtClass' id = "txtClass" size="20" value = '<%=rs.Fields("Class")%>'></td>

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

 
  </tr>

  <tr>
    <td width="17%" height="25">Security</td>
    <td width="83%" height="25"><input readonly = 'true'  STYLE="width: 300px"   class=readonly  type = 'text' name ='txtSecurity' id = 'txtSecurity' size="20" value = '<%=rs.Fields("OrdDetailSecurity")%>'>
    <input type = 'hidden' name ='txtSecurityID' id = "txtSecurityID" size="20" value = '<%=rs.Fields("Security_DPA_")%>'></td>
  </tr>
  <tr>
    <td width="17%" height="25">Balance</td>
    <td width="83%" height="25"><input readonly = 'true' STYLE="width: 150px; text-align: right"  class=readonly  type = 'text' name ='txtBalance' id = 'txtBalance' size="20" value = '<%= FormatNum(rs.Fields("BalanceQty")) %>'></td>
  </tr>
  <tr>
    <td width="17%" height="25">Settlement Date</td>
    <td width="83%" height="25"><SCRIPT language="JavaScript">calDate.writeControl();</SCRIPT></td>
  </tr>

  <tr>
  <td colspan = '2' height="58">
  
  <table border="0" width="100%">
    <tr>
      <td width="23%"><b><font color="#000080">Ref</font></b></td>
      <td width="13%"><b><font color="#000080">Date&nbsp;</font></b></td>
      <td width="17%"><b><font color="#000080">Quantity</font></b></td>
      <td width="14%"><b><font color="#000080">Price</font></b></td>
      <td width="33%"><b><font color="#000080">Broker</font></b></td>
    </tr>
    
    <tr>
      <td width="23%" valign="top"><input type = 'text' name ='txtSlip' id = 'txtSlip' size="15">&nbsp;</td>
      <SCRIPT language="javascript">
      document.all.item('txtSlip').focus();
    </script>
      <td width="13%" valign="top"><SCRIPT language="JavaScript">calTDate.writeControl();</SCRIPT></td>
      <td width="17%" valign="top"><input type = 'text' name ='txtQty' id = 'txtQty' size="11"></td>
      <td width="14%" valign="top"><input type = 'text' name ='txtPrice' id = 'txtPrice' size="10"></td>
      <td width="33%" valign="top"><select name="cboBroker" id="cboBroker" size="1" 
			onKeypress="return (dodefaultaction()==''); " 
			onKeydown="return (dodefaultaction()==''); " 
			onKeyup="return (change(cboBroker));" 
			onfocus="txtval = '';inputIsItemCode = 1;" 
			onblur="txtval = '';inputIsItemCode = 1;">
          <option selected SearchCode = "0" SearchText = ""  value=""></option>
          <%
		Set conn = GetActiveConnection("KBroker")
        sqlStr = "SELECT * FROM [BrokerList] Order By BrokerCode"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF%>
					<option SearchCode = "<%=rs.Fields("BrokerCode")%>" SearchText = "<%=rs.Fields("BrokerName")%>" value = '<%=rs.Fields("Broker_DPA_")%>'><%=rs.Fields("BrokerNameEx")%></option>
					<%rs.MoveNext
                Loop
        End If
          %>
        </select>
      </td>
    </tr>
  </table>
  </td>
  </tr>
  <tr>
    <td width="100%" colspan=2 align=right>
		<input type = 'submit' class=buttons name ='cmdAdd' id = 'cmdAdd' value=" Save " onclick = "AllowedNavigation()">
    	&nbsp; <input type = 'button' class=buttons name ='cmdClose' id = "cmdClose" value=" Close " onclick = "window.self.close();">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
	</td>
  </tr>
</table>

</form></body>

</html>