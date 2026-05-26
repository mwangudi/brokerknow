<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Activity Class</title>
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
<script language="javascript">
	function  UpdateClientAccess(theChk)
	{
		var holdVal = "0"; //no client access
		if (theChk.checked)
		{
			holdVal = "1";//client can access
		}
		
		document.getElementById("ClientAccess").value = holdVal;
				
		//document.frmMain.elements("CompoundStatus").value = holdVal;
	}
	
	function ChangePrice(thecbo)
	{
	if(thecbo.value == 1)
		{		
		document.frmMain.elements("txtValue").value=document.frmMain.elements("txtLevy").value;
		}
	else
		{
		document.frmMain.elements("txtValue").value=document.frmMain.elements("txtAmount").value;
		}
	}
	
</script>
</head>

<body Class="Dialog">

<!--#include file="../libroutines.asp"-->
<%
	
	Dim action
	Dim conn 
   	Dim sqlStr
   	Dim rs
   	Dim ID
   	Dim rsEdit
	Dim volComm
	Dim volBound
	Dim minComm
	Dim cma
	Dim imobRate
	Dim secImob
	Dim regularComm
	Dim LevyRate
	Dim orderIsSaleType	
	Dim commission
	Dim agentCommission
	Dim staffCommission

	Set TimeLimitRs = CreateObject("ADODB.Recordset")   						        
   	TimeLimitRs.CursorLocation = adUseClient
	
	UserId=Session("UserID")	 	
	action = ucase(Request.Form("action"))
	ID = Request("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		window.self.close
                </script>
                <% response.end
        End If

	if action = "EXECUTE" then
	  
	  cmdCancel = Request.Form("cmdCancel")
        Set conn = GetActiveConnection("KBroker")
        
	  LevyValue=Request.Form("txtValue")
      Fieldtype=Cint(Request.Form("cboField"))
      
        If cmdCancel <> "" Then
			
			WriteDialogCancelScript
			Set Conn = Nothing
			Response.End
        End If      
      
        		If Trim(LevyValue) = "" Then
        				if(Fieldtype=1) then
        				%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Levy Rate."
				         		
				         </script>
				         <% response.end
				         else
				         %>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Levy Amount."
				         		
				         </script>
				         <% response.end				        
				         end if
				 End If
				 'ensure Slip is numeric
				If (Not IsNumeric(LevyValue)) Then
						if(Fieldtype=1) then
						%>
						<script language = 'vbscript'>
						ShowMessage "Levy Rate. must be numeric"
						
						</script>
						<% response.end
						else
						%>
						<script language = 'vbscript'>
						ShowMessage "Levy Amount. must be numeric"
						
						</script>
						<% response.end						
						end if
				End If


		
        'save data
        sqlStr="SELECT dbo.OrdDetailList.*, dbo.Lots.Contract_DPA_,dbo.Lots.LotGrossAmount FROM dbo.OrdDetailList INNER JOIN" & _
               " dbo.Lots ON dbo.OrdDetailList.OrdDetail_DPA_ = dbo.Lots.OrdDetail_DPA_ WHERE (dbo.Lots.Contract_DPA_ = " & ID & ")"
        
	  Set Rs=Conn.Execute(sqlStr)
  	  
	  if not(rs.eof or rs.bof) then
		volComm=rs("VolumeRate")
		volBound=rs("VolumeBoundary")
		minComm=rs("MinimumCommission")
		cma=rs("CMARegulated")
		imobRate=rs("PostImmobilisedRate")
		secImob=rs("SecurityImmobilised")		
		orderIsSaleType=rs("OrderTypeSale")			
		agentCommission=rs("AgentCommission")
		orderSecType=rs("ordDetailSectype") 
		grossAmount=rs("LotGrossAmount")
	  end if	  
		commission=LevyValue
		regularComm=LevyValue
        'Apply broker commission
		if(Fieldtype=1) then				
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
												regularComm = LevyValue
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
													commission = LevyValue
												end if
						

										else
												regularComm = LevyValue
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
						Dim dblPrevCommission
						Dim dblCommissionDifference
						Dim dblVATSubtract
						Dim dblVATRate

						commissionAmount = CCur((commission/100.00) * grossAmount)
												
						if commissionAmount < minComm then
							commissionAmount = minComm
						end if
						
						else
							commissionAmount=LevyValue
							commission=(commissionAmount*100)/grossAmount
						end if		
						levyAmount = commissionAmount
						tmpLevy = levyAmount								
						
						'Recalculate the VAT amount*************************
						sqlStr = "SELECT LevyAmount FROM LevyContract WHERE SystemMaintained = 11 AND Contract_DPA_=" & ID
						dblPrevCommission = 0
						dblVATSubtract = 0
						dblVATRate = 0

						If Not conn.Execute(sqlStr).EOF then
							dblPrevCommission = conn.Execute(sqlStr)(0)
						End If

						dblCommissionDifference = dblPrevCommission - commissionAmount
						
						If dblCommissionDifference <> 0 then
							sqlStr = "SELECT LevyRate FROM LevyContract WHERE LevyShortName = 'VAT' AND Contract_DPA_ = " & ID
							If Not conn.Execute(sqlStr).EOF then dblVATRate = conn.Execute(sqlStr)(0)

							dblVATSubtract = CCur((dblVATRate/100.00) * dblCommissionDifference)
						End If

						If dblVATSubtract < 0 then
							dblVATSubtract = dblVATSubtract * -1
							sqlStr="update LevyContract Set LevyAmount=LevyAmount+"& RoundPoint05(dblVATSubtract) &" where LevyShortName = 'VAT' and Contract_DPA_=" & ID
						Else
							sqlStr="update LevyContract Set LevyAmount=LevyAmount-"& RoundPoint05(dblVATSubtract) &" where LevyShortName = 'VAT' and Contract_DPA_=" & ID
						End If

						conn.Execute(sqlStr)
						'END Recalculate the VAT amount*************************

						conn.BeginTrans					
						sqlStr="update LevyContract Set LevyAmount="& RoundPoint05(levyAmount) &",LevyRate="& commission & "," & _
  							 "LevyRatePercentage='" & commission & "%" & "' ,ChangedBy="& UserId & ",TimeChanged=GetDate() where SystemMaintained=11 and Contract_DPA_=" & ID
        				
						

                		conn.Execute sqlStr
						
					levyAmount = CCur((agentCommission/100.00) * tmpLevy)

					sqlStr="update LevyContract Set LevyAmount="& RoundPoint05(levyAmount) &",LevyRate="& agentCommission * (commission/100.00) & "," & _
  							 "LevyRatePercentage='" & agentCommission & "%" & "',ChangedBy="& UserId & ",TimeChanged=GetDate() where SystemMaintained=12 and Contract_DPA_=" & ID
        					
                conn.Execute sqlStr	  				
        conn.CommitTrans
               
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript
   	end If
   	

   	
   
%>
<BR>

<form name = 'frmEditActvtyClass' id="frmMain" method = 'post' action = 'EditBrokerCommission.asp' >
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<%
        Set conn = GetActiveConnection("KBroker")
       
        
        sqlStr = "SELECT * FROM [BrokerCommissionList] WHERE Contract_DPA_  = " & ID
        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected Contract does not exist"
                		window.self.close
                </script>
                <% response.end
        End If
        
        

%>
  <tr>
    <td width="30%">Client</td>
    <td width="70%">&nbsp;&nbsp;<input readonly = 'true' class=readonly  STYLE="WIDTH: 200px; text-align: left" type = 'text' name ='txtClient' id = "txtClient" value = '<%=rs.Fields("ClientName")%>' size="20"></td>
  </tr>
  <tr>
    <td width="30%">Security</td>
    <td width="70%">&nbsp;&nbsp;<input readonly = 'true' class=readonly  STYLE="WIDTH: 200px; text-align: left" type = 'text' name ='txtSecurity' id = 'txtSecurity' size="20" value = '<%=rs.Fields("SecurityCode")%>'></td>
  </tr>
  <tr>
    <td width="30%">Order No</td>
    <td width="70%">&nbsp;&nbsp;<input readonly = 'true' class=readonly  STYLE="WIDTH: 200px; text-align: left" type = 'text'  name ='txtOrder' id = 'txtOrder' size="20" readonly value = '<%=rs.Fields("Order_DPA_")%>'></td>
  </tr>
  <tr>
    <td width="30%">Contract Number</td>
    <td width="70%">&nbsp;&nbsp;<input readonly = 'true' class=readonly  STYLE="WIDTH: 200px; text-align: left" type = 'text'  name ='txtContract' id = 'txtContract' size="20" value = '<%=rs.Fields("ContractNumber")%>'></td>
  </tr>
  <tr>
    <td width="30%">Quantity</td>
    <td width="70%">&nbsp;&nbsp;<input readonly = 'true' class=readonly  STYLE="WIDTH: 200px; text-align: left" type = 'text'  name ='txtQuantity' id = 'txtQuantity' size="20" readonly value = '<%=formatNumEx(rs.Fields("LotQty"),0)%>'></td>
  </tr>
  <tr>
    <td width="30%">Price</td>
    <td width="70%">&nbsp;&nbsp;<input readonly = 'true' class=readonly  STYLE="WIDTH: 200px; text-align: left" type = 'text'  name ='txtPrice' id = 'txtPrice' size="20" value = '<%=FormatNum(rs.Fields("LotPrice"))%>'></td>
  </tr>
  <tr>
    <td width="30%">Gross</td>
    <td width="70%">&nbsp;&nbsp;<input readonly = 'true' class=readonly  STYLE="WIDTH: 200px; text-align: left" type = 'text'  name ='txtGross' id = 'txtGross' size="20" readonly value = '<%=formatNum(rs.Fields("LotGrossAmount"))%>'></td>
  </tr>    	
  <tr>
    <td width="30%">Commission Amount</td>
    <td width="70%">&nbsp;&nbsp;<input readonly = 'true' class=readonly  STYLE="WIDTH: 200px; text-align: left" type = 'text'  name ='txtAmount' id = 'txtAmount' size="20" readonly value = '<%=formatNum(rs.Fields("LevyAmount"))%>'></td>
  </tr>
  <tr>
    <td width="30%">Commission Rate</td>
    <td width="70%">&nbsp;&nbsp;<input readonly = 'true' class=readonly STYLE="WIDTH: 200px; text-align: left" type = 'text'  name ='txtLevy' id = 'txtLevy' size="20" value = '<%=FormatNum(rs.Fields("LevyRate"))%>'></td>
  </tr>
  <tr>
  <td width="30%">Field</td>
  <td>&nbsp;&nbsp;<select name="cboField" onchange='ChangePrice(this);'>			
			<option selected SearchCode = "1" SearchText = "Commission Rate" value = '1'>Commission Rate</option>			
			<option SearchCode = "2" SearchText = "Commission Amount" value='2'>Commission Amount</option>			
		</select>
      </td>
  </tr>
  <tr>
    <td width="30%">Value</td>
    <td width="70%">&nbsp;&nbsp;<input STYLE="WIDTH: 200px; text-align: left" type = 'text'  name ='txtValue' id = 'txtValue' size="20" value = '<%=FormatNum(rs.Fields("LevyRate"))%>'></td>
  </tr>      
  <tr>
     <td width="100%" COLSPAN=2 align="right" valign=absBottom>
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save">
		&nbsp;
		<input type = 'button' Class=Buttons name ='cmdClose' id = 'cmdClose' value=" Cancel " OnClick="JavaScript: window.self.close();">		
		&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">		
    </td>
  </tr>
</table>
</form>

</body>

</html>
