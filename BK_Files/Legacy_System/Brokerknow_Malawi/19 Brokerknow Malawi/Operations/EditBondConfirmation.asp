<!--#include file="../libroutines.asp"-->
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Bond Confirmation</title>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>

<!--END CALENDAR -->
<script language="javascript">
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
</script>
</head>
<%
	
	Dim action
	Dim conn 
   	Dim sqlStr
   	Dim rs   	
   	Dim SettlementDate
	Dim ForwardRate
	Dim NoDaysMaturity
	Dim Convexity
	Dim Duration
	Dim ConsiderationI
	Dim CommissionI
	Dim CounterPartyI
	Dim CounterParty
	Dim Contact	
	Dim Title	

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
        
	  LevyRate=Request.Form("txtLevy")

        If cmdCancel <> "" Then
			
			WriteDialogCancelScript
			Set Conn = Nothing
			Response.End
        End If   
       
        SettlementDate=Request.Form("txtDate")
		ForwardRate=Request.Form("txtfrate")
		NoDaysMaturity=Request.Form("txtNoDays")
		Convexity=Request.Form("txtConvexity")
		Duration=Request.Form("txtDuration")
		ConsiderationI=Request.Form("txtConsiderationP")
		CommissionI=Request.Form("txtCommissionP")
		CounterPartyI=Request.Form("txtCounter")
		CounterParty=Request.Form("txtClientCode")
		Contact	=Request.Form("txtContact")
		Title	=Request.Form("txtTitle")

				'Check for Forward Rate
				 
				 'validate Order Hold Option 
				 If Trim(SettlementDate) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Settlement Date"
				         </script>
				         <% response.end
				 End If			 	 
				
				 if(ForwardRate="") then
					ForwardRate=0
				 end if
				 if(Convexity="") then
					Convexity=0
				 end if
				
				 if(Duration="") then
					Duration=0
				 end if
				 if(CounterParty="") then
					CounterParty=0
				 end if
				if(NoDaysMaturity="") then
					NoDaysMaturity=0
				 end if
				if(ForwardRate="") then
					ForwardRate=0
				 end if


					conn.BeginTrans					
						sqlStr = "Update [BondConfirmation] " & _
				         "Set SettlementDate='" & FormatDate(SettlementDate) & "'" & _
				         ",ForwardRate=" & ccur(ForwardRate)  & _
				         ",NoDaysMaturity=" & " " & NoDaysMaturity & " " & " " & _
				         ",Convexity=" & " " & ccur(Convexity) & " " & " " & _
				         ",Duration=" & " " & Duration & " " & " " & _
				         ",ConsiderationI='" & ConsiderationI & "'" & _
				         ",CommissionI='" & CommissionI & "'" & _
				         ",CounterPartyI='" & CounterPartyI & "'"  & _
				         ",CounterParty_DPA_=" & " " & CounterParty & " " & " " & _
				         ",ContactPerson='" & Contact & "'"  & _
				         ",Title='" & Title & "'"  & _
				         "where Order_DPA_=" & ID	         
						

                			'sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))				
							
											
						    conn.Execute sqlStr	
					
        conn.CommitTrans
               
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript
   	end If
   	

Set conn = GetActiveConnection("KBroker")
       
        
        sqlStr = "SELECT * FROM BondConfirmationList INNER JOIN " & _
                 " BondConfirmation ON BondConfirmationList.Order_DPA_ = BondConfirmation.Order_DPA_ WHERE Deleted=0 and BondConfirmation.Order_DPA_  = " & ID
        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If (rs.EOF Or rs.BOF) Then
		 %>
                <script language = 'vbscript'>
                		ShowMessage "Use New button to add bond confirmation"
                window.self.close		
                </script>
                <%
                'WriteDialogRelocateScript "AddInterBank.asp?ID=" & ID 
                'WritefraEnabledDialogCloseScript
                response.end
		end if
        
           	
   
%>
<body Class="Dialog">
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frmAddBondConfirmation", "txtDate","cmdDate","<%=FormatDate(rs("SettlementDate"))%>",1);
	var calValidity=new ctlSpiffyCalendarBox("calValidity", "frmAddBondConfirmation", "txtValidity","cmdValDate","<%=FormatDate(Date)%>",1);
	var calReleaseDate=new ctlSpiffyCalendarBox("calReleaseDate", "frmAddBondConfirmation", "txtReleaseDate","cmdReleaseDate","<%=FormatDate(Date)%>",1);
</SCRIPT>
<form name = 'frmAddBondConfirmation' id="frmMain" method = 'post' action = 'EditBondConfirmation.asp' >
<tr>
     <td width="100%" COLSPAN=2 align="right" valign=absBottom>		
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save">
		&nbsp;
		<input type = 'button' Class=Buttons name ='cmdClose' id = 'cmdClose' value=" Cancel " OnClick="JavaScript: window.self.close();">		
		&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
    </td>
 </tr>
 <tr>
 <td colspan="2">&nbsp;</td>
 </tr>
 <table width="100%" cellspacing="0" cellpadding="0">
 <tr>
 <td>
<p><b>Contract Details</b></p>
<table border="1" width="90%" cellspacing="0" cellpadding="0" bordercolor="#000000">
  <tr>
    <td width="35%">&nbsp;&nbsp;Trade Date</td>
    <td width="65%"><input readonly = 'true' class=readonly  STYLE="WIDTH: 150px; text-align: left" type = 'text' name ='txtTransDate' id = "txtTransDate" value = '<%=FormatDate(rs.Fields("TransDate"))%>'></td>
  </tr>
  <tr>
    <td width="35%">&nbsp;&nbsp;Client</td>
    <td width="65%"><input readonly = 'true' class=readonly  type = 'text' name ='txtClient' id = 'txtClient' size="30" value = '<%=rs.Fields("ClientName")%>'></td>
  </tr>
  <tr>
    <td width="30%">&nbsp;&nbsp;Order No</td>
    <td width="70%"><input readonly = 'true' class=readonly  STYLE="WIDTH: 150px; text-align: left" type = 'text'  name ='txtOrder' id = 'txtOrder' size="20" readonly value = '<%=rs.Fields("Order_DPA_")%>'></td>
  </tr>
  <tr>
    <td width="30%">&nbsp;&nbsp;Contract No</td>
    <td width="70%"><input readonly = 'true' class=readonly  STYLE="WIDTH: 150px; text-align: left" type = 'text'  name ='txtContract' id = 'txtContract' size="20" value = '<%=rs.Fields("ContractNumber")%>'></td>
  </tr>
  <tr>
    <td width="30%">&nbsp;&nbsp;Issue No</td>
    <td width="70%"><input readonly = 'true' class=readonly  STYLE="WIDTH: 150px; text-align: left" type = 'text'  name ='txtIssueNo' id = 'txtIssueNo' size="20" readonly value = '<%=rs.Fields("BondIssue")%>'></td>
  </tr>
  <tr>
    <td width="30%">&nbsp;&nbsp;Ref</td>
    <td width="70%"><input readonly = 'true' class=readonly  STYLE="WIDTH: 150px; text-align: left" type = 'text'  name ='txtSlipNo' id = 'txtSlipNo' size="20" readonly value = '<%=rs.Fields("SlipNo")%>'></td>
  </tr>  
  <tr>
    <td width="30%">&nbsp;&nbsp;Price</td>
    <td width="70%"><input readonly = 'true' class=readonly  STYLE="WIDTH: 150px; text-align: left" type = 'text'  name ='txtPrice' id = 'txtPrice' size="20" value = '<%=FormatNumEx(rs.Fields("Price"),4)%>'></td>
  </tr>
  <tr>
    <td width="30%">&nbsp;&nbsp;Face Value</td>
    <td width="70%"><input readonly = 'true' class=readonly  STYLE="WIDTH: 150px; text-align: left" type = 'text'  name ='txtGross' id = 'txtGross' size="20" readonly value = '<%=formatNum(rs.Fields("FaceValue"))%>'></td>
  </tr>    	
  <tr>
    <td width="30%">&nbsp;&nbsp;Consideration</td>
    <td width="70%"><input readonly = 'true' class=readonly  STYLE="WIDTH: 150px; text-align: left" type = 'text'  name ='txtConsideration' id = 'txtConsideration' size="20" readonly value = '<%=FormatNum(rs.Fields("Gross"))%>'></td>
  </tr>
  <tr>
    <td width="30%">&nbsp;&nbsp;Commission</td>
    <td width="70%"><input readonly = 'true' class=readonly  STYLE="WIDTH: 150px; text-align: left" type = 'text'  name ='txtcommission' id = 'txtcommission' size="20" readonly value = '<%=FormatNum(rs.Fields("Commission"))%>'></td>
  </tr>  
  <tr>
    <td width="30%">&nbsp;&nbsp;Net Amount</td>
    <td width="70%"><input readonly = 'true' class=readonly  STYLE="WIDTH: 150px; text-align: left" type = 'text'  name ='txtNetAmount' id = 'txtNetAmount' size="20" value = '<%=FormatNumEx(rs.Fields("NetAmount"),2)%>'></td>
  </tr>
  <tr>
    <td width="30%">&nbsp;&nbsp;Issue Date</td>
    <td width="70%"><input readonly = 'true' class=readonly  STYLE="WIDTH: 150px; text-align: left" type = 'text'  name ='txtIssue' id = 'txtissue' size="20" readonly value = '<%=formatDate(rs.Fields("IssueDate"))%>'></td>
  </tr>    	  
  <tr>
    <td width="30%">&nbsp;&nbsp;Maturity Date</td>
    <td width="70%"><input readonly = 'true' class=readonly  STYLE="WIDTH: 150px; text-align: left" type = 'text'  name ='txtMaturity' id = 'txtMaturity' size="20" readonly value = '<%=formatDate(rs.Fields("MaturityDate"))%>'></td>
  </tr>   
</table>
</td>
<%
Set rs =nothing 
sql="Select * From BondConfirmation where Order_DPA_=" & ID 
	Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected Contract does not exist"
                		window.self.close
                </script>
                <% response.end
        End If
%>
  <td valign="top">
  <p><b>Confirmation Details</b></p>
  <table border="1" width="100%" cellspacing="0" cellpadding="0" bordercolor="#000000">
  <tr>
    <td width="40%">&nbsp;&nbsp;Settlement Date</td>
    <td width="60%"><SCRIPT language="JavaScript">cal.writeControl();</SCRIPT></td>
  </tr>
  <tr>
    <td width="40%">&nbsp;&nbsp;Forward Rate</td>
    <td width="60%"><input  STYLE="WIDTH: 100px; text-align: left" type = 'text' name ='txtfrate' id = 'txtClient' size="30" value='<%=FormatNum(rs("ForwardRate"))%>'></td>
  </tr>
  <tr>
    <td width="40%">&nbsp;&nbsp;No Of Days To Maturity</td>
    <td width="60%"><input STYLE="WIDTH: 100px; text-align: left" type = 'text'  name ='txtNoDays' id = 'txtNoDays' size="20" value='<%=rs("NoDaysMaturity")%>'></td>
  </tr>
  <tr>
    <td width="40%">&nbsp;&nbsp;Convexity</td>
    <td width="60%"><input STYLE="WIDTH: 100px; text-align: left" type = 'text'  name ='txtConvexity' id = 'txtConvexity' size="20" value='<%=FormatNum(rs("Convexity"))%>'></td>
  </tr>
  <tr>
    <td width="40%">&nbsp;&nbsp;Duration</td>
    <td width="60%"><input STYLE="WIDTH: 100px; text-align: left" type = 'text'  name ='txtDuration' id = 'txtDuration' size="20" value='<%=formatnum(rs("Duration"))%>'></td>
  </tr>
  <tr>
    <td width="40%">&nbsp;&nbsp;Payment Instruction(Consideration)</td>
    <td width="60%"><input STYLE="WIDTH: 270px; text-align: left" type = 'text'  name ='txtConsiderationP' id = 'txtConsiderationP' size="20" value='<%=rs("ConsiderationI")%>'></td>
  </tr>  
  <tr>
    <td width="40%">&nbsp;&nbsp;Payment Instruction(Commission)</td>
    <td width="60%"><input STYLE="WIDTH: 270px; text-align: left" type = 'text'  name ='txtCommissionP' id = 'txtCommissionP' size="20" value='<%=rs("CommissionI")%>'></td>
  </tr>
  <tr>
    <td width="40%">&nbsp;&nbsp;Payment Instruction(Counter Party)</td>
    <td width="60%"><input STYLE="WIDTH: 270px; text-align: left" type = 'text'  name ='txtCounter' id = 'txtCounter' size="20" value='<%=rs("CounterPartyI")%>'></td>
  </tr>    	
  <tr>
    <td width="40%">&nbsp;&nbsp;Counter Party </td>
				<td width="60%"><input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="10" value='<%=rs("CounterParty_DPA_")%>' onBlur="txtval = this.value; selectItem(cboBank);">
				<select name = 'cboBank' id = "cboBank" size="1" 
    					onchange='UpdateCode(true,cboBank,txtClientCode)'
					onKeypress="return (dodefaultaction()==''); " 
					onKeydown="return (dodefaultaction()==''); " 
					onKeyup="return (FilterData(this,1,UpdateCode(change(cboBank,0),cboBank,txtClientCode)));" 
					onfocus="txtval = '';inputIsItemCode = 1;" 
					onblur="txtval = '';inputIsItemCode = 1;">
					<option selected SearchCode = "" SearchText = ""  value = ''></option>
					<%

					        Set conn = GetActiveConnection("KBroker")
					        
					        sqlStr = "SELECT * FROM Client ORDER BY ClientName"
					        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					        If Not (rsEdit.EOF Or rsEdit.BOF) Then
					                rsEdit.MoveFirst
					                Do Until rsEdit.EOF
					                if(rsEdit.Fields("Client_DPA_")=rs("CounterParty_DPA_")) then
					                %>					                        
									<option Selected SearchCode = "<%=rsEdit.Fields("Client_DPA_")%>" SearchText = "<%=rsEdit.Fields("ClientName")%>" value = '<%=rsEdit.Fields("Client_DPA_")%>'><%=Mid(rsEdit.Fields("ClientName"),1,20)%></option>
					                <%
					                else
					                %>					                        
									<option SearchCode = "<%=rsEdit.Fields("Client_DPA_")%>" SearchText = "<%=rsEdit.Fields("ClientName")%>" value = '<%=rsEdit.Fields("Client_DPA_")%>'><%=Mid(rsEdit.Fields("ClientName"),1,20)%></option>
					                <%
					                end if
					                rsEdit.MoveNext
					                Loop
					        End If
					%>

					    </select>
				</td>
  </tr>
  <tr>
    <td width="40%">&nbsp;&nbsp;Contact Person</td>
    <td width="60%"><input readonly = 'true' class=readonly STYLE="WIDTH: 270px; text-align: left" type = 'text'  name ='txtContact' id = 'txtContact' size="20" value='<%=rs("ContactPerson")%>'></td>
  </tr>  
  <tr>
    <td width="40%">&nbsp;&nbsp;Title</td>
    <td width="60%"><input STYLE="WIDTH: 270px; text-align: left" type = 'text'  name ='txtTitle' id = 'txtTitle' size="20" value='<%=rs("Title")%>'></td>
  </tr>
  <tr>
    <td width="40%">&nbsp;&nbsp;Account Manager</td>
    <td width="60%"><input readonly = 'true' class=readonly  STYLE="WIDTH: 270px; text-align: left" type = 'text'  name ='txtAccount' id = 'txtAccount' size="20" value='<%=rs("ClientOwner")%>'></td>
  </tr>      
</table>
  </td>
  </tr>
</table>
</form>

</body>

</html>
