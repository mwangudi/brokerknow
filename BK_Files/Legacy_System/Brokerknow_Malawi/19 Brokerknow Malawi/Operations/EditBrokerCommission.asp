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

	function forceSubmit()
	{
		setOpener();
		var targetPage = window.dialogArguments.opener.document.all.item("frmMain").elements("EditPage").value		
		document.frmEditActvtyClass.method='post';
		document.frmEditActvtyClass.target='_self';
		document.frmEditActvtyClass.submit();
	}
	
	function setOpener()
	{
		window.self.opener = window.dialogArguments.opener;
	}
	
</script>
</head>

<body Class="Dialog" onLoad="javascript: setOpener()">

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
	  
	  cmdCancel = Request.Form("buttonAction")
        Set conn = GetActiveConnection("KBroker")
     
	  LevyValue=Request.Form("txtValue")
      Fieldtype=Cint(Request.Form("cboField"))
      
        If cmdCancel = "" Then
			
			WriteDialogCancelScript2
			Set Conn = Nothing
			Response.End
        End If      
      
        		If Trim(LevyValue) = "" Then
        				if(Fieldtype=1) then
        				%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Levy Rate."
				         		
				         </script>
				         <% 
						 ReloadPage(ID)
						 response.end
				         else
				         %>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Levy Amount."
				         		
				         </script>
				         <% 
						ReloadPage(ID)
						 response.end				        
				         end if
				 End If
				 'ensure Slip is numeric
				If (Not IsNumeric(LevyValue)) Then
						if(Fieldtype=1) then
						%>
						<script language = 'vbscript'>
						ShowMessage "Levy Rate. must be numeric"
						
						</script>
						<% 
						ReloadPage(ID)
						response.end
						else
						%>
						<script language = 'vbscript'>
						ShowMessage "Levy Amount. must be numeric"
						
						</script>
						<% 
						ReloadPage(ID)
						response.end						
						end if
				End If


		
        'save data
        
		

		procstr =  ID & ", " & cdbl(LevyValue) & " ," & Fieldtype & "," & session("UserID") & "" 
  	   %>
				         <script language = 'vbscript'>
				         		'ShowMessage "<%=procstr%>"
				         		
				         </script>
				         <% 'response.end
	 	conn.begintrans
			conn.execute (sqlserverformat("cont_EditBrokerComm  " & procstr))
        conn.CommitTrans
               
        conn.Close
        Set conn = Nothing
		%>
		<SCRIPT LANGUAGE="JAVASCRIPT">
			window.opener.location= window.opener.location;
		</script>
		<%
        WritefraEnabledDialogCloseScript2
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
  <td width="30%">Field(Excl of tax)</td>
  <td>&nbsp;&nbsp;<select name="cboField" onchange='ChangePrice(this);'>			
			<option  SearchCode = "1" SearchText = "Commission Rate" value = '1'>Commission Rate</option>			
			<option selected SearchCode = "2" SearchText = "Commission Amount" value='2'>Commission Amount</option>			
		</select>
      </td>
  </tr>
  <tr>
    <td width="30%">Value</td>
    <td width="70%">&nbsp;&nbsp;<input STYLE="WIDTH: 200px; text-align: left" type = 'text'  name ='txtValue' id = 'txtValue' size="20" value = '<%=FormatNum(rs.Fields("LevyAmount"))%>'></td>
  </tr>      
  <tr>
     <td width="100%" COLSPAN=2 align="right" valign=absBottom>
		<BR>
		<input type = 'button' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save" onClick= "javascript: forceSubmit();">
		&nbsp;
		<input type = 'button' Class=Buttons name ='cmdClose' id = 'cmdClose' value=" Cancel " OnClick="JavaScript: window.self.close();">		
		&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
		<input type = 'hidden' name ='buttonAction' id = 'action' value="Save">		
    </td>
  </tr>
</table>
</form>

</body>

</html>
