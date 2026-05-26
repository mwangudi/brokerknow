<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Security</title>
 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>

<!--CALENDAR -->
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>
<script language="JavaScript" src="CALENDAR/calendar.js"></script>

<!--END CALENDAR -->


<script language="javascript">
		function  UpdateImmobilised(theChk)
		{
			if (theChk.checked)
			{
				document.frmMain.elements("txtImmobilised").value = "1";
			}
			else
			{
				document.frmMain.elements("txtImmobilised").value = "0";
			}
				
			
		}
		function  UpdateCanTrade(theChk)
		{
			if (theChk.checked)
			{
				document.frmMain.elements("txtCanTrade").value = "1";
			}
			else
			{
				document.frmMain.elements("txtCanTrade").value = "0";
			}
				
			
		}
</script>
</head>

<body Class="Dialog">

<!--#include file="../libroutines.asp"-->
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>

<%
	 
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim InterBankRs
   Dim guidStr 
   Dim guid 
	
	action = ucase(Request.Form("action"))
	ID=Request("ID")
	
	
	If Trim(ID)="" OR ID=0 Then
    %>
                <script language = 'vbscript'>
                		ShowMessage "No contracts To make interbanks"
                window.self.close		
                </script>
                <%
            	response.end
    end if


	if action = "EXECUTE" then
	   Dim SourceEntityType_DPA_
       Dim SourceEntity_DPA_
       Dim TargetEntityType_DPA_
       Dim TargetEntity_DPA_
       Dim TransferAmount
       Dim TransferDate
       Dim ChangedBy       
       Dim Reference
       Dim Narrative
       Dim InterTransferType_DPA_
       Dim InterTransfer_EIT_
       
       Set conn = GetActiveConnection("KBroker")
       
       Set InterBankRs = Server.CreateObject("ADODB.Recordset")
		
		InterBankRs.CursorLocation = adUseClient
		
       
       sqlStr="SELECT InterBankTransferList.Client_DPA_ AS SourceEntity_DPA_, Client.EntityType_DPA_ AS SourceEntityType_DPA_, InterBankTransferList.LotGrossAmount, " & _
              "InterBankTransferList.OrderTypeSale FROM InterBankTransferList INNER JOIN Client ON InterBankTransferList.Client_DPA_ = Client.Client_DPA_" & _
			  " WHERE (InterBankTransferList.Contract_DPA_ = " & ID & ")"
        
        Set InterBankRs=conn.Execute(sqlStr)		
		
		if Not(InterBankRs.EOF and InterBankRs.bof) then
			if(InterBankRs("OrderTypeSale")=0) then
				SourceEntityType_DPA_=InterBankRs("SourceEntityType_DPA_")
				SourceEntity_DPA_=InterBankRs("SourceEntity_DPA_")
				TargetEntityType_DPA_=8
				TargetEntity_DPA_=14
			else
				SourceEntityType_DPA_=8
				SourceEntity_DPA_=14
				TargetEntityType_DPA_=InterBankRs("SourceEntityType_DPA_")
				TargetEntity_DPA_=InterBankRs("SourceEntity_DPA_")
			end if
         TransferAmount=InterBankRs("LotGrossAmount")
         TransferDate=Request.Form("txtADate")
         ChangedBy=Session("UserID")       
         Reference=Request.Form("txtReference")
         Narrative=Request.Form("txtNarrative")
         InterTransferType_DPA_=1       
		end if		 
            
        'save header
        set guid = server.createobject("NDUtils.CGUID")
        InterTransfer_EIT_ = guid.GenerateGUID
        
       sqlStr = "INSERT INTO [InterTransfer] (SourceEntityType_DPA_,InterTransfer_EIT_,SourceEntity_DPA_" & _
				"       ,TargetEntityType_DPA_,TargetEntity_DPA_,TransferAmount,TransferDate,ChangedBy,TimeChanged,TransferReference,TransferNarrative,InterTransferType_DPA_,Contract_DPA_)" & _
				" Values( " & "'" & SourceEntityType_DPA_ & "'" & " " & _
				"       ," & "'" & InterTransfer_EIT_ & "'" & " " & _
				"       ," & " " & SourceEntity_DPA_ & " " & "," & " " & TargetEntityType_DPA_ & " " & " " & _
				"		," & " " & TargetEntity_DPA_ & " " & " " & _
				"		," & " " & TransferAmount & " " & " " & _
				"		," & "#" & FormatDate(TransferDate) & "#" & " " & _
				"		," & " " & ChangedBy & " " & " " & _
				"		," & "'" & FormatDate(Now()) & "'" & " " & _
				"		," & "'" & Reference & "'" & " " & _
				"		," & "'" & Narrative & "'" & " " & _
				"		," & " " & InterTransferType_DPA_ & " " & " " & _				
				"       ," & " " & ID & " " & ")"
        
        
        'Response.Write(sqlStr)
        'Response.End 
        
        
        
        conn.BeginTrans
			conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))		
		conn.CommitTrans
        Set Conn = Nothing
        WritefraEnabledDialogCloseScript
        Response.End 
   	end If
   	
%>

<form name = 'frmAddSecurity' method = 'post' id="frmMain" action = "AddInterBank.asp" >
<table border="0" width="100%" cellpadding=2 cellspacing=2>
 <tr>
 <%
 Set conn = GetActiveConnection("KBroker")
        
	 sqlStr = "SELECT ContractNumber From LotList WHERE Contract_DPA_  = " & ID 
    Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
    If Not(rs.EOF Or rs.BOF) Then
    ContractNumber=rs("ContractNumber")
    end if
Set rs=nothing

sqlStr = "SELECT Contract_DPA_ From InterTransfer WHERE Contract_DPA_  = " & ID 

Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
    
    If Not(rs.EOF Or rs.BOF) Then
    %>
                <script language = 'vbscript'>
                		ShowMessage "Use Edit button to effect changes"
                window.self.close		
                </script>
                <%
                'WriteDialogRelocateScript "AddInterBank.asp?ID=" & ID 
                'WritefraEnabledDialogCloseScript
                response.end
    end if

%>
    
                <td width="30%">Contract</td>
                <td width="70%"><input readonly = 'true' class=readonly type="text" name="txtContract" id="txtContract" readonly size="25" value='<%=ContractNumber%>'></td>
              </tr>
              <tr>
				<td width="30%">Date</td>
				<td width="70%">
				<SCRIPT language="JavaScript">			
				var cal=new ctlSpiffyCalendarBox("cal", "frmAddSecurity", "txtADate","cmdDate","<%= FormatDate(Date) %>",1);
				cal.writeControl();
				</SCRIPT>		
				</td>
				</tr>
              <tr>
                <td width="30%">Reference</td>
                <td width="70%"><input type="text" name="txtreference" id="txtreference" size="25"></td>
              </tr>
              <tr>
                <td width="30%">Narrative</td>
                <td width="70%">
                <textarea rows=3 name ='txtNarrative' id = "txtNarrative" cols="20"></textarea></td>
              </tr>
 
  <tr>
     <td width="100%" colspan="2" align=right>
		<BR>
		<BR>
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
		&nbsp;&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
		<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">		
      </td>
  </tr>
</table>
</form>


</body>

</html>