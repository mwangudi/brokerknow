<!--#include file="../libroutines.asp"-->
<%

'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "StockWatchList"
		const DataEntity = "StockWatchList"
		const DataEntityPlural = "StockWatchLists"
		const ActionFolder = "Operations"
		const ActionPage = "StockWatchList"
'======================= End_Alter_Across_Entities =================================
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim guidStr 
   Dim guid 
   Dim buttonAction
	dim currentEntityType
	
	
	action = ucase(Request.Form("action"))
	ID = Request.Form("ID")
	currentEntityType = 1
	UserID=Session("UserID")
	select case action
		case "EXECUTE"
		
				Dim reloadRequired
		
				reloadRequired = false
				buttonAction = Trim(Trim(Ucase(Request.Form("cmdAdd"))))
		
				if instr(1,buttonAction,"CONTINUE") > 0 then
						Dim client
						Dim security
						Dim user
					 
						user = Session("UserID")
						client= Request.Form("cboClient")
						security= Request.Form("cbosecurity")
								
						 'validate client
						 If Trim(client) = "" Then%>
						         <script language = 'vbscript'>
						         		ShowMessage "Please specify the client"
						         		
						         </script>
						         <% response.end
						 End If
						 'validate security
						 If Trim(security) = "" Then%>
						         <script language = 'vbscript'>
						         		ShowMessage "Please specify the Security"
						         		
						         </script>
						         <% response.end
						 End If

						   
						 'save header				 
						 set guid = server.createobject("NDUtils.CGUID")
						 guidStr = guid.GenerateGUID
						 
						 sqlStr = "INSERT INTO [StockWatch] (Client_DPA_,Security_DPA_,StockWatch_DPA_,StockWatch_EIT_,ChangedBy) SELECT " & "#" & client & "#" & " as client_DPA_," & " " & security & " " & " as Security_DPA_," & _
						         "iif(isnull(max([StockWatch_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'StockWatch'),max([StockWatch_DPA_]) + 1)" & " " & " as StockWatch_DPA_" & _
						         "," & "'" & guidStr & "'" & " as StockWatch_EIT_" & _
						         "," & " " & UserId & " " & " as ChangedBy FROM [StockWatch]"
						 Set conn = GetActiveConnection("KBroker")
						 conn.BeginTrans

						sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
								
							
								conn.Execute sqlStr
						     
								'obtain header key value
								sqlStr = "SELECT StockWatch_DPA_] FROM [StockWatch] WHERE [StockWatch_EIT_] = " & "'" & guidStr & "'"
						     
								Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
								If (rs.EOF Or rs.BOF) Then%>
						         			<script language = 'vbscript'>
						         					ShowMessage "A serious error has been encountered while saving the data. Try saving again"
						         					
						         			</script>
						         			<% response.end
								End If
						     conn.CommitTrans
						%>
						<Script Language="JavaScript">
							try{
								window.parent.dialogArguments.opener.location.reload();
								}
							catch(e){window.self.close()}
						</Script>
							<%
						
						WriteDialogRelocateScript "EditStockWatch.asp?ID=" & rs.Fields("StockWatch_DPA_")& "-"& client & ""
						Response.end
				end if
		case "FETCH_ACCOUNTS"
			currentEntityType = cint(ID)
   	end select
   	
   	
%>

<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add <%=DataEntity%></title>

<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
	<SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>
	<SCRIPT language=Javascript src="../scripts/accountList.js"></SCRIPT>
	

<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frm<%=DataSource%>", "txtDate","cmdDate","<%=FormatDate(Date)%>",1);
	var calJournalDate=new ctlSpiffyCalendarBox("calJournalDate", "frm<%=DataSource%>", "txtJournalDate","cmdJournalDate","<%=FormatDate(Date)%>",1);
</SCRIPT>
<!--END CALENDAR -->

<script language='vbscript'>

					function EntitySelected(itemID)
 							frm<%=DataSource%>.elements("ID").value = itemID
 							frm<%=DataSource%>.elements("action").value = "Fetch_Accounts"
 							frm<%=DataSource%>.submit
 							
					end function

</script>
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
		
		function UpdateCertField(theList)
		{
			//handle the certificate
			var i = 0;
			for (i=0; i < theList.options.length; i++) {
				if((theList.options(i).selected))
				{
					if(theList.options(i).RequireCertificate == "True")
					{
						document.frmMain.elements("txtCert").disabled = false;
					}
					else
					{
						document.frmMain.elements("txtCert").disabled = true;
						document.frmMain.elements("txtCert").value = "";
					}
				}
			}
		}
		
		function UpdateSecurityListing(theList)
		{
			//swap lists
			if(theList.currentSecType == "S")
			{
				document.frmMain.elements("cboFixed").style.display = "block";
				document.frmMain.elements("cboFixed").name = "cboSecurity";
				
				document.frmMain.elements("cboSecurity").style.display = "none";
				document.frmMain.elements("cboSecurity").name = "cboSecurityHidden";
				
				theList.currentSecType = "F"
			}
			else
			{
				document.frmMain.elements("cboFixed").style.display = "none";
				document.frmMain.elements("cboFixed").name = "cboSecurityHidden";
				
				document.frmMain.elements("cboSecurity").style.display = "block";
				document.frmMain.elements("cboSecurity").name = "cboSecurity";
				
				theList.currentSecType = "S"
			}
		}
		
		function FetchAccounts(theList)
		{
			var i = 0;
			var entity = theList.value;
			var toList = document.frmMain.cboAccount;
			
			frm = document.frmMain;				
			xmlhttp = createXMLHTTPObj();
			
			url="GetList.asp?ID="+entity+"&action=GetAccountList";
			xmlhttp.open("GET",url,true);
			xmlhttp.onreadystatechange=function() {
				if (xmlhttp.readyState==4) {
				returnStr = xmlhttp.responseText;
				returnStr = getBodyHTML(returnStr);
			
				var secList = "<select name = '" + toList.name + "' id = '" + toList.name + "' size='1' ";
				secList += "OnClick='event.cancelBubble=true;' " ;
				secList += "onChange='event.cancelBubble=true;' " ;
				secList += "onKeypress='return (dodefaultaction()==\"\"); ' "  ;
				secList += "onKeydown='return (dodefaultaction()==\"\");' " ; 
				secList += "onKeyup='return (change(" + toList.name + "));' " ; 
				secList += "onfocus='txtval = \"\";inputIsItemCode = 1;' "  ;
				secList += "onblur='txtval = \"\";inputIsItemCode = 1;'>" ;
				secList += returnStr ;
				secList += "</select>";
				
				toList.outerHTML = secList;														
				}
				}
			xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
			xmlhttp.send(); 
		
		
		}

		
				function UpdatePrice(){
					
					var security = document.frm<%=DataSource%>.elements("cboSecurity");
					
					document.frm<%=DataSource%>.elements("txtPrice").value = security[security.selectedIndex].price;
					
							
					}
</script>
</head>

<body Class="Dialog">

<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>

<form name = 'frm<%=DataSource%>' method = 'post' action = 'AddStockWatch.asp' id = 'frmMain'>


<table border="0" width="100%">
 <tr>		<td>Client</td>
			<td nowrap height="10">
			<input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="10" onBlur="txtval = this.value; selectItem(cboClient);UpdateCodes(true,cboClient,txtCdsNo);">
			<input type = 'text' name ='txtCdsNo' id = 'txtCdsNo' size="16" onBlur="txtval = this.value; selectItems(cboClient);UpdateCode(true,cboClient,txtClientCode);">
			<select name = 'cboClient' id = 'cboClient' size="1" 
					onKeypress="return (dodefaultaction()==''); " 
					onKeydown="return (dodefaultaction()==''); " 
					onKeyup="return (UpdateCode(change(cboClient,0),cboClient,txtClientCode));" 
					onChange="UpdateCode(true,cboClient,txtClientCode);UpdateCodes(true,cboClient,txtCdsNo);"
					onfocus="txtval = '';inputIsItemCode = 1;" 
					onblur="txtval = '';inputIsItemCode = 1;" readonly>
				    	
					<%
										
					Set Conn = GetActiveConnection("KBroker")
					
					Set rs = Server.CreateObject("ADODB.Recordset")

					sqlStr = "SELECT * FROM [Client] where Deleted=0 order by RTrim(Ltrim(ClientName))"
					sqlStr = "SELECT     * " & _
							" FROM Client " & _
							" WHERE (Deleted <>1) " & _
							" ORDER BY ClientName ASC"
'					Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					Set rs = conn.Execute(sqlStr)
					If Not (rs.EOF Or rs.BOF) Then
						    rs.MoveFirst
					        Do Until rs.EOF					                
								        
					        ClientName=rs.Fields("ClientName")
						    NameClient=Mid(ClientName,1,30)
								    %>                    
					                <option SearchCode = "<%=rs.Fields("Client_DPA_")%>" SearchText = "<%=rs.Fields("ClientName")%>" SearchCds = "<%=rs.Fields("ClientCDSNo")%>" value = '<%=rs.Fields("Client_DPA_")%>'><%=NameClient%></option>
					                <%rs.MoveNext
					        Loop
					End If
					
					Set rs = Nothing
					Set Conn = Nothing
					
					%>
			</select>		
			</td>
		</tr>
  <tr>
    <td width="10%">Security</td>
    <td width="50%"><select name="cboSecurity" onchange="UpdatePrice();">

					<option value="" price="">Select a security</option>

					<%
						set rs=server.createobject("Adodb.recordset")
						Set Conn = GetActiveConnection("KBroker")

						rs.open "Select * from datastream_SecurityPriceList order by security_DPA_",conn, 0,1

						while not rs.eof 
					%>
						<option value='<%=rs("Security_DPA_")%>' price='<%=formatnum(rs("Price"))%>'><%=rs("SecurityCode")%></option>
					<%
						rs.movenext
						wend
						rs.close
						set rs= nothing
						set conn= nothing
					%>
				</Select>
	</td>
  </tr>
  <tr><td>Price</td><td><input type="text" name="txtPrice" style="text-align:right" value="0" readonly class="readonly"></td></tr>
 </table>

<table>
	  <tr>
    <td width="20%" colspan=4 align=right>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Continue " onclick = "AllowedNavigation()">
        &nbsp;&nbsp; <input type = 'button' Class=Buttons name ='cmdClose' id = "cmdClose" value=" Close " onclick = "window.self.close();">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ID' id = 'ID'>
	</td>
  </tr>
</table>
</form></body>

</html>