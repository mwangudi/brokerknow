<!--#include file="../libroutines.asp"-->
<%

'======================= Begin_Alter_Across_Entities =================================
		
	const UDLName = "KBroker"
		const DataSource = "StockWatchList"
		const DataEntity = "StockWatch"
		const DataEntityPlural = "StockWatchLists"
		const ActionFolder = "Operations"
	'======================= End_Alter_Across_Entities =================================
		const LinkedIndependent = 1
		const LinkedDependent = 2
		Dim action
		Dim conn 
		Dim sqlStr
		Dim rs
		Dim guidStr 
		Dim guid 
		Dim buttonAction
		dim currentEntityType

	action = ucase(Request.Form("action"))
	ID = Request("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		
                </script>
                <% response.end
        End If   

	currentEntityType = 1

	UserID=Session("UserID")
	select case action
		case "SAVE"
		Set conn = GetActiveConnection("KBroker")
		
				conn.BeginTrans
		'update the detail data
			Dim i
			Dim TotalAmount
			TotalAmount=0
			i=0

        
		sqlStr = "UPDATE StockWatch SET Security_DPA_ = " & request.form("cboSecurity1") & " " & _
				" WHERE StockWatch_DPA_=" & Request.Form("StockNo")	
				sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
				i = i+1

		
		
		conn.Execute sqlStr
		

			  conn.CommitTrans
		
		%>
			<Script Language="JavaScript">
				try{
						window.parent.dialogArguments.opener.location.reload();
					}
				catch(e){window.self.close()}
			</Script>
		<%
		case "SAVECLIENT"
		Set conn = GetActiveConnection("KBroker")
		
		conn.BeginTrans

			MobileNo = trim(request.form("txtMobileNo"))
			if MobileNo = "" OR isnumeric(MobileNo) = false then MobileNo = "NULL"

			sqlClient = "Update Client set ClientCellTel = " & MobileNo & " Where Client_DPA_ = " & ID
			
			conn.execute (sqlClient)

		conn.CommitTrans
		
		%>
			<Script Language="JavaScript">
				try{
						window.parent.dialogArguments.opener.location.reload();
					}
				catch(e){window.self.close()}
			</Script>
		<%

		case "ADD"
			Set conn = GetActiveConnection("KBroker")
			'Insert a new record here
			user = Session("UserID")
			
			security= Request.Form("cbosecurity")
			
			 'validate security
			 If Trim(security) = "" Then%>
					 <script language = 'vbscript'>
							ShowMessage "Please specify the Security"
							
					 </script>
					 <% response.end
			 End If

			set guid = server.createobject("NDUtils.CGUID")
			 guidStr = guid.GenerateGUID
			 sqlStr = "INSERT INTO [StockWatch] (Client_DPA_,Security_DPA_,StockWatch_DPA_,StockWatch_EIT_,ChangedBy) SELECT " & "#" & ID & "#" & " as client_DPA_," & " " & security & " " & " as Security_DPA_," & _
						         "iif(isnull(max([StockWatch_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'StockWatch'),max([StockWatch_DPA_]) + 1)" & " " & " as StockWatch_DPA_" & _
						         "," & "'" & guidStr & "'" & " as StockWatch_EIT_" & _
						         "," & " " & UserId & " " & " as ChangedBy FROM [StockWatch]"	

			

			conn.BeginTrans
				conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
				
			conn.CommitTrans
		%>
			<Script Language="JavaScript">
				try{
						window.parent.dialogArguments.opener.location.reload();
						//window.self.close();
					}
				catch(e){window.self.close()}
			</Script>
		
			
		<%
		case "DELETE"
		
		'response.write request.form("txtDelete")
		ItemID=request.form("txtDelete")
		'response.end
		'find out whether any child records exist
		Set conn = GetActiveConnection("KBroker")

		
			'delete from database
			sqlStr = "Update [StockWatch] Set Deleted=1 WHERE StockWatch_DPA_ = " & ItemID
			conn.Execute SQLServerFormat(HandleQuote(sqlStr))
			conn.Close
			Set conn = Nothing
			%>
			<Script Language="JavaScript">
				try{
					window.parent.dialogArguments.opener.location.reload();
					}
				catch(e){window.self.close()}
			</Script>
				<%
				'respons
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
<title>Edit Stock Watch</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
 <SCRIPT language=Javascript src="../scripts/accountList.js"></SCRIPT>

<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frm<%=DataSource%>", "txtDate","cmdDate","<%=FormatDate(Date)%>",1);
	var calJournalDate=new ctlSpiffyCalendarBox("calJournalDate", "frm<%=DataSource%>", "txtJournalDate","cmdJournalDate","<%=FormatDate(Date)%>",1);
</SCRIPT>
<SCRIPT LANGUAGE="JavaScript1.2" src="../scripts/valjavavalidate.js" TYPE="text/javascript"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" >
<!--
	var isTrue=false;
function validate(){
	
	}
//-->
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
		function DeleteSecItem(ctrlName)		
	{
		var isDelete;
			if(ctrlName.checked==true){
			if(parseInt(document.all.frm<%=Datasource%>.totItems.value) >1){
			isDelete = confirm('Are you sure you want to delete this Stockwatch Item?');
				if (isDelete ==true){
				 document.all.frm<%=Datasource%>.submit();
				}
				else{
					ctrlName.checked=false;
				}
			}
			else{
				alert('The Stockwatch must have atleast one Item');
				ctrlName.checked=false;
			}
			}
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
	function UpdatePrice1(){
	
	var security = document.frm<%=DataSource%>.elements("cboSecurity1");
	
	document.frm<%=DataSource%>.elements("txtPrice1").value = security[security.selectedIndex].price;
	
			
	}
</script>
</head>

<body Class="Dialog">

<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>


<form name = 'frm<%=DataSource%>' id='frmMain' method = 'post' action = 'EditStockWatchItem.asp' onsubmit="return commit();">
<%
        Set conn = GetActiveConnection("KBroker")
             
        sqlStr = "SELECT * FROM " & Datasource & " WHERE Client_DPA_ = " & ID & " order by StockWatch_DPA_"
               
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected <%=DataEntity%> cannot be retrieved for editing"
                		
                </script>
                <% response.end
        End If
		
%>
<table border="0" width="100%">
  <tr>
    <td width="50">Client </td>
    <td width="416" align="left"><input type="text" name="" value="<%=rs("clientName")%>" readonly class="readonly" style="width:300;"></td>
  </tr>
  <tr>
    <td width="50">Mobile No </td>
    <td width="416" align="left"><input type="text" name="txtMobileNo" id="txtMobileNo" value="<%=rs("ClientCellTel")%>"  size="20">
	<input type = 'submit' Class=Buttons name ='cmdSaveClient' id = 'cmdSaveClient' value=" Save " onclick="javascript:document.all.frm<%=Datasource%>.action.value='SAVECLIENT';">
	</td>
  </tr>
  <tr>
    <td width="30">&nbsp;</td>
    <td width="416">&nbsp;</td>
  </tr>
    
  <tr>
  </table>
 <table border="0" width="90%" cellspacing="0" cellpadding="1">
 <tr><td>
   <table border="0" width="90%" cellspacing="0" cellpadding="1">
    <tr>
	  <td width="25%"><b><font color="#000080">No.</font></b></td>	
	  <td width="25%"><b><font color="#000080">Security</font></b></td>	
	  <td align="right"><b><font color="#000080">&nbsp;Price</td>
	  <td align="right"><b><font color="#000080">&nbsp;</td>
	  <td><b><font color="#000080">Delete</td>
    </tr>
	<%  
	 Set conn = GetActiveConnection("KBroker")

	set rsEdit= server.createobject("Adodb.recordset")
	sqlStr = "SELECT * FROM " & Datasource & " WHERE Client_DPA_ = " & ID & " order by StockWatch_DPA_"
    Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))

	itemID=trim(request.Querystring("itemID"))
	if itemID="" then
				
		dim cnt
		cnt =0
		' Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		 if not rsEdit.eof or not rsEdit.bof then 
			rsEdit.moveFirst

		 dim bgcolor
		 bgcolor="#FFFFFF"
		
		 do until rsEdit.eof 
		  if bgcolor="#F0EFDB" then  bgcolor="#FFFFFF" else bgcolor ="#F0EFDB"
		 ItemIDNo =rsEdit("StockWatch_DPA_")
		 if isnull(ItemIDNo) or trim(ItemIDNo)="" then ItemIDNo=0 
		 %>
		<tr bgcolor="<%=bgcolor%>" onMouseover="JavaScript: this.bgColor='#99CCFF'" onMouseout="JavaScript: this.bgColor='<%=bgcolor%>'" >
		 <td ><b><a href="EditStockWatchItem.asp?ID=<%=ID%>&itemID=<%=rsEdit("StockWatch_DPA_")%>"><font color="#000080"><%=rsEdit("StockWatch_DPA_")%><input type="hidden" Name="StockNo<%=cnt%>" value="<%=rsEdit("StockWatch_DPA_")%>"></font></b></td></a>

		<td><%=rsEdit.Fields("SecurityCode")%></td>
		<td align="right"><%=rsEdit.Fields("Price")%></td>
		
		 <td align="right"><b><font color="#000080">&nbsp;&nbsp;&nbsp;</td>
		  <td><input type="hidden" name="txtDel<%=cnt%>" value ="<%=rsEdit("StockWatch_DPA_")%>"><input type = 'CheckBox' name ='Del<%=cnt%>' id = 'Del' size="20" value="<%=rsEdit("StockWatch_DPA_")%>" onclick="javascript:document.all.frm<%=Datasource%>.action.value='DELETE';document.all.frm<%=Datasource%>.txtDelete.value=document.all.frm<%=Datasource%>.txtDel<%=cnt%>.value;DeleteSecItem(document.all.frm<%=DataSource%>.Del<%=cnt%>);"></td>
		</tr>
		<%
			cnt = cnt+1
			rsEdit.movenext
			loop
		end if

	else
	
		cnt =0
		sqlStr = "SELECT * FROM " & DataEntity & "FullList WHERE " & DataEntity & "_DPA_ = " & ID & " and journalEntry_DPA_= "& itemID &" order by journalEntry_DPA_"

		sqlStr = "SELECT * FROM " & Datasource & " WHERE StockWatch_DPA_ = " & itemID & " order by StockWatch_DPA_"
		 Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		 if not rsEdit.eof or not rsEdit.bof then 
			rsEdit.moveFirst

		
		 bgcolor="#FFFFFF"
	 do until rsEdit.eof 

	 ItemIDNo =rsEdit("StockWatch_DPA_")
	 if isnull(ItemIDNo) or trim(ItemIDNo)="" then ItemIDNo=0 
	 %>
    <tr>
	   <td width="40"><b><font color="#000080"><%=rsEdit("StockWatch_DPA_")%><input type="hidden" Name="StockNo" value="<%=rsEdit("StockWatch_DPA_")%>"></font></b></td>

       <td width="50%"><select name="cboSecurity1" onchange="UpdatePrice1();">

					<option value="" price="">Select a security</option>

					<%
						set rs=server.createobject("Adodb.recordset")
						Set Conn = GetActiveConnection("KBroker")

						'rs.open "Select * from datastream_SecurityPriceList order by securityCode ",conn, 0,1
						rs.open "Select * from StockWatchDataStreamPriceList order by securityCode ",conn, 0,1
						

						while not rs.eof 
					%>
						<option value='<%=rs("Security_DPA_")%>' price='<%=formatnum(rs("Price"))%>' <%if rsEdit("security_DPA_")=rs("Security_DPA_") then response.write "Selected"%>><%=rs("SecurityCode")%></option>
					<%
						rs.movenext
						wend
						rs.close
						set rs= nothing
						set conn= nothing
					%>
				</Select>
	</td>
     <td><input type="text" name="txtPrice1" style="text-align:right" value='<%=formatnum(rsEdit("Price"))%>' readonly class="readonly"></td>
	   <td>&nbsp;</td>
    	  <td align="right"><b><font color="#000080">&nbsp;&nbsp;&nbsp;</td>
	  <td><input type = 'CheckBox' name ='Del<%=cnt%>' id = 'Del' size="20" disabled></td>
    </tr>
	<%
		cnt = cnt+1
		rsEdit.movenext
		loop
	end if	
	
	end if
	%>
	
		</table>
	</td><tr>
	<tr><td>
	<div style="display:none" ID="AddLine">
	<table border="0" width="98%" cellspacing="0" cellpadding="1">
	<tr>
       <td width="50%"><select name="cboSecurity" onchange="UpdatePrice();">

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
     <td><input type="text" name="txtPrice" style="text-align:right" value="0" readonly class="readonly"></td>
	   <td>&nbsp;</td>
    </tr>
	<tr> <td nowrap colspan="9" align="center"><input type ="submit" name="SaveLine" value ="SAVE" onclick="javascript:document.all.frm<%=Datasource%>.action.value='ADD';">
	  <INPUT TYPE="Button" name="Cancel" value ="Cancel" onclick="javascript:document.all.AddLine.style.display='none';javascript:document.all.frm<%=Datasource%>.cmdAdd.style.display='';javascript:document.all.frm<%=Datasource%>.cmdClose.style.display='';">
	   <INPUT TYPE="Button" name="Close1" value ="Close" onclick="javascript:window.close();">
	  </td>
	  </tr>
	

  </table>
  </div>
</td></tr>
<tr><td>
<table>
	  <td width="20%"  align="center" >
		
     &nbsp;&nbsp;
		 <%if itemID="" then%><input type = 'button' Class=Buttons name ='cmdAdd' id = 'cmdAddMore' value=" Add Line/s " onclick="javascript:document.all.AddLine.style.display='block';javascript:document.all.frm<%=Datasource%>.cmdAdd.style.display='none';javascript:document.all.frm<%=Datasource%>.cmdClose.style.display='none';">
			<%end if%>
		 &nbsp;&nbsp; 
		
              <%if itemID<>"" then%>
        <input type = 'submit' Class=Buttons name ='cmdSave' id = 'cmdSave' value=" Save" onclick="javascript:document.all.frm<%=Datasource%>.action.value='SAVE';">
		<%end if%>
        &nbsp;&nbsp; 
		<%if itemID="" then%>
		<input type = 'button' Class=Buttons name ='cmdClose' id = 'cmdClose' value=" Close " onclick="javascript:window.close();">
		<%else%>
			<input type = 'submit' Class=Buttons name ='cmdCancel1' id = 'cmdCancel1' value=" Cancel ">
		<%end if%>
		<input type ="hidden" name="totItems" value="<%=cnt%>">
		<input type = 'hidden' name ='action' id = 'action' value="">
		<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
		<input type ="hidden" name="txtDelete" Id="txtDelete" value ="">
				
	</td>
  </tr>
</table>
</td></tr>
</table>
</form></body>

</html>
<SCRIPT LANGUAGE="JavaScript" >
<!--
	//check when the commit button is clicked
	function commit(){
		if (document.frm<%=DataSource%>.action.value=='COMMIT'){
			if (document.frm<%=DataSource%>.Total.value!=0){
				alert('You cannot commit this Journal \n The Total should be Zero');
				return false;
			}
		}
		
	}
	//validate the detail Items
function validateItems(){
	 var doc= document.frm<%=DataSource%>;
	 <%
		z=0
		do while z<cnt
			z=z+1
		loop
	%>
	}
	
		<%z=0
		 do while z<cnt	
		%>
		function FetchAccounts<%=z%>(theList)
		{
			var i = 0;
			var entity = theList.value;
			var toList = document.frmMain.cboAccount<%=z%>;
			
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
//-->
		<%
			z=z+1
			loop
		%>
</SCRIPT>
