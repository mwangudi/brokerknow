<!--#include file="../libroutines.asp"-->
<%

'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "AddJournal"
		const ActionPage = "JournalList"
		const DataEntity = "Journal"
		const DataEntityPlural = "Journals"
		const ActionFolder = "Operations"
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
						Dim user
						Dim jDate
						Dim narrative
						Dim Entity
						Dim Account
						Dim debit	
						Dim credit
					 
						user = Session("UserID")
						jDate = Request.Form("txtJournalDate") & " " & Time
						narrative = Request.Form("txtNarrative")
						Entity = Request.Form("cboEntity")
						Account = Request.Form("cboAccount")
						debit = Request.Form("txtDebit")
						credit = Request.Form("txtCredit")						

		
						 'validate Entity
						 If Trim(Entity) = "" Then%>
						         <script language = 'vbscript'>
						         		ShowMessage "Please specify the Entity"						         		
						         </script>
						         <% response.end
						 End If
							
						'validate Narrative
						 If Trim(narrative) = "" or len(Trim(narrative)) < 1 Then%>
						         <script language = 'vbscript'>
						         		ShowMessage "Please specify the Narrative"						         		
						         </script>
						         <% response.end
						 End If

						 'validate Account
						 If Trim(Account) = "" Then%>
						         <script language = 'vbscript'>
						         		ShowMessage "Please specify the Account"
						         		
						         </script>
						         <% response.end
						 End If
						 'validate size of Narrative
						 If Len(Narrative) > 500 Then%>
						         <script language = 'vbscript'>
						         ShowMessage "Narrative can only be 500 characters in length"
						         
						         </script>
						         <% response.end
						 End If
						         							
						'ensure Debit is numeric
						If (Debit <> "") And (Not IsNumeric(Debit)) Then%>
							<script language = 'vbscript'>
								ShowMessage "Debit must be numeric"
											
							</script>
							<% response.end
						End If
						'ensure Credit is numeric
						If (Credit <> "") And (Not IsNumeric(Credit)) Then%>
							<script language = 'vbscript'>
								ShowMessage "Credit must be numeric"
											
							</script>
							<% response.end
						End If
						 
						 if (trim(debit) = "") and (trim(credit) = "") then%>
								<script language = 'vbscript'>
									ShowMessage "You must enter either a debit amount or a credit amount"
												
								</script>
								<% response.end
						End If
						  
						 debit = iif(trim(debit) = "",0,debit)
						 credit = iif(trim(credit) = "",0,credit)
						   
						 'save header				 
						 set guid = server.createobject("NDUtils.CGUID")
						 guidStr = guid.GenerateGUID
						 
						 sqlStr = "INSERT INTO [Journal] (JournalDate,UserID,JournalNarrative,Journal_DPA_,Journal_EIT_,ChangedBy) SELECT " & "#" & FormatDate(jDate) & "#" & " as JournalDate," & " " & UserID & " " & " as UserID" & _
						         "," & "'" & narrative & "'" & " as JournalNarrative," & " " & "iif(isnull(max([Journal_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Journal'),max([Journal_DPA_]) + 1)" & " " & " as Journal_DPA_" & _
						         "," & "'" & guidStr & "'" & " as Journal_EIT_" & _
						         "," & " " & UserId & " " & " as ChangedBy FROM [Journal]"

						 response.write sqlStr & vbcrlf & vbcrlf ': response.end
						 Set conn = GetActiveConnection("KBroker")

						 conn.BeginTrans
								sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
								conn.Execute sqlStr
						     
								'obtain header key value
								sqlStr = "SELECT Journal_DPA_] FROM [Journal] WHERE [Journal_EIT_] = " & "'" & guidStr & "'"
						     
								Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
								If (rs.EOF Or rs.BOF) Then%>
						         			<script language = 'vbscript'>
						         					ShowMessage "A serious error has been encountered while saving the data. Try saving again"
						         					
						         			</script>
						         			<% response.end
								End If
						     
								'save detail data
								sqlStr = "INSERT INTO [JournalEntry] (JournalEntryDebit,JournalEntryCredit" & _
										",JournalEntry_DPA_,Journal_DPA_,EntityType_DPA_,Entity_DPA_) SELECT " & " " & debit & " " & " as JournalEntryDebit" & _
										"," & " " & credit & " " & " as JournalEntryCredit" & _
										"," & " " & "iif(isnull(max([JournalEntry_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'JournalEntry'),max([JournalEntry_DPA_]) + 1)" & " " & " as JournalEntry_DPA_" & _
										"," & " " & rs.Fields("Journal_DPA_") & " " & " as Journal_DPA_" & _
										"," & " " & Entity & " " & " as EntityType_DPA_" & _
										"," & " " & Account & " " & " as Entity_DPA_" & _
										" FROM [JournalEntry]"
						     
							 response.write sqlStr
								sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
								conn.Execute = sqlStr
						
						if(Cint(entity)=1) then
							conn.execute ("Exec ClientTotalProcedure " & account)							
							
							'conn.execute ("Exec ClientStatementProcBrief " & account)
							conn.execute ("Exec ClientBalanceProcedure " & account)	
						end if
							
						conn.CommitTrans
						
						 'retrieve the item ID
						'sqlStr = "SELECT JournalEntry_DPA_  FROM JournalEntry WHERE Journal_DPA_=" & rs.Fields("Journal_DPA_")
						'Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
						'If (rs.EOF Or rs.BOF) Then%>
						         	<script language = 'vbscript'>
						         			'ShowMessage "An error has been encountered while saving the order. Try editing the Order if you wish to add more entries"
						         			
						         	</script>
						         	<% 'response.end
						'End If
						
						'WriteDialogRelocateScript "Edit" & DataEntity & ".asp?ID=" & rs.Fields("JournalEntry_DPA_")
						WriteDialogRelocateScript "Edit" & DataEntity & ".asp?ID=" & rs.Fields("Journal_DPA_") & "&IDHeld=Journal"
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
</script>
</head>

<body Class="Dialog">

<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>

<form name = 'frm<%=DataSource%>' method = 'post' action = '<%=DataSource%>Test.asp' id = 'frmMain'>


<table border="0" width="510">
  <tr>
    <td width="80">Date</td>
    <td width="416"><SCRIPT language="JavaScript">calJournalDate.writeControl();</SCRIPT></td>
  </tr>
  <tr>
    <td width="10%">Narrative</td>
    <td width="50%"><input type = 'text' name ='txtNarrative' id = 'txtNarrative' size="20"></td>
  </tr>
  <tr>
  <td colspan = '2' width="502">
  
  <table border="0" width="700">
    <tr>
      <td width="5%"><b><font color="#000080">Entity</font></b></td>
      <td width="85%"><b><font color="#000080">Account</font></b></td>
      <td width="5%"><b><font color="#000080">Debit</font></b></td>
      <td width="5%"><b><font color="#000080">Credit</td>
    </tr>
    <tr>
      <td width="32%"><select name = 'cboEntity' id = 'cboEntity' size="1" onchange='FetchAccounts(this)'>
    	
<%
		Set conn = GetActiveConnection("KBroker")
		
		'delete earlier incomplete journal entries
		
		conn.BeginTrans
		
			sqlStr = "DELETE FROM JournalEntry WHERE Journal_DPA_ IN (SELECT Journal_DPA_ FROM Journal WHERE UserID = " & Session("UserID") & " AND JournalCommitted = 0)" 
			conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		
			sqlStr = "DELETE FROM Journal WHERE UserID = " & Session("UserID") & " AND JournalCommitted = 0"
			conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		
		conn.CommitTrans
		
		
        sqlStr = "SELECT * FROM [FullEntityTypeList] Order By EntityTypeName"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF
						if rs.Fields("EntityType_DPA_").value = currentEntityType then%>
								<option selected value = '<%=rs.Fields("EntityType_DPA_")%>'><%=rs.Fields("EntityTypeName")%></option>
                        <%else%>
								<option value = '<%=rs.Fields("EntityType_DPA_")%>'><%=rs.Fields("EntityTypeName")%></option>
                        <%end if
                        rs.MoveNext
                Loop
        End If
%>

    </select></td>
    <td width="50%">
    <input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="10" onBlur="txtval = this.value; selectItem(cboAccount);UpdateCode(true,cboAccount,txtClientCode);">
    <select name = 'cboAccount' id = "cboAccount" size="1" 
    onchange='UpdateCode(true,cboAccount,txtClientCode)'
	onKeypress="return (dodefaultaction()==''); " 
	onKeydown="return (dodefaultaction()==''); " 
	onKeyup="return (FilterData(this,<%=currentEntityType%>,UpdateCode(change(cboAccount,0),cboAccount,txtClientCode)));" 
	onfocus="txtval = '';inputIsItemCode = 1;" 
	onblur="txtval = '';inputIsItemCode = 1;" tabindex=-1>
	<option selected SearchCode = "" SearchText = ""  value = ''></option>

 
<%
        
        sqlStr = "SELECT * FROM [CompleteEntityList] WHERE EntityType_DPA_ =" & currentEntityType & " Order By EntityName"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF
                AccountName=Mid(rs.Fields("EntityName"),1,30)
                %>
                        <option SearchCode = "<%=rs.Fields("EntityCode")%>" SearchText = "<%=rs.Fields("EntityName")%>" value = '<%=rs.Fields("Entity_DPA_")%>'><%=AccountName%></option>
                        <%rs.MoveNext
                Loop
        End If
%>

    </select></td>
      <td width="12%"><input type = 'text' name ='txtDebit' id = 'txtDebit' size="9"></td>
      <td width="24%"><input type = 'text' name ='txtCredit' id = 'txtCredit' size="18"></td>
    </tr>
  </table>
  
  </td>
  </tr>
 
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