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

		Dim user
		Dim jDate
		Dim narrative
		Dim Entity
		Dim Account
		Dim debit	
		Dim credit
		totItems=(Request.Form("totItems"))
		user = Session("UserID")
		jDate = Request.Form("txtJournalDate") & " " & Time
		PDateComp = Request.Form("txtJournalDate")
		narrative = Request.Form("txtNarrative")
		Entrynarrative = Request.Form("txtEntryNarrative")
		Entity = Request.Form("cboEntity")
		Account = Request.Form("cboAccount")
		debit = Request.Form("txtDebit")
		credit = Request.Form("txtCredit")
		debit = iif(trim(debit) = "",0,debit)
		credit = iif(trim(credit) = "",0,credit)
		currentEntityType = 1
		'debit = iif(trim(debit) = "",0,debit)
		'credit = iif(trim(credit) = "",0,credit)

		'-----------------------------------------------
		'Get Clients Current Acount Balance
		if (Cint(Entity) = 1) or (Cint(Entity) = 2) then
							
		  Set myconn = GetActiveConnection("KBroker")
							
		  if (Cint(Entity) = 1) then
						mysql = " SELECT     ISNULL(ClientBalances.CurrentBal, 0) + ISNULL(Client.CreditLimit, 0) - ISNULL(ClientTotal.Total, 0) AS AvailableCredit, Client.Client_DPA_ "  & _
									" FROM         Client LEFT OUTER JOIN " & _
									" ClientTotal ON Client.Client_DPA_ = ClientTotal.Client_DPA_ LEFT OUTER JOIN " & _
									" ClientBalances ON Client.Client_DPA_ = ClientBalances.client_DPA_ " & _
									" WHERE     (Client.Deleted = 0) AND (Client.IsFrozen <> 1)  AND  Client.client_DPA_=" & Account
		  elseif (Cint(Entity) = 2) then
		  	'myconn.Execute("Execute AgentStatementProcNew " & Account)		
		  	mysql = "SELECT CurrentBal FROM AgentBalances WHERE Agent_DPA_=" & Account
		  end if
								
		  if ccur(debit) > 0 then	
								 	
		  	 	Set myrs = myconn.Execute(mysql)
								 
		  		 if not myrs.bof or myrs.eof then
				 
				 	if (Cint(Entity) = 1) then
						currentclibl = myrs.fields("AvailableCredit").value
					elseif (Cint(Entity) = 2) then
						currentclibl = myrs.fields("CurrentBal").value
					end if
					
		  	 			balanceafter = ccur(currentclibl) - ccur(debit) 
										
		  			if balanceafter < 0 then
		  			 Set myconn = nothing
		  		 	 Set myrs = nothing
		  			 %>
		  	         <script language = 'vbscript'>
		  	         		ShowMessage "There is insufficient funds in this account"
		  					window.history.go(-1)
		  	         </script>
		  	         <% 
		  			 Set myconn = nothing
		  		 	 Set myrs = nothing
		  			 response.end
		  			end if
									
		  		 end if
		  		 Set myconn = nothing
		  	 	 Set myrs = nothing
		  	end if	
								 
		  end if	 
		   '-----------------------------------------------		
		
		'Response.write(Entity)
			if(Cint(Entity)=100) then
				Entity=7
			end if

		UserID=Session("UserID")
		select case action
		case "SAVE"
		
		'Prevent Back-Dating
		If DateDiff("d",cdate(PDateComp),Date) > 0 Then%>
		         <script language = 'vbscript'>
		         		ShowMessage "The system does not allow back-dating of Journals."
		         </script>
		         <% response.end
		 End If
		
		Set conn = GetActiveConnection("KBroker")
				
		conn.BeginTrans
		'update the detail data
			Dim i
			Dim TotalAmount
			TotalAmount=0
			i=0
	
	mysqlstr = "UPDATE    ClientBalances " & _
							 " SET              CurrentBal = CurrentBal - CASE WHEN JournalEntry.JournalEntryDebit = 0 THEN JournalEntry.JournalEntryCredit ELSE 0 - JournalEntry.JournalEntryDebit " & _
							 "                        END " & _
							 " FROM         ClientBalances INNER JOIN " & _
							 "                       JournalEntry ON ClientBalances.client_DPA_ = JournalEntry.Entity_DPA_ INNER JOIN " & _
							 "                       Journal ON JournalEntry.Journal_DPA_ = Journal.Journal_DPA_ " & _
							 " WHERE     (JournalEntry.EntityType_DPA_ = 1) AND (JournalEntry.Deleted = 0) AND (Journal.JournalCommitted = 1) AND (Journal.Released = 1) AND  " & _
							 "                       (JournalEntry.Journal_DPA_ = " & ID & ")"

	conn.Execute(mysqlstr)

	while i< cint(totItems) 
	
	if trim(Request.Form("txtDebit"&i))="" then debit=0 else debit = ccur(trim(Request.Form("txtDebit"&i)))
	if trim(Request.Form("txtCredit"&i))="" then credit=0 else credit = ccur(trim(Request.Form("txtCredit"&i)))

	Entity=Request.Form("cboEntity"&i)

		'debit = iif(trim(Request.Form("txtDebit"&i)) = "",0,debit)
		'credit = iif(trim(Request.Form("txtCredit"&i)) = "",0,credit)

	if(Cint(Entity)=100) then
			Entity=7
	end if
	sqlStr = "UPDATE JournalEntry SET JournalEntryDebit = " & debit & "," & _
			" JournalEntryCredit = " & credit & ", EntityType_DPA_ = " & Entity & "," & _
			" Narrative = '" &  Request.Form("txtEntryNarrative"&i) & "'"  & _
			", Entity_DPA_ = " & Request.Form("cboAccount"&i) & _
			" WHERE JournalEntry_DPA_=" & Request.Form("JournalNo"& i)	
		sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
				i = i+1
				'response.write  sqlStr & "<br>"
				'response.end				
				conn.Execute sqlStr
			     wend				 
				
				'update the data to the table 	
				sqlStr = "UPDATE [Journal] SET Released=0, JournalDate = " & "#" & jDate & "#" & ",JournalNarrative = " & "'" & narrative & "'" & "" & _
				" WHERE Journal_DPA_  = " & ID       
				'sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
					
				'response.write  sqlStr & "<br>"
				
				conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
				sqlStr1="UPDATE Journal Set JournalCommitted =0,Released=0, ChangedBy=" & UserId & ",TimeChanged=GetDate() Where( Journal_DPA_=" & ID & ")"
				conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr1))

				'conn.execute ("ClientTotalsDelete")
				'conn.execute ("ClientTotalsProcedure")
				'conn.execute ("ClientBalancesDelete")
				'conn.execute ("ClientBalancesProcedure")
		
				conn.CommitTrans
				
				Conn.close
				Set Conn = nothing

		%>
			<Script Language="JavaScript">
				try{
						window.parent.dialogArguments.opener.location.reload();
						//window.self.close();
					}
				catch(e){window.self.close()}
			</Script>
		<%
		case "ADD"
			Set conn = GetActiveConnection("KBroker")
			'save detail data

			mysqlstr = "UPDATE    ClientBalances " & _
							 " SET              CurrentBal = CurrentBal - CASE WHEN JournalEntry.JournalEntryDebit = 0 THEN JournalEntry.JournalEntryCredit ELSE 0 - JournalEntry.JournalEntryDebit " & _
							 "                        END " & _
							 " FROM         ClientBalances INNER JOIN " & _
							 "                       JournalEntry ON ClientBalances.client_DPA_ = JournalEntry.Entity_DPA_ INNER JOIN " & _
							 "                       Journal ON JournalEntry.Journal_DPA_ = Journal.Journal_DPA_ " & _
							 " WHERE     (JournalEntry.EntityType_DPA_ = 1) AND (JournalEntry.Deleted = 0) AND (Journal.JournalCommitted = 1) AND (Journal.Released = 1) AND  " & _
							 "                       (JournalEntry.Journal_DPA_ = " & ID & ")"

	'conn.Execute(mysqlstr)

					sqlStr = "INSERT INTO [JournalEntry] (JournalEntryDebit,JournalEntryCredit" & _
							",JournalEntry_DPA_,Journal_DPA_,EntityType_DPA_,Entity_DPA_,Narrative) SELECT " & " " & debit & " " & " as JournalEntryDebit" & _
							"," & " " & credit & " " & " as JournalEntryCredit" & _
							"," & " " & "iif(isnull(max([JournalEntry_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'JournalEntry'),max([JournalEntry_DPA_]) + 1)" & " " & " as JournalEntry_DPA_" & _
							"," & " " & ID & " " & " as Journal_DPA_" & _
							"," & " " & Entity & " " & " as EntityType_DPA_" & _
							"," & " " & Account & " " & " as Entity_DPA_" & _
							"," & "'" & Entrynarrative & "'" & " as Narrative" & _
							" FROM [JournalEntry]"
				
				sqlStr1="UPDATE Journal Set JournalCommitted =0, Released=0, ChangedBy=" & UserId & ",TimeChanged=GetDate() Where( Journal_DPA_=" & ID & ")"
				conn.BeginTrans
				conn.Execute mysqlstr
				conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
				conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr1))

				'conn.execute ("ClientTotalsDelete")
				'conn.execute ("ClientTotalsProcedure")
				'conn.execute ("ClientBalancesDelete")
				'conn.execute ("ClientBalancesProcedure")
		
				conn.CommitTrans

				Conn.close
				Set Conn = nothing

		%>
			<Script Language="JavaScript">
				try{
						window.parent.dialogArguments.opener.location.reload();
						//window.self.close();
					}
				catch(e){window.self.close()}
			</Script>
		<%	
		case "COMMIT"

			Set conn = GetActiveConnection("KBroker")
			set rs = server.createobject("Adodb.recordset")	'update the data to the table 	
				sqlStr = "UPDATE [Journal] SET Released=0, JournalDate = " & "#" & jDate & "#" & ",JournalNarrative = " & "'" & narrative & "'" & "" & _
				" WHERE Journal_DPA_  = " & ID       
				sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
					
				'response.write  sqlStr & "<br>"
				conn.Execute = sqlStr
				sqlStr1="UPDATE Journal Set JournalCommitted =1,Released=0, ChangedBy=" & UserId & ",TimeChanged=GetDate() Where( Journal_DPA_=" & ID & ")"
				conn.Execute = sqlStr1
				
				mysqlstr = "UPDATE    ClientBalances " & _
							 " SET              CurrentBal = CurrentBal + CASE WHEN JournalEntry.JournalEntryDebit = 0 THEN JournalEntry.JournalEntryCredit ELSE 0 - JournalEntry.JournalEntryDebit " & _
							 "                        END " & _
							 " FROM         ClientBalances INNER JOIN " & _
							 "                       JournalEntry ON ClientBalances.client_DPA_ = JournalEntry.Entity_DPA_ INNER JOIN " & _
							 "                       Journal ON JournalEntry.Journal_DPA_ = Journal.Journal_DPA_ " & _
							 " WHERE     (JournalEntry.EntityType_DPA_ = 1) AND (JournalEntry.Deleted = 0) AND (Journal.JournalCommitted = 1) AND (Journal.Released = 1) AND  " & _
							 "                       (JournalEntry.Journal_DPA_ = " & ID & ")"

				'conn.Execute(mysqlstr)
				conn.execute("UpdateClientBalanceFromJournal " & ID)
				'conn.execute ("ClientTotalsDelete")
				'conn.execute ("ClientTotalsProcedure")
				'conn.execute ("ClientBalancesDelete")
				'conn.execute ("ClientBalancesProcedure")

				Conn.close
				Set Conn = nothing

		
		%>
			<Script Language="JavaScript">
				try{
					window.parent.dialogArguments.opener.location.reload();
					window.self.close();

					}
				catch(e){window.self.close()}
			</Script>
		<%
		case "DELETE"
		sqlStr ="Select "


		'response.write request.form("txtDelete")
		ItemID=request.form("txtDelete")
		'response.end
		'find out whether any child records exist
		Set conn = GetActiveConnection("KBroker")
		set rs = server.createobject("Adodb.recordset")
			sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = 'JournalEntry') AND (ChildType = " & LinkedIndependent & ")"
			
			Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			If Not (rs.BOF Or rs.EOF) Then
					Dim childRS
					Dim tableName
	                
					rs.MoveFirst
					Do Until rs.EOF
                				tableName = rs.Fields("Child")
							sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE JournalEntry_DPA_ = " & ItemID & " and Deleted=0"
							Set childRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
							If Not (childRS.BOF Or childRS.EOF) Then%>
                					<script language = 'vbscript'>
                						ShowMessage "<%=rs.Fields("DeletionMessage")%>"
                						
                					</script>
                					<%response.end
							End If
							rs.MoveNext
					Loop
			End If
						
			
			'Mark the records as deleted in the database

			mysqlstr = "UPDATE    ClientBalances " & _
							 " SET              CurrentBal = CurrentBal - CASE WHEN JournalEntry.JournalEntryDebit = 0 THEN JournalEntry.JournalEntryCredit ELSE 0 - JournalEntry.JournalEntryDebit " & _
							 "                        END " & _
							 " FROM         ClientBalances INNER JOIN " & _
							 "                       JournalEntry ON ClientBalances.client_DPA_ = JournalEntry.Entity_DPA_ INNER JOIN " & _
							 "                       Journal ON JournalEntry.Journal_DPA_ = Journal.Journal_DPA_ " & _
							 " WHERE     (JournalEntry.EntityType_DPA_ = 1) AND (JournalEntry.Deleted = 0) AND (Journal.JournalCommitted = 1) AND (Journal.Released = 1) AND  " & _
							 "                       (JournalEntry.Journal_DPA_ = " & ID & ")"

			conn.Execute(mysqlstr)

			sqlStr = "Update [JournalEntry] Set Deleted=1 WHERE JournalEntry_DPA_ = " & ItemID
			conn.Execute SQLServerFormat(HandleQuote(sqlStr))

			sqlStr1="UPDATE Journal Set JournalCommitted =0,Released=0, ChangedBy=" & UserId & ",TimeChanged=GetDate() Where( Journal_DPA_=" & ID & ")"
			
			'conn.execute ("ClientTotalsDelete")
			'conn.execute ("ClientTotalsProcedure")
			'conn.execute ("ClientBalancesDelete")
			'conn.execute ("ClientBalancesProcedure")
		
			conn.Execute = sqlStr1
			
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
<SCRIPT LANGUAGE="JavaScript1.2" src="../scripts/valjavavalidate.js" TYPE="text/javascript"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" >
<!--
	var isTrue=false;
function validate(){
	 var doc= document.frmMain;
	 //if (checkcomboempty(document.frmMain.elements("cboEntity"),' Please Select an Entity')==false) return false;
	 //if (checkcomboempty(document.frmMain.elements("cboAccount"),' Please Select an Account')==false) return false;
	 //if ((trim(doc.elements("txtDebit").value)=='') && (trim(doc.elements("txtCredit").value)=='')){
		//alert("Please enter the Debit or Credit Amount");
		//doc.elements("txtDebit").focus();
		//return false;
	 //}
 	 //if (checkempty(doc.txtDebit,' Please Enter Debit ')==false) return false;
	 //if (checkempty(doc.txtCredit,' Please Enter Credit ')==false) return false;
	 //if (checknumericWithCommas(doc.txtDebit," Debit must be numeric")==false) return false;
	 //if (checknumericWithCommas(doc.txtCredit," Credit must be numeric")==false) return false;
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
		function DeleteJournalItem(ctrlName)
	{
		var isDelete;
			if(ctrlName.checked==true){
			if(parseInt(document.all.frm<%=Datasource%>.totItems.value) >1){
			isDelete = confirm('Are you sure you want to delete this Journal Item?');
				if (isDelete ==true){
				 document.all.frm<%=Datasource%>.submit();
				}
				else{
					ctrlName.checked=false;
				}
			}
			else{
				alert('The Journal must have atleast one Item');
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

		function ClearFields(element)
		{
		   if (element == 'txtClientCode')
		   {
			document.frmAddJournal.elements("txtClientCode").value = ''
			return;
		   }
		}

		function LoadClient(accountno, element, guidstr)
		{
		 var clientcode = document.frmAddJournal.elements("txtClientCode").value
		  var clientcobo = document.getElementById("cboAccount");		 
		 
	     var x_clientname;
			
		 if (element == 'txtClientCode')
		 {
			clientcds = ''
			clientname = ''
			
			if (clientcode == '')
			{
			document.frmAddJournal.elements("txtClientCode").value = 'Code'
			clientcobo.length = 1;
			clientcobo[0].text = 'Load Account';
			clientcobo[0].value = '';
			return;
			}
			
		 }

				xmlhttp = createXMLHTTPObj();
				
				url="GetList.asp?clientcode="+clientcode+"&cdsno="+clientcds+"&clientname="+clientname+"&action=SLoadClient&guidstr="+guidstr;
				
				//alert(url);

				xmlhttp.open("GET",url,true);

				xmlhttp.onreadystatechange=function() 
				  {
							if (xmlhttp.readyState==4) 
							{
								returnStr = xmlhttp.responseText;
								returnStr = getBodyHTML(returnStr);
								
								//alert(returnStr);

								myArray = returnStr.split("<->");
								
								x_clientname = myArray[7]

								if (x_clientname.length > 25) 
								{
									x_clientname = x_clientname.substring(0,23)  + '...';
								}
								
								//document.getElementById("cboClient").options.length = 0;
								clientcobo.length = 1;
								if (element != 'txtclientname')
								{
									document.frmAddJournal.elements("txtClientCode").value = myArray[5];
									clientcobo[0].text = x_clientname;
									clientcobo[0].value = myArray[5];
								}
							}
					}
				 
				 xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
				 xmlhttp.send();
		}

</script>
</head>

<body Class="Dialog">

<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>


<form name = 'frm<%=DataSource%>' id='frmMain' method = 'post' action = 'EditJournalItem.asp' onsubmit="return commit();">
<%
        Set conn = GetActiveConnection("KBroker")
             
        sqlStr = "SELECT JournalDate,Released,Journal_DPA_,JournalNarrative FROM " & DataEntity & "FullList WHERE " & DataEntity & "_DPA_ = " & ID
        
		
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))

		'response.write(sqlStr)
		'Response.end

        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected <%=DataEntity%> cannot be retrieved for editing"
                		
                </script>
                <% response.end
        End If
%>
<table border="0" width="100%">
  <tr>
    <td width="80"><%=DataEntity%> No.</td>
    <td width="416">&nbsp;<input readonly = 'true' class=readonly  type = 'text' name ='txt<%=DataEntity%>No' id = 'txt<%=DataEntity%>No' value = '<%=ID%>' size="20"></td>
  </tr>
  <tr>
    <td width="80">Date</td>
    <SCRIPT language="JavaScript">
	var calJournalDate=new ctlSpiffyCalendarBox("calJournalDate", "frm<%=DataSource%>", "txtJournalDate","cmdJournalDate","<%= FormatDate(rs.Fields("JournalDate")) %>",1);
	</SCRIPT>
    <td width="416">&nbsp;<SCRIPT language="JavaScript">calJournalDate.writeControl();</SCRIPT></td>
  </tr>
  <tr>
    <td width="80">Released</td>
    <%If rs.Fields("Released") = 1 Then%>
    <td width="416"><input type=checkbox class='BorderLess' style="border: 0;" checked disabled value='False' name='chkRelease' onClick='UpdateReleaseStatus(this,"<%=rs.Fields("Journal_DPA_")%>");'></td>
    <%Else%>
    <td width="416"><input type=checkbox class='BorderLess' disabled style="border: 0;"></td>
    <%End If%>
  </tr>
  <tr>
    <td width="80">Narrative</td>
    <td width="416">&nbsp;<input type = 'text' name ='txtNarrative' id = 'txtNarrative' size="20" value = '<%=rs.Fields("JournalNarrative")%>'></td>
  </tr>
  <%
	set rsTotal= server.createobject("Adodb.recordset")

	mysqlstr = "SELECT     SUM(JournalEntry.JournalEntryCredit) - SUM(JournalEntry.JournalEntryDebit) AS Total " & _
				 " FROM         JournalEntry INNER JOIN " & _
				 "                       Journal ON JournalEntry.Journal_DPA_ = Journal.Journal_DPA_ " & _
				 " WHERE     (Journal.Deleted = 0) AND (JournalEntry.Deleted = 0) AND (Journal.Journal_DPA_ = " & ID & ")"

	
		
	rsTotal.open mysqlstr, conn,0,1
	
	'"Select sum(JournalEntryCredit)-sum(JournalEntryDebit) as Total from JournalFulllist where Journal_DPA_=" & ID & " And Deleted =0", conn, 0,1

	if not rsTotal.eof or not rsTotal.bof then
		Total = CCur(rsTotal("Total"))
		if Total<0 then Total1=(0-Total) & " Dr" else Total1= Total & " Cr"
	end if
  %>
   <tr>
    <td width="10%">Total</td>
    <td width="50%">
	<input type ="hidden" name="Total" id ="Total" value ="<%=Total%>">
	<input type = 'text' name ='txtTotal' id = 'txtTotal' size="20" class="readonly" readonly value ="<%=Total1%>"></td>
  </tr>
  <tr>
  </table>
 <table border="0" width="98%" cellspacing="0" cellpadding="1">
 <tr><td>
   <table border="0" width="98%" cellspacing="0" cellpadding="1">
    <tr>
	  <td width="40"><b><font color="#000080">No.</font></b></td>	
      <td><b><font color="#000080">Entity</font></b></td>
	  <td><b><font color="#000080">&nbsp;&nbsp;</font></b></td>
      <td><b><font color="#000080">Account</font></b></td>
	  <td><b><font color="#000080">Narrative</font></b></td>
      <td align="right"><b><font color="#000080">Debit</font></b></td>
      <td align="right"><b><font color="#000080">Credit</td>
	  <td align="right"><b><font color="#000080">&nbsp;</td>
	  <td><b><font color="#000080">Delete</td>
    </tr>
	<%  
	 Set conn = GetActiveConnection("KBroker")
	set rsEdit= server.createobject("Adodb.recordset")
	itemID=trim(request.Querystring("itemID"))
	if itemID="" then

		sqlStr = "SELECT * FROM " & DataEntity & "FullList WHERE " & DataEntity & "_DPA_ = " & ID & " order by journalEntry_DPA_"
		
		'response.write(mysqlstr)
		'Response.end
		dim cnt
		cnt =0
		 Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		 if not rsEdit.eof or not rsEdit.bof then 
			rsEdit.moveFirst

		 dim bgcolor
		 bgcolor="#FFFFFF"
		
		 do until rsEdit.eof 
		  if bgcolor="#F0EFDB" then  bgcolor="#FFFFFF" else bgcolor ="#F0EFDB"
		 ItemIDNo =rsEdit("Journal_DPA_")
		 if isnull(ItemIDNo) or trim(ItemIDNo)="" then ItemIDNo=0 
		 %>
		<tr bgcolor="<%=bgcolor%>" onMouseover="JavaScript: this.bgColor='#99CCFF'" onMouseout="JavaScript: this.bgColor='<%=bgcolor%>'" >
		 <td ><b><a href="EditJournalItem.asp?ID=<%=ID%>&itemID=<%=rsEdit("JournalEntry_DPA_")%>"><font color="#000080"><%=rsEdit("JournalEntry_DPA_")%><input type="hidden" Name="JournalNo<%=cnt%>" value="<%=rsEdit("JournalEntry_DPA_")%>"></font></b></td></a>

		<td><%=rsEdit.Fields("JournalEntryEntity")%></td>
		<td><%=rsEdit.Fields("Entity_DPA_")%></td>
		<td nowrap><%=mid(rsEdit.Fields("JournalEntryAccount"),1,25)%></td>
		<td nowrap align="left"><%=rsEdit("JournalEntryNarrative")%></td>
		 <td nowrap align="right"><%=FormatNum(rsEdit("JournalEntryDebit"))%></td>
		 <td nowrap align="right"><%=FormatNum(rsEdit("JournalEntryCredit"))%></td>
		 <td align="right"><b><font color="#000080">&nbsp;&nbsp;&nbsp;</td>
		  <td><input type="hidden" name="txtDel<%=cnt%>" value ="<%=rsEdit("JournalEntry_DPA_")%>"><input type = 'CheckBox' name ='Del<%=cnt%>' id = 'Del' size="20" value="<%=rsEdit("JournalEntry_DPA_")%>" onclick="javascript:document.all.frm<%=Datasource%>.action.value='DELETE';document.all.frm<%=Datasource%>.txtDelete.value=document.all.frm<%=Datasource%>.txtDel<%=cnt%>.value;DeleteJournalItem(document.all.frm<%=DataSource%>.Del<%=cnt%>);"></td>
		</tr>
		<%
			cnt = cnt+1
			rsEdit.movenext
			loop
		end if

	else
	
		cnt =0
		sqlStr = "SELECT * FROM " & DataEntity & "FullList WHERE " & DataEntity & "_DPA_ = " & ID & " and journalEntry_DPA_= "& itemID &" order by journalEntry_DPA_"
		 Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		 if not rsEdit.eof or not rsEdit.bof then 
			rsEdit.moveFirst

		
		 bgcolor="#FFFFFF"
	 do until rsEdit.eof 

	 ItemIDNo =rsEdit("Journal_DPA_")
	 if isnull(ItemIDNo) or trim(ItemIDNo)="" then ItemIDNo=0 
	 %>
    <tr>
	   <td width="40"><b><font color="#000080"><%=rsEdit("JournalEntry_DPA_")%><input type="hidden" Name="JournalNo<%=cnt%>" value="<%=rsEdit("JournalEntry_DPA_")%>"></font></b></td>
	<%
		Set conn = GetActiveConnection("KBroker")
		currentEntityType =rsEdit("EntityType_DPA_")
		if cint(currentEntityType) = 7 then
			set rs1= server.createobject("Adodb.recordset")
			rs1.open "select AccountManager from owner where Owner_DPA_=" & rsEdit("Entity_DPA_"), conn,0,1
			if not rs1.eof or not rs1.bof then
				isAccManager=rs1("AccountManager")
			end if
			rs1.close
			set rs1 = nothing
		end if


	%>
      <td width="32%"><select name = 'cboEntity<%=cnt%>' id = 'cboEntity<%=cnt%>' size="1"onchange='FetchAccounts<%=cnt%>(cboEntity<%=cnt%>)'>
    	
<%
		
        sqlStr = "SELECT  EntityType_DPA_,EntityTypeName FROM [FullEntityTypeList] Order By EntityTypeName"
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
	<option value = '100' <%if isAccManager=true then response.write "Selected"%>>Account Manager</option>
    </select></td>
    <td nowrap colspan="2">
    <input type = 'text' name ='txtClientCode<%=cnt%>' id = 'txtClientCode<%=cnt%>' size="10" onBlur="txtval = this.value; selectItem(cboAccount<%=cnt%>);UpdateCode(true,cboAccount<%=cnt%>,txtClientCode<%=cnt%>);" value="<%=trim(rsEdit("EntityCode"))%>">
    <select name = 'cboAccount<%=cnt%>' id = "cboAccount<%=cnt%>" size="1" 
    onchange='UpdateCode(true,cboAccount<%=cnt%>,txtClientCode<%=cnt%>)'
	onKeypress="return (dodefaultaction()==''); " 
	onKeydown="return (dodefaultaction()==''); " 
	onKeyup="return (FilterData(this,<%=currentEntityType%>,UpdateCode(change(cboAccount<%=cnt%>,0),cboAccount<%=cnt%>,txtClientCode<%=cnt%>)));" 
	onfocus="txtval = '';inputIsItemCode = 1;" 
	onblur="txtval = '';inputIsItemCode = 1;">
	<option selected SearchCode = "" SearchText = ""  value = ''></option>

 
<%
        
        sqlStr = "SELECT EntityName,Entity_DPA_,EntityCode FROM [tblCompleteEntityList] WHERE EntityType_DPA_ =" & currentEntityType & " Order By EntityName"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF
                AccountName=Mid(rs.Fields("EntityName"),1,20)
                %>
               <option SearchCode = "<%=rs.Fields("EntityCode")%>" SearchText = "<%=rs.Fields("EntityName")%>" value = '<%=rs.Fields("Entity_DPA_")%>' <%if trim(rs("Entity_DPA_"))= trim(rsEdit("Entity_DPA_")) then response.write "Selected"%>><%=AccountName%></option>
                        <%rs.MoveNext
                Loop
        End If
%>

    </select></td>
	  <td><TEXTAREA NAME="txtEntryNarrative<%=cnt%>" id = 'txtEntryNarrative<%=cnt%>' ROWS="1" COLS="17"><%=rsEdit("Narrative")%></TEXTAREA>
	 </td>
      <td width="12%"><input type = 'text' name ='txtDebit<%=cnt%>' id = 'txtDebit<%=cnt%>' size="12" value="<%=rsEdit("JournalEntryDebit")%>"></td>
      <td width="24%"><input type = 'text' name ='txtCredit<%=cnt%>' id = 'txtCredit<%=cnt%>' size="12" value="<%=rsEdit("JournalEntryCredit")%>"></td>
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
    <td colspan="2"><select name = 'cboEntity' id = 'cboEntity' size="1"onchange='FetchAccounts(this)'>
    	
<%
		Set conn = GetActiveConnection("KBroker")
		
		
		'delete earlier incomplete journal entries
		
		conn.BeginTrans
		
			sqlStr = "DELETE FROM JournalEntry WHERE Journal_DPA_ IN (SELECT Journal_DPA_ FROM Journal WHERE UserID = " & Session("UserID") & " AND JournalCommitted = 0)" 
			'conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		
			sqlStr = "DELETE FROM Journal WHERE UserID = " & Session("UserID") & " AND JournalCommitted = 0"
			'conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		
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
	<option value = '100'>Account Manager</option>
    </select></td>
    <td nowrap>
    <input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="8" onBlur="LoadClient(cboAccount, this.name)" onClick = "ClearFields(this.name);">
    <select name = 'cboAccount' id = "cboAccount" size="1" >
	<option selected SearchCode = "" SearchText = ""  value = ''>Load Account</option>
    </select></td>
	 <td>
	 <TEXTAREA NAME="txtEntryNarrative" id = 'txtEntryNarrative' ROWS="1" COLS="17"></TEXTAREA>
	 </td>
      <td><input type = 'text' name ='txtDebit' id = 'txtDebit' size="12"></td>
      <td><input type = 'text' name ='txtCredit' id = 'txtCredit' size="12"></td>
	   <td>&nbsp;</td>
    </tr>
	<tr> <td nowrap colspan="9" align="center"><input type ="submit" name="SaveLine" value ="SAVE" onclick="javascript:document.all.frm<%=Datasource%>.action.value='ADD';validate();">
	  <INPUT TYPE="Button" name="Cancel" value ="Cancel" onclick="javascript:document.all.AddLine.style.display='none';javascript:document.all.frm<%=Datasource%>.cmdAdd.style.display='';javascript:document.all.frm<%=Datasource%>.cmdClose.style.display='';javascript:document.all.frm<%=Datasource%>.cmdCommit.style.display='';">
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
		 <%if itemID="" then%><input type = 'button' Class=Buttons name ='cmdAdd' id = 'cmdAddMore' value=" Add Line/s " onclick="javascript:document.all.AddLine.style.display='block';javascript:document.all.frm<%=Datasource%>.cmdAdd.style.display='none';javascript:document.all.frm<%=Datasource%>.cmdClose.style.display='none';javascript:document.all.frm<%=Datasource%>.cmdCommit.style.display='none';">
			<%end if%>
		 &nbsp;&nbsp; 
		<%if itemID="" then%><input type = 'submit' Class=Buttons name ='cmdCommit' id = 'cmdCommit' value=" Commit " onclick="javascript:document.all.frm<%=Datasource%>.action.value='COMMIT';">
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
		<input type ="hidden" name="txtDelAmount" Id="txtDelAmount" value ="">
		<input type = 'hidden' name ='txtcalendar' id = 'txtcalendar' value='0'>
		
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
	 %>
	 if (checkcomboempty(doc.cboEntity<%=z%>,' Please Select an Entity')==false) return false;
	 if (checkcomboempty(doc.cboAccount<%=z%>,' Please Select an Account')==false) return false;
	 if ((trim(doc.txtDebit<%=z%>.value)=='') && (trim(doc.txtCredit<%=z%>.value)=='')){
		alert("Please enter the Debit or Credit Amount");
		doc.txtDebit<%=z%>.focus();
		return false;
	 }
 	// if (checkempty(doc.txtDebit,' Please Enter Debit ')==false) return false;
	 //if (checkempty(doc.txtCredit,' Please Enter Credit ')==false) return false;
	 if (checknumericWithCommas(doc.txtDebit<%=z%>," Debit must be numeric")==false) return false;
	 if (checknumericWithCommas(doc.txtCredit<%=z%>," Credit must be numeric")==false) return false;
	<%
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
