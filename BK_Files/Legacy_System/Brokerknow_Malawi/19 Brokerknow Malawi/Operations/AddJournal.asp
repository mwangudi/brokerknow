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
				buttonAction = Trim(Trim(Ucase(Request.Form("buttonAction"))))
		
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
						Entrynarrative = Request.Form("txtEntryNarrative")

		
						 'validate Entity
						 If Trim(Entity) = "" Then%>
						         <script language = 'vbscript'>
						         		ShowMessage "Please specify the Entity"
						         		
						         </script>
						         <% ReloadPage(ID)
							 response.end
						 End If
						 'validate Account
						 If Trim(Account) = "" or ucase(trim(Account))="CODE" Then%>
						         <script language = 'vbscript'>
						         		ShowMessage "Please specify the Account"
						         		
						         </script>
						         <% ReloadPage(ID)
							 response.end
						 End If

						 'validate Narrative
						 If Trim(narrative) = "" or len(Trim(narrative)) < 1 Then%>
						         <script language = 'vbscript'>
						         		ShowMessage "Please specify the Narrative"						         		
						         </script>
						         <% ReloadPage(ID)
							 response.end
						 End If

						 'validate size of Narrative
						 If Len(Narrative) > 500 Then%>
						         <script language = 'vbscript'>
						         ShowMessage "Narrative can only be 500 characters in length"
						         
						         </script>
						         <% ReloadPage(ID)
							 response.end
						 End If
						         							
						'ensure Debit is numeric
						If (Debit <> "") And (Not IsNumeric(Debit)) Then%>
							<script language = 'vbscript'>
								ShowMessage "Debit must be numeric"
											
							</script>
							<% ReloadPage(ID)
							 response.end
						End If
						'ensure Credit is numeric
						If (Credit <> "") And (Not IsNumeric(Credit)) Then%>
							<script language = 'vbscript'>
								ShowMessage "Credit must be numeric"
											
							</script>
							<% ReloadPage(ID)
							 response.end
						End If
						 
						 if (trim(debit) = "") and (trim(credit) = "") then%>
								<script language = 'vbscript'>
									ShowMessage "You must enter either a debit amount or a credit amount"
												
								</script>
								<% ReloadPage(ID)
							 response.end
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

						sqlStr = "INSERT INTO [Journal] (JournalDate,UserID,JournalNarrative,Journal_EIT_,ChangedBy) Values( " & "'" & FormatDate(jDate) & "'," &  UserID  & _
						         "," & "'" & narrative & "'" & _
						         "," & "'" & guidStr & "'" & _
						         "," & " " & UserId & ")" 
								
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
						         			<% ReloadPage(ID)
							 response.end
								End If
						     
								'save detail data
								sqlStr = "INSERT INTO [JournalEntry] (JournalEntryDebit,JournalEntryCredit" & _
										",JournalEntry_DPA_,Journal_DPA_,EntityType_DPA_,Entity_DPA_,Narrative) SELECT " & " " & debit & " " & " as JournalEntryDebit" & _
										"," & " " & credit & " " & " as JournalEntryCredit" & _
										"," & " " & "iif(isnull(max([JournalEntry_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'JournalEntry'),max([JournalEntry_DPA_]) + 1)" & " " & " as JournalEntry_DPA_" & _
										"," & " " & rs.Fields("Journal_DPA_") & " " & " as Journal_DPA_" & _
										"," & " " & Entity & " " & " as EntityType_DPA_" & _
										"," & " " & Account & " " & " as Entity_DPA_" & _
										 ",'"  & Entrynarrative & "' " & _
										" FROM [JournalEntry]"
						     
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

						%>
						<SCRIPT LANGUAGE="JAVASCRIPT">
							//alert('hapa');
							//window.parent.parent.frames['maininfo'].location.reload();
							window.opener.location= window.opener.location;
							//alert('hapa');
							window.self.location.href='EditJournal.asp?ID=<%=rs.Fields("Journal_DPA_")%>&IDHeld=Journal';
							
						</SCRIPT><%
						'WriteDialogRelocateScript "Edit" & DataEntity & ".asp?ID=" & rs.Fields("Journal_DPA_") & "&IDHeld=Journal"
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
			forceSubmit()
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



			function ClearFields(element)
		{
		   if (element == 'txtClientCode')
		   {
			document.frm<%=DataSource%>.elements("txtClientCode").value = '';
			document.frm<%=DataSource%>.elements("txtcdsno").value = 'CDS No.';
			document.frm<%=DataSource%>.elements("txtclientname").value = 'Client Name';
			return;
		   }
		   if (element == 'txtcdsno')
		   {
			document.frm<%=DataSource%>.elements("txtClientCode").value = 'Code';
			document.frm<%=DataSource%>.elements("txtcdsno").value = '';
			document.frm<%=DataSource%>.elements("txtclientname").value = 'Client Name';
			return;
		   }
		   if (element == 'txtclientname')
		   {
		    document.frm<%=DataSource%>.elements("txtclientname").value = '';
			document.frm<%=DataSource%>.elements("txtClientCode").value = 'Code';
			document.frm<%=DataSource%>.elements("txtcdsno").value = 'CDS No.';
			return;
		   }		
		   
		}

		function updatefields(selectedclient)
		{
		 document.frm<%=DataSource%>.elements("txtclientname").value = '';
		 document.frm<%=DataSource%>.elements("txtcdsno").value = '';
		 document.frm<%=DataSource%>.elements("txtClientCode").value = selectedclient.value;
		 //alert(selectedclient.value);
			LoadMyClient();
		}

		function LoadMyClient()
		{
			var clientcode = document.frm<%=DataSource%>.elements("txtClientCode").value
			var clientcds = document.frm<%=DataSource%>.elements("txtCdsNo").value
			var clientcobo = document.getElementById("cboAccount");	
			var entitycobo = document.getElementById("cboEntity");
			var guidstr = Math.random();
			//alert(entitycobo.value);
			if (entitycobo.value != 1) 
			 {
				return;
			 }
			xmlhttp = createXMLHTTPObj();
				
			url="GetList.asp?clientcode="+clientcode+"&cdsno="+clientcds+"&clientname=&action=SLoadClient&guidstr="+guidstr;
			
			xmlhttp.open("GET",url,true);

			xmlhttp.onreadystatechange=function() 
			{
				if (xmlhttp.readyState==4) 
				{
					returnStr = xmlhttp.responseText;
					returnStr = getBodyHTML(returnStr);
									
					myArray = returnStr.split("<->");
					//alert(myArray);
									
					document.frm<%=DataSource%>.elements("txtclientname").value = myArray[7];
					document.frm<%=DataSource%>.elements("txtClientCode").value = myArray[5];
					document.frm<%=DataSource%>.elements("txtCdsNo").value = myArray[9]; 
				}
		   }
				 
		xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
		xmlhttp.send();
		}

		function LoadClient(accountno, element)
		{
		 var clientcode = document.frm<%=DataSource%>.elements("txtClientCode").value;
		 var clientcds = document.frm<%=DataSource%>.elements("txtcdsno").value;
		 var clientname = document.frm<%=DataSource%>.elements("txtclientname").value;
		 var clientcobo = document.getElementById("cboAccount");		 
		 var entitycobo = document.getElementById("cboEntity");
		 //alert(entitycobo.value)
	     var x_clientname;
		
		 if (entitycobo.value != 1) 
		 {
		 	return;
		 }
		 
		 if (element == 'txtClientCode')
		 {
			clientcds = '';
			clientname = '';
			
			if (clientcode == '')
			{
			document.frm<%=DataSource%>.elements("txtClientCode").value = 'Code'
			clientcobo.length = 1;
			clientcobo[0].text = 'Load Account';
			clientcobo[0].value = '';
			//loadIframe(0);
			return;
			}
			
		 }
		 else if (element == 'txtcdsno')
		 {
			clientcode = '';
			clientname = '';

			if (clientcds == '')
			{
			document.frm<%=DataSource%>.elements("txtcdsno").value = 'CDS No.'
			clientcobo.length = 1;
			clientcobo[0].text = 'Load Account';
			clientcobo[0].value = '';
			//loadIframe(0);
			return;
			}
						
		 }
		 else if (element == 'txtclientname')
		 {
			clientcode = '';
			clientcds = '';

			if (clientname == '')
			{
			document.frm<%=DataSource%>.elements("txtclientname").value = 'Client Name';
			clientcobo.length = 1;
			clientcobo[0].text = 'Load Account';
			clientcobo[0].value = '';
			//loadIframe(0);
			return;
			}
			
		 }

				xmlhttp = createXMLHTTPObj();
				
				url="GetList.asp?clientcode="+clientcode+"&cdsno="+clientcds+"&clientname="+clientname+"&action=SLoadClient";
				
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
								
								x_clientname = myArray[7];

								if (x_clientname.length > 12) 
								{
									x_clientname = x_clientname.substring(0,16)  + '...';
								}
								
								//document.getElementById("cboAccount").options.length = 0;
								clientcobo.length = 1;
								if (element != 'txtclientname')
								{
									
									document.frm<%=DataSource%>.elements("txtclientname").value = x_clientname;
									document.frm<%=DataSource%>.elements("txtClientCode").value = myArray[5];
									document.frm<%=DataSource%>.elements("txtCdsNo").value = myArray[9]; 
									
									clientcobo[0].Credit = myArray[0];
									clientcobo[0].CurrentBal = myArray[1];
									clientcobo[0].Agent = myArray[2];
									clientcobo[0].Owner = myArray[3];
									clientcobo[0].AgentID = myArray[4];
									clientcobo[0].SearchCode = myArray[5];
									clientcobo[0].OrderContact = myArray[6];
									clientcobo[0].SearchText = myArray[7];
									clientcobo[0].OwnerID = myArray[8];
									clientcobo[0].SearchCDS = myArray[9];
									clientcobo[0].IsCustodian = myArray[10];
									
									clientcobo[0].text = myArray[7];
									clientcobo[0].value = myArray[5];

								}
								else
								{
									var myArrayx;
									var myArrayz;
									
									//alert(returnStr);
									myArrayx = returnStr.split("|");
									myArrayxsize = myArrayx.length - 1;
									
									//alert(myArrayxsize);

									for (i=myArrayxsize; i>=0; i--)
									{
										
										myArrayz = myArrayx[i].split("<->");

										//alert(myArrayz)
										
										document.frm<%=DataSource%>.elements("txtClientCode").value = '';
										document.frm<%=DataSource%>.elements("txtcdsno").value = '';
										document.frm<%=DataSource%>.elements("txtclientname").value = '';
										document.frm<%=DataSource%>.elements("txtClientCode").value = myArray[5];
										document.frm<%=DataSource%>.elements("txtCdsNo").value = myArray[9]; 
										document.frm<%=DataSource%>.elements("txtClientname").value = myArray[7];
										
										document.getElementById("cboAccount").options[i] = new Option(myArrayz[7],myArrayz[5],myArrayz[6],myArrayz[10],myArrayz[4],myArrayz[8],myArrayz[3],myArrayz[3],myArrayz[0],myArrayz[1],myArrayz[5],myArrayz[7],myArrayz[9]);
																				
									}
									
								}
							}
					}
				 
				 xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
				 xmlhttp.send();
		}

	function evaluateEntity(Val, Entity)
	{
		var clientcobo = document.getElementById("cboAccount");		 
      //Enable Printing if entity is Broker or Client
	  
	  
	  if (Val != 1) 
	  {
			document.frm<%=DataSource%>.elements("txtClientCode").disabled  = true;
			document.frm<%=DataSource%>.elements("txtcdsno").disabled  = true;
			document.frm<%=DataSource%>.elements("txtclientname").disabled  = true;
	  }
	  else
	  {
	  		document.frm<%=DataSource%>.elements("txtClientCode").disabled  = false;
			document.frm<%=DataSource%>.elements("txtcdsno").disabled  = false;
			document.frm<%=DataSource%>.elements("txtclientname").disabled  = false;
	  }
	  
	  if (Val == 1) 
	  {
	 		clientcobo.length = 1;
			clientcobo[0].text = 'Load Account';
			clientcobo[0].value = '';
			
			document.frm<%=DataSource%>.elements("txtClientCode").value = 'Code';
			document.frm<%=DataSource%>.elements("txtcdsno").value = 'CDS No.';
			document.frm<%=DataSource%>.elements("txtclientname").value = 'Client Name';
			
			return;
	  }
	  //Furher Processing
	  FetchAccounts(Entity)
	}

	function FetchAccounts(theList)
		{
			var i = 0;
			var entity = theList.value;
			var toList = document.frm<%=DataSource%>.cboAccount;
			
			currentEntityType = entity;
			frm = document.frm<%=DataSource%>;				
			xmlhttp = createXMLHTTPObj();
			
			url="GetList.asp?ID="+entity+"&action=GetAccountList";
			xmlhttp.open("GET",url,true);
			xmlhttp.onreadystatechange=function() 
			{
				if (xmlhttp.readyState==4) 
				{
					returnStr = xmlhttp.responseText;
					returnStr = getBodyHTML(returnStr);
					
					var secList = "<select name = '" + toList.name + "' id = '" + toList.id + "' size='1' ";
					secList += "onChange='updatefields(this)'>"; 
					secList += returnStr ;
					secList += "</select>";
					
					toList.outerHTML = secList;															
				}
			}
			xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
			xmlhttp.send(); 
			totalContractAmt = 0;	
			
		}


		function forceSubmit()
		{
			//setOpener();
			//var targetPage = window.dialogArguments.opener.document.all.item("frmMain").elements("EditPage").value;
					
			document.frm<%=DataSource%>.method='post';
			document.frm<%=DataSource%>.target='_self';
			document.frm<%=DataSource%>.submit();	
			
		}
		
		function setOpener()
		{

			window.self.opener = window.dialogArguments.opener;
					//alert(window.dialogArguments.opener.location);
		}
		
</script>
</head>

<body Class="Dialog" onLoad="javascript: setOpener()" >

<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>

<form name = 'frm<%=DataSource%>' method = 'post' action = '<%=DataSource%>.asp' id = 'frmMain'>


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
 <!-- <td colspan = '2' width="502">-->
   <td width="5%"><b><font color="#000080">Entity</font></b></td>
   <td width="32%"><select name = 'cboEntity' id = 'cboEntity' size="1" onchange='evaluateEntity(this.value,this)'>
    	
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
   </tr>
	<td ><b><font color="#000080">Account</font></b></td>
	<td width="50%" nowrap>
    <input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="5" value = 'Code' onBlur="txtval = this.value;LoadClient(cboAccount, this.name);"  onClick  = "ClearFields(this.name)">
	&nbsp;
	<input style="display:none" type = 'text' name ='txtcdsno' id = 'txtcdsno' size="10" value = 'CDS NO.' onBlur="txtval = this.value;LoadClient(cboAccount, this.name);" onClick  = "ClearFields(this.name)">

	<input type = 'text' name ='txtclientname' id = 'txtclientname' size="15" value = 'Client Name' onBlur="txtval = this.value;LoadClient(cboAccount, this.name);" onClick  = "ClearFields(this.name)">

	<select name = 'cboAccount' id = "cboAccount" size="1" onChange ="updatefields(this)">		
		<option SearchCode = "" SearchText = "" value = '' >Load Account</option>
	</select> </td>
   <tr>
</tr>

  
 <!--   <table border="0" width="700">
    <tr>
      <td width="5%"><b><font color="#000080">Entity</font></b></td>
      <td width="85%"><b><font color="#000080">Account</font></b></td>
      <td width="5%"><b><font color="#000080">Debit</font></b></td>
      <td width="5%"><b><font color="#000080">Credit</td>
    </tr>
    <tr>
      
    
      <td width="12%"><input type = 'text' name ='txtDebit' id = 'txtDebit' size="9"></td>
      <td width="24%"><input type = 'text' name ='txtCredit' id = 'txtCredit' size="18"></td>
    </tr>
  </table>
  
  </td>
  </tr>-->
   <tr> 
      <td height="119" colspan = '2'> <table border="0" width="98%">
          <tr> 
            <td width="44%" ><b></b></td>
            <td width="27%" ><b></b></td>
            <td width="29%" ><b></b></td>
          </tr>
          <tr valign="bottom"> 
            <td><b><font color="#000080">Entry Narrative</font></b> <TEXTAREA NAME="txtEntryNarrative" id = 'txtEntryNarrative' ROWS="1" COLS="17"></TEXTAREA></td>
            <td ><b><font color="#000080">Debit</font></b> <input type = 'text' name ='txtDebit' id = 'txtDebit' size="12"></td>
            <td ><b><font color="#000080">Credit</font></b> <input type = 'text' name ='txtCredit' id = 'txtCredit' size="12"></td>
          </tr>
        </table>
       <table>
		<tr>
			<td width="20%" colspan=4 align=left>
				<input type = 'button' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Continue " onclick = "AllowedNavigation()">
				&nbsp;&nbsp; <input type = 'button' Class=Buttons name ='cmdClose' id = "cmdClose" value=" Close " onclick = "window.self.close();">
				<input type = 'hidden' name ='action' id = 'action' value="Execute">
				<input type = 'hidden' name ='ID' id = 'ID'>
				<input type = 'hidden' name ='buttonAction' id = 'action' value="CONTINUE">
			</td>
		  </tr>
	</table>
        <p>&nbsp;</p></td>
    </tr>
 
</table>
<!--  
-->
</form></body>

</html>