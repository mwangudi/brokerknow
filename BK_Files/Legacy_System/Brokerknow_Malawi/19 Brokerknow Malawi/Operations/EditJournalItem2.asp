<!--#include file="../libroutines.asp"-->


<%
	const UDLName = "KBroker"
	const DataSource = "EditJournal"
	const DataEntity = "Journal"
	const DataEntityPlural = "Journals"
	const ActionFolder = "Operations"
	
	const LinkedIndependent = 1
   const LinkedDependent = 2
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim ID
   Dim rsEdit
   dim currentEntityType
	
	action = ucase(Request.Form("action"))
	ItemID = Request.Form("ItemID")
	UserId=Session("UserID")
	
	ID = Request("ID")
	EntityID = Request("EntityID")
	currentEntityType = 11

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		
                </script>
                <% response.end
        End If
        if action <> "" then
				
				If (Trim(ItemID) = "") and (trim(EntityID) = "") Then%>
						<script language = 'vbscript'>
                				ShowMessage "No item specified"
                				
						</script>
						<%response.end
				End If
		end if
	select case action 
		case "EXECUTE_DETAIL"
   			Dim Entity
			Dim Account
			Dim debit	
			Dim credit		
	        
	        if itemID = "-1" then
					Entity = Request.Form("cboEntity")
					Account = Request.Form("cboAccount")
					debit = Request.Form("txtDebit")
					credit = Request.Form("txtCredit")
					JournalEntryNarrative = Request.Form("txtJournalEntryNarrative")
					
			else
					Entity = Request.Form("cboEntityInPlace")
					Account = Request.Form("cboAccountInPlace")
					debit = Request.Form("Debit")
					credit = Request.Form("Credit")
					JournalEntryNarrative = Request.Form("txtJournalEntryNarrative")
					
			end if

			'validate Entity
			If Trim(Entity) = "" Then%>
			        <script language = 'vbscript'>
			        		ShowMessage "Please specify the Entity"
						         		
			        </script>
			        <% 
					ReloadPage(ID)
					response.end
			End If
			'validate Account
			If Trim(Account) = "" Then%>
			        <script language = 'vbscript'>
			        		ShowMessage "Please specify the Account"
						         		
			        </script>
			        <% ReloadPage(ID)
					response.end
			End If
			'ensure Debit is numeric
			If (Debit <> "") And (Not IsNumeric(Debit)) Then%>
				<script language = 'vbscript'>
					ShowMessage "Debit must be numeric"
											
				</script>
				<%ReloadPage(ID) 
				response.end
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
	                        
			if itemID = "-1" then
					response.Write ID & chr(13)
					'save detail data
					sqlStr = "INSERT INTO [JournalEntry] (JournalEntryDebit,JournalEntryCredit" & _
							",JournalEntry_DPA_,Journal_DPA_,EntityType_DPA_,Entity_DPA_,Narrative) SELECT " & " " & debit & " " & " as JournalEntryDebit" & _
							"," & " " & credit & " " & " as JournalEntryCredit" & _
							"," & " " & "iif(isnull(max([JournalEntry_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'JournalEntry'),max([JournalEntry_DPA_]) + 1)" & " " & " as JournalEntry_DPA_" & _
							"," & " " & ID & " " & " as Journal_DPA_" & _
							"," & " " & Entity & " " & " as EntityType_DPA_" & _
							"," & " " & Account & " " & " as Entity_DPA_" & _
							"," & " " & JournalEntryNarrative & " " & " as JournalEntryNarrative" & _
							" FROM [JournalEntry]"
		    else
					'edit detail data
					sqlStr = "UPDATE JournalEntry SET JournalEntryDebit = " & debit & "," & _
							" JournalEntryCredit = " & credit & ", EntityType_DPA_ = " & Entity & "," & _
							" Entity_DPA_ = " & Account & _
							" Narrative =  '" & JournalEntryNarrative & "' " &_
							" WHERE JournalEntry_DPA_=" & itemID
					
		    end if 
		    
			sqlStr1="UPDATE Journal Set ChangedBy=" & UserId & ",TimeChanged=GetDate() Where( Journal_DPA_=" & ID & ")"
      
			Set conn = GetActiveConnection("KBroker")
	       
			conn.BeginTrans
					conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
					conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr1))
					
					if(Cint(entity)=1) then
						'conn.execute ("Exec ClientTotalProcedure " & account)							
						conn.execute ("Exec ClientBalanceProcedure " & account)		
					end if			
			conn.CommitTrans
			
			 'retrieve the item ID
				'sqlStr = "SELECT JournalEntry_DPA_  FROM JournalEntry WHERE Journal_DPA_=" & ID
				'Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
				'ID = Rs.Fields("JournalEntry_DPA_")
			
			conn.Close
			Set conn = Nothing
			WriteDialogRelocateScript DataSource & "Item.asp?ID=" & ID
			Response.End
			
		case "EXECUTE_DELETE"
			'ensure at least one detail record is left over	
			sqlStr = "SELECT COUNT(JournalEntry_DPA_) as Total FROM [JournalEntry] WHERE Journal_DPA_=" & ID & " and Deleted=0"
			Set conn = GetActiveConnection("KBroker")
	        
			Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			If rs.EOF Or rs.BOF Then%>
					<script language = 'vbscript'>
                			ShowMessage "The database is corrupted"
                			
					</script>
					<%response.end
			End If
			If (CInt(rs.Fields("Total")) < 2) Then%>
					<script language = 'vbscript'>
                			ShowMessage "There must be at least one Journal entry"
                			
					</script>
					<%response.end
			End If
	        
			'find out whether any child records exist
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
			 'retrieve the item ID
			'	sqlStr = "SELECT JournalEntry_DPA_  FROM JournalEntry WHERE Journal_DPA_=" & ID & " AND JournalEntry_DPA_ <> " & ItemID
			'	Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			'	otherItemID = Rs.Fields("JournalEntry_DPA_")
				
			
			'delete from database
			sqlStr = "Update [JournalEntry] Set Deleted=1 WHERE JournalEntry_DPA_ = " & ItemID
			conn.Execute SQLServerFormat(HandleQuote(sqlStr))
			
			conn.Close
			Set conn = Nothing
			WriteDialogRelocateScript DataSource & "Item.asp?ID=" & ID
			Response.End
		case "EXECUTE_CLOSE"
		
			'check that the journal balances
			sqlStr = "SELECT * FROM [JournalTotalList] WHERE Journal_DPA_=" & ID
			Set conn = GetActiveConnection("KBroker")
	        
			Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			If rs.EOF Or rs.BOF Then%>
					<script language = 'vbscript'>
                			ShowMessage "The database is corrupted"
                			
					</script>
					<%response.end
			End If
			If (rs.Fields("JournalDebitTotal") - rs.Fields("JournalCreditTotal")) <> 0 Then%>
					<script language = 'vbscript'>
                			ShowMessage "The debit and credit columns must add up to 0"
                			
					</script>
					<%response.end
			End If
			
			sqlStr = "UPDATE Journal SET JournalCommitted = 1 WHERE Journal_DPA_=" & ID
			conn.Execute SQLServerFormat(HandleQuote(sqlStr))
			
			WritefraEnabledDialogCloseScript
			Response.End
		case "EXECUTE_CANCEL"
			Dim journalCommitted
			
			journalCommitted = abs(cint(cbool(request("JournalCommitted"))))
			
			if journalCommitted = 0 then
					'delete the journal
					sqlStr = "DELETE FROM [JournalEntry] WHERE Journal_DPA_ = " & ID
					Set conn = GetActiveConnection("KBroker")
					conn.Execute SQLServerFormat(HandleQuote(sqlStr))
	        
					sqlStr = "DELETE FROM [Journal] WHERE Journal_DPA_ = " & ID
					conn.Execute SQLServerFormat(HandleQuote(sqlStr))
			end if
			WritefraEnabledDialogCloseScript
			Response.End
		case "FETCH_ACCOUNTS"
			currentEntityType = cint(EntityID)
   	end select
   		
   	
%>
<%
 Set conn = GetActiveConnection("KBroker")
 
	Dim rowCount
	Dim entityList
	Dim accountList
	Dim quote 
	Dim secType
	Dim entitytype
	
	'sqlStr = "SELECT * FROM JournalList WHERE Journal_DPA_ IN (SELECT Journal_DPA_ FROM JournalList WHERE " & _
    '         "	JournalEntry_DPA_=" & ID & ")"
    
    sqlStr = "SELECT * FROM JournalFullList WHERE Journal_DPA_ = " & ID
       
    Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
	
	quote =  chr(34)
	entityList = GetEntityList("cboEntity", "cboAccount","txtClientCode")
	accountList = GetAccountList("cboAccount","txtClientCode")
	     
    
    If Not(rs.EOF Or rs.BOF) Then%>
        
		<!-- grid data -->
		<% 'row data
		 rowCount = 0
		rs.MoveFirst
		Do Until rs.EOF 

'======================= Begin_Alter_Across_Entities =================================
			'row ID 
			entitytype=rs("EntityType_DPA_")
			rowData = rowData & quote & rs.Fields("JournalEntry_DPA_") & quote & " : " 
			
			'row data 
			rowData = rowData & "[" 
			rowData = rowData & quote & rs.Fields("JournalEntry_DPA_") & quote & "," 
			rowData = rowData & quote & rs.Fields("JournalEntryEntity") & quote & ","
			rowData = rowData & quote & rs.Fields("JournalEntryAccount") & quote & ","
			rowData = rowData & quote & rs.Fields("JournalEntryNarrative") & quote & ","

			
			rowData = rowData & quote & rs.Fields("JournalEntryDebit") & quote & ","
			rowData = rowData & quote & rs.Fields("JournalEntryCredit") & quote & ","
			rowData = rowData & quote & " " & quote  
			rowData = rowData & "]" 

			rowIDs = rowIDs & quote & rs.Fields("JournalEntry_DPA_") & quote 
			rowCount = rowCount + 1
			
'======================= End_Alter_Across_Entities =================================

			rs.MoveNext 
			
			'build the row IDs array
			rowIDs = rowIDs & "," 
			rowData = rowData & ","	
		Loop
		rs.MoveFirst
	End if
	
	'row ID 
	rowData = rowData & quote & -1 & quote & " : " 
			
	'row data 
	rowData = rowData & "[" 
	rowData = rowData & quote & "New Line" & quote & "," 
	rowData = rowData & quote & entityList & quote & ","
	rowData = rowData & quote & accountList & quote & ","

	rowData = rowData & quote & "<input type = 'text' name ='txtJournalEntryNarrative' id = 'txtJournalEntryNarrative' size='9' onChange = 'event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>" & quote & ","
	rowData = rowData & quote & "<input type = 'text' name ='txtDebit' id = 'txtDebit' size='12' onChange = 'event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>" & quote & ","
	rowData = rowData & quote & "<input type = 'text' name ='txtCredit' id = 'txtCredit' size='9' onChange = 'event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>" & quote & ","
	rowData = rowData & quote & "<input type=button value='Add' Class=Buttons OnClick='JavaScript: AddRowInProgress();'>&nbsp;&nbsp;<input type='reset' value='Cancel' Class=Buttons>" & quote 
	rowData = rowData & "]" 
	

	'build the row IDs array 
	rowIDs = rowIDs & quote & -1 & quote 
	rowCount = rowCount + 1
'======================= Begin_Alter_Across_Entities =================================%> 
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit <%=DataEntity%> Item</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
 <script language='vbscript'>
			function ItemSelected(itemID)
					
 					frm<%=DataSource%>Item.elements("ItemID").value = itemID
			end function
			
			
			function SaveInPlaceEdit()
				    Dim myOwnerFrame				
					'UpdateID
					Set window.parent.dialogArguments.opener.parent.frames("footer").editDocOpener = window.self
					frm<%=DataSource%>Item.target = "deleteFrame" 					
					frm<%=DataSource%>Item.submit
					'restoreID
			end function
			
			function CloseDocumentWithValidate()
				    Dim myOwnerFrame	
				    'UpdateID
				    frm<%=DataSource%>Item.elements("ItemID").value =0			
					Set window.parent.dialogArguments.opener.parent.frames("footer").editDocOpener = window.self
					frm<%=DataSource%>Item.target = "deleteFrame" 					
					frm<%=DataSource%>Item.submit
					'restoreID
			end function
			
			function CancelDocument()
				    Dim myOwnerFrame	
				    'UpdateID
				    frm<%=DataSource%>Item.elements("ItemID").value =0			
					Set window.parent.dialogArguments.opener.parent.frames("footer").editDocOpener = window.self
					frm<%=DataSource%>Item.target = "deleteFrame" 					
					frm<%=DataSource%>Item.submit
			end function
		</script>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
		
		<!-- ActiveUI stylesheet and scripts -->
		<link href="../runtime/classic/activeui.css" rel="stylesheet" type="text/css">
		<script src="../runtime/activeui.js"></script>
		<!-- Include patches here -->
		<script src="../runtime/paging1.js"></script>
		<!-- grid format -->
		
		<style> 
			.active-controls-grid {height: 100%; font: menu;}
			.active-row-highlight .active-row-cell {background-color: skyblue}
		    		    		     	
			.active-column-0 {width: 50px;}
			.active-column-1 {width: 100px;}
			.active-column-2 {width: 270px;}
			.active-column-4 {width: 100px;}
			.active-column-5 {width: 80px;}
			.active-column-6 {width: 80px;}
			.active-column-7 {width: 150px;}		
			
			.active-grid-row,
			.active-grid-row.active-list-item,
			.active-scroll-left .active-list-item {height: 22px;}
			
			
			.active-selection-true, .active-selection-true .active-row-cell {
				color: blue!important;
				background-color: bisque!important;
				}
				
		
		</style>
<script language='javascript'>
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

<body Class="Dialog" onload="setOpener()">	
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>

<form name = 'frm<%=DataSource%>Item' id='frmMain' method = 'post' action = '<%=DataSource%>Item2.asp'>

<script language='vbscript'>

					function EntitySelected(itemID)
 							frm<%=DataSource%>Item.elements("EntityID").value = itemID
 							frm<%=DataSource%>Item.elements("action").value = "Fetch_Accounts"
 							frm<%=DataSource%>Item.submit
 							
					end function
</script>

		<script language="javascript">
		
			//column titles 
			var colCount = 7;
			var colNames = ["", "Entity",  
					 "Account","Narrative", "Debit","Credit", ""];
			
			var myColumns = ["Line No", "Entity",  
					 "Account","Narrative", "Debit","Credit", ""];
		</script>
<%'======================= End_Alter_Across_Entities =================================%>			
		<script language="javascript">
			
			//data
			var myData = {<%=rowData%>}; 
			var myRowIDs = [<%=rowIDs%>]; 
			
			
			//editing
			var inPlaceEdit = false;
			var addInProgress = false;
			var clickedRowID = -1; 
			var dataChanged = false;
			var prevRow = -1;//the row currently under in-place edit
			
			function EditInPlaceDataChanged()
			{
				dataChanged = true;
			}
			
			function AddRowInProgress()
			{
				addInProgress = true;
			}
			
			var currentEntityName = "";
			var currentAccountName = "";
			var RowEditFn = function(src)
			{
				var rowIndex = src.getProperty("row/index");
				var i;
				if (rowIndex==0) return;
				for(i = 0; i < colCount; i++)
				{
					//alert(colNames[i]+':::'+myData[prevRow][i]);
					if(colNames[i] != "")
					{
						if(prevRow >= 0)
						{
							if(colNames[i]=="Entity")
							{
								myData[prevRow][i] = currentEntityName;
							}
							else
							{
								if(colNames[i]=="Account")
								{
									myData[prevRow][i] = currentAccountName;
								}
								else
								{
									myData[prevRow][i] = document.frmMain.elements(colNames[i]).value;
								}
							}
						}
						if(colNames[i]=="Entity")
						{	
							currentEntityName =  myData[rowIndex][i];
							myData[rowIndex][i] = inPlaceEntityList;
							
						}
						else
						{
							if(colNames[i]=="Account")
							{	
								currentAccountName =  myData[rowIndex][i];
								myData[rowIndex][i] = inPlaceAccountList;
								
							}
							else
							{
								myData[rowIndex][i] = "<INPUT TYPE='text' NAME='" + colNames[i] + "' ID='" + colNames[i] + "' VALUE='" + myData[rowIndex][i] + "' onChange = 'EditInPlaceDataChanged();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>";																
								
							}
						}
						
					}
				}
				
				
				myData[rowIndex][colCount - 1] = "<INPUT TYPE='button' class='Buttons' VALUE='Save' onClick = 'SaveInPlaceEdit();event.cancelBubble=true;'>&nbsp;<INPUT TYPE='button' class='Buttons' VALUE='Cancel' onClick = 'cancelEditRow();event.cancelBubble=true;'>";				
				inPlaceEdit = true;
				prevRow = rowIndex;
				grid.refresh();
				//select the appropriate item
				var entityList = document.frmMain.elements("cboEntityInPlace");
				var accountList = document.frmMain.elements("cboAccountInPlace");
				
				for (i=0; i < entityList.options.length; i++) {
					if(entityList.options(i).text == currentEntityName)
					{
							entityList.options(i).selected = true;
					}
				}
				
				//get the right account list
				var itemFound = false;
				
				//accountList = document.frmMain.elements("cboAccountInPlace");
				
				for (i=0; i < accountList.options.length; i++) {
					if(accountList.options(i).text == currentAccountName)
					{
							accountList.options(i).selected = true;
							itemFound = true;
					}
				}
				
				if(!itemFound)
				{
					var currEntity = entityList.value;
					FetchAccounts(entityList, accountList, document.frmMain.txtClientCodeInPlace);
					accountList = document.frmMain.elements("cboAccountInPlace");
				
					for (i=0; i < accountList.options.length; i++) {
						if(accountList.options(i).text == currentAccountName)
						{
								accountList.options(i).selected = true;
								itemFound = true;
						}
					}
				}
			}
			
			var RowChangeFn = function(src)
			{
				if(inPlaceEdit || addInProgress)
				{
					if(dataChanged || addInProgress)
					{
						ItemSelected(prevRow);
						SaveInPlaceEdit();
					}
					else
					{
						if(prevRow != clickedRowID)
						{
							cancelEditRow();
						}
					}
				}
			}
			
			var HandleClick = function(src)
			{
				clickedRowID = src.getProperty("row/index");
				ItemSelected(clickedRowID);
			}
			
			var headerID;
			
			
			function cancelEditRow(){
					var i;
					for(i = 0; i < colCount; i++)
					{
						if(colNames[i] != "")
						{
							if(colNames[i]=="Entity")
							{
								myData[prevRow][i] = currentEntityName;
							}
							else
							{
								if(colNames[i]=="Account")
								{
									myData[prevRow][i] = currentAccountName;
								}
								else
								{
									myData[prevRow][i] = document.frmMain.elements(colNames[i]).value;
									
								}
							}
						}
					}
					myData[prevRow][colCount - 1] = "";
					inPlaceEdit = false;
					prevRow = -1;
					grid.refresh();
			}
			
			function UpdateID(){
				headerID = window.parent.frames["header"].document.all.item("ID").value;
				document.all.item("ID").value = headerID;
			}
			
			function restoreID(){
				document.all.item("ID").value = headerID;
			}
			
			function HandleDeleteAction()
			{
					document.frmMain.elements("action").value = "Execute_Delete";
					SaveInPlaceEdit();
					document.frmMain.elements("action").value = "Execute_Detail";
			}
			
			function HandleCloseAction()
			{
					
					document.frmMain.elements("action").value = "Execute_Close"
					CloseDocumentWithValidate();
					document.frmMain.elements("action").value = "Execute_Detail"
			}
			
			function HandleCancelOperation(journalCommitted)
			{
					//alert(journalCommitted);
					document.frmMain.elements("JournalCommitted").value = journalCommitted;
					document.frmMain.elements("action").value = "Execute_Cancel"
					CancelDocument();
			}
			
			function FetchAccounts(theList, toList, codeInput)
			{
				var i = 0, xmlhttp, frm, url, returnStr;
				var entity = theList.value;
				
				
				frm = document.frmMain;				
				xmlhttp = createXMLHTTPObj();
				
				url="GetList.asp?ID="+entity+"&action=GetAccountList";
				xmlhttp.open("GET",url,true);
				xmlhttp.onreadystatechange=function() {
				  if (xmlhttp.readyState==4) {
					returnStr = xmlhttp.responseText;
					returnStr = getBodyHTML(returnStr);
					
					var secList = "<select name = '" + toList.name + "' id = '" + toList.name + "' size='1' style='width:290px'";
					secList += "OnClick='event.cancelBubble=true;' " ;
					secList += "onChange='UpdateCode(true," + toList.name + "," + codeInput.name + "); event.cancelBubble=true;' " ;
					secList += "onKeypress='return (dodefaultaction()==\"\"); ' "  ;
					secList += "onKeydown='return (dodefaultaction()==\"\");' " ; 
					secList += "onKeyup='return (UpdateCode(change(" + toList.name + ",0)," + toList.name + "," + codeInput.name + "));' " ; 
					secList += "onfocus='txtval = \"\";inputIsItemCode = 1;' "  ;
					secList += "onblur='txtval = \"\";inputIsItemCode = 1;'>" ;
					secList += returnStr ;
					secList += "</select>";
					
					toList.outerHTML = secList;															
				  }
				 }
				xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
				xmlhttp.send();
				
				document.getElementById(codeInput.name).value = '';
				
			}
			
			//get ready for in-place edit
			var inPlaceEntityList = "<%=GetEntityList("cboEntityInPlace", "cboAccountInPlace","txtClientCodeInPlace")%>"
			var inPlaceAccountList = "<%=GetAccountList("cboAccountInPlace","txtClientCodeInPlace")%>"
		</script> 
		
		<script language="javascript"> 

			// create ActiveUI Grid javascript object 
			var grid = new Active.Controls.Grid; 
			//grid.setModel("row", new Active.Rows.Page);
			
			// set rows ids 
			//grid.setProperty("row/count", <%=rowCount%>);
			//grid.setProperty("row/values", myRowIDs);
			grid.setRowValues(myRowIDs); 
			

			// set number of columns 
			//grid.setProperty("column/count", colCount);
			grid.setColumnCount(colCount); 

			// provide cells and headers text 
			//grid.setProperty("data/text", function(i, j){return myData[i][j]});
			grid.setDataText(function(i, j){return myData[i][j]}); 
			//grid.setProperty("column/texts", myColumns);
			grid.setColumnText(function(i){return myColumns[i]}); 

			//introduce paging
			//grid.setProperty("row/pageSize", 5);
			
			// set click action handler 
			grid.setAction("click", HandleClick); 
			grid.setAction("dblclick", RowEditFn); 
			grid.setAction("selectionChanged", RowChangeFn); 

			//stripes 
			var alternate = function(){ return this.getProperty("row/order") % 2 ? "gainsboro" : "white";} 
			var row = new Active.Templates.Row; row.setStyle("background", alternate); 
			row.setEvent("onmouseover", "mouseover(this, 'active-row-highlight')"); 
			row.setEvent("onmouseout", "mouseout(this, 'active-row-highlight')"); 
			grid.setTemplate("row", row); 
			var column = new Active.Templates.Text; 
			column.setStyle("border-right", "1px solid white");  
			grid.setTemplate("column", column);  grid.setRowHeaderWidth("0px"); 

			// write grid html to the page 
			document.write(grid); 
			//ShowMessage(grid.toString());
		</script> 
		
        <%
       
  function GetAccountList(listName,txtName)
		Dim accList
		Dim rsAccount

		accList = "<input type = 'text' name ='" & txtName & "' id = '" & txtName & "' size='9' onBlur='txtval = this.value; selectItem(" & listName & ");' onChange = 'event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>"
		accList = accList & "&nbsp;<select name = '" & listName & "' id = '" & listName & "' size='1' style='width:290px'"
		accList = accList & "OnClick='event.cancelBubble=true;' "  
		accList = accList & "onChange='UpdateCode(true," & listName & "," & txtName & "); event.cancelBubble=true;' " 
		accList = accList & "onKeypress='return (dodefaultaction()==\""\""); ' "  
		accList = accList & "onKeydown='return (dodefaultaction()==\""\"");event.cancelBubble=true;' "  
		accList = accList & "onKeyup='return (UpdateCode(change(" & listName & ",0)," & listName & "," & txtName & "));' "  
		accList = accList & "onfocus='txtval = \""\"";inputIsItemCode = 1;' "  
		accList = accList & "onblur='txtval = \""\"";inputIsItemCode = 1;'>"
		
		accList = accList & "<option selected SearchCode = '0' SearchText = ''  value = ''></option>"
		
		Dim displayField
		
		displayField = "EntityNameEx"
		
        sqlStr = "SELECT * FROM [CompleteEntityList] WHERE EntityType_DPA_ =" & currentEntityType & " Order By EntityName"
        Set rsAccount = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsAccount.EOF Or rsAccount.BOF) Then
                rsAccount.MoveFirst
                Do Until rsAccount.EOF
                AccountName=Mid(Trim(rsAccount.Fields(displayField)),1,15)
                        accList = accList & "<option SearchCode = '" & rsAccount.Fields("EntityCode") & "' SearchText = '" & AccountName & "'  value = '" & rsAccount.Fields("Entity_DPA_") & "'>" & AccountName & "</option>"
                        rsAccount.MoveNext
                Loop
        End If
	    accList = accList & "</select>"
	    GetAccountList = accList
  end function
  
  function GetEntityList(listName, tolistName, codeInput)
		Dim entList
		Dim rsEntity
		entList = "<select name = '" & listName & "' id = '" & listName & "' size='1' onChange = 'FetchAccounts(this, document.frmMain." & tolistName & ", document.frmMain." & codeInput & ");event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>"
		entList = entList & "<option selected value = ''></option>"
		
        sqlStr = "SELECT * FROM [FullEntityTypeList] Order By EntityTypeName"
        Set rsEntity = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsEntity.EOF Or rsEntity.BOF) Then
                rsEntity.MoveFirst
                Do Until rsEntity.EOF
                        if rsEntity.Fields("EntityType_DPA_").value = currentEntityType then
								entList = entList & "<option selected value = '" & rsEntity.Fields("EntityType_DPA_") & "'>" & rsEntity.Fields("EntityTypeName") & "</option>"
                        else
								entList = entList & "<option value = '" & rsEntity.Fields("EntityType_DPA_") & "'>" & rsEntity.Fields("EntityTypeName") & "</option>"
                        end if
                        rsEntity.MoveNext
                Loop
        End If
	    entList = entList & "</select>"
	    GetEntityList = entList
  end function
 %>
 <table border="0" width="100%" ID="Table1">
<tr><td>
<input type = 'hidden' name ='ItemID' id = 'ItemID'>
<input type = 'hidden' name ='ID' id = 'ID' value="<%= ID %>">
<input type = 'hidden' name ='action' id = 'action' value="Execute_Detail">
<input type = 'hidden' name ='EntityID' id = 'EntityID'>
<input type = 'hidden' name ='JournalCommitted' id = 'JournalCommitted' value="0" >
</td>
</tr>
</table>
</form>
</body>

</html>