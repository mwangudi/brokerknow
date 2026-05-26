<!--#include file="../libroutines.asp"-->


<%
	const UDLName = "KBroker"
		const DataSource = "BnkBranchList"
		const DataEntity = "Branch"
		const DataEntityPlural = "Bank Branches"
		const ActionFolder = "Data"
	
	const LinkedIndependent = 1
   const LinkedDependent = 2
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim ID
   Dim rsEdit
	
	action = ucase(Request.Form("action"))
	ItemID = Request.Form("ItemID")
	ID = Request("ID")


		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		alert "No record specified for editing"
                		window.self.close
                </script>
                <% response.end
        End If
        if action <> "" then
				If Trim(ItemID) = "" Then%>
						<script language = 'vbscript'>
                				alert "No item specified"
                				
						</script>
						<%response.end
				End If
		end if

	select case action 
		case "EXECUTE_DETAIL"
			Dim name    
			Dim branchCode
			Dim swiftCode 
			
			ItemID = Request.Form("ItemID")   
			
			If ItemID = "-1" Then 
					name = Request.Form("txtName")
					branchCode = Request.Form("txtBranchCode")
					swiftCode = Request.Form("txtBranchSwiftCode")
			else
					name = Request.Form("Branch")
					branchCode = Request.Form("BranchCode")
					swiftCode = Request.Form("SwiftCode")
			end if
						       
			 'validate Name
			If Trim(Name) = "" Then%>
			        <script language = 'vbscript'>
			        		alert "Please specify the Name"
			        		
			        </script>
			        <% response.end
			End If
			'validate size of Name
			If Len(name) > 100 Then%>
			        <script language = 'vbscript'>
			        alert "Description can only be 100 characters in length"
			        
			        </script>
			        <% response.end
			End If
			'validate size of Bank Branch Code
            If Len(branchCode) > 20 Then%>
                	<script language = 'vbscript'>
                    alert "Bank Branch Code can only be 20 characters in length"
                    
                	</script>
                	<% response.end
            End If
            'validate size of Bank Branch Swift Code
            If Len(swiftCode) > 20 Then%>
                	<script language = 'vbscript'>
                    alert "Bank Branch Swift Code can only be 20 characters in length"
                    
                	</script>
                	<% response.end
            End If
        
			Set conn = GetActiveConnection("KBroker")
       
			'save data
			If ItemID <> "-1" Then
				'update
				sqlStr = "UPDATE [BnkBranch] SET BnkBranchName  = " & "'" & name & "'" & _
						",BnkBranchCode  = " & "'" & branchCode & "'" & _
						",BnkBranchSwiftCode  = " & "'" & swiftCode & "'" & " WHERE BnkBranch_DPA_  = " & ItemID
			Else
				 'save data
				sqlStr = "INSERT INTO [BnkBranch] (Bank_DPA_,BnkBranchName,BnkBranchCode,BnkBranchSwiftCode,BnkBranch_DPA_) SELECT " & ID & " AS Bank_DPA_, " & "'" & name & "'" & " as BnkBranchName" & _
					   ", " & "'" & branchCode & "'" & " as BnkBranchCode" & _
					   ", " & "'" & swiftCode & "'" & " as BnkBranchSwiftCode" & _
					   "," & " " & "iif(isnull(max([BnkBranch_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'BnkBranch'),max([BnkBranch_DPA_]) + 1)" & " " & " as Branch_DPA_" & _
						" FROM [BnkBranch]"
                
			End If   
			
			conn.BeginTrans
			        conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
			conn.CommitTrans
			
			sqlStr = "SELECT BnkBranch_DPA_ FROM BnkBranch WHERE " & _
					"	Bank_DPA_= " & ID
					
			Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			ID = Rs.Fields("BnkBranch_DPA_").Value

			conn.Close
			Set conn = Nothing
			WriteDialogRelocateScript "BankListItem.asp?ID=" & ID
			
			Response.End
   	
			
		case "EXECUTE_DELETE"	
			Set conn = GetActiveConnection("KBroker")
			
			'ensure at least one detail record is left over	
			sqlStr = "SELECT COUNT(BnkBranch_DPA_) as Total FROM [BnkBranch] WHERE Bank_DPA_=" & ID
			        
			Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			If rs.EOF Or rs.BOF Then%>
					<script language = 'vbscript'>
                			MsgBox "The database is corrupted"
                			
					</script>
					<%response.end
			End If
			If (CInt(rs.Fields("Total")) < 2) Then%>
					<script language = 'vbscript'>
                			MsgBox "There must be at least one Bank Branch"
                			
					</script>
					<%response.end
			End If
			
			'find out whether any child records exist
			sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = 'BnkBranch') AND (ChildType = " & LinkedIndependent & ")"
			Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			If Not (rs.BOF Or rs.EOF) Then
					Dim childRS
					Dim tableName
	                
					rs.MoveFirst
					Do Until rs.EOF
                				tableName = rs.Fields("Child")
							sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE BnkBranch_DPA_ = " & ItemID
							Set childRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
							If Not (childRS.BOF Or childRS.EOF) Then%>
                					<script language = 'vbscript'>
                						MsgBox "<%=rs.Fields("DeletionMessage")%>"
                						
                					</script>
                					<%response.end
							End If
							rs.MoveNext
					Loop
			End If
			
					
			'delete from database
			sqlStr = "DELETE FROM [BnkBranch] WHERE BnkBranch_DPA_ = " & ItemID
			conn.Execute SQLServerFormat(HandleQuote(sqlStr))
			sqlStr = "SELECT BnkBranch_DPA_ FROM BnkBranch WHERE " & _
					"	Bank_DPA_= " & ID
					
			Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			ID = Rs.Fields("BnkBranch_DPA_").Value

			conn.Close
			Set conn = Nothing
			WriteDialogRelocateScript "BankListItem.asp?ID=" & ID
			
			Response.End
   	end select
   		
   	
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
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
					UpdateID
					Set window.parent.dialogArguments.opener.parent.frames("footer").editDocOpener = window.self
					'frm<%=DataSource%>Item.target = "deleteFrame" 					
					frm<%=DataSource%>Item.submit
					'forceSumbmit()
					restoreID
			end function
		</script>

		<!-- ActiveUI stylesheet and scripts -->
		<link href="../runtime/classic/activeui.css" rel="stylesheet" type="text/css">
		<script src="../runtime/activeui.js"></script>
		<!-- Include patches here -->
		<script src="../runtime/paging1.js"></script>
		<!-- grid format -->
		<style> .active-controls-grid {height: 100%; font: menu;}
			.active-row-highlight .active-row-cell {background-color: skyblue}
			.active-selection-true, .active-selection-true .active-row-cell {
				color: blue!important;
				background-color: bisque!important;
				}
			.active-grid-row,
			.active-grid-row.active-list-item,
			.active-scroll-left .active-list-item {height: 22px;}
			
			.active-column-0 {width: 100px;}
			.active-column-1 {width: 100px;}
			.active-column-2 {width: 100px;}
			.active-column-3 {width: 120px;}
				
		</style>

		 <SCRIPT language='Javascript'>
 function forceSubmit()
		{
			setOpener();
			//var targetPage = window.dialogArguments.opener.document.all.item("frmMain").elements("EditPage").value;
					
			document.frm<%=DataSource%>Item.method='post';
			document.frm<%=DataSource%>Item='_self';
			document.frm<%=DataSource%>Item.submit();		
		}
		
		function setOpener()
		{
			window.parent.opener = window.parent.dialogArguments.opener;					
		}
</script>

</head>

<body Class="Dialog" onload="setOpener()" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" SCROLL="NO">


<form name = 'frm<%=DataSource%>Item' id='frmMain' method = 'post' action = 'BankListItem.asp'>

  
 <%
 Set conn = GetActiveConnection("KBroker")
 
	Dim rowCount
	Dim securityList
	Dim quote 
	
	quote =  chr(34)
	
	
             
    sqlStr = "SELECT * FROM BankList WHERE " & _
             "	Bank_DPA_ IN (SELECT Bank_DPA_ FROM BnkBranch WHERE " & _
             "	BnkBranch_DPA_= " & ID & ")"
       
    Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
    If Not(rs.EOF Or rs.BOF) Then%>
        
		<!-- grid data -->
		<% 'row data
		 
		rs.MoveFirst
		Do Until rs.EOF 

'======================= Begin_Alter_Across_Entities =================================
			'row ID 
			'rowData = rowData & ","
			rowData = rowData & quote & rs.Fields("BnkBranch_DPA_") & quote & " : " 
			
			'row data 
			rowData = rowData & "[" 			
			rowData = rowData & quote & rs.Fields("BnkBranchName") & quote & ","
			rowData = rowData & quote & rs.Fields("BnkBranchCode") & quote & ","
			rowData = rowData & quote & rs.Fields("BnkBranchSwiftCode") & quote & ","
			rowData = rowData & quote & " " & quote '& ","
			rowData = rowData & "]" 
			'build the row IDs array
			'rowIDs = rowIDs & "," 
			rowIDs = rowIDs & quote & rs.Fields("BnkBranch_DPA_") & quote 
			rowCount = rowCount + 1
			
			
'======================= End_Alter_Across_Entities =================================

			rs.MoveNext 
				'build the row IDs array
				rowIDs = rowIDs & "," 
				rowData = rowData & ","	
			
		Loop
		
	End if
'======================= Begin_Alter_Across_Entities =================================

	'row ID 
	rowData = rowData & quote & 0 & quote & " : " 
			
	'row data 
	rowData = rowData & "[" 
	'rowData = rowData & quote & "New Line" & quote & ","  
	rowData = rowData & quote & "<input type = 'text' name ='txtName' STYLE='WIDTH: 100px' id = 'txtName' size='9' OnClick='event.cancelBubble=true;'>" & quote & ","
	rowData = rowData & quote & "<input type = 'text' name ='txtBranchCode'  STYLE='WIDTH: 100px' id = 'txtBranchCode' size='9' OnClick='event.cancelBubble=true;'>" & quote & ","
	rowData = rowData & quote & "<input type = 'text' name ='txtBranchSwiftCode'  STYLE='WIDTH: 100px' id = 'txtBranchSwiftCode' size='9' OnClick='event.cancelBubble=true;'>" & quote & ","
	rowData = rowData & quote & "<input type = 'button' name ='cmdAddGridNew'   id = 'cmdAddGridNew' size='9' onClick = 'AddRowInProgress();;' value='Add' class=Buttons>&nbsp;&nbsp;<input type='reset' value='Cancel' Class=Buttons>" & quote 
	rowData = rowData & "]" 
	

	'build the row IDs array 
	rowIDs = rowIDs & quote & 0 & quote 
	rowCount = rowCount + 1
		


%> 


		<script language="javascript">
			//column titles 
			var colCount = 4;
			var colNames = ["Branch", "BranchCode", "SwiftCode", ""];
			
			var myColumns = ["Branch", "Branch Code", "Swift Code", ""];
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
			
			var RowEditFn = function(src)
			{
				var rowIndex = src.getProperty("row/index");
				var i;
				if (rowIndex==0) return;
				for(i = 0; i < colCount; i++)
				{
					if(colNames[i] != "")
					{
						if(prevRow >= 0)
						{
							myData[prevRow][i] = document.frmMain.elements(colNames[i]).value;
						}
						myData[rowIndex][i] = "<INPUT TYPE='text' NAME='" + colNames[i] + "' ID='" + colNames[i] + "' VALUE='" + myData[rowIndex][i] + "' onChange = 'EditInPlaceDataChanged();'>";
						
					}
				}
				
				myData[rowIndex][colCount - 1] = "<INPUT TYPE='button' class='Buttons' VALUE='Save' onClick = 'SaveInPlaceEdit();event.cancelBubble=true;'>&nbsp;<INPUT TYPE='button' class='Buttons' VALUE='Cancel' onClick = 'cancelEditRow();event.cancelBubble=true;'>";
				inPlaceEdit = true;
				prevRow = rowIndex;
				grid.refresh();
			}
			
			
			function cancelEditRow(){
					var i;
							for(i = 0; i < colCount; i++)
							{
								if(colNames[i] != "")
								{
									myData[prevRow][i] = document.frmMain.elements(colNames[i]).value;
								}
							}
							myData[prevRow][colCount - 1] = "";						
							inPlaceEdit = false;
							prevRow = -1;
							grid.refresh();
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
			
			function UpdateID(){
				headerID = window.parent.frames["header"].document.all.item("ID").value;
				document.all.item("ID").value = headerID;
			}
			
			function restoreID(){
				document.all.item("ID").value = headerID;
			}
			
			function DoDelete(){
				var selItem = document.all.item("itemID").value;
				if (selItem=="") {
					alert("Please select an item");
					return;
				}
				
				if (window.confirm("Delete selected item?")){
					document.all.item("action").value = "EXECUTE_DELETE";
					SaveInPlaceEdit();
					document.all.item("action").value = "Execute_Detail";
				}	
			}
			
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
			//alert(grid.toString());
		</script> 
       
 <table border="0" width="100%" ID="Table1">
<tr><td>
<input type = 'hidden' name ='ItemID' id = 'ItemID'>
<input type = 'hidden' name ='ID' id = 'ID' value="<%= ID %>">
<input type = 'hidden' name ='action' id = 'action' value="Execute_Detail">
</td>
</tr>
</table>
</form>
</body>

</html>
