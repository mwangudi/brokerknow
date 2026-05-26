<%OPTION EXPLICIT%>
<html>	

	<head>
		<title>Question 4</title>
		
		<style type="text/css">
			.deleteButton
			{
				border: 0;
				cursor: hand;
				font-family: Arial;
				font-size: 11px;
				font-weight: bold;
				color: blue;
				
			}
			.editableTextBox
			{
				width:100%;
				height:100%;
				border:0px;
			}	
				
			TABLE.dataTable
			{
				BACKGROUND: #ffffff;
						    
			}
			
			TABLE.dataTable TD.dataHeader, TD.dataCounter
			{
				background-color: #C0C0C0;
				font-family: Arial;
				font-size: 10pt;
				font-weight: bold;
			}
			
			TABLE.dataTable TD
			{
				WHITE-SPACE: nowrap;
				BORDER-BOTTOM: #C0C0C0 1px solid;
			    BORDER-TOP: #C0C0C0 0px solid;
			    BORDER-LEFT: #C0C0C0 1px solid;
			    BORDER-RIGHT: #C0C0C0 1px solid;	
				PADDING-RIGHT: 5px;			    
			    PADDING-LEFT: 5px;	
			 } 
		</style>
		
	</head>
		
	<body>
		<!--#include file="libraries/dbConn.inc"-->
		<%
			'variables declared here...
			Dim Conn 'ADODB connection object to db
			Dim Rs	'Recordset object
			Dim SQLStat	'sql statement
			Dim companyName	'company name	
			Dim companyTag	'company tag line
			Dim companyDesc	'company description	
			Dim companyId	'company id
			Dim displayCounter 'data counter
			Dim deleteFlag
			
			
			'open db using procedure found in dbConn.inc above
			Set Conn = OpenDatabase
				
			companyId = Trim(Request.Form("svCompanyId"))
			deleteFlag = Request.Form("deleteFlag")
			
			If companyId <> "" Then
			
				If deleteFlag = "1" Then
					SQLStat = "DELETE FROM Companies WHERE CompanyId = " & companyId
					Conn.Execute SQLStat
				
				Else
					'the save process is to be run			
					companyName = Trim(Request.Form(companyId & "cName")) 
					companyTag = Trim(Request.Form(companyId & "cTagLine"))
					companyDesc = Trim(Request.Form(companyId & "cDescription"))		
				
					'still, conduct a validation for the company name
					If companyName = "" Then%>
						<Script Language="JavaScript">
							window.parent.alert('The company name is required.')
						</Script>
					<%Response.End
					End If			
				
				
					Set Rs = Server.CreateObject("ADODB.Recordset")
					SQLStat = "SELECT * FROM Companies"
				
					'the enumerators can be used here because of the first
					'line inclusion of the ADODB library class found in the
					'included dbConn.inc
					Rs.CursorLocation = adUseClient
					Rs.Open SQLStat,  Conn.ConnectionString, adOpenDynamic, adLockBatchOptimistic
	
					With Rs
						If companyId = "-1" Then
							.AddNew 
						Else
							.Filter = "CompanyId = " & companyId 	
						End If
							
							.Fields("CompanyName").Value = companyName
							.Fields("CompanyTagLine").Value = companyTag
							.Fields("CompanyDescription").Value = companyDesc
							
						.UpdateBatch adAffectCurrent
					End With			
				End If
				
				Set Rs = Nothing
				Set Conn = Nothing
				
				'if a new addition has just happened, refresh the 
				'display table, and also when a deletion occurs
				If companyId = "-1" Or deleteFlag = "1" Then%>
					<Script Language="JavaScript">
						window.parent.self.location.reload();
					</Script>	
				<%End If		
						
				'stop execution
				Response.End  
			
			End If %>		
			
			
			<Script Language="JavaScript">
				function Delete(id)
				{					
					if (confirm('Are you sure you want to delete the selected company?')){
						document.getElementById('svCompanyId').value = id;
						document.getElementById('deleteFlag').value = 1;
						document.getElementById('frmMain').submit();						
					}				
				}
				
				function saveChanges(doc, enforce){					
					if (enforce)
					{
						if (doc.value=='') 
						{
							alert('The selected field is required.');
							doc.focus();
							return;
						}
					}				
					
					//update the current id of the data...
					document.getElementsByTagName('INPUT').item('svCompanyId').value = doc.id;					
					
					if (doc.id==-1)					
					{
						//new company data set being entered...
						//ensure company name has been entered..
						var coValue = document.getElementsByTagName('INPUT').item('-1cName').value
						if (coValue=='') return;
					}
					
					//a valid update has occured
					document.getElementById('frmMain').submit();
					
				}
				
				function doImport(fl)
				{	
					var selFl =  fl.value;					
					if (selFl.substring(selFl.length - 3, selFl.length).toUpperCase() !== 'XLS'){
						alert('The file selected is not a valid MS Excel file');
						fl.value = '';
						return;
					}
					document.getElementById('frmUpload').submit();
					//document.getElementById('uploadTbl').style.display = 'none';
				}			
			
				
			</Script>
			
			<DIV style="position:absolute;visibility:hidden;display:none">
				<iFrame src="" id="saveFrame" name="saveFrame"></iFrame>
			</DIV>
			
			<form method="post" action="manage_company.asp" id="frmMain" name="frmMain" target="saveFrame">
			<input type="hidden" name="svCompanyId" id="svCompanyId">
			<input type="hidden" name="deleteFlag" id="deleteFlag" value="0">
			<center>	
				<table border="0" cellspacing="0" cellpadding="3" class="dataTable">
					<tr>
						<td class="dataHeader">&nbsp;</td>
						<td class="dataHeader">Name</td>
						<td class="dataHeader">Tag Line</td>
						<td class="dataHeader">Description</td>
						<td class="dataHeader">&nbsp;</td>
					</tr>
					<%
					Set Rs = Conn.Execute("SELECT * FROM Companies")
					If Not (Rs.EOF Or Rs.BOF) Then
						displayCounter = 0
						Do Until Rs.EOF
							displayCounter = displayCounter + 1
							companyId = Rs.Fields("CompanyId").Value  %>
						<tr>
							<td class="dataCounter"><%= displayCounter %></td>
							<td><input type="text" id="<%= companyId %>" name="<%= companyId %>cName" OnChange="JavaScript: saveChanges(this, true)" value="<%= Rs.Fields("CompanyName").Value  %>" class="editableTextBox"></td>
							<td><input type="text" id="<%= companyId %>" name="<%= companyId %>cTagLine" OnChange="JavaScript: saveChanges(this, false)" value="<%= Rs.Fields("CompanyTagLine").Value  %>" class="editableTextBox"></td>
							<td><input type="text" id="<%= companyId %>" name="<%= companyId %>cDescription" OnChange="JavaScript: saveChanges(this, false)" value="<%= Rs.Fields("CompanyDescription").Value  %>" class="editableTextBox"></td>
							<td class="dataCounter"><input type="button" class="deleteButton" value=" X " title="Delete" OnClick="JavaScript: Delete('<%= companyId %>')"></td>
						</tr>
						<%Rs.MoveNext
						Loop
					End If%>
					<tr>
						<td class="dataCounter"><font color="navy" size="4">*</font></td>
						<td><input type="text" id="-1" OnChange="JavaScript: saveChanges(this, true)"  name="-1cName" value="" class="editableTextBox"></td>
						<td><input type="text" id="-1" OnChange="JavaScript: saveChanges(this, false)" name="-1cTagLine" value="" class="editableTextBox"></td>
						<td><input type="text" id="-1" OnChange="JavaScript: saveChanges(this, false)" name="-1cDescription" value="" class="editableTextBox"></td>
						<td class="dataCounter">&nbsp;</td>
					</tr>
				</table>
				
				
			
			</form>
			
			<form action="importXL.asp" enctype="multipart/form-data" method="post" name="frmUpload" id="frmUpload" target="saveFrame">
				<table style="border:1px lightblue ridge" id="uploadTbl">
					<tr>
						<td>
							 Pick a MS Excel file to upload.<br>
							 Ensure that the information in the Excel file is in the first sheet and is as follows:<br>
							 <i>Company Name ('A' column)<br>
							 Company Tag Line ('B' column)<br>
							 Company Description ('C' column)<br></i>
							<input type="file" onChange="JavaScript: doImport(this)" name="sourcefile" size="50">							
							
						</td>
						<td>
						</td>
					</tr>
				</table>					
			</form>
			
		</center>
	</body>
</html>
