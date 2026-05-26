<!--#include virtual="libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "EditEntity"
	const DataEntity = "Entity"
	const DataEntityPlural = "Entities"
	const ActionFolder = "Operations"
	
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim rsEdit
	Dim guid
	Dim guidStr
	
	action = ucase(Request.Form("action"))
	ID = Request("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		
                </script>
                <% response.end
        End If
        
	if action = "EXECUTE" then
		Dim buttonAction
		Dim reloadRequired
		
		reloadRequired = false
		buttonAction = Trim(Ucase(Request.Form("cmdAdd")))
		if buttonAction = "SAVE" then
				Dim EntType
				Dim EntName
				Dim gen1
				Dim gen2
				Dim gen3
		        Dim OpeningBal
		        Dim Code
       
				Code = Request.Form("txtCode")
				OpeningBal = Request.Form("txtOpeningBal")
				EntType = Request.Form("cboEntityType") 
				EntName = Request.Form("txtEntName")
				gen1 = Request.Form("cboGenericSetting1")
				gen2 = Request.Form("cboGenericSetting2")
				gen3 = Request.Form("cboGenericSetting3")			
				 
				'validate Account Type
				If Trim(EntType) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Entity Type"
				         		
				         </script>
				         <% response.end
				End If

				'validate Generic1
				If Trim(gen1) = "" Then
						gen1 = "NULL"
				End If
				'validate Generic2
				If Trim(gen2) = "" Then
						gen2 = "NULL"
				End If
				'validate Generic3
				If Trim(gen3) = "" Then
						gen3 = "NULL"
				End If
				'validate Name
				If Trim(EntName) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Name"
				         		
				         </script>
				         <% response.end
				End If
				'validate size of Name
				If Len(EntName) > 100 Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Name can only be 100 characters in length"
				         		
				         </script>
				         <% response.end
				End If
				'validate Code
				If Trim(Code) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Code"
					         		
				         </script>
				         <% response.end
				End If
				'validate size of Code
				If Len(Code) > 50 Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Code can only be 50 characters in length"
					         		
				         </script>
				         <% response.end
				End If
					
				If Not IsNumeric(OpeningBal) Then%>
						<script language = 'vbscript'>
						ShowMessage "Opening Balance can only be numeric"						
						</script>
						<% response.end
				End If

				'save data		
				Set conn = GetActiveConnection("KBroker")
				
				sqlStr = "UPDATE [Entity] SET " & _
						"		GenericSetting_DPA_ = " & " " & Gen1 & " " & _
						"		,GenericSetting_DPA_2 = " & " " & Gen2 & " " & "" & _
						"       ,GenericSetting_DPA_3 = " & " " & Gen3 & " " & ",EntityName = " & "'" & EntName & "'" & "" & _
						"       ,EntityType_DPA_ = " & " " & EntType & " " & _
						"       ,EntityOpeningBal = " & " " & OpeningBal & " " & _
						"       ,EntityCode = " & "'" & Code & "'" & _
						"		 WHERE Entity_DPA_  = " & ID
				sqlStr = SQLServerFormat(HandleQuote(sqlStr))


				conn.Execute sqlStr
				
				conn.Close
				Set conn = Nothing
				WritefraEnabledDialogCloseScript
				Response.End
			end if
			Dim clientCode
        
			clientCode = "var validNavigate = true;" & chr(13)
			%>
			<script>
				<%=clientCode%>
			</script>
			<%
			response.End
   	end If
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit <%=DataEntity%></title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<script language="JavaScript" src="CALENDAR/calendar.js"></script>
<!--END CALENDAR -->

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
		
		var gen1 = "";
		var gen2 = "";
		var gen3 = "";
				
		function FilterGenericList(theList)
		{
			var i = 0;
			var entType = theList.value;
			var genBagList = document.frmMain.cboGenericBag;
			

			var genList = document.frmMain.cboGenericSetting1;
			var genList2 = document.frmMain.cboGenericSetting2;
			var genList3 = document.frmMain.cboGenericSetting3;
			RemoveOptions(genList)  ;
			RemoveOptions(genList2)  ;
			RemoveOptions(genList3)  ;
					
			//add default selection (none)
			var NewOption = new Option();   			    
		   	NewOption.text = '';
		   	NewOption.value = '';
		   	NewOption.selected = true;
		   	genList.add(NewOption, 0);
		   	
		   	NewOption = new Option();   			    
		   	NewOption.text = '';
		   	NewOption.value = '';
		   	NewOption.selected = true;
		   	genList2.add(NewOption, 0);
		   	
		   	NewOption = new Option();   			    
		   	NewOption.text = '';
		   	NewOption.value = '';
		   	NewOption.selected = true;
		   	genList3.add(NewOption, 0);
					
					
			for (i=0; i < genBagList.options.length; i++) {
				if((genBagList.options(i).EntityTpe == entType))
				{
					var NewOption = new Option();   			    
		   		    NewOption.text = genBagList.options[i].text;
		   		    NewOption.value = genBagList.options[i].value;
		   		    
					if((genBagList.options(i).Generic == 1))
					{
		   				if(NewOption.value == gen1)
		   				{
		   					NewOption.selected = true;
		   				}
		   				genList.add(NewOption, 0);
		   			}
		   			
		   			if((genBagList.options(i).Generic == 2))
					{
		   				if(NewOption.value == gen2)
		   				{
		   					NewOption.selected = true;
		   				}
		   				genList2.add(NewOption, 0);
		   			}
		   			
		   			if((genBagList.options(i).Generic == 3))
					{
		   				if(NewOption.value == gen3)
		   				{
		   					NewOption.selected = true;
		   				}
		   				genList3.add(NewOption, 0);
		   			}
		   			
				}
						
			}
		}
		
		function RemoveOptions(Field){		   
		   if (Field.length==0) return;
		  
		   for (loop=Field.length - 1; loop >= 0; loop--) {
		       var GoneOption = Field.options[loop]		  
		       Field.remove(GoneOption.index);		        
		       }
		   
		 }
		
</script>
</head>
<%
Set conn = GetActiveConnection("KBroker")
sqlStr = "SELECT * FROM Entity WHERE SystemMaintained = 0 AND Entity_DPA_=" & ID
        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		window.self.ShowMessage "The selected <%=DataEntity%> cannot be retrieved for editing"
                		
                </script>
                <% response.end
        End If
%>
<body Class="Dialog">

<form name = 'frm<%=DataSource%>' method = 'post' action = '<%=DataSource%>.asp' id = "frmMain">
<table border="0" width="100%" cellspacing="1" cellpadding="1">
 
  <tr>
    <td width="15%">Type</td>
    <td width="35%"><select name = 'cboEntityType' id = 'cboEntityType' size="1" onchange="FilterGenericList(this)">
<%		
        sqlStr = "SELECT * FROM [EntityTypeList] Order By EntityTypeName"
        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsEdit.EOF Or rsEdit.BOF) Then
                rsEdit.MoveFirst
                Do Until rsEdit.EOF
                		if rsEdit.Fields("EntityType_DPA_") = rs.Fields("EntityType_DPA_") Then%>
                			<option selected value = '<%=rsEdit.Fields("EntityType_DPA_")%>'><%=rsEdit.Fields("EntityTypeName")%></option>
                		<%else%>
                        <option value = '<%=rsEdit.Fields("EntityType_DPA_")%>'><%=rsEdit.Fields("EntityTypeName")%></option>
                     <%end if
						rsEdit.MoveNext
                Loop
        End If
%>

    </select></td>
    <td width="31%">

	</td>
  </tr>
  <tr>
    <td width="15%">Name</td>
    <td width="54%"><input type = 'text' name ='txtEntName' id = 'txtEntName' size="20" value = '<%=rs.Fields("EntityName")%>'></td>
    <td width="31%">

	</td>
  </tr>
   <tr>
    <td width="15%">Code</td>
    <td width="54%"><input type = 'text' name ='txtCode' id = 'txtCode' size="20" value = '<%=rs.Fields("EntityCode")%>'></td>
    <td width="31%">

	</td>
  </tr>
  <tr>
    <td nowrap> Opening Balance</td>
    <td nowrap><input type = 'text' name ='txtOpeningBal' STYLE="TEXT-ALIGN: RIGHT;" id = "txtOpeningBal" size="20" value = '<%=rs.Fields("EntityOpeningBal")%>'></td>
  </tr>
  <tr style="display:none">
    <td width="15%"></td>
    <td width="54%"><select name = 'cboGenericBag' id = 'cboGenericBag' size="1">
<%
        Dim genericRS
        sqlStr = "SELECT * FROM [GenericSettingList] Order By GenericSettingDescription"
        Set genericRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        'genericRS.Filter = "Generic_DPA_ = 1"
        If Not (genericRS.EOF Or genericRS.BOF) Then
                Do Until genericRS.EOF%>
						<option Generic = '<%=genericRS.Fields("Generic_DPA_")%>' EntityTpe = '<%=genericRS.Fields("EntityType_DPA_")%>' value = '<%=genericRS.Fields("GenericSetting_DPA_")%>'><%=genericRS.Fields("GenericSettingDescription")%></option>
                        <%genericRS.MoveNext
                Loop
        End If
%>

    </select></td>
    <td width="31%">

	</td>
  </tr>
  <tr>
    <td width="15%">Generic 1</td>
    <td width="54%"><select name = 'cboGenericSetting1' id = 'cboGenericSetting1' size="1"></select></td>
    <td width="31%">

	</td>
  </tr>
  <tr>
    <td width="15%">Generic 2</td>
    <td width="54%"><select name = 'cboGenericSetting2' id = 'cboGenericSetting2' size="1"></select></td>
    <td width="31%">

	</td>
  </tr>
  <tr>
    <td width="15%">Generic 3</td>
    <td width="54%"><select name = 'cboGenericSetting3' id = 'cboGenericSetting3' size="1"></select></td>
    <td width="31%">

	</td>
  </tr>
  <tr>
    <td nowrap> Quarter 1</td>
    <td nowrap><input type = 'text' name ='txtQuarter1' STYLE="TEXT-ALIGN: RIGHT;" id = "txtQuarter1" size="20" value='<%=rs.Fields("Quarter1")%>'></td>
  </tr>
  <tr>
    <td nowrap> Quarter 2</td>
    <td nowrap><input type = 'text' name ='txtQuarter2' STYLE="TEXT-ALIGN: RIGHT;" id = "txtQuarter2" size="20" value='<%=rs.Fields("Quarter2")%>'></td>
  </tr>
  <tr>
    <td nowrap> Quarter 3</td>
    <td nowrap><input type = 'text' name ='txtQuarter3' STYLE="TEXT-ALIGN: RIGHT;" id = "txtQuarter3" size="20" value='<%=rs.Fields("Quarter3")%>'></td>
  </tr>
  <tr>
    <td nowrap> Quarter 4</td>
    <td nowrap><input type = 'text' name ='txtQuarter4' STYLE="TEXT-ALIGN: RIGHT;" id = "txtQuarter4" size="20" value='<%=rs.Fields("Quarter4")%>'></td>
  </tr>  
  <script language="javascript">
		gen1 = <%=Rs.Fields("GenericSetting_DPA_").Value%>
		gen2 = <%=Rs.Fields("GenericSetting_DPA_2").Value%>
		gen3 = <%=Rs.Fields("GenericSetting_DPA_3").Value%>
		FilterGenericList(document.frmMain.cboEntityType);  
  
  </script>
  <tr>
	  <td width="100%" colspan=3 align="right" valign=absBottom>
		<BR><BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save" onclick = "AllowedNavigation()">
    	<input type = 'button' Class=Buttons name ='cmdCancel' id = "cmdCancel" value="Cancel" onclick = "JavaScript: window.self.close()">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ID' id = 'ID' value="<%= Rs.Fields("Entity_DPA_").Value %>">
	</td>
  </tr>
</table>

</form>
</body>

</html>
