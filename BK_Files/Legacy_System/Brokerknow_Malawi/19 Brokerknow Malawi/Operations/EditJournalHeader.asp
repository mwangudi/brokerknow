<!--#include file="../libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "EditJournal"
	const DataEntity = "Journal"
	const DataEntityPlural = "Journals"
	const ActionFolder = "Operations"
	
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim ID
   Dim rsEdit
	
	action = ucase(Request.Form("action"))
	ID = Request("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		
                </script>
                <% response.end
        End If   
   
    action2 = ucase(Request.Form("delAction"))
    
	select case action2
		case "EXECUTE"
			Dim release
			dim ReleaseDate
		        
			release = Request.Form("Release")
			ReleaseDate = trim(Request.Form("ReleaseDate"))
				
			'validate release
			If Trim(release) = "" Then%>
					<script language = 'vbscript'>
			    			MsgBox "Invalid release status"
				        			
					</script>
					<% response.end
			End If
			
			Set conn = GetActiveConnection("KBroker")
		        
			'save data
			Dim userID
			Dim manualReleaseDate
				
			'manualReleaseDate = "GETDATE()"
			manualReleaseDate = "Date()"
				
			sqlStr = "UPDATE Journal SET Released = " & " " & release & " " & "" & _
					" WHERE Journal_DPA_  = " & ID                
'conn.Execute sqlStr 
			conn.BeginTrans
					conn.Execute SQLServerFormat(HandleQuote(sqlStr))
			conn.CommitTrans
			conn.Close
			Set conn = Nothing

			response.redirect "EditJournalHeader.asp?ID=" & ID
		end select

	select case action 
		case "EXECUTE_HEADER"
			Dim jDate
			Dim narrative		
						
			jDate = Request.Form("txtJournalDate")
			narrative = Request.Form("txtNarrative")       
			toCancel = Request.Form("cmdCancel")
									    
			
			If toCancel <> "" Then
				WriteDialogCloseScript
				Response.End
			End If   
			
				'validate size of Narrative
				If Len(Narrative) > 500 Then%>
				        <script language = 'vbscript'>
				        ShowMessage "Narrative can only be 500 characters in length"
						         
				        </script>
				        <% response.end
				End If

			    
				Set conn = GetActiveConnection("KBroker")
			    
				'save data
				sqlStr = "UPDATE [Journal] SET JournalDate = " & "#" & FormatDate(jDate) & "#" & ",JournalNarrative = " & "'" & narrative & "'" & "" & _
						" WHERE Journal_DPA_  = " & ID                

				'Response.Write sqlStr
				'Response.End		
				
				conn.BeginTrans
						conn.Execute SQLServerFormat(HandleQuote(sqlStr))
				conn.CommitTrans
				
				Set conn = Nothing			
				Set conn = GetActiveConnection("KBroker")
				'retrieve the item ID
				sqlStr = "SELECT JournalEntry_DPA_ FROM JournalEntry WHERE (Journal_DPA_= " & ID & ")"
								
				Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
				Set rs=conn.execute(sqlStr)
								
				conn.Close				
				
		Case Else
				'ID = GetOrderID(ID)	
    end select
    
    Function GetOrderID(detailID)
		Dim getRs
		Set getConn = GetActiveConnection("KBroker")
		sqlStr = "SELECT Journal_DPA_ FROM " & DataEntity & "List WHERE " & _
                "  JournalEntry_DPA_ =" & detailID
                
        set getRs = getConn.Execute(sqlStr)
        If Not (getRs.EOF OR getRs.BOF) Then
			GetOrderID = getRs("Journal_DPA_")
		Else
			GetOrderID = detailID
		End If	        
		
    End Function	
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
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>

<script language='vbscript'>
	function ItemSelected(itemID)
			frm<%=DataSource%>Header.elements("ID").value = itemID
	end function
					
	function SaveInPlaceEdit()
			frm<%=DataSource%>Header.submit
	end function
</script>

<script>
function DeleteItem()
{
		window.parent.frames("detail").HandleDeleteAction();
}

function ValidateItems()
{
		window.parent.frames("detail").HandleCloseAction();		
}

function CancelOperation()
{
		window.parent.frames("detail").HandleCancelOperation();
}
		function forceSubmit()
		{
			//setOpener();
			//var targetPage = window.dialogArguments.opener.document.all.item("frmMain").elements("EditPage").value;
					
			document.frm<%=DataSource%>Header.method='post';
			document.frm<%=DataSource%>Header.target='_self';
			document.frm<%=DataSource%>Header.submit();	
			
		}
		function setOpener()
		{
			
			window.parent.opener = window.parent.dialogArguments.opener;
		}
</script>		
</head>

<body Class="Dialog" onLoad="setOpener();">

<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<form name = 'frm<%=DataSource%>Header' id='frmMain' method = 'post' action = '<%=DataSource%>Header.asp' >
<%
        Set conn = GetActiveConnection("KBroker")
             
        sqlStr = "SELECT * FROM " & DataEntity & "FullList WHERE " & DataEntity & "_DPA_ = " & ID
                
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
    <td width="80"><%=DataEntity%> No.</td>
    <td width="416">&nbsp;<input readonly = 'true' class=readonly  type = 'text' name ='txt<%=DataEntity%>No' id = 'txt<%=DataEntity%>No' value = '<%=ID%>' size="20"></td>
  </tr>
  <tr>
    <td width="80">Date</td>
    <SCRIPT language="JavaScript">
	var calJournalDate=new ctlSpiffyCalendarBox("calJournalDate", "frm<%=DataSource%>Header", "txtJournalDate","cmdJournalDate","<%= FormatDate(rs.Fields("JournalDate")) %>",1);
	</SCRIPT>
    <td width="416">&nbsp;<SCRIPT language="JavaScript">calJournalDate.writeControl();</SCRIPT></td>
  </tr>
  <tr>
    <td width="80">Released</td>
    <%
     
     if rs.Fields("Released") = "" or isnull(rs.Fields("Released")) then Rlsed = 0 else Rlsed = rs.Fields("Released")
    
    If Rlsed = 1 Then%>
    <td width="416"><input type=checkbox class='BorderLess' style="border: 0;" checked value='False' name='chkRelease' onClick='UpdateReleaseStatus(this,"<%=rs.Fields("Journal_DPA_")%>");'></td>
    <%Else%>
    <td width="416"><input type=checkbox class='BorderLess' disabled style="border: 0;" id=checkbox1 name=checkbox1></td>
    <%End If%>
  </tr>
  <tr>
    <td width="80">Narrative</td>
    <td width="416">&nbsp;<input type = 'text' name ='txtNarrative' id = 'txtNarrative' size="20" value = '<%=rs.Fields("JournalNarrative")%>'></td>
  </tr>
  <tr>
  <tr>
    <td width="100%" colspan=4 align=center>
    <input type = 'button' Class=Buttons name ='cmdDelete' id = 'cmdDelete' value=" Delete " onClick='DeleteItem();'>
	 &nbsp;&nbsp;
	 <input type = 'button' Class=Buttons name ='cmdAdd' id = 'cmdSave' value=" Save " onclick="document.getElementById('buttonAction').value='Save';forceSubmit();" >
	 &nbsp;&nbsp;     
	 <input type = 'button' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Commit " onClick="document.getElementById('buttonAction').value='Commit';ValidateItems(); forceSubmit();">
	 &nbsp;&nbsp;
     <!--<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " onClick='CancelOperation();'>-->
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " onClick='JavaScript: window.parent.opener.location=window.parent.opener.location;window.self.close();'>
    	<input type = 'hidden' name ='action' id = 'action' value="Execute_Header">
    	<input type = 'hidden' name ='ID' id = 'ID' value="<%= Rs.Fields("Journal_DPA_").Value %>">
		<input type = 'hidden' name ='Release' id = 'Release'>
		<input type = 'hidden' name ='buttonAction' id = 'action' value="">
		<input type = 'hidden' name ='ReleaseDate' id = 'ReleaseDate'>
		<!--The message in the input below is meant for the Footer page. It is replaced with a different string if deletion proceeds-->
		<input type = 'hidden' name ='delAction' id = 'delAction' value="This action will delete the <%=DataEntity%> that contains the selected item. If you want to delete the whole <%=DataEntity%>, click Yes. Otherwise, click No and then click Edit to delete the item from the <%=DataEntity%>.">
		<!-- ----------------------------------------------------------------------------------------------------------------------- -->
		</td>
  </tr>
  </table>
  </form>
</body>

<script>
function  UpdateReleaseStatus(theChk, theItem)
{
	var relVal = "0";
	if (theChk.checked)
	{
		relVal = "1";
	}
				
	document.frmMain.elements("Release").value = relVal;
	document.frmMain.elements("ReleaseDate").value = theChk.ReleaseDate;
	document.frmMain.elements("delAction").value = "Execute";
	ItemSelected(theItem);
	SaveInPlaceEdit();
}
</script>
</html>
