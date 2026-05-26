<!--#include file="../libroutines.asp"-->
<%
	
	'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "EditLevyOrder"
		const DataEntity = "LevyOrder"
		const DataEntityPlural = "LevyOrders"
		const ActionFolder = "Reports"
'======================= End_Alter_Across_Entities =================================
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim guidStr 
   Dim guid 
	
	action = ucase(Request.Form("action"))
	
	if action = "EXECUTE" then
		Dim LevyOrdersSel
		Dim VDate
		Dim LevyOrdersSelArray, existingLevyArray()
		Dim filterArray
		
		LevyOrdersSel = Request.Form ("LevyOrdersSel")		
		VDate = Request.Form ("txtVDate")
		
		If (Trim(LevyOrdersSel) = "") Then%>
				<script language = 'vbscript'>
                		ShowMessage "No levies specified"
				</script>
				<%response.end
		End If
		
		LevyOrdersSelArray = Split(LevyOrdersSel, ",")
		ReDim  existingLevyArray(UBound(LevyOrdersSelArray))

		
		Set conn = GetActiveConnection("KBroker")
		Set chkRs = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM LevyReportOrder"	
		chkRs.Open SQL, Conn.ConnectionString, adOpenKeySet, adLockOptimistic
		
		For i = 0 To UBound(LevyOrdersSelArray)
			thisVal = Trim(LevyOrdersSelArray(i))
			chkRs.Filter = "LevyName = '" & Replace(thisVal, Chr(39), "''")  & "'"
			If Not (chkRs.EOF Or chkRs.BOF) Then
				existingLevyArray(i) = chkRs.Fields("LevyOrder_DPA_").Value 
			Else
				existingLevyArray(i) = "NON"
			End If
			chkRs.Cancel
		Next
		
		Set chkRs = Nothing
		
	
		conn.BeginTrans
		
		For i = 0 To UBound(LevyOrdersSelArray)
			 	thisVal = Trim(LevyOrdersSelArray(i))
			 	existingLevy = IsNumeric(existingLevyArray(i))
				If existingLevy Then
					sqlStr = "UPDATE LevyReportOrder SET LevyOrder = " & i & " WHERE LevyOrder_DPA_ = " & existingLevyArray(i)
				Else
					 sqlStr = "INSERT INTO [LevyReportOrder] (LevyOrder" & _
								",LevyName,LevyOrder_DPA_) SELECT " & " " & i & " " & " as LevyOrder" & _
								",'" & thisVal & "' " & " as LevyName" & _
								"," & " " & "iif(isnull(max([LevyOrder_DPA_])),1,max([LevyOrder_DPA_]) + 1)" & " " & " as LevyOrder_DPA_" & _
								" FROM [LevyReportOrder]"
				End If
				
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
		Next
		
		
		conn.CommitTrans
        conn.Close
        
        Set conn = Nothing
						
		WritefraEnabledDialogCloseScript2
		Response.End
	
	end if
	


    Set conn = GetActiveConnection("KBroker")
    
    SQL = "SELECT * FROM LevyOrderListSource"
	Set contractRs = Server.CreateObject("ADODB.Recordset")
	contractRs.Open SQL, Conn.ConnectionString, adOpenKeySet, adLockOptimistic
%>

<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add <%=DataEntity%></title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
<SCRIPT language="JavaScript">
	var calVDate=new ctlSpiffyCalendarBox("calVDate", "frm<%=DataSource%>", "txtVDate","cmdVDate","<%= FormatDate(Date) %>",1);
</SCRIPT>
<!--END CALENDAR -->

<script >
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

		function forceSubmit()
		{
			setOpener();
			//var targetPage = window.dialogArguments.opener.document.all.item("frmMain").elements("EditPage").value;
					
			document.frm<%=DataSource%>.method='post';
			document.frm<%=DataSource%>.target='_self';
			document.frm<%=DataSource%>.submit();		
		}
		
		function setOpener()
		{
			window.self.opener = window.dialogArguments.opener;					
		}
</script>
</head>

<body Class="Dialog" onload="setOpener()">
<div id="spiffycalendar" STYLE="z-index: 10" class="text"></div>

<form name = 'frm<%=DataSource%>' method = 'post' action = '<%=DataSource%>.asp' target = 'deleteFrame' id='frmMain' OnSubmit="JavaScript: UpdateDialogHandle();">

<table border="0" width="100%" height="204">
  <tr>
    <td width="17%" height="10"></td>
    <td width="83%" height="10"></td>
  </tr>
  
  <td colspan = '2' height="106">
  
  <table border="0" cellspacing="1" cellpadding="2" id="mainTable">
	<tr>
	<td><b><font size="2" face="Verdana" color="#0000FF">Available Levies:&nbsp;</font></b></td>
	<td colspan="2" rowspan="2" valign="Top">
 
			<table border="0" width="100%" cellspacing="1" CellPadding="1">
			<tr>
			<td width="200%" colspan="2"><b><font size="2" face="Verdana" color="#0000FF">Selected
              Levies:</font></b></td>
			</tr>
			<tr>
			<td width="100%"><input type="button" style="display: none" value=" &lt; " name="MoveBtn" Class=Buttons OnClick="JavaScript: Move(this)">&nbsp;&nbsp;&nbsp;<input type="button" value=" &gt; " name="MoveBtn" Class=Buttons OnClick="JavaScript: Move(this); "></td>
			<td width="100%">&nbsp;&nbsp;
			<select size="4" name="LevyOrdersSel" multiple  OnKeyPress="JavaScript: if (event.keyCode==46) Move(this)" id="LevyOrdersSel">
			<%
			
			SQL = "SELECT * FROM LevyOrderList ORDER BY LevyOrder"
			Set levOrdrs = Conn.Execute(SQL)
			
			If Not (levOrdrs.EOF Or levOrdrs.BOF) Then
				Do Until levOrdrs.EOF
					displayValue = levOrdrs("LevyName")
					If StrComp(levOrdrs.Fields("LevyOrder_DPA_").Value, Request("ID"), vbTextCompare) = 0 Then
						selText = "selected"
					Else
						selText = ""
					End If	
					  %>
					<Option <%= selText %>  Value="<%= displayValue  %>"><%= displayValue %></Option>
					<%
					levOrdrs.MoveNext
				Loop
			End If
			
			
				
		%>
			</select></td>
			</tr>
			</table>               
			
	</td>
	</tr>
	<tr valign="Top">
	
	<td nowrap ID="userTD">
		<select size="4" name="ContractsAvail" multiple id="ContractsAvail">
		<option ContractAmount=0>Select a levy to add</option>
		<%
		If Not (contractRs.BOF Or contractRs.EOF) Then
		Do Until contractRs.EOF
			displayValue = contractRs("LevyName")
			%>
			<Option Value="<%=  displayValue  %>"><%= displayValue %></Option>
			<%
			contractRs.MoveNext
		Loop
		End If%>
		</select>
		<%
	
	Set contractRs = Nothing
	Set Conn = Nothing %></td>
	        
  </tr>
  
  
</table>

  
  </td>
  
  <tr>
	<td colspan="2" align="right">
		<input type="button" class="buttons" value="Move Down" OnClick="JavaScript: moveDown();"> &nbsp;&nbsp;<input type="button" class="buttons" value="Move Up" OnClick="JavaScript: moveUp();">
	</td>
  </tr>
  
  <tr>
    <td width="100%" colspan=2 align=right height="55">
		<input type = 'button' class=buttons name ='cmdAdd' id = 'cmdAdd' value=" Save " OnClick="forceSubmit()" OnClick="VBScript: SelectForm">
        &nbsp; <input type = 'button' class=buttons name ='cmdClose' id = "cmdClose" value=" Close " onclick = "window.self.close();">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
	</td>
  </tr>
</table>

</form>
<!--JAVA CODE-->
<Script Language="JavaScript">
<!--Begin
	document.all.item("ContractsAvail").size = Math.round(document.all.item("userTD").clientHeight / 10) ;
	document.all.item("LevyOrdersSel").size = document.all.item("ContractsAvail").size;
	
//===========BEGIN MOVE FUNCTION FOR SELECTED STAFF============================= 
function Move(Btn){

var todo = Btn.value;
var Users = document.all.item("ContractsAvail");
var loop;
var InsertList;

if (todo.search(">")>0){
 InsertList = document.all.item("LevyOrdersSel") ;	
 if (Users.selectedIndex==-1) return(ShowMessage("Select a levy from the contract list."))  
  AddOption(Users,InsertList) 
  }
else{
 RemoveOption(document.all.item("LevyOrdersSel"))
 }

	 
}



//=========END FUNCTION ====================================================

//=========BEGIN DROP-DOWN SELECT FUNCTION FOR FORM POSTING======

 function SelectAll(Object){
 //select all upwards
  for (loop=Object.length-1; loop>-1;
   loop--)
    {
     Object.options[loop].selected = true
    }
  }

//============END SELECT FUNCTION===================================

//==========BEGIN REMOVE OPTION/S FROM DROP-DOWN FUNCTION ON THE FLY=====
 function RemoveOption(Field){
	Selection = new Boolean();
	if (Field.length==0) return(ShowMessage("The list is empty."))
	for (loop=Field.length - 1; loop >= 0; loop--) {
	    var GoneOption = Field.options[loop]
		if (GoneOption.selected==true) {
	      		Selection = true;
	      		Field.remove(GoneOption.index);
	      }
	    }
	    
   if (Selection==false) ShowMessage("Select a levy to remove from the List.")
   
  }

//==============END REMOVE OPTION/S FUNCTION====================

//=========BEGIN ADD OPTION TO DROP-DOWN ON THE FLY=============

  function AddOption(Input,Output){    
    for (loop=0; loop < Input.length; loop++){
    		if (Input.options[loop].selected && loop !== 0){
    		    NewOption = new Option();   			    
			    NewOption.text = Input.options[loop].text;
			    NewOption.value = Input.options[loop].value;
			    NewOption.selected = false;	
			    if (!CheckDuplicates(Output, NewOption.value)) Output.add(NewOption, 0)
    		}
    		
    }
    
  }

function CheckDuplicates(DupPut, valText){
	var loop;
   for (loop=0; loop < DupPut.length;loop++){
      if (DupPut.options[loop].value==valText){
	       	return(true) ;
       }
     }
 }   
   
  
  
//========END ON THE FLY ADD OPTION FUNCTION=====================


function moveDown(){
	var doc = document.all.item('LevyOrdersSel');
	var currIndex = doc.selectedIndex;
	if (currIndex >= 0){
		try{
			var nextOptionIndex = currIndex - 1 + 2;
			var tempOptionText = doc.options[nextOptionIndex].text;
			var tempOptionValue = doc.options[nextOptionIndex].value;
			doc.options[nextOptionIndex].text = doc.options[currIndex].text;
			doc.options[nextOptionIndex].value = doc.options[currIndex].value;
			doc.options[currIndex].text = tempOptionText;
			doc.options[currIndex].value = tempOptionValue;
			doc.selectedIndex = nextOptionIndex;
		}
		catch(e){}
	}

}

function moveUp(){
	var doc = document.all.item('LevyOrdersSel');
	var currIndex = doc.selectedIndex;
	if (currIndex >= 0){
		try{
			var nextOptionIndex = currIndex - 1;
			var tempOptionText = doc.options[nextOptionIndex].text;
			var tempOptionValue = doc.options[nextOptionIndex].value;
			doc.options[nextOptionIndex].text = doc.options[currIndex].text;
			doc.options[nextOptionIndex].value = doc.options[currIndex].value;
			doc.options[currIndex].text = tempOptionText;
			doc.options[currIndex].value = tempOptionValue;
			doc.selectedIndex = nextOptionIndex;
		}
		catch(e){}
	}

}

-->
</Script>
 <!--END CODE-->
<Script Language="VBScript">
	Function FormatNum(num)		
		FormatNum = FormatNumber(num, 2)
	End Function
	
	Function SelectForm
		For Each Thing In frm<%=DataSource%>
			If InStr(1, Thing.Name, "LevyOrdersSel") > 0 Then
				SelectAll Thing
			End If
		Next
	End Function
	
	Function DoResizeWin	
		Dim nHeight, nWidth
		Dim defMaxHeight, defMaxWidth
		
		defMaxHeight = (screen.availHeight) - 100
		defMaxWidth = (screen.availWidth) - 100
		
		nHeight = (document.all.item("mainTable").clientHeight) + 50
		nWidth = (document.all.item("mainTable").clientWidth) + 50
		
		If nHeight > defMaxHeight Then
			nHeight = defMaxHeight
		End If
		
		If nWidth > defMaxWidth Then	
			nWidth = defMaxWidth
		End If		
				
		window.parent.dialogHeight = nHeight & "px"
		window.parent.dialogWidth = nWidth & "px"
		
		
	End Function
</Script>

</body>

</html>
