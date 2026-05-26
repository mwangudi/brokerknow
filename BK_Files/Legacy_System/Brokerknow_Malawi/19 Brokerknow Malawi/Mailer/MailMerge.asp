<html>
<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Mail Merge Template</title>
<LINK REL="stylesheet" TYPE="text/css" HREF="../STYLE/default.css">
<link rel="stylesheet" type="text/css" href="../STYLE/webparts.css">
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
</head>

<BODY>
<!--#include file="../libroutines.asp"-->

<%

	Set Conn = GetActiveConnection("KBroker")
	mailSQL = "SELECT * FROM MailMerge ORDER BY MergeTable"
	Set Rs = Conn.Execute(mailSQL)
	
	MailMergeOptions = Request.Form("MailMergeOptions")
	
	If MailMergeOptions =  "1" Then
		If Not (Rs.EOF Or Rs.BOF) Then
			Do Until Rs.EOF
				thisMergeTable = Rs.Fields("MergeTable").Value
				theSelectedValues = Request.Form(thisMergeTable) 
				
				If theSelectedValues <> "" Then
					'selected items for mail merge
					'get the mail merge items
					theSelectedMailMergeItems = Request.Form(thisMergeTable & "Details")
					If  theSelectedMailMergeItems <> "" Then
						Dim selIDs
						selIDs = Split(theSelectedValues, ",")
						For i = 0 To UBound(selIDs)
							i =i
						Next
						
					End If
					
				End If		
				
				Rs.MoveNext
			Loop			
		End If
		
		Set Rs = Nothing
		Set Conn = Nothing
		Response.End 
	End If


	


%>

<CENTER>
	<DIV class="ListNugget" id="AdvSearchHead" style="WIDTH: 640px" name="AdvSearchHead">
		<TABLE class="ListNuggetHeader" cellPadding="0" cellSpacing="0" width="100%" name="AdvSearchtestHeader"> 
			<TR>
			<TD class="ListNuggetTitleCellWhite"
					onselectstart="window.event.cancelBubble=true; return false;"   
					onclick="PartWrapperToggle('AdvSearchHead');">
					<A class=ListNuggetTitle onclick="return PartWrapperToggle('AdvSearchHead');"  
					 href="javascript:PartWrapperToggle('AdvSearchHead');">System Mail Merge Options
					</A>
				</TD>
			 
				<TD class=ListNuggetButtonCellWhite onclick="PartWrapperToggle('AdvSearchHead');">
				<DIV class=ListNuggetButton>
					<IMG class=ListNuggetUpButton id=AdvSearchUp height=17 alt="Hide options" src="../images/blue-chevron_up.gif" width=17 align=right border=0 name=AdvSearchHeadUp>
					<IMG class=ListNuggetDownButton id=AdvSearchDown height=17 alt=Options src="../images/gray-chevron_down.gif" width=17 align=right border=0 name=AdvSearchHeadDown>
				</DIV>
			</TD>
			</TR>
		</TABLE>
		
<DIV class="ListNuggetBody" id="AdvSearchHeadBody" name="AdvSearchHeadBody" style="WIDTH: 640px">
<table class="srch_bg" style="MARGIN-TOP: 0px" cellPadding="1" width=100% cellSpacing="0" border="0">  
<tr><td>

<form method="POST" action="MailMerge.asp" name="frmMailMerge">    
	<input type=hidden value="MailMergeOptions" value=1>
 <table align=center class="srch_bg" style="MARGIN-TOP: 0px" cellPadding="4" cellSpacing="0" border="0"> 
            <tr> 
              <td><b>Available Mail Merge Options:</b>
              </td>                            
            </tr>
            <%
            Dim myCount            
            myCount = 0
            
            If Not (Rs.EOF Or Rs.BOF) Then
				Do Until Rs.EOF
					myCount = myCount + 1%>
            <tr> 
              <td valign="top">
              
              	<DIV class="ListNugget" id="AdvSearchHead<%= myCount %>" style="WIDTH: 640px" name="AdvSearchHead<%= myCount %>">
					<TABLE class="ListNuggetHeader" cellPadding="0" cellSpacing="0" width="100%" name="AdvSearchtestHeader"> 
					<TR>
					<TD class="ListNuggetTitleCellWhite"
					onselectstart="window.event.cancelBubble=true; return false;"   
					onclick="PartWrapperToggle('AdvSearchHead<%= myCount %>');">
					<A class=ListNuggetTitle onclick="return PartWrapperToggle('AdvSearchHead<%= myCount %>');"  
					 href="javascript:PartWrapperToggle('AdvSearchHead<%= myCount %>');"><%= Rs.Fields("Description") %>
					</A>
				</TD>                
					<TD class=ListNuggetButtonCellWhite onclick="PartWrapperToggle('AdvSearchHead<%= myCount %>');">
					<DIV class=ListNuggetButton>
					<IMG class=ListNuggetUpButton id=AdvSearchUp height=17 alt="Hide options" src="../images/blue-chevron_up.gif" width=17 align=right border=0 name=AdvSearchHeadUp>
					<IMG class=ListNuggetDownButton id=AdvSearchDown height=17 alt=Options src="../images/gray-chevron_down.gif" width=17 align=right border=0 name=AdvSearchHeadDown>
					</DIV>
					</TD>
					</TR>
					</TABLE>
					<DIV class="ListNuggetBody" id="AdvSearchHeadBody" name="AdvSearchHeadBody" style="WIDTH: 640px">
					<table border="0" cellspacing="0" cellpadding="0">
					<tr> 
						<td>Merge Details</td>
						<td>&nbsp;&nbsp;Merge Fields</td>
					</tr>
					<tr>
				<%
				Set myInnerRs = Conn.Execute ("SELECT * FROM " & Rs.Fields("MergeTable").Value)
				myIDField = Rs.Fields("IDField").Value
				myDisplayField = Rs.Fields("DisplayField").Value
				
				If Not (myInnerRs.EOF Or myInnerRs.BOF) Then%>                                             						
					
                    <td> 
                      <select STYLE="WIDTH: 200px" size="5" name="<%= Rs.Fields("MergeTable").Value %>" multiple>                        
						<%Do Until myInnerRs.EOF %>
							<option value="<%= myInnerRs.Fields(myIDField).Value %>"><%= myInnerRs.Fields(myDisplayField	).Value %></option>
						<%	
							myInnerRs.MoveNext
						  Loop										
						%>
                      </select>
                    </td>                  
				<%
				End If				
				
				Set myInnerRs = Conn.Execute ("SELECT * FROM MailMergeDetails WHERE MailMergeID LIKE '" & Rs.Fields("ID").Value & "'")				
				
				If Not (myInnerRs.EOF Or myInnerRs.BOF) Then%>                                                             
                    <td> &nbsp;
                      <select STYLE="WIDTH: 200px" size="5" name="<%= Rs.Fields("MergeTable").Value %>Details" multiple>                        
						<%Do Until myInnerRs.EOF %>
							<option value="<%= myInnerRs.Fields("ID").Value %>" ><%= myInnerRs.Fields("Description").Value %></option>
						<%	
							myInnerRs.MoveNext
						  Loop										
						%>
                      </select>
                    </td>                 
                <%End If
                Set myInnerRs = Nothing%>
                 </tr>
                </table>
                </DIV>
              </td>
              <td > 
			</td>
			</tr>
			
			<%Rs.MoveNext
			Loop
			End If
			%>
			
			<tr> 
			<td><input type="button" value="Mail Merge" OnClick="JavaScript: DoMailMerge()" id=button1 name=button1></td>
			</td>			
			</td>
			</tr>
			</table>
			</td>
			</tr>
			</table>


</form>


</td>
</tr>
</table>
</div>
</div>

</CENTER>			

<%Set Rs = Nothing
Set Conn = Nothing%>
<Script Language="JavaScript">
	function DoMailMerge(){
		var Object = document.all.item("MailMergeOptions");
		var sendToList;
		
		sendToList = "";
		for (loop=Object.length-1; loop>-1; loop-- ){
			if (Object.options[loop].selected == true){
				if (sendToList=="") sendToList = Object.options[loop].value
				else sendToList = sendToList + ";" + Object.options[loop].value
			}
		}
		
		if (sendToList=="") (alert("There are no selected options"))
		else{
			try{
				document.all.item("frmMailMerge").submit(); 
			}
			catch(e){}
		}
		
	}
</Script>
</body>
</html>
