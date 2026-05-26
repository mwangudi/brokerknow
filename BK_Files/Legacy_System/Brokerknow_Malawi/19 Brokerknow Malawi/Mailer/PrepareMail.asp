<html>
<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Mail Preparation</title>
<LINK REL="stylesheet" TYPE="text/css" HREF="../STYLE/default.css">
<link rel="stylesheet" type="text/css" href="../STYLE/webparts.css">
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
</head>

<BODY>
<!--#include file="../libroutines.asp"-->

<%

Set Conn = GetActiveConnection("KBroker")
mailSQL = "SELECT ClientName As Name, ClientEmail As Email, 'Client' As Description   FROM Client " & _
			"	UNION " & _
			"	SELECT AgentName As Name, AgentEmail As Email, 'Agent' As Description   FROM Agent"
Set Rs = Conn.Execute(mailSQL)


%>

<CENTER>
	<DIV class="ListNugget" id="AdvSearchHead" style="WIDTH: 640px" name="AdvSearchHead">
		<TABLE class="ListNuggetHeader" cellPadding="0" cellSpacing="0" width="100%" name="AdvSearchtestHeader"> 
			<TR>
			<TD class="ListNuggetTitleCellWhite"
					onselectstart="window.event.cancelBubble=true; return false;"   
					onclick="PartWrapperToggle('AdvSearchHead');">
					<A class=ListNuggetTitle onclick="return PartWrapperToggle('AdvSearchHead');"  
					 href="javascript:PartWrapperToggle('AdvSearchHead');">System Mail Addresses
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

<form method="POST" action="Preparemail.asp" name="form1">    

 <table align=center class="srch_bg" style="MARGIN-TOP: 0px" cellPadding="4" cellSpacing="0" border="0"> 
            <tr> 
              <td><b>Available Mail Addresses:</b>
              </td>                            
            </tr>
            <tr> 
              <td valign="top"> 
                <table border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td> 
                      <select STYLE="WIDTH: 200px" size="23" name="Users" multiple>
                        <%                        
                        If Not (rs.EOF Or Rs.BOF) Then
							Do Until Rs.EOF
								If Not IsNull(Rs.Fields("Email").Value) Then
									If Rs.Fields("Email").Value <> "" Then%>
										<option value="<%= Rs.Fields("Email").Value %>" ><%= Rs.Fields("Name").Value & " (" & Rs.Fields("Description").Value & ")"%></option>
							<%		End If
								End If
							Rs.MoveNext
							Loop				
						End If						
						%>
                      </select>
                    </td>
                  </tr>
                </table>
              </td>
              <td > 
			</td>
			</tr>
			<tr> 
			<td><input class="buttons" type="button" value="Send Mail to Selected Options" OnClick="JavaScript: DoSend()"></td>
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
	function DoSend(){
		var Object = document.all.item("Users");
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
			window.open("mailto: " + sendToList)
		}
		
	}
</Script>
</body>
</html>
