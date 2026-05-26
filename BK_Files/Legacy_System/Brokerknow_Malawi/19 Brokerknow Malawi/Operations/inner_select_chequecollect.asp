<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Cheque Collection</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
</head>

<body Class="Dialog" topMargin=0 leftMargin=0 rightMargin=0 bottomMargin=0>
<!--#include file="../libroutines.asp"-->

  <table border="0" width="100%" cellspacing="0" cellpadding="0">
    <tr>
    	
	  <!-- Your main data goes here (centrally located on the page) -->
          	
        <TD valign=top align=left width="100%" height="1" class="layeredContent">
		<table width="100%" cellspacing="0" cellpadding="0" border="0" bgColor="#FFFFFF">
        			<%
        						Dim Conn	'db connection object
        						Dim Rs	'recordset object of previous articles
        						Dim prevArtSQL	'sql statement to retrive previous articles
        						Dim articleCount	'counter
        						Dim idStr	'id for the info holder
        						Dim filterClause	'filter info sent in querystring
        						
        						'open db
        						Set Conn = GetActiveConnection("KBroker")    						
        						
        						'check for previous articles
        						'stop
        					filterClause = Request.QueryString
        				
        					if filterClause <> "" then
        						if(trim(Request.QueryString("action"))<>"") then
								prevArtSQL = "SELECT Payment.* FROM Payment WHERE (ChequeCollection = N'COLLECTED') AND (EntityType_DPA_ = 1) AND (Entity_DPA_ = " & Request.QueryString("ent_type") & ")"        						
        						else        						
        						prevArtSQL = "SELECT Payment.* FROM Payment WHERE (ChequeCollection = N'AWAITING COLLECTION') AND (EntityType_DPA_ = 1) AND (Entity_DPA_ = " & Request.QueryString("ent_type") & ")"
        						end if
        						
        					
								Set Rs = Server.CreateObject("ADODB.Recordset")
								Rs.Open prevArtSQL, Conn.ConnectionString, adOpenKeySet, adLockOptimistic
	
        						If Not (Rs.EOF Or Rs.BOF) Then        							%>
        						<tr bgColor="#C0C0C0">
        							<td align="top" nowrap></td>
        							<td align="top" nowrap><b>Cheque No</b></td>
								</tr>
        						<%
        							articleCount = 0
        							Do Until Rs.EOF        								
        									idStr = "input-" & articleCount
        									if(trim(Request.QueryString("action"))<>"") then
        									%>
        									<tr>
        										<td align="top" nowrap><input type="checkbox" class="BorderLess" checked disabled style="background-color: transparent" name="chk<%=trim(Rs.Fields("Payment_DPA_").Value)%>" ID="chk<%=trim(Rs.Fields("Payment_DPA_").Value)%>" OnClick="JavaScript: updateOwnerDocument(<%=Rs.Fields("Payment_DPA_").Value%>)"></td>
        										<td align="top" nowrap><%= Rs.Fields("PaymentReference").Value   %></td>
											</tr>
        									<%        									
        									else
        									%>
        									<tr>
        										<td align="top" nowrap><input type="checkbox" class="BorderLess" style="background-color: transparent" name="chk<%=trim(Rs.Fields("Payment_DPA_").Value)%>" ID="chk<%=trim(Rs.Fields("Payment_DPA_").Value)%>" OnClick="JavaScript: updateOwnerDocument(<%=Rs.Fields("Payment_DPA_").Value%>)"></td>
        										<td align="top" nowrap><%= Rs.Fields("PaymentReference").Value   %></td>
											</tr>
        									<%
        									end if
        								Rs.MoveNext
        							Loop        							
        						End If
        						
        						'clean up
        						Set Rs = Nothing
        						Set Conn = Nothing
        					end if
        						%>
		</table>				   
       
       </TD>       
       <!-- End page's main data -->	      
    </tr>
  
  </table> 	

<Script Language="JavaScript">
	function updateOwnerDocument(theID){
		var chkBox = document.all.item("chk" + theID)
		if (chkBox.checked==true)
		{
			window.parent.UpdateVoucherAmount(theID,"add");
		}
		else
		{
			window.parent.UpdateVoucherAmount(theID,"remove");
		}
	}
</Script>

<Script Language="VBScript">
	Function UpdateSelContracts(updateVals)	
		Dim isChecked, returnVal
		
		On Error Resume Next
		
		For Each Thing In document.getElementsByTagName("INPUT") 
			isChecked = Thing.Checked
			If  isChecked = True Then
				If returnVal <> "" Then
					returnVal = returnVal & "," & Thing.Name
				Else
					returnVal = Thing.Name
				End If			
				
				If updateVals = True Then
					window.parent.UpdateVoucherAmount Thing.value
				End If
			End If
		Next 
		
		
		window.parent.document.getElementById("ContractsSel").value = returnVal
		
	End Function	

</Script>


</body>

</html>
