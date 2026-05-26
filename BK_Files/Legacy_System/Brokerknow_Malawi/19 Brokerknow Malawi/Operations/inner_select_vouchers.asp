<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Order</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
</head>

<body Class="Dialog" topMargin=0 leftMargin=0 rightMargin=0 bottomMargin=0>
<!--#include file="../libroutines.asp"-->

  <table border="0" width="100%" cellspacing="0" cellpadding="0">
    <tr>
    	<td width="100%">
    	
    		
   		</td>
	 </tr>
    
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
        						
        						'get filter clause if any
        						
        						For Each Thing In Request.QueryString 
        							
        							If filterClause <> "" Then        						
        								if Thing = "XTra" then
        									filterClause = filterClause & " AND " & Request.QueryString("XTra")
        								else
        									filterClause = filterClause & " AND [" & Thing & "] = " & Request.QueryString(Thing) 
        								end if 
        							Else
        								if Thing = "XTra" then
        									filterClause = Request.QueryString("XTra")
        								else
        									filterClause = " [" & Thing & "] = " & Request.QueryString(Thing)
        								end if
        								 
        							End If	
        							
        						Next
        						
        						
        						'WHERE user_rowguid = '" & Session("UserID") & "' 
        						'prevArtSQL = "SELECT * FROM I_Articles " & filterClause & " ORDER BY CreateDate DESC"
        						
        						If filterClause = "" Then        							
        							Set Conn = Nothing
        							Response.End         
        						Else
        							filterClause = " WHERE " & filterClause								
        						End If
        					
        						prevArtSQL = "SELECT * FROM ContractListWithNet " & filterClause & "  "
        						'prevArtSQL = "SELECT * FROM ContractList " & filterClause & "  "
        						
        					
								Set Rs = Server.CreateObject("ADODB.Recordset")
								Rs.Open prevArtSQL, Conn.ConnectionString, adOpenKeySet, adLockOptimistic
	
        						
        						If Not (Rs.EOF Or Rs.BOF) Then        							%>
        						<tr bgColor="#C0C0C0">
        							<td align="top" nowrap></td>
        							<td align="top" nowrap><b>Ref</b></td>
        							<td align="top" nowrap><b>Trade Date</b></td>
        							<td align="top" nowrap><b>Contract</b></td>
        							<td align="top" nowrap><b>Quantity</b></td>
        							<td align="top" nowrap><b>Price</b></td>
        							<td align="top" nowrap><b>Gross</b></td>
								</tr>
        						<%
        							articleCount = 0
        							Do Until Rs.EOF        								
        									'grab values here,
        									'to be placed in the form
        									
        									displayValue = Rs("OrdDetailClient") & " : " & Rs("ContractNumber") 
        									articleCount = articleCount + 1
        									idStr = "input-" & articleCount
        									displayAmt = cdbl(Rs.Fields("NetAmount")) 
        									
        									If CBool(Rs.Fields("ContractClientVouchered").Value) = True Then
        										chkString = " checked "
        									Else
        										chkString = " "
        									End If	
        									
        									%>
        									<tr>
        										<td align="top" nowrap> 
        											<input type="checkbox" <%= chkString %> class="BorderLess" style="background-color: transparent" name="<%= Rs.Fields("Contract_DPA_").Value %>" ID="<%= idStr %>" value="<%=displayAmt%>" OnClick="JavaScript: updateOwnerDocument(this)">        											
        										</td>
        										       										
        										<td align="top" nowrap><%= Rs.Fields("LotSlipNo").Value   %></td>
        										<td align="top" nowrap><%= FormatDate(Rs.Fields("LotTDate").Value)   %></td>
        										<td align="top" nowrap><%= Rs.Fields("ContractNumber").Value   %></td>
        										<td align="top" nowrap><%= FormatNumCommasOnly(Rs.Fields("LotQty").Value)   %></td>
        										<td align="top" nowrap><%= FormatNum(Rs.Fields("LotPrice").Value)   %></td>
        										<td align="top" nowrap><%= FormatNum(displayAmt)  %></td>
											</tr>
											
											<%If Trim(chkString) = "checked" Then%>
											
										
        									<%
        									End If
        								Rs.MoveNext
        							Loop        							
        						End If
        						
        						'clean up
        						Set Rs = Nothing
        						Set Conn = Nothing
        						
        						
        						%>
        			
			
		</table>				   
       
       </TD>       
       <!-- End page's main data -->	      
    </tr>
  
  </table> 	

<Script Language="JavaScript">
	function updateOwnerDocument(chkBox){
		
		if (chkBox.checked==true){
			window.parent.UpdateVoucherAmount(chkBox.value, 1); 
		}
		else {
			window.parent.UpdateVoucherAmount(chkBox.value, 0); 
		}
		
		UpdateSelContracts(false);
	}
</Script>

<Script Language="VBScript">
 'Updates the list of selected contracts.
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
							
				'Update Contract Net Amount in Payment Total
				If updateVals = True Then
					window.parent.UpdateVoucherAmount Thing.value , 1 
				End If
			End If
		Next 
		
		window.parent.document.getElementById("ContractsSel").value = returnVal
		
	End Function	

</Script>


</body>

</html>
