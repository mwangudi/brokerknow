<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit CDS Trade</title>
   <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
</head>

<body Class="Dialog">
<!--#include file="../libroutines.asp"-->



<%
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim ID
   Dim rsEdit
	
	action = ucase(Request.Form("action"))
	ID = Request("ID")
	itemids=Split(ID,"<->")
	itemid=itemids(1)
	itemtypeid=itemids(0)				

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		
                </script>
                <%WriteDialogRefuseOpenScript 
                response.end
        End If

	if action = "EXECUTE" then
		Dim description
        
        Commission = Request.Form("txtCommission")
      
       
        'validate Description
        If Trim(Commission) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Commission"
                		
                </script>
                <% response.end
        End If
        'validate size of Description
        If Len(Commission) > 100 Then%>
                <script language = 'vbscript'>
                ShowMessage "Commission can only be 100 characters in length"
                
                </script>
                <% response.end
        End If
		
		'ensure Commission is numeric
				If (Not IsNumeric(Commission)) Then%>
				    <script language = 'vbscript'>
						ShowMessage "Commission must be numeric"
						
				    </script>
				    <% response.end
				End If

		Set conn = GetActiveConnection("KBroker")
        
        'save data
        
        sqlStr = "UPDATE [_CDS_Imported_Trades_] SET CommissionRate = " & "" & Commission & "" & " WHERE CDSImport_DPA_  = " & itemtypeid					

	   conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
        
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript
        Response.End
   	end If
%>
<BR>
<form name = 'frmEditResidency' method = 'post' action = 'EditCDSTrade.asp' >
<table border="0" width="100%" cellspacing="2" cellpadding="2">
<%
        Set conn = GetActiveConnection("KBroker")
        
       
        sqlStr = "SELECT CommissionRate, CDSImport_DPA_ FROM  _CDS_Imported_Trades_" & _
				 " WHERE (CDSImport_DPA_ = " & itemtypeid & ")"

        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected CDST Trade cannot be retrieved for editing"
                		
                </script>
                <% response.end
        End If
       

%>
  <tr>
    <td width="17%">Commission&nbsp;Rate</td>
    <td width="83%"><input type = 'text' name ='txtCommission' id = 'txtCommission' size="10" value = '<%=rs.Fields("CommissionRate")%>'></td>
  </tr>
  <tr>
    <td width="100%" colspan="2" align=right>
		<BR>
	
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
		&nbsp;&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
    	<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
      </td>
  </tr>
</table>
</form>

</body>

</html>
