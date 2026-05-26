<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Commission</title>
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
	
	toCancel = Request.Form("cmdCancel")
       
       Set conn = GetActiveConnection("KBroker")
       
       If toCancel <> "" Then
			
			WriteDialogCancelScript
			Set Conn = Nothing
			Response.End
       End If
       

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		
                </script>
                <% response.end
        End If

	if action = "EXECUTE" then
		Dim description
       Dim rate
       Dim BondRate
       Dim secBoundary
       Dim bondBoundary
       Dim secRateAbove
       Dim bondRateAbove
       Dim secMin
       Dim bondMin
	   BondRate = Request.Form("txtBondRate")
       description = Request.Form("txtDescription")
       rate = Request.Form("txtRate")
       secBoundary = Request.Form("txtSecurityBoundary")
       bondBoundary = Request.Form("txtBondBoundary")
       secRateAbove = Request.Form("txtSecurityAbove")
       bondRateAbove = Request.Form("txtBondAbove")
       secMin = Request.Form("txtSecurityMin")
       bondMin = Request.Form("txtBondMin")
	   'validate Minimum Bond Commission
        If Trim(bondMin) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the  Minimum Bond Commission"
                		
                </script>
                <% response.end
        End If
        'ensure  Minimum Bond Commission is numeric
        If (bondMin <> "") And (Not IsNumeric(bondMin)) Then%>
                <script language = 'vbscript'>
                ShowMessage " Minimum Bond Commission  must be numeric"
                
                </script>
                <% response.end
        End If
       
       'validate Minimum Security Commission
        If Trim(secMin) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the  Minimum Security Commission"
                		
                </script>
                <% response.end
        End If
        'ensure  Minimum Security Commission is numeric
        If (secMin <> "") And (Not IsNumeric(secMin)) Then%>
                <script language = 'vbscript'>
                ShowMessage " Minimum Security Commission  must be numeric"
                
                </script>
                <% response.end
        End If
       
       'validate Upper Bond Rate
        If Trim(bondRateAbove) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Upper Bond Rate"
                		
                </script>
                <% response.end
        End If
        'ensure Upper Bond Rate is numeric
        If (bondRateAbove <> "") And (Not IsNumeric(bondRateAbove)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Upper Bond Rate  must be numeric"
                
                </script>
                <% response.end
        End If
       
       'validate Upper Security Rate
        If Trim(secRateAbove) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Upper Security Rate"
                		
                </script>
                <% response.end
        End If
        'ensure Upper Security Rate is numeric
        If (secRateAbove <> "") And (Not IsNumeric(secRateAbove)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Upper Security Rate  must be numeric"
                
                </script>
                <% response.end
        End If
      
       'validate Bond Boundary
        If Trim(bondBoundary) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Bond Boundary"
                		
                </script>
                <% response.end
        End If
        'ensure Bond Boundary is numeric
        If (bondBoundary <> "") And (Not IsNumeric(bondBoundary)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Bond Boundary  must be numeric"
                
                </script>
                <% response.end
        End If
      
       'validate Security Boundary
        If Trim(secBoundary) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Security Boundary"
                		
                </script>
                <% response.end
        End If
        'ensure Security Boundary is numeric
        If (secBoundary <> "") And (Not IsNumeric(secBoundary)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Security Boundary  must be numeric"
                
                </script>
                <% response.end
        End If
        
        
        'validate Lower Security Rate
        If Trim(Rate) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Lower Security Rate"
                		
                </script>
                <% response.end
        End If
        'ensure Lower Security Rate is numeric
        If (Rate <> "") And (Not IsNumeric(Rate)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Lower Security Rate  must be numeric"
                
                </script>
                <% response.end
        End If
        'validate Lower Bond Rate
        If Trim(BondRate) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Lower Bond Rate"
                		
                </script>
                <% response.end
        End If
        'ensure Lower Bond Rate is numeric
        If (BondRate <> "") And (Not IsNumeric(BondRate)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Lower Bond Rate  must be numeric"
                
                </script>
                <% response.end
        End If
        'validate Description
        If Trim(Description) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Description"
                		
                </script>
                <% response.end
        End If
        'validate size of Description
        If Len(Description) > 100 Then%>
                <script language = 'vbscript'>
                ShowMessage "Description can only be 100 characters in length"
                
                </script>
                <% response.end
        End If
        
        'save data
        
        sqlStr = "UPDATE [Commission] SET CommissionDescription = " & "'" & description & "'" & _
				",CommissionRate = " & " " & rate & " " & "" & _
				",BondCommission = " & " " & BondRate & " " & "" & _
				",SecurityBoundary = " & " " & secBoundary & " " & "" & _
				",BondBoundary = " & " " & bondBoundary & " " & "" & _
				",UpperBondCommission = " & " " & bondRateAbove & " " & "" & _
				",UpperSecurityCommission = " & " " & secRateAbove & " " & "" & _
				",MinimumBondCommission = " & " " & bondMin & " " & "" & _
				",MinimumSecurityCommission = " & " " & secMin & " " & "" & _
                " WHERE Commission_DPA_  = " & ID                
        
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
        
       
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript
        response.end
   	end If

%>

<form name = 'frmEditCommission' method = 'post' action = 'EditCommission.asp' >
<table border="0" width="100%" cellspacing="1" cellpadding="1">
<%
        Set conn = GetActiveConnection("KBroker")
        
        sqlStr = "SELECT CommissionDescription,CommissionRate,BondCommission," & _
				"SecurityBoundary,BondBoundary,UpperBondCommission,UpperSecurityCommission," & _
				"MinimumBondCommission,MinimumSecurityCommission,Commission_DPA_ FROM [Commission] WHERE Commission_DPA_  = " & ID
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected Commission cannot be retrieved for editing"
                		
                </script>
                <% response.end
        End If
        
       
%>
 
  <tr>
    <td> Description</td>
    <td><input type = 'text' name ='txtDescription' id = 'txtDescription' size="20" value = '<%=rs.Fields("CommissionDescription")%>'></td>
  </tr>
  
   <tr>
    <td><b>Security Structure</b></td>
    <td></td>
   </tr>
   <tr>
    <td>Gross Amount Boundary</td>
    <td><input type = 'text' name ='txtSecurityBoundary' id = 'txtSecurityBoundary' size="20" value = '<%=rs.Fields("SecurityBoundary")%>'></td>
   </tr>
   <tr>
    <td>Lower Rate</td>
    <td><input type = 'text' name ='txtRate' id = 'txtRate' size="20" value = '<%=rs.Fields("CommissionRate")%>'></td>
   </tr>
   <tr>
    <td>Upper&nbsp; Rate</td>
    <td><input type = 'text' name ='txtSecurityAbove' id = 'txtSecurityAbove' size="20" value = '<%=rs.Fields("UpperSecurityCommission")%>'></td>
   </tr>
   <tr>
    <td>Minimum Amount</td>
    <td><input type = 'text' name ='txtSecurityMin' id = 'txtSecurityMin' size="20" value = '<%=rs.Fields("MinimumSecurityCommission")%>'></td>
   </tr>
  
  <tr>
    <td><b>Bond Structure</b></td>
    <td></td>
  </tr>
  <tr>
    <td>Gross Amount Boundary</td>
    <td><input type = 'text' name ='txtBondBoundary' id = 'txtBondBoundary' size="20" value = '<%=rs.Fields("BondBoundary")%>'></td>
  </tr>
  
  
  <tr>
    <td>Lower Rate</td>
    <td><input type = 'text' name ='txtBondRate' id = 'txtBondRate' size="20" value = '<%=rs.Fields("BondCommission")%>'></td>
  </tr>
  
  
  <tr>
    <td>Upper&nbsp; Rate</td>
    <td><input type = 'text' name ='txtBondAbove' id = 'txtBondAbove' size="20" value = '<%=rs.Fields("UpperBondCommission")%>'></td>
  </tr>
  
  
  <tr>
    <td>Minimum Amount</td>
    <td><input type = 'text' name ='txtBondMin' id = 'txtBondMin' size="20" value = '<%=rs.Fields("MinimumBondCommission")%>'></td>
  </tr>
  <tr>
    <td width="100%" Colspan=2 align=right>
		<BR><BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
		&nbsp;&nbsp;
		<input type = 'button' Class=Buttons name ='cmdClose' id = 'cmdClose' value=" Cancel " OnClick="JavaScript: window.self.close();">
    	<input type = 'hidden' name ='action' id = 'action' value="Execute">
    	<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
    </td>
  </tr>
</table>
</form>

</body>

</html>
