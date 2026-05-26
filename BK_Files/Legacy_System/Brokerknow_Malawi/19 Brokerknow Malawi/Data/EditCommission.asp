
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
 <script language='javascript'>
		function forceSubmit()
		{
			setOpener();
			//var targetPage = window.dialogArguments.opener.document.all.item("frmMain").elements("EditPage").value;
					
			document.frmEditCommission.method='post';
			document.frmEditCommission.target='_self';
			document.frmEditCommission.submit();		
		}
		
		function setOpener()
		{
			window.self.opener = window.dialogArguments.opener;					
		}

</script>

 </head>

<body Class="Dialog" onload="setOpener()">
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
	   Dim Vatable
	         
       BondRate = Request.Form("txtBondRate")
       description = Request.Form("txtDescription")
       rate = Request.Form("txtRate")
       secBoundary = Request.Form("txtSecurityBoundary")
       bondBoundary = Request.Form("txtBondBoundary")
       secRateAbove = Request.Form("txtSecurityAbove")
       bondRateAbove = Request.Form("txtBondAbove")
       secMin = Request.Form("txtSecurityMin")
       bondMin = Request.Form("txtBondMin")
	 ' Vatable = Request.Form("cboVatable")

	 
	  secondBondBoundary = trim(Request.Form("txtBondBoundaryLevel2"))
	  midBondRate = trim(Request.Form("txtBondRateMid"))
	  secondSecurityBoundary = trim(Request.Form("txtSecurityBoundaryLevel2"))
	  midSecurityRate = trim(Request.Form("txtMiddleRate"))

	  'validate Minimum Bond Commission
        If Trim(bondMin) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the  Minimum Bond Commission"
                		
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If
        'ensure  Minimum Bond Commission is numeric
        If (bondMin <> "") And (Not IsNumeric(bondMin)) Then%>
                <script language = 'vbscript'>
                ShowMessage " Minimum Bond Commission  must be numeric"
                
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If
       
       'validate Minimum Security Commission
        If Trim(secMin) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the  Minimum Security Commission"
                		
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If
        'ensure  Minimum Security Commission is numeric
        If (secMin <> "") And (Not IsNumeric(secMin)) Then%>
                <script language = 'vbscript'>
                ShowMessage " Minimum Security Commission  must be numeric"
                
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If
       
       'validate Upper Bond Rate
        If Trim(bondRateAbove) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Upper Bond Rate"
                		
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If
        'ensure Upper Bond Rate is numeric
        If (bondRateAbove <> "") And (Not IsNumeric(bondRateAbove)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Upper Bond Rate  must be numeric"
                
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If
       
       'validate Upper Security Rate
        If Trim(secRateAbove) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Upper Security Rate"
                		
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If
        'ensure Upper Security Rate is numeric
        If (secRateAbove <> "") And (Not IsNumeric(secRateAbove)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Upper Security Rate  must be numeric"
                
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If

		
'******************

		'validate Minimum Security Commission
        If Trim(secondBondBoundary) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the  Second Boundary range"
                		
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If
        'ensure  Minimum Security Commission is numeric
        If (secondBondBoundary <> "") And (Not IsNumeric(secondBondBoundary)) Then%>
                <script language = 'vbscript'>
                ShowMessage " Second Boundary range  must be numeric"
                
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If


		If Trim(midBondRate) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the  Middle Bond rate"
                		
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If
        'ensure  Minimum Security Commission is numeric
        If (midBondRate <> "") And (Not IsNumeric(midBondRate)) Then%>
                <script language = 'vbscript'>
                ShowMessage " Middle Bond rate must be numeric"
                
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If

		'


		 If Trim(secondSecurityBoundary) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the  Second Security Boundary Rate"
                		
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If
        'ensure  Minimum Security Commission is numeric
        If (secondSecurityBoundary <> "") And (Not IsNumeric(secondSecurityBoundary)) Then%>
                <script language = 'vbscript'>
                ShowMessage " Second Security Boundary range  must be numeric"
                
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If


		If Trim(midSecurityRate) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the  Middle Security rate"
                		
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If
        'ensure  Minimum Security Commission is numeric
        If (midSecurityRate <> "") And (Not IsNumeric(midSecurityRate)) Then%>
                <script language = 'vbscript'>
                ShowMessage " Middle Security rate must be numeric"
                
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If

'********
      
       'validate Bond Boundary
        If Trim(bondBoundary) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Bond Boundary"
                		
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If
        'ensure Bond Boundary is numeric
        If (bondBoundary <> "") And (Not IsNumeric(bondBoundary)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Bond Boundary  must be numeric"
                
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If
      
       'validate Security Boundary
        If Trim(secBoundary) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Security Boundary"
                		
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If
        'ensure Security Boundary is numeric
        If (secBoundary <> "") And (Not IsNumeric(secBoundary)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Security Boundary  must be numeric"
                
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If
        
        
        'validate Lower Security Rate
        If Trim(Rate) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Lower Security Rate"
                		
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If
        'ensure Lower Security Rate is numeric
        If (Rate <> "") And (Not IsNumeric(Rate)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Lower Security Rate  must be numeric"
                
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If
        'validate Lower Bond Rate
        If Trim(BondRate) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Lower Bond Rate"
                		
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If
        'ensure Lower Bond Rate is numeric
        If (BondRate <> "") And (Not IsNumeric(BondRate)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Lower Bond Rate  must be numeric"
                
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If
        'validate Description
        If Trim(Description) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Description"
                		
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If
        'validate size of Description
        If Len(Description) > 100 Then%>
                <script language = 'vbscript'>
                ShowMessage "Description can only be 100 characters in length"
                
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If
		 If trim(vatable) = "" or len(trim(vatable))=0 Then
               vatable=0
        End If
        
        
        'save data
		

		sqlStr = "INSERT INTO [Commission] (CommissionDescription,CommissionRate,BondCommission," & _
				"SecurityBoundary,BondBoundary,UpperBondCommission,UpperSecurityCommission,MinimumBondCommission,MinimumSecurityCommission,Vatable) values( " & "'" & description & "'"  & _
                "," & " " & rate & " "  & _
                "," & " " & BondRate & " "  & _
                "," & " " & secBoundary & " "  & _
                "," & " " & bondBoundary & " "  & _
                "," & " " & bondRateAbove & " "  & _
                "," & " " & secRateAbove & " "   & _
                "," & " " & bondMin & " "   & _
                "," & " " & secMin & " " & " " & _
                "," & " " &  Vatable & " )"
		 sqlStr = "UPDATE [Commission] SET CommissionDescription = " & "'" & description & "'" & _
				",CommissionRate = " & " " & rate & " " & "" & _
				",BondCommission = " & " " & BondRate & " " & "" & _
				",SecurityBoundary = " & " " & secBoundary & " " & "" & _
				",BondBoundary = " & " " & bondBoundary & " " & "" & _
				",UpperBondCommission = " & " " & bondRateAbove & " " & "" & _
				",UpperSecurityCommission = " & " " & secRateAbove & " " & "" & _
				",MinimumBondCommission = " & " " & bondMin & " " & "" & _
				",MinimumSecurityCommission = " & " " & secMin & " " & "" & _
				",MedianSecurityCommission="  &  midSecurityRate & " " & _
				",MedianBondCommission="  &  midBondRate & " " & _
				",SecondSecurityBoundary="  &  secondSecurityBoundary & " " & _
				",SecondBondBoundary="  &  secondBondBoundary & " " & _
                " WHERE Commission_DPA_  = " & ID   
				

			'	response.write sqlStr : response.end
        
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
        
       
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript2
        response.end
   	end If

%>

<form name = 'frmEditCommission' method = 'post' action = 'EditCommission.asp' >
<table border="0" width="100%" cellspacing="1" cellpadding="1">
<%
        Set conn = GetActiveConnection("KBroker")
        
        sqlStr = "SELECT CommissionDescription,CommissionRate,BondCommission," & _
				"SecurityBoundary,BondBoundary,UpperBondCommission,UpperSecurityCommission," & _
				"MinimumBondCommission,MinimumSecurityCommission,Commission_DPA_,MedianSecurityCommission,MedianBondCommission,SecondSecurityBoundary,SecondBondBoundary FROM [Commission] WHERE Commission_DPA_  = " & ID
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
    <td>Gross Amount Boundary(Level 1 and 2)</td>
    <td><input type = 'text' name ='txtSecurityBoundary' id = 'txtSecurityBoundary' size="20" value = '<%=rs.Fields("SecurityBoundary")%>'></td>
   </tr>
	<tr>
    <td>Gross Amount Boundary(Level 2 and 3)</td>
    <td><input type = 'text' name ='txtSecurityBoundaryLevel2' id = 'txtSecurityBoundaryLevel2' size="20" value = '<%=rs.Fields("SecondSecurityBoundary")%>'></td>
   </tr>
   <tr>
    <td>Lower Rate (Level 1)</td>
    <td><input type = 'text' name ='txtRate' id = 'txtRate' size="20" value = '<%=rs.Fields("CommissionRate")%>'></td>
   </tr>
   <tr>
    <td>Middle Rate(Level 2)</td>
    <td><input type = 'text' name ='txtMiddleRate' id = 'txtMiddleRate' size="20" value = '<%=rs.Fields("MedianSecurityCommission")%>'></td>
   </tr>

     
  
  
   <tr>
    <td>Upper&nbsp; Rate (Level 1)</td>
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
    <td>Gross Amount Boundary (Level 1 and 2)</td>
    <td><input type = 'text' name ='txtBondBoundary' id = 'txtBondBoundary' size="20" value = '<%=rs.Fields("BondBoundary")%>'></td>
  </tr>
   <tr>
    <td>Gross Amount Boundary (Level 2 and 3)</td>
    <td><input type = 'text' name ='txtBondBoundaryLevel2' id = 'txtBondBoundaryLevel2' size="20" value = '<%=rs.Fields("SecondBondBoundary")%>'></td>
  </tr>
  
  <tr>
    <td>Lower Rate (Level 1)</td>
    <td><input type = 'text' name ='txtBondRate' id = 'txtBondRate' size="20" value = '<%=rs.Fields("BondCommission")%>'></td>
  </tr>
  
   <tr>
    <td>Middle Rate (Level 2)</td>
    <td><input type = 'text' name ='txtBondRateMid' id = 'txtBondRateMid' size="20" value = '<%=rs.Fields("MedianBondCommission")%>'></td>
  </tr>
  <tr>
    <td>Upper&nbsp; Rate (Level 3)</td>
    <td><input type = 'text' name ='txtBondAbove' id = 'txtBondAbove' size="20" value = '<%=rs.Fields("UpperBondCommission")%>'></td>
  </tr>

  
  <tr>
    <td>Minimum Amount</td>
    <td><input type = 'text' name ='txtBondMin' id = 'txtBondMin' size="20" value = '<%=rs.Fields("MinimumBondCommission")%>'></td>
  </tr>
  <tr>
     <!--<tr>
    <td>VAT</td>
    <td><input type = 'checkbox' name ='cboVatable' id = 'chkvat' size="3" value=1></td>
  </tr>-->
  <tr>
    <td width="100%" Colspan=2 align=right>
		<BR><BR>
		<input type = 'button' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save " onclick="forceSubmit();">
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
