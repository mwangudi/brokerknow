<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Commission Type</title>
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 

</head>


 <Script Language="VBScript">
	Function SelectForm
		For Each Thing In frmAddLevy
			If InStr(1, Thing.Name, "SecuritiesSel") > 0 Then
				SelectAll Thing
			End If
		Next
	End Function
 </Script>
 
 <script language='javascript'>
		function forceSubmit()
		{
			setOpener();
			//var targetPage = window.dialogArguments.opener.document.all.item("frmMain").elements("EditPage").value;
					
			document.frmAddCommission.method='post';
			document.frmAddCommission.target='_self';
			document.frmAddCommission.submit();		
		}
		
		function setOpener()
		{
			window.self.opener = window.dialogArguments.opener;					
		}

</script>

<body Class="Dialog" onload="setOpener()">
<!--#include file="../libroutines.asp"-->
<%
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
	
	action = ucase(Request.Form("action"))
	
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
		secondSecurityBoundary= trim(request.form("txtSecurityBoundaryMiddle"))
		middleSecurityrate = trim(request.form("txtMiddleRate"))
		secondBondBoundary= trim(request.form("txtMiddleBondBoundary"))
		middleBondRate = trim(request.form("txtMiddleBondRate"))

	          
       'validate Minimum Bond Commission
        If Trim(bondMin) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the  Minimum Bond Commission"
                		
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If
        'ensure  Minimum Bond Commission is numeric
        If (bondMin <> "") And (Not IsNumeric(bondMin)) Then%>
                <script language = 'vbscript'>
                ShowMessage " Minimum Bond Commission  must be numeric"
                
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If
       
       'validate Minimum Security Commission
        If Trim(secMin) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the  Minimum Security Commission"
                		
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If
        'ensure  Minimum Security Commission is numeric
        If (secMin <> "") And (Not IsNumeric(secMin)) Then%>
                <script language = 'vbscript'>
                ShowMessage " Minimum Security Commission  must be numeric"
                
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If
		


		'validate Minimum Security Commission
        If Trim(middleSecurityrate) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the  Middle Security Commission"
                		
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If


		'ensure  Median Security Commission is numeric
        If (middleSecurityrate <> "") And (Not IsNumeric(middleSecurityrate)) Then%>
                <script language = 'vbscript'>
                ShowMessage " Middle Security Commission  must be numeric"
                
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If
       
       'validate Upper Bond Rate
        If Trim(bondRateAbove) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Upper Bond Rate"
                		
                </script>
                <%
				ReloadPage(ID)
				response.end
        End If
        'ensure Upper Bond Rate is numeric
        If (bondRateAbove <> "") And (Not IsNumeric(bondRateAbove)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Upper Bond Rate  must be numeric"
                
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If
'*************
		'validate middle bond Commission
        If Trim(middleBondRate) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the  Middle Bond Commission"
                		
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If

		
		'ensure  Median bond Commission is numeric
        If (middleBondRate <> "") And (Not IsNumeric(middleBondRate)) Then%>
                <script language = 'vbscript'>
                ShowMessage " Middle Bond Commission  must be numeric"
                
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If

		'*************
       
       'validate Upper Security Rate
        If Trim(secRateAbove) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Upper Security Rate"
                		
                </script>
                <% ReloadPage(ID)
				response.end
        End If
        'ensure Upper Security Rate is numeric
        If (secRateAbove <> "") And (Not IsNumeric(secRateAbove)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Upper Security Rate  must be numeric"
                
                </script>
                <% ReloadPage(ID)
				response.end
        End If
      
       'validate Bond Boundary
        If Trim(bondBoundary) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Bond Boundary"
                		
                </script>
                <% ReloadPage(ID)
				response.end
        End If
        'ensure Bond Boundary is numeric
        If (bondBoundary <> "") And (Not IsNumeric(bondBoundary)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Bond Boundary  must be numeric"
                
                </script>
                <% ReloadPage(ID)
				response.end
        End If
      
       'validate Security Boundary
        If Trim(secBoundary) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Security Boundary"
                		
                </script>
                <% ReloadPage(ID)
				response.end
        End If
        'ensure Security Boundary is numeric
        If (secBoundary <> "") And (Not IsNumeric(secBoundary)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Security Boundary  must be numeric"
                
                </script>
                <% ReloadPage(ID)
				response.end
        End If
        
		'*******
		 'validate second Security Boundary
        If Trim(secondSecurityBoundary) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Second Security Boundary"
                		
                </script>
                <% ReloadPage(ID)
				response.end
        End If

        'ensure Second Security Boundary is numeric
        If (secondSecurityBoundary <> "") And (Not IsNumeric(secondSecurityBoundary)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Second Security Boundary  must be numeric"
                
                </script>
                <% ReloadPage(ID)
				response.end
        End If

		'validate Second Bond Boundary
        If Trim(secondBondBoundary) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Second Bond Boundary"
                		
                </script>
                <% ReloadPage(ID)
				response.end
        End If
        'ensure Second Bond is numeric
        If (secondBondBoundary <> "") And (Not IsNumeric(secondBondBoundary)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Second Bond Boundary  must be numeric"
                
                </script>
                <% ReloadPage(ID)
				response.end
        End If
		'********
        
        'validate Lower Security Rate
        If Trim(Rate) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Lower Security Rate"
                		
                </script>
                <% ReloadPage(ID)
				response.end
        End If
        'ensure Lower Security Rate is numeric
        If (Rate <> "") And (Not IsNumeric(Rate)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Lower Security Rate  must be numeric"
                
                </script>
                <% ReloadPage(ID)
				response.end
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
                <% ReloadPage(ID)
				response.end
        End If
        'validate Description
        If Trim(Description) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Description"
                		
                </script>
                <% ReloadPage(ID)
				response.end
        End If
        'validate size of Description
        If Len(Description) > 100 Then%>
                <script language = 'vbscript'>
                ShowMessage "Description can only be 100 characters in length"
                
                </script>
                <% ReloadPage(ID)
				response.end
        End If
        
        If trim(vatable) = "" or len(trim(vatable))=0 Then
               vatable=0
        End If
        
        
       
        'save data
       sqlStr = "INSERT INTO [Commission] (CommissionDescription,CommissionRate,BondCommission," & _
				"SecurityBoundary,BondBoundary,UpperBondCommission,UpperSecurityCommission,MinimumBondCommission,MinimumSecurityCommission,Commission_DPA_,Vatable) SELECT " & "'" & description & "'" & " as CommissionDescription" & _
                "," & " " & rate & " " & " as CommissionRate" & _
                "," & " " & BondRate & " " & " as BondCommission" & _
                "," & " " & secBoundary & " " & " as SecurityBoundary" & _
                "," & " " & bondBoundary & " " & " as BondBoundary" & _
                "," & " " & bondRateAbove & " " & " as UpperBondCommission" & _
                "," & " " & secRateAbove & " " & " as UpperSecurityCommission" & _
                "," & " " & bondMin & " " & " as MinimumBondCommission" & _
                "," & " " & secMin & " " & " as MinimumSecurityCommission" & _
                "," & " " & "iif(isnull(max([Commission_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Commission'),max([Commission_DPA_]) + 1)" & " " & " as Commission_DPA_" & _
                " FROM [Commission]," & Vatable & " As VAT"

	  sqlStr = "INSERT INTO [Commission] (CommissionDescription,CommissionRate,BondCommission," & _
				"SecurityBoundary,BondBoundary,UpperBondCommission,UpperSecurityCommission,MinimumBondCommission,MinimumSecurityCommission,MedianSecurityCommission,MedianBondCommission,SecondSecurityBoundary,SecondBondBoundary) values( " & "'" & description & "'"  & _
                "," & " " & rate & " "  & _
                "," & " " & BondRate & " "  & _
                "," & " " & secBoundary & " "  & _
                "," & " " & bondBoundary & " "  & _
                "," & " " & bondRateAbove & " "  & _
                "," & " " & secRateAbove & " "   & _
                "," & " " & bondMin & " "   & _
                "," & " " & secMin & " " & " " & _
                "," & " " &  middleSecurityrate & " "  & _
				"," & " " &  middleBondRate & " "  & _
				"," & " " &  secondSecurityBoundary & " "  & _
				"," & " " &  secondBondBoundary & ") "  


		



				'response.write sqlStr: response.end
        Set conn = GetActiveConnection("KBroker")
        
        conn.BeginTrans
                conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
        conn.CommitTrans
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript2
   	end If
%>


<form name = 'frmAddCommission' method = 'post' id="frmMain" action = 'AddCommission.asp' target="deleteFrame" OnSubmit="JavaScript: UpdateDialogHandle();">
    <input type = 'hidden' name ='cboCMA' id = 'cboCMA' value='0'>

<table border="0" cellspacing="1" cellpadding="1">
   <tr>
    <td> Description</td>
    <td><input type = 'text' name ='txtDescription' id = 'txtDescription' size="20"></td>
  </tr>
  
   <tr>
    <td><b>Security Structure</b></td>
    <td></td>
   </tr>
   <tr>
    <td>Gross Amount Boundary(Level 1 and 2)</td>
    <td><input type = 'text' name ='txtSecurityBoundary' id = 'txtSecurityBoundary' size="20"></td>
   </tr>

   <tr>
    <td>Gross Amount Boundary(Level 2 and 3)</td>
    <td><input type = 'text' name ='txtSecurityBoundaryMiddle' id = 'txtSecurityBoundaryMiddle' size="20"></td>
   </tr>
   <tr>
    <td>Lower Rate(Level 1)</td>
    <td><input type = 'text' name ='txtRate' id = 'txtRate' size="20"></td>
   </tr>

   <tr>
    <td>Middle Rate(Level 2)</td>
    <td><input type = 'text' name ='txtMiddleRate' id = 'txtMiddleRate' size="20"></td>
   </tr>
   <tr>
    <td>Upper&nbsp; Rate(Level 3)</td>
    <td><input type = 'text' name ='txtSecurityAbove' id = 'txtSecurityAbove' size="20"></td>
   </tr>
   <tr>
    <td>Minimum Amount</td>
    <td><input type = 'text' name ='txtSecurityMin' id = 'txtSecurityMin' size="20"></td>
   </tr>
  
  <tr>
    <td><b>Bond Structure</b></td>
    <td></td>
  </tr>
  <tr>
    <td>Gross Amount Boundary (Level 1 and 2)</td>
    <td><input type = 'text' name ='txtBondBoundary' id = 'txtBondBoundary' size="20"></td>
  </tr>
  
   <tr>
    <td>Gross Amount Boundary (Level 2 and 3)</td>
    <td><input type = 'text' name ='txtMiddleBondBoundary' id = 'txtMiddleBondBoundary' size="20"></td>
  </tr>
  
  <tr>
    <td>Lower Rate(Level 1)</td>
    <td><input type = 'text' name ='txtBondRate' id = 'txtBondRate' size="20"></td>
  </tr>


  <tr>
    <td>Middle Rate(Level 2)</td>
    <td><input type = 'text' name ='txtMiddleBondRate' id = 'txtMiddleBondRate' size="20"></td>
  </tr>

  
  <tr>
    <td>Upper&nbsp; Rate(Level 3)</td>
    <td><input type = 'text' name ='txtBondAbove' id = 'txtBondAbove' size="20"></td>
  </tr>
  
  
  <tr>
    <td>Minimum Amount</td>
    <td><input type = 'text' name ='txtBondMin' id = 'txtBondMin' size="20"></td>
  </tr>
  
  <!--<tr>
    <td>VAT</td>
    <td><input type = 'checkbox' name ='cboVatable' id = 'chkvat' size="20" value=1></td>
  </tr>-->
    
  <tr>
    <td colspan=2 align=right >
		<BR><BR>
		<input type = 'button' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save "  onclick="forceSubmit();">
		&nbsp;
		&nbsp;
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
      </td>
  </tr>
  
</table>
</form>


</body>

