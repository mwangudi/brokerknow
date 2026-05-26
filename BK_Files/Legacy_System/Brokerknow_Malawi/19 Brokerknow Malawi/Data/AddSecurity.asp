<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Security</title>
 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>

<!--CALENDAR -->
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>
<script language="JavaScript" src="CALENDAR/calendar.js"></script>

<!--END CALENDAR -->


<script language="javascript">
		function  UpdateImmobilised(theChk)
		{
			if (theChk.checked)
			{
				document.frmMain.elements("txtImmobilised").value = "1";
			}
			else
			{
				document.frmMain.elements("txtImmobilised").value = "0";
			}			
		}

		function forceSubmit()
		{
			setOpener();
			//var targetPage = window.dialogArguments.opener.document.all.item("frmMain").elements("EditPage").value;
					
			document.frmAddSecurity.method='post';
			document.frmAddSecurity.target='_self';
			document.frmAddSecurity.submit();		
		}
		
		function setOpener()
		{
			window.self.opener = window.dialogArguments.opener;					
		}
</script>
</head>

<body Class="Dialog" onLoad="setOpener()">

<!--#include file="../libroutines.asp"-->
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<SCRIPT language="JavaScript">			
			var cal=new ctlSpiffyCalendarBox("cal", "frmAddSecurity", "txtADate","cmdDate","<%= FormatDate(Date) %>",1);
			var cal1=new ctlSpiffyCalendarBox("cal1", "frmAddSecurity", "txtEDate","cmdDate","",1);
			
</SCRIPT>
<%
	 
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim guidStr 
   Dim guid 
	
	action = ucase(Request.Form("action"))
	
	if action = "EXECUTE" then
	   Dim name
       Dim addr
       Dim transFee
       Dim mktPrice
       Dim code
       Dim adate
       Dim fee
       Dim secType
       Dim immob
       Dim Sector
       dim ImportCode
	   Dim Edate
	   Dim chkCDSstatus
       
        immob = trim(Request.Form("txtImmobilised"))
        code = trim(Request.Form("txtCode"))
        name = trim(Request.Form("txtName"))
		addr = trim(Request.Form("txtAddr"))
        mktPrice = trim(Request.Form("txtMktPrice"))
        adate = trim(Request.Form("txtADate"))
        fee = trim(Request.Form("txtFee"))
        Sector = trim(Request.Form("cboSector"))
        importCode = trim(Request.Form("txtimportCode"))
		Edate = trim(Request.Form("txtEDate"))
		
       
       'secType =  Request.Form("cboSecType")
       secType =  2
       
        'validate Name
        If Trim(Name) = "" Then%>
				<script language="vbscript">
                						ShowMessage "Please specify the Name"
                						
								</script>
				<% 
				ReloadPage(ID)
				response.end
        End If
        
        'validate Market Price
        If Trim(mktPrice) = "" Then%>
				<script language="vbscript">
                						ShowMessage "Please specify the Market Price"
                						
								</script>
				<% 
				ReloadPage(ID)
				response.end
        End If
        'validate size of Address
        If Len(Addr) > 100 Then%>
				<script language="vbscript">
								ShowMessage "Address can only be 100 characters in length"
								
								</script>
				<% 
				ReloadPage(ID)
				response.end
        End If
        'validate size of Code
        If Len(Code) > 50 Then%>
				<script language="vbscript">
								ShowMessage "Code cannot be more than 50 characters"
								
								</script>
				<% 
				ReloadPage(ID)
				response.end
						End If
        
        
        if Trim(Sector) = "" Then%>
				<script language="vbscript">
                						ShowMessage "Please specify the Market Sector"
                						
								</script>
				<% 
				ReloadPage(ID)
				response.end
        End If
        'ensure Market Price is numeric
        If (MktPrice <> "") And (Not IsNumeric(MktPrice)) Then%>
			<script language="vbscript">
							ShowMessage "Market Price  must be numeric"
							
							</script>
			<% 
			ReloadPage(ID)
			response.end
        End If
        'validate size of Name
        If Len(Name) > 100 Then%>
			<script language="vbscript">
							ShowMessage "Name can only be 100 characters in length"
							
							</script>
			<% 
			ReloadPage(ID)
			response.end
        End If
     	'validate Security Type
        If Trim(secType) = "" Then%>
				<script language="vbscript">
                						ShowMessage "Please specify the Security Type"
                						
								</script>
				<% 
				ReloadPage(ID)
				response.end
        End If
		'validate detail info
				'validate Transfer Fee
                If Trim(Fee) = "" Then%>
                		<script language = 'vbscript'>
                        ShowMessage "Please specify the Fee"
                        
                		</script>
                		<% 
						ReloadPage(ID)
						response.end
                End If
                'ensure Transfer Fee is numeric
                If (Fee <> "") And (Not IsNumeric(Fee)) Then%>
                		<script language = 'vbscript'>
                        ShowMessage "Fee must be numeric"
                        
                		</script>
                		<% 
						ReloadPage(ID)
						response.end
                End If
              
			  'Validate Expiry date

			  if Edate <> "" AND isdate(Edate) = false then
			     %>
                		<script language = 'vbscript'>
                        ShowMessage "Please specify a valid Expiry date."
                        
                		</script>
                		<% 
						ReloadPage(ID)
						response.end
			  end if

		   if Edate = "" then
			Edate = "NULL"
		   Else
			Edate = "#" & formatdate(Edate) & "#"
		   End if
                
        'save header
        set guid = server.createobject("NDUtils.CGUID")
        guidStr = guid.GenerateGUID
        
        sqlStr = "INSERT INTO [Security] (SecurityAddr,SecurityCode,SecurityMktPrice,SecurityName" & _
                "       ,Security_DPA_,OrderSecType_DPA_,Sector_DPA_,Immobilised,ImportCode,Security_EIT_,ExpiryDate) SELECT " & "'" & addr & "'" & " as SecurityAddr," & "'" & Code & "'" & " as SecurityCode" & _
                "       ," & " " & MktPrice & " " & " as SecurityMktPrice" & _
                "       ," & "'" & Name & "'" & " as SecurityName," & " " & _
				"Security_DPA_ = case (isnull(max([Security_DPA_]),0)) " & _
				"                when 0 then (SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Security') " & _
				"                else max([Security_DPA_]) + 1 " & _
				"                end " & _
                "		," & " " & secType & " " & " as OrderSecType_DPA_" & _
                "		," & " " & Sector & " " & " as Sector_DPA_" & _
                "		," & " " & immob & " " & " as Immobilised" & _
                "		," & "'" & ImportCode & "'" & " as ImportCode" & _
                "       ," & "'" & guidStr & "'" & " as Security_EIT_ " & _
				 "       ," & " " & Edate & " " & " as ExpiryDate FROM [Security]"

		'response.write sqlStr
		'response.end

				         
        Set conn = GetActiveConnection("KBroker")
        
        conn.BeginTrans
		 
			conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
	        
			'obtain header key value
			sqlStr = "SELECT [Security.Security_DPA_] FROM [Security] WHERE [Security.Security_EIT_] = " & "'" & guidStr & "'"
	        
			Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			If (rs.EOF Or rs.BOF) Then%>
					<script language = 'vbscript'>
                			ShowMessage "A serious error has been encountered while saving the data. Try saving again"
                			
					</script>
					<% 					
					ReloadPage()
					response.end
			End If
	        
			'save detail data
			 sqlStr = "INSERT INTO [SecTransFee] (SecTransFeeADate,SecTransFeeFee,SecTransFee_DPA_" & _
                "                       ,Security_DPA_) SELECT " & "#" & FormatDate(adate) & " " & Time & "#" & " as SecTransFeeADate" & _
                "       ," & " " & fee & " " & " as SecTransFeeFee" & _
                "       ," & " " & "iif(isnull(max([SecTransFee_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'SecTransFee'),max([SecTransFee_DPA_]) + 1)" & " " & " as SecTransFee_DPA_" & _
                "       ," & " " & rs.Fields("Security_DPA_") & " " & " as Security_DPA_" & _
                "        FROM [SecTransFee]"
	        
			conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
		conn.CommitTrans
        Set Conn = Nothing
        WritefraEnabledDialogCloseScript2
        Response.End 
   	end If
   	
%>

<form name = 'frmAddSecurity' method = 'post' id="frmMain" action = "AddSecurity.asp" >
<table border="0" width="100%" cellpadding=2 cellspacing=2>
 <tr>
                <td width="30%">Name</td>
                <td width="70%"><input type="text" name="txtName" id="txtName" size="25"></td>
              </tr>
              <tr>
                <td width="30%">Code</td>
                <td width="70%"><input type="text" name="txtCode" id="txtCode" size="25"></td>
              </tr>
              <tr>
                <td width="30%">Address</td>
                <td width="70%"><textarea rows=3 name ='txtAddr' id = "txtAddr"></textarea></td>
              </tr>
              <tr>
                <td width="30%">Market Sector</td>
                <td width="30%"><select name = 'cboSector' id = 'cboSector' size="1">
    	
	<%
		Set conn = GetActiveConnection("KBroker")
        sqlStr = "SELECT * FROM [MarketSectorList] Order By ShortDescription ASC"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF
                %>
						<option value = '<%=rs.Fields("Sector_DPA_")%>'><%=rs.Fields("ShortDescription")%></option>
             <%  
                    rs.MoveNext
                Loop
        End If
	%>

    </select></td>
    </td> 
    
              </tr>
              <tr>
              <tr>
                <td width="30%">Market Price</td>
                <td width="70%"><input type="text" name="txtMktPrice" id="txtMktPrice" size="25"></td>
              </tr>
              <tr>
        <tr>
  <td width="30%">Expiry Date</td>
  <td width="70%">
		<SCRIPT language="JavaScript">			
			cal1.writeControl();
		</SCRIPT>
		
		</td>
  </tr>              
     <tr>
  <td width="30%">Activation Date</td>
  <td width="70%">
		<SCRIPT language="JavaScript">			
			cal.writeControl();
		</SCRIPT>
		
		</td>
  </tr>          
  
 <tr>
    <td>Immobilised</td>
    <td><input type=checkbox Class="BorderLess"   value='False' name='chkImmobilised' onClick = 'UpdateImmobilised(this);'> 
      </td>
  </tr>
  
  <tr>
  <td width="30%">Transfer Fee</td>
  <td width="70%"><input type = 'text' name ='txtFee' id = 'txtFee' size="25"></td>
  </tr>
   <tr>
  <td width="30%">Import Code</td>
  <td width="70%"><input type = 'text' name ='txtImportCode' id = 'txtImportCode' size="25"></td>
  </tr>
  <tr>
     <td width="100%" colspan="2" align=right>
		<BR>
		<BR>
		<BR>
		<b  name="hide" id="hide">
		<input type = 'button' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save " onclick="forceSubmit();"></b>
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
		&nbsp;&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
		<input type = 'hidden' name ='txtImmobilised' id = 'txtImmobilised' value='0'>

      </td>
  </tr>
</table>
</form>


</body>

</html>
