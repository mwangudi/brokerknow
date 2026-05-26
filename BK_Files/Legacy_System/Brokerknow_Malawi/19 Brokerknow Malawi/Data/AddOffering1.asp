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
		function  UpdateCanTrade(theChk)
		{
			if (theChk.checked)
			{
				document.frmMain.elements("txtCanTrade").value = "1";
			}
			else
			{
				document.frmMain.elements("txtCanTrade").value = "0";
			}
				
			
		}
</script>
</head>

<body Class="Dialog">

<!--#include file="../libroutines.asp"-->
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>

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
       dim cantrade
       
       immob = Request.Form("txtImmobilised")
       cantrade = Request.Form("txtCanTrade")
       code = Request.Form("txtCode")
       name = Request.Form("txtName")
	   addr = Request.Form("txtAddr")
       mktPrice = Request.Form("txtMktPrice")
       adate = Request.Form("txtADate")
       fee = Request.Form("txtFee")       
       secType =3      
       
        'validate Name
        If Trim(Name) = "" Then%>
				<script language="vbscript">
                						ShowMessage "Please specify the Name"                						
								</script>
				<% response.end
        End If
        
        'validate Market Price
        If Trim(mktPrice) = "" Then%>
				<script language="vbscript">
                						ShowMessage "Please specify the Market Price"
                						
								</script>
				<% response.end
        End If
        'validate size of Address
        If Len(Addr) > 100 Then%>
				<script language="vbscript">
								ShowMessage "Address can only be 100 characters in length"
								
								</script>
				<% response.end
        End If
        'validate size of Code
        If Len(Code) > 10 Then%>
				<script language="vbscript">
								ShowMessage "Code can only be 10 characters in length"
								
								</script>
				<% response.end
						End If
        'ensure Market Price is numeric
        If (MktPrice <> "") And (Not IsNumeric(MktPrice)) Then%>
			<script language="vbscript">
							ShowMessage "Market Price  must be numeric"
							
							</script>
			<% response.end
        End If
        'validate size of Name
        If Len(Name) > 100 Then%>
			<script language="vbscript">
							ShowMessage "Name can only be 100 characters in length"
							
							</script>
			<% response.end
        End If
     	'validate Security Type
        If Trim(secType) = "" Then%>
				<script language="vbscript">
                						ShowMessage "Please specify the Security Type"
                						
								</script>
				<% response.end
        End If
		'validate detail info
				'validate Transfer Fee
                If Trim(Fee) = "" Then%>
                		<script language = 'vbscript'>
                        ShowMessage "Please specify the Fee"
                        
                		</script>
                		<% response.end
                End If
                'ensure Transfer Fee is numeric
                If (Fee <> "") And (Not IsNumeric(Fee)) Then%>
                		<script language = 'vbscript'>
                        ShowMessage "Fee must be numeric"
                        
                		</script>
                		<% response.end
                End If
              
                
        'save header
        set guid = server.createobject("NDUtils.CGUID")
        guidStr = guid.GenerateGUID
        
        sqlStr = "INSERT INTO [SecurityOffering] (SecurityAddr,SecurityCode,SecurityMktPrice,SecurityName" & _
                "       ,Security_DPA_,OrderSecType_DPA_,Immobilised,CanTrade,Security_EIT_) SELECT " & "'" & addr & "'" & " as SecurityAddr," & "'" & Code & "'" & " as SecurityCode" & _
                "       ," & " " & MktPrice & " " & " as SecurityMktPrice" & _
                "       ," & "'" & Name & "'" & " as SecurityName," & " " & "iif(isnull(max([Security_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'SecurityOffering'),max([Security_DPA_]) + 1)" & " " & " as Security_DPA_" & _
                "		," & " " & secType & " " & " as OrderSecType_DPA_" & _
                "		," & " " & immob & " " & " as Immobilised" & _
                "		," & " " & cantrade & " " & " as CanTrade" & _
                "       ," & "'" & guidStr & "'" & " as Security_EIT_ FROM [SecurityOffering]"
        
        
        'Response.Write(sqlStr)
        'Response.End 
        
        Set conn = GetActiveConnection("KBroker")
        
        conn.BeginTrans
			conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
	        
			'obtain header key value
			sqlStr = "SELECT [SecurityOffering.Security_DPA_] FROM [SecurityOffering] WHERE [SecurityOffering.Security_EIT_] = " & "'" & guidStr & "'"
	        
			Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			If (rs.EOF Or rs.BOF) Then%>
					<script language = 'vbscript'>
                			ShowMessage "A serious error has been encountered while saving the data. Try saving again"
                			
					</script>
					<% response.end
			End If
	        
			'save detail data
			 sqlStr = "INSERT INTO [SecTransFee] (SecTransFeeADate,SecTransFeeFee,SecTransFee_DPA_" & _
                "                       ,Security_DPA_) SELECT " & "#" & adate & " " & Time & "#" & " as SecTransFeeADate" & _
                "       ," & " " & fee & " " & " as SecTransFeeFee" & _
                "       ," & " " & "iif(isnull(max([SecTransFee_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'SecTransFee'),max([SecTransFee_DPA_]) + 1)" & " " & " as SecTransFee_DPA_" & _
                "       ," & " " & rs.Fields("Security_DPA_") & " " & " as Security_DPA_" & _
                "        FROM [SecTransFee]"
	        
			conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
		conn.CommitTrans
        Set Conn = Nothing
        WritefraEnabledDialogCloseScript
        Response.End 
   	end If
   	
%>

<form name = 'frmAddSecurity' method = 'post' id="frmMain" action = "AddOffering.asp" >
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
                <td width="30%">Market Price</td>
                <td width="70%"><input type="text" name="txtMktPrice" id="txtMktPrice" size="25"></td>
              </tr>
              <tr>              
     <tr>
  <td width="30%">Activation Date</td>
  <td width="70%">
		<SCRIPT language="JavaScript">			
			var cal=new ctlSpiffyCalendarBox("cal", "frmAddSecurity", "txtADate","cmdDate","<%= FormatDate(Date) %>",1);
			cal.writeControl();
		</SCRIPT>
		
		</td>
  </tr>          
 <tr>
    <td>Can Trade</td>
    <td><input type=checkbox Class="BorderLess"   value='False' name='chkCanTrade' onClick = 'UpdateCanTrade(this);'> 
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
     <td width="100%" colspan="2" align=right>
		<BR>
		<BR>
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
		&nbsp;&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
		<input type = 'hidden' name ='txtImmobilised' id = 'txtImmobilised' value='0'>
		<input type = 'hidden' name ='txtCanTrade' id = 'txtCanTrade' value='0'>
      </td>
  </tr>
</table>
</form>


</body>

</html>
