<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Bond Issue</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<script language="JavaScript" src="CALENDAR/calendar.js"></script>

<!--END CALENDAR -->
</head>

<body Class="Dialog">

<!--#include file="../libroutines.asp"-->
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<%
	
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	
	action = ucase(Request.Form("action"))
	
	if action = "EXECUTE" then
		Dim security
		Dim number
		Dim idate
		Dim mdate		
		Dim life
		Dim pay
		Dim rate
		dim facevalue
		dim determination
		
       security = Request.Form("cboSecurity") 
       number = Request.Form("txtIssue")
       idate = Request.Form("txtIDate")
       mdate = Request.Form("txtMDate")
       sdate = Request.Form("txtSDate")
       life = Request.Form("txtLife")
       pay = Request.Form("cboPayments")
       rate = Request.Form("txtRate")
	   facevalue=Request.Form("txtface")	
	   determination=Request.Form("cboDetermination")	
	   
       'validate Security
        If Trim(Security) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Security"
                
                </script>
                <% response.end
        End If
        'validate Issue Number
        If Trim(number) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Issue Number"
                
                </script>
                <% response.end
        End If
        'validate size of Issue Number
        If Len(number) > 20 Then%>
                <script language = 'vbscript'>
                ShowMessage "Issue Number can only be 20 characters in length"
                
                </script>
                <% response.end
        End If
        'validate Life
        If Trim(life) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Life"
                
                </script>
                <% response.end
        End If
        'ensure Life is numeric
        If (life <> "") And (Not IsNumeric(life)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Life  must be numeric"
                
                </script>
                <% response.end
        End If
        'validate Coupon Payment
        If Trim(pay) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Coupon Payment"
                
                </script>
                <% response.end
        End If
        'validate Coupon Rate
        If Trim(rate) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Coupon Rate"
                
                </script>
                <% response.end
        End If
        'ensure Coupon Rate is numeric
        If (rate <> "") And (Not IsNumeric(rate)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Coupon Rate  must be numeric"
                
                </script>
                <% response.end
        End If   
        
        If Trim(determination) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Coupon determination"
                
                </script>
                <% response.end
        End If
        If Trim(facevalue) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Face Value"
                
                </script>
                <% response.end
        End If
        'ensure face value is numeric
        If (facevalue <> "") And (Not IsNumeric(facevalue)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Face Value  must be numeric"
                
                </script>
                <% response.end
        End If    
        'save data
        sqlStr = "INSERT INTO [Bond] (BondMDate, BondIDate,BondIssue,Determination,FaceValue,BondLife,BondPayment,BondRate,Bond_DPA_ " & _
                "       ,Security_DPA_) SELECT " & "#" & FormatDate(mdate) & "#" & " as BondMDate," & "#" & FormatDate(idate) & "#" & " as BondIDate," & "'" & number & "'" & " as BondIssue" & _
                "       ," & "'" & Determination & "'" & " as Determination," & " " & facevalue & " " & " as FaceValue," & " " & life & " " & " as BondLife," & "'" & pay & "'" & " as BondPayment" & _
                "       ," & " " & rate & " " & " as BondRate," & " " & "iif(isnull(max([Bond_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Bond'),max([Bond_DPA_]) + 1)" & " " & " as Bond_DPA_" & _
                "       ," & " " & security & " " & " as Security_DPA_" & _
                "        FROM [Bond]"
        Set conn = GetActiveConnection("KBroker")
        sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
        conn.BeginTrans
                conn.Execute sqlStr
        conn.CommitTrans
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript
        Response.End
   	end If
%>

<form name = 'frmAddIssue' method = 'post' action = 'AddIssue.asp' >
<table border="0" width="100%" cellspacing="1" cellpadding="1">
  <tr>
    <td width="40%">Bond</td>
    <td width="60%"><select name = 'cboSecurity' id = 'cboSecurity' size="1">
    	<option selected value = ''></option>
<%
        Set conn = GetActiveConnection("KBroker")
        
        sqlStr = "SELECT * FROM [SecurityList] WHERE OrderSecTypeDescription = 'F'"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF%>
                        <option value = '<%=rs.Fields("Security_DPA_")%>'><%=rs.Fields("SecurityName")%></option>
                        <%rs.MoveNext
                Loop
        End If
%>

    </select></td>
  </tr>
  <tr>
    <td width="40%">Issue No</td>
    <td width="60%"><input type = 'text' name ='txtIssue' id = 'txtIssue' size="20"></td>
  </tr>
  <tr>
    <td width="40%">Face Value</td>
    <td width="60%"><input type = 'text' name ='txtface' id = 'txtface' size="20"></td>
  </tr> 
  <tr>
    <td width="40%">Issue Date</td>
    <td width="60%">
		<SCRIPT language="JavaScript">
			var calIDate=new ctlSpiffyCalendarBox("calIDate", "frmAddIssue", "txtIDate","cmdIDate","<%= FormatDate(Date) %>",1);
			calIDate.writeControl();
		</SCRIPT>
	</td>
  </tr>
  <tr>
    <td width="40%">Maturity Date</td>
    <td width="60%">
		<SCRIPT language="JavaScript">
			var calMDate=new ctlSpiffyCalendarBox("calMDate", "frmAddIssue", "txtMDate","cmdMDate","<%= FormatDate(Date) %>",1);
			calMDate.writeControl();
		</SCRIPT>
	</td>
  </tr>
  <tr>
    <td width="40%">Coupon Determination</td>
    <td width="60%">
		<select name="cboDetermination">			
			<option selected SearchCode = "0" SearchText = "Fixed" value = 'Fixed'>Fixed</option>			
			<option value='Floating'>Floating</option>			
		</select>
    </td>
  </tr>    
  <tr>
    <td width="40%">Life</td>
    <td width="60%"><input type = 'text' name ='txtLife' id = "txtLife" size="20"></td>
  </tr>
  <tr>
    <td width="40%">Coupon Payments</td>
    <td width="60%">
		<select name="cboPayments">
			<option selected value=''></option>
			<option value='Monthly'>Monthly</option>
			<option value='Quarterly'>Quarterly</option>
			<option value='Semi-annually'>Semi-annually</option>
			<option value='Annually'>Annually</option>
		</select>
    </td>
  </tr>
  <tr>
    <td width="40%">Coupon Rates</td>
    <td width="60%"><input type = 'text' name ='txtRate' id = "txtRate" size="20"></td>
  </tr>
  <tr>
   <td colspan=2 align=right>
		<BR><BR><BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
		&nbsp;&nbsp;
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close()">
		&nbsp;&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
     </td>
  </tr>
 
</table>
</form>

</body>

</html>
