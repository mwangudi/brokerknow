<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Bond Issue</title>

  <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>
 
  <!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">


</head>

<body Class="Dialog">
<!--#include file="../libroutines.asp"-->
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<%
	
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim ID
	Dim rsEdit
	
	action = ucase(Request.Form("action"))
	ID = Request("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		
                </script>
                <% WriteDialogRefuseOpenScript
                response.end
        End If
            
	if action = "EXECUTE" then	    
		Dim security
		Dim number
		Dim idate
		Dim mdate		
		Dim life
		Dim pay
		Dim rate
		Dim facevalue
		Dim determination
		
        
       security = Request.Form("cboSecurity") 
       number = Request.Form("txtIssue")
       idate = Request.Form("txtIDate")
       mdate = Request.Form("txtIDate1")       
       life = Request.Form("txtLife")
       pay = Request.Form("cboPayments")
       rate = Request.Form("txtRate")
       facevalue = Request.Form("txtface")
       determination = Request.Form("cbodetermination")
              
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
        'validate size of Coupon Payment
        If Len(pay) > 20 Then%>
                <script language = 'vbscript'>
                ShowMessage "Coupon Payment can only be 20 characters in length"
                
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
        'validate Life
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
        
        'update data
        Set conn = GetActiveConnection("KBroker")
        
        'save data
       sqlStr = "UPDATE [Bond] SET Determination = " & "'" & determination & "'" & ",FaceValue=" & " " & facevalue & " " & ",BondMDate = " & "#" & FormatDate(mdate) & "#" & ", BondIDate = " & "#" & FormatDate(idate) & "#" & ",BondIssue = " & "'" & number & "'" & "" & _
                "       ,BondLife = " & " " & life & " " & ",BondPayment = " & "'" & pay & "'" & "" & _
                "       ,BondRate = " & " " & rate & " " & ",Security_DPA_ = " & " " &  security & ", " & _
                "ModifiedBy = " & session("UserID") & " ,DateModified = #" & FormatDate(now()) & "# WHERE Bond_DPA_  = " & ID
        
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
        
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript
        Response.End
   	end If
   	
   	'Fetch Bond Issue Data
   	Set conn = GetActiveConnection("KBroker")
   
		sqlStr = "SELECT Security.SecurityCode, Security.SecurityName, Bond.BondIssue, Bond.BondMDate,Bond.FaceValue,Bond.Determination," & _		" Bond.BondIDate, Bond.BondLife, Bond.BondPayment, Bond.BondRate, Security.Security_DPA_" & _
		" FROM Bond INNER JOIN Security ON Bond.Security_DPA_ = Security.Security_DPA_" & _
		" WHERE Security.OrderSecType_DPA_=1 AND Bond_DPA_  = " & ID

        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected Bond Issue cannot be retrieved for editing"
                		
                </script>
                <% WriteDialogRefuseOpenScript
                response.end
        End If
%>

<form name = 'frmEditBond' method = 'post' action = 'EditIssue.asp' >
<table border="0" width="100%" cellspacing="1" cellpadding="1">
  <tr>
    <td width="40%">Bond</td>
    <td width="60%"><select name = 'cboSecurity' id = 'cboSecurity' size="1">
<%
        'List only fixed securities
        sqlStr = "SELECT * FROM [SecurityList] WHERE OrderSecTypeDescription = 'F'"
        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsEdit.EOF Or rsEdit.BOF) Then
                rsEdit.MoveFirst
                Do Until rsEdit.EOF
                		if rsEdit.Fields("Security_DPA_") = rs.Fields("Security_DPA_") Then%>
                			<option selected value = '<%=rsEdit.Fields("Security_DPA_")%>'><%=rsEdit.Fields("SecurityName")%></option>
                		<%else%>
                        <option value = '<%=rsEdit.Fields("Security_DPA_")%>'><%=rsEdit.Fields("SecurityName")%></option>
                     <%end if
						rsEdit.MoveNext
                Loop
        End If
%>

    </select></td>
  </tr>
<script language="JavaScript" src="../CALENDAR/calendar.js"></script>
<SCRIPT language="JavaScript">
	var calIDate=new ctlSpiffyCalendarBox("calIDate", "frmEditBond", "txtIDate","cmdIDate","<%= FormatDate(rs.Fields("BondIDate")) %>",1);
</SCRIPT>
<!--END CALENDAR -->
  <tr>
    <td width="40%">Issue Number</td>
    <td width="60%"><input type = 'text' name ='txtIssue' id = 'txtIssue' size="20" value = '<%=rs.Fields("BondIssue")%>'></td>
  </tr>
  <tr>
    <td width="40%">Face Value</td>
    <td width="60%"><input type = 'text' name ='txtface' id = 'txtface' size="20"value='<%=rs("FaceValue")%>'></td>
  </tr>    
  <tr>
    <td width="40%">Issue Date</td>
    <td width="60%">
		<SCRIPT language="JavaScript">
			var calIDate=new ctlSpiffyCalendarBox("calIDate", "frmEditBond", "txtIDate","cmdIDate","<%= FormatDate(rs.Fields("BondIDate")) %>",1);
			calIDate.writeControl();
		</SCRIPT>
	</td>
  </tr>
  <tr>
    <td width="40%">Maturity Date</td>
    <td width="60%">
		<SCRIPT language="JavaScript">
			var calIDate=new ctlSpiffyCalendarBox("calIDate", "frmEditBond", "txtIDate1","cmdIDate","<%= FormatDate(rs.Fields("BondMDate")) %>",1);
			calIDate.writeControl();
		</SCRIPT>
	</td>
  </tr>
  <tr>
    <td width="40%">Coupon Determination</td>
    <td width="60%">
		<select name="cboDetermination">		     
			<option selected value='<%=rs.Fields("Determination")%>'><%=rs.Fields("Determination")%></option>						
			<% if(trim(rs("determination"))="Fixed") then
			%>
			<option value='Floating'>Floating</option>			
			<%			
			else			
			%>
			<option value='Fixed'>Fixed</option>			
			<%
			end if
			%>
		</select>
    </td>
  </tr>  
  <tr>
    <td width="40%">Life</td>
    <td width="60%"><input type = 'text' name ='txtLife' id = "txtLife" size="20" value = '<%=rs.Fields("BondLife")%>'></td>
  </tr>
  <tr>
    <td width="40%">Coupon Payments</td>
    <td width="60%">
		<select name="cboPayments">
			<option selected value='<%=rs.Fields("BondPayment")%>'><%=rs.Fields("BondPayment")%></option>
			<option value='Monthly'>Monthly</option>
			<option value='Quarterly'>Quarterly</option>
			<option value='Biannually'>Biannually</option>
			<option value='Annually'>Annually</option>
		</select>
    </td>
  </tr>
  <tr>
    <td width="40%">Coupon Rates</td>
    <td width="60%"><input type = 'text' name ='txtRate' id = "txtRate" size="20" value = '<%=rs.Fields("BondRate")%>'></td>
  </tr>
  <tr>
  <tr>
   <td width="100%" colspan=2 align=right>
		<BR><BR><BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
		&nbsp;&nbsp;
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close()">
		&nbsp;&nbsp;
    	<input type = 'hidden' name ='action' id = 'action' value="Execute">
    	<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
    </td>
  </tr>
</table>
</form>
</body>

</html>














