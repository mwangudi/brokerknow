<!--#include virtual="libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "BondsEvaluationList"
	const DataEntity = "BondsEvaluation"
	const DataEntityPlural = "BondsEvaluations"
	const ActionFolder = "Data"
	
	Dim UserId
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim rst
	Dim guid
	Dim guidStr
	Dim ID
	
	set rst = server.CreateObject ("ADODB.Recordset")
	
	UserId=Session("UserID")
	action = ucase(Request.Form("action"))
	ID = Request("ID")

    If Trim(ID) = "" Then%>
       <script language = 'vbscript'>
         ShowMessage "No record specified for editing"	
       </script>
                <% response.end
    End If
            
	if action = "EXECUTE" then
           
	        Dim ForwardRate
	        Dim DaysInCoupon
	        Dim Basis
	        Dim code
	        Dim Included
	        Dim CouponRate
	        
			ForwardRate = Cdbl(Request.Form("ForwardRate"))
			DaysInCoupon = Cdbl(Request.Form("DaysInCoupon"))
			Basis = Cdbl(Request.Form("Basis"))
			CouponRate = Cdbl(Request.Form("CouponRate"))
			ForwardRateID = Request.Form("ForwardRateID")
			
			if ucase(trim(Request.Form("Included"))) = "ON" then
			 Included = 1
			else
			 Included = 0
			end if
				
			'Validations
			 
			If trim(CouponRate) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Coupon Rate."
                window.self.close
                </script>
                <% response.end
            End If 
            
            If not isnumeric(trim(CouponRate)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Invalid Coupon Rate."
                window.self.close
                </script>
                <% response.end
            End If 
            
		if ForwardRateID <> "" then
			If trim(ForwardRate) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Forward Rate."
                window.self.close
                </script>
                <% response.end
            End If 
            
            If not isnumeric(trim(ForwardRate)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Invalid Forward Rate."
                window.self.close
                </script>
                <% response.end
            End If 
		end if
			
			If trim(DaysInCoupon) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the No of Days In Coupon."
                window.self.close
                </script>
                <% response.end
            End If 
            
            If not isnumeric(trim(DaysInCoupon))  Then%>
                <script language = 'vbscript'>
                ShowMessage "Invalid No of Days In Coupon."
                window.self.close
                </script>
                <% response.end
            End If 
			
			If trim(Basis) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Basis."
                window.self.close
                </script>
                <% response.end
            End If 
            
            If not isnumeric(trim(Basis))  Then%>
                <script language = 'vbscript'>
                ShowMessage "Invalid Basis."
                window.self.close
                </script>
                <% response.end
            End If 
            
			'save data
					
		     Set conn = GetActiveConnection("KBroker")
				 conn.BeginTrans		
						
				  IF ucase(mid(ID,1,3)) = "PRY" then		' Primary Issues	      
				    sqlStr = "Update [PrimaryIssues] set Basis = " & Basis & ", " & _
									"DaysInCoupon = " & DaysInCoupon & ", Included = " & Included & ", " & _ 
									"ModifiedBy = " & session("UserID") & " ,DateModified = #" & FormatDate(now()) & "# Where PrimaryIssues_DPA_=" & cint(mid(ID,4,len(ID) - 3))
									
				  elseif ucase(mid(ID,1,3)) = "CON" then 'Contract
		
					sqlStr = "Update [Contract] set Basis = " & Basis & ", " & _
									"DaysInCoupon = " & DaysInCoupon & ", Included = " & Included & ", " & _ 
									"ModifiedBy = " & session("UserID") & " ,DateModified = #" & FormatDate(now()) & "# Where Contract_DPA_ =" & cint(mid(ID,4,len(ID) - 3))
									
				  else ' External Trade
					sqlStr = "Update [BondTrades] set Basis = " & Basis & ", " & _
									"DaysInCoupon = " & DaysInCoupon & ", Included = " & Included & ",  " & _ 
									"ModifiedBy = " & session("UserID") & " ,DateModified = #" & FormatDate(now()) & "# Where BondTrades_DPA_=" & cint(mid(ID,4,len(ID) - 3))
					
				  end if
				  		 			
					sqlStr = SQLServerFormat(HandleQuote(sqlStr))
                     
					conn.Execute sqlStr
					
				  'Update Coupon Rate in Bonds Table
				   	
				   	sqlStr2 = " Update [Bond] set BondRate = " & CouponRate & " Where Bond_DPA_ IN " & _
				   	          " (Select Bond_DPA_ from BondsEvaluationList where Rtrim(Evaluation_DPA_) like '" & ID & "')"
				   	
				   	sqlStr2 = SQLServerFormat(HandleQuote(sqlStr2))
				   	
				   	conn.Execute sqlStr2
				   	
				  'Úpdate Forward Rate
				  
				  if ForwardRateID <> "" then
				   sqlStr3 = "Update [ForwardRate] Set ForwardRate = " & ForwardRate & ", ActivationDate = #" & now() & "#, " & _
									"ModifiedBy = " & session("UserID") & " , DateModified = #" & FormatDate(now()) & "#" & _
									" Where ForwardRate_DPA_ = " & ForwardRateID
				   sqlStr3 = SQLServerFormat(HandleQuote(sqlStr3))
				   	
				   	conn.Execute sqlStr3
				  end if
				   	
					conn.CommitTrans
					
					WritefraEnabledDialogCloseScript
		
				    conn.Close
					Set conn = Nothing
					
					Response.End	
   	End if
   	
 Set conn = GetActiveConnection("KBroker")
	sqlStr = "SELECT * FROM BondsEvaluationList WHERE RTrim(Evaluation_DPA_) like '" & ID & "'"
			        
	Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
    If rs.EOF Or rs.BOF Then%>
	   <script language = 'vbscript'>
	    window.self.ShowMessage "The selected <%=DataEntity%> cannot be retrieved for editing"             		
	   </script>
		<% response.end
	 End If
	   	
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit <%=DataEntity%></title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/accountList.js"></SCRIPT>
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<script language="JavaScript" src="CALENDAR/calendar.js"></script>
<!--END CALENDAR -->
</head>
<body Class="Dialog">
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<form name = 'frmBondsEvaluation' method = 'post' action = 'EditBondsEvaluation.asp' id = "frmBondsEvaluation" >
<br>&nbsp;
<table border="1" width="90%" cellspacing="0" cellpadding="0" bordercolor="#000000">
 <TR>
    <TD width="20%">&nbsp;Included</TD>
    <TD width="80%">
    <%
    
     If cint(iif(isnull(rs.Fields("Included")),0,rs.Fields("Included"))) = 1 then
     %><Input type="checkbox" name="Included" id="Included" checked readonly><%
     else
     %><Input type="checkbox" name="Included" id="Included" readonly> <%
     end if
    %>
    </TD></TR>
    <TR>
    <TD width="30%">&nbsp;Type</TD>
    <TD width="70%"><input Type="Text" name="Type" id="Type" class="readonly" value="<%=rs.Fields("Type")%>" readonly></TD>
    </TR>
    <TR>
    <TD width="30%">&nbsp;Reference</TD>
    <TD width="70%"><input Type="Text" name="Reference" class="readonly" id="Reference" value="<%=rs.Fields("Reference")%>" readonly></TD></TR>
  <TR>
    <TD width="30%">&nbsp;Date of Purchase</TD>
    <TD width="70%"><input Type="Text" name="PDate" id="PDate" class="readonly" value="<%=formatdate(rs.Fields("TradeDate"))%>" readonly></TD></TR>
  <TR>
    <TD width="30%">&nbsp;Client</TD>
    <TD width="70%"><input type="text" name="client" id="client" class="readonly" value="<%=rs.Fields("Client_DPA_") & " - " & rs.Fields("ClientName")%>" size="40" readonly></td>
  </tr>
  <TR>
    <TD width="30%">&nbsp;Bond Issue</TD>
    <TD width="70%"><input Type="BondIssue" name="BondIssue" id="Type" class="readonly" value="<%=rs.Fields("BondIssue")%>" readonly>
    </TD></TR>
  <TR>
    <TD width="30%">&nbsp;Issue Date</TD>
    <TD width="70%"><input Type="Text" name="IssueDate" id="IssueDate" class="readonly" value="<%=formatdate(rs.Fields("IssueDate"))%>" readonly></TD></TR>
  <TR>
    <TD width="30%">&nbsp;Maturity Date</TD>
    <TD width="70%"><input Type="Text" name="MaturityDate" id="MaturityDate" class="readonly" value="<%=formatdate(rs.Fields("MaturityDate"))%>" readonly></TD></TR>
  <TR>
    <TD width="30%">&nbsp;Coupon Rate</TD>
    <TD width="70%"><input type ="text" Name ="CouponRate"  id="CouponRate" size="20" Value="<%=formatnumber(rs.Fields("CouponRate"),2)%>"></TD></TR>
  <TR>
    <TD width="30%">&nbsp;Face Value</TD>
    <TD width="70%"><input type ="text" Name ="FaceValue"  id="FaceValue" class="readonly" size="20" Value="<%=formatnumber(rs.Fields("FaceValue"),2)%>" readonly></TD></TR>
    <TD width="30%">&nbsp;Forward Rate</TD>
    <TD width="70%">
    <input type ="hidden" Name ="ForwardRateID"  id="ForwardRateID" size="20" Value="<%=rs.Fields("ForwardRate_DPA_")%>" >
    <% if isnull(rs.Fields("ForwardRate_DPA_")) then%>
     <input type ="hidden" Name ="ForwardRate"  id="ForwardRate" size="20" Value="0">
     <input type ="text" Name ="ForwardRate1"  id="ForwardRate1" size="20" Value="Undefined"  class="readonly" readonly>
     
    <% else %>
     <input type ="text" Name ="ForwardRate"  id="ForwardRate" size="20" Value="<%=formatnumber(iif(isnull(rs.Fields("ForwardRate")),0,rs.Fields("ForwardRate")),2)%>" >
    <%End if%>
   
    </TD></TR>
  <TR>
    <TD width="30%">&nbsp;Buying Price</TD>
    <TD width="70%"><input type ="text" name="BPrice" id ="BPrice" class="readonly" value="<%=formatnumber(rs.Fields("BPrice"),4)%>" readonly></TD></TR>
  <%
    dim BasisValue
    dim DaysValue
  
    if Cint(iif(isnull(rs.Fields("Basis")),0,rs.Fields("Basis"))) = 0 then
     BasisValue = 364
    else
     BasisValue = rs.Fields("Basis")
    end if
    
    if Cint(iif(isnull(rs.Fields("DaysInCoupon")),0,rs.Fields("DaysInCoupon"))) = 0 then
     DaysValue = 182
    else
     DaysValue = rs.Fields("DaysInCoupon")
    end if
     
  %>
  <TR>
    <TD width="30%">&nbsp;No. Of Days In Coupon</TD>
    <TD width="70%"><input type ="text" Name ="DaysInCoupon"  id="DaysInCoupon" size="20" value="<%=DaysValue%>"></TD></TR>
  
  <TR>
    <TD width="30%">&nbsp;Basis</TD>
    <TD width="70%"><input type ="text" Name ="Basis"  id="Basis" size="20" value="<%=BasisValue%>"></TD></TR>
 
  <tr>
	  <td width="100%" colspan=2 align="center" valign=absBottom>
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdSave' id = 'cmdSave' value=" Save ">
    	<input type = 'button' Class=Buttons name ='cmdCancel' id = "cmdCancel" value="Cancel" onclick = "JavaScript: window.self.close()">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
	</td>
  </tr>
</table>

</form>
</body>

</html>