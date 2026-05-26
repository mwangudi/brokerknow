<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Gender</title>
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
<script language='javascript'>
		function forceSubmit()
		{
			setOpener();
			document.frmEditGender.method='post';
			document.frmEditGender.target='_self';
			document.frmEditGender.submit();		
		}
		
		function setOpener()
		{
			window.self.opener = window.dialogArguments.opener;					
		}
</script>

</head>

<body Class="Dialog" onload="javascript: setOpener(); ">

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

	If Trim(ID) = "" Then%>
            <script language = 'vbscript'>
            		ShowMessage "No record specified for editing"
                		
            </script>
            <% response.end
    End If

	if action = "EXECUTE" then
		Dim description
        
        Description = Request.Form("txtCompanyName")
        phonenumber = Request.Form("txtPhoneNumber")
        faxnumber = Request.Form("txtFaxNumber")
        caddress = Request.Form("txtAddress")
        city = Request.Form("txtCity")
        country = Request.Form("txtCountry")
        branchid = Request.Form("cboBranch")
        photo = Request.Form("txtPhoto")   
		stateorprovince = Request.Form("txtStateProvince") 
		PostalCode = Request.Form("txtPostalCode") 
		Broker_DPA_ = Request.Form("cboBroker") 
       
        'validate Description
        If Trim(Description) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Description"
                		
                </script>
                <% reloadpage(ID)
				response.end
        End If
		
		'validate Phone number
        If Trim(phonenumber) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Phone Number"                		
                </script>
                <%reloadpage(ID)
				response.end
        End If
        
        'validate Address
        If Trim(caddress) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Address"                		
                </script>
                <% reloadpage(ID)
				response.end
        End If  
        
        'validate City
        If Trim(city) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the City"                		
                </script>
                <% response.end
        End If 

		 'validate Country
        If Trim(country) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Country"                		
                </script>
                <% reloadpage(ID)
				response.end
        End If 


		Set conn = GetActiveConnection("KBroker")
       
        'save data
        
        If Photo <> "" Then
			photoNew = Photo
		Else
			PhotoNew = "/Data/Photos/_blank.jpg"	
		End If
			    
        
        sqlStr = "UPDATE [CompanyInfo] SET CompanyName = '" & Description & "', " & _
				 " PhoneNumber = '" & phonenumber & "', " & _
				 " FaxNumber = '" & faxnumber & "', " & _	
				 " Address = '" & caddress & "', " & _	
				 " City = '" & city & "', " & _	
				 " Country = '" & country & "', " & _	
				 " BranchID = '" & branchid & "', " & _	
				 " Photo = '" & photoNew & "', " & _	
				 " PostalCode = '" & PostalCode & "', " & _
				 " Broker_DPA_ = " & Broker_DPA_ & ", " & _		
				 " StateOrProvince = '" & stateorprovince & "' " & _				 		
				 " WHERE SetupID  = " & ID                
				 
		
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
        
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript2
   	end If
%>

<form name = 'frmEditGender' method = 'post' action = 'EditCompanyInfo.asp' >
<table border="0" width="100%" cellspacing="1" cellpadding="1">
<%
        Set conn = GetActiveConnection("KBroker")
        
        
        sqlStr = "SELECT * FROM [CompanyInfo] WHERE SetupID  = " & ID
        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected set-up information cannot be retrieved for editing"
                		
                </script>
                <% response.end
        End If
     

%>
  <tr>
    <td width="40%"> Company Name</td>
    <td width="70%"><input type = 'text' name ='txtCompanyName' id = 'txtCompanyName' size="25" value = '<%=rs.Fields("CompanyName")%>'></td>
  </tr>
  <tr>
    <td width="40%"> Phone Number</td>
    <td width="70%"><input type = 'text' name ='txtPhoneNumber' id = 'txtPhoneNumber' size="25" value = '<%=rs.Fields("PhoneNumber")%>'></td>
  </tr>
  <tr>
    <td width="40%"> Fax Number</td>
    <td width="70%"><input type = 'text' name ='txtFaxNumber' id = 'txtFaxNumber' size="25" value = '<%= rs.Fields("FaxNumber") %>'></td>
  </tr>  
  <tr>
    <td width="40%"> Address </td>
    <td width="70%"><input type = 'text' name ='txtAddress' id = 'txtAddress' size="25" value = '<%=rs.Fields("Address")%>'></td>
  </tr>
  <tr>
    <td width="40%"> PostalCode</td>
    <td width="70%"><input type = 'text' name ='txtPostalCode' id = 'txtPostalCode' size="25" value = '<%=rs.Fields("PostalCode")%>'></td>
  </tr>
  <tr>
    <td width="40%"> City</td>
    <td width="70%"><input type = 'text' name ='txtCity' id = 'txtCity' size="25" value = '<%=rs.Fields("City")%>'></td>
  </tr>
  <tr>
    <td width="40%"> State/Province</td>
    <td width="70%"><input type = 'text' name ='txtStateProvince' id = 'txtStateProvince' size="25" value = '<%=rs.Fields("StateOrProvince")%>'></td>
  </tr>
  
  <tr>
    <td width="40%"> Country</td>
    <td width="70%"><input type = 'text' name ='txtCountry' id = 'txtCountry' size="25" value = '<%= rs.Fields("Country") %>'></td>
  </tr>
  <tr>
    <td width="40%"> Company Logo</td>
    <td width="70%">
    <INPUT name="txtPhoto" ID="txtPhoto" type="hidden" value="<%= rs.Fields("Photo") %>">
    <IFRAME FRAMEBORDER=0 SCROLLING=NO SRC="upload.asp" width="170px" height="170px"></IFRAME>		
    </td>
  </tr>
  
  <tr>
    <td width="100%" colspan=2><HR width="100%"></td>    
  </tr> 
  <tr>
    <td width="40%"> Branch </td>
    <td width="70%">
    <select name = 'cboBranch' id = 'cboBranch' size="1">
		<%
		        Set conn = GetActiveConnection("KBroker")       
		        
		        sqlStr = "SELECT * FROM [BranchList]"
		        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		        If Not (rsEdit.EOF Or rsEdit.BOF) Then
		                rsEdit.MoveFirst
		                Do Until rsEdit.EOF
		                		if rsEdit.Fields("Branch_DPA_") = rs.Fields("BranchID") Then%>
		                			<option selected value = '<%=rsEdit.Fields("Branch_DPA_")%>'><%=rsEdit.Fields("BranchName")%></option>
		                		<%else%>
		                        <option value = '<%=rsEdit.Fields("Branch_DPA_")%>'><%=rsEdit.Fields("BranchName")%></option>
		                     <%end if
								rsEdit.MoveNext
		                Loop
		        End If
		%>

    </select>
    
    </td>
  </tr>
   <tr>
    <td width="40%"> Brokerage Identity </td>
    <td width="70%">
    <select name = 'cboBroker' id = 'cboBroker' size="1">
		<%
		        Set conn = GetActiveConnection("KBroker")       
		        
		        sqlStr = "SELECT * FROM [Broker] ORDER BY BrokerName"
		        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		        If Not (rsEdit.EOF Or rsEdit.BOF) Then
		                rsEdit.MoveFirst
		                Do Until rsEdit.EOF
		                		if rsEdit.Fields("Broker_DPA_") = rs.Fields("Broker_DPA_") Then%>
		                			<option selected value = '<%=rsEdit.Fields("Broker_DPA_")%>'><%= rsEdit.Fields("BrokerName") %></option>
		                		<%else%>
		                        <option value = '<%=rsEdit.Fields("Broker_DPA_").Value %>'><%= rsEdit.Fields("BrokerName").Value %></option>
		                     <%end if
								rsEdit.MoveNext
		                Loop
		        End If
		Set rsEdit = Nothing
		%>	
			

    </select>
    
    </td>
  </tr>
  <tr>
    <td width="100%" colspan=2 align=right>		
		<BR>
	<input type = 'button' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save " onclick="javascript:forceSubmit();">
	&nbsp;&nbsp;
	<input type = 'button' Class=Buttons name ='cmdClose' id = 'cmdClose' value=" Cancel " OnClick="JavaScript: window.self.close();">
    	<input type = 'hidden' name ='action' id = 'action' value="Execute">
    	<input type = 'hidden' name ='ID' id = 'ID' value="<%= ID %>">
    </td>
  </tr>
</table>
</form>
<%
Set Rs = Nothing
Set Conn = Nothing
%>

</body>

</html>
