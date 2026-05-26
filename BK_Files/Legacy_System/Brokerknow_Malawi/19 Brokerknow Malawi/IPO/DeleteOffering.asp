<html>

<head>
<title>Delete Security</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>

</head>

<body Class="Dialog">

<!--#include file="../libroutines.asp"-->
<%
	Dim conn 
	Dim sqlStr
	Dim rs
	
	Set conn = GetActiveConnection("KBroker")
    
	action = ucase(Request("delAction"))
	
	if action = "EXECUTE" then
			  
		ID = Request("ID")
		
		const LinkedIndependent = 1
		const LinkedDependent = 2
			
		If Trim(ID) = "" Then
			%>
		    <script language = 'vbscript'>
		    		ShowMessage "No record specified for deletion"
		    </script>
		    <%
		    response.end
		End If

        '----------------------------------

		'get voucher number
		sqlStr = "SELECT Offerings.PAL_No, Offerings.Client_DPA_, Offerings.Offering_Price, Offerings.Alloted_Rights, " & _
				  " Offerings.Offerings_Date, Security.SecurityName, Offerings.Receipt,Offerings.CitiAccepted, Security.BankAccount_DPA_," & _
				  " Offerings.Batch_No,isnull(Certificate,0) as Certificate,isnull(Offerings.CDSCharge,0) as CDSCharge " & _
                  " ,isnull(Security.CDSBankAccount_DPA_,0) as CDSBankAccount_DPA_ FROM Offerings INNER JOIN Security " & _
				  " ON Offerings.Offering = Security.Security_DPA_ " & _
				  " WHERE Offerings.Offering_DPA_ = " & ID	
				  
		'Response.write sqlStr
		'Response.End
		

		Set rs = conn.Execute(sqlStr)

		If   (rs.BOF Or rs.EOF) Then%>
		        <script language = 'vbscript'>
		        		ShowMessage "Serious error. The Application cannot be retrieved for deletion"
		        		window.self.close
		        </script>
		        <%response.end
		End If
				
		if isnull(rs.fields("Receipt")) then
				 Payment = 0
		else
				Payment = rs.fields("Receipt")
				'voucherType = 3
		end if
		
		'obtain values for journal
		clientID = rs.fields("Client_DPA_")
		oBank = rs.fields("BankAccount_DPA_")
		oDate = rs.fields("Offerings_Date")
		oQty = trim(rs.fields("Alloted_Rights"))
		oPrice = trim(rs.fields("Offering_Price"))
		oSecurity = iif(Len(trim(rs.fields("SecurityName")))>20,Left(trim(rs.fields("SecurityName")),20) & "...",trim(rs.fields("SecurityName")))
		oParticulars = FormatNumEx(oQty,0) & " " & oSecurity & " @" & FormatNumEx(oPrice,2) & " [ " & rs("Batch_No") & "]"
		IssueCertificate = trim(rs.fields("Certificate"))
		CDSCharge = trim(rs.fields("CDSCharge"))
		CDSBank = trim(rs.fields("CDSBankAccount_DPA_"))
		
        if(IssueCertificate=true) then
			TheCertificate = 1
        else
			TheCertificate = 0
        end if

		 procStr = "@userID = " & Session("UserID") & ",@clientDPA = " & clientID & ",@offerBank = " & oBank & ",@jDate = '" & oDate & "',@offerQty = " & oQty & ",@offerPrice = " & oPrice & ",@JournalNarrative = '" & oParticulars & "',@Certificate = " & TheCertificate '& ",@CDSCharge=" & CDSCharge & ",@CDSBank=" & CDSBank

		sqlStr = "update Offerings set deleted=1,Status=2, changedBy = " & Session("UserID") & ", TimeChanged = '"& Now() &"' WHERE Offering_DPA_ = " & ID
        sqlStr = SQLServerFormat(HandleQuote(sqlStr))
		
		'conn.begintrans
			
			 
		'Response.write procStr
		'Response.End
			Conn.execute("DeleteApplication " & procStr)
			Conn.execute(sqlStr)
		
			'conn.execute ("Exec ClientTotalProcedure " & clientID)							
			conn.execute ("Exec ClientBalanceProcedure " & clientID)
											
		'conn.CommitTrans

		'--------------------------------------
        
        Set Conn = Nothing
        WriteDeleteCloseScript
        Response.End
   	end If
%>