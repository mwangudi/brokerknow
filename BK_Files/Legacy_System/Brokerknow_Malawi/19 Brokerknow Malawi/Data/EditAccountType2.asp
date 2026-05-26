<!--#include virtual="libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "EditAccountType2"
	const DataEntity = "AccountType"
	const DataEntityPlural = "Account Types"
	const ActionFolder = "Data"
	
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim rsEdit
	Dim guid
	Dim guidStr
	
	action = ucase(Request.Form("action"))
	ID = Request("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		
                </script>
                <% response.end
        End If
        
	if action = "EXECUTE" then
		Dim buttonAction
		Dim reloadRequired
		
		reloadRequired = false
		buttonAction = Trim(Ucase(Request.Form("cmdAdd")))
		if buttonAction = "SAVE" then
				Dim parentType
				Dim Setting				
				
		        Quarter1 = Request.Form("txtQuarter1")
				Quarter2 = Request.Form("txtQuarter2")
				Quarter3 = Request.Form("txtQuarter3")
				Quarter4 = Request.Form("txtQuarter4")				
				
				if(trim(Quarter1)="") then
				Quarter1=0
				end if
				
				if(trim(Quarter2)="") then
				Quarter2=0
				end if
				
				if(trim(Quarter3)="") then
				Quarter3=0
				end if
				
				if(trim(Quarter4)="") then
				Quarter4=0
				end if

				Quarter1=Replace(Quarter1,",","")
				Quarter2=Replace(Quarter2,",","")
				Quarter3=Replace(Quarter3,",","")
				Quarter4=Replace(Quarter4,",","")				
				
				If Not IsNumeric(Quarter1 ) Then%>
						<script language = 'vbscript'>
						ShowMessage "Quarter1 can only be numeric"						
						</script>
						<% response.end
				End If
				
				If Not IsNumeric(Quarter2) Then%>
						<script language = 'vbscript'>
						ShowMessage "Quarter2 can only be numeric"						
						</script>
						<% response.end
				End If
				
				If Not IsNumeric(Quarter3) Then%>
						<script language = 'vbscript'>
						ShowMessage "Quarter3 can only be numeric"						
						</script>
						<% response.end
				End If
				
				If Not IsNumeric(Quarter4) Then%>
						<script language = 'vbscript'>
						ShowMessage "Quarter4 can only be numeric"						
						</script>
						<% response.end
				End If
				
				Quarter1=Ccur(Quarter1)
				Quarter2=Ccur(Quarter2)
				Quarter3=Ccur(Quarter3)
				Quarter4=Ccur(Quarter4)				
				
				'save data		
				Set conn = GetActiveConnection("KBroker")
						
				sqlStr = "UPDATE [AccountType] SET " & _
						"       Quarter1 = " & " " & Quarter1 & " " & _
						"       ,Quarter2 = " & " " & Quarter2 & " " & _
						"       ,Quarter3 = " & " " & Quarter3 & " " & _
						"       ,Quarter4 = " & " " & Quarter4 & " " & _
						"        WHERE AccountType_DPA_  = " & ID
				sqlStr = SQLServerFormat(HandleQuote(sqlStr))


				conn.Execute sqlStr
				
				conn.Close
				Set conn = Nothing
				WritefraEnabledDialogCloseScript
				Response.End
			end if
			Dim clientCode
        
			clientCode = "var validNavigate = true;" & chr(13)
			%>
			<script>
				<%=clientCode%>
			</script>
			<%
			response.End
   	end If
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit <%=DataEntity%></title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<script language="JavaScript" src="CALENDAR/calendar.js"></script>
<!--END CALENDAR -->

<script >
		var validNavigate = false;
		function ReleaseRecord()
		{
			if(!validNavigate)
			{
 				event.returnValue = "Please use the cancel button to close the dialog"
 			}
		}
		
		function AllowedNavigation()
		{
			validNavigate = true;
		}
		
	function formatnumber(theTxt)
	{	
	
	var theprice = theTxt.value;
		
		theprice =format2Number(format_number(theprice,2));
		
		theTxt.value=theprice;
	}
	
function format_number(p,d) 
	{
  	var r;
  	if(p<0)
  		{
  		p=-p;
  		r=format_number2(p,d);
  		r="-"+r;
  		}
  	else
  		{
  		r=format_number2(p,d);
  		}
  return r;
	}

function format_number2(pnumber,decimals) 
	{
  	var strNumber = new String(pnumber);
  	var arrParts = strNumber.split('.');
  	var intWholePart = parseInt(arrParts[0],10);
  	var strResult = '';
  	if (isNaN(intWholePart))
    intWholePart = '0';
  	if(arrParts.length > 1)
  		{
    	var decDecimalPart = new String(arrParts[1]);
    	var i = 0;
    	var intZeroCount = 0;
     	while ( i < String(arrParts[1]).length )
     		{
       		if( parseInt(String(arrParts[1]).charAt(i),10) == 0 )
       			{
         		intZeroCount += 1;
         		i += 1;
       			}
       		else
         	break;
    		}
    	decDecimalPart = parseInt(decDecimalPart,10)/Math.pow(10,parseInt(decDecimalPart.length-decimals-1)); 
    	Math.round(decDecimalPart); 
    	decDecimalPart = parseInt(decDecimalPart)/10; 
    	decDecimalPart = Math.round(decDecimalPart); 

    	//If the number was rounded up from 9 to 10, and it was for 1 'decimal' 
    	//then we need to add 1 to the 'intWholePart' and set the decDecimalPart to 0. 

    	if(decDecimalPart==Math.pow(10, parseInt(decimals)))
    		{ 
      		intWholePart+=1; 
      		decDecimalPart="0"; 
    		} 
    	var stringOfZeros = new String('');
    	i=0;
    	if( decDecimalPart > 0 )
    		{
      		while( i < intZeroCount)
      			{
        		stringOfZeros += '0';
        		i += 1;
      			}
    		}
    	decDecimalPart = String(intWholePart) + "." + stringOfZeros + String(decDecimalPart); 
    	var dot = decDecimalPart.indexOf('.');
    	if(dot == -1)
    		{
      		decDecimalPart += '.'; 
      		dot = decDecimalPart.indexOf('.'); 
    		} 
    	var l=parseInt(dot)+parseInt(decimals); 
    	while(decDecimalPart.length <= l) 
    		{
      		decDecimalPart += '0'; 
    		}
    	strResult = decDecimalPart;
  		}
  	else
  		{
    	var dot; 
    	var decDecimalPart = new String(intWholePart); 

    	decDecimalPart += '.'; 
    	dot = decDecimalPart.indexOf('.'); 
    	var l=parseInt(dot)+parseInt(decimals); 
    	while(decDecimalPart.length <= l) 
    		{
      		decDecimalPart += '0'; 
    		}
    	strResult = decDecimalPart;
  		}
  	return strResult;
	}

</script>
</head>
<%
Set conn = GetActiveConnection("KBroker")
sqlStr = "SELECT * FROM AccountTypeLevel2 WHERE AccountType_DPA_=" & ID
        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		window.self.ShowMessage "The selected <%=DataEntity%> cannot be retrieved for editing"
                		
                </script>
                <% response.end
        End If
%>
<body Class="Dialog">

<form name = 'frm<%=DataSource%>' method = 'post' action = '<%=DataSource%>.asp' id = "frmMain">
<table border="0" width="100%" cellspacing="1" cellpadding="1">
   <tr>
    <td >Account Type Name</td>
    <td ><input readonly = 'true' class=readonly type = 'text' name ='txtSetting' id = 'txtSetting' size="20" value = '<%=rs.Fields("AccountTypeName")%>'></td>
    <td >

	</td>
  </tr>
  <%
	Quarter1 = rs.Fields("Quarter1").value
	Quarter2 = rs.Fields("Quarter2").value
	Quarter3 = rs.Fields("Quarter3").value
	Quarter4 = rs.Fields("Quarter4").value

	if Quarter1 = "" then 
	Quarter1 = 0
	end if
	
	if Quarter2 = "" then 
	Quarter2 = 0
	end if

	if Quarter3 = "" then 
	Quarter3 = 0
	end if
	
	if Quarter4 = "" then 
	Quarter4 = 0
	end if
	%>

  <tr>
    <td nowrap> Quarter 1</td>
    <td nowrap><input type = 'text' name ='txtQuarter1' STYLE="TEXT-ALIGN: RIGHT;" id = "txtQuarter1" size="20" value='<%=formatnum(rs.Fields("Quarter1"))%>' onchange='format2Number(this);'></td>
  </tr>
  <tr>
    <td nowrap> Quarter 2</td>
    <td nowrap><input type = 'text' name ='txtQuarter2' STYLE="TEXT-ALIGN: RIGHT;" id = "txtQuarter2" size="20" value='<%=formatNum(rs.Fields("Quarter2"))%>' onchange='format2Number(this);'></td>
  </tr>
  <tr>
    <td nowrap> Quarter 3</td>
    <td nowrap><input type = 'text' name ='txtQuarter3' STYLE="TEXT-ALIGN: RIGHT;" id = "txtQuarter3" size="20" value='<%=FormatNum(rs.Fields("Quarter3"))%>' onchange='format2Number(this);'></td>
  </tr>
  <tr>
    <td nowrap> Quarter 4</td>
    <td nowrap><input type = 'text' name ='txtQuarter4' STYLE="TEXT-ALIGN: RIGHT;" id = "txtQuarter4" size="20" value='<%=FormatNum(rs.Fields("Quarter4"))%>' onchange='format2Number(this);'></td>
  </tr>  
  <tr>
	  <td width="100%" colspan=3 align="right" valign=absBottom>
		<BR><BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save" onclick = "AllowedNavigation()">
    	<input type = 'button' Class=Buttons name ='cmdCancel' id = "cmdCancel" value="Cancel" onclick = "JavaScript: window.self.close()">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ID' id = 'ID' value="<%= Rs.Fields("AccountType_DPA_").Value %>">
	</td>
  </tr>
</table>

</form>
</body>

</html>
