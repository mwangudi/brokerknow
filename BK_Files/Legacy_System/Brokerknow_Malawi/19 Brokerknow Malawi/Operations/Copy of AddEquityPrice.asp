<!--#include virtual="libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "EquityList"
	const DataEntity = "MarketPrice"
	const DataEntityPlural = "Market Prices"
	const ActionFolder = "Operations"
	
	Dim UserId
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim guid
	Dim guidStr
	Dim ID
    

	ID = Request("ID")

	If Trim(ID) = "" Then%>
            <script language = 'vbscript'>
                	ShowMessage "No record specified for editing"
            </script>
            <% response.end
    End If
	
	UserId=Session("UserID")
	action = ucase(Request.Form("action"))
	
	if action = "EXECUTE" then
	        
		    Dim code
			Dim MarketPrice
			Dim MarketDate
			
			code = Request.Form("code")
			MarketPrice = Request.Form("MarketPrice")
			MarketDate = Request.Form("txtDate")
			Security_DPA = Request.Form("ID")
			
			if  MarketDate = "" then
			 %><script language = 'Javascript'>
					alert('Please specify the Market Date.')
					window.history.back()
				</script><% response.end
			end if     
			
			if  not isDate(MarketDate) then
			 %><script language = 'Javascript'>
					alert('Invalid Market Date.')
					window.history.back()
				</script><% response.end
			end if  
			
			if  MarketPrice = "" then
			 %><script language = 'Javascript'>
					alert('Please specify the Market Price.')
					window.history.back()
				</script><% response.end
			end if 
			
			if  not isnumeric(MarketPrice) then
			 %><script language = 'Javascript'>
					alert('Invalid Market Price.')
					window.history.back()
				</script><% response.end
			end if 
			
			'save data
					
		     Set conn = GetActiveConnection("KBroker")
		  
			 conn.BeginTrans	
			 
			 'Check that there is no any other market price for the specified Date 
	         
			    sqlstr = " select * from EquityList where Security_DPA_ = " & Security_DPA & _
				         " AND [Date] = #" & MarketDate & " # "
	         
				Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
	            
			    if not (rs.bof or rs.eof) then
			       
			       if not isnull(rs.fields("Price")) then
						%>
						   <script language = 'vbscript'>
						      window.self.ShowMessage "The selected date already has a market price."     		
						   </script>
						<%
						 
						response.end
			       end if
			       
			    end if	
			    
				'Add Record to Datastream Security if non-existant
				sqlstr = " select * from datastream_securities where SecKnow_DPA = " & Security_DPA & " "
	         
				Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
	         
				if (rs.BOF or rs.EOF) then
					'Security Non - Existant in Datastream Security table. So add it.
					 sqlStr = "INSERT INTO [Datastream_Securities] (SecFirstListing, SecKnow_DPA, SecCode) " & _
							  " Values  (" & "#" & FormatDate(MarketDate) & "#, " & Security_DPA & ", '" & code & "')"
							 	  
					 sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
					 
					 conn.Execute sqlStr
				end if
	         
					'Add record to Datastream Market
					sqlStr = "INSERT INTO [Datastream_Market] (MktUnique, MktClose, MktCode, MktDate) " & _
							  " SELECT " & " " & "iif(isnull(max([MktUnique])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Datastream_Market'),max([MktUnique]) + 1)" & " " & " as MktUnique" & _
					         "," & " " & MarketPrice & " " & " as MktClose  " & _
							  "," & "'" & code & "'" & " as MktCode" & _
							  "," & "#" & FormatDate(MarketDate) & "#" & " as MktDate" & _
							  " FROM [Datastream_Market]"
							  
					sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
           
				conn.Execute (sqlStr)
			 	
				conn.CommitTrans
					
			 WritefraEnabledDialogCloseScript
		
			 conn.Close
			 Set conn = Nothing
					
			 Response.End	
   	End if
   	
Set conn = GetActiveConnection("KBroker")

'Retrieve Security_DPA_ AND Market Date from the passed ID value

Security_DPA = mid(ID,1,instr(1,ID,"-")-1)
MktDate = mid(ID,instr(1,ID,"-")+1,len(ID)-1)

sqlStr = "SELECT * FROM " &  DataSource & " WHERE Security_DPA_=" & Security_DPA & " AND [Date] = #" & MktDate & "#"
        
Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))

If rs.EOF Or rs.BOF Then%>
    <script language = 'vbscript'>
       window.self.ShowMessage "The selected <%=DataEntity%> cannot be retrieved."     		
   </script>
 <%  
 response.end
else
	if not isnull(rs.fields("Price")) then
	   if FormatDate(rs.fields("Date")) = FormatDate(Now()) then
			%>
			   <script language = 'vbscript'>
			      window.self.ShowMessage "The selected Security already has market price."     		
			  </script>
			<%  WriteDeleteCloseScript
			response.end
		end if
	end if
 
End If
         	
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add <%=DataEntity%></title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/accountList.js"></SCRIPT>
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<script language="JavaScript" src="CALENDAR/calendar.js"></script>
<!--END CALENDAR -->


<script language='javascript'>
		var validNavigate = false;
	
	function UpdateCode(){
	 var frm = document.frmMain ;
	 
	 frm.code.value = frm.cboSecurity[frm.cboSecurity.selectedIndex].SecCode;
	 
	}	
		
    function FormatNumber2(Obj){
	 var theVal = Obj.value
	 Obj.value = formatNum(theVal)
	}
	
	function FormatNumber(Obj,decimals){
	 var theVal = Obj.value
	 Obj.value = format_number2(theVal,decimals)
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
<body Class="Dialog">
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "txtDate","cmdDate","<%=FormatDate(now())%>",1);
</SCRIPT>

<form name = 'frmMain' method = 'post' action = 'AddEquityPrice.asp' id = "frmMain" >
<table border="0" width="100%" cellspacing="1" cellpadding="1">
 <tr>
    <td width="15%">Code</td>
    <td width="54%">
    <input type="text" name="code" id="code" value="<%=rs.fields("SecurityCode")%>" size="20" class="readonly" readonly></td>
  </tr>
 <tr>
    <td width="20%">Security</td>
    <td width="54%">
    <input type="text" name="Security" id="Security" value="<%=rs.fields("SecurityName")%>" size="40" 
    style= "BORDER-RIGHT: silver 1px ridge;BORDER-TOP: silver 1px solid;FONT-WEIGHT: normal;FONT-SIZE: 8pt;BORDER-LEFT: silver 1px solid;WIDTH: 300px;COLOR: navy; BORDER-BOTTOM: silver 1px outset;FONT-FAMILY: verdana, arial, helvetica, sans-serif;BACKGROUND-COLOR: #c0c0c0" 
    readonly></td>
  
  </tr>
  
  <tr >
	<td width="15%" valign="top">Market Date</td>
	<td ><SCRIPT language="JavaScript">cal.writeControl();</SCRIPT></td>
  </tr>
  <tr>
    <td width="15%">Market Price</td>
    <td width="54%">
    <input type="text" name="MarketPrice" id="MarketPrice" value="" size="20" onChange="Javascript: FormatNumber(this,4)"></td>
  </tr>
  <tr>
	  <td width="100%" colspan=2 align="center" valign=absBottom>
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdSave' id = 'cmdSave' value=" Save ">
    	<input type = 'button' Class=Buttons name ='cmdCancel' id = "cmdCancel" value="Cancel" onclick = "JavaScript: window.self.close()">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ID' id = 'ID' value="<%=Security_DPA%>">
	</td>
  </tr>
</table>

</form>
</body>

</html>