<!--#include virtual="libroutines.asp"-->
<%
	
	Dim UserId
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim guid
	Dim guidStr
	
    
	
	UserId=Session("UserID")
	action = ucase(Request.Form("action"))
	
	if action = "EXECUTE" then
	        
			Dim MarketCap
			Dim MarketDate
			
			MarketCap = replace(Request.Form("MarketCap"),",","")
			MarketDate = Request.Form("txtDate")
			
			'response.write DateDiff("d",cdate(Date),cdate(MarketDate)):response.end
			if  MarketDate = "" then
			 %><script language = 'Javascript'>
					alert('Please specify the Market Date.')
					//window.history.back()
				</script><% response.end
			end if     
			
			if  not isDate(MarketDate) then
			 %><script language = 'Javascript'>
					alert('Invalid  Date.')
					//window.history.back()
				</script><% response.end
			end if  
			
			if  MarketCap = "" then
			 %><script language = 'Javascript'>
					alert('Please specify the Market Price.')
					//window.history.back()
				</script><% response.end
			end if 
			
			if  not isnumeric(MarketCap) then
			 %><script language = 'Javascript'>
					alert('Invalid Market Price.')
					//window.history.back()
				</script><% response.end
			end if 
			
			
			'Validate that the date is not less that 3 days from the current date
			'get days to get the difference
			if WeekDay(FormatDate(Date), VBMonday) >= 3 then
				DayDiff=-3
			else
				DayDiff=-5
			end if
			
			
			If (DateDiff("d",cdate(Date),cdate(MarketDate)) < DayDiff)  Then
			 %><script language = 'Javascript'>
					alert('You are not allowed to enter values for trades more than three WORKING days old.')
					//window.history.back()
				</script><% response.end
			end if 
			
			'Validate for weekends
			
			If WeekDay(FormatDate(MarketDate), VBMonday) = 6 or WeekDay(FormatDate(MarketDate), VBMonday) = 7 then
				%><script language = 'Javascript'>
					alert('You are not allowed to enter values for the weekend.')
					//window.history.back()
				</script><% response.end
			
			end if
			
						
			
					
		     Set conn = GetActiveConnection("KBroker")
			 
			 sqlstr = "SELECT     TradeDate " & _
					 " FROM         tblTurnOver " & _
					 " WHERE     (CAST(FLOOR(CAST(TradeDate AS float)) AS datetime) = '"& formatdate(MarketDate) &"')"
			set RSDup = conn.execute(sqlstr)
			
			if not(RSDup.eof or RSDup.bof) then
			 %><script language = 'Javascript'>
					alert('The value for that date has already been entered.')
					window.history.back()
				</script><% 
				conn.close
				set RSDup = nothing
				set conn = nothing
				response.end
			
			end if
		  
			 conn.BeginTrans	
			 
			 
			 				    
			'save data
			sqlStr = "INSERT INTO tblTurnOver (TradeDate, MarketTurnOver, CreatedBy, TimeCreated) " & _
					 " VALUES     ('"& formatdate(MarketDate) &"',"& MarketCap &", "& UserId &", GETDATE())"
					  
			'response.write SQLServerFormatWithCustomMax(sqlStr):response.end
           
			conn.Execute (SQLServerFormatWithCustomMax(sqlStr))			 	
			conn.CommitTrans
					
			 WritefraEnabledDialogCloseScript
		
			 conn.Close
			 Set conn = Nothing
					
			 Response.End	
   	End if
   	
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
	
	
		
    function FormatNumber2(Obj){
	 var theVal = Obj.value
	 Obj.value = formatNum(theVal)
	}
	
	function FormatNumber(Obj,decimals){
	 var theVal = Obj.value
	 Obj.value = format_number11(theVal,decimals)
	}
	
	function format_number11(pnumber,decimals){
			var snum = new String(pnumber);
			
			snum=snum.replace(/,/gi,"");
			
			if (isNaN(snum)) { return 0};
			if (snum=='') { return 0};
			
			
			var sec = snum.split('.');
			var whole = parseFloat(sec[0]);
			var result = '';
			
			if(sec.length > 1){
				var dec = new String(sec[1]);
				dec = String(parseFloat(sec[1])/Math.pow(10,(dec.length - decimals)));
				dec = String(whole + Math.round(parseFloat(dec))/Math.pow(10,decimals));
				var dot = dec.indexOf('.');
				if(dot == -1){
					dec += '.'; 
					dot = dec.indexOf('.');
				}
				while(dec.length <= dot + decimals) { dec += '0'; }
				result = dec;
			} else{
				var dot;
				var dec = new String(whole);
				dec += '.';
				dot = dec.indexOf('.');		
				while(dec.length <= dot + decimals) { dec += '0'; }
				result = dec;
			}	
			return addCommas(result);
		}
	
	function addCommas(number)
		{
		  var num1 = ((number.length > 3) ? (addCommas(number.substring(0, number.length - 3)) + "," + number.substring(number.length - 3, number.length)) : 
			 String(number));
		  return num1.replace(',.','.')
		}
</script>
</head>
<body Class="Dialog">
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "txtDate","cmdDate","<%=FormatDate(now())%>",1);
</SCRIPT>

<form name = 'frmMain' method = 'post' action = 'AddMarketTurnover.asp' id = "frmMain" >
<table border="0" width="100%" cellspacing="1" cellpadding="1">

 
  
  <tr >
	<td width="15%" valign="top"> Date</td>
	<td ><SCRIPT language="JavaScript">cal.writeControl();</SCRIPT></td>
  </tr>
  <tr>
    <td width="15%">Market Capitalisation</td>
    <td width="54%">
    <input type="text" name="MarketCap" id="MarketCap" value="" size="20" onChange="Javascript: FormatNumber(this,2)"></td>
  </tr>
  <tr>
	  <td width="100%" colspan=2 align="center" valign=absBottom>
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdSave' id = 'cmdSave' value=" Save ">
    	<input type = 'button' Class=Buttons name ='cmdCancel' id = "cmdCancel" value="Cancel" onclick = "JavaScript: window.self.close()">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		
	</td>
  </tr>
</table>

</form>
</body>

</html>