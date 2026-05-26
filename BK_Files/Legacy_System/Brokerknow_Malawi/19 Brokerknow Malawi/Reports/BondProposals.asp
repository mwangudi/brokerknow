<html>

<head>
	<meta http-equiv="Content-Language" content="en-uk">
	<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
	
	<title>Bond Proposals</title>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	  <!--CALENDAR -->
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
	<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
	<SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>
    <style media="print">
		@page {
			@top{font-family: Helvetica, Arial, sans-serif;
				font-size: 150%;
				font-weight: bolder;
				text-align: left;
				content: "<%= FormatDate(Date) %>";			
			}		
			
			
			margin-left: 2cm;
			margin-right: 5cm;
			margin-top: 1cm;    
			margin-bottom: 2cm;
			size: portrait;
			
			br.newpage{
				page-break-before:always;
			}
			
			tr.pageNumbering{
				display:none;
			}
			
		}

	</style>
</head>

<body Class="Reports">

<!--#include file="../libroutines.asp"-->


<%

genReport = Request.Form("genReport")
selectedTradeDate = Request.Form("txtDate")
useOwner = Request.Form("useOwner")
selOwners = Request.Form("selOwners")

If genReport <> "1" Or Not IsDate(selectedTradeDate) Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
	
		function validateForm(frm){			
			if (frm.txtDate.value==''){
				alert("Select a date");
				frm.txtDate.focus();
				return;
			}
			
			frm.target = '_self';			
			frm.submit();
		}
		function switchDisplay(obj){
			if (obj.style.display=='none') obj.style.display = '';
			else obj.style.display = 'none';
		}
		
		
		function calendarChange(dateValue){
			var Input = document.all.item('AllSelOwners');
			var Output =  document.all.item('selOwners');
			var dateVal;
			var mainObj = document.all.item('txtDate');
			
			if (dateValue==null || dateValue=='undefined') dateVal = mainObj.value;
			else dateVal = dateValue;
			
			Output.length = 0;
		     for (loop=0; loop < Input.length; loop++){
		     		if (Input.options[loop].TAG == clientFormatDate(dateVal)){		     			
		     		    NewOption = new Option();   			    
		   			    NewOption.text = Input.options[loop].text;
		   			    NewOption.value = Input.options[loop].value;			
		   			    Output.add(NewOption, 0);
		     		}	
		     }
		}
		
		
		var cal = new ctlSpiffyCalendarBox("cal", "frmMain", "txtDate","cmdDate","<%= FormatDate(Date) %>",1);
		
		
	</Script>
	
	<Script Language="VBScript">
		Function clientFormatDate(theDate)
			On Error Resume Next
			clientFormatDate = Day(theDate) & "-" & MonthName(Month(theDate), True) & "-" & Year(theDate)
			If Err.Number > 0 Then
				clientFormatDate = theDate
			End If
		End Function
	</Script>
	
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="BondProposals.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<input type="hidden" Name="timeLimit" value="1">
		<table>
				
			<tr>
				<td>Select any day of month</td>
				<td>
					<SCRIPT language="JavaScript">
						cal.writeControl();
					</SCRIPT>	
				</td>
			</tr>
			
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... " id=Button1 name=Button1>&nbsp;&nbsp; </td>
			</tr>
		</table>
		
	</form>
		
	<%Set rs = Nothing
	Set Conn = Nothing
	Response.End
End If

%>

<% DrawPageFunctions True, True, True %>

<%
   Dim conn 
   Dim sqlStr
   Dim rs
   Dim isDifferentBroker
   Dim currBrokerCode
   
   Set conn = GetActiveConnection("KBroker")
   
   sqlStr = "SELECT * FROM Bondoffer WHERE Proposaldate = '" & formatdate(selectedTradeDate) & "'"
   Set Rs = Conn.Execute (sqlStr)
   If Rs.EOF Or Rs.BOF Then%>
		<Script Language="JavaScript">	
			ShowMessage('There were no bond proposals under the selected date');
			window.parent.history.go(-1)
		</Script>
		<%Set Conn = Nothing
		Set Rs = Nothing
		Response.End
   End If
   
   
   
   
  '' sqlStr ="select * from BondOffer order by Proposal_DPA_"
   
   Set Rs = Conn.Execute(sqlStr)         
   Dim pageNumber
	
	pageNumber = 0
	Dim balance
	Dim acntManag
	Dim commission
	Dim NetProceeds
	
   While not Rs.EOF
   
   
		pageNumber = pageNumber + 1
%> 

<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" width="100%">
	<tr class="pageNumbering">
		<td align="left" >
			<FONT FACE=ARIAL SIZE=2><B>Page <%=pageNumber%>	</B></FONT>	
		</td>		
	</tr>
	
	<tr>
		<td align="right">           			
        <table border="0" width="100%">
         <tr>        
           <td align="right" colspan=2 height="200px" style="border-bottom: 2px inset #000000" width="670">
			<Img Src="../data/photos/aaprintlogo.bmp">			
		</td>
          </tr>
        </table>           			
		</td>		
	</tr>
	
</table>	  

  <table border="0" cellspacing="0" cellpadding="5" style="font-family: Arial" width="70%" align ="center">
	<tr>
      <td colspan="2" align="left" width=100%">
       <%if  trim(Rs.Fields("contact").Value)="" then
		else
		''Response.Write ucase(Rs.Fields("contact").Value)
		end if
		
       %>
       <BR>
	   <% ''replace(Rs.Fields("AccountAddress").Value,",",",<br>")
        response.write Rs.Fields("ClientName").Value %> 
       <br>
        
       <% ''replace(Rs.Fields("AccountAddress").Value,",",",<br>")
        response.write Rs.Fields("AccountAddress").Value %> 
       <br>
        <%
         if trim(Rs.Fields("Fax").Value) <> ""  then
			Response.Write "Fax: "& Rs.Fields("Fax").Value
		end if
          %> 
       </td>
    </tr>
    <%
       if trim(Rs.Fields("Contact").Value) <> ""  then
    %>
    <tr>
      <td colspan="2" align="left" height="16" valign="bottom"><b>ATTENTION:&nbsp;&nbsp;<%=UCASE(Rs.Fields("Contact").Value)  %> </td>
    </tr>
    <%end if%>
        <tr>
      <td colspan="2" align="left">Dear&nbsp;<% if trim(Rs.Fields("Salutation"))="" then 
					response.write "Sir"
				else
				response.write Rs.Fields("Salutation")
			end if%>, </td>
    </tr>
    <tr>
      <td colspan="2" align="center"><b><font size="4"><u>
      <%If cint(rs("TradeType"))=cint(1) then%> 
		 BOND PURCHASE PROPOSAL
        <%elseif cint(rs("TradeType"))=cint(2) then%>
        BOND SALE PROPOSAL
        <%
        End if
             
        %>
        </u></font></b></td>
    </tr>
  
      <tr>
      <td width="50%">Issue No</td>
      <td width="50%" align="right"><%= Rs.Fields("BondIssue").Value %></td>
    </tr>
   
    <tr>
      <td width="50%"> <span lang="en-us">Price</span></td>
      <td align="right">
      <% 
      
      if isnull(Rs.Fields("AlternatePrice").Value) then 
			AlternatePrice =0 
		else
			AlternatePrice=Rs.Fields("AlternatePrice").Value
		end if
			
    	if isnumeric(AlternatePrice)then Price =AlternatePrice
		 	if  Price <> 0 then
		 	 	Price =AlternatePrice
			 else
			 	Price =rs.Fields("BondDirtyPrice")
		end if
			 
      
    	Response.write formatnumber(Price,4) & " %"
    	%>
    </td>
    </tr>
   
    <tr>
      <td width="50%"> Face Value</td>
      <td align="right"><%= formatnumber(Rs.Fields("FaceValue").Value,2) %></td>
    </tr>
    <tr>
    
    
      <td width="50%">Consideration</td>
      <td align="right"><%
      	if isnumeric(price) then
		 Consideration =((Price * Rs.Fields("FaceValue").Value)/100)
		end if
     ' Consideration= Rs.Fields("FaceValue").Value*Rs.Fields("BondDirtyPrice").Value 
			Response.Write formatnumber(RoundPoint05(Consideration),2)
			%></td>
    </tr>
    <tr>
      <td>Commission</td>
      <td align="right">
      <%
		'' Commission=((Rs.Fields("Commission").Value * Consideration)/100)
		Commission=(Rs.Fields("Commission").Value)
		Response.Write formatnumber(RoundPoint05(Commission),2)
	  %></td>
    </tr>
    <tr>
      <td><b>Total</b></td>
      <td align="right"><b>
     		<%if cint(Rs.Fields("TradeType").Value)=cint(1) then
				Response.Write formatnumber(RoundPoint05(formatnumber((Consideration+Commission),2)),2)						
			elseif cint(Rs.Fields("TradeType").Value)=cint(2)then
				Response.Write formatnumber(RoundPoint05(formatnumber((Consideration-Commission),2)),2)	
			end if%></b>
		</td>
    </tr>
     <tr>
      <td width="50%">Settlement date</td>
      <td align="right"><%= FormatDate(Rs.Fields("settlementDate").Value) %></td>
    </tr>
    <tr>
      <td width="50%">Issue date</td>
      <td align="right"><%= FormatDate(Rs.Fields("BondIDate").Value) %></td>
    </tr>
     <tr>
    <td width="50%">Maturity date</td>
      <td align="right"><%= FormatDate(Rs.Fields("BondMDate").Value) %></td>
    </tr>
    <tr>
      <td width="50%">No. of days to maturity</td>
      <td align="right">
      <%response.write (Rs.Fields("RemainingDaysToCoupon").value + (Rs.Fields("PreviousCouponPayments").Value* Rs.Fields("CouponPeriodDays").Value))%>
      </td>
    </tr>
    <tr>
      <td align="left">Buyer's Yield</td>
      <td align="right"><%= formatnumber(Rs.Fields("ForwardRate").Value,2)&" %" %></td>
    </tr>
    <tr>
      <td width="50%">Duration</td>
      <td align="right"><%= Formatnumber(Rs.Fields("Duration").Value,2) %></td>
    </tr>
    <tr>
      <td width="50%">Convexity</td>
      <td align="right"><%= Formatnumber(Rs.Fields("Convexity").Value,2) %></td>
    </tr>   
    <tr>
      <td colspan="2">&nbsp;<br>Kindly confirm.</td>
      
    </tr>
  </table>
<br>  
<table border="0" cellspacing=2 cellpadding=2 width="70%" align="center">
	<tr>
		<td align="left" valign="top">
<PRE><font face="Arial"><b>
Yours faithfully, <BR>	
FOR: DYER AND BLAIR LIMITED</b>
</font></PRE>			
		</td>		
	</tr>
	<tr>
		<td align="left" valign="bottom" height="70">
<PRE><font face="Arial"><b>		
<%if trim(Rs.Fields("Owner").Value)<>"" then
	Response.Write UCASE(Rs.Fields("Owner").Value)
  else
	Response.write "MARTIN MBUGUA"
 end if  
 %> <BR>RELATIONSHIP MANAGER</b>
</font></PRE>			
		</td>		
	</tr>
	
	<tr>
		<td  align="right" height="30">
            &nbsp;			
		</td>		
	</tr>
	<tr><td><!--#Include file="DirectorFooter.asp"--></td></tr>
</table>  



<%		Rs.MoveNext
		'important!		
			If Not Rs.EOF Then %>
				<BR class="newpage">
		<%	End If
	wend

	Set Rs = Nothing
	Set Rs = Nothing
Set Conn = Nothing%>
</body>

</html>