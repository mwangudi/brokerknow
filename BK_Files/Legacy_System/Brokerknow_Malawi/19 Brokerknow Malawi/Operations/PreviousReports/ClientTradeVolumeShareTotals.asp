<!--#include file="../libroutines.asp"-->

<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Portfolio BS</title>  
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
<SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>



<style media="print">
		@page {
				size: landscape;
				margin-left: 2cm;
				margin-right: 5cm;
				margin-top: 1cm;    
				margin-bottom: 2cm;
				writing-mode: tb-rl;
				height: 80%;
				margin: 10% 0%;						
				br.newpage{
					page-break-before:always;
				}		
			}		 
	</style>
</head>

<body Class="Reports">



<%
'FirstDay=DateSerial(Year(Date), Month(Date) + iOffset, 1)
FirstDay=Date-90
genReport = Request.Form("genReport")
selectedClient = Request.Form("cboClient")
selectedFromDate = Request.Form("txtFromDate")
selectedToDate = Request.Form("txtToDate")

mingross =Request.Form("txtMinGross")
maxgross=Request.Form("txtMaxGross")
minvol=Request.Form("txtMinVol")
maxvol=Request.Form("txtMaxVol")
mincomm=Request.Form("txtMinComm")
maxcomm=Request.Form("txtMaxComm")

dates =Cint(Request.Form("txtdates"))
grosses=Cint(Request.Form("gross"))
volumes=Cint(Request.Form("volume"))
commissions=Cint(Request.Form("commission"))

If genReport <> "1" Then%>
		<Script Language="JavaScript">
		report_SetBodyClass();
		function validateForm(frm){			
			//if (frm.txtDate.value==''){
			//	alert("Select a date");
			//	frm.txtDate.focus();
			//	return;
			//}
			
			frm.target = '_self';			
			frm.submit();
		}
		
		function UpdateDate (chkbox)
			{			
			if(chkbox.checked)
				{
				if(chkbox.name=='chkDates')
					{					
					document.frmMain.elements("txtdates").value="1"
					}
				if(chkbox.name=='chkGross')
					{					
					document.frmMain.elements("gross").value="1"
					}
				if(chkbox.name=='chkCommission')
					{					
					document.frmMain.elements("commission").value="1"
					}
				if(chkbox.name=='chkVolume')
					{					
					document.frmMain.elements("volume").value="1"
					}
				}
				else
				{
				if(chkbox.name=='chkDates')
					{					
					document.frmMain.elements("txtdates").value="0"
					}
				if(chkbox.name=='chkGross')
					{					
					document.frmMain.elements("gross").value="0"
					}
				if(chkbox.name=='chkCommission')
					{					
					document.frmMain.elements("commission").value="0"
					}
				if(chkbox.name=='chkVolume')
					{					
					document.frmMain.elements("volume").value="0"
					}												
				}			
			}
		function formatnumber(theTxt)
	{	
	var theprice = theTxt.value;
	//var thesectype=document.frmMain.elements("cboOrderSecType").value
	
	//if(thesectype==1)
	//	{
		theprice =format_number(theprice,2);
	//	}
	//	else
	//	{
	//	theprice =format_number(theprice,2);
	//	}
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


		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "txtFromDate","cmdDate","<%= FormatDate(FirstDay) %>",1);
		var cal1=new ctlSpiffyCalendarBox("cal1", "frmMain", "txtToDate","cmdDate","<%= FormatDate(Date) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="ClientTradeVolumeShareTotals.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<table>			
			<tr>
				<td> <input type="checkbox" OnClick="JavaScript: UpdateDate (this); " class="BorderLess" name="chkDates" id="chkDates" value="0"><label for="useOwnFields" style="cursor: hand"><b>Dates</b></label></td>
				<td>From Date:&nbsp;	
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
				<td>To date:&nbsp;
					<SCRIPT language="JavaScript">cal1.writeControl();</SCRIPT>	
				</td>
			</tr>		
			<tr>
				<td> <input type="checkbox" OnClick="JavaScript: UpdateDate (this); " class="BorderLess" name="chkGross" id="chkGross" value="0"><label for="useOwnFields" style="cursor: hand"><b>Gross</b></label></td>
				<td>Min:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<input STYLE="WIDTH: 100px; text-align: left" type = 'text'  name ='txtMinGross' id = 'txtMinGross' size="20"  OnBlur="JavaScript: format2Number(this)">
				</td>
				<td>Max:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<input STYLE="WIDTH: 100px; text-align: left" type = 'text'  name ='txtMaxGross' id = 'txtMaxGross' size="20"  OnBlur="JavaScript: format2Number(this)">
				</td>
			</tr>		
			<tr>
				<td> <input type="checkbox" OnClick="JavaScript: UpdateDate (this); " class="BorderLess" name="chkCommission" id="chkCommission" value="0"><label for="useOwnFields" style="cursor: hand"><b>Commission</b></label></td>
				<td>Min:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<input STYLE="WIDTH: 100px; text-align: left" type = 'text'  name ='txtMinComm' id = 'txtMinComm' size="20"  OnBlur="JavaScript: format2Number(this)">
				</td>
				<td>Max:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<input STYLE="WIDTH: 100px; text-align: left" type = 'text'  name ='txtMaxComm' id = 'txtMaxComm' size="20"  OnBlur="JavaScript: format2Number(this)">
				</td>
			</tr>		
			<tr>
				<td> <input type="checkbox" OnClick="JavaScript: UpdateDate (this); " class="BorderLess" name="chkVolume" id="chkVolume" value="0"><label for="useOwnFields" style="cursor: hand"><b>Volume</b></label></td>
				<td>Min:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<input STYLE="WIDTH: 100px; text-align: left" type = 'text'  name ='txtMinVol' id = 'txtMinVol' size="20" OnBlur="JavaScript: format2Number(this)">
				</td>
				<td>Max:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<input STYLE="WIDTH: 100px; text-align: left" type = 'text'  name ='txtMaxVol' id = 'txtMaxVol' size="20"  OnBlur="JavaScript: format2Number(this)">
				</td>
			</tr>		
			
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... ">&nbsp;&nbsp;</td>
				<input type = 'hidden' name ='txtdates' id = 'txtdates' value='0'>
				<input type = 'hidden' name ='gross' id = 'gross' value='0'>
				<input type = 'hidden' name ='commission' id = 'commission' value='0'>
				<input type = 'hidden' name ='volume' id = 'volume' value='0'>		
			</tr>
		</table>
		
	</form>
	
	<%
	Response.End
End If

%>
<%DrawPageFunctions True, True, True%>


<%
selectedFromDate = Request.Form("txtFromDate")
selectedToDate = Request.Form("txtToDate")

mingross =Replace(Request.Form("txtMinGross"),",","")
maxgross=Replace(Request.Form("txtMaxGross"),",","")
minvol=Replace(Request.Form("txtMinVol"),",","")
maxvol=Replace(Request.Form("txtMaxVol"),",","")
mincomm=Replace(Request.Form("txtMinComm"),",","")
maxcomm=Replace(Request.Form("txtMaxComm"),",","")

dates =Cint(Request.Form("txtdates"))
grosses=Cint(Request.Form("gross"))
volumes=Cint(Request.Form("volume"))
commissions=Cint(Request.Form("commission"))

SelectedSearchArgs=""
SelectedGroupedArgs=""
FilterCriteria=""

if(dates=1) then
SelectedSearchArgs="TransDate between '" & CDate(selectedFromDate) & "' and '" & CDate(selectedToDate)+1 & "'"
FilterCriteria="Filtered By Trade Date Between '" & FormatDate(CDate(selectedFromDate)) & "' and '" & FormatDate(CDate(selectedToDate)) & "'"
end if

if(grosses=1) then
	if(mingross="" or maxgross="") then
      %>
		<Script Language="JavaScript">
			alert("Please Specify the criteria for Gross.")
			window.parent.history.go(-1);			
		</Script>
		<%
		Response.End
	end if
	If (mingross <> "" And Not IsNumeric(mingross)) or (maxgross <> "" And Not IsNumeric(maxgross)) Then%>
		<script language = 'JavaScript'>
			alert("Gross Must be Numeric.")
			window.parent.history.go(-1);
		</script>
		<% response.end
	End If

	if(SelectedGroupedArgs="") then
	SelectedGroupedArgs="Sum(Gross) between " & mingross & " and " & maxgross
	else
	SelectedGroupedArgs=SelectedGroupedArgs & " and Sum(Gross) between " & mingross & " and " & maxgross
	end if

      if(FilterCriteria="") then
	FilterCriteria="Filtered By Gross between " & formatnum(mingross) & " and " & formatnum(maxgross)
      else	  
      FilterCriteria=FilterCriteria & " then by Gross between " & formatnum(mingross) & " and " & formatnum(maxgross)
	end if
end if

if(commissions=1) then
	if(mincomm="" or maxcomm="") then
      %>
		<Script Language="JavaScript">
			alert("Please Specify the criteria for Commission.")
			window.parent.history.go(-1);			
		</Script>
		<%
		Response.End
	end if
	
	If (mincomm <> "" And Not IsNumeric(mincomm)) or (maxcomm <> "" And Not IsNumeric(maxcomm)) Then%>
		<script language = 'JavaScript'>
			alert("Gross Must be Numeric.")
			window.parent.history.go(-1);
		</script>
		<% response.end
	End If

	if(SelectedGroupedArgs="") then
	SelectedGroupedArgs="Sum(Commission) between " & mincomm & " and " & maxcomm
	else
	SelectedGroupedArgs=SelectedGroupedArgs & " and Sum(Commission) between " & mincomm & " and " & maxcomm
	end if
	
	if(FilterCriteria="") then
	FilterCriteria="Filtered By Commission between " & formatnum(mincomm) & " and " & formatnum(maxcomm)
      else	  
      FilterCriteria=FilterCriteria & " then by Commission between " & formatnum(mincomm) & " and " & formatnum(maxcomm)
	end if
end if

if(volumes=1) then
	if(minvol="" or maxvol="") then
      %>
		<Script Language="JavaScript">
			alert("Please Specify the criteria for Volumes.")
			window.parent.history.go(-1);			
		</Script>
		<%
		Response.End
	end if

	If (minvol <> "" And Not IsNumeric(minvol)) or (maxvol <> "" And Not IsNumeric(maxvol)) Then%>
		<script language = 'JavaScript'>
			alert("Gross Must be Numeric.")
			window.parent.history.go(-1);
		</script>
		<% response.end
	End If

	if(SelectedGroupedArgs="") then
	SelectedGroupedArgs="SUM(Volume) between " & minvol & " and " & maxvol
	else
	SelectedGroupedArgs=SelectedGroupedArgs & " and SUM(Volume) between " & minvol & " and " & maxvol
	end if
	
	if(FilterCriteria="") then
	FilterCriteria="Filtered By Volume between " & formatnum(minvol) & " and " & formatnum(maxvol)
      else	  
      FilterCriteria=FilterCriteria & " then by Volume between " & FormatNum(minvol) & " and " & FormatNum(maxvol)
	end if
end if

'sqlStr = "SELECT " & selColumns & " FROM ClientTradeVolumeBondsTotals " & SelectedSearchArgs & " " &  orderByCols 

if(SelectedSearchArgs<>"") then
	if(SelectedGroupedArgs<>"") then
		sqlStr="SELECT [Client Code], [Client Name], SUM(Volume) AS Volume, SUM(Commission) AS Commission, SUM(Gross) AS Gross" & _
	 		" FROM ClientTradeVolumeSharesTotals WHERE (" &  SelectedSearchArgs & ")" & _
	 		" GROUP BY [Client Code], [Client Name] HAVING (" & SelectedGroupedArgs & ")"	
     else
	sqlStr="SELECT [Client Code], [Client Name], SUM(Volume) AS Volume, SUM(Commission) AS Commission, SUM(Gross) AS Gross" & _
	 		" FROM ClientTradeVolumeSharesTotals WHERE (" &  SelectedSearchArgs & ")" & _
	 		" GROUP BY [Client Code], [Client Name]"	

     end if
else
	if(SelectedGroupedArgs<>"") then
		sqlStr="SELECT [Client Code], [Client Name], SUM(Volume) AS Volume, SUM(Commission) AS Commission, SUM(Gross) AS Gross" & _
	 		" FROM ClientTradeVolumeSharesTotals " & _
	 		" GROUP BY [Client Code], [Client Name] HAVING (" & SelectedGroupedArgs & ")"	
     else	
		sqlStr="SELECT [Client Code], [Client Name], SUM(Volume) AS Volume, SUM(Commission) AS Commission, SUM(Gross) AS Gross" & _
	 		" FROM ClientTradeVolumeSharesTotals " & _
	 		" GROUP BY [Client Code], [Client Name]" 
     end if
end if
Set Conn = GetActiveConnection("KBroker")

Set Rs = Conn.Execute(sqlStr)

 If Rs.EOF Or Rs.BOF Then%>
		<Script Language="JavaScript">	
			ShowMessage('No information was found using the criteria entered.');
			window.parent.history.go(-1);
		</Script>
		<%Set Conn = Nothing
		Set Rs = Nothing
		Response.End
  End If


 %>

<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
     <tr>
		<td nowrap><b><font face="Arial Narrow" size="4">CLIENT TRADE VOLUME FOR SHARES</font></b></td>
		<td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
	</tr>	
       <tr>
		  <td COLSPAN=2><font face="Arial" size="2"><%=filtercriteria%></font></td>
	</tr>
</table>			



    <table border="0" width="100%" cellPadding="2" cellSpacing=0>
    <tr bgColor="#000000">
			
	<%For i = 0 To Rs.Fields.Count - 1
	if(i<>1 and i<>0) then
	%>		
	 	 <td nowrap align="right"><b><font color="#FFFFFF"><%= Rs.Fields(i).Name %></font></b></td>
   	<%	
	else
	%>		
	 	 <td nowrap><b><font color="#FFFFFF"><%= Rs.Fields(i).Name %></font></b></td>
   	<%
	end if
	Next %>
	</tr>
	
	<% 	GrossTotal=0
		CommissionTotal=0
      	VolumeTotal=0

		Do Until rs.EOF%>
        		<tr>
      				
      				<%For i = 0 To Rs.Fields.Count - 1
						if(i<>0) then
							if(i=1) then
							%>		
					 	 	<td nowrap><%= Mid(Rs.Fields(i).Value,1,30) %></td>
				   			<%							
							else
								if(i=2) then
								VolumeTotal=VolumeTotal + Rs.Fields(i).Value
								end if
								if(i=3) then
								CommissionTotal=CommissionTotal + Rs.Fields(i).Value
								end if
								if(i=4) then
								GrossTotal=GrossTotal + Rs.Fields(i).Value
								end if				
								
							%>		
					 	 	<td nowrap align="right"><%= FormatNum(Rs.Fields(i).Value) %></td>
				   			<%
							end if
						else
						%>		
					 	 <td nowrap><%= Rs.Fields(i).Value %></td>
				   		<%						
						end if
					Next %>		
	        </tr>
	 <%  Rs.MoveNext
	 Loop%>
	<tr>
	<td colspan="5">&nbsp;</td>
	</tr>

	<tr>
	<td colspan="2" align="right"></td>
	<td align="right"><b><%=FormatNum(VolumeTotal)%></b></td>
	<td align="right"><b><%=FormatNum(CommissionTotal)%></b></td>
	<td align="right"><b><%=FormatNum(GrossTotal)%></b></td>
	</tr>	  	
	  	
  </table>
  
	 <%	 
	 Set Rs = Nothing
	 Set Conn = Nothing
     %>	
</body>



</html>
