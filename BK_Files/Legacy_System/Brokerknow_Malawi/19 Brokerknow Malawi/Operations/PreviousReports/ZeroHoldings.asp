<html>
<%
	' DB Connection
Dim Conn 
    Set Conn = CreateObject("ADODB.Connection")
    theDBName = "KBroker" 
    Conn.ConnectionString =  "FILE NAME=" & GetUDLPath(theDBName) 
    Conn.Open
    
   Function GetUDLPath(theDBName) 
    Dim tmpStr
    
    tmpStr = StrReverse(Request.ServerVariables("APPL_PHYSICAL_PATH"))
    
    tmpStr = Mid(tmpStr, InStr(1, tmpStr, "\") + 1)
    
    tmpStr = StrReverse(tmpStr)
    
    GetUDLPath = tmpStr & "\UDL\" & Trim(theDBName) & ".UDL"

End Function
%>
<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>ZERO HOLDINGS</title>
	<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
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
			
			
		}

	</style>
</head>

<body Class="Reports">
<!--#include file="../libroutines.asp"-->

<%

genReport = Request.Form("genReport")
selectedClient = Request.Form("cboClient")
selectedFromDate = Request.Form("transFromDate")



%>

<% DrawPageFunctions True, True, True %>
	
<%
	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")						        
	sqlStr = "select * from ZeroHoldings order by Client_DPA_"
	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(sqlStr), conn.ConnectionString, 0, 1

%>	

<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
		<td width="10%" nowrap><font face="Impact" size="4">ZERO HOLDINGS</font></td>
      <td width="60%" nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
      
    </tr>

  </table>
<br>
<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
      <td ><b> As of: <%= FormatDate(now()) %></b></td>
      
    </tr>

</table>
<BR>



    

   <%
     Rs.PageSize=50

		Rs.CacheSize = Rs.PageSize
		intPageCount = Rs.PageCount 
		intRecordCount = Rs.RecordCount 
	
		first=cint(0)
		
	' Now you must double check to make sure that you are not before the start
	' or beyond end of the recordset.  If you are beyond the end, set 
	' the current page equal to the last page of the recordset.  If you are
	' before the start, set the current page equal to the start of the recordset.	

       PageNumber1=PageNumber1 + 1
        
        intPage=0
        
         intPageCount=Cint(intPageCount)
        
      t=0      
	Do while Cint(intPage) < intPageCount	
	intPage=intPage + 1
	if(Cint(first)=cint(1)) then
	%>
             <BR class="newpage">
    <%
		
	end if
	If CInt(intPage) > CInt(intPageCount) Then intPage = intPageCount
	If CInt(intPage) <= 0 Then intPage = 1
	
	 'Make sure that the recordset is not empty.  If it is not, then set the 
	 'AbsolutePage property and populate the intStart and the intFinish variables.
	
	'if Not(Rs.eof and Rs.bof) Then

	If intRecordCount > 0 Then 'and Not(Rs.eof and Rs.bof) 
		Rs.AbsolutePage = intPage
		intStart = Rs.AbsolutePosition
		'Response.write(intStart)
		
		If CInt(intPage) = CInt(intPageCount) Then
			intFinish = intRecordCount
		Else
			intFinish = intStart + (Rs.PageSize - 1)
		End if
	End If	  

    %>
	  <table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow; LEFT-MARGIN:100PX"  width="100%">
		<tr>
      <td style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Code</font></b></td>
      <td align="left" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  align=left><b><font face="Arial Narrow" size="3">Client</font></b></td>
      <td align="left" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  align=left><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
      <td style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">REF No.</font></b></td>
      
    </tr>
	<%
	   'Do Until Rs.EOF       

		For intRecord = 1 to Rs.PageSize

	%>
		<tr>	
		  <td nowrap valign="top"><font size="1"><%= rs.Fields("Client_DPA_").Value %></font></td>		  
		  <td  style="text-align: left" nowrap valign="top"><font size="1" >&nbsp;&nbsp;<%= mid(Rs.Fields("ClientName").Value,1,35) %></font></td>
		  <td nowrap valign="top"><font size="1" >&nbsp;</font></td>
		  <td nowrap valign="top"><font size="1" >&nbsp;&nbsp;<%= Rs.Fields("ClientCDSNo").Value %></font></td>
		</tr>
	
	<%		'Rs.MoveNext	
		rs.MoveNext
        
		If Rs.EOF Then Exit for

        Next
	%>
	
    <tr>
      <td colspan="6" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        &nbsp;&nbsp;&nbsp; </td>

    </tr>
	</table>
	<%	

	loop

	%>


   
<%Set Rs = Nothing
Set Conn = Nothing%>   
</body>

</html>