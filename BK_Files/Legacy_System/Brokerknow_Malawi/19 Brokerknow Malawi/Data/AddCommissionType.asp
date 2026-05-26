<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Commission Type</title>
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="scripts/common.js"></SCRIPT> 

</head>

<body><!--#include file="../libroutines.asp"-->
<%
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
	
	action = ucase(Request.Form("action"))
	
	if action = "EXECUTE" then
		Dim description
       Dim rate
       
       description = Request.Form("txtDescription")
       rate = Request.Form("txtRate")
      
       
        'validate Rate
        If Trim(Rate) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Rate"
                		window.history.back
                </script>
                <% response.end
        End If
        'validate size of Description
        If Len(Description) > 100 Then%>
                <script language = 'vbscript'>
                ShowMessage "Description can only be 100 characters in length"
                window.history.back
                </script>
                <% response.end
        End If
        'ensure Rate is numeric
        If (Rate <> "") And (Not IsNumeric(Rate)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Rate  must be numeric"
                window.history.back
                </script>
                <% response.end
        End If
        
       
        'save data
       sqlStr = "INSERT INTO [Commission] (CommissionDescription,CommissionRate,Commission_DPA_) SELECT " & "'" & description & "'" & " as CommissionDescription" & _
                "," & " " & rate & " " & " as CommissionRate," & " " & "iif(isnull(max([Commission_DPA_])),1,max([Commission_DPA_]) + 1)" & " " & " as Commission_DPA_" & _
                " FROM [Commission]"
        Set conn = GetActiveConnection("KBroker")
        
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
        conn.Close
        Set conn = Nothing
   	end If
%>

<CENTER>
	<DIV class="ListNugget" id="AdvSearchHead" style="WIDTH: 640px" name="AdvSearchHead">
		<TABLE class="ListNuggetHeader" cellPadding="0" cellSpacing="0" width="100%" name="AdvSearchtestHeader"> 
			<TR>
			<TD class="ListNuggetTitleCellWhite"
					onselectstart="window.event.cancelBubble=true; return false;"   
					onclick="PartWrapperToggle('AdvSearchHead');">
					<A class=ListNuggetTitle onclick="return PartWrapperToggle('AdvSearchHead');"  
					 href="javascript:PartWrapperToggle('AdvSearchHead');">Add Commission Type
					</A>
				</TD>
			 
				<TD class=ListNuggetButtonCellWhite onclick="PartWrapperToggle('AdvSearchHead');">
				<DIV class=ListNuggetButton>
					<IMG class=ListNuggetUpButton id=AdvSearchUp height=17 alt="Hide options" src="../images/blue-chevron_up.gif" width=17 align=right border=0 name=AdvSearchHeadUp>
					<IMG class=ListNuggetDownButton id=AdvSearchDown height=17 alt=Options src="../images/gray-chevron_down.gif" width=17 align=right border=0 name=AdvSearchHeadDown>
				</DIV>
			</TD>
			</TR>
		</TABLE>
		
<DIV class="ListNuggetBody" id="AdvSearchHeadBody" name="AdvSearchHeadBody" style="WIDTH: 640px">
<table class="srch_bg" style="MARGIN-TOP: 0px" cellPadding="1" width=100% cellSpacing="0" border="0">  
<tr><td>
<form name = 'frmAddCommission' method = 'post' action = 'AddCommissionType.asp' >
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td width="18%"> Rate</td>
    <td width="82%"><input type = 'text' name ='txtRate' id = 'txtRate' size="20"></td>
  </tr>
  <tr>
    <td width="18%"> Description</td>
    <td width="82%"><input type = 'text' name ='txtDescription' id = 'txtDescription' size="20"></td>
  </tr>
  <tr>
    <td width="18%"><input type = 'submit' name ='cmdAdd' id = 'cmdAdd' value="Save"></td>
    <td width="82%"><input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
      </td>
  </tr>
  <tr>
    <td width="18%"><input type = 'button' name ='cmdEdit' id = 'cmdEdit' value="Edit Existing " OnClick="JavaScript: window.location.replace('EditCommissionList.asp')"></td>
    <td width="82%"></td>
  </tr>
  <tr>
    <td width="18%"><input type = 'button' name ='cmdDelete' id = 'cmdDelete' value="Delete" OnClick="JavaScript: window.location.replace('DeleteCommissionList.asp')"></td>
    <td width="82%"></td>
  </tr>
</table>
</form>


</td>
</tr>
</table>
</div>
</div>


</body>

