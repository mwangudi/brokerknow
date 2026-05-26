<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Order Type</title>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
 <script language="JavaScript" src="../scripts/common.js"></script>

</head>

<body Class="Dialog">
<!--#include file="../libroutines.asp"-->
<%
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
	
	action = ucase(Request.Form("action"))
	
	if action = "EXECUTE" then
		Dim description
		Dim saleType
		Dim handlingfee
       
       description = Request.Form("txtDescription")
       saleType = Request.Form("cboSaleType")
	   handlingfee = Request.Form("txthandlingfee")
      
       
        'validate Description
        If Trim(Description) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Description"
                		
                </script>
                <% response.end
        End If
        'validate size of Description
        If Len(Description) > 100 Then%>
                <script language = 'vbscript'>
                ShowMessage "Description can only be 100 characters in length"
                
                </script>
                <% response.End
                If trim(handlingfee) = "" or len(trim(handlingfee))=0 Then
               handlingfee=0
        End If
        End If
        'validate Sale Type
        If Trim(saleType) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Sale Type"
                		
                </script>
                <% response.end
        End If
       
        'save data
        sqlStr = "INSERT INTO [OrderType] (OrderTypeSale, OrderTypeDescription,OrderType_DPA_) SELECT " & " " & saleType & " " & " as OrderTypeSale," & "'" & description & "'" & " as OrderTypeDescription" & _
                "," & " " & "iif(isnull(max([OrderType_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'OrderType'),max([OrderType_DPA_]) + 1)" & " " & " as OrderType_DPA_,Handlingfee" & _
                " FROM [OrderType]"

				
        Set conn = GetActiveConnection("KBroker")
        
        conn.BeginTrans
        
                conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
        conn.CommitTrans
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript
        Response.End
   	end If
%>


<form name = 'frmAddOrderType' method = 'post' action = 'AddOrderType.asp' target="deleteFrame" OnSubmit="JavaScript: UpdateDialogHandle();">
<table border="0" width="100%">
  <tr>
   <!-- <td width="17%"> Description</td>
    <td width="83%"><input type = 'text' name ='txtDescription' id = 'txtDescription' size="20" value = 
	'<%=rs.Fields("OrderTypeDescription")%>'></td>-->
  </tr>
     <tr>
    <td width="17%"> Handling Fee</td>
    <td width="83%"><input type = 'text' name ='txthandlingfee' id = 'txthandlingfee' size="20" ></td>
  </tr>
  <tr>
    <td width="18%">Sale type</td>
    <td width="82%"><b><select name = 'cboSaleType' id = 'cboSaleType' size="1">
    	<option selected value = ''></option>
    	<option value = '1'>Yes</option>
    	<option value = '0'>No</option>
    </select></b></td>
  </tr>
  <tr>
     <td width="100%" colspan="2" align=right>
		<BR>
	
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
		&nbsp;&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
      </td>
  </tr>
 
</table>
</form>

</body>

</html>
