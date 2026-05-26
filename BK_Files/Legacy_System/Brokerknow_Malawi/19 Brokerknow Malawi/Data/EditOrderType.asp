<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Order Type</title>
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
                <% response.end
        End If
		If trim(handlingfee) = "" or len(trim(handlingfee))=0 Then
               handlingfee=0
        End If
        'validate Sale Type
        If Trim(saleType) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Sale Type"
                		
                </script>
                <% response.end
        End If


        if ucase(saleType) = "YES" then saleType = 1 else saleType = 0
        
		Set conn = GetActiveConnection("KBroker")
        
        'save data
        
        sqlStr = "UPDATE [OrderType] SET OrderTypeSale = " & " " & saleType & ",handlingfee="  &  handlingfee & " " & ", OrderTypeDescription = " & "'" & description & "'" & " WHERE OrderType_DPA_  = " & ID
       Set conn = GetActiveConnection("KBroker")
        
        conn.BeginTrans
		        
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
               Set conn = Nothing
        WriteFraEnabledDialogCloseScript
        response.end
   	end If
%>

<form name = 'frmEditOrderType' method = 'post' action = 'EditOrderType.asp' target="deleteFrame" OnSubmit="JavaScript: UpdateDialogHandle();">
<table border="0" width="100%">
<%
        Set conn = GetActiveConnection("KBroker")
        
        
        sqlStr = "SELECT OrderTypeSale,OrderTypeDescription,OrderType_DPA_,handlingfee FROM [OrderType] WHERE OrderType_DPA_  = " & ID        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected Order type cannot be retrieved for editing"
                		
                </script>
                <% response.end
        End If
		
		%>
  <tr>
    <td width="17%"> Description</td>
    <td width="83%"><input type = 'text' name ='txtDescription' id = 'txtDescription' size="20" value = '<%=rs.Fields("OrderTypeDescription")%>'></td>
  </tr>
 
  <tr>
    <td width="17%">Sale type</td>
    <td width="83%"><select name = 'cboSaleType' id = 'cboSaleType' size="1">
<%
		Dim default
		Dim other
        
        if rs.Fields("OrderTypeSale") then
        		default = "Yes"
        		other = "No"
        else
        		default = "No"
        		other = "Yes"
        end if%>
       <option selected value = '<%=default%>'><%=default%></option>
       <option value = '<%=other%>'><%=other%></option>
    </select></td>
  </tr>
 <tr>
    <td width="17%"> Handling Fee</td>
    <td width="83%"><input type = 'text' name ='txthandlingfee' id = 'txthandlingfee' size="5"value = '<%=rs.Fields("handlingfee")%>' ></td>
  </tr>
  <tr>
	<td width="100%" colspan="2" align=right>
		<BR>
		<BR>		
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
		&nbsp;&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
    	<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
      </td>
    
  </tr>
</table>
</form>

</body>

</html>
