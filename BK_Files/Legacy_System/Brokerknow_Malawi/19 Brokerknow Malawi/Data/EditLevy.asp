
<!--#include file="../libroutines.asp"-->

<%
	
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim ID
	Dim rsEdit
	Dim guidStr 
	Dim guid
	
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
       Dim levyType
       Dim amount
       Dim block
       Dim app2Bond
       Dim app2Sec
       Dim active
       Dim i
       Dim affectedSec
	   Dim Vatable
       
       description = Request.Form("txtDescription")
       levyType = Request.Form("cboType")
       amount = Request.Form("txtAmount")
       block = Request.Form("txtBlock")
       app2Sec = Request.Form("cboAppSecurity")
       app2Bond = Request.Form("cboAppBond")
       active = Request.Form("cboActive")
       ShortName = Request.Form("txtShortName")
	   Vatable = Request.Form("cboVatable")
       
       
       'fetch affected securities
		set affectedSec = Request.Form("SecuritiesSel")
		
		
        'validate Amount
        If Trim(amount) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Amount/Percentage"
                		
                </script>
                <% 
				Reloadpage(ID)
				response.end
        End If
        
        'validate Apply To Bond
        If Trim(app2Bond) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the levy applies to Bonds"
                		
                </script>
                <% 
				Reloadpage(ID)
				response.end
        End If
        
        'validate Apply To Security
        If Trim(app2Sec) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the levy applies to Securities"
                		
                </script>
                <% response.end
        End If
       'validate Active
        If Trim(Active) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please whether the levy is active"
                		
                </script>
                <% 
				Reloadpage(ID)
				response.end
        End If
       'validate Description
        If Trim(Description) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Description"
                	</script>
                <% 
				Reloadpage(ID)
				response.end
        End If
        
        'validate Type
        If Trim(levyType) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the levy Type"
                		
                </script>
                <% 
				Reloadpage(ID)
				response.end
        End If
        
        'validate Block
        If Trim(block) = "" Then
			block = 0
		end if

		If trim(vatable) = "" or len(trim(vatable))=0 Then
               vatable=0
        End If
		'ensure Amount is numeric
        If (Amount <> "") And (Not IsNumeric(Amount)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Amount  must be numeric"
                
                </script>
                <% 
				Reloadpage(ID)
				response.end
        End If
        'ensure Block is numeric
        If (Block <> "") And (Not IsNumeric(Block)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Block  must be numeric"
                
                </script>
                <% 
				Reloadpage(ID)
				response.end
        End If
        'validate size of Description
        If Len(Description) > 100 Then%>
                <script language = 'vbscript'>
                ShowMessage "Description can only be 100 characters in length"
                
                </script>
                <% 
				Reloadpage(ID)
				response.end
        End If
        'validate size of Type
        If Len(levyType) > 2 Then%>
                <script language = 'vbscript'>
                ShowMessage "Type can only be 2 characters in length"
                
                </script>
                <% 
				Reloadpage(ID)
				response.end
        End If
        
        If Trim(levyType) = "S" Then
                If block < 1 Then%>
						<script language = 'vbscript'>
                				ShowMessage "Please specify the Block"
                				
						</script>
						<% response.end
				End If
        End If
        'validate short name
        If Len(ShortName) > 100 Then%>
                <script language = 'vbscript'>
                ShowMessage "Short Name can only be 100 characters in length"                
                </script>
                <% 
				Reloadpage(ID)
				response.end
        End If 
        
        Set conn = GetActiveConnection("KBroker")
        
        'save data
        
        sqlStr = "UPDATE [Levy] SET LevyActive = " & " " & Active & _
				"		,LevyAmount = " & " " & Amount & " " & _
                "       ,LevyAppBond = " & " " & App2Bond & " " & _
				"		,LevyAppSecurity = " & " " & App2Sec & " " & _
                "       ,LevyBlock = " & " " & Block & " " & _
                "		,LevyDescription = " & "'" & Description & "'" & _
                "       ,LevyType = " & "'" & LevyType & "'" & _
                "		,LevyShortName = '" & ShortName & "'" & _ 
				"		,Vatable = '" & Vatable & "'" & _ 
                "		 WHERE Levy_DPA_  = " & ID
        		 

        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
                
                'clear current securities
				sqlStr = "DELETE FROM LevySecurity WHERE Levy_DPA_ = " & ID
				conn.Execute sqlStr
				
				set guid = server.createobject("NDUtils.CGUID")
				'save affected securities
				for i = 1 to affectedSec.count
						guidStr = guid.GenerateGUID
						sqlStr = "INSERT INTO [LevySecurity] (LevySecurity_DPA_, LevySecurity_EIT_, Security_DPA_,Levy_DPA_) SELECT " & _
								"	" & " " & "iif(isnull(max([LevySecurity_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'LevySecurity'),max([LevySecurity_DPA_]) + 1)" & " " & " as LevySecurity_DPA_" & _
								"	," & "'" & guidStr & "'" & " as LevySecurity_EIT_" & _
								"	," & " " & affectedSec(i) & " " & " as Security_DPA_" & _
								"	," & " " & ID & " " & " as Levy_DPA_ " & _ 
								"	," & " " & vatable & " " & " as vatable " & _ 
								"	FROM [LevySecurity]"
						conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
				next
        conn.CommitTrans
        
        Set conn = Nothing
        WritefraEnabledDialogCloseScript2
        Response.End
   	end If
%>

<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Levy</title>
   <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 
 
 <Script Language="VBScript">
	Function SelectForm
		For Each Thing In frmEditLevy
			If InStr(1, Thing.Name, "SecuritiesSel") > 0 Then
				SelectAll Thing
			End If
		Next
	End Function
 </Script>
 
 <script language='javascript'>
function  ShowSecurityList(theChk)
{
	if (theChk.checked)
	{
		document.frmMain.elements("cboAppSecurity").value = "1";
		document.getElementById('SecurityList').style.display = '';
		document.getElementById('SecurityListTitle').style.display = '';
		document.getElementById('SecurityListCaption').style.display = '';
	}
	else
	{
	   document.frmMain.elements("cboAppSecurity").value = "0";
	   document.getElementById('SecurityList').style.display = 'none';
	   document.getElementById('SecurityListTitle').style.display = 'none';
	   document.getElementById('SecurityListCaption').style.display = 'none';
	}
	
}

function  UpdateBond(theChk)
{
	if (theChk.checked)
	{
		document.frmMain.elements("cboAppBond").value = "1";
	}
	else
	{
	   document.frmMain.elements("cboAppBond").value = "0";
	}
	
}
		
//===========BEGIN MOVE FUNCTION FOR SELECTED STAFF============================= 
function Move(Btn)
{
	var todo = Btn.value;
	var Users = document.all.item("SecuritiesAvail");
	var loop;
	var InsertList;

	if (todo.search(">")>0)
	{
		InsertList = document.all.item("SecuritiesSel") ;	
		if (Users.selectedIndex==-1) return(ShowMessage("Select a security from the list."))  
		 AddOption(Users,InsertList) 
	}
	else
	{
		RemoveOption(document.all.item("SecuritiesSel"))
	}	 
}


//=========END FUNCTION ====================================================

//=========BEGIN DROP-DOWN SELECT FUNCTION FOR FORM POSTING======

 function SelectAll(Object){
 //select all upwards
  for (loop=Object.length-1; loop>-1;
   loop--)
    {
     Object.options[loop].selected = true
    }
  }

//============END SELECT FUNCTION===================================

//==========BEGIN REMOVE OPTION/S FROM DROP-DOWN FUNCTION ON THE FLY=====
 function RemoveOption(Field){
	Selection = new Boolean();
	if (Field.length==0) return(ShowMessage("The list is empty."))
	for (loop=Field.length - 1; loop >= 0; loop--) {
	    var GoneOption = Field.options[loop]
		if (GoneOption.selected==true) {
	      		Selection = true;
	      		Field.remove(GoneOption.index);
	      }
	    }
	    
   if (Selection==false) ShowMessage("Select a security to remove from the List.")
   
  }

//==============END REMOVE OPTION/S FUNCTION====================

//=========BEGIN ADD OPTION TO DROP-DOWN ON THE FLY=============

  function AddOption(Input,Output){    
    for (loop=0; loop < Input.length; loop++){
    		if (Input.options[loop].selected && loop !== 0){
    		    NewOption = new Option();   			    
			    NewOption.text = Input.options[loop].text;
			    NewOption.value = Input.options[loop].value;		
			    //NewOption.ContractAmount = Input.options[loop].ContractAmount;		
			    NewOption.selected = false;	
			    if (!CheckDuplicates(Output, NewOption.value)) Output.add(NewOption, 0)
    		}
    		
    }
    
  }

function CheckDuplicates(DupPut, valText){
	var loop;
   for (loop=0; loop < DupPut.length;loop++){
      if (DupPut.options[loop].value==valText){
	       	return(true) ;
       }
     }
 }   
   
  
//========END ON THE FLY ADD OPTION FUNCTION=====================

function UpdateListingDisplay()
{
	UpdateBond(document.frmMain.elements("chkSecurity"));
	UpdateBond(document.frmMain.elements("chkBond"));
 }
		
		function forceSubmit()
		{
			setOpener();
			//var targetPage = window.dialogArguments.opener.document.all.item("frmMain").elements("EditPage").value;
					
			document.frmEditLevy.method='post';
			document.frmEditLevy.target='_self';
			document.frmEditLevy.submit();		
		}
		
		function setOpener()
		{
			window.self.opener = window.dialogArguments.opener;					
		}
</script>

</head>

<body Class="Dialog" onload="UpdateListingDisplay();setOpener()">
<form id='frmMain' name = 'frmEditLevy' method = 'post' action = 'EditLevy.asp' action="deleteFrame" OnSubmit="JavaScript: UpdateDialogHandle();">
<table border="0" width="100%" cellspacing="1" cellpadding="1">
<%
        Set conn = GetActiveConnection("KBroker")
       
        sqlStr = "SELECT LevyShortName, LevyActive,LevyAmount,LevyAppBond,LevyAppSecurity,LevyBlock" & _
                "       ,LevyDescription,LevyType,Levy_DPA_,vatable FROM [Levy] WHERE Levy_DPA_  = " & ID
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected Levy cannot be retrieved for editing"
                		
                </script>
                <% response.end
        End If

	
%>
  <tr>
    <td width="17%"> Description</td>
    <td width="83%"><input type = 'text' name ='txtDescription' id = 'txtDescription' size="20" value = '<%=rs.Fields("LevyDescription")%>'></td>
  </tr>
   <tr>
    <td width="18%" >Short Name</td>
    <td width="82%" ><input type = 'text' name ='txtShortName' id = 'txtShortName' size="20" value = '<%=rs.Fields("LevyShortName")%>'></td>
  </tr>
  <tr>
    <td width="17%"> Type</td>
    <td width="83%"><b>

<select name = 'cboType' id = 'cboType' size="1">
<%
		Dim default
		Dim defaultVal
		Dim other
		Dim otherVal
        
        if rs.Fields("LevyType") = "P" then
        		default = "Percentage"
        		other = "Schedule"
        else
        		default = "Schedule"
        		other = "Percentage"
        end if%>
       <option selected value = '<%=left(default,1)%>'><%=default%></option>
       <option value = '<%=left(other,1)%>'><%=other%></option>
    </select></b></td>
  </tr>
  <tr>
    <td width="17%"> Amount/Percentage</td>
    <td width="83%"><input type = 'text' name ='txtAmount' id = 'txtAmount' size="20" value = '<%=rs.Fields("LevyAmount")%>'></td>
  </tr>
  <tr>
    <td width="17%"> Block</td>
    <td width="83%"><input type = 'text' name ='txtBlock' id = 'txtBlock' size="20" value = '<%=rs.Fields("LevyBlock")%>'></td>
  </tr> 
  <tr>

  <%
  
  if rs.Fields("Vatable")=1 then 
	strchecked = "checked" 
	applicableVat = "1"
	else 
	strchecked =  ""
	applicableVat = "1"
	end if 
  
  %>
  <td width="17%"> VAT</td>
    <td width="83%"><input type = 'checkbox' name ='cboVatable' <% = strchecked %> id = 'chkvat' size="3" value =<% = applicableVat %>  ></td>
	</tr>
   <tr>
    <td width="19%" > Apply To Security</td>
    <td width="81%" colspan="3" >
    <%	Dim chkState
		Dim applicable
    
		if rs.Fields("LevyAppSecurity") then
			chkState = "checked"
			applicable = "1"
		else
			chkState = ""
			applicable = "0"
		end if
	%>
    <input type=checkbox   <%=chkState%> name='chkSecurity' onClick = 'ShowSecurityList(this);'>
    <input type = 'hidden' name ='cboAppSecurity' id = 'cboAppSecurity' value='<%=applicable%>'>
    
       
</tr>
  <tr id="SecurityListTitle">
    <td width="43%" colspan="2" > Specify Affected Securities:</td>
    <td width="57%" colspan="2" >
    
  </tr>
  <tr id="SecurityListCaption">
    <td width="19%" > &nbsp;</td>
    <td width="24%" >
    
    Available Securities
    <td width="11%" >
    
    
    <td width="46%" >
    
    Affected Securities

  </tr>
  <tr id="SecurityList">
    <td width="19%" > &nbsp;</td>
    <td width="24%" >
    <select size="10" style="width:200px" name="SecuritiesAvail" multiple id="SecuritiesAvail">
<%
	Dim displayValue
	Dim AvailsecurityRs
	
	
	sqlStr = "SELECT * FROM SecurityList " 'WHERE " & _
           ' "  Security_DPA_ NOT IN(SELECT Security_DPA_ FROM LevySecurity WHERE " & _
           ' "  (Levy_DPA_ = " & rs.fields("Levy_DPA_") & "))"
            
    set AvailsecurityRs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
	If Not (AvailsecurityRs.EOF Or AvailsecurityRs.BOF) Then
		Do Until AvailsecurityRs.EOF
			displayValue = AvailsecurityRs("SecurityCode") & " : " & AvailsecurityRs("SecurityName")
			%>
			<Option Value="<%= AvailsecurityRs("Security_DPA_")  %>"><%= displayValue %></Option>
			<%
			AvailsecurityRs.MoveNext
		Loop
	End If
%>
	</select>
    <td width="11%" >
    
    <input type="button" value=" &lt; " name="MoveBtn" Class=Buttons OnClick="JavaScript: Move(this)">&nbsp;&nbsp;&nbsp;<input type="button" value=" &gt; " name="MoveBtn" Class=Buttons OnClick="JavaScript: Move(this); ">
    <td width="46%" >
    
    <select size="10" style="width:200px" name="SecuritiesSel" multiple  OnKeyPress="JavaScript: if (event.keyCode==46) Move(this)" id="SecuritiesSel">
	<%
	Dim SelsecurityRs
	
	
	sqlStr = "SELECT * FROM SecurityList WHERE " & _
            "  Security_DPA_ IN(SELECT Security_DPA_ FROM LevySecurity WHERE " & _
            "  (Levy_DPA_ = " & rs.fields("Levy_DPA_") & "))"
            
    set SelsecurityRs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
	If Not (SelsecurityRs.EOF Or SelsecurityRs.BOF) Then
		Do Until SelsecurityRs.EOF
			displayValue = SelsecurityRs("SecurityCode") & " : " & SelsecurityRs("SecurityName")
			%>
			<Option Value="<%= SelsecurityRs("Security_DPA_")  %>"><%= displayValue %></Option>
			<%
			SelsecurityRs.MoveNext
		Loop
	End If
			%>
	</select>
  </tr>
  <tr>
    <td width="19%" > Apply To Bond</td>
    <td width="81%" colspan="3" >
    
    <%    
		if rs.Fields("LevyAppBond") then
			chkState = "checked"
			applicable = "1"
		else
			chkState = ""
			applicable = "0"
		end if
	%>
    <input type=checkbox <%=chkState%> name='chkBond' onClick = 'UpdateBond(this);'>
    <input type = 'hidden' name ='cboAppBond' id = 'cboAppBond' value='<%=applicable%>'>
    
  </tr>
  <tr>
    <td width="17%"> Active</td>
    <td width="83%"><b><select name = 'cboActive' id = 'cboActive' size="1">
    	<%
        if rs.Fields("LevyActive") then
        		default = "Yes"
        		defaultVal = 1
        		other = "No"
        		otherVal = 0
        else
        		default = "No"
        		defaultVal =  0
        		other = "Yes"
        		otherVal = 1
        end if%>
       <option selected value = '<%=defaultVal%>'><%=default%></option>
       <option value = '<%=otherVal%>'><%=other%></option>
    </select></b></td>
  </tr>
  <tr>
   
     <td width="100%" colspan="2" align=right>
		<BR>
		<BR>		
		<input type = 'button' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save " OnClick="Javascript: forceSubmit();" OnClick="VBScript: SelectForm" >
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
		&nbsp;&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='buttonAction' id = 'buttonAction' value="Save">
    	<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
      </td>
  </tr>
</table>
</form>

</body>

</html>
