
<!--#include file="../libroutines.asp"-->
<%
	
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
    Dim guidStr 
	Dim guid
	
	action = ucase(Request.Form("action"))
	
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
	   'Vatable = Request.Form("cboVatable")
     'response.write app2Sec & "<br>"
		'fetch affected securities
		set affectedSec = Request.Form("SecuritiesSel")
       
        'validate Amount
        If Trim(amount) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Amount/Percentage"
                		
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If
        
        'validate Apply To Bond
        If Trim(app2Bond) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the levy applies to Bonds"
                		
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If
        
        'validate Apply To Security
        If Trim(app2Sec) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the levy applies to Securities"
                		
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If
       'validate Active
        If Trim(Active) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please whether the levy is active"
                		
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If
       'validate Description
        If Trim(Description) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Description"
                		
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If
        
        'validate Type
        If Trim(levyType) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the levy Type"
                		
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If
        
        'validate Block
        If Trim(block) = "" Then
			block = 0
		end if
		'ensure Amount is numeric
        If (Amount <> "") And (Not IsNumeric(Amount)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Amount  must be numeric"
                
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If
        'ensure Block is numeric
        If (Block <> "") And (Not IsNumeric(Block)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Block  must be numeric"
                
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If
        'validate size of Description
        If Len(Description) > 100 Then%>
                <script language = 'vbscript'>
                ShowMessage "Description can only be 100 characters in length"
                
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If
        'validate size of Type
        If Len(levyType) > 2 Then%>
                <script language = 'vbscript'>
                ShowMessage "Type can only be 2 characters in length"
                
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If
        
        If Trim(levyType) = "S" Then
                If block < 1 Then%>
						<script language = 'vbscript'>
                				ShowMessage "Please specify the Block"
                				
						</script>
						<% 
						ReloadPage(ID)
						response.end
				End If
        End If

		 'If trim(vatable) = "" or len(trim(vatable))=0 Then
              ' vatable=0
       ' End If
        
        'validate short name
        If Len(ShortName) > 100 Then%>
                <script language = 'vbscript'>
                ShowMessage "Short Name can only be 100 characters in length"                
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If 
        
        
		set guid = server.createobject("NDUtils.CGUID")
		guidStr = guid.GenerateGUID
       
        'save data
       sqlStr = "INSERT INTO [Levy] (LevyActive,LevyAmount,LevyAppBond,LevyAppSecurity,Levy_EIT_,LevyShortName,LevyBlock,LevyDescription" & _
                ",LevyType,Levy_DPA_,SystemMaintained) SELECT " & " " & Active & " " & " as LevyActive," & " " & amount & " " & " as LevyAmount," & " " & app2Bond & " " & " as LevyAppBond" & _
                "," & " " & app2Sec & " " & " as LevyAppSecurity" & _
                "," & "'" & guidStr & "'" & " as Levy_EIT_" & _
                "," & "'" & ShortName & "'" & " as LevyShortName" & _
                "," & " " & block & " " & " as LevyBlock," & "'" & description & "'" & " as LevyDescription" & _
                "," & "'" & levyType & "'" & " as LevyType" & _
                "," & " " & "iif(isnull(max([Levy_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Levy'),max([Levy_DPA_]) + 1)" & " " & " as Levy_DPA_" & _
                "," & " " & "(SELECT iif(isnull(max([LevyContractSysMaintainedList.SystemMaintained])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'LevySysMaintained'),max([LevyContractSysMaintainedList.SystemMaintained]) + 1) AS SystemMaintained FROM LevyContractSysMaintainedList)"  & " as SystemMaintained" & _
                " FROM [Levy]	," & " " & "(SELECT iif(isnull(max([LevyContract.SystemMaintained])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'LevySysMaintained'),max([LevyContract.SystemMaintained]) + 1) AS SystemMaintained FROM LevyContract)" & " " & " as SystemMaintained ," & Vatable & " As vatable" 


	  

        sqlStr = "INSERT INTO [Levy] (LevyActive,LevyAmount,LevyAppBond,LevyAppSecurity,Levy_EIT_,LevyShortName,LevyBlock,LevyDescription" & _
                ",LevyType,Levy_DPA_,SystemMaintained,vatable) SELECT " & " " & Active & " " & " as LevyActive," & " " & amount & " " & " as LevyAmount," & " " & app2Bond & " " & " as LevyAppBond" & _
                "," & " " & app2Sec & " " & " as LevyAppSecurity" & _
                "," & "'" & guidStr & "'" & " as Levy_EIT_" & _
                "," & "'" & ShortName & "'" & " as LevyShortName" & _
                "," & " " & block & " " & " as LevyBlock," & "'" & description & "'" & " as LevyDescription" & _
                "," & "'" & levyType & "'" & " as LevyType" & _
                "," & " " & "iif(isnull(max([Levy_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Levy'),max([Levy_DPA_]) + 1)" & " " & " as Levy_DPA_" & _
                ","  & "(SELECT iif(isnull(max([LevyContractSysMaintainedList.SystemMaintained]),0),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'LevySysMaintained'),max([LevyContractSysMaintainedList.SystemMaintained]) + 1) AS SystemMaintained FROM LevyContractSysMaintainedList)" & " " & " as SystemMaintained," &  Vatable & " As vatable"  & _    
				" FROM [Levy]"

				sqlStr = "INSERT INTO [Levy] (LevyActive,LevyAmount,LevyAppBond,LevyAppSecurity,Levy_EIT_,LevyShortName,LevyBlock,LevyDescription" & _
                ",LevyType,Levy_DPA_,SystemMaintained)  SELECT " & " " & Active & " " & " as LevyActive," & " " & amount & " " & " as LevyAmount," & " " & app2Bond & " " & " as LevyAppBond" & _
                "," & " " & replace(app2Sec,",","") & " " & " as LevyAppSecurity" & _
                "," & "'" & guidStr & "'" & " as Levy_EIT_" & _
                "," & "'" & ShortName & "'" & " as LevyShortName" & _
                "," & " " & block & " " & " as LevyBlock," & "'" & description & "'" & " as LevyDescription" & _
                "," & "'" & levyType & "'" & " as LevyType" & _
                "," & " " & "iif(isnull(max([Levy_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Levy'),max([Levy_DPA_]) + 1)" & " " & " as Levy_DPA_" & _
                "," & " " & "(SELECT iif(isnull(max([LevyContractSysMaintainedList.SystemMaintained])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'LevySysMaintained'),max([LevyContractSysMaintainedList.SystemMaintained]) + 1) AS SystemMaintained FROM LevyContractSysMaintainedList)" & " " & " as SystemMaintained" & _
                " FROM [Levy]"
           '"," & " " & "(SELECT iif(isnull(max([LevyContract.SystemMaintained])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'LevySysMaintained'),max([LevyContract.SystemMaintained]) + 1) AS SystemMaintained FROM LevyContract)" & " " & " as SystemMaintained" 
			
Set conn = GetActiveConnection("KBroker")
set rsSysMaintained=conn.execute("SELECT     ISNULL(MAX(SystemMaintained), 0) + 1 as SystemMaintained FROM LevyContractSysMaintainedList")
sysmaintained=rsSysMaintained("SystemMaintained")
set rsSysMaintained = Nothing

		   sqlStr = "INSERT INTO [Levy] (LevyActive,LevyAmount,LevyAppBond,LevyAppSecurity,Levy_EIT_,LevyShortName,LevyBlock,LevyDescription" & _
                ",LevyType,SystemMaintained)  values(" &  Active & " " & " ," & " " & amount & " " & " ," & " " & app2Bond & _
                "," & " " & replace(app2Sec,",","")  & _
                "," & "'" & guidStr & "'"  & _
                "," & "'" & ShortName & "'"  & _
                "," & " " & block & " " & " ," & "'" & description & "'"  & _
                "," & "'" & levyType & "'"  & _
                "," & " " &  sysmaintained &")"
		'response.write(sqlstr)
	'response.end
	%><script language='vbscript' > 'inputbox "title"," sasa", "<%=sqlstr%>" </script><%
		
		
        
        conn.BeginTrans
                sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
               
                conn.Execute sqlStr
                
                'obtain header key value
				sqlStr = "SELECT Levy_DPA_, LevyShortName, SystemMaintained FROM Levy WHERE Levy_EIT_ = " & "'" & guidStr & "'"
				     
				Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
				If (rs.EOF Or rs.BOF) Then%>
				 			<script language = 'vbscript'>
				 					ShowMessage "A serious error has been encountered while saving the data. Try saving again"
				         					
				 			</script>
				 			<% response.end
				End If
				
				'save affected securities
				for i = 1 to affectedSec.count
						guidStr = guid.GenerateGUID
						sqlStr = "INSERT INTO [LevySecurity] (LevySecurity_DPA_, LevySecurity_EIT_, Security_DPA_,Levy_DPA_ ) SELECT " & _
								"	" & " " & "iif(isnull(max([LevySecurity_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'LevySecurity'),max([LevySecurity_DPA_]) + 1)" & " " & " as LevySecurity_DPA_" & _
								"	," & "'" & guidStr & "'" & " as LevySecurity_EIT_" & _
								"	," & " " & affectedSec(i) & " " & " as Security_DPA_" & _
								"	," & " " & rs.Fields("Levy_DPA_") & " " & " as Levy_DPA_" & _
                               "	," & " " & vatable("vatable") & " " & " as vatable" & _
								"	FROM [LevySecurity]"
							 
						conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
				next
				
				'create an account for the levy
				sqlStr = "INSERT INTO [Entity] (EntityGeneric1,EntityGeneric2,EntityGeneric3,EntityName" & _
						"       ,Entity_DPA_,EntityType_DPA_,EntityOpeningBal,LevySystemMaintained,SystemMaintained) SELECT " & "''" & " as EntityGeneric1" & _
						"       ," & "''" & " as EntityGeneric2" & _
						"       ," & "''" & " as EntityGeneric3" & _
						"       ," & "'" & rs.Fields("LevyShortName") & "'" & " as EntityName" & _
						"		," & " " & "iif(isnull(max([Entity_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Entity'),max([Entity_DPA_]) + 1)" & " " & " as Entity_DPA_" & _
						"       ," & " 6 " & " as EntityType_DPA_" & _
						"		," & " 0 " & " as EntityOpeningBal" & _
						"		," & " " & rs.Fields("SystemMaintained") & " " & " as LevySystemMaintained" & _
						"		," & " 1 " & " as SystemMaintained" & _
						"        FROM [Entity]"			
							
				sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
				
				conn.Execute sqlStr
        conn.CommitTrans
        conn.Close
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
<title>Add Levy</title>
  <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 
 
 <Script Language="VBScript">
	Function SelectForm
		For Each Thing In frmAddLevy
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
	//alert(document.frmMain.elements("cboAppSecurity").value );
	
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


function forceSubmit()
		{
			setOpener();
			//var targetPage = window.dialogArguments.opener.document.all.item("frmMain").elements("EditPage").value;
					
			document.frmAddLevy.method='post';
			document.frmAddLevy.target='_self';
			document.frmAddLevy.submit();		
		}
		
		function setOpener()
		{
			window.self.opener = window.dialogArguments.opener;					
		}

</script>
</head>

<body Class="Dialog" onload="setOpener()">
<form id='frmMain' name = 'frmAddLevy' method = 'post' action = 'AddLevy.asp' target="deleteFrame" OnSubmit="JavaScript: UpdateDialogHandle();">
<table border="0" width="100%" cellspacing="1" cellpadding="1">
  <tr>
    <td width="19%" > Description</td>
    <td width="81%" colspan="3" ><input type = 'text' name ='txtDescription' id = 'txtDescription' size="20"></td>
  </tr>
  <tr>
    <td width="19%" >Short Name</td>
    <td width="81%" colspan="3" ><input type = 'text' name ='txtShortName' id = 'txtShortName' size="20"></td>
  </tr>
 
  <tr>
    <td width="19%" > Type</td>
    <td width="81%" colspan="3" >
<b>

<select name = 'cboType' id = 'cboType' size="1">
    	<option selected value = 'P'>Percentage</option>
    	<option value = 'S'>Schedule</option>
    </select></b></td>
  </tr>
  <tr>
    <td width="19%" > Amount/Percentage</td>
    <td width="81%" colspan="3" ><input type = 'text' name ='txtAmount' id = 'txtAmount' size="20"></td>
  </tr>
  <tr>
    <td width="19%" > Block</td>
    <td width="81%" colspan="3" ><input type = 'text' name ='txtBlock' id = 'txtBlock' size="20"></td>
  </tr>
  <tr>
   <!-- <tr>
    <td>VAT</td>
    <td><input type = 'checkbox' name ='cboVatable' id = 'chkvat' size="20" value=1></td>
  </tr>-->
    <input type = 'hidden' name ='cboAppSecurity' id = 'cboAppSecurity' value=''>
  </tr>
  <tr>
    <td width="19%" > Apply To Security</td>
    <td width="81%" colspan="3" >
    <input type=checkbox  Class="BorderLess"   checked name='chkSecurity' onClick = 'ShowSecurityList(this);'>
    <input type = 'hidden' name ='cboAppSecurity' id = 'cboAppSecurity' value='1'>
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
			Set conn = GetActiveConnection("KBroker")
    
			sqlStr = "SELECT * FROM SecurityList WHERE (OrderSecType_DPA_ = 2)"
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open sqlStr, Conn.ConnectionString, adOpenKeySet, adLockOptimistic
			
			If Not (rs.BOF Or rs.EOF) Then%>
				<option >Select a security to add</option>
				<%
				Do Until rs.EOF
					displayValue = rs("SecurityCode") & " : " & rs("SecurityName") 
					%>
					<Option Value="<%= rs("Security_DPA_")  %>"><%= displayValue %></Option>
					<%
					rs.MoveNext
				Loop
			End If
			Set rs = Nothing
			Set Conn = Nothing %>
	</select>
    <td width="11%" >
    
    <input type="button" value=" &lt; " name="MoveBtn" Class=Buttons OnClick="JavaScript: Move(this)">&nbsp;&nbsp;&nbsp;<input type="button" value=" &gt; " name="MoveBtn" Class=Buttons OnClick="JavaScript: Move(this); ">
    <td width="46%" >
    
    <select size="10" style="width:200px" name="SecuritiesSel" multiple  OnKeyPress="JavaScript: if (event.keyCode==46) Move(this)" id="SecuritiesSel">
	</select>

  </tr>
  <tr>
    <td width="19%" > Apply To Bond</td>
    <td width="81%" colspan="3" >
    
    <input type=checkbox  Class="BorderLess"  name='chkBond' onClick = 'UpdateBond(this);'>
    <input type = 'hidden' name ='cboAppBond' id = 'cboAppBond' value='0'>

  </tr>
  <tr>
    <td width="19%" > Active</td>
    <td width="81%" colspan="3" >
<b>

<select name = 'cboActive' id = 'cboActive' size="1">
    	<option selected value = '1'>Yes</option>
    	<option value = '0'>No</option>
    </select></b></td>
  </tr>
  <tr>
    <td width="100%" colspan="4" align=right>
		<BR>
		<input type = 'button' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save " OnClick="Javascript:forceSubmit(); " OnClick="VBScript: SelectForm " >
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
		&nbsp;&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
      </td>
  </tr>
 
</table>
</form>

</body>

