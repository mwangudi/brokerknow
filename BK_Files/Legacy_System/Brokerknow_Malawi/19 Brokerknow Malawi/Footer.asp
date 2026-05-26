<html>

<head><TITLE>BrokerKnow Toolbar (Bottom)</TITLE>
<LINK href="STYLE/default.css" type=TEXT/CSS rel=STYLESHEET> 
<LINK href="STYLE/webparts.css" type=TEXT/CSS rel=STYLESHEET>
<SCRIPT language=Javascript src="scripts/common.js"></SCRIPT>
<base target="maininfo">
</head>

<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">

<!--#include file="libroutines.asp"-->

<%
Dim Conn
Dim RS

UserID = Session("UserID")

If UserID = "" Then
	'session expired
	Response.Write "Session Expired"
	Response.End 
End If


mnuID = Request.QueryString("mnuID")

If mnuID = "" Then Response.End

Set Conn = GetActiveConnection("KBroker")
'Conn.Execute("UpdateClientBalances")

Conn.Close
Set Conn = Nothing
%>

<div>

<table width="100%" border="0" style="FONT-SIZE: 8pt; FONT-FAMILY: Tahoma" cellpadding="0" cellspacing="0">
  <TR>	
	<TD align=left valign="Bottom" width="50%">
		<Table>
			<TR valign="Bottom">
				<TD STYLE="display: none" ID="cancelSort" class="footerHighlightNavOn" onMouseover="JavaScript: this.bgColor='grey'; window.status=this.innerText" onMouseout="JavaScript: this.className='footerHighlightNavOn'; window.status=''" OnClick="JavaScript: cancelOperation(this)">
					&nbsp;Cancel Sort&nbsp;
				</TD>
				<TD STYLE="display: none" ID="cancelFilter" class="footerHighlightNavOn" onMouseover="JavaScript: this.className='nav_over'; window.status=this.innerText" onMouseout="JavaScript: this.className='footerHighlightNavOn'; window.status=''" OnClick="JavaScript: cancelOperation(this)">
					&nbsp;Cancel Filter&nbsp;
				</TD>
				<TD STYLE="display: none" ID="cancelSearch" class="footerHighlightNavOn" onMouseover="JavaScript: this.className='nav_over'; window.status=this.innerText" onMouseout="JavaScript: this.className='footerHighlightNavOn'; window.status=''" OnClick="JavaScript: cancelOperation(this)">
					&nbsp;Cancel Search&nbsp;
				</TD>
			</TR>
		</TABLE>
	</TD>
	
	<TD align=right nowrap width="50%">
		
		<TABLE border="0" style="FONT-SIZE: 8pt; FONT-FAMILY: Tahoma" cellpadding="2" height="40px" cellspacing="2">
			<TR>
			
				<%
				'-1 is reserverd for reports!
				'-2 is reserverd for imports!
				If (mnuID <> "-1") And (mnuID <> "-2") Then
					Set Conn = GetActiveConnection("KBroker")
					SQL = "SELECT MenuGroups.* FROM  UserGroups INNER JOIN " & _
								"                      MenuGroups ON UserGroups.GroupID = MenuGroups.groupID " & _
								"			WHERE     (UserGroups.UserID = " & userID & ") AND (MenuGroups.MenuID = " & mnuID & ")"

					Set Rs = Conn.Execute(SQL)
				
					If Not (Rs.EOF Or Rs.BOF) Then 
						canAdd = 0
						canDelete = 0
						canEdit = 0
						canSort = 0 
						canFilter = 0
						canSearch = 0
						
  						Do Until Rs.EOF	
  							If Rs.Fields("CanAdd").Value = "1" Then	canAdd = 1
  							If Rs.Fields("CanDelete").Value = "1" Then canDelete = 1
  							If Rs.Fields("CanEdit").Value = "1" Then canEdit = 1
  							If Rs.Fields("CanSort").Value = "1" Then canSort = 1
  							If Rs.Fields("CanFilter").Value = "1" Then canFilter = 1
  							If Rs.Fields("CanSearch").Value = "1" Then canSearch = 1
						
							Rs.MoveNext
							
						Loop
						
					 End If
					 
					 Set Rs = Nothing
					 Set Conn = Nothing
					 
				ElseIf (mnuID <> "-2") then
					'report special
					canAdd = 1
					canDelete = 0
					canEdit = 0
					canSort = 1 
					canFilter = 1
					canSearch = 1
				Else
					'import special
					canAdd = 0
					canDelete = 0
					canEdit = 0
					canSort = 1 
					canFilter = 1
					canSearch = 1
					
						 
				End If	 
	
				
				
				For i = 1 To 6
					mnuPermission = 0
  					Select Case i
  						Case 1
  							'add
  							Action = "JavaScript:  DoAdd()" 
							mnuPermission = canAdd 
							mnuCaption = "New"
							tdClientScript = "onMouseover=""JavaScript: this.className='nav_over'"" onMouseout=""JavaScript: this.className='footerHighlightnav'"""
						Case 2
  							'delete
  							Action = "JavaScript:  DoDelete()"
							mnuPermission = canDelete 
							mnuCaption = "Delete"
							tdClientScript = "onMouseover=""JavaScript: this.className='nav_over'"" onMouseout=""JavaScript: this.className='footerHighlightnav'"""
						Case 3
  							'edit  							
							action = "JavaScript:  DoEdit()"
							mnuPermission = canEdit 
							mnuCaption = "Edit"					
							tdClientScript = "onMouseover=""JavaScript: this.className='nav_over'"" onMouseout=""JavaScript: this.className='footerHighlightnav'"""
						Case 4
  							'sort  							
							action = "JavaScript:  DoSort()"
							mnuPermission = canSort 
							mnuCaption = "Sort"	
							tdClientScript = "onMouseover=""JavaScript: if (this.className=='footerHighlightnav') this.className='nav_over';"" onMouseout=""JavaScript: if (this.className=='nav_over') this.className='footerHighlightnav';""" 
						Case 5
  							'filter  							
							action = "JavaScript:  DoFilter()"
							mnuPermission = canFilter 
							mnuCaption = "Filter"	
							tdClientScript = "onMouseover=""JavaScript: if (this.className=='footerHighlightnav') this.className='nav_over';"" onMouseout=""JavaScript: if (this.className=='nav_over') this.className='footerHighlightnav';""" 
						Case 6
  							'search  							
							action = "JavaScript:  DoSearch()"
							mnuPermission = canSearch 
							mnuCaption = "Search"				
							tdClientScript = "onMouseover=""JavaScript: if (this.className=='footerHighlightnav') this.className='nav_over';"" onMouseout=""JavaScript: if (this.className=='nav_over') this.className='footerHighlightnav';""" 
					End Select	 
					
					If mnuPermission = 1 Then 	%>
						<TD OnClick="<%= Action %>" ID="<%= mnuCaption %>TD" nowrap width="20px" class="footerHighlightnav"  <%= tdClientScript %> >
							&nbsp;<b><%= mnuCaption  %></b>&nbsp;
						</TD>		
					<%					
					End If
					
				Next%>	
			</TR>
		</TABLE>	
			
	</TD>
  </TR>
</table>
</div>

<Div style="display: none">
	<IFRAME marginwidth="0" marginheight="0" FRAMEBORDER=0 SRC="loading.htm" ID="hiddenFrame" NAME="deleteFrame" TAG=""></IFRAME>	
</Div>

</BODY>

<Script Language="JavaScript">
	var dWidth, dHeight, editDocOpener=null;
	var ActionWin;
	var reportMode = "<%= mnuID %>";
	window.defaultStatus = window.parent.document.title;
	
	function searchDialog(){
		var opener;
		var searchArgs;		
		var searchTarget;
		var searchTD;
		var cancelButton;
	}
	function DoSearch(){
		try{
			var newWinWidth = "30em";
			var newWinHeight = "22em";			
			var searchArgs = window.parent.frames["maininfo"].document.all.item("frmMain").elements("searchArgs").value;	
			searchDialog.searchArgs = searchArgs;
			searchDialog.searchTD = document.all.item("SearchTD");
			searchDialog.cancelButton = document.all.item("cancelSearch") ;
			searchDialog.opener = window.parent.frames["maininfo"];	
			searchDialog.searchTarget = window.parent.frames["maininfo"].document.all.item("ActionPage").value;	    	      
			ActionWin = window.showModalDialog("Search/search.asp", searchDialog, "dialogWidth:" + newWinWidth + ";dialogHeight:" + newWinHeight + ";status:0;dialogHide:false;help:no;scroll:yes;resizable:no;edge:sunken;unadorned:yes");
		}
		catch(e){}
	}
	function filterDialog(){
		var opener;
		var filterArgs;
		var filterTarget;
		var filterTD;
		var cancelButton;
	}
	

	function DoFilter(){
		try{
			var newWinWidth = "25em";
			var newWinHeight = "16em";
			var filterArgs = window.parent.frames["maininfo"].document.all.item("frmMain").elements("filterArgs").value;
			filterDialog.filterArgs = filterArgs;
			filterDialog.filterTD = document.all.item("FilterTD");
			filterDialog.cancelButton = document.all.item("cancelFilter") ;
			filterDialog.opener = window.parent.frames["maininfo"];
			filterDialog.filterTarget = window.parent.frames["maininfo"].document.all.item("ActionPage").value;	    	    
			ActionWin = window.showModalDialog("Filter/filter.asp", filterDialog, "dialogWidth:" + newWinWidth + ";dialogHeight:" + newWinHeight + ";status:0;dialogHide:false;help:no;scroll:yes;resizable:no;edge:sunken;unadorned:yes");
		}
		catch(e){}
	}
	function sortDialog(){
		var opener;
		var sortArgs;
		var sortTarget;
		var sortTD;
		var cancelButton;
	}

	function DoSort(){
		try{
			var newWinWidth = "15em";
			var newWinHeight = "20em";
			var sortArgs = window.parent.frames["maininfo"].document.all.item("frmMain").elements("SortArgs").value;
			sortDialog.sortArgs = sortArgs;
			sortDialog.opener = window.parent.frames["maininfo"];
			sortDialog.sortTD = document.all.item("sortTD");
			sortDialog.cancelButton = document.all.item("cancelSort") ;
			sortDialog.sortTarget = window.parent.frames["maininfo"].document.all.item("ActionPage").value;	    
			ActionWin = window.showModalDialog("Sort/sort.asp", sortDialog, "dialogWidth:" + newWinWidth + ";dialogHeight:" + newWinHeight + ";status:0;dialogHide:false;help:no;scroll:yes;resizable:no;edge:sunken;unadorned:yes");
		}
		catch(e){}
	}
	
	function editDialog() {
		var selectedID;
		var opener;
	}
	
	
	function DoEdit()
	{
		try{
			var selID = window.parent.frames["maininfo"].document.all.item("frmMain").elements("ID").value
			if (selID=="") 
			{
				alert ("Please select a record to edit");
				return;
			}			
			editDialog.selectedID = selID;

			editDialog.opener = window.parent.frames["maininfo"];
			//var editDialog1 = window.parent.frames["maininfo"];
			var dialogTitle = getDialogTitle();  
			if (dialogTitle!=="") dialogTitle = "Edit " + dialogTitle;
			var targetPage = window.parent.frames["maininfo"].document.all.item("frmMain").elements("EditPage").value;
			targetPage += "?ID=" + selID;
			launchDialog(dialogTitle, targetPage, editDialog);
			//launchDialog(dialogTitle, targetPage, editDialog1);
		}	
		catch(e){}	
	}
	
	function launchDialog(pageTitle, targetPage, dialogArg){

		targetPage = "Dialog.asp?titleDoc=" + pageTitle + "&dialogDoc=" + targetPage;
		if (reportMode=="-1"){						
			reportCount ++;
			ActionWin = window.open(targetPage, 'report' + reportCount, 'height=600,width=700,status=no,scrollbars=yes,toolbar=no,menubar=no,location=no,resizable=yes');
			//alert(ActionWin.opener);
			ActionWin.opener = window.self;
		}
		else{
			getDialogLayout();	
			window.status = pageTitle;			
			ActionWin = window.showModalDialog(targetPage, dialogArg, "dialogWidth:" + dWidth + ";dialogHeight:" + dHeight + ";status:0;dialogHide:false;help:no;scroll:yes;resizable:no;edge:sunken;unadorned:yes");	
			//ActionWin = window.open(targetPage, "_default", "dialogWidth:" + dWidth + ";dialogHeight:" + dHeight + ";status:0;dialogHide:false;help:no;scroll:yes;resizable:no;edge:sunken;unadorned:yes");	
			//alert(ActionWin.opener);
			ActionWin.opener = window.self;
			//ActionWin = window.open(targetPage, 'test', 'height=600,width=700,status=no,scrollbars=yes,toolbar=no,menubar=yes,location=no,resizable=yes');
		}	
	}
	
	function addDialog() {
		var selectedID;
		var opener;
	}
	
	function DoAdd(){
		try{
			var selID = window.parent.frames["maininfo"].document.all.item("frmMain").elements("ID").value
			if (selID=="") {
				selID = 0;
			}
			
			var args = "";

			if(window.parent.frames["maininfo"].document.all.item("frmMain").elements("args") != null)
			{
				args = window.parent.frames["maininfo"].document.all.item("frmMain").elements("args").value;
			}
			
			var dialogTitle = getDialogTitle(); 
			if (reportMode=="-1") dialogTitle = "View " + dialogTitle;
			
			else{
				if (dialogTitle!=="") dialogTitle = "Add " + dialogTitle;				
			}	
			var targetPage = window.parent.frames["maininfo"].document.all.item("frmMain").elements("AddPage").value;	
			targetPage += "?ID=" + selID;
			targetPage += "&args=" + args;					
			addDialog.selectedID = selID;
			
			addDialog.opener = window.parent.frames["maininfo"];
			
			launchDialog(dialogTitle, targetPage, addDialog);
		}	
		catch(e){}	
	}	
	
	function getDialogTitle(){
		try{
			if (reportMode=="-1") return 'Report';
			else return (window.parent.frames["header"].document.all.item("DataDescription").innerHTML);
		}
		catch(e){return ('');}
	}
	
	function getDialogLayout(){
		var defWidth = "42em";
		var defHeight = "30em";
		try{
			var defLayout = window.parent.frames["maininfo"].document.all.item("dialogLayout").value;	
			var thisHeight = defLayout.substr(defLayout.indexOf("height:"), defLayout.indexOf(";"))
			thisHeight = thisHeight.replace("height:", "");
			var thisWidth = defLayout.substr(defLayout.indexOf("width:"), defLayout.length)
			thisWidth = thisWidth.replace("width:", "");
			defWidth = thisWidth;
			defHeight = thisHeight;
		}
		catch(e){}
		
		dWidth = defWidth;
		dHeight = defHeight;
	
	}
	
	function doRefreshOpener(){
		if (editDocOpener!==null){
			editDocOpener.document.forms["frmMain"].elements("action").value = "";
			editDocOpener.document.forms["frmMain"].target = "_self";
			editDocOpener.document.forms["frmMain"].submit();
			editDocOpener = null;
		}	
	}
	
	function relocateDocOpener(pageTo){
		if (editDocOpener!==null){
			editDocOpener.location.replace(pageTo);
			editDocOpener = null;
		}	
	}
	
	function closeDocOpener(){
		if (editDocOpener!==null){
			editDocOpener.self.close();
			editDocOpener = null;
		}	
	}
	
	function DoDelete()
	{
		try{
		
			var selID = window.parent.frames["maininfo"].document.all.item("frmMain").elements("ID").value
			
			if (selID!="")
			{
				//The message contained in the input will be overwritten if deletion proceeds. It is however restored on source page reload.
				var delInput = window.parent.frames["maininfo"].document.all.item("frmMain").elements("delAction");
				
				if(ConfirmDelete(delInput.value))
				{
					if (window.parent.frames["maininfo"].document.all.item("frmMain").elements("ActionPage").value=='SendClientReports.asp')
					{
					window.parent.frames["maininfo"].location.href = '../Data/DeleteClientReports.asp?del=1&ID='+selID;
					}
					else
					{
					ActionWin = window.frames["deleteFrame"];
					ActionWin.opener = window.parent.frames["maininfo"];
					var targetPage = window.parent.frames["maininfo"].document.all.item("frmMain").elements("DeletePage").value;
					window.parent.frames["maininfo"].document.all.item("frmMain").action = targetPage;
					window.parent.frames["maininfo"].document.all.item("frmMain").target = ActionWin.name;
					
					delInput.value = "Execute"
					window.parent.frames["maininfo"].document.all.item("frmMain").submit();
					}
				}
			}
			else alert ("No item selected for deletion");
		}			
		catch(e){}	
	}
	
	function cancelOperation(dObj){
		try{
			switch(dObj.id){
				case "cancelSort":
					window.parent.frames["maininfo"].document.all.item("SelectedSortArgs").value = "";	
					window.parent.frames["maininfo"].document.all.item("frmMain").action = window.parent.frames["maininfo"].document.all.item("ActionPage").value;
					window.parent.frames["maininfo"].document.all.item("frmMain").submit();
					document.all.item("SortTD").className = "footerHighlightnav";
					break;
				case "cancelFilter":
					window.parent.frames["maininfo"].document.all.item("SelectedFilterArgs").value = "";	
					window.parent.frames["maininfo"].document.all.item("frmMain").action = window.parent.frames["maininfo"].document.all.item("ActionPage").value;
					window.parent.frames["maininfo"].document.all.item("frmMain").submit();
					document.all.item("FilterTD").className = "footerHighlightnav";
					break;
				case "cancelSearch":
					window.parent.frames["maininfo"].document.all.item("SelectedSearchArgs").value = "";	
					window.parent.frames["maininfo"].document.all.item("frmMain").action = window.parent.frames["maininfo"].document.all.item("ActionPage").value;
					window.parent.frames["maininfo"].document.all.item("frmMain").submit();
					document.all.item("SearchTD").className = "footerHighlightnav";
					break;		
				default:
			}
			
			dObj.style.display="none";
		}
		catch(e){dObj.style.display="none";}
	
	}
	
	
</Script>

<script language="vbscript">
		Function ConfirmDelete(delMessage)
 			Dim ans
 			Dim msg
 			
 			msg = trim(delMessage)
 			if (msg = "")or (msg = "Execute") then
 				delMessage = "Delete the selected item?"
 			end if
 			ans = MsgBox(delMessage, vbYesNo + vbExclamation, "Delete")
 			if ans = vbNo Then
 				ConfirmDelete = false
 			else
 				ConfirmDelete = true
 			end if
		End function
</script>
</html>
