<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Group Access</title>
  <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
</head>

<!--#include file="../libroutines.asp"-->

<%
	
   Dim action
   Dim conn 
   Dim sqlStr
   Dim rs
	
	action = ucase(Request.Form("action"))
	ID = Request("ID")
	
	If ID = "" Then%>
		<Script Language="JavaScript">
			ShowMessage("No record specified for editing")
		</Script>
		<%WriteDialogRefuseOpenScript
		Response.End
	End If
	
	Set conn = GetActiveConnection("KBroker")
	  
	
	If action = "EXECUTE" Then	
		groupID = Request.Form("GroupID")
		
		SQL = "SELECT Menus.mnuCaption, Menus.IsReport, MenuGroups.* FROM Menus INNER JOIN MenuGroups " & _
			  " ON MenuGroups.MenuID = Menus.MenuID WHERE  Menus.MainMenuID = " & _
			  " (SELECT MenuID FROM Menus WHERE MenuID = (SELECT MenuID FROM MenuGroups WHERE [ID] = " & ID & ")) " & _
			  " AND Menus.isMainMenu <> 1 AND MenuGroups.GroupID = " & groupID & " ORDER BY Menus.mnuCaption"
		
		Set childRs = Conn.Execute(SQL)
		
		If Not (childRs.EOF Or childRs.BOF) Then
			Do Until childRs.EOF
				mnuID = childRs.Fields("menuID").Value	
				
				If childRs.Fields("IsReport").Value = "1" Then
					If Request.Form("CanAdd" & mnuID) = "1" Then	
						canAdd = 1
					Else
						canAdd = 0
					End If	
					
					updateSQL = "UPDATE MenuGroups SET CanAdd = " & canAdd & " WHERE [ID] = " & _
								" (SELECT [ID] FROM MenuGroups WHERE MenuID = " & mnuID & " AND GroupID = " & groupID & ")"
				Else 				
					canAdd = 0
					canDelete = 0
					canEdit = 0
					canSort = 0 
					canFilter = 0
					canSearch = 0  						
  					If Request.Form("CanAdd" & mnuID) = "1" Then	canAdd = 1
  					If Request.Form("CanDelete" & mnuID) = "1" Then canDelete = 1
  					If Request.Form("CanEdit" & mnuID) = "1" Then canEdit = 1
  					If Request.Form("CanSort" & mnuID) = "1" Then canSort = 1
  					If Request.Form("CanFilter" & mnuID) = "1" Then canFilter = 1
  					If Request.Form("CanSearch" & mnuID) = "1" Then canSearch = 1 				
  					
  					updateSQL = "UPDATE MenuGroups SET CanAdd = " & canAdd & ", CanDelete = " & canDelete & "," & _
  								" CanEdit = " & canEdit & " WHERE [ID] = " &  _
  								" (SELECT [ID] FROM MenuGroups WHERE MenuID = " & mnuID & " AND GroupID = " & groupID & ")"
				End If  		
				
				Conn.Execute updateSQL
				
				SQL = "SELECT Menus.mnuCaption, Menus.IsReport, MenuGroups.* FROM Menus INNER JOIN MenuGroups " & _
								" ON MenuGroups.MenuID = Menus.MenuID WHERE Menus.isMainMenu <> 1 AND " & _
								" Menus.mainmenuID = " & mnuID & " AND MenuGroups.GroupID = " & groupID & " ORDER BY mnuCaption"

				Set innerChildRs = Conn.Execute(SQL) 
							
							
				If Not (innerChildRs.EOF Or innerChildRs.BOF) Then 
								
					Do Until innerChildRs.EOF									
							mnuID = childRs.Fields("menuID").Value	
				
							If childRs.Fields("IsReport") = "1" Then
								If Request.Form("CanAdd" & mnuID).Value = "1" Then	
									canAdd = 1
								Else
									canAdd = 0
								End If	
								
								updateSQL = "UPDATE MenuGroups SET CanAdd = " & canAdd & " WHERE [ID] = " & _
											" (SELECT [ID] FROM MenuGroups WHERE MenuID = " & mnuID & " AND GroupID = " & groupID & ")"
							Else 				
								canAdd = 0
								canDelete = 0
								canEdit = 0
								canSort = 0 
								canFilter = 0
								canSearch = 0  						
  								If Request.Form("CanAdd" & mnuID) = "1" Then	canAdd = 1
  								If Request.Form("CanDelete" & mnuID) = "1" Then canDelete = 1
  								If Request.Form("CanEdit" & mnuID) = "1" Then canEdit = 1
  								If Request.Form("CanSort" & mnuID) = "1" Then canSort = 1
  								If Request.Form("CanFilter" & mnuID) = "1" Then canFilter = 1
  								If Request.Form("CanSearch" & mnuID) = "1" Then canSearch = 1 				
  								
  								updateSQL = "UPDATE MenuGroups SET CanAdd = " & canAdd & ", CanDelete = " & canDelete & "," & _
  											" CanEdit = " & canEdit & " WHERE [ID] = " & _
  											" (SELECT [ID] FROM MenuGroups WHERE MenuID = " & mnuID & " AND GroupID = " & groupID & ")"
							End If  		
				
							Conn.Execute updateSQL
				
				
						innerChildRs.MoveNext
					Loop	
				End If						
				
				
				
				childRs.MoveNext
			Loop
		End If		
		
		Set conn = Nothing
        WriteFraEnabledDialogCloseScript
        Response.End
		
   	End If  	
   	
   	
%>

<body Class="Dialog" OnLoad="VBScript: DoResizeWin" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">

<SCRIPT language=JavaScript>
		
	var cnt;
	var objSpanCollection;
	var menuHeightCollection = new Array();
	var objMenu;
	var lastObjMenu;
	var Ccounter = 0;
	var Mcounter = 1;
	var firstNode = 0;
	var lastNode = 0;
	var emptyCat = new Array(); var ec = 1; //track empty categories
	var currCat = 1, newCat;
	var menuDisplaced = new Boolean();
	var currMenuItem;
	
	function dopreload()
	{
		var the_images = new Array('../images/ftv2blank.gif','../images/ftv2lastnode.gif','../images/ftv2mlastnode.gif','../images/ftv2mnode.gif','../images/ftv2node.gif','../images/ftv2plastnode.gif','../images/ftv2pnode.gif','../images/ftv2vertline.gif','../images/ftv2pfirstnode.gif','../images/ftv2mfirstnode.gif');
		preloadImages(the_images);
	}

	function preloadImages(the_images_array) 
	{
		for(loop = 0; loop < the_images_array.length; loop++)
		{
	   		var an_image = new Image();
			an_image.src = the_images_array[loop];
		}
	}

	function openTree(caption)
	{
		
		document.write('<div style="margin-left:0; display:; ">');		
		document.write('<table cellpadding=0 cellspacing=0><tr>');
		document.write('<td valign=top><img src="../images/ftv2mfirstnode.gif" ID=MI-1 onClick="expandCollapseClick();">');
		document.write('</td>');
		document.write('<td valign=top><font size="2" color="brown" face="Arial" ID=MT-1><b>');
		document.write(caption);
		document.write('</a></b></font></td>');
		document.writeln('</tr></table>');
		
	}

	function closeTree()
	{
		document.writeln ('</DIV>\n');
	}

	// This function handles to onClick event for expanding/closing
	// nodes of the tree.
	function expandCollapseClick ()
	{
		var parentID;
		var child;
		var parentImage;
		var otherImage;
		

		parentID = window.event.srcElement.id;
		var imageNumber = parentID.substr(3);

		// used to detect which category is being clicked.
		if(imageNumber != '1')
			if(currCat != '1' && imageNumber != currCat){
				expandMenu("MI-" + currCat);
				currCat = 1;
		}

		if(parentID.charAt (0) == 'M' && !emptyCat[parentID.substr(3)])
		{
			child = document.all ('C-' + imageNumber);
			parentImage = document.all ('MI-' + imageNumber);
			if (child.style.display == 'none') // hidden
			{
				//otherImage = document.all ('FI-' + imageNumber);
				//otherImage.src = '../images/foldere2.gif';	
				child.style.display = '';

				// used only if tree root is not clicked.
				if(imageNumber != '1') currCat = imageNumber;

				/* use last minus sign when appropiate - determine which type of node */
				
				switch(parentID) {
					case 'MI-1' :
						parentImage.src = '../images/ftv2mfirstnode.gif';
						break;
					case ('MI-' + Mcounter) :
						parentImage.src = '../images/ftv2mlastnode.gif';			
						break;
					default : 
						if (parentImage.src.indexOf("ftv2plastnode.gif") != -1) parentImage.src = '../images/ftv2mlastnode.gif';
						else parentImage.src = '../images/ftv2mnode.gif';
						break;
				}
					
			}
			else
			{
				//otherImage = document.all ('FI-' + imageNumber);
				//otherImage.src = '../images/folderc2.gif';

				// used only if tree root is not clicked.
				if(imageNumber != '1') currCat = '1';
			
				child.style.display = 'none';
				switch(parentID) {
					case 'MI-1' :
						parentImage.src = '../images/ftv2pfirstnode.gif';
						break;
					case ('MI-' + Mcounter) :
						parentImage.src = '../images/ftv2plastnode.gif';			
						break;
					default :
						if (parentImage.src.indexOf("ftv2mlastnode.gif") != -1) parentImage.src = '../images/ftv2plastnode.gif';
						else parentImage.src = '../images/ftv2pnode.gif';
						break;
				}
			}
		}
	}

	// This function gets called from a script tag at the bottom of
	// your html to expand the menu upon page opening.
	function expandMenu(cat)
	{
		var parentID;
		var child;
		var parentImage;
		var otherImage;

		// Get First level desired to expand. You want to specify
		// Top level (MI-1) when calling this function.
		parentID = cat.substr(0);

		if (parentID.charAt(0) == 'M' && !emptyCat[parentID.substr(3)])
		{
		
			child = document.all ('C' + parentID.substr (2));
			parentImage = document.all ('MI' + parentID.substr (2));			
			
			if (child.style.display == 'none') // hidden
			{
				//otherImage = document.all ('FI' + parentID.substr (2));
				//otherImage.src = '../images/foldere2.gif';
				child.style.display = '';
				/* use last minus sign when appropiate - determine which type of node */			
				switch(parentID) {
					case 'MI-1' :
						parentImage.src = '../images/ftv2mfirstnode.gif';
						break;
					case ('MI-' + Mcounter) :
						parentImage.src = '../images/ftv2mlastnode.gif';			
						break;
					default : 
						if (parentImage.src.indexOf("ftv2plastnode.gif") != -1) parentImage.src = '../images/ftv2mlastnode.gif';
						else parentImage.src = '../images/ftv2mnode.gif';
						break;
				}
			}
			else
			{
				//otherImage = document.all ('FI' + parentID.substr (2));
				//otherImage.src = '../images/folderc2.gif';
				child.style.display = 'none';
				switch(parentID) {
					case 'MI-1' :
						parentImage.src = '../images/ftv2pfirstnode.gif';
						break;
					case ('MI-' + Mcounter) :
						parentImage.src = '../images/ftv2plastnode.gif';			
						break;
					default : 
						if (parentImage.src.indexOf("ftv2mlastnode.gif") != -1) parentImage.src = '../images/ftv2plastnode.gif';
						else parentImage.src = '../images/ftv2pnode.gif';
						break;
				}
			}
		}
	}
	
	function makeCheckBoxes(nodeID, canAdd, canDelete, canEdit, canSort, canFilter, canSearch){
		var checkStatus = 0;
		var categoryToDisplay = new String();
		var returnStr = new String();
		var isDisabled = new String ();
		var idHead = new String();
		var extraTDText = new String();
		
		for (loop=1; loop < 7; loop++){
			switch(loop){
				case 1:
				    categoryToDisplay = "New";
				    checkStatus = canAdd;
				    idHead = "canAdd";
				    if (checkStatus==null) extraTDText= ' STYLE="display: none" ';
					 break;
				case 2:
					categoryToDisplay = "Delete";
					checkStatus = canDelete;
					idHead = "canDelete";
					if (checkStatus==null) extraTDText = ' STYLE="display: none" ';
					break;
				case 3:
					categoryToDisplay = "Edit";
					checkStatus = canEdit;
					idHead = "canEdit";
					if (checkStatus==null) extraTDText = ' STYLE="display: none" ';
					break;
				case 4:
					isDisabled = "disabled";
					categoryToDisplay = "Sort";
					checkStatus = canSort;					
					idHead = "canSort";
					break;
				case 5:
					isDisabled = "disabled";
					categoryToDisplay = "Filter";
					checkStatus = canFilter;
					idHead = "canFilter";
					break;
				case 6:
					isDisabled = "disabled";
					categoryToDisplay = "Search";
					checkStatus = canSearch;
					idHead = "canSearch";
					break;						
			}
			
			
			if (categoryToDisplay !== ""){
				var thisID = nodeID + ' ' + categoryToDisplay;
				var thisStatus = "";
				if (checkStatus==1) thisStatus=" checked ";
				returnStr += '<td Class="tdRule"' + extraTDText + '><input type=checkbox ' + thisStatus + ' ' + isDisabled + ' name="' + idHead + nodeID + '" Value="1" Class="BorderLess" ID="' + thisID + '"><label for="' + thisID + '" style="cursor: hand"><font face=verdana size=1 color=brown>' + categoryToDisplay + '</font></label></td>'; 
			}
			isDisabled = ""
			categoryToDisplay = "";
			checkStatus = 0;
			extraTDText = "";
		}
		
		return returnStr;
	
	}
	
	function makeReportCheckBox(nodeID, canAdd)
	{
		var thisStatus;
		var thisID = nodeID + ' ' + name;
		var idHead = 'canAdd';
		if (canAdd==1) thisStatus = "checked";
		else thisStatus = "";			
		return ('<td Class="tdRule"><input type=checkbox ' + thisStatus + ' name="' + idHead + nodeID + '" Value="1" Class="BorderLess" ID="' + thisID + '"><label for="' + thisID + '" style="cursor: hand"><font face=verdana size=1 color=brown>Access to report</font></label></td>');
	}


	// This function creates the end nodes under any given category.
	function makeNode (isReport, name, nodeID, canAdd, canDelete, canEdit, canSort, canFilter, canSearch)
	{
		
		++lastNode; // increment to track last node of node division.
		var myID = 'MN-' + lastNode;
		document.writeln ('<table CELLPADDING=0 CELLSPACING=0 STYLE="margin-left: 1;margin-top: 0"><tr><td valign=top><img ID=img1-' + lastNode + ' src="../images/ftv2vertline.gif"><img ID=img2-' + lastNode + ' src="../images/ftv2node.gif"></td>'); 		
		var categoryToDisplay = name;		
		document.writeln('<td Class="tdRule"  STYLE="cursor: hand" OnMouseOver="window.parent.status=\'' + name + '\'" OnMouseOut="window.parent.status=\'\'" valign=middle>');
		document.writeln('<font face=Arial size="1" ID="' + nodeID + '"><font face="Arial" STYLE="font-size: 8pt;text-decoration: none;"');
		
		if (isReport==1) {
			document.writeln(' color="blue">' + categoryToDisplay + '</font></font></td>');
			document.writeln(makeReportCheckBox(nodeID, canAdd)); 
		}	
		else {
			document.writeln(' color="black">' + categoryToDisplay + '</font></font></td>');
			document.writeln(makeCheckBoxes(nodeID, canAdd, canDelete, canEdit, canSort, canFilter, canSearch));
		}	
		document.writeln ('</tr></table>');
	}
	

	// This function creates the category.
	function makeCategory (isReport, name, nodeID, canAdd, canDelete, canEdit, canSort, canFilter, canSearch, mnuType)
	{
	
		Mcounter++;
		var myID = 'MT-' + Mcounter ;
		var myCatID = 'MK-' + Mcounter ;
		document.writeln ('<table CELLPADING=0 STYLE="margin-left: 2; position: relative" CELLSPACING=0><tr><td valign=top><img src="../images/ftv2pnode.gif" ID=MI-' + Mcounter + '  onClick=expandCollapseClick()></td><td valign=top STYLE="cursor: hand" OnMouseOver="window.parent.status=\'' + name + '\'" OnMouseOut="window.parent.status=\'\'"><font face="Arial" style="cursor: hand;text-decoration: none;font-weight: bold;font-size: 7.9pt" ');
		
		if (isReport==1) {
			document.writeln(' color="blue" ID="' + myCatID + '" onClick="JavaScript: expandCollapseClick();">' + name +  '</font></td>');
			document.writeln(makeReportCheckBox(nodeID, canAdd)); 
		}	 
		else {
			document.writeln(' color="black" ID="' + myCatID + '" onClick="JavaScript: expandCollapseClick();">' + name +  '</font></td>');
			if (mnuType=="0") 	document.writeln(makeCheckBoxes(nodeID, canAdd, canDelete, canEdit, canSort, canFilter, canSearch));
			if (mnuType=="1") 	document.writeln(makeCheckBoxes(nodeID, null, null, canEdit));
		}	
		document.write ('</tr></table>');		
	}
	   
	// This function creates the opening div tag for hiding and
	// showing with the expandCollapseClick function.
	function openDiv ()
	{
		emptyCat[ec] = false;
		Ccounter++;		
		document.writeln ('<DIV ID=C-' + Ccounter + ' STYLE="margin-left: 2;">');
	}
	
	function openNodeDiv ()
	{
		firstNode = lastNode + 1; // remember first node of node division.
		++Ccounter;			
		document.writeln ('<DIV ID=C-' + Ccounter + ' STYLE="margin-left: 2; display:None;">');
		
	}

	// This function closes the div tag.
	function closeDiv ()
	{
		var image;

		//document.writeln ('</DIV>');

		// Choose correct image for last catagory
		image = document.all('MI-' + Mcounter);

		// if last category is empty no plus sign else plus sign
		if(emptyCat[ec])
			image.src = '../images/ftv2lastnode.gif';
		else
			image.src = '../images/ftv2plastnode.gif';

		// blank last category's node's first image(vertical line).
		for(i = firstNode; i <= lastNode; ++i)
		{
			image = document.all('img1-' + i);
			//Added width and height for proper alignment - should be OK  
			image.width = '24';
			image.height = '0';
			image.src = '../images/ftv2blank.gif';
		}		
		
	}
	function closeNodeDiv ()
	{
		var image;

		document.writeln('</DIV>');		
		
		// if category empty(no nodes) then no plus sign else set last node image
		if(firstNode == lastNode + 1){
			emptyCat[++ec] = true;
			image = document.all('MI-' + ec);
			image.src = '../images/ftv2node.gif';
		}
		else{
			emptyCat[++ec] = false;
			image = document.all('img2-' + lastNode);
			image.src = '../images/ftv2lastnode.gif';
		}
			
	}
	
</SCRIPT>

<DIV id=MainMenu>
<form name = 'frmEditGroupAccess' method = 'post' action = 'EditGroupAccess.asp'>

<%
		Set groupRs = Conn.Execute ("SELECT * FROM MenuGroups WHERE [ID] = " & ID)
		
		If Not (groupRs.EOF Or groupRs.BOF) Then
			mnuID = groupRs.Fields("MenuID").Value
			groupID = groupRs.Fields("GroupID").Value
			
			Set capRs = Conn.Execute("SELECT mnuCaption FROM Menus WHERE MenuID = " & mnuID)
			If Not (capRs.EOF Or capRs.BOF) Then
				caption = capRs.Fields("mnuCaption").Value
			Else%>
				<Script Language="JavaScript">
					ShowMessage("The place-holders based on the access rights information cannot be retrieved.")
				</script>
				<%
				WriteDialogRefuseOpenScript
				Response.End	
			End If		
			
			
		Else%>
			<Script Language="JavaScript">
				ShowMessage("The access rights information cannot be retrieved.")
			</script>
			<%
			WriteDialogRefuseOpenScript
			Response.End	
		End If			
		 
	    SQL = "SELECT Menus.mnuCaption, Menus.IsReport, Menus.mnuType, MenuGroups.* FROM Menus INNER JOIN MenuGroups " & _
			  " ON MenuGroups.MenuID = Menus.MenuID WHERE  Menus.MainMenuID = " & _
			  " (SELECT MenuID FROM Menus WHERE MenuID = (SELECT MenuID FROM MenuGroups WHERE [ID] = " & ID & ")) " & _
			  " AND Menus.isMainMenu <> 1 AND MenuGroups.GroupID = " & groupID & " ORDER BY Menus.mnuCaption"
		
		Set childRs = Conn.Execute(SQL)
		containsSubCategories = False
		If Not (childRs.EOF Or childRs.BOF) Then %>
						<SCRIPT Language="JavaScript"> 
							dopreload();
							openTree("<%= Trim(caption) %>");
							openDiv();
							 
						</SCRIPT>
						<%
						 containsSubCategories = True
						 Do Until childRs.EOF							
							mnuCaption = Trim(childRs("mnuCaption"))								
							mnuID = childRs.Fields("menuID").Value		
							canAdd = 0
							canDelete = 0
							canEdit = 0
							canSort = 0 
							canFilter = 0
							canSearch = 0  		
											
  							If childRs.Fields("CanAdd").Value = "1" Then	canAdd = 1
  							If childRs.Fields("CanDelete").Value = "1" Then canDelete = 1
  							If childRs.Fields("CanEdit").Value = "1" Then canEdit = 1
  							If childRs.Fields("CanSort").Value = "1" Then canSort = 1
  							If childRs.Fields("CanFilter").Value = "1" Then canFilter = 1
  							If childRs.Fields("CanSearch").Value = "1" Then canSearch = 1
  							isReport = childRs.Fields("IsReport").Value			
  							mnuType = childRs.Fields("mnuType").Value			%>						
							
							<SCRIPT Language="JavaScript">
								makeCategory ('<%= isReport %>', '<%= Replace(Trim(mnuCaption), "'", "\'")  %>', '<%= mnuID %>', '<%= canAdd %>', '<%= canDelete %>', '<%= canEdit %>', '<%= canSort %>', '<%= canFilter %>', '<%= canSearch %>', '<%= mnuType %>');
								openNodeDiv ();								
							</SCRIPT>
							
							<%
							

							SQL = "SELECT Menus.mnuCaption, Menus.IsReport, MenuGroups.* FROM Menus INNER JOIN MenuGroups " & _
								" ON MenuGroups.MenuID = Menus.MenuID WHERE Menus.isMainMenu <> 1 AND " & _
								" Menus.mainmenuID = " & mnuID & " AND MenuGroups.GroupID = " & groupID & " ORDER BY mnuCaption"

							Set innerChildRs = Conn.Execute(SQL) 
							
							
							If Not (innerChildRs.EOF Or innerChildRs.BOF) Then 
								
								Do Until innerChildRs.EOF									
										canAdd = 0
										canDelete = 0
										canEdit = 0
										canSort = 0 
										canFilter = 0
										canSearch = 0  						
  										If innerChildRs.Fields("CanAdd").Value = "1" Then	canAdd = 1
  										If innerChildRs.Fields("CanDelete").Value = "1" Then canDelete = 1
  										If innerChildRs.Fields("CanEdit").Value = "1" Then canEdit = 1
  										If innerChildRs.Fields("CanSort").Value = "1" Then canSort = 1
  										If innerChildRs.Fields("CanFilter").Value = "1" Then canFilter = 1
  										If innerChildRs.Fields("CanSearch").Value = "1" Then canSearch = 1
  										isReport = childRs.Fields("IsReport").Value	
										nodeID = innerChildRs.Fields("menuID").Value%>
										
										<SCRIPT Language="JavaScript">									
											makeNode ('<%= isReport %>','<%= Replace(Trim(innerChildRs.Fields("mnucaption").value), "'", "\'") %>', '<%= nodeID %>', '<%= canAdd %>', '<%= canDelete %>', '<%= canEdit %>', '<%= canSort %>', '<%= canFilter %>', '<%= canSearch %>');
										</SCRIPT>																				
										
									<%
									innerChildRs.MoveNext
								Loop 																
							End If
							
							Set innerChildRs = Nothing %>
							<SCRIPT Language="JavaScript">
									closeNodeDiv ();
							</SCRIPT>
								
							<%														
							childRs.MoveNext
						Loop%>											
						<SCRIPT LANGUAGE="JAVASCRIPT">
							closeDiv();
							closeTree();
						</SCRIPT>
					<%
			Else%>
					<Script Language="JavaScript">
						ShowMessage("There are currently no place-holders for the selected group access right")
					</Script>	
				 <%WriteDialogRefuseOpenScript
			End If			
			Set childRs = Nothing%>

		<table align=right cellpadding=0 cellspacing=0>
			<tr>
			 <td width="100%" align=right>
				<BR>
			
				<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
				&nbsp;&nbsp;
				<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">				
				<input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
				<input type = 'hidden' name ='ID' id = 'ID' value="<%= ID %>">
				<input type = 'hidden' name ='GroupID' id = 'GroupID' value="<%= GroupID %>">
		      </td>
				</tr>

		</table>
			

	
	</form>
</Div>


<Script Language="VBScript">
	
	
	Function DoResizeWin	
		On Error Resume Next
		
		Dim nHeight, nWidth
		Dim defMaxHeight, defMaxWidth
		
		defMaxHeight = (screen.availHeight) - 100
		defMaxWidth = (screen.availWidth) - 100
		
		nHeight = (document.all.item("mainTable").clientHeight) + 50
		nWidth = (document.all.item("mainTable").clientWidth) + 50
		
		If nHeight > defMaxHeight Then
			nHeight = defMaxHeight
		End If
		
		If nWidth > defMaxWidth Then	
			nWidth = defMaxWidth
		End If		
				
		window.parent.dialogHeight = nHeight & "px"
		window.parent.dialogWidth = nWidth & "px"
		
		
	End Function
</Script>
</body>

