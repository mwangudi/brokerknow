<%

	
		const WorkArea = "maininfo"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"><html>

<head><TITLE>BrokerKnow Menus</TITLE>
<META http-equiv=Content-Type content="text/html">
<LINK href="STYLE/default.css" type=TEXT/CSS rel=STYLESHEET> 
<LINK href="STYLE/webparts.css" type=TEXT/CSS rel=STYLESHEET>
<SCRIPT language=Javascript src="scripts/common.js"></SCRIPT>
<base target=<%=WorkArea%>>
<STYLE type=text/css>
	body {
		background-color: #FFFFFF;
		font-size: 8pt;
		font-family: Verdana;
	}
		
	.Menu_Head {
		filter:alpha(opacity=80);
		background-color: #FFFFFF;
		color: white;
		cursor: pointer;
		width: 175px;
		height: 20px;
	}

	.Menu_Items {
		filter:alpha(opacity=100);		
		padding: 10;
		display: block;
		width: 175px;		
		background-color: white;
	}
	
	.CaptionImage {
		filter:alpha(opacity=80);
		background-color: #FFFFFF;
		color: white;
		cursor: pointer;
	}

	.Description {
		filter:alpha(opacity=90);
		background-color: #6699cc;
		display: none;
		width: 190;
		height: 90;
		position: absolute;
		border: 1 solid #006699;
	}
	.DescTitle {
		background-color: #006699;
		color: white;
		font-weight: bold;
	}
</STYLE>

<META content="MSHTML 6.00.2600.0" name=GENERATOR>

</head>
<BODY leftMargin=0 topMargin=0 onload="JavaScript: InitializeMenu(); resizeObjects()" marginheight="0" marginwidth="0">
<!--#include file="libroutines.asp"-->
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
		var the_images = new Array('images/ftv2blank.gif','images/ftv2lastnode.gif','images/ftv2mlastnode.gif','images/ftv2mnode.gif','images/ftv2node.gif','images/ftv2plastnode.gif','images/ftv2pnode.gif','images/ftv2vertline.gif','images/ftv2pfirstnode.gif','images/ftv2mfirstnode.gif');
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
		document.write('<font face="Tahoma,Arial,Helvetica,Sans Serif" size="2">');
		document.write('<div style="margin-left:0; display:; ">');		
		document.write('<table cellpadding=0 cellspacing=0><tr>');
		document.write('<td valign=top><img src="images/ftv2mfirstnode.gif" ID=MI-1 onClick="expandCollapseClick();">');
		document.write('<img ID=FI-1 src="images/foldere2.gif"></td>');
		document.write('<td valign=top><font size="-1" ID=MT-1><b>');
		document.write(caption);
		document.write('</a></b></font></td>');
		document.writeln('</tr></table>');
		
	}

	function closeTree()
	{
		document.writeln ('</DIV>\n</FONT>');
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
				//otherImage.src = 'images/foldere2.gif';	
				child.style.display = '';

				// used only if tree root is not clicked.
				if(imageNumber != '1') currCat = imageNumber;

				/* use last minus sign when appropiate - determine which type of node */
				
				switch(parentID) {
					case 'MI-1' :
						parentImage.src = 'images/ftv2mfirstnode.gif';
						break;
					case ('MI-' + Mcounter) :
						parentImage.src = 'images/ftv2mlastnode.gif';			
						break;
					default : 
						if (parentImage.src.indexOf("ftv2plastnode.gif") != -1) parentImage.src = 'images/ftv2mlastnode.gif';
						else parentImage.src = 'images/ftv2mnode.gif';
						break;
				}
					
			}
			else
			{
				//otherImage = document.all ('FI-' + imageNumber);
				//otherImage.src = 'images/folderc2.gif';

				// used only if tree root is not clicked.
				if(imageNumber != '1') currCat = '1';
			
				child.style.display = 'none';
				switch(parentID) {
					case 'MI-1' :
						parentImage.src = 'images/ftv2pfirstnode.gif';
						break;
					case ('MI-' + Mcounter) :
						parentImage.src = 'images/ftv2plastnode.gif';			
						break;
					default :
						if (parentImage.src.indexOf("ftv2mlastnode.gif") != -1) parentImage.src = 'images/ftv2plastnode.gif';
						else parentImage.src = 'images/ftv2pnode.gif';
						break;
				}
			}
		}
		
		resizeObjects();
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
				//otherImage.src = 'images/foldere2.gif';
				child.style.display = '';
				/* use last minus sign when appropiate - determine which type of node */			
				switch(parentID) {
					case 'MI-1' :
						parentImage.src = 'images/ftv2mfirstnode.gif';
						break;
					case ('MI-' + Mcounter) :
						parentImage.src = 'images/ftv2mlastnode.gif';			
						break;
					default : 
						if (parentImage.src.indexOf("ftv2plastnode.gif") != -1) parentImage.src = 'images/ftv2mlastnode.gif';
						else parentImage.src = 'images/ftv2mnode.gif';
						break;
				}
			}
			else
			{
				//otherImage = document.all ('FI' + parentID.substr (2));
				//otherImage.src = 'images/folderc2.gif';
				child.style.display = 'none';
				switch(parentID) {
					case 'MI-1' :
						parentImage.src = 'images/ftv2pfirstnode.gif';
						break;
					case ('MI-' + Mcounter) :
						parentImage.src = 'images/ftv2plastnode.gif';			
						break;
					default : 
						if (parentImage.src.indexOf("ftv2mlastnode.gif") != -1) parentImage.src = 'images/ftv2plastnode.gif';
						else parentImage.src = 'images/ftv2pnode.gif';
						break;
				}
			}
		}
	}


	// This function creates the end nodes under any given category.
	function makeNode (name, URLpath, target, nodeID, nodeDesc)
	{
		
		++lastNode; // increment to track last node of node division.
		var myID = 'MN-' + lastNode;
		document.writeln ('<table CELLPADDING=0 CELLSPACING=0 STYLE="margin-left: 1;margin-top: 0"><tr><td valign=top><img ID=img1-' + lastNode + ' src="images/ftv2vertline.gif"><img ID=img2-' + lastNode + ' src="images/ftv2node.gif"></td>'); 
		
		var categoryToDisplay, targetstr;

		if (name.length > 12)
		{
			categoryToDisplay = name.substring (0, 50);
		}
		else
		{
			categoryToDisplay = name;
		}
		document.writeln('<td width="100" STYLE="cursor: hand" OnMouseOver="window.parent.status=\'' + name + '\'" OnMouseOut="window.parent.status=\'\'" TITLE="Open &#145;' + name + '&#146;." valign=middle OnClick="JavaScript:  OpenLink(\'' + URLpath + '\', \'' + target + '\'); UpdateToolBar(\'' + nodeID + ' \'); UpdateHeader (\'' + nodeDesc.toUpperCase() + ' \');">');
		document.writeln('<font face=Arial size="1" ID="' + nodeID + '"><font face="Arial" STYLE="font-size: 8pt;text-decoration: none;" >' + categoryToDisplay + '</font></font>');
		document.writeln ('</td></tr></table>');
	}
	
	//This function opens the URL link
	function OpenLink(path, target){
		window.parent.frames["tabs"].document.all.item("ReportTitle").innerHTML = "";
		targetstr = target + ""; // Explicitly change target to string
		try{
			if (targetstr != "") // Test to see if target is empty, if not empty add the target attribute to href tag
			{
				window.parent.frames[target].location.replace(path);
			}
			else // else leave tag alone.
			{
				window.location.replace(path);
			}	
		}
		catch(e){}
	}

	// This function creates the category.
	function makeCategory (name, URLpath, target, theCatID, catDesc)
	{
	
		Mcounter++;
		var myID = 'MT-' + Mcounter ;
		var myCatID = 'MK-' + Mcounter ;
		if (URLpath!="#") document.writeln ('<table CELLPADING=0 STYLE="margin-left: 2; position: relative" CELLSPACING=0><tr><td valign=top><img src="images/ftv2pnode.gif" ID=MI-' + Mcounter + '  onClick=expandCollapseClick()></td><td valign=top STYLE="cursor: hand" OnMouseOver="window.parent.status=\'' + name + '\'" OnMouseOut="window.parent.status=\'\'"><font face="Arial" color="black" style="cursor: hand;text-decoration: none;font-weight: bold;font-size: 7.9pt" ID="' + myCatID + '" onClick="JavaScript: OpenLink(\'' + URLpath + '\', \'' + target + '\'); expandCollapseClick(); UpdateToolBar(\'' + theCatID + ' \'); UpdateHeader(\'' + catDesc + ' \')">' + name +  '</font></td></tr></table>');
		else document.writeln ('<table CELLPADING=0 STYLE="margin-left: 2; position: relative" CELLSPACING=0><tr><td valign=top><img src="images/ftv2pnode.gif" ID=MI-' + Mcounter + '  onClick=expandCollapseClick()></td><td valign=top OnMouseOver="window.parent.status=\'' + name + '\'" OnMouseOut="window.parent.status=\'\'"><font face="Arial" color="black" style="cursor: hand;text-decoration: none;font-weight: bold;font-size: 7.9pt" ID="' + myCatID + '" onClick="JavaScript: expandCollapseClick();">' + name +  '</font></td></tr></table>');
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
			image.src = 'images/ftv2lastnode.gif';
		else
			image.src = 'images/ftv2plastnode.gif';

		// blank last category's node's first image(vertical line).
		for(i = firstNode; i <= lastNode; ++i)
		{
			image = document.all('img1-' + i);
			//Added width and height for proper alignment - should be OK  
			image.width = '24';
			image.height = '0';
			image.src = 'images/ftv2blank.gif';
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
			image.src = 'images/ftv2node.gif';
		}
		else{
			emptyCat[++ec] = false;
			image = document.all('img2-' + lastNode);
			image.src = 'images/ftv2lastnode.gif';
		}
			
	}
	

	function InitializeMenu() 	{
		
		objSpanCollection =  document.body.all.MainMenu.getElementsByTagName("SPAN");		  		  
		for (var i = 0; i < objSpanCollection.length; i++)
		{

			var objSpan = objSpanCollection(i);
			menuHeightCollection[i] = objSpan.childNodes(1).clientHeight;
			var objLinkCollection = document.body.all.MainMenu.getElementsByTagName("A");
			for (var j = 0; j < objLinkCollection.length; j++)
			{
				var objA = objLinkCollection(j);
			}
			
			objSpan.childNodes(0).onclick = ControlMenu;
			objSpanCollection(i).childNodes(1).style.display = "none";
			objSpanCollection(i).childNodes(1).childNodes(0).style.display = "none";
		}
		
	}
	
	function ShowDescription()
	{	
		x = event.clientX + document.body.scrollLeft; /* get the mouse left position */
		y = event.clientY + document.body.scrollTop + 35; /* get the mouse top position  */
		this.parentNode.childNodes(2).style.display = "block";
		this.parentNode.childNodes(2).style.left = x;
		this.parentNode.childNodes(2).style.top = y;
	}
	
	function HideDescription()
	{
		this.parentNode.childNodes(2).style.display = "none";
	}
	
	function ControlMenu() 
	{
		try{
			if (lastObjMenu != this.parentNode.childNodes(1)){
				objMenu = lastObjMenu;
				HideMenuQuick();
				//update the class of the previously-clicked menu button
				objMenu.parentNode.childNodes(0).childNodes(0).childNodes(0).childNodes(0).className = "nav_Menu";
			}	
		}
		catch(e){}		
		cnt = 1;
		objMenu = this.parentNode.childNodes(1);				
		if (objMenu.style.display == "none") {				
			menuDisplaced = false;
			ShowMenu();
			}
		else {
			HideMenu();			
		}
		
	}
	
	function ShowMenu()
	{		
		var objList = objMenu.childNodes(0);				
		if (cnt < 10)
		{
			objMenu.style.display = "block";			
			objMenu.filters["alpha"].opacity = objMenu.filters["alpha"].opacity + 100;
			for (var i = 0; i < objSpanCollection.length; i++)
				if (objMenu.parentNode == objSpanCollection[i]){
					objMenu.style.height = objMenu.clientHeight + (menuHeightCollection[i]/10);
					currMenuItem = i;
					}
			cnt++;		
			setTimeout("ShowMenu()",1)
		}
		
		if (cnt >= 10) {
			objList.style.display = "block";  			
			lastObjMenu = objMenu;
			if (menuDisplaced==false){
				//resizeMenuDisplay(objMenu);
				resizeObjects();
				menuDisplaced = true;
			}
		}
		

	}
	
	function resizeMenuDisplay(theMenu){
		var resultHeight;
		resultHeight = document.body.clientHeight - ((objSpanCollection.length - currMenuItem) * theMenu.parentNode.childNodes(0).clientHeight) - theMenu.offsetTop;
		resultHeight = (resultHeight + theMenu.parentNode.childNodes(0).clientHeight) - document.all.item("BottomDiv").clientHeight;
		theMenu.style.height = resultHeight;
	}
	
	function HideMenu()
	{	
		var objList = objMenu.childNodes(0);
		if (cnt < 10)
		{
			objMenu.filters["alpha"].opacity = objMenu.filters["alpha"].opacity - 100;
			for (var i = 0; i < objSpanCollection.length; i++)
				if (objMenu.parentNode == objSpanCollection[i])
					objMenu.style.height = objMenu.clientHeight - (menuHeightCollection[i]/10);
			objList.style.display = "none";
			cnt++;
			setTimeout("HideMenu()",1)
		}
		if (cnt >= 10){
			objMenu.style.display = "none";
			resizeObjects();
		}	
	}	
	
	function HideMenuQuick()
	{	
		var objList = objMenu.childNodes(0);
		for (cnt = 1; cnt < 11; cnt++){
			if (cnt < 10)
				{
				for (var i = 0; i < objSpanCollection.length; i++)
					if (objMenu.parentNode == objSpanCollection[i])
						objMenu.style.height = objMenu.clientHeight - (menuHeightCollection[i]/10);
				objList.style.display = "none";
				}
			if (cnt >= 10)
				objMenu.style.display = "none";
		}		
	}	
	
	function UpdateToolBar(theMenu){
		window.parent.frames["footer"].location.replace("Footer.asp?mnuID=" + theMenu)		  
		
	}
	
	function UpdateReportsTab(reportID){
		if (reportID !== "" && reportID!=="undefined"){
				window.parent.frames["tabs"].document.all.item("CurrentMenuID").value = reportID;	
				window.parent.frames["tabs"].document.all.item("InputURL").value = "";
		}
		else window.parent.frames["tabs"].document.all.item("CurrentMenuID").value = "";	
		
		window.parent.frames["tabs"].selectCurrentTab(0);
	}
	
	function UpdateWorkArea(thePath, theTarget, theCatID, catDesc)
	{
		OpenLink(thePath,theTarget);
		UpdateToolBar(theCatID); 
		UpdateHeader(catDesc);
	}
	
	function UpdateHeader(theDesc){
		try{
			window.parent.frames["header"].document.all.item("DataDescription").innerHTML = theDesc.toUpperCase();
			if (theDesc!==""){
				window.parent.frames["tabs"].document.all.item("TabsTable").style.display = "";
				window.parent.frames["tabs"].document.all.item("InputURL").value = "";
			}
			
			window.parent.frames["tabs"].selectCurrentTab(0);
		}
		catch(e){}	
	}
	
	function resizeObjects(){
		try{
			HidePreferences();			
			if (lastObjMenu != 'undefined' && lastObjMenu != null){	
				if (lastObjMenu.style.display=="block") resizeMenuDisplay(lastObjMenu);
			}				
			var bottomObj = document.all.item("BottomDiv");	
			var v_h, b_h, bottomObjNewTop;
			v_h = parseInt(objSpanCollection[objSpanCollection.length - 1].offsetTop) + parseInt(objSpanCollection[objSpanCollection.length - 1].offsetHeight);
			b_h = document.body.clientHeight - bottomObj.clientHeight;
			if (b_h > v_h) bottomObjNewTop = b_h
			else bottomObjNewTop = v_h;
			bottomObj.style.top = bottomObjNewTop ;
		}
		catch(e){}	
	}
	function CloseApp(){	
		if (window.confirm("Are you sure you want to quit this application?")){
			window.parent.close();  
		}
	}
	
	function CenterWin(ActionWin, width, height){
		ActionWin.moveTo(((screen.availWidth - width) / 2), ((screen.availHeight - height) / 2) );
	}
	
	function LoadHelp(){
		window.showHelp('Help', '', 'helpWindow');
	}
	function CancelSelect(){
		window.event.cancelBubble = true;
	}
	
	function HidePreferences(){
		document.all.item("Preferences").style.display = "none";
	}
	

	function ShowPreferences(){
		var prefCallee = document.all.item("Preferences");
		var prefCaller = document.all.item("PrefRow");
		var prefCallerOwner = document.all.item("BottomDiv");
		
		if (prefCallee.style.display=="none"){
			prefCallee.style.top = (prefCallerOwner.style.pixelTop - prefCallee.clientHeight) + 2;
			prefCallee.style.width = prefCaller.clientWidth;
			prefCallee.style.left  = "56";
			prefCallee.style.display = "";
		}
		else {HidePreferences();}
		
		window.event.cancelBubble = true;		
	}
	
	function fullScreen(){
		var hdiff;
		window.parent.moveTo(-4,-4);
		hdiff=window.parent.screenTop;
		window.parent.moveTo(-6,-hdiff-7);
		window.parent.resizeTo(screen.width+13,screen.height+hdiff+33);
	}
	
	function normalScreen(){
		var hdiff;
		window.parent.moveTo(0,0);
		hdiff=window.parent.screenTop;
		window.parent.resizeTo(screen.width,screen.height);
	}
	
	window.onresize = resizeObjects;
	window.document.onscroll = resizeObjects;
	window.document.body.onselectstart = CancelSelect;
	window.document.onclick = HidePreferences;
	
	
</SCRIPT>
<BR>
<DIV id=MainMenu>
<%
		
		Dim Conn

		Dim RS
		Dim m_mnuArray(100)
		Dim mnuCaption
		Dim rsChild, rsSubChild 

		UserID = Session("UserID")

		Set Conn = GetActiveConnection("KBroker")		%>
 
	<SCRIPT Language="JavaScript"> 
		openDiv();								
	</SCRIPT>
	
 	<% 	
 	Dim SQL, containsSubCategories   
 	
  
    SQL = "SELECT * FROM MainMenuList WHERE IsReport <> 1 AND IsMainMenu = 1 AND EXISTS(SELECT     MenuGroups.ID " & _
			" FROM         UserGroups INNER JOIN " & _
			"                      MenuGroups ON UserGroups.GroupID = MenuGroups.groupID " & _
			"			WHERE     (UserGroups.UserID = " & userID & ") AND (MenuGroups.MenuID = MainMenuList.menuID))   ORDER BY mnuCaption"
	
    Set RS = Conn.Execute(SQL)
	
	If Not (Rs.EOF Or Rs.BOF) Then	%>
		<SCRIPT Language="JavaScript"> dopreload(); </SCRIPT>
		<%Do Until Rs.EOF%>		
		<SPAN>		
			
			<table cellpadding="0" cellspacing="1" class="Menu_Head">
				<tr>						
					<td class="nav_Menu" onMouseover="JavaScript: if (this.className!=='nav_clicked') this.className='nav_Menu_over'; window.status=this.innerText;" onMouseout="JavaScript: if (this.className=='nav_Menu_over') this.className='nav_Menu'; window.status='';" OnClick="JavaScript: if (this.className=='nav_clicked') this.className='nav_Menu'; else this.className='nav_clicked'; UpdateReportsTab('<%= Rs.Fields("MenuID").Value %>'); UpdateWorkArea('<%= Rs.Fields("DefaultChildAction").Value %>','<%=WorkArea%>','<%= Rs.Fields("DefaultChildID").Value %>','<%= Rs.Fields("DefaultChildDescription").Value %>')"><%= Trim(Rs("mnuCaption")) %></td>
				</tr>
			</table>
				
			<table border="0" cellspacing=0 style="BORDER-RIGHT: #cccccc 1px inset; BORDER-TOP: #cccccc 1px inset; BORDER-LEFT: #cccccc 1px inset; BORDER-BOTTOM: #cccccc 1px inset" class="Menu_Items" STYLE="display: none">
				<TR>
					<TD valign="absTop">
							
			
		<%
		 mnuID = Rs.Fields("menuID").Value
		 
		 SQL = "SELECT * FROM Menus WHERE IsReport <> 1 AND mainmenuID = " & mnuID & " AND isMainMenu <> 1 AND EXISTS(SELECT     MenuGroups.ID " & _
			" FROM         UserGroups INNER JOIN " & _
			"                      MenuGroups ON UserGroups.GroupID = MenuGroups.groupID " & _
			"			WHERE     (UserGroups.UserID = " & userID & ") AND (MenuGroups.MenuID = Menus.menuID))    ORDER BY mnuCaption"
		
		Set childRs = Conn.Execute(SQL)
		containsSubCategories = False
		If Not (childRs.EOF Or childRs.BOF) Then
						 containsSubCategories = True
						 Do Until childRs.EOF	
						
							mnuCaption = Trim(childRs("mnuCaption"))
							mnuDescription = childRs("mnuDescription")
							
							If IsNull(childRs("Image")) Then
								mnuImage = "images/linknote.gif"
							ElseIf childRs("Image") = "" Then
								mnuImage = "images/linknote.gif"
							Else
								mnuImage = childRs("Image")
							End If
							
							id = childRs.Fields("menuID").Value
							
							If Not (childRs("mnuaction") = "") Then 'Action is a html link
                        		Link = childRs("mnuaction")
                   			Else
                        		Link = "#"
							End If 
						
						
							%>						
							
							<SCRIPT Language="JavaScript">
								makeCategory ("<%= Replace(Trim(mnuCaption), "'", "\'") %>", "<%= Link %>", "<%=WorkArea%>", "<%= id %>", "<%= mnuDescription %>");
								openNodeDiv ();								
							</SCRIPT>
							
							<%
							mnuID = ID

							SQL = "SELECT * FROM Menus  WHERE IsReport <> 1 AND isMainMenu <> 1 AND mainmenuID = " & mnuID & " AND EXISTS(SELECT     MenuGroups.ID " & _
									" FROM         UserGroups INNER JOIN " & _
									"                      MenuGroups ON UserGroups.GroupID = MenuGroups.groupID " & _
									"			WHERE     (UserGroups.UserID = " & userID & ") AND (MenuGroups.MenuID = Menus.menuID))   ORDER BY mnuCaption"

							Set innerChildRs = Conn.Execute(SQL) 
							
							
							If Not (innerChildRs.EOF Or innerChildRs.BOF) Then %>							
								<%
								Do Until innerChildRs.EOF
									
										If Not (innerChildRs.Fields("mnuaction").Value  = "") Then
											Action = innerChildRs.Fields("mnuaction").Value 
										Else
											Action = "#"
										End If       
										
										nodeID = innerChildRs.Fields("menuID").Value
										nodeDesc = ucase(innerChildRs.Fields("mnuDescription").Value)%>
										<SCRIPT Language="JavaScript">									
											makeNode ('<%= Replace(Trim(innerChildRs.Fields("mnucaption").value), "'", "\'") %>', '<%= Action %>', 'maininfo', '<%= nodeID %>', '<%= nodeDesc %>');
										</SCRIPT>																				
									<%innerChildRs.MoveNext
								Loop 																
							End If
							
							Set innerChildRs = Nothing %>
							<SCRIPT Language="JavaScript">
									closeNodeDiv ();
							</SCRIPT>
								
							<%														
							childRs.MoveNext
						Loop%>
											
						
					<%
			End If			
			Set childRs = Nothing%>
				
				</TD>
						</TR>
					</table>		
						
		</SPAN>
		
			<%	If containsSubCategories Then%>
					<SCRIPT LANGUAGE="JAVASCRIPT">closeDiv()</SCRIPT>			
		<%		End If
			Rs.MoveNext
		Loop	%>
		
	<%		
	End If
	
	Set Rs = Nothing%>
	
	<Div STYLE="position:absolute; margin-left: 0;" ID="BottomDiv">
		<Table Class="Menu_Head">
			<TR valign=left>
			<TD  nowrap width="20px" class="nav" onMouseover="JavaScript: this.className='nav_over'; window.status='Close BrokerKnow'" onMouseout="JavaScript: this.className='nav'; window.status=''" OnSelectStart="JavaScript: window.event.cancelBubble = true;" OnClick="JavaScript: CloseApp()">
				&nbsp;Close&nbsp;
			</TD>
			<TD ID="PrefRow" nowrap width="20px" class="nav" onMouseover="JavaScript: this.className='nav_over'; window.status='Preferences'" onMouseout="JavaScript: this.className='nav'; window.status=''" OnSelectStart="JavaScript: window.event.cancelBubble = true;" OnClick="JavaScript: ShowPreferences()">
				&nbsp;Preferences&nbsp;
			</TD>
			<TD  nowrap width="20px" class="nav" onMouseover="JavaScript: this.className='nav_over'; window.status='Help'" onMouseout="JavaScript: this.className='nav'; window.status=''" OnSelectStart="JavaScript: window.event.cancelBubble = true;" OnClick="JavaScript: LoadHelp()">
				&nbsp;<b>?</b>&nbsp;
			</TD>
			</TR>
		</TABLE>
	</Div>

<DIV STYLE="position:absolute; left:0; display: none" ID="Preferences">
	<Table STYLE="BORDER-RIGHT: #cccccc 1px inset; Left: 0; BORDER-TOP: #cccccc 1px inset; BORDER-LEFT: #cccccc 1px inset; BORDER-BOTTOM: #cccccc 1px inset"  Class="cal-Textbox" border="0">
		<TR valign=left>
			<TD  nowrap OnMouseOver="JavaScript: window.status=this.innerText" OnMouseOut="JavaScript: window.status=''" STYLE="cursor: hand">
				<b>Theme</b>
			</TD>
		</TR>
		<TR valign=left>
			<TD  nowrap  OnMouseOver="JavaScript: window.status=this.innerText" OnMouseOut="JavaScript: window.status=''" STYLE="cursor: hand" OnClick="JavaScript: fullScreen();">
				<b>Full Screen</b>
			</TD>
		</TR>
		<TR valign=left>
			<TD  nowrap  OnMouseOver="JavaScript: window.status=this.innerText" OnMouseOut="JavaScript: window.status=''" STYLE="cursor: hand" OnClick="JavaScript: normalScreen();">
				<b>Normal Screen</b>
			</TD>
		</TR>
	</TABLE>
</DIV>
</BODY>
</html>



























































