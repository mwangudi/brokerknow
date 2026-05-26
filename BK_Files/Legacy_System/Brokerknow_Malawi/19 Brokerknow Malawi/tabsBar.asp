<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2//EN">

<html dir=ltr><head><title>BrokerKnow Navigational Tabs Bar</title>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=Windows-1252">
<SCRIPT language=Javascript src="scripts/common.js"></SCRIPT>
<LINK href="STYLE/default.css" type=TEXT/CSS rel=STYLESHEET> 
</head>

<BODY  leftMargin=0 topMargin=2 marginwidth="0" marginheight="0" >
<TABLE ID="TabsTable" valign="absTop" STYLE="display: none" cellpadding=0 cellspacing=0 width=100%> 
	<TR>
	<TD class="taMenu" valign=Top align=Left width=33%>
		<DIV Class="taMenu">
			<UL>
				<LI id="current"  name="0"><a href="#" OnClick="JavaScript: NavigateURL(0)" title="Input/Modify"><SPAN>
					 <FONT face=Arial size=2><b>
						Input/Modify
					</b></FONT>	
					</SPAN></a></LI>
				<LI id="1" name="1"><a href="#"  OnClick="JavaScript: NavigateURL(1)" title="Reports"><SPAN>
					<FONT face=Arial size=2><b>
						Reports
					</b></FONT>	
					</SPAN></a></LI>
			</UL>
		</DIV>
	</TD>
	<td valign=Top align=left >&nbsp <font color='green' size=2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<b id="ReportTitle">Report Title</b></font><td>
	<td valign=Top align=right >&nbsp <font face=Arial size=2>&nbsp;</font><td>
	</TR>
</TABLE>

<input type=hidden name="InputURL">
<input type=hidden name="CurrentMenuID">

<Script Language="JavaScript">
	var currentViewedMnuID;
	
	
	function launchReports(){		
		window.showModalDialog("reports_dialog.asp?ID=" + document.all.item("CurrentMenuID").value, null, "dialogWidth:300px;dialogHeight:450px");
	}
	
	function selectCurrentTab(tabCount){
		try{
			//highlight current tab
			var prevObj = document.getElementById('current');
			if (prevObj !== null && prevObj !== 'undefined') prevObj.id = prevObj.name;
			var currObj = document.getElementById(tabCount);
			if (currObj !== null && currObj !== 'undefined') currObj.id = 'current';
		}
		catch(e){}		
	}
		
	function NavigateURL(urlType){
		var urlTo = new String();
		var urlInput = new String();
		
		document.all.item("ReportTitle").innerHTML = "";
		switch(urlType){
			case 0:	
					urlTo = document.all.item("InputURL").value;
					if (urlTo!==""){
						window.parent.frames["maininfo"].location.replace(urlTo);											
						document.all.item("InputURL").value = "";
						window.parent.frames["contents"].UpdateToolBar(document.all.item("CurrentMenuID").value);
					}				
				break;
			case 1:			
				urlTo = document.all.item("CurrentMenuID").value;					
				urlInput = document.all.item("InputURL").value;
				if (urlTo!==""){					
					if (currentViewedMnuID !== urlTo)	currentViewedMnuID = urlTo;
					else {
						//if (urlInput !== "")	break;
					}	
					if (urlInput=="") document.all.item("InputURL").value = window.parent.frames["maininfo"].location;
					//urlTo = "Reports/ReportsList.asp?mnuID=" + urlTo;
					urlTo = "Reports.asp?mnuID=" + urlTo;
					window.parent.frames["maininfo"].location.replace(urlTo);
					window.parent.frames["contents"].UpdateToolBar(-1);
				}
				
				break;		
			case 2:			
				urlTo = document.all.item("CurrentMenuID").value;					
				urlInput = document.all.item("InputURL").value;
				if (urlTo!==""){					
					if (currentViewedMnuID !== urlTo)	currentViewedMnuID = urlTo;					
					if (urlInput=="") document.all.item("InputURL").value = window.parent.frames["maininfo"].location;
					//urlTo = "Reports/ReportsList.asp?mnuID=" + urlTo;
					urlTo = "/imports/import.asp?mnuID=" + urlTo;
					window.parent.frames["maininfo"].location.replace(urlTo);
					window.parent.frames["contents"].UpdateToolBar(-2);
				}
				break;			
		}
		
		selectCurrentTab (urlType);
		
	}
	
	function DoPrint(){
		var printWin = window.parent.frames["maininfo"];
		try{
			printWin.document.all.item("BottomDiv").style.display = "none";
		}
		catch(e){}
		
		printWin.focus();
		printWin.print();
		
		try{
			printWin.document.all.item("BottomDiv").style.display = "";
		}
		catch(e){}
		
		
	}
	
	
</Script>
</BODY>
</HTML>