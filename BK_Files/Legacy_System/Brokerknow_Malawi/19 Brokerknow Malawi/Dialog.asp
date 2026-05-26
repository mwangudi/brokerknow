
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2//EN">
<html dir=ltr>
<head>
	<title><%= Trim(Request.QueryString("titleDoc"))  & Replace(Space(700), " ", "&nbsp;") %></title>
	<META HTTP-EQUIV="Content-Type" content="text/html; charset=Windows-1252">
	<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>
</head>
<BODY Class=Dialog leftMargin=0 topMargin=0 marginwidth="0" marginheight="0" OnLoad="JavaScript: Init();" SCROLL="NO">
<Script Language="JavaScript">	
	function Init(){
	    var myLoc = new String();
	    var searchStr = "dialogDoc=";
	    myLoc = window.location.search;
	    
	    var myPos = myLoc.indexOf(searchStr);
	    if (myPos=="-1") window.self.close();	
	    else{
			myPos = myPos + searchStr.length;
			var myPage = myLoc.substr(myPos, myLoc.length - myPos);
			//var xHeight = document.body.offsetHeight;
			//*var xWidth = document.body.offsetWidth; 
			var xHeight = "100%";
			var xWidth = "100%";
			document.body.innerHTML += '<IFRAME OnLoad="UpdateFormAttributes(this)" marginwidth="5" marginheight="2" FRAMEBORDER=0 SRC="' + myPage + '" ID="dialogFrame" NAME="dialogFrame" height="' + xHeight + '" width="' + xWidth + '"></IFRAME>';
			//window.open(myPage);
	    }	    
	}
	
	function UpdateFormAttributes(winFrame){
		try{
			var winFrame = window.parent.frames["dialogFrame"];
			
			if (winFrame.document.forms.length > 0){	
				for (i=0; i <= winFrame.document.forms.length; i++){							
					var wkingForm = winFrame.document.forms[i];
					wkingForm.target = "deleteFrame";
					wkingForm.onsubmit = winFrame.UpdateDialogHandle;
				 }
				}
			}
		catch(e){}	
	
		
	}
	
	
</Script>

</BODY>
</HTML>