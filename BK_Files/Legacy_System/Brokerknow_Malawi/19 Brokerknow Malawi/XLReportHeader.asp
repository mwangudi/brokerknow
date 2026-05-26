<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>XL Report Header</title>
  
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
 

</head>

<body Class="Report">


<table border=0 cellpadding=2 cellspacing=4>
	<tr name=functionRow id=functionRow>
	    <td OnClick="Javascript: printFramedXLDoc()" nowrap class=nav onMouseover="JavaScript: if (this.className=='nav') this.className='nav_over';" onMouseout="JavaScript: if (this.className=='nav_over') this.className='nav';">
			<img src="../images/printLink.gif" border=0 alt=Print>Print
	    </td>
	    <td OnClick="JavaScript: window.parent.window.parent.location.reload();;" nowrap class=nav onMouseover="JavaScript: if (this.className=='nav') this.className='nav_over';" onMouseout="JavaScript: if (this.className=='nav_over') this.className='nav';">
			<font color=blue face="Webdings" size=1>3</font>Back
	    </td>
	     <td OnClick="javascript: window.parent.window.parent.close()" nowrap class=nav onMouseover="JavaScript: if (this.className=='nav') this.className='nav_over';" onMouseout="JavaScript: if (this.className=='nav_over') this.className='nav';">
			<font color=blue face="Wingdings" size=3>x</font>Close
	    </td>
	    
	</tr>
</table>

<Script Language="JavaScript">
	function printXLDoc(){
		window.parent.frames("maininfoR").print();
	}
	
	
</Script>
</body>

</html>
