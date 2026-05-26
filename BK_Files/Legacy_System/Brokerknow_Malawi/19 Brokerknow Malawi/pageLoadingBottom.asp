<%Response.Flush%>
	<SCRIPT LANGUAGE="JavaScript">
	if(upLevel) {
		var splash = document.getElementById("splashScreen");
		}
	else if(ns4) {
		var splash = document.splashScreen;
		}
	else if(ie4) {
		var splash = document.all.splashScreen;
		}
	hideObject(splash);
	</SCRIPT>