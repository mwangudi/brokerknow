<SCRIPT LANGUAGE="JavaScript">
		// This script is intended for use with a minimum of Netscape 4 or IE 4.
		if(document.getElementById) {
			var upLevel = true;
			}
		else if(document.layers) {
			var ns4 = true;
			}
		else if(document.all) {
			var ie4 = true;
			}

		function showObject(obj) {
			if (ns4) obj.visibility = "show";
			else if (ie4 || upLevel) obj.style.visibility = "visible";
			}
		function hideObject(obj) {
			if (ns4) {
				obj.visibility = "hide";
				}
			if (ie4 || upLevel) {
				obj.style.visibility = "hidden";
				}
			}

	</SCRIPT>

	