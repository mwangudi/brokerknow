
// alert("fhsupport_js Start");


function faqInitialize() {
	if(document.getElementsByTagName){
		gss_hidefaqs();
		var faqarray = new Array();
		faqarray = document.getElementsByTagName("div");
		for(i=0;i<faqarray.length;i++){
			if(faqarray[i].id.substring(0,6) == "faqdiv"){
				faqswitch(faqarray[i].id.substring(6,faqarray[i].id.length));
				i = faqarray.length+1;
			}
		}
	}
}

function faqProdFaqIntialize() {
	if(document.getElementsByTagName){
		if(document.faqform) {
			var optionnum = 1;
			var faqarray = new Array();
			var displayString = new String();
			faqarray = document.getElementsByTagName("div");
			for(i=0;i<faqarray.length;i++){
				if(faqarray[i].className == "faqHeaderText"){
					displayString = faqarray[i].innerText;
					if(displayString.lastIndexOf(">") > 0) {
						displayString = displayString.substr(displayString.lastIndexOf(">")+1);
					}
					document.faqform.faqsection.options[optionnum] = new Option(displayString, 'faq' + faqarray[i].id);
					optionnum++;
				}
			}
		}
	}
}

function faqswitch(idnum)
{
	if(document.getElementById){
		if(document.getElementById("faqdiv"+idnum)){
			var elem = document.getElementById("faqdiv"+idnum);
			var hdr = document.getElementById("faqHeader"+idnum)
			if(elem == null) { return; }
			if(elem.style.display == "none")
			{
				if(hdr != null)	{	hdr.className="faqHeaderOpen";	}
				elem.style.display = "block";
				if(document.getElementById("faqplus"+idnum)){document.getElementById("faqplus"+idnum).style.display = "none";}
				if(document.getElementById("faqminus"+idnum)){document.getElementById("faqminus"+idnum).style.display = "";}
				if(document.getElementById("chevdown"+idnum)){document.getElementById("chevdown"+idnum).style.display = "none";}
				if(document.getElementById("chevup"+idnum)){document.getElementById("chevup"+idnum).style.display = "";}
			}
			else
			{
				elem.style.display = "none";
				if(hdr != null)	{	hdr.className="faqHeaderClosed";	}
				if(document.getElementById("faqplus"+idnum)){document.getElementById("faqplus"+idnum).style.display = "";}
				if(document.getElementById("faqminus"+idnum)){document.getElementById("faqminus"+idnum).style.display = "none";}
				if(document.getElementById("chevdown"+idnum)){document.getElementById("chevdown"+idnum).style.display = "";}
				if(document.getElementById("chevup"+idnum)){document.getElementById("chevup"+idnum).style.display = "none";}
			}
		}
		window.event.cancelBubble = true;
		return false;
	}

}


function gss_hidefaqs()
{
	if(document.getElementsByTagName){
		var divarray = new Array();
		divarray = document.getElementsByTagName("div");
		for(i=0; i<divarray.length; i++){
			if(divarray[i].id){
				if(divarray[i].id.indexOf("faqdiv") > -1){
					if(divarray[i].style.display != "none"){
						faqswitch(divarray[i].id.replace("faqdiv", ""));
					}
				}
			}
		}
	}
}

function gss_focusToc(focusid){
	if(document.getElementById("faq"+focusid)) top.location.href = top.location.href + "#faq" + focusid ;
}

function gss_HideCategoryToc(clid){
	if(document.getElementById("faq"+clid)){
		if(document.getElementsByTagName){
			var faqarray = new Array();
			faqarray = document.getElementsByTagName("div");
			// set all hidden
			for(i=0;i<faqarray.length;i++)
			{
				if(faqarray[i].className == "faqCategoryContainer"){faqarray[i].style.display = "none";}
			}
			faqarray = document.getElementsByTagName("UL");
			for(i=0;i<faqarray.length;i++)
			{
				if(faqarray[i].id.indexOf("faqTocEntry") > -1) {faqarray[i].style.display = "none";}
			}			
			// except this one
			// if(document.getElementById("faq"+clid).style.display == "none")
			document.getElementById("faq"+clid).style.display = "";
			if(document.getElementById("faqTocEntry"+clid)) {
				document.getElementById("faqTocEntry"+clid).style.display = "";
			}
			faqshow(clid, 'open');
		}
	}	
}

function faqshow(idnum, prop)
{
	if(document.getElementById)
	{
		if(document.getElementById("faqdiv"+idnum)) {
			var elem = document.getElementById("faqdiv"+idnum);
			var hdr = document.getElementById("faqHeader"+idnum)
			if(elem == null) { return; }
			if (prop == "open")
			{
				if(hdr != null)	{	hdr.className = "faqHeaderOpen";	}
				elem.style.display = "block";
				if(document.getElementById("faqplus"+idnum)){document.getElementById("faqplus"+idnum).style.display = "none";}
				if(document.getElementById("faqminus"+idnum)){document.getElementById("faqminus"+idnum).style.display = "";}
				if(document.getElementById("chevdown"+idnum)){document.getElementById("chevdown"+idnum).style.display = "none";}
				if(document.getElementById("chevup"+idnum)){document.getElementById("chevup"+idnum).style.display = "";}
			}else if (prop == "close") {
				elem.style.display = "none";
				if(hdr != null)	{ hdr.className="faqHeaderClosed";	}
				if(document.getElementById("faqplus"+idnum)){document.getElementById("faqplus"+idnum).style.display = "";}
				if(document.getElementById("faqminus"+idnum)){document.getElementById("faqminus"+idnum).style.display = "none";}
				if(document.getElementById("chevdown"+idnum)){document.getElementById("chevdown"+idnum).style.display = "";}
				if(document.getElementById("chevup"+idnum)){document.getElementById("chevup"+idnum).style.display = "none";}
			}
		}
	}
}

function clickExpandCollapse(){
	if(document.getElementsByTagName){
		var faqarray = new Array();
		faqarray = document.getElementsByTagName("div");
		for(i=0; i<faqarray.length; i++){
			if(faqarray[i].id.substring(0,6) == "faqdiv"){
				if(document.getElementById("ExpandCollapse").innerHTML == "+ Show All"){if(faqarray[i].style.display == "none") faqswitch(faqarray[i].id.substring(6,faqarray[i].id.length));}
				else{if(faqarray[i].style.display == "") faqswitch(faqarray[i].id.substring(6,faqarray[i].id.length));}
			}
		}
		if(document.getElementById("ExpandCollapse").innerHTML == "+ Show All"){
			document.getElementById("ExpandCollapse").innerHTML="- Hide All";
		}else{
			document.getElementById("ExpandCollapse").innerHTML="+ Show All";
		}
	}
}

var shownsection = 0;
function sortfaq(){
	if(document.getElementsByTagName){
		var faqarray = new Array();
		faqarray = document.getElementsByTagName("div");
		shownsection = document.faqform.faqsection.selectedIndex;
	
		if(document.faqform.faqsection.selectedIndex < 1){
			for(i=0;i<faqarray.length;i++){
				if(faqarray[i].className == "faqCategoryContainer"){faqarray[i].style.display = "block";}
				if(faqarray[i].className == "faqHeaderLeftSwitches"){faqarray[i].style.visibility = "visible";}
				if(faqarray[i].className == "faqHeaderRightSwitches"){faqarray[i].style.visibility = "visible";}
			
			}
		}else{
			for(i=0;i<faqarray.length;i++){
				if(faqarray[i].className == "faqCategoryContainer"){faqarray[i].style.display = "none";}
				if(faqarray[i].className == "faqHeaderLeftSwitches"){faqarray[i].style.visibility = "hidden";}
				if(faqarray[i].className == "faqHeaderRightSwitches"){faqarray[i].style.visibility = "hidden";}
			}
			var elem = document.getElementById(document.faqform.faqsection.options[document.faqform.faqsection.selectedIndex].value);
			if(typeof(elem) != "undefined") {
				elem.style.display = "block";
				var idnum = elem.id.substring(3, elem.id.length);
				if(document.getElementById("faqdiv"+idnum).style.display == "none") document.getElementById("faqdiv"+idnum).style.display = "" ;
			}
		}
	}
}

function sortNonProdfaq(clid){
    if(document.getElementById("faq"+clid))
    {
	if(document.getElementsByTagName){
		var faqarray = new Array();
		faqarray = document.getElementsByTagName("div");
		
			for(i=0;i<faqarray.length;i++){
				if(faqarray[i].className == "faqcontainer"){faqarray[i].style.display = "none";}
				if(faqarray[i].className == "faqswitches"){faqarray[i].style.display = "none";}
			}
			
			var elem = document.getElementById("faq"+clid);
			elem.style.display = "";
			if(document.getElementById("faqdiv"+clid).style.display == "none") faqswitch(clid);
			shownsection = parseInt(clid);
			
		}
	}	
}

function PartWrapperToggle(elementName, wapi) {
	var HeaderElement = null;
	var BodyElement = null; 
	if(document.getElementsByName) {
		HeaderElement = document.getElementsByName(elementName+"Header");
		BodyElement = document.getElementsByName(elementName+"Body");
		UpImage = document.getElementsByName(elementName+"Up");
		DownImage = document.getElementsByName(elementName+"Down");
		if(BodyElement) {	
			if(BodyElement[0].style.display == "none") {
				BodyElement[0].style.display = "block";
				//HeaderElement[0].className = "ListNuggetHeader";
				//currObjMenu = BodyElement[0];
				DownImage[0].style.display = "none";
				UpImage[0].style.display = "block";
				//cnt = 1;
				//origHeight = BodyElement[0].style.height;				
				//ShowMenu();
			} else {
				BodyElement[0].style.display = "none";
				//HeaderElement[0].className = "ListNuggetHeaderClosed";
				//currObjMenu = document.all.item(elementName+"Body")
				//currObjMenu = BodyElement[0];
				UpImage[0].style.display = "none";
				DownImage[0].style.display = "block";
				//cnt = 1;
				//origHeight = BodyElement[0].style.height;
				//HideMenu();
			}				
			
			
		}	
	}
	
	if (wapi!='link') {
		window.event.cancelBubble = true;
	}
	return false;

}


