
var d = document;

function gsfx_drawselectlist(po, co, pn){
	
	if(failure==0){

		var potl;

		if(po){
			if(po.type){
				potl = po.type.toLowerCase();
				if(potl.indexOf("select") > -1) ctl = po.selectedIndex;
				else ctl = 0;
			}else if(po.length){
				potl = "radio";
				ctl = -1;
				for(i=0; i<po.length; i++){
					if(po[i].checked) ctl = i;
				}
				if(ctl < 0) ctl = 0;
			}
		}else ctl = 0;
		
		var arraylen;
		if(co && co.options){
			arraylen = eval(pn+"_"+co.name+"_"+ctl+'.length;');
			if(co.children && navigator.appVersion.indexOf("Macintosh") < 0){
				while(co.children.length > 0){
					for(i=0;i<co.children.length;i++){
						co.remove(co.children[i]);
					}
				}
			}else for(i=0;i<co.options.length;i++){co.options[i] = null;}
		}
	
		var si = 0;
		var ms = true;

		if(po){
			if(potl.indexOf("select") > -1){
				if(po.options[ctl].value.indexOf("PRODLISTSRC=OFF") > -1) ms = false;
			}else if(potl.indexOf("hidden") > -1){
				if(po.value.indexOf("PRODLISTSRC=OFF") > -1) ms = false;
			}else if(potl == "radio"){
				if(po[ctl].value.indexOf("PRODLISTSRC=OFF") > -1) ms = false;
			}
		}
		
		if(co && co.options){
			if(ms){
				for(i=0;i<arraylen;i++){
					ts = eval(pn+"_"+co.name+"_"+ctl+'['+i+']');
					//eval("if("+pn+"_"+co.name+"_"+ctl+"default > 0) si = "+i+";");
					eval("co.options[i] = new Option(\""+ts+"\")");
					co.disabled = false;
				}
			si = eval(pn + "_" + co.name + "_" + ctl + "default");		
			}else{
				co.options[0] = new Option("", "");
				si = 0;
				co.disabled = true;
			}
			
			co.selectedIndex = si;
		}
	}
}

function gsfx_switchUISrchMode(unid){

	var e, c;

	if(d.getElementsByName){
		
		var dm = "";
		var adv;
		if(navigator.appVersion.indexOf("Mac") > -1){
			adv = d.getElementsByTagName("table");
			for(i=0; i<adv.length; i++){
				if(adv[i].name == unid+"_advanced"){
					if(dm == ""){
						dm = adv[i].style.display;
						if(dm == "none") dm = "block";
						else dm = "none";
					}
					adv[i].style.display = dm;
				}
			}
		}else{
			adv = d.getElementsByName(unid+"_advanced");
			if(adv.length > 0){
				dm = adv[0].style.display;
				if(dm == "none") dm = "block";
				else dm = "none";
			}
			for(i=0; i<adv.length; i++){adv[i].style.display = dm;}
		}
		if(adv.length > 0){
			e = d.getElementById(unid+"_expand");
			c = d.getElementById(unid+"_collapse");
		}
	}else if(d.all){
		var dm;
		for(i=0; i<d.all.length;i++){
			if(d.all[i].id == unid+"_advanced"){
				dm = d.all[i].style.display;
				break;
			}
		}

		if(dm == "none") dm = "block";
		else dm = "none";

		for(i=0; i<d.all.length;i++){
			if(d.all[i].id == unid+"_advanced") d.all[i].style.display = dm;
			if(d.all[i].id == unid+"_expand") e = d.all[i];
			if(d.all[i].id == unid+"_collapse") c = d.all[i];
		}
	}

	if(e && c){
		if(dm == "" || dm == "block"){
			e.style.display = "none";
			c.style.display = "block";
		}else{
			e.style.display = "block";
			c.style.display = "none";
		}
		setcookieval(unid+"_srchMode", dm);
	}
}

function getUrlParam(param){
	var rv = "";
	var q = new Array();
	if(param.indexOf("_") > 0) {
		argParam = param.split("_")[1];
	} else {
		argParam = param;
	}
	var here = document.location.toString();
	q = here.split("?");
	if(q.length > 1){
		var argarray = new Array();
		argarray = q[1].toString().split("&");
		var keyval;
		for(i=0; i<argarray.length; i++){
			keyval = argarray[i].toString().split("=");
			if(keyval[0].toLowerCase() == argParam.toLowerCase()){
				rv = keyval[1];
				break;
			}
		}
	}
	if(rv == "") {
		rv = unescape(fetchcookieval(param));
		if(rv == "blank") { rv = ""; }
	}
	return rv;
}

function gsfx_loadUISrchMode(unid){

	var e, c;
	var dm = fetchcookieval(unid+"_srchMode");
	if(dm == "blank") dm = "none";

	if(d.getElementsByName){
		var adv;
		if(navigator.appVersion.indexOf("Mac") > -1){
			adv = d.getElementsByTagName("table");
			for(i=0; i<adv.length; i++){
				if(adv[i].name == unid+"_advanced") adv[i].style.display = dm;
			}
		}else{
			adv = d.getElementsByName(unid+"_advanced");
			for(i=0; i<adv.length; i++){
				adv[i].style.display = dm;
			}
		}
		if(adv.length > 0){
			e = d.getElementById(unid+"_expand");
			c = d.getElementById(unid+"_collapse");
		}
	}else if(d.all){
		for(i=0; i<d.all.length;i++){
			if(d.all[i].id == unid+"_advanced") d.all[i].style.display = dm;
			if(d.all[i].id == unid+"_expand") e = d.all[i];
			if(d.all[i].id == unid+"_collapse") c = d.all[i];
		}
	}

	if(e && c){
		if(dm == "" || dm == "block"){
			e.style.display = "none";
			c.style.display = "block";
		}else{
			e.style.display = "block";
			c.style.display = "none";
		}
	}
}

function srch_AppendSelectScript(catalogElement,selectElement,webpartName,selectName){
gsfx_drawselectlist(catalogElement,selectElement,webpartName)
tval = fetchcookieval(webpartName + '_' + selectName )
if(tval != "blank"){
	for(i=0; i<selectElement.options.length; i++){
		if(selectElement.options[i].value == unescape(tval)){
			selectElement.selectedIndex = i;
			break;
		}
	}
}

}

function srch_AppendRadioScript(radioElement,webpartName,radioName){

tval = fetchcookieval(webpartName + '_' + radioName )
if(tval != "blank"){
	for(i=0; i<radioElement.length; i++){
		if(radioElement[i].value == unescape(tval)){
			radioElement[i].checked = true;
			radioElement[i].click();
			break;
		}
	}
}

}

function srch_AppendTextScript(textElement,webpartName,textName,defaultstring){
tval = fetchcookieval(webpartName + '_' + textName)
	if(tval != 'blank' && tval != '')
	{ 
		textElement.value =UnicodeFixup(unescape(tval)); 
	} 
	else 
	{ 
		textElement.value = defaultstring;
	}
}

function AdvanceSrchTextScript(textElement,webpartName,textName){
	var arg = getUrlParam(webpartName + '_' + textName);
	if(arg != ""){
		textElement.value = arg;
	}

}

function AdvanceSrchSelectScript(catalogElement,selectElement,webpartName,selectName){

gsfx_drawselectlist(catalogElement,selectElement,webpartName)
var arg = getUrlParam(webpartName + "_" + selectName);
if(arg != ""){
	for(i=0; i<selectElement.options.length; i++){
		if(selectElement.options[i].value.toLowerCase() == arg.toLowerCase())
		{	selectElement.selectedIndex = i;
			break;
		}
	}
}

}

function AdvanceSrchRadioScript(radioElement,webpartName,radioName){
var arg = getUrlParam(webpartName + "_" + radioName);
if(arg != ""){
	for(i=0; i<radioElement.length; i++){
		if(radioElement[i].value.toLowerCase() == arg.toLowerCase())
		{
			radioElement[i].checked = true;
			radioElement[i].click();
			break;
		}
	}
}
}

function AddSubmitSelectScript(selectElement,webpartName,selectName)
{
	srch_setcookieval(webpartName + "_" + selectName,escape(selectElement.options[selectElement.selectedIndex].value));
}


function AddSubmitTextScript(textElement,leftTextElement,webpartName,textName,defaultstring)
{
if(leftTextElement.value == defaultstring) leftTextElement.value = '';
	textElement.value = unescape(leftTextElement.value);
	srch_setcookieval(webpartName+ "_" + textName, UnicodeFixup(escape(leftTextElement.value)));
}

function AddSubmitRadioScript(radioElement,webpartName,radioName)
{

for(i=0;i<radioElement.length;i++){
		if(radioElement[i].checked){
			srch_setcookieval(webpartName + '_' + radioName,escape(radioElement[i].value));
			
		}
	}
}




function UnicodeFixup(s){
	var result = new String();
	var c = '';
	var i = -1; 
	var l = s.length;
	result = "";
	for(i = 0; i<l; i++) {
		c = s.substring(i, i+1);
		if(c == "%") {
			result += c; i++;
			c = s.substring(i, i+1);
			if(c != "u") {
				if(parseInt("0x" + s.substring(i, i+2)) > 128) result += "u00";
			}
		}
		result += c;
	}
	return result;
}