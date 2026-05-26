var reportCount = 0;
var ExemptFromDefaultTabIndex = false;
var ACCEPT_IMG_TYPES = 'PNG,GIF,JPG,JPEG,TIF,BMP';

function printFramedXLDoc(){
	document.all.item("printTable").style.display = 'none';
	window.self.print();
	document.all.item("printTable").style.display = '';	
}




function getBodyHTML(data){
	//grab all HTML Source	
	var thisDocHTML ;
	if (data==null || data=='undefined') thisDocHTML = window.document.body.parentNode.outerHTML;
	else thisDocHTML  = data;

	//remove the print tags and scripts
	var beginSplit, endSplit, firstPart, secondPart;
	var beginStr = "<!--BEGINPAGEFUNCS-->";
	var endStr = "<!--ENDPAGEFUNCS-->";

	
	beginSplit = thisDocHTML.indexOf(beginStr) ;
	endSplit = thisDocHTML.indexOf(endStr) ;
	if (endSplit!==-1) endSplit = endSplit - 0 + endStr.length;  

	firstPart = thisDocHTML.substr(0, beginSplit) 	;
	secondPart = thisDocHTML.substr(endSplit);	

	return (firstPart + secondPart) ;
	
		
}

function cancelMenu(){
	window.event.returnValue = false 
	window.event.cancelBubble = true;
}

function UpdateDialogHandle(){
	try{
		window.parent.dialogArguments.opener.parent.frames("footer").editDocOpener = window.self;
	}
	catch(e){}
}


function setDefaultAttributes(){	
	if (!ExemptFromDefaultTabIndex)	{
		var docForms = document.all.tags('FORM');
		for (i=0; i<docForms.length; i++)
		    setAttributes(docForms[i]);
	}

}

function setAttributes(frm){
	var defTags = new String();
	defTags = '  INPUT,SELECT,TEXTAREA';	
	var tCount = 1;
	var firstItemDone = new Boolean();
	for (el in frm.elements){
		try{
			var tagN = frm[el].tagName;
			var searchResult = defTags.search(tagN);
			if (searchResult !== -1 && searchResult !==0){
				frm[el].tabIndex = tCount; 				
				tCount ++;
				if (firstItemDone!==true){
					if (frm[el].type!=='hidden'){
						if (frm[el].readOnly==false && frm[el].style.display=='' && frm[el].style.visibility==''){
							frm[el].focus();
							firstItemDone = true;
						}
					}
				}
				
										
			}
		}
		catch(e){}
	}
	
	

}

function resizePagingDisplay(){
	try{			
		var bottomObj = document.all.item("BottomDiv");
		bottomObj.style.top = (document.body.clientHeight - bottomObj.clientHeight) - 5;					
	}
		catch(e){}	
}

function Paging(pageTo){
	try{
		document.all.item("Page").value = pageTo;
		document.forms[0].action = document.all.item("ActionPage").value;
		document.forms[0].submit();
	}
	catch(e){}
}

function updateBestPrice(tBox){
	try{
		if (tBox.value.toLowerCase()=='b'){
			tBox.value = 'BEST';
			document.all.item('txtCert').focus();
		}
	}
	catch(e){}	
	}

		
function format2NumberCommasOnly(tBox){
	var d = FormatNumCommasOnly(tBox.value);
	tBox.value = d;
}
		
function format2Number(tBox){
	var d = FormatNum(tBox.value);
	tBox.value = d;
}

function createXMLHTTPObj()			
				{
					 var xmlhttp;
					 try {
						xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
					 } catch (e) 
						{
							try {
								xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
								// For internet Explorer
							} catch (E)
							   {
									if (!xmlhttp && typeof XMLHttpRequest!='undefined') {
										xmlhttp = new XMLHttpRequest();
										// For Mozilla, Firefox, Safari, and Netscape
									}	
								}
						}	
						
						return 	xmlhttp;	
					}


//window.document.oncontextmenu = cancelMenu;/
//window.onload = setDefaultAttributes;

// autocomplete functionality		
var txtval = '';
var curlist;
var inputIsItemCode = 1;

function select(trigger)
{
	curlist = trigger;
	
	for (n=0; n < curlist.length; n++)
	{
		var theText;
		
		if(inputIsItemCode == 1)
		{
			theText = curlist[n].SearchCode;
		}
		else
		{
			theText = curlist[n].SearchText;
		}
	
		if(theText.toLowerCase().indexOf(txtval.toLowerCase())==0)
		{
			curlist.selectedIndex=n;
			break; 
		}
		else
		{ 
			curlist.selectedIndex=0;
		}
		
		
	}

}

function dodefaultaction(e)
{ 
	var code; 
	
	if (!e) var e = window.event;
	
	if (e.keyCode) 
		code = e.keyCode; 
		
	else 
		if (e.which) 
			code = e.which;
	
	if(code == '9' | code == '40' | code == '38') 
		return ''; 
	else 
		return code;
}

function change(trigger, useCode)
{
		var code = dodefaultaction();
	
		if(useCode) 
		{
			inputIsItemCode = 1;
		}
		else
		{
			inputIsItemCode = 0;
		}
		
		if(code == '')
		{
			txtval='';
			return false;
		} 
		else
		{
		if(code == '8')
		{ 
			txtval = txtval.substring(0,txtval.length - 1);
		} 
		else
		{
			txtval = txtval + String.fromCharCode(code);
		}

		select(trigger); 
		return true;
		}
}

function UpdateCode(retVal,codeSource,codeTarget)
{
	codeTarget.value = codeSource[codeSource.selectedIndex].SearchCode;
	return retVal;
}
function UpdateCodes(retVal,codeSource,codeTarget)
{
	codeTarget.value = codeSource[codeSource.selectedIndex].SearchCds;
	return retVal;
}

function selectItem(trigger)
{
	curlist=trigger;
	for (n=0;n<curlist.length;n++)
	{
		var theText;
		
		theText = curlist[n].SearchCode;
		if(theText.toLowerCase() == txtval.toLowerCase())
		{
			curlist.selectedIndex=n;
			return; 
		}
	}
	curlist.selectedIndex=0;
}

function selectItems(trigger)
{
	curlist=trigger;
	for (n=0;n<curlist.length;n++)
	{
		var theText;
		
		theText = curlist[n].SearchCds;
		if(theText.toLowerCase() == txtval.toLowerCase())
		{
			curlist.selectedIndex=n;
			return; 
		}
	}
	curlist.selectedIndex=0;
}

// end autocomplete functionality
