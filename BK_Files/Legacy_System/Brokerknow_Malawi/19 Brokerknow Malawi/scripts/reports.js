var cal, cal2;
	
function report_drawFilterInterface(fieldTypei, i){
	var interfaceTD, r;
	interfaceTD = document.all.item("inputTD" + i);
		
	r = parseInt(fieldTypei);
	if (r=="NaN" || fieldTypei==""){
		interfaceTD.innerHTML = "&nbsp;";			
		return;
	}
		
	checkToDefaultOn(i) = false;
		
	switch (myFieldTypes(fieldTypei)){
		case "0": // 0 = text, default
			interfaceTD.innerHTML = "<INPUT TYPE=TEXT NAME=SearchValue" + i + ">";
			document.all.item('SearchValue' + i).focus();
			break;
		case "1": 	//1 = date
			cal = new ctlSpiffyCalendarBox("cal", "frmSearch", "SearchValue" + i, "cmdDate", "<%= FormatDate(Date) %>", 1)
			cal.returnOutStringOnWrite();
			interfaceTD.innerHTML = cal.writeControl();
			break;
		case "2": 	//2 = number	
			interfaceTD.innerHTML = '<INPUT TYPE=TEXT NAME=SearchValue' + i + ' OnKeyDown="persistNumbers();">';
			document.all.item('SearchValue' + i).focus();
			break;
		case "3": 	//3 = date range
			cal = new ctlSpiffyCalendarBox("cal", "frmSearch", "SearchValueFrom" + i, "cmdDate", "<%= FormatDate(Date) %>", 1)
			cal2 = new ctlSpiffyCalendarBox("cal2", "frmSearch", "SearchValueTo" + i, "cmdDate2", "<%= FormatDate(Date) %>", 1)
			cal.returnOutStringOnWrite();
			cal2.returnOutStringOnWrite();
			interfaceTD.innerHTML = "From:&nbsp;" + cal.writeControl() + "<p>To:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" +   cal2.writeControl() ;
			checkToDefaultOn(i) = true;
			break;
		case "4": 	//4 = number range
			interfaceTD.innerHTML = 'From:&nbsp;<INPUT TYPE=TEXT NAME=SearchValueFrom' + i + ' OnKeyDown="persistNumbers();">';
			interfaceTD.innerHTML += '<p>To:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME=SearchValueTo' + i + ' OnKeyDown="persistNumbers();">';
			checkToDefaultOn(i) = true;
			break;	
		case "5":	//5 = boolean field	
			interfaceTD.innerHTML = '<INPUT TYPE=radio class=borderless NAME=SearchValueBool' + i + '1 value=True id=chkOpt' + i + '><label for=chkOpt' + i + '>True</label>';
			interfaceTD.innerHTML += '&nbsp;<INPUT TYPE=radio class=borderless NAME=SearchValueBool' + i + ' value=False id=chkOpt2' + i + '><label for=chkOpt2' + i + '>False</label>';
			checkToDefaultOn(i) = true;
			break;	
		case "6":	//6 = boolean field, yes no	
			interfaceTD.innerHTML = '<INPUT TYPE=radio class=borderless NAME=SearchValueBool' + i + '1 value=True id=chkOpt1><label for=chkOpt1>Yes</label>';
			interfaceTD.innerHTML += '&nbsp;<INPUT TYPE=radio class=borderless NAME=SearchValueBool' + i + ' value=False id=chkOpt2><label for=chkOpt2>No</label>';
			checkToDefaultOn(i) = true;
			break;	
				
		default:
			interfaceTD.innerHTML = "<INPUT TYPE=TEXT NAME=SearchValue" + i + ">";
			document.all.item('SearchValue' + i).focus();
		
	break;	
	}
		
}
	
function persistNumbers(){
		 event.returnValue = IsANumber(event.keyCode);	 
}
	
	
function checkToDefault(theSel, iTurn){
	if (checkToDefaultOn(iTurn)) theSel.selectedIndex = 0;
}


function report_validateFrm(frm){			
	var doc = document.getElementById('customCols');
	if (doc.length >= 0) SelectAll(doc);			
	DoSearch();
	DoSort(); 
	//if (window.parent.name !== 'KNWNG' && window.name !== 'maininfoR')
	frm.target = '_self';			
	//else
		//frm.target = '_blank';
					
	frm.submit();
}
		
function report_SetBodyClass(){
	try{
				
		if (window.parent.name !== 'KNWNG' && window.name !== 'maininfoR')
				document.body.className = 'dialog';
	}			
	catch(e){}
}


function switchDisplay(obj){
	if (obj.style.display=='none') obj.style.display = '';
	else obj.style.display = 'none';
}
		
 function SelectAll(Object){
 //select all upwards
  for (loop=Object.length-1; loop>-1; loop--)	
    {	
     Object.options[loop].selected = true
    }	
 }

function moveDown(){
	var doc = document.all.item('customCols');
	var currIndex = doc.selectedIndex;
	if (currIndex >= 0){
		try{
			var nextOptionIndex = currIndex - 1 + 2;
			var tempOptionText = doc.options[nextOptionIndex].text;
			var tempOptionValue = doc.options[nextOptionIndex].value;
			doc.options[nextOptionIndex].text = doc.options[currIndex].text;
			doc.options[nextOptionIndex].value = doc.options[currIndex].value;
			doc.options[currIndex].text = tempOptionText;
			doc.options[currIndex].value = tempOptionValue;
			doc.selectedIndex = nextOptionIndex;
		}	
		catch(e){}
	}

}

function moveUp(){
	var doc = document.all.item('customCols');
	var currIndex = doc.selectedIndex;
	if (currIndex >= 0){
		try{
			var nextOptionIndex = currIndex - 1;
			var tempOptionText = doc.options[nextOptionIndex].text;
			var tempOptionValue = doc.options[nextOptionIndex].value;
			doc.options[nextOptionIndex].text = doc.options[currIndex].text;
			doc.options[nextOptionIndex].value = doc.options[currIndex].value;	
			doc.options[currIndex].text = tempOptionText;
			doc.options[currIndex].value = tempOptionValue;
			doc.selectedIndex = nextOptionIndex;
		}
		catch(e){}
	}
}
		
function report_AddOption(Input,Output){    
    	NewOption = new Option();   			    
	    NewOption.text = Input.id;
	    NewOption.value = Input.value;
	    NewOption.selected = false;	
		Output.add(NewOption, 0);   		  	    
 }
 		 
 function report_RemoveOption(Input, Field){
	Selection = new Boolean();
	for (loop=Field.length - 1; loop >= 0; loop--) {
		 var GoneOption = Field.options[loop]
		if (GoneOption.value==Input.value) {
  			Selection = true;
  			Field.remove(GoneOption.index);
  		}
  	}	
}
		
function evalCheck(chk){
	if (chk.checked){
		report_AddOption(chk, document.getElementById('customCols'));
	}
	else {
		report_RemoveOption(chk, document.getElementById('customCols'));
	}
}
