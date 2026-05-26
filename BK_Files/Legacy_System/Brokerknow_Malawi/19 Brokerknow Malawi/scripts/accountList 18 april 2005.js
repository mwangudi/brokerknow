
function FilterData(obj,entityType,retVal)
{
	switch (entityType)
	{
	case 1:
		FilterOrderList(obj);
		break;
	case 3:
		FilterContracts(obj);
		break;
	case "1":
		FilterOrderList(obj);
		break;
	case "3":
		FilterContracts(obj);
		break;
	default:
		return retVal;
	}
			
	return retVal;
}
		
function FilterContracts(obj)
{
	var IDName = "";
	var frameName = "";
	var framePageName = "";
	var voucherParamName = "";
	var saleType = "";
	var selValue = "";
			
	IDName = "BrokerCode";
	frameName = "fraInnerBrokerSelects";
	framePageName = "inner_select_voucherBroker";
	voucherParamName = "BrokerReceiptVouchered";
	saleType = "1";
	selValue = obj.options[obj.selectedIndex].SearchCode;
				
			
	var fra = document.getElementById(frameName);
			
	var pageTo = framePageName + '.asp?' + voucherParamName + '=0&OrderTypeSale=' + saleType + '&' + IDName + '=' + selValue; 
			
	if(selValue.length == 0)
	{
		pageTo = framePageName + '.asp';
	}
			
	fra.src = pageTo;
}
		
function FilterOrderList(theList)
{
	var i = 0;
	var clientNo = theList.value;
	var orderList = document.frmMain.cboOrder;
	var orderBagList = document.frmMain.cboOrderBag;
	

	
	RemoveOptions(orderList)  ;
			
	//add default selection (none)
	var NewOption = new Option();   			    
   	NewOption.text = '';
   	NewOption.value = '0';
   	NewOption.selected = true;
   	orderList.add(NewOption, 0);
			
			
	//ShowMessage(orderList.options(3).ClientTag);
	for (i=0; i < orderBagList.options.length; i++) {
		if((orderBagList.options(i).ClientTag == clientNo))
		{
			var NewOption = new Option();   			    
   		    NewOption.text = orderBagList.options[i].text;
   		    NewOption.value = orderBagList.options[i].value;
   		    orderList.add(NewOption, 0);
		}
				
	}
}

function GetAccountsList(entity,toList,filterRequired)
{
	currentEntityType = entity;
	frm = document.frmMain;				
	xmlhttp = createXMLHTTPObj();
			
	url="GetList.asp?ID="+entity+"&action=GetAccountList";
	xmlhttp.open("GET",url,true);
	xmlhttp.onreadystatechange=function() {
		if (xmlhttp.readyState==4) {
		returnStr = xmlhttp.responseText;
		returnStr = getBodyHTML(returnStr);
				
		var secList = "<select name = '" + toList.name + "' id = '" + toList.name + "' size='1' ";
		
		if(filterRequired)
		{
			secList += "onChange='FilterData(this," + currentEntityType + ",UpdateCode(true,cboAccount,txtClientCode))' " ;
		}
		else
		{
			secList += "onChange='UpdateCode(true,cboAccount,txtClientCode)' " ;
		}
		
		secList += "onKeypress='return (dodefaultaction()==\"\"); ' "  ;
		secList += "onKeydown='return (dodefaultaction()==\"\");' " ; 
		
		if(filterRequired)
		{
			secList += "onKeyup='return (FilterData(this," + currentEntityType + ",UpdateCode(change(" + toList.name + ",0)," + toList.name + ",txtClientCode)));' " ;
		}
		else
		{
			secList += "onKeyup='return (UpdateCode(change(" + toList.name + ",0)," + toList.name + ",txtClientCode));' " ;
		}
		
		
		secList += "onfocus='txtval = \"\";inputIsItemCode = 1;' "  ;
		secList += "onblur='txtval = \"\";inputIsItemCode = 1;'>" ;
		secList += returnStr ;
		secList += "</select>";
				
		toList.outerHTML = secList;															
		}
		}
	xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
	xmlhttp.send();
	
	document.getElementById('txtClientCode').value = '';
}

function FetchAccounts(theList)
{			
	var i = 0;
	var entity = theList.value;
	var toList = document.frmMain.cboAccount;
			
	GetAccountsList(entity,toList,true);
			
	var fra = document.getElementById('fraInnerBrokerSelects');		
	var pageTo = 'inner_select_voucherBroker.asp';  
	fra.src = pageTo;
			
	totalContractAmt = 0;
	document.all.item("txtTotal").value = totalContractAmt;
	document.all.item("ContractsSel").value = "";	
			
	if (entity==3){
		document.getElementById('brokerVoucherRow').style.display = '';
		document.getElementById('orderRow').style.display = 'none';
		//document.getElementById('txtClientCode').style.display = 'none';
		document.getElementById('txtVoucherType').value = currentEntityType;
	} 
			
	else{
		 if (entity==1){
			document.getElementById('brokerVoucherRow').style.display = 'none';
			document.getElementById('orderRow').style.display = '';
			//document.getElementById('txtClientCode').style.display = '';
			document.getElementById('txtVoucherType').value = '0';
		}			
		else {
			document.getElementById('brokerVoucherRow').style.display = 'none';
			document.getElementById('orderRow').style.display = 'none';
			//document.getElementById('txtClientCode').style.display = 'none';
			document.getElementById('txtVoucherType').value = '0';
		}
	}
			
}

function FetchJournalAccounts(theList)
{
	var i = 0;
	var entity = theList.value;
	var toList = document.frmMain.cboAccount;
			
	GetAccountsList(entity,toList,false);	
}