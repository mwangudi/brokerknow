/*
	'File Name		: valjavavalidate.asp
	'File Date		: 15th August 2004
	'Developer		: Silas Njoroge
	'Modified By	: Developer name 
	'Modified Date	: 17th August 2004
	'Comments		: function to validate in javascript
	1)  string trim(string)
	2)  boolean checkempty(object,errmsg)
	3)	boolean checkemail(object,errmsg)
	4)	boolean chemnumeric(object,errmsg)	
	5)	boolean checktelno(object,errmsg)
	6)  boolean checkmaxlength(object,length,errmsg)
	7)	boolean checkminlength(object,length,errmsg)
	8)  boolean checkinteger(object,errmsg)
*/	

//Function trim similiar to Visual Basic Trim()
//Removes Leading and trailing spaces and tabs from the argument passed
//returns a string
//trim all the required fields using this function
function trim(str)
{
	var x;
	var ch;
	
	for(x=0;x<str.length;x++)
	{
		ch=str.substr(x,1);
		if(ch==' ' || ch=='\t')
		{
			str=str.substr(x+1,str.length-1);
		}
		else
			break;
	}
	
	for(x=str.length-1;x>=0;x=x-1)
	{
		ch=str.substr(x,1);
		if(ch==' ' || ch=='\t')
		{
			str=str.substr(0,str.length-1);
		}
		else
			break;
	}
	
	return str;
}



	//this function accepts the object as a parameter and display the appropriate error messages if the text box is empty
	function checkempty(objname,errmsg)
	{
		if (trim(objname.value)=="")
		{
			alert(errmsg);
			objname.focus();
			return false
		}
		return true;
	}

	//this function accepts the object as a parameter and display the appropriate error messages if the text box is empty
	//similar to the above function only that it checks for combo box
	function checkcomboempty(objname,errmsg)
	{
		if (trim(objname[objname.selectedIndex].value)=="")
		{
			alert(errmsg);
			objname.focus();
			return false
		}
		return true;
	}

	function checknumeric(objname,errmsg)
	{
		if (isNaN(trim(objname.value)))
		{
			alert(errmsg);
			objname.focus();
			return false;
		}
		return true;
	}
//Email validation starts from here
	function checkemail(objname,errmsg)
	{
		vvalue=trim(objname.value);
		atPos = vvalue.indexOf('@');
		sppos = vvalue.indexOf(" ");
		dopos = vvalue.indexOf(".");
		if (atPos < 1 || atPos == (vvalue.length - 1) || (sppos != -1)|| (dopos == -1))
		{
			alert(errmsg);
			objname.focus();
			return false;
		}
		return true;
	}
//Email validation ends from here

//telno validation starts from here 
//it checks that the entered value does not contain anything except numeric characters and hyphen
	function checktelno(objname,errmsg)
	{
		var str = trim(objname.value);
		for(x=0;x<str.length;x++)
		{
			ch=str.substr(x,1);
			if ((ch < '0' || ch >'9')&&(ch!='-')&&(ch!=' ')&&(ch!='/')&&(ch!='(')&&(ch!=')'))
			{
				alert(errmsg);
				objname.focus();
				return false;
			}
		}
		return true;
	}
//telno validation ends over here

//maxlength validation starts over here
	function checkmaxlength(objname,maxlength,errmsg)
	{
		var str = objname.value;
		if (str.length>maxlength)
		{
			alert(errmsg);
			objname.focus();
			return false;
		}
		return true;
	}
//maxlength validation ends over here

//minlength validation starts over here
	function checkminlength(objname,minlength,errmsg)
	{
		var str = objname.value;
		if (str.length<minlength)
		{
			alert(errmsg);
			objname.focus();
			return false;
		}
		return true;
	}
//minlength validation ends over here
	function checkinteger(objname,errmsg)
	{
		var str = trim(objname.value);
		for(x=0;x<str.length;x++)
		{
			ch=str.substr(x,1);
			if ((ch < '0' || ch >'9'))
			{
				alert(errmsg);
				objname.focus();
				return false;
			}
		}
		return true;
	}
/*			code to validate the radio buttons
			done by silas

*/
	function checkradio(objname,errmsg){
		//set a variable radio_choice to false
				var radio_choice = false;

				for (counter = 0;  counter < objname.length; counter++)
				{
				// If a radio button has been selected it will return true
				// (If not it will return false)
				if (objname[counter].checked)
				radio_choice = true; 
				}

				if (!radio_choice)
				{
				// If there were no selections made display an alert box 
				alert(errmsg)
				return (false);
				}
		}


// validation starts from here
	function checknum(objname,errmsg)
	{
		var check=false;
		var str = trim(objname.value);
		for(x=0;x<str.length;x++)
		{
			ch=str.substr(x,1);
			if ((ch =='0'||ch =='1'||ch =='2'||ch =='3'||ch =='4'||ch =='5'||ch =='6'||ch =='7'||ch =='8'||ch =='9'||ch =='0' ))
			check=true
				
		}
			if (!check)
				{
				alert(errmsg);
				objname.focus();
				return false;
				}
		
		return true;
	}
		function checkchar(objname)
	{
		var check=false;
		var str = trim(objname.value);
		for(x=0;x<str.length;x++)
		{
			ch=str.substr(x,1);
			if ((ch ==' '||ch =='+'||ch ==':'||ch =='!'||ch ==';'||ch =='~'||ch =='@'||ch =='#'||ch =='$'||ch =='`'||ch =='^'||ch =='%'||ch =='&'||ch =='?'||ch =='('||ch ==')' ||ch =='='||ch =='-'||ch =='_'||ch =='\/'||ch =='\\'||ch =='"'||ch =='\''||ch =='['||ch ==']'||ch =='{'||ch =='}'||ch ==','||ch=='.'  ))
			check=true
				
		}
			if (check)
				{
				alert('Field should not contain characters "~!@#$$%^&()_+ -?\><,.\'{}[]=`\/\\" or spaces');
				objname.focus();
				return false;
				}
		
		return true;
	}

/* validate the textfield for numerals*/

function dodacheck(val,errmsg) {
var characters = /[$\\@\\\#%\^\&\*\(\)\[\]\+\_\{\}\`\~\=\0123456789]/;
var strPass = val.value;
var strLength = strPass.length;
var lchar = val.value.charAt((strLength) - 1);
if(lchar.search(characters) != -1) {
var tst = val.value.substring(0, (strLength) - 1);
val.value = tst;
return (false)
   }
}

//Date validation
function isLeap(year){
    if(year % 400 == 0){
        return true;
    } else if((year % 4 == 0) && (year % 100 != 0)){
        return true
    } else return false;
};
    
function days_in(month, year){
    if(month == 4 || month == 6 || month == 9 || month == 11){
        return  30;
    } else if(!isLeap(year) && month == 2){
        return 28;
    } else if(isLeap(year) && month == 2){
        return 29;
    }
	else 
	
	return 31;
};
    
function checkDate(myDob){
    var myArrayDate, myDay, myMonth, myYear, myString, myYearDigit;
    myString = myDob.value + "";

    myArrayDate = myString.split("/");

   // myDay = Math.round(parseFloat(myArrayDate[1]));
    //myMonth = Math.round(parseFloat(myArrayDate[0]));
    //myYear = Math.round(parseFloat(myArrayDate[2]));
  myDay = Math.round(parseFloat(myArrayDate[0]));
   myMonth = Math.round(parseFloat(myArrayDate[1]));
    myYear = Math.round(parseFloat(myArrayDate[2]));

    myString = myYear + "";
    myYearDigit = myString.length;
  
    if (isNaN(myDay) || isNaN(myMonth) || isNaN(myYear) || (myYear < 1) || (myDay < 1) || (myMonth < 1) || (myMonth > 12) || (myYearDigit != 4) || (myDay > days_in(myMonth, myYear))){
		alert("Please enter a valid date. (dd/mm/yyyy)");
		myDob.focus();
        return false;
    } 
	else{
        return true;
  }

};