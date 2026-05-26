<%
 filetext=Request.QueryString("filetext")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>File Upload</TITLE>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>
<SCRIPT Language="JavaScript">
 
	function Upload(fileObj){
		try{
			var img_AcceptArr = ACCEPT_IMG_TYPES.split(',');
			var selFile = fileObj.value;
			var isAccepted = new Boolean();
			var filetext;
			
			if (selFile.length > 3){
				var flExtension = selFile.substr(selFile.length - 3).toLowerCase();
				
				for (i = 0; i < img_AcceptArr.length; i++){
					if (img_AcceptArr[i].toLowerCase() == flExtension){
						isAccepted = true;
						break;
					}					
				}

				if (isAccepted==true){
					newpath=GetFilePath(selFile);	
					filetext = window.document.getElementById('filetext').value;
						
					window.parent.document.forms[0].elements(filetext).value =newpath; 
					passedhere=1;
					document.forms[0].submit();   
				}
				else {
					ShowMessage('The selected type of file is not accepted.\nEnsure that the file picked out is a valid image file.');
				}	
			}	
		}	
		catch(e){alert("The file could not be uploaded: " + e.description)}
	}
	function refreshPhoto(){
		try{
			var toRefresh = window.parent.document.forms[0].elements("handlePhotoRefresh").value ;
			if (toRefresh=="1") 
			{
			drawPicture();
			alert('here');
			}
		}
		catch(e){}
	}
	
	function drawPicture(){
		var outPutHTML = new String();
		var defHTML = new String();
		var tdHold = document.all.item("PictureTD"); 		
		defHTML = '<img ID="defPhoto" width="170px" height="150px" src="/Data/Photos/_blank.jpg">';
		try{								
			var thePath = window.parent.document.forms[0].elements(filetext).value ;			
			
			if (thePath!==""){				
				var thePathName = thePath;
				outPutHTML = '<img ID="defPhoto" width="170px" height="150px" src="' + thePathName + '">';
			}			
		}
		catch(e){}
		
		if (outPutHTML=="") tdHold.innerHTML = defHTML;
		else	tdHold.innerHTML = outPutHTML;
	}		
	
	function deleteImage(){
		try{
			if (window.parent.confirm('Are you sure you want to delete the current image?')==false){
				return;
			}
			var default_Img = '/Data/Photos/_blank.jpg'
			var currServPhoto = window.parent.document.forms[0].elements(filetext).value;
			if (currServPhoto == default_Img){
				ShowMessage ('No photo selected')
			}
			else {
				//document.forms[0].action = 'upload.asp';    
				//document.forms[0].target = '_self';    
				//document.all.item('delPhoto').value = '1';
				//document.all.item('delPhotoPath').value = currServPhoto;				
				window.parent.document.forms[0].elements(filetext).value = default_Img;
				drawPicture();
				//document.forms[0].submit();    
				
			}
		}		
		catch(e){ShowMessage('The deletion process could not continue.' + e.description);}
	}
</SCRIPT>
</head>
<!--#include virtual="libroutines.asp"-->

<%

delPhoto = Request.Form("delPhoto")
filetext=Request.QueryString("filetext")
%>
<SCRIPT Language="JavaScript">
var filetext='<%=filetext%>';
//var filetext='txtPhoto';
var newpath;
var passedhere;

</Script>
<%
If delPhoto = "1" Then
	delPhotoPath = Request.Form("delPhotoPath")
	
	If delPhotoPath <> "" Then
		Dim fs
		Set fs = CreateObject("Scripting.FileSystemObject")
		fs.DeleteFile delPhotoPath, True
	
		Response.Redirect "Upload.asp"
		Response.End	
	End If	
End If
%>


<Script Language="VBScript">
	Function GetFilePath(photo)
		Dim photoNew
		photoNew = StrReverse(photo)
		If  InStr(1, photoNew, "\") > 0 Then
			photoNew = Mid(photoNew, 1, InStr(1, photoNew, "\") - 1)
		ElseIf 	InStr(1, photoNew, "/") > 0 Then
			photoNew = Mid(photoNew, 1, InStr(1, photoNew, "/") - 1)
		End If	
		photoNew = "/Data/Photos/" & StrReverse(photoNew)
		GetFilePath = photoNew
	End Function	
	
	
</Script>
<body class="Dialog" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" OnLoad="JavaScript: drawPicture();">
	<form method="post" enctype="multipart/form-data" action="savePhoto.asp?filetext=<%=filetext%>" name="frmFile" id="frmFile">
	<input type=hidden name=delPhoto value=0>
	<input type=hidden name=delPhotoPath value=0>
	<input type=hidden name="filetext" value='<%=filetext%>'>
	<table border=0 cellspacing=0 cellpadding=0>
		<tr>
			<td ID="PictureTD">
				
			</td>
		</tr>
		<tr>	
			<td>
				<INPUT class="cal-textboxBlue" value="<%= Request("File") %>" name="File" ID="File" type="File" size="12" OnChange="JavaScript: Upload(this);" TITLE="Upload image">&nbsp;				
				<input type=Button Class=cal-textboxBlue Value=" X " TITLE="Delete current image" OnClick="JavaScript: deleteImage();">
			</td>
		</tr>	
	</table>
	</form>
</body>

</html>