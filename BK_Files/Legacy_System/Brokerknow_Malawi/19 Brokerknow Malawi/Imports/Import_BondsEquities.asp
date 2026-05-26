<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>CDS File Upload</TITLE>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>
</head>
<!--#include virtual="libroutines.asp"-->

<%

delPhoto = Request.Form("delPhoto")

If delPhoto = "1" Then
	delPhotoPath = Request.Form("delPhotoPath")
	
	If delPhotoPath <> "" Then
		Dim fs
		Set fs = CreateObject("Scripting.FileSystemObject")
		fs.DeleteFile delPhotoPath, True
	
		Response.Redirect "import.asp"
		Response.End	
	End If	
End If
%>

<SCRIPT Language="JavaScript">
	function Upload(fileObj){
		var ACCEPT_TXT_TYPES = 'xls';
		
		try{
			var img_AcceptArr = ACCEPT_TXT_TYPES.split(',');
			var selFile = fileObj.value;
			var isAccepted = new Boolean();
			
			if (selFile.length > 3){
				var flExtension = selFile.substr(selFile.length - 3).toLowerCase();
				
				for (i = 0; i < img_AcceptArr.length; i++){
					if (img_AcceptArr[i].toLowerCase() == flExtension){
						isAccepted = true;
						break;
					}					
				}
			
				if (isAccepted==true){
				   
					document.forms[0].submit();        
				}
				else {
					ShowMessage('The selected type of file is not accepted.\nEnsure that the file picked out is a valid TEXT file.');
				}	
			}	
		}	
		catch(e){alert("The file could not be uploaded: " + e.description)}
	}
	function refreshPhoto(){
		try{
			var toRefresh = window.parent.document.forms[0].elements("handlePhotoRefresh").value ;
		}
		catch(e){}
	}
	
	
	
	function ShowUnprocessedImports(){
		window.location.replace("CDSMatchedTradesList.asp");
	}
	
	function ShowUnmatchedTrades(){
		window.location.replace("CDSUnMatchedTradesList.asp");
	}

	function ImportPriceList(){
		window.location.replace("ImportPriceList.asp");
	}
	
	function UpdateCommissions(){
		window.location.replace("UpdateCommissions.asp");
	}
	
</SCRIPT>
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

<body class="Reports" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
	<form method="post" enctype="multipart/form-data"  action="Import_BondsEquities2.asp" name="frmFile">
	<input type=hidden name=delPhoto value=0>
	<input type=hidden name=delPhotoPath value=0>
	<table border=0 cellspacing=5 cellpadding=5>
	
	
		<tr>
			<td>
				<input type="radio" class="BorderLess" value="1" id="dTrade" name="CDSMode">&nbsp;&nbsp;<label for="dTrade" style="cursor: hand">Daily Trades</label>
				<br><input type="radio" class="BorderLess" value="1" id="dTrade2" name="CDSMode"  onclick ="javascript: window.location.replace('importHoldings.asp');">&nbsp;&nbsp;<label for="dTrade2" style="cursor: hand" >Client Holdings<label>
				<br><input type="radio" class="BorderLess" value="1" id="dTrade3" name="CDSMode">&nbsp;&nbsp;<label for="dTrade3" style="cursor: hand">Transaction Details</label>
				<br><input type="radio" class="BorderLess" value="1" id="price" name="CDSMode"  onclick ="javascript: window.location.replace('importPriceList.asp');">&nbsp;&nbsp;<label for="price" style="cursor: hand">Import Price List</label>
				<br><input type="radio" checked class="BorderLess" value="1" id="turnover" name="CDSMode"  onclick ="javascript: window.location.replace('Import_BondsEquities.asp');">&nbsp;&nbsp;<label for="turnover" style="cursor: hand">Import Bonds & Equities Turnovers</label>
				
			</td>
		</tr>
		<tr>	
			<td>
				<INPUT class="cal-textboxBlue" value="<%= Request("File") %>" name="sourcefile" ID="sourcefile" type="File" size="30" OnChange="JavaScript: Upload(this);" TITLE="Upload image">&nbsp;<input type=Button Class=cal-textboxBlue Value=" X " TITLE="Delete current image" STYLE="display: none" OnClick="JavaScript: deleteImage();" id=Button1 name=Button1>
			</td>
		</tr>	
		<!--<tr>	
			<td>
				<INPUT type=Button  value="View Imported" name="showImported" ID="showImported" TITLE="Show unprocessed imported trades" OnClick="JavaScript: ShowUnprocessedImports();">
			</td>
		</tr>
		<tr>	
			<td>
				<INPUT type=Button  value="View Unmatched" name="showUnmatched" ID="showUnmatched" TITLE="Show unmatched trades" OnClick="JavaScript: ShowUnmatchedTrades();">
			</td>
		</tr>
		
		<tr>	
			<td>
				<INPUT type=Button  value="Update Commissions For Compounded Contracts" name="updateCommissions" ID="updateCommissions" TITLE="Update commissions for compounded contracts" OnClick="JavaScript: UpdateCommissions();">
			</td>
		</tr>-->
		
		<!-- <tr>	
			<td>
				<INPUT type=Button  value="Import Price list" name="ImportPriceLsit" ID="ImportPriceLsit" TITLE="Import NSE PriceList" OnClick="JavaScript: ImportPriceList();">
			</td>
		</tr> -->
	</table>
	</form>
</body>

</html>