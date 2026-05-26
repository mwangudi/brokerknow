<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2//EN">

<html dir=ltr><HEAD><TITLE>Contents</TITLE>
<META NAME="ROBOTS" CONTENT="NOINDEX"></HEAD>

<BODY bgcolor="deepskyblue" border="0"><font face="ms sans serif">
<SCRIPT LANGUAGE="VBScript">
<!--
Function Dec(strHex)
    Dec = InStr("123456789ABCDEF", UCase(Left(strHex,1))) * 16
    Dec = Dec + InStr("123456789ABCDEF", UCase(Mid(strHex,2,1)))
End Function

Function FixHex(ByVal strURL)
    Dim x
    FixHex = ""    
    x = InStr(1,strURL,"%")
    Do While (x > 0)
        FixHex = FixHex & Left(strURL,x-1)
        FixHex = FixHex & Chr(Dec(Mid(strURL,x+1)))
        strURL = Mid(strURL,x+3)
        x = InStr(1,strURL,"%")
    Loop
    FixHex = FixHex & strURL
End Function


Sub TOCPrint_Click()
MyUrl=parent.frames(2).location
x=InStr(MyUrl,"/iishelp")
y=Len(MyUrl)
NewUrl=FixHex(Right(MyUrl,y-(x-1)))
hhctrl.syncURL(NewUrl)
hhctrl.syncURL(NewUrl)
hhctrl.Print()
End Sub
-->
</SCRIPT>

<% MachType=Request.ServerVariables("HTTP_UA-CPU")
If MachType="Alpha" Then
	ContHref="contalph.asp"
Else
	ContHref="Contents.asp"
End If %>

<SPAN STYLE="LEFT:  0px; POSITION: relative; TOP: 4px">
<A HREF="<%= ContHref%>"><IMG alt=Contents src="NoCont.gif" border=0></A>
</SPAN>

<SPAN STYLE="LEFT:  -4px; POSITION: relative; TOP: 4px">
<IMG alt=Index src="Index.gif" border=0>
</SPAN>

<SPAN STYLE="LEFT:  -8px; POSITION: relative; TOP: 4px">

<A href="Search.asp"><IMG alt=Search src="NoSearch.gif" border=0></A>

</SPAN>

<table width="262" bgcolor="#ffffff" border="0" cellpadding="10">
<TR border="1" bgcolor="#ffffff">
<td>
<SPAN STYLE="LEFT:  0px; POSITION: relative; TOP: 18px" 
     >
<font size="-2">Click the text box, then type the word you are looking for. Select an index entry and click <b>Display.</b><br></font>
</SPAN>
</td>
</TR>
</table>

<TABLE bgcolor="#ffffff" width="262" height="69%" font="verdana">
<TR border="0" bgcolor="#ffffff">
<TD valign="top">
<div style="FONT-SIZE: 9pt; FONT-FAMILY: verdana" 
     >
<SPAN STYLE="LEFT:  0px; POSITION: relative; TOP: 10px">
<center>
      <OBJECT id=hhctrl codeBase=../../common/hhctrl.cab 
      type=application/x-oleobject height=220 width=238 
      classid=clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11><PARAM NAME="Width" VALUE="6297"><PARAM NAME="Height" VALUE="5821"><PARAM NAME="Command" VALUE="index"><PARAM NAME="Item1" VALUE="cohhk.hhk"></OBJECT>
</center>
</SPAN>
</div>
</TD>
</TR>

</TABLE></font>

</BODY>
</HTML>