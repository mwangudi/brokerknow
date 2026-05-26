<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2//EN">

<html dir=ltr><HEAD><TITLE>Contents</TITLE>
<META NAME="ROBOTS" CONTENT="NOINDEX">
<META HTTP-EQUIV="Content-Type" content="text/html; charset=Windows-1252">
</HEAD>

<BODY bgcolor="deepskyblue">
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

Sub TOCSynch_Click()
MyUrl=parent.frames(2).location
x=InStr(MyUrl,"/iishelp")
y=Len(MyUrl)
NewUrl=FixHex(Right(MyUrl,y-(x-1)))
call hhctrl.syncURL(NewUrl)
call hhctrl.syncURL(NewUrl)
end sub

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

<SPAN STYLE="LEFT:  0px; POSITION: relative; TOP: 4px">
<IMG alt=Contents src="Cont.gif" border=0>
</SPAN>

<SPAN STYLE="LEFT:  -4px; POSITION: relative; TOP: 4px">
<A href="Index.asp"><IMG alt=Index src="NoIndex.gif" border=0></A>
</SPAN>


<SPAN STYLE="LEFT:  -8px; POSITION: relative; TOP: 4px">
<A href="Search.asp"><IMG alt=Search src="NoSearch.gif" border=0></A>
</SPAN>


<table width=262 height="31" border="0" cellspacing=2 bgcolor="white" bordercolor="white">
<TR><TD width="208">

</TD>
<TD align="right">
<SPAN STYLE="LEFT:  1px; POSITION: relative; TOP: 0px">
<A onclick=TOCPrint_Click() href="#Ptoc"><IMG alt="Print a topic or node from the TOC" src ="print.gif" border=0 ></a><a name="Ptoc"></a>
</SPAN>
</TD>
<td align="left">
<A onclick=TOCSynch_Click() href="#Stoc"><IMG alt="Synchronize the TOC with the content pane" src ="synch.gif" border=0 ></a><a name="Stoc"></a>
</td></TR></table>
<OBJECT id=hhctrl codeBase=../../common/alpha.cab#version=4,73,8412,0 
type=application/x-oleobject height="74%" width=262 
classid=clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11><PARAM NAME="Width" VALUE="6932"><PARAM NAME="Height" VALUE="15981"><PARAM NAME="Command" VALUE="Contents"><PARAM NAME="Item1" VALUE="cohhc.hhc"></OBJECT>
</BODY>
</HTML>