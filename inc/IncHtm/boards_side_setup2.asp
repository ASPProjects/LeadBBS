<%
Function SideBoard_GetContent()
Dim Str,Tmp
Tmp = Topic_AnnounceList(GBL_Board_ID,10,0,"yes","1","0","")
If Tmp <> "" Then Str = Str & "		<div class=""content_side_box"">" & VbCrLf &_
"			<div class=""title""><b>��龫��</b></div>" & VbCrLf &_
"			" & Tmp & VbCrLf &_
"		</div>" & VbCrLf
SideBoard_GetContent = Str
End Function
Const GBL_B_SubBoard_Flag = 0
%>