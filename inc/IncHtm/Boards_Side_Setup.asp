<%
Str = Str & "		<div class=""content_side_box"">" & VbCrLf &_
"			<div class=""title""><b>最新精华</b></div>" & VbCrLf &_
"			" & Topic_AnnounceList(0,10,0,"yes","1","0","") & VbCrLf &_
"		</div>" & VbCrLf
Str = Str & "		<div class=""content_side_box"">" & VbCrLf &_
"			<div class=""title""><b>最新图片</b></div>" & VbCrLf &_
"			" & Topic_PicInfo(140,105,4) & VbCrLf &_
"		</div>" & VbCrLf
Str = Str & "		<div class=""content_side_box"">" & VbCrLf &_
"			<div class=""title""><b>版块排行</b></div>" & VbCrLf &_
"			" & Topic_AnnounceList(0,10,0,"yes","2","0","") & VbCrLf &_
"		</div>" & VbCrLf
Str = Str & "		<div class=""content_side_box"">" & VbCrLf &_
"			<div class=""title""><b>论坛推荐</b></div>" & VbCrLf &_
"			" & Topic_AnnounceList(0,16,54,"yes","0","0","") & VbCrLf &_
"		</div>" & VbCrLf
Str = Str & "		<div class=""content_side_box"">" & VbCrLf &_
"			<div class=""title""><b>最新帖子</b></div>" & VbCrLf &_
"			" & Topic_AnnounceList(0,16,0,"yes","0","0","") & VbCrLf &_
"		</div>" & VbCrLf
%>