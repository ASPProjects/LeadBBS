<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Board_popfun.asp -->
<!-- #include file=../../User/inc/UserTopic.asp -->
<%
DEF_BBS_homeUrl="../../"

Main

Sub Main

	BBS_SiteHead DEF_SiteNameString & " - 汉字简体繁体转换",0,"<span class=navigate_string_step>汉字简体繁体转换</span>"
	UserTopicTopInfo("plug")

	If GBL_CHK_User = "" then
		Response.write "<div class=alert>您没有使用简体繁体转换的权限，请先登陆或者注册为论坛会员。</div>"
	Else
		Main_ChineseCode
	End If
	UserTopicBottomInfo
	SiteBottom

End Sub

Sub Main_ChineseCode

%>
	<div class="clear"></div>
	<a href=javascript:; onclick='app_load("汉字简体繁体转换","ChineseCode.htm","500px","400px");'>汉字简繁转换</a>
	- <a href=javascript:; onclick='app_load("万年历","cal/cal.htm","540px","475px");'>万年历</a>
	- <a href=javascript:; onclick='app_load("黄金矿工","../flash_gold/default.asp?appflag=1","580px","1024px");'>黄金矿工</a>
	- <a href=javascript:; onclick='app_load("聊天室","../bbschat/default.asp?appflag=1","500px","900px");'>聊天室(LeadChat)</a>
	<br>
	<br>
		<div id="appTitle" class="apptitle" style="margin-bottom:10px;font-weight:bold;color: blue;font-size:14px;"></div>
	<div class="appmain" style="border:1px #888888 dashed;background:#eeeeee;width:auto;float:left;padding:5px;margin-bottom:35px;">
		<iframe src="ChineseCode.htm" name="appFrame" id="appFrame" hidefocus="" frameborder="no" scrolling="no" style="margin:0px;padding:0px;font-size:12px;overflow-x:hidden;"></iframe>
	</div>
	<script>
	function app_load(title,url,width,height)
	{
		$id("appFrame").style.width = width;
		$id("appFrame").style.height = height;
		$id("appFrame").src = url;
		$id("appTitle").innerHTML = title;
	}
	app_load("汉字简体繁体转换","ChineseCode.htm","500px","400px");
	</script>
		
<%

End Sub%>