<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Board_popfun.asp -->
<!-- #include file=../../User/inc/UserTopic.asp -->
<%
DEF_BBS_homeUrl="../../"

Main

Sub Main

	BBS_SiteHead DEF_SiteNameString & " - ���ּ��己��ת��",0,"<span class=navigate_string_step>���ּ��己��ת��</span>"
	UserTopicTopInfo("plug")

	If GBL_CHK_User = "" then
		Response.write "<div class=alert>��û��ʹ�ü��己��ת����Ȩ�ޣ����ȵ�½����ע��Ϊ��̳��Ա��</div>"
	Else
		Main_ChineseCode
	End If
	UserTopicBottomInfo
	SiteBottom

End Sub

Sub Main_ChineseCode

%>
	<div class="clear"></div>
	<a href=javascript:; onclick='app_load("���ּ��己��ת��","ChineseCode.htm","500px","400px");'>���ּ�ת��</a>
	- <a href=javascript:; onclick='app_load("������","cal/cal.htm","540px","475px");'>������</a>
	- <a href=javascript:; onclick='app_load("�ƽ��","../flash_gold/default.asp?appflag=1","580px","1024px");'>�ƽ��</a>
	- <a href=javascript:; onclick='app_load("������","../bbschat/default.asp?appflag=1","500px","900px");'>������(LeadChat)</a>
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
	app_load("���ּ��己��ת��","ChineseCode.htm","500px","400px");
	</script>
		
<%

End Sub%>