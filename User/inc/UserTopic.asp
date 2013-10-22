<%
Const LMT_MaxMessageNumber = 5000 '用户收件箱允许的最多接收记录，超过将无法接收新消息。

Sub DisplayUserCenter(info)

	%>
	<script language="JavaScript" type="text/javascript">
	function swap_view(str,sobj)
	{
		var obj=$id(str);
		obj.style.display=(obj.style.display=='none'?'':'none');
		sobj.className=(sobj.className=='swap_collapse'?'swap_open':'swap_collapse');
	}
	</script>
	<%
	If info = "user" Then
	%>
			<div class="title">个人专区</div>
			<div class="user_itemlist">
			<div class="swap_collapse" onclick="swap_view('master_part_1',this);"><span>个人信息</span></div>
			<ul id="master_part_1">
			<%If GetBinarybit(GBL_CHK_UserLimit,1) = 1 or GBL_CHK_UserLimit = "" Then%><li><a href=<%=DEF_BBS_HomeUrl%>User/UserGetPass.asp?act=active><span class=redfont>激活我的账号</span></a></li><%End If%>
			<li><a href=<%=DEF_BBS_HomeUrl%>User/UserModify.asp>修改我的资料</a></li>
			<li><a href=<%=DEF_BBS_HomeUrl%>User/LookUserInfo.asp>个人信息</a></li>
			<li><a href=<%=DEF_BBS_HomeUrl%>User/MyInfoBox.asp>短消息</a></li>
			<li><a href=<%=DEF_BBS_HomeUrl%>User/LookUserInfo.asp?Evol=f>我的好友</a></li>
			<li><a href=<%=DEF_BBS_HomeUrl%>User/LookUserInfo.asp?Evol=bind>绑定网站</a></li>
			</ul>
			<div class="swap_collapse" onclick="swap_view('master_part_2',this);"><span>帖子与附件</span></div>
			<ul id="master_part_2">
			<li><a href=<%=DEF_BBS_HomeUrl%>User/LookUserInfo.asp?Evol=n>我的帖子</a></li>
			<li><a href=<%=DEF_BBS_HomeUrl%>User/LookUserInfo.asp?Evol=l>已上传的附件</a></li>
			<li><a href=<%=DEF_BBS_HomeUrl%>User/LookUserInfo.asp?Evol=bag>收藏的帖子</a></li>
			</ul>
			<div class="swap_collapse" onclick="swap_view('master_part_3',this);"><span>充值</span></div>
			<ul id="master_part_3">
			<li><a href="<%=DEF_BBS_HomeUrl%>User/alipay/Payment.asp"><div class=ttt><%=DEF_PointsName(1)%>充值</div></A></li>
			</ul>
			</div>
	<%
	ElseIf info = "forum" Then
	%>
			<div class=title>论坛信息</div>
			<div class=user_itemlist>
			<ul>
			<li><a href=<%=DEF_BBS_HomeUrl%>User/UserTop.asp>用户排行榜</a></li>
			<li><a href=<%=DEF_BBS_HomeUrl%>User/UserOnline.asp>在线用户</a></li>
			<li><a href=<%=DEF_BBS_HomeUrl%>User/UserTop.asp?r>查找论坛用户</a></li>
			<li><a href=<%=DEF_BBS_HomeUrl%>User/UserTop.asp?e>新入用户</a></li>
			</ul>
			<hr class=splitline2>
			<ul>
			<li><a href=<%=DEF_BBS_HomeUrl%>User/UserTop.asp?b>版面排行</a></li>
			<li><a href=<%=DEF_BBS_HomeUrl%>Search/List.asp?1>论坛帖子</a></li>
			<li><a href=<%=DEF_BBS_HomeUrl%>Search/Search.asp>帖子搜索</a></li>
			<li><a href=<%=DEF_BBS_HomeUrl%>Search/UploadList.asp>论坛附件</a></li>
			</ul>
			<hr class=splitline2>
			<ul>
			<li><a href=<%=DEF_BBS_HomeUrl%>User/help/about.asp>论坛管理团队</a></li>
			</ul>
			</div>
	<%
	ElseIf info = "help" Then
	%>
			<div class=title>帮助中心</div>
			<div class=user_itemlist>
			<ul>
			<li><a href=<%=DEF_BBS_HomeUrl%>User/Help/Help.asp>使用手册</a></li>
			<li><a href=<%=DEF_BBS_HomeUrl%>User/Help/Ubb.asp>UBB代码</a></li>
			<li><a href=<%=DEF_BBS_HomeUrl%>User/Help/Ubb.asp?icon>论坛表情</a></li>
			</ul>
			<hr class=splitline2>
			<ul>
			<li><a href=<%=DEF_BBS_HomeUrl%>User/help/Ubb.asp?colo>颜色对照表</a></li>
			</ul>
			</div>
<%
	ElseIf info = "plug" Then
	%>
			<div class=title>插件/工具</div>
			<div class=user_itemlist>
			<ul>
			<li><a href=<%=DEF_BBS_HomeUrl%>plug-ins/ChineseCode/default.asp>论坛插件</a></li>
			</ul>
			<hr class=splitline2>
			<ul>
			<li><a href=<%=DEF_BBS_HomeUrl%>plug-ins/LeadCard/Default.asp>LeadCard</a></li>
			</ul>
			</div>
<%
	End If

End Sub

Function DisplayLoginForm(title)

Dim AjaxFlag
If Request("AjaxFlag") = "1" Then
	AjaxFlag = 1
Else
	AjaxFlag = 0
End If

Dim Url
Url = filterUrlstr(Left(Request("dir"),100))
If Url = "" and (inStr(Request.QueryString,"dir=") = 0) and (inStr(Request.form,"dir=") = 0) Then
	Url = DEF_BBS_HomeUrl
End If

Dim action,command
action = Left(Request("action"),5)
command = Left(request("command"),5)

%>
<div class="title" id="login_title"><%=title%></div>
<form action=<%=Url%>User/<%If action = "bind" and command = "bind" Then
		Response.Write DEF_RegisterFile
	Else
		Response.Write "login.asp"
	End If%> method="post" id="login_form" onsubmit="submit_disable(this);"<%
	If AjaxFlag = 1 Then
		Response.Write " target=""hidden_frame"""
	End If
	%>>
	<div class=value2><span class=a>账号：</span><input name=user tabindex=91 type=text maxlength=20 size=22 value="<%
	If action = "bind" and command = "bind" Then
	Else
		If GBL_CHK_user = "" or isNull(GBL_CHK_user) Then
			Response.Write htmlencode(Request("user"))
		Else
			Response.Write htmlencode(GBL_CHK_user)
		End If
	End If%>" class='fminpt input_2'> <a href=<%=Url%>User/<%=DEF_RegisterFile%>>注册</a>
	<a href=<%=Url%>User/UserGetPass.asp?act=active><span class=redfont>激活</span></a>
	<input type=hidden value="<%
	'If Request("submitflag") <> "ddddls-+++" Then
		If Request("u") <> "" Then
			Response.Write htmlencode(Request("u"))
		Else
			Dim HomeUrl,u
			HomeUrl = "http://"&Request.ServerVariables("server_name")
			u = filterUrlstr(Request.QueryString("u"))
			If Left(u,1) <> "/" and Left(u,1) <> "\" and Left(u,Len(HomeUrl)) <> HomeUrl Then u = ""
			If u = "" Then
				u = Lcase(Request.ServerVariables("HTTP_REFERER"))
				If Request.ServerVariables("SERVER_PORT") <> "80" Then HomeUrl = HomeUrl & ":" & Request.ServerVariables("SERVER_PORT")
				If Left(u,Len(HomeUrl)) <> Lcase(HomeUrl) Then u = ""
				If inStr(u,"/user/login.asp") > 0 Then u = ""
			End If
			Response.Write htmlencode(u)
		End If
	'End If%>" name=u></div>
	<input type=hidden name=AjaxFlag value="<%=htmlencode(Left(Request("AjaxFlag"),1))%>">
	<input type=hidden name=JsFlag value="1">
	<input type=hidden name=action value="<%=htmlencode(action)%>">
	<input type=hidden name=command value="<%=htmlencode(command)%>">
	<div class=value2><span class=a>密码：</span><input name=pass tabindex=92 type=password maxlength=20 size=22 value="<%'=htmlencode(GBL_CHK_pass)%>" class='fminpt input_2'>
	<a href=<%=Url%>User/UserGetPass.asp>忘记密码？</a>
	</div>
	<div class=value2><span class=a>保存：</span><select name=CkiExp>
			<option value="-99">安全模式
			<option value="-1">浏览进程
			<option value=7 selected>一周
			<option value="3650">永久
		</select>密码保留时间
	</div>
	<br />
	<div class=value2>
	<input name=submitflag type=hidden value="ddddls-+++">
	<input type=submit value="登录" class="fmbtn btn_2">
	</div>
</form>
	<br />
	<div class=value2>注意：选择安全模式，将不会在本地存储账户信息</div><%
If GetBinarybit(DEF_Sideparameter,10) = 1 Then%>
<span class="grayfont">其它登录：</span><a href="<%=Url%>app/qqlogin/login.asp"><img src="<%=Url%>images/app/1.gif" border="0" style="position:absolute;" /><span style="padding-left:18px;">QQ登录</span></a><%
End If%></div>
<%
End Function

Sub UserTopicTopInfo(info)
%>
<div class="area"><%
	Global_TableHead
%>
<div class="main user_table">
	<%If info <> "" Then%>
	<div class="content_side_left tdleft" id="p_side"><%DisplayUserCenter(info)%>
	</div><%End If%>
	<div class="content_main_left">
		<div class="content_main_2_left">
		<div class="content_main_body tdright">
			<div class="tdright_collapse">

<%End Sub

Sub UserTopicBottomInfo

%>				</div>
			</div>
	</div>
	</div>
</div>
</div><%Global_TableBottom%></div><%

End Sub

Sub Processor_LoginMsg(str,obj,evl)

	If AjaxFlag = 0 Then
		Response.Write str
	Else
		If AjaxFlag = 1 and Request.Form("JsFlag")="1" Then%>
		<script>parent.layer_outmsg("<%=obj%>","<span class=\"redfont\"><%=Replace(Replace(Replace(Str,"\","\\"),"""","\"""),VbCrLf,"\n")%></span>","","<%=Replace(Replace(Replace(evl,"\","\\"),"""","\"""),VbCrLf,"\n")%>");</script>
		<%
		Else%>
		<span class="redfont">
			<%=Str%>
		</span>
	<%	End If
	End If

End Sub%>