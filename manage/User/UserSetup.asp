<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/User_setup.asp -->
<!-- #include file=../../inc/Board_popfun.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<%
DEF_BBS_HomeUrl = "../../"
Const MaxLinkNum = 100
initDatabase
GBL_CHK_TempStr = ""
checkSupervisorPass

Dim Form_DEF_UserEnableUserTitle,Form_DEF_UserUserTitleNeedLevel,Form_LMT_UserNameEnableEnglishWords
Dim Form_LMT_UserNameEnableChineseChar,Form_LMT_UserNameEnableChineseWords
Dim Form_DEF_User_RegPoints,Form_LMT_EnableRegNewUsers,Form_DEF_ShortestUserName,Form_DEF_RegNewUserTotalRestTime
Dim Form_DEF_UserNewRegAttestMode,Form_DEF_UserActivationExpiresDay,Form_DEF_User_GetPassMode
Dim Form_DEF_UserLevelPoints,Form_DEF_UserLevelString,Form_DEF_UserOfficerString
Redim Form_DEF_UserLevelPoints(DEF_UserLevelNum),Form_DEF_UserLevelString(DEF_UserLevelNum),Form_DEF_UserOfficerString(DEF_UserOfficerNum)
Dim Form_DEF_FiltrateUserNameString,Form_DEF_UserShortestPassword,Form_DEF_UserShortestPasswordMaster,Form_Def_UserTestNumber
Dim Form_DEF_seller_email,Form_DEF_seller_minpoints,Form_DEF_seller_exchangescale

GetDefaultValue

Manage_sitehead DEF_SiteNameString & " - 管理员",""
frame_TopInfo
DisplayUserNavigate("用户注册参数设置")
If GBL_CHK_Flag=1 Then
	UserSetup
Else
DisplayLoginForm
End If
frame_BottomInfo
closeDataBase
Manage_Sitebottom("none")

Function UserSetup

%>
<form name="pollform3sdx" method="post" action="UserSetup.asp">
<input type="hidden" name="SubmitFlag" value=yes>
<p>
		设置：<a href=../SiteManage/SiteSetup.asp>论坛常用参数</a> <a href=../SiteManage/UploadSetup.asp>上传参数</a>
		<span class=grayfont>用户注册参数</span>
		<a href=../SiteManage/UbbcodeSetup.asp>UBB编码参数</a>
		<br><span class=grayfont>(下面是您网站的用户注册参数，错误的设置将会发生严重错误)<br><br>
		如果在设置后发现网站不能正常运行，请将LeadBBS最新版的User_Setup.asp覆盖回去</span>
</p>
<%If Request.Form("SubmitFlag") <> "" Then
	CheckLinkValue
End If%>
<b><span class=redfont><%=GBL_CHK_TempStr%></span></b>
<p>
<%
If Request.Form("SubmitFlag") <> "" Then
	If GBL_CHK_TempStr <> "" Then
		DisplayDatabaseLink
	Else
		MakeDataBaseLinkFile
		Exit Function
	End If
Else
	DisplayDatabaseLink
End If
%>
<br>
<input type=submit name=提交 value=提交 class=fmbtn>
<input type=reset name=取消 value=取消 class=fmbtn>
</form>
<%

End Function

Function CheckLinkValue

	GetFormValue

End Function

Function DisplayDatabaseLink

	Dim N
		%>
		<table border=0 cellpadding=0 cellspacing=0 width="100%" class=frame_table>
		<tr>
			<td class=tdbox width=120>自定头衔</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_UserEnableUserTitle value=0<%If Form_DEF_UserEnableUserTitle = 0 Then%> checked<%End If%>></td><td>禁止</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_UserEnableUserTitle value=1<%If Form_DEF_UserEnableUserTitle = 1 Then%> checked<%End If%>></td><td>允许 (<span class=grayfont>是否允许用户自定义头衔</span>)</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>头衔<%=DEF_PointsName(3)%></td>
			<td class=tdbox>
				<select name=Form_DEF_UserUserTitleNeedLevel><%
				For N = 0 to DEF_UserLevelNum
					If N = Form_DEF_UserUserTitleNeedLevel Then
						Response.write "				<option value=" & N & " selected>" & N & "." & DEF_UserLevelString(N) & "</option>" & VbCrLf
					Else
						Response.write "				<option value=" & N & ">" & N & "." & DEF_UserLevelString(N) & "</option>" & VbCrLf
					End If
				Next%>(<span class=grayfont>用户自定义头衔所需要的<%=DEF_PointsName(3)%></span>)
				</select> 如允许自定义头衔，请指定自定义头衔所要求达到的<%=DEF_PointsName(3)%></td>
		</tr>
		<tr>
			<td class=tdbox>用 户 名</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_LMT_UserNameEnableEnglishWords value=0<%If Form_LMT_UserNameEnableEnglishWords = 0 Then%> checked<%End If%>></td><td>禁止使用西文字符(字母数字)</td>
          		<td><input class=fmchkbox type=radio name=Form_LMT_UserNameEnableEnglishWords value=1<%If Form_LMT_UserNameEnableEnglishWords = 1 Then%> checked<%End If%>></td><td>允许使用西文字符(字母数字)</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>用 户 名</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_LMT_UserNameEnableChineseChar value=0<%If Form_LMT_UserNameEnableChineseChar = 0 Then%> checked<%End If%>></td><td>禁止使用中文符号(标点,日文等字符)</td>
          		<td><input class=fmchkbox type=radio name=Form_LMT_UserNameEnableChineseChar value=1<%If Form_LMT_UserNameEnableChineseChar = 1 Then%> checked<%End If%>></td><td>允许使用中文符号(标点,日文等字符)</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>用 户 名</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_LMT_UserNameEnableChineseWords value=0<%If Form_LMT_UserNameEnableChineseWords = 0 Then%> checked<%End If%>></td><td>禁止使用中文汉字</td>
          		<td><input class=fmchkbox type=radio name=Form_LMT_UserNameEnableChineseWords value=1<%If Form_LMT_UserNameEnableChineseWords = 1 Then%> checked<%End If%>></td><td>允许使用中文汉字</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>注册<%=DEF_PointsName(0)%></td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_User_RegPoints" maxlength="4" size="10" value="<%=htmlencode(Form_DEF_User_RegPoints)%>"><span class=grayfont>(刚注册用户就拥有的<%=DEF_PointsName(0)%>点数，默认为0)</span></td>
		</tr>
		<tr>
			<td class=tdbox>开关注册</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_LMT_EnableRegNewUsers value=0<%If Form_LMT_EnableRegNewUsers = 0 Then%> checked<%End If%>></td><td>禁止注册新用户</td>
          		<td><input class=fmchkbox type=radio name=Form_LMT_EnableRegNewUsers value=1<%If Form_LMT_EnableRegNewUsers = 1 Then%> checked<%End If%>></td><td>允许新用户注册</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>用户名长</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_ShortestUserName" maxlength="2" size="10" value="<%=htmlencode(Form_DEF_ShortestUserName)%>"><span class=grayfont>(允许注册的用户名的最短字符个数，单位字节)</span></td>
		</tr>
		<tr>
			<td class=tdbox>最短密码</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_UserShortestPassword" maxlength="2" size="10" value="<%=htmlencode(Form_DEF_UserShortestPassword)%>"><span class=grayfont>(允许使用的用户密码的最短字符个数，单位字节，针对普通用户)</span></td>
		</tr>
		<tr>
			<td class=tdbox>管理密码</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_UserShortestPasswordMaster" maxlength="2" size="10" value="<%=htmlencode(Form_DEF_UserShortestPasswordMaster)%>"><span class=grayfont>(允许使用的<%=DEF_PointsName(8)%>以上成员的最短密码字符个数)</span></td>
		</tr>
		<tr>
			<td class=tdbox>注册间隔</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_RegNewUserTotalRestTime" maxlength="6" size="10" value="<%=htmlencode(Form_DEF_RegNewUserTotalRestTime)%>"><span class=grayfont>(限制论坛在此时间内只允许注册一名新用户，单位秒)</span></td>
		</tr>
		<tr>
			<td class=tdbox>注册认证</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_UserNewRegAttestMode value=0<%If Form_DEF_UserNewRegAttestMode = 0 Then%> checked<%End If%>></td><td>无激活,注册即为正式会员</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_UserNewRegAttestMode value=1<%If Form_DEF_UserNewRegAttestMode = 1 Then%> checked<%End If%>></td><td>邮件激活(此项务必开启邮件发送功能)</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_UserNewRegAttestMode value=2<%If Form_DEF_UserNewRegAttestMode = 2 Then%> checked<%End If%>></td><td>其它激活(管理员后台更改)</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>激活时间</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_UserActivationExpiresDay" maxlength="3" size="10" value="<%=htmlencode(Form_DEF_UserActivationExpiresDay)%>"><span class=grayfont>(注册后，用户必须在指定天数内激活，否则系统将作删除用户处理，单位：天，填写0表示无限制，永久保留)</span></td>
		</tr>
		<tr>
			<td class=tdbox>密码找回</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_User_GetPassMode value=0<%If Form_DEF_User_GetPassMode = 0 Then%> checked<%End If%>></td><td>禁止找回</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_User_GetPassMode value=1<%If Form_DEF_User_GetPassMode = 1 Then%> checked<%End If%>></td><td>即时密码更改</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_User_GetPassMode value=2<%If Form_DEF_User_GetPassMode = 2 Then%> checked<%End If%>></td><td>即时密码更改 若未激活同时发送邮件重新通知激活码</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox><%=DEF_PointsName(3)%>定义</td>
			<td class=tdbox>
				<table border=0 cellpadding=1 cellspacing=0>
				<tr>
					<td>&nbsp;<%=DEF_PointsName(3)%></td>
					<td>&nbsp;称号</td>
					<td>&nbsp;要求发表文章</td>
				</td><%
			For n = 0 to DEF_UserLevelNum
				%>
				<tr>
					<td>&nbsp;<%=Right(" " & N,2)%>级</td>
					<td>&nbsp;<input class=fminpt type="text" name="Form_DEF_UserLevelString<%=N%>" maxlength="18" size="20" value="<%=htmlencode(Form_DEF_UserLevelString(n))%>"></td>
					<td>&nbsp;<input class=fminpt type="text" name="Form_DEF_UserLevelPoints<%=N%>" maxlength="14" size="10" value="<%=htmlencode(Form_DEF_UserLevelPoints(n))%>"></td>
				</td>
				<%
			Next
			%>
				</table>
				&nbsp;<span class=grayfont>发表文章指的是曾经发表过的文章(包括删除数量)</span></td>
		</tr>
		<tr>
			<td class=tdbox><%=DEF_PointsName(9)%></td>
			<td class=tdbox>
				<table border=0 cellpadding=1 cellspacing=0>
				<tr>
					<td>&nbsp;编号</td>
					<td>&nbsp;称呼</td>
				</tr><%
			For n = 0 to DEF_UserOfficerNum
				%>
				<tr>
					<td>&nbsp;=<%=Right(" " & N,2)%></td>
					<td>&nbsp;<input class=fminpt type="text" name="Form_DEF_UserOfficerString<%=N%>" maxlength="100" size="50" value="<%=htmlencode(Form_DEF_UserOfficerString(n))%>"></td>
				</tr>
				<%
			Next
			%>
				</table>
				&nbsp;<span class=grayfont><%=DEF_PointsName(9)%>名称允许使用html代码，但不允许使用引号．</span></td>
		</tr>
		<tr>
			<td class=tdbox>过滤名字</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_FiltrateUserNameString" maxlength="1024" size="50" value="<%=htmlencode(Form_DEF_FiltrateUserNameString)%>">
			<br><span class=grayfont>(使用|分隔，注册用户名及用户头衔和签名将不能包含此类名字)</span></td>
		</tr>
		<tr>
			<td class=tdbox>验 证 码</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_Def_UserTestNumber value=0<%If Form_Def_UserTestNumber = 0 Then%> checked<%End If%>></td><td>论坛系统默认(继承论坛参数设置)</td><tr>
          		<tr><td><input class=fmchkbox type=radio name=Form_Def_UserTestNumber value=1<%If Form_Def_UserTestNumber = 1 Then%> checked<%End If%>></td><td>肯定使用注册验证码</td></tr>
          		<tr><td><input class=fmchkbox type=radio name=Form_Def_UserTestNumber value=2<%If Form_Def_UserTestNumber = 2 Then%> checked<%End If%>></td><td>肯定不使用注册验证码</td></tr>
          		</table></td>
		</tr>
		<tr>
			<td class=tdbox>支付宝账号</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_seller_email" maxlength="150" size="50" value="<%=htmlencode(Form_DEF_seller_email)%>">
			<br><span class=grayfont>填写网站充值<%=DEF_PointsName(1)%>入账账号,一般是EMAIL地址,不填写表示不开启支付功能</span></td>
		</tr>
		<tr>
			<td class=tdbox>支付宝一次最少充值</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_seller_minpoints" maxlength="6" size="10" value="<%=htmlencode(Form_DEF_seller_minpoints)%>"><span class=grayfont>(单位,RMB元)</span></td>
		</tr>
		<tr>
			<td class=tdbox>支付宝元兑换率</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_seller_exchangescale" maxlength="6" size="10" value="<%=htmlencode(Form_DEF_seller_exchangescale)%>"><span class=grayfont>(一元RMB兑换多少<%=DEF_PointsName(1)%>)</span></td>
		</tr>
		</table>
		<%

End Function

Function GetDefaultValue

	Dim N
	Form_DEF_UserEnableUserTitle = DEF_UserEnableUserTitle
	Form_DEF_UserUserTitleNeedLevel = DEF_UserUserTitleNeedLevel
	Form_LMT_UserNameEnableEnglishWords = LMT_UserNameEnableEnglishWords
	Form_LMT_UserNameEnableChineseChar = LMT_UserNameEnableChineseChar
	Form_LMT_UserNameEnableChineseWords = LMT_UserNameEnableChineseWords
	Form_DEF_User_RegPoints = DEF_User_RegPoints
	Form_LMT_EnableRegNewUsers = LMT_EnableRegNewUsers
	Form_DEF_ShortestUserName = DEF_ShortestUserName
	Form_DEF_RegNewUserTotalRestTime = DEF_RegNewUserTotalRestTime
	Form_DEF_UserNewRegAttestMode = DEF_UserNewRegAttestMode
	Form_DEF_UserActivationExpiresDay = DEF_UserActivationExpiresDay
	Form_DEF_User_GetPassMode = DEF_User_GetPassMode
	For n = 0 to DEF_UserLevelNum
		Form_DEF_UserLevelString(n) = DEF_UserLevelString(n)
		Form_DEF_UserLevelPoints(n) = DEF_UserLevelPoints(n)
	Next
	For n = 0 to DEF_UserOfficerNum
		Form_DEF_UserOfficerString(n) = DEF_UserOfficerString(n)
	Next
	Form_DEF_FiltrateUserNameString = DEF_FiltrateUserNameString
	Form_DEF_UserShortestPassword = DEF_UserShortestPassword
	Form_DEF_UserShortestPasswordMaster = DEF_UserShortestPasswordMaster
	Form_Def_UserTestNumber = Def_UserTestNumber
	Form_DEF_seller_email = DEF_seller_email
	Form_DEF_seller_minpoints = DEF_seller_minpoints
	Form_DEF_seller_exchangescale = DEF_seller_exchangescale

End Function


Function GetFormValue

	Dim N
	Form_DEF_UserEnableUserTitle = Trim(Request.Form("Form_DEF_UserEnableUserTitle"))
	Form_DEF_UserUserTitleNeedLevel = Trim(Request.Form("Form_DEF_UserUserTitleNeedLevel"))
	Form_LMT_UserNameEnableEnglishWords = Trim(Request.Form("Form_LMT_UserNameEnableEnglishWords"))
	Form_LMT_UserNameEnableChineseChar = Trim(Request.Form("Form_LMT_UserNameEnableChineseChar"))
	Form_LMT_UserNameEnableChineseWords = Trim(Request.Form("Form_LMT_UserNameEnableChineseWords"))
	Form_DEF_User_RegPoints = Trim(Request.Form("Form_DEF_User_RegPoints"))
	Form_LMT_EnableRegNewUsers = Trim(Request.Form("Form_LMT_EnableRegNewUsers"))
	Form_DEF_ShortestUserName = Trim(Request.Form("Form_DEF_ShortestUserName"))
	Form_DEF_RegNewUserTotalRestTime = Trim(Request.Form("Form_DEF_RegNewUserTotalRestTime"))
	Form_DEF_UserNewRegAttestMode = Trim(Request.Form("Form_DEF_UserNewRegAttestMode"))
	Form_DEF_UserActivationExpiresDay = Trim(Request.Form("Form_DEF_UserActivationExpiresDay"))
	Form_DEF_User_GetPassMode = Trim(Request.Form("Form_DEF_User_GetPassMode"))
	For n = 0 to DEF_UserLevelNum
		Form_DEF_UserLevelString(n) = Trim(Request.Form("Form_DEF_UserLevelString" & N))
		Form_DEF_UserLevelPoints(n) = Trim(Request.Form("Form_DEF_UserLevelPoints" & N))
	Next
	For n = 0 to DEF_UserOfficerNum
		Form_DEF_UserOfficerString(n) = Trim(Request.Form("Form_DEF_UserOfficerString" & N))
	Next
	Form_DEF_FiltrateUserNameString = Trim(Request.Form("Form_DEF_FiltrateUserNameString"))
	Form_DEF_UserShortestPassword = Trim(Request.Form("Form_DEF_UserShortestPassword"))
	Form_DEF_UserShortestPasswordMaster = Trim(Request.Form("Form_DEF_UserShortestPasswordMaster"))
	Form_Def_UserTestNumber = Trim(Request.Form("Form_Def_UserTestNumber"))
	
	Form_DEF_seller_email = Left(Trim(Request.Form("Form_DEF_seller_email")),150)
	Form_DEF_seller_minpoints = Trim(Request.Form("Form_DEF_seller_minpoints"))
	Form_DEF_seller_exchangescale = Trim(Request.Form("Form_DEF_seller_exchangescale"))
	
	If isNumeric(Form_DEF_UserEnableUserTitle) = 0 Then GBL_CHK_TempStr = "是否允许用户自定义头衔必须为数字<br>" & VbCrLf
	If isNumeric(Form_DEF_UserUserTitleNeedLevel) = 0 Then GBL_CHK_TempStr = "指定自定义头衔所要求达到的" & DEF_PointsName(3) & "必须为数字<br>" & VbCrLf
	If isNumeric(Form_LMT_UserNameEnableEnglishWords) = 0 Then GBL_CHK_TempStr = "是否允许用户名使用西文字符(字母数字)必须为数字<br>" & VbCrLf
	If isNumeric(Form_LMT_UserNameEnableChineseChar) = 0 Then GBL_CHK_TempStr = "是否允许用户使用中文符号(标点,日文等字符)必须为数字<br>" & VbCrLf
	If isNumeric(Form_LMT_UserNameEnableChineseWords) = 0 Then GBL_CHK_TempStr = "是否允许用户使用中文汉字必须为数字<br>" & VbCrLf
	If isNumeric(Form_DEF_User_RegPoints) = 0 Then GBL_CHK_TempStr = "注册用户就拥有的" & DEF_PointsName(0) & "点数必须为数字<br>" & VbCrLf
	If isNumeric(Form_LMT_EnableRegNewUsers) = 0 Then GBL_CHK_TempStr = "开关注册必须为数字<br>" & VbCrLf
	If isNumeric(Form_DEF_ShortestUserName) = 0 Then GBL_CHK_TempStr = "户名的最短字符个数必须为数字<br>" & VbCrLf
	If isNumeric(Form_DEF_RegNewUserTotalRestTime) = 0 Then GBL_CHK_TempStr = "注册间隔必须为数字<br>" & VbCrLf
	If isNumeric(Form_DEF_UserNewRegAttestMode) = 0 Then GBL_CHK_TempStr = "注册认证方式必须为数字<br>" & VbCrLf
	If isNumeric(Form_DEF_UserActivationExpiresDay) = 0 Then GBL_CHK_TempStr = "激活时间必须为数字<br>" & VbCrLf
	If isNumeric(Form_DEF_User_GetPassMode) = 0 Then GBL_CHK_TempStr = "密码找回必须为数字<br>" & VbCrLf
	For n = 0 to DEF_UserLevelNum
		If inStr(Form_DEF_UserLevelString(n),"%") Then
			GBL_CHK_TempStr = "第" & N & DEF_PointsName(3) & "名称不能包含百分号<br>" & VbCrLf
		End If
		If isNumeric(Form_DEF_UserLevelPoints(n)) = 0 Then
			GBL_CHK_TempStr = "第" & N & DEF_PointsName(3) & "要求发表文章数量必须为数字<br>" & VbCrLf
			Exit Function
		End If
	Next
	For n = 0 to DEF_UserOfficerNum
		If inStr(Form_DEF_UserOfficerString(n),"%") Then
			GBL_CHK_TempStr = "第" & N & "编号" & DEF_PointsName(9) & "不能包含百分号<br>" & VbCrLf
		End If
	Next
	If inStr(Form_DEF_FiltrateUserNameString,"%") Then GBL_CHK_TempStr = "密码找回不能包含百分号<br>" & VbCrLf
	If isNumeric(Form_DEF_UserShortestPassword) = 0 Then GBL_CHK_TempStr = "户名密码的最短字符个数必须为数字<br>" & VbCrLf
	If isNumeric(Form_DEF_UserShortestPasswordMaster) = 0 Then GBL_CHK_TempStr = "论坛管理成员户名密码的最短字符个数必须为数字<br>" & VbCrLf
	If inStr(Form_DEF_seller_email,"%") Then GBL_CHK_TempStr = "支付宝账号不能包含百分号<br>" & VbCrLf
	If isNumeric(Form_Def_UserTestNumber) = 0 Then Form_Def_UserTestNumber = 0
	If isNumeric(Form_DEF_seller_minpoints) = 0 Then GBL_CHK_TempStr = "支付宝一次最少充值必须为数字<br>" & VbCrLf
	If isNumeric(Form_DEF_seller_exchangescale) = 0 Then GBL_CHK_TempStr = "支付宝元兑换率必须为数字<br>" & VbCrLf

End Function

Function ReplaceStr(str)

	ReplaceStr = Replace(Str,"""","""""")

End Function

Function MakeDataBaseLinkFile

	Dim TempStr,N
	TempStr = ""
	TempStr = TempStr & chr(60) & "%" & VbCrLf
	TempStr = TempStr & "const DEF_UserEnableUserTitle = " & Form_DEF_UserEnableUserTitle & VbCrLf	
	TempStr = TempStr & "const DEF_UserUserTitleNeedLevel = " & Form_DEF_UserUserTitleNeedLevel & VbCrLf
	TempStr = TempStr & "const LMT_UserNameEnableEnglishWords = " & Form_LMT_UserNameEnableEnglishWords & VbCrLf
	TempStr = TempStr & "const LMT_UserNameEnableChineseChar = " & Form_LMT_UserNameEnableChineseChar & VbCrLf
	TempStr = TempStr & "const LMT_UserNameEnableChineseWords = " & Form_LMT_UserNameEnableChineseWords & VbCrLf
	TempStr = TempStr & "const DEF_User_RegPoints = " & Form_DEF_User_RegPoints & VbCrLf
	TempStr = TempStr & "const LMT_EnableRegNewUsers = " & Form_LMT_EnableRegNewUsers & VbCrLf
	TempStr = TempStr & "const DEF_ShortestUserName = " & Form_DEF_ShortestUserName & VbCrLf
	TempStr = TempStr & "const DEF_RegNewUserTotalRestTime = " & Form_DEF_RegNewUserTotalRestTime & VbCrLf
	TempStr = TempStr & "const DEF_UserNewRegAttestMode = " & Form_DEF_UserNewRegAttestMode & VbCrLf
	TempStr = TempStr & "const DEF_UserActivationExpiresDay = " & Form_DEF_UserActivationExpiresDay & VbCrLf
	TempStr = TempStr & "const DEF_User_GetPassMode = " & Form_DEF_User_GetPassMode & VbCrLf

	TempStr = TempStr & "Dim DEF_UserLevelString,DEF_UserLevelNum,DEF_UserLevelPoints" & VbCrLf
	TempStr = TempStr & "DEF_UserLevelString = Array("
	For n = 0 to DEF_UserLevelNum
		If n = 0 Then
			TempStr = TempStr & """" & ReplaceStr(Form_DEF_UserLevelString(n)) & """"
		Else
			TempStr = TempStr & ",""" & ReplaceStr(Form_DEF_UserLevelString(n)) & """"
		End If
	Next
	TempStr = TempStr & ")" & VbCrLf
	
	TempStr = TempStr & "DEF_UserLevelPoints = Array("
	For n = 0 to DEF_UserLevelNum
		If n = 0 Then
			TempStr = TempStr & Form_DEF_UserLevelPoints(n)
		Else
			TempStr = TempStr & "," & Form_DEF_UserLevelPoints(n)
		End If
	Next
	TempStr = TempStr & ")" & VbCrLf

	TempStr = TempStr & "DEF_UserLevelNum = Ubound(DEF_UserLevelString,1)" & VbCrLf	
	TempStr = TempStr & "Dim DEF_UserOfficerString,DEF_UserOfficerNum" & VbCrLf	
	TempStr = TempStr & "DEF_UserOfficerString = Array("
	For n = 0 to DEF_UserOfficerNum
		If n = 0 Then
			TempStr = TempStr & """" & ReplaceStr(Form_DEF_UserOfficerString(n)) & """"
		Else
			TempStr = TempStr & ",""" & ReplaceStr(Form_DEF_UserOfficerString(n)) & """"
		End If
	Next
	TempStr = TempStr & ")" & VbCrLf

	TempStr = TempStr & "DEF_UserOfficerNum = Ubound(DEF_UserOfficerString,1)" & VbCrLf
	TempStr = TempStr & "const DEF_FiltrateUserNameString = " & Chr(34) & LCase(Form_DEF_FiltrateUserNameString) & chr(34) & VbCrLf

	TempStr = TempStr & "const DEF_UserShortestPassword = " & Form_DEF_UserShortestPassword & VbCrLf
	TempStr = TempStr & "const DEF_UserShortestPasswordMaster = " & Form_DEF_UserShortestPasswordMaster & VbCrLf
	TempStr = TempStr & "const Def_UserTestNumber = " & Form_Def_UserTestNumber & VbCrLf
	TempStr = TempStr & "const DEF_seller_email = " & Chr(34) & LCase(Form_DEF_seller_email) & chr(34) & VbCrLf
	TempStr = TempStr & "const DEF_seller_minpoints = " & Form_DEF_seller_minpoints & VbCrLf
	TempStr = TempStr & "const DEF_seller_exchangescale = " & Form_DEF_seller_exchangescale & VbCrLf
	TempStr = TempStr & "%" & chr(62) & VbCrLf

	ADODB_SaveToFile TempStr,"../../inc/User_Setup.ASP"
	CALL Update_InsertSetupRID(1051,"inc/User_Setup.ASP",2,TempStr," and ClassNum=" & 2)
	If GBL_CHK_TempStr = "" Then
		Response.Write "<br><span class=greenfont>2.成功完成设置！</span>"
	Else
		%><%=GBL_CHK_TempStr%><br>服务器不支持在线写入文件功能，请使用FTP等功能，将<span class=redfont>inc/User_Setup.ASP</span>文件替换成框中内容(注意备份)<p>
		<textarea name="fileContent" cols="80" rows="30" class=fmtxtra><%=Server.htmlencode(TempStr)%></textarea><%
		GBL_CHK_TempStr = ""
	End If

End Function
%>