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

Manage_sitehead DEF_SiteNameString & " - ����Ա",""
frame_TopInfo
DisplayUserNavigate("�û�ע���������")
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
		���ã�<a href=../SiteManage/SiteSetup.asp>��̳���ò���</a> <a href=../SiteManage/UploadSetup.asp>�ϴ�����</a>
		<span class=grayfont>�û�ע�����</span>
		<a href=../SiteManage/UbbcodeSetup.asp>UBB�������</a>
		<br><span class=grayfont>(����������վ���û�ע���������������ý��ᷢ�����ش���)<br><br>
		��������ú�����վ�����������У��뽫LeadBBS���°��User_Setup.asp���ǻ�ȥ</span>
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
<input type=submit name=�ύ value=�ύ class=fmbtn>
<input type=reset name=ȡ�� value=ȡ�� class=fmbtn>
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
			<td class=tdbox width=120>�Զ�ͷ��</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_UserEnableUserTitle value=0<%If Form_DEF_UserEnableUserTitle = 0 Then%> checked<%End If%>></td><td>��ֹ</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_UserEnableUserTitle value=1<%If Form_DEF_UserEnableUserTitle = 1 Then%> checked<%End If%>></td><td>���� (<span class=grayfont>�Ƿ������û��Զ���ͷ��</span>)</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>ͷ��<%=DEF_PointsName(3)%></td>
			<td class=tdbox>
				<select name=Form_DEF_UserUserTitleNeedLevel><%
				For N = 0 to DEF_UserLevelNum
					If N = Form_DEF_UserUserTitleNeedLevel Then
						Response.write "				<option value=" & N & " selected>" & N & "." & DEF_UserLevelString(N) & "</option>" & VbCrLf
					Else
						Response.write "				<option value=" & N & ">" & N & "." & DEF_UserLevelString(N) & "</option>" & VbCrLf
					End If
				Next%>(<span class=grayfont>�û��Զ���ͷ������Ҫ��<%=DEF_PointsName(3)%></span>)
				</select> �������Զ���ͷ�Σ���ָ���Զ���ͷ����Ҫ��ﵽ��<%=DEF_PointsName(3)%></td>
		</tr>
		<tr>
			<td class=tdbox>�� �� ��</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_LMT_UserNameEnableEnglishWords value=0<%If Form_LMT_UserNameEnableEnglishWords = 0 Then%> checked<%End If%>></td><td>��ֹʹ�������ַ�(��ĸ����)</td>
          		<td><input class=fmchkbox type=radio name=Form_LMT_UserNameEnableEnglishWords value=1<%If Form_LMT_UserNameEnableEnglishWords = 1 Then%> checked<%End If%>></td><td>����ʹ�������ַ�(��ĸ����)</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>�� �� ��</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_LMT_UserNameEnableChineseChar value=0<%If Form_LMT_UserNameEnableChineseChar = 0 Then%> checked<%End If%>></td><td>��ֹʹ�����ķ���(���,���ĵ��ַ�)</td>
          		<td><input class=fmchkbox type=radio name=Form_LMT_UserNameEnableChineseChar value=1<%If Form_LMT_UserNameEnableChineseChar = 1 Then%> checked<%End If%>></td><td>����ʹ�����ķ���(���,���ĵ��ַ�)</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>�� �� ��</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_LMT_UserNameEnableChineseWords value=0<%If Form_LMT_UserNameEnableChineseWords = 0 Then%> checked<%End If%>></td><td>��ֹʹ�����ĺ���</td>
          		<td><input class=fmchkbox type=radio name=Form_LMT_UserNameEnableChineseWords value=1<%If Form_LMT_UserNameEnableChineseWords = 1 Then%> checked<%End If%>></td><td>����ʹ�����ĺ���</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>ע��<%=DEF_PointsName(0)%></td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_User_RegPoints" maxlength="4" size="10" value="<%=htmlencode(Form_DEF_User_RegPoints)%>"><span class=grayfont>(��ע���û���ӵ�е�<%=DEF_PointsName(0)%>������Ĭ��Ϊ0)</span></td>
		</tr>
		<tr>
			<td class=tdbox>����ע��</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_LMT_EnableRegNewUsers value=0<%If Form_LMT_EnableRegNewUsers = 0 Then%> checked<%End If%>></td><td>��ֹע�����û�</td>
          		<td><input class=fmchkbox type=radio name=Form_LMT_EnableRegNewUsers value=1<%If Form_LMT_EnableRegNewUsers = 1 Then%> checked<%End If%>></td><td>�������û�ע��</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>�û�����</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_ShortestUserName" maxlength="2" size="10" value="<%=htmlencode(Form_DEF_ShortestUserName)%>"><span class=grayfont>(����ע����û���������ַ���������λ�ֽ�)</span></td>
		</tr>
		<tr>
			<td class=tdbox>�������</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_UserShortestPassword" maxlength="2" size="10" value="<%=htmlencode(Form_DEF_UserShortestPassword)%>"><span class=grayfont>(����ʹ�õ��û����������ַ���������λ�ֽڣ������ͨ�û�)</span></td>
		</tr>
		<tr>
			<td class=tdbox>��������</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_UserShortestPasswordMaster" maxlength="2" size="10" value="<%=htmlencode(Form_DEF_UserShortestPasswordMaster)%>"><span class=grayfont>(����ʹ�õ�<%=DEF_PointsName(8)%>���ϳ�Ա����������ַ�����)</span></td>
		</tr>
		<tr>
			<td class=tdbox>ע����</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_RegNewUserTotalRestTime" maxlength="6" size="10" value="<%=htmlencode(Form_DEF_RegNewUserTotalRestTime)%>"><span class=grayfont>(������̳�ڴ�ʱ����ֻ����ע��һ�����û�����λ��)</span></td>
		</tr>
		<tr>
			<td class=tdbox>ע����֤</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_UserNewRegAttestMode value=0<%If Form_DEF_UserNewRegAttestMode = 0 Then%> checked<%End If%>></td><td>�޼���,ע�ἴΪ��ʽ��Ա</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_UserNewRegAttestMode value=1<%If Form_DEF_UserNewRegAttestMode = 1 Then%> checked<%End If%>></td><td>�ʼ�����(������ؿ����ʼ����͹���)</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_UserNewRegAttestMode value=2<%If Form_DEF_UserNewRegAttestMode = 2 Then%> checked<%End If%>></td><td>��������(����Ա��̨����)</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>����ʱ��</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_UserActivationExpiresDay" maxlength="3" size="10" value="<%=htmlencode(Form_DEF_UserActivationExpiresDay)%>"><span class=grayfont>(ע����û�������ָ�������ڼ������ϵͳ����ɾ���û�������λ���죬��д0��ʾ�����ƣ����ñ���)</span></td>
		</tr>
		<tr>
			<td class=tdbox>�����һ�</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_User_GetPassMode value=0<%If Form_DEF_User_GetPassMode = 0 Then%> checked<%End If%>></td><td>��ֹ�һ�</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_User_GetPassMode value=1<%If Form_DEF_User_GetPassMode = 1 Then%> checked<%End If%>></td><td>��ʱ�������</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_User_GetPassMode value=2<%If Form_DEF_User_GetPassMode = 2 Then%> checked<%End If%>></td><td>��ʱ������� ��δ����ͬʱ�����ʼ�����֪ͨ������</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox><%=DEF_PointsName(3)%>����</td>
			<td class=tdbox>
				<table border=0 cellpadding=1 cellspacing=0>
				<tr>
					<td>&nbsp;<%=DEF_PointsName(3)%></td>
					<td>&nbsp;�ƺ�</td>
					<td>&nbsp;Ҫ�󷢱�����</td>
				</td><%
			For n = 0 to DEF_UserLevelNum
				%>
				<tr>
					<td>&nbsp;<%=Right(" " & N,2)%>��</td>
					<td>&nbsp;<input class=fminpt type="text" name="Form_DEF_UserLevelString<%=N%>" maxlength="18" size="20" value="<%=htmlencode(Form_DEF_UserLevelString(n))%>"></td>
					<td>&nbsp;<input class=fminpt type="text" name="Form_DEF_UserLevelPoints<%=N%>" maxlength="14" size="10" value="<%=htmlencode(Form_DEF_UserLevelPoints(n))%>"></td>
				</td>
				<%
			Next
			%>
				</table>
				&nbsp;<span class=grayfont>��������ָ�������������������(����ɾ������)</span></td>
		</tr>
		<tr>
			<td class=tdbox><%=DEF_PointsName(9)%></td>
			<td class=tdbox>
				<table border=0 cellpadding=1 cellspacing=0>
				<tr>
					<td>&nbsp;���</td>
					<td>&nbsp;�ƺ�</td>
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
				&nbsp;<span class=grayfont><%=DEF_PointsName(9)%>��������ʹ��html���룬��������ʹ�����ţ�</span></td>
		</tr>
		<tr>
			<td class=tdbox>��������</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_FiltrateUserNameString" maxlength="1024" size="50" value="<%=htmlencode(Form_DEF_FiltrateUserNameString)%>">
			<br><span class=grayfont>(ʹ��|�ָ���ע���û������û�ͷ�κ�ǩ�������ܰ�����������)</span></td>
		</tr>
		<tr>
			<td class=tdbox>�� ֤ ��</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_Def_UserTestNumber value=0<%If Form_Def_UserTestNumber = 0 Then%> checked<%End If%>></td><td>��̳ϵͳĬ��(�̳���̳��������)</td><tr>
          		<tr><td><input class=fmchkbox type=radio name=Form_Def_UserTestNumber value=1<%If Form_Def_UserTestNumber = 1 Then%> checked<%End If%>></td><td>�϶�ʹ��ע����֤��</td></tr>
          		<tr><td><input class=fmchkbox type=radio name=Form_Def_UserTestNumber value=2<%If Form_Def_UserTestNumber = 2 Then%> checked<%End If%>></td><td>�϶���ʹ��ע����֤��</td></tr>
          		</table></td>
		</tr>
		<tr>
			<td class=tdbox>֧�����˺�</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_seller_email" maxlength="150" size="50" value="<%=htmlencode(Form_DEF_seller_email)%>">
			<br><span class=grayfont>��д��վ��ֵ<%=DEF_PointsName(1)%>�����˺�,һ����EMAIL��ַ,����д��ʾ������֧������</span></td>
		</tr>
		<tr>
			<td class=tdbox>֧����һ�����ٳ�ֵ</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_seller_minpoints" maxlength="6" size="10" value="<%=htmlencode(Form_DEF_seller_minpoints)%>"><span class=grayfont>(��λ,RMBԪ)</span></td>
		</tr>
		<tr>
			<td class=tdbox>֧����Ԫ�һ���</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_seller_exchangescale" maxlength="6" size="10" value="<%=htmlencode(Form_DEF_seller_exchangescale)%>"><span class=grayfont>(һԪRMB�һ�����<%=DEF_PointsName(1)%>)</span></td>
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
	
	If isNumeric(Form_DEF_UserEnableUserTitle) = 0 Then GBL_CHK_TempStr = "�Ƿ������û��Զ���ͷ�α���Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_UserUserTitleNeedLevel) = 0 Then GBL_CHK_TempStr = "ָ���Զ���ͷ����Ҫ��ﵽ��" & DEF_PointsName(3) & "����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_LMT_UserNameEnableEnglishWords) = 0 Then GBL_CHK_TempStr = "�Ƿ������û���ʹ�������ַ�(��ĸ����)����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_LMT_UserNameEnableChineseChar) = 0 Then GBL_CHK_TempStr = "�Ƿ������û�ʹ�����ķ���(���,���ĵ��ַ�)����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_LMT_UserNameEnableChineseWords) = 0 Then GBL_CHK_TempStr = "�Ƿ������û�ʹ�����ĺ��ֱ���Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_User_RegPoints) = 0 Then GBL_CHK_TempStr = "ע���û���ӵ�е�" & DEF_PointsName(0) & "��������Ϊ����<br>" & VbCrLf
	If isNumeric(Form_LMT_EnableRegNewUsers) = 0 Then GBL_CHK_TempStr = "����ע�����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_ShortestUserName) = 0 Then GBL_CHK_TempStr = "����������ַ���������Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_RegNewUserTotalRestTime) = 0 Then GBL_CHK_TempStr = "ע��������Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_UserNewRegAttestMode) = 0 Then GBL_CHK_TempStr = "ע����֤��ʽ����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_UserActivationExpiresDay) = 0 Then GBL_CHK_TempStr = "����ʱ�����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_User_GetPassMode) = 0 Then GBL_CHK_TempStr = "�����һر���Ϊ����<br>" & VbCrLf
	For n = 0 to DEF_UserLevelNum
		If inStr(Form_DEF_UserLevelString(n),"%") Then
			GBL_CHK_TempStr = "��" & N & DEF_PointsName(3) & "���Ʋ��ܰ����ٷֺ�<br>" & VbCrLf
		End If
		If isNumeric(Form_DEF_UserLevelPoints(n)) = 0 Then
			GBL_CHK_TempStr = "��" & N & DEF_PointsName(3) & "Ҫ�󷢱�������������Ϊ����<br>" & VbCrLf
			Exit Function
		End If
	Next
	For n = 0 to DEF_UserOfficerNum
		If inStr(Form_DEF_UserOfficerString(n),"%") Then
			GBL_CHK_TempStr = "��" & N & "���" & DEF_PointsName(9) & "���ܰ����ٷֺ�<br>" & VbCrLf
		End If
	Next
	If inStr(Form_DEF_FiltrateUserNameString,"%") Then GBL_CHK_TempStr = "�����һز��ܰ����ٷֺ�<br>" & VbCrLf
	If isNumeric(Form_DEF_UserShortestPassword) = 0 Then GBL_CHK_TempStr = "�������������ַ���������Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_UserShortestPasswordMaster) = 0 Then GBL_CHK_TempStr = "��̳�����Ա�������������ַ���������Ϊ����<br>" & VbCrLf
	If inStr(Form_DEF_seller_email,"%") Then GBL_CHK_TempStr = "֧�����˺Ų��ܰ����ٷֺ�<br>" & VbCrLf
	If isNumeric(Form_Def_UserTestNumber) = 0 Then Form_Def_UserTestNumber = 0
	If isNumeric(Form_DEF_seller_minpoints) = 0 Then GBL_CHK_TempStr = "֧����һ�����ٳ�ֵ����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_seller_exchangescale) = 0 Then GBL_CHK_TempStr = "֧����Ԫ�һ��ʱ���Ϊ����<br>" & VbCrLf

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
		Response.Write "<br><span class=greenfont>2.�ɹ�������ã�</span>"
	Else
		%><%=GBL_CHK_TempStr%><br>��������֧������д���ļ����ܣ���ʹ��FTP�ȹ��ܣ���<span class=redfont>inc/User_Setup.ASP</span>�ļ��滻�ɿ�������(ע�ⱸ��)<p>
		<textarea name="fileContent" cols="80" rows="30" class=fmtxtra><%=Server.htmlencode(TempStr)%></textarea><%
		GBL_CHK_TempStr = ""
	End If

End Function
%>