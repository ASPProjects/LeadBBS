<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Board_popfun.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<%
DEF_BBS_HomeUrl = "../../"
Const MaxLinkNum = 100
initDatabase
GBL_CHK_TempStr = ""
CheckSupervisorPass

Dim Form_SavePoints
Dim Form_DEF_ManageDir
Dim Form_DEF_BBS_Name
Dim Form_DEF_BBS_DarkColor,Form_DEF_BBS_LightDarkColor,Form_DEF_BBS_Color,Form_DEF_BBS_LightColor,Form_DEF_BBS_LightestColor,Form_DEF_BBS_TableHeadColor
Dim Form_DEF_BBS_MaxLayer,Form_DEF_UsedDataBase,Form_DEF_BBS_SearchMode

Dim Form_DEF_BBS_AnnouncePoints,Form_DEF_BBS_PrizeAnnouncePoints,Form_DEF_BBS_MakeGoodAnnouncePoints,Form_DEF_BBS_MaxTopAnnounce,Form_DEF_BBS_MaxAllTopAnnounce
Dim Form_DEF_BBS_DisplayTopicLength,Form_DEF_BBS_ScreenWidth,Form_DEF_BBS_LeftTDWidth
Dim Form_DEF_MasterCookies,Form_DEF_SiteNameString
Dim Form_DEF_SupervisorUserName,Form_DEF_MaxTextLength

Dim Form_DEF_MaxListNum,Form_DEF_TopicContentMaxListNum
Dim Form_DEF_MaxJumpPageNum,Form_DEF_DisplayJumpPageNum
Dim Form_DEF_MaxBoardMastNum
Dim Form_DEF_EnableUserHidden,Form_DEF_VOTE_MaxNum
Dim Form_DEF_MaxLoginTimes,Form_DEF_EnableUpload,Form_DEF_EnableGFL
Dim Form_DEF_UserOnlineTimeOut,Form_DEF_faceMaxNum
Dim Form_DEF_AllDefineFace,Form_DEF_AllFaceMaxWidth
Dim Form_DEF_BBS_EmailMode,Form_DEF_EnableAttestNumber,Form_DEF_AttestNumberPoints
Dim Form_DEF_EnableUnderWrite,Form_DEF_NeedOnlineTime
Dim Form_DEF_EnableForbidIP,Form_DEF_TopAdString
Dim Form_DEF_RestSpaceTime,Form_DEF_LoginSpaceTime,Form_DEF_AccessDatabase,Form_DEF_SiteHomeUrl
Dim Form_DEF_DefaultStyle
Dim Form_DEF_EnableFlashUBB,Form_DEF_EnableImagesUBB,Form_DEF_AnnounceFontSize,Form_DEF_EditAnnounceDelay
Dim Form_DEF_DisplayOnlineUser,Form_DEF_EnableSpecialTopic,Form_DEF_UBBiconNumber,Form_DEF_EnableDelAnnounce
Dim Form_DEF_PointsName,Form_DEF_EnableMakeTopAnc,Form_DEF_EnableDatabaseCache
Dim Form_DEF_WriteEventSpace,Form_DEF_EnableTreeView,Form_DEF_EditAnnounceExpires
Dim Form_DEF_RepeatLoginTimeOut,Form_DEF_FSOString,Form_DEF_Now
Redim Form_DEF_PointsName(Ubound(DEF_PointsName))
Dim Form_DEF_LineHeight,Form_DEF_RegisterFile,Form_DEF_LimitTitle,Form_DEF_DownKey

Dim Form_DEF_UpdateInterval,Form_DEF_BottomInfo,Form_DEF_GBL_Description

Dim DEF_PointsNameBak
DEF_PointsNameBak = Array("����","����","����","�ȼ�","����","��֤��Ա","�ܰ���","������","��̳����","����","רҵ�û�")

Dim DEF_Sideparameter_String,Form_DEF_Sideparameter
DEF_Sideparameter_String = Array("����-��ҳ��ֹ��ʾ","����-��ҳ������ʾ�������Ĭ�Ϲر�״̬(�û���ͨ�������ʾ)","����-���濪����ʾ","����-���濪����ʾ�������Ĭ�Ϲر�״̬(�û���ͨ�������ʾ)","���-������Ӷ�ҳ�ظ�������ʾ��ϸҳ(Ĭ�Ͻ���ʾβҳ)","<span class=grayfont>������</span>","<span class=grayfont>������</span>","<span class=grayfont>������</span>","<span class=grayfont>������</span>","���û�������(������ѶQQ����,��Ҫ������չ����)","<span class=grayfont>������</span>","<span class=grayfont>������</span>","<span class=grayfont>������</span>","<span class=grayfont>������</span>","<span class=grayfont>������</span>","����Rewriteα��̬(���ô��ȷ���ռ�����ȷ��װ������Rewrite)","���ð�������İ�鵼��","���ò鿴����ҳ������İ�鵼��","��������б�Ĭ�Ͻ���ʾ����")

GetDefaultValue

Manage_sitehead DEF_SiteNameString & " - ����Ա",""
frame_TopInfo
DisplayUserNavigate("��̳��������")
If GBL_CHK_Flag=1 Then
	SiteLink
Else
	DisplayLoginForm
End If
frame_BottomInfo
closeDataBase
Manage_Sitebottom("none")

Function SiteLink

%>
<form name="pollform3sdx" method="post" action="SiteSetup.asp">
<input type="hidden" name="SubmitFlag" value=yes>
<p>
		���ã�<span class=grayfont>��̳���ò���</span> <a href=UploadSetup.asp>�ϴ�����</a>
		<a href=../User/UserSetup.asp>�û�ע�����</a>
		<a href=UbbcodeSetup.asp>UBB�������</a>
		<br>
		<span class=grayfont>(����Ϊ��վ��������ע���޸ģ���������ý��ᷢ�����ش���)<br><br>
		��������ú�����վ�����������У��뽫LeadBBS���°��BBSSetup.asp���ǻ�ȥ</span>
</p>
<%If Request.Form("SubmitFlag") <> "" Then
	CheckLinkValue
End If%>
<b><span class=redfont><%=GBL_CHK_TempStr%></span></b>
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
</form>
<%

End Function

Function CheckLinkValue

	GetFormValue

End Function

Function DisplayDatabaseLink

		%>
		<table border=0 cellpadding=0 cellspacing=0 width="100%" class=frame_table>
		<tr>
			<td class=tdbox width=120>��̳����</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_BBS_Name" maxlength="30" size="30" value="<%=htmlencode(Form_DEF_BBS_Name)%>"><span class=note>(�255��)</span></td>
		</tr>
		<tr>
			<td class=tdbox>��̳����<br>����ʣ��</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_SavePoints" maxlength="14" size="30" value="<%=htmlencode(Form_SavePoints)%>"><br><span class=note>(���а���������Ա�Է����������еĽ�������ʹ����ֵ�����ܽ���һ�ɼ�ȥָ��������ʹ�����Զ����٣�ֱ���������Ϊֹ)</span></td>
		</tr>
		<tr>
			<td class=tdbox>����Ŀ¼</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_ManageDir" maxlength="30" size="30" value="<%=htmlencode(Form_DEF_ManageDir)%>"><span class=note>(��̳ʹ�õ�Ŀ¼��Ĭ��Ϊmanage��ע����ʵ�Ĺ���Ŀ¼��˱���һ��)</span></td>
		</tr>
		<tr bgcolor=<%=DEF_BBS_LightColor%> class=TBBG1>
			<td class=tdbox colspan=2>��ɫ����(ΪĳЩ�������֧����ʽ����趨)</td>
		</tr>
		<tr>
			<td class=tdbox>�� �� ɫ</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_BBS_DarkColor" maxlength="255" size="30" value="<%=htmlencode(Form_DEF_BBS_DarkColor)%>"><span class=note>(DEF_BBS_DarkColor)</span></td>
		</tr>
		<tr>
			<td class=tdbox>�� �� ɫ</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_BBS_LightDarkColor" maxlength="255" size="30" value="<%=htmlencode(Form_DEF_BBS_LightDarkColor)%>"><span class=note>(DEF_BBS_LightDarkColor)</span></td>
		</tr>
		<tr>
			<td class=tdbox>��̳��ɫ</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_BBS_Color" maxlength="255" size="30" value="<%=htmlencode(Form_DEF_BBS_Color)%>"><span class=note>(DEF_BBS_Color)</span></td>
		</tr>
		<tr>
			<td class=tdbox>�� �� ɫ</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_BBS_LightColor" maxlength="255" size="30" value="<%=htmlencode(Form_DEF_BBS_LightColor)%>"><span class=note>(DEF_BBS_LightColor)</span></td>
		</tr>
		<tr>
			<td class=tdbox>�� �� ɫ</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_BBS_LightestColor" maxlength="255" size="30" value="<%=htmlencode(Form_DEF_BBS_LightestColor)%>"><span class=note>(DEF_BBS_LightestColor������ɫ)</span></td>
		</tr>
		<tr>
			<td class=tdbox>���ͷɫ</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_BBS_TableHeadColor" maxlength="255" size="30" value="<%=htmlencode(Form_DEF_BBS_TableHeadColor)%>"><span class=note>(DEF_BBS_TableHeadColor)</span></td>
		</tr>

		<tr bgcolor=<%=DEF_BBS_LightColor%> class=TBBG1>
			<td class=tdbox colspan=2>��������</td>
		</tr>
		<tr>
			<td class=tdbox>�ظ�����</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_BBS_MaxLayer" maxlength="255" size="10" value="<%=htmlencode(Form_DEF_BBS_MaxLayer)%>"><span class=note>(��״���ʱ��ʾ�����ظ����������ڵ���Ϊ�������������������)</span></td>
		</tr>
		<tr>
			<td class=tdbox>�� �� ��</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_UsedDataBase value=1<%If Form_DEF_UsedDataBase = 1 Then%> checked<%End If%>></td><td>Access</td>
          		<td>
          			<input class=fmchkbox type=radio name=Form_DEF_UsedDataBase value=0<%If Form_DEF_UsedDataBase = 0 Then%> checked<%End If%>>
          		</td>
          		<td>Microsoft SQL Server</td>
          		<td>
          			<input class=fmchkbox type=radio name=Form_DEF_UsedDataBase value=2<%If Form_DEF_UsedDataBase = 2 Then%> checked<%End If%>>
          		</td>
          		<td>MySQL</td>
          		<td><span class=note>&nbsp; (֧��ACCESS��MSSQL�������ݿ⣬<span class=redfont>��С������</span>)</span></td></tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>����ģʽ</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_BBS_SearchMode value=0<%If Form_DEF_BBS_SearchMode = 0 Then%> checked<%End If%>></td><td>����������</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_BBS_SearchMode value=1<%If Form_DEF_BBS_SearchMode = 1 Then%> checked<%End If%>></td><td>ģ����ѯ</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_BBS_SearchMode value=2<%If Form_DEF_BBS_SearchMode = 2 Then%> checked<%End If%>></td><td>ȫ�ļ���(��MSSQL����ע���Ƿ��Ѿ���װȫ�ķ���)</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>��������</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_BBS_AnnouncePoints" maxlength="4" size="10" value="<%=htmlencode(Form_DEF_BBS_AnnouncePoints)%>"><span class=note>(��������<%=DEF_PointsName(0)%>����)</span></td>
		</tr>
		<tr>
			<td class=tdbox>���۳ͷ�</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_BBS_PrizeAnnouncePoints" maxlength="4" size="10" value="<%=htmlencode(Form_DEF_BBS_PrizeAnnouncePoints)%>"><span class=note>(�����������<%=DEF_PointsName(1)%>�������ͷ�<%=DEF_PointsName(0)%>����Ϊ���ʾ��ֹ)</span></td>
		</tr>
		<tr>
			<td class=tdbox>��������</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_BBS_MakeGoodAnnouncePoints" maxlength="4" size="10" value="<%=htmlencode(Form_DEF_BBS_MakeGoodAnnouncePoints)%>"><span class=note>(�������Ӽ�<%=DEF_PointsName(0)%>����)</span></td>
		</tr>
		<tr>
			<td class=tdbox>��ඥ��</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_BBS_MaxTopAnnounce" maxlength="4" size="10" value="<%=htmlencode(Form_DEF_BBS_MaxTopAnnounce)%>"><span class=note>(ÿ������������ö���������)</span></td>
		</tr>
		<tr>
			<td class=tdbox>����ܹ�</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_BBS_MaxAllTopAnnounce" maxlength="2" size="10" value="<%=htmlencode(Form_DEF_BBS_MaxAllTopAnnounce)%>"><span class=note>(��̳���������̶ܹ���)</span></td>
		</tr>
		<tr>
			<td class=tdbox>���ⳤ��</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_BBS_DisplayTopicLength" maxlength="3" size="10" value="<%=htmlencode(Form_DEF_BBS_DisplayTopicLength)%>"><span class=note>(��ʾ��������ĳ��ȣ�������̳�����б���λ�ֽ�)</span></td>
		</tr>
		<tr>
			<td class=tdbox>��̳���</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_BBS_ScreenWidth" maxlength="10" size="10" value="<%=htmlencode(Form_DEF_BBS_ScreenWidth)%>"><span class=note>(�££ӵĿ�ȣ������ǰٷֱ�)</span></td>
		</tr>
		<tr>
			<td class=tdbox>�������</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_BBS_LeftTDWidth" maxlength="10" size="10" value="<%=htmlencode(Form_DEF_BBS_LeftTDWidth)%>"><span class=note>(�££ӵ�������ȣ������ǰٷֱ�)</span></td>
		</tr>
		<tr>
			<td class=tdbox>Cookies </td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_MasterCookies" maxlength="20" size="20" value="<%=htmlencode(Form_DEF_MasterCookies)%>"><span class=note>(��װCookie������ǰ�)</span></td>
		</tr>
		<tr>
			<td class=tdbox>��վ����</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_SiteNameString" maxlength="255" size="20" value="<%=htmlencode(Form_DEF_SiteNameString)%>"><span class=note>(��̳��վ������)</span></td>
		</tr>
		<tr>
			<td class=tdbox>�� �� Ա</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_SupervisorUserName" maxlength="255" size="20" value="<%=htmlencode(Form_DEF_SupervisorUserName)%>"><span class=redfont>(��ע���Сд�����ú������µ�¼�������Ա���ŷָ�)</span></td>
		</tr>
		<tr>
			<td class=tdbox>���ݳ���</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_MaxTextLength" maxlength="5" size="10" value="<%=htmlencode(Form_DEF_MaxTextLength)%>"><span class=note>(�����������ݡ�����Ϣ���ȣ�����Ա�������ֵ�ı������ݣ���λ�ֽ�)</span></td>
		</tr>
		<tr>
			<td class=tdbox>��ʾ��¼</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_MaxListNum" maxlength="3" size="10" value="<%=htmlencode(Form_DEF_MaxListNum)%>"><span class=note>(Ĭ�Ϸ�ҳ�г�������¼��)</span></td>
		</tr>
		<tr>
			<td class=tdbox>��ʾ����</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_TopicContentMaxListNum" maxlength="3" size="10" value="<%=htmlencode(Form_DEF_TopicContentMaxListNum)%>"><span class=note>(�鿴����ÿҳ��ʾ�����������������ʾ��)</span></td>
		</tr>
		<tr>
			<td class=tdbox>��תҳ��</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_MaxJumpPageNum" maxlength="3" size="10" value="<%=htmlencode(Form_DEF_MaxJumpPageNum)%>"><span class=note>(��ҳʱ�������������ֱ����תҳ��,��������ת��Ϊ0)</span></td>
		</tr>
		<tr>
			<td class=tdbox>��ʾ��ת</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_DisplayJumpPageNum" maxlength="3" size="10" value="<%=htmlencode(Form_DEF_DisplayJumpPageNum)%>"><span class=note>(��ʾ��תҳ����ע�ⲻҪ�����������ֱ����תҳ��)</span></td>
		</tr>
		<tr>
			<td class=tdbox>��������</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_MaxBoardMastNum" maxlength="3" size="10" value="<%=htmlencode(Form_DEF_MaxBoardMastNum)%>"><span class=note>(ÿ��������������Ŀ)</span></td>
		</tr>
		<tr>
			<td class=tdbox>��������</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr><td><input class=fmchkbox type=radio name=Form_DEF_EnableUserHidden value=1<%If Form_DEF_EnableUserHidden = 1 Then%> checked<%End If%>></td><td>����</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_EnableUserHidden value=0<%If Form_DEF_EnableUserHidden = 0 Then%> checked<%End If%>></td><td>��ֹ</td><td><span class=note>&nbsp; (�Ƿ����������û�����)</span></td></tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>ͶƱ��Ŀ</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_VOTE_MaxNum" maxlength="3" size="10" value="<%=htmlencode(Form_DEF_VOTE_MaxNum)%>"><span class=note>(���ͶƱ��Ŀ)</span></td>
		</tr>
		<tr>
			<td class=tdbox>��¼����</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_MaxLoginTimes" maxlength="6" size="10" value="<%=htmlencode(Form_DEF_MaxLoginTimes)%>"><span class=note>(����ĳһ�û��ظ��Ĵ����¼����)</span></td>
		</tr>
		<tr>
			<td class=tdbox>�������</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_RestSpaceTime" maxlength="6" size="10" value="<%=htmlencode(Form_DEF_RestSpaceTime)%>"><span class=note>(����ĳЩ������Ҫ��ʱ���������緢�������ŵȣ���λ��)</span></td>
		</tr>
		<tr>
			<td class=tdbox>��¼���</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_LoginSpaceTime" maxlength="6" size="10" value="<%=htmlencode(Form_DEF_LoginSpaceTime)%>"><span class=note>�û���¼�ۻ���������Ҫ������¼��ʱ��</span></td>
		</tr>
		<tr>
			<td class=tdbox>�ϴ�Ȩ��</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_EnableUpload value=0<%If Form_DEF_EnableUpload = 0 Then%> checked<%End If%>></td><td>�������ϴ�</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_EnableUpload value=1<%If Form_DEF_EnableUpload = 1 Then%> checked<%End If%>></td><td>ȫ���˿����ϴ�</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_EnableUpload value=2<%If Form_DEF_EnableUpload = 2 Then%> checked<%End If%>></td><td>������Ա���ϴ�</td>
          </tr>
          <tr>
          		<td><input class=fmchkbox type=radio name=Form_DEF_EnableUpload value=3<%If Form_DEF_EnableUpload = 3 Then%> checked<%End If%>></td><td>�����������Ͽ��ϴ�</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_EnableUpload value=4<%If Form_DEF_EnableUpload = 4 Then%> checked<%End If%>></td><td>��<%=DEF_PointsName(5)%>���ϴ�</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_EnableUpload value=5<%If Form_DEF_EnableUpload = 5 Then%> checked<%End If%>></td><td>��<%=DEF_PointsName(5)%>���������Ͽ��ϴ�</td>
          		</tr></table><span class=note>�����ָ�ϴ�����</span></td>
		</tr>
		<tr>
			<td class=tdbox>ͼ�����</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_EnableGFL value=1<%If Form_DEF_EnableGFL = 1 Then%> checked<%End If%>></td><td>����ʹ��AspJpeg���</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_EnableGFL value=0<%If Form_DEF_EnableGFL = 0 Then%> checked<%End If%>></td><td>��ֹʹ��</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>���߳�ʱ</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_UserOnlineTimeOut" maxlength="6" size="10" value="<%=htmlencode(Form_DEF_UserOnlineTimeOut)%>"><span class=note>(�����û���ָ��ʱ���ڲ������κη��ʣ�������ߣ���λ��)</span></td>
		</tr>
		<tr>
			<td class=tdbox>ͷ�����</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_faceMaxNum" maxlength="6" size="10" value="<%=htmlencode(Form_DEF_faceMaxNum)%>"><span class=note>(��̳Ĭ��ͷ��ĸ���)</span></td>
		</tr>
		<tr>
			<td class=tdbox>�Զ�ͷ��</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_AllDefineFace value=0<%If Form_DEF_AllDefineFace = 0 Then%> checked<%End If%>></td><td>��ֹ�Զ���ͷ��</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_AllDefineFace value=1<%If Form_DEF_AllDefineFace = 1 Then%> checked<%End If%>></td><td>�����Զ����κ�ͷ��</td>
          	 	</tr>
          	 	<tr>
          		<td><input class=fmchkbox type=radio name=Form_DEF_AllDefineFace value=2<%If Form_DEF_AllDefineFace = 2 Then%> checked<%End If%>></td><td>����վ��ͼƬ��������������վ��ͼƬ��Ϊͷ��</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_AllDefineFace value=3<%If Form_DEF_AllDefineFace = 3 Then%> checked<%End If%>></td><td>��������վ��ͼƬ���������ϴ�ͼƬ��Ϊͷ��</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>ͷ���С</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_AllFaceMaxWidth" maxlength="3" size="10" value="<%=htmlencode(Form_DEF_AllFaceMaxWidth)%>"><span class=note>(�Զ���ͷ�����󳤶ȺͿ�ȣ���λ����)</span></td>
		</tr>
		<tr>
			<td class=tdbox>�ʼ�����</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_BBS_EmailMode value=0<%If Form_DEF_BBS_EmailMode = 0 Then%> checked<%End If%>></td><td>��ֹ�ʼ�����</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_BBS_EmailMode value=1<%If Form_DEF_BBS_EmailMode = 1 Then%> checked<%End If%>></td><td>ʹ��EasyMail����</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_BBS_EmailMode value=2<%If Form_DEF_BBS_EmailMode = 2 Then%> checked<%End If%>></td><td>ʹ��Jmail����</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_BBS_EmailMode value=3<%If Form_DEF_BBS_EmailMode = 3 Then%> checked<%End If%>></td><td>ʹ��CDO����</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>�� ֤ ��</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<tr><td colspan=4><b>1.ȫ������</b></td></tr>
				<td width=5><input class=fmchkbox type=radio name=Form_DEF_EnableAttestNumber value=0<%If Form_DEF_EnableAttestNumber = 0 Then%> checked<%End If%>></td><td colspan=3>����</td>
				</tr>
				<tr><td colspan=4><b>2.������������̳��֤�빦��</b></td></tr>
				<tr>
          		<td><input class=fmchkbox type=radio name=Form_DEF_EnableAttestNumber value=1<%If Form_DEF_EnableAttestNumber = 1 Then%> checked<%End If%>></td><td>ʹ��ASPJPEG���</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_EnableAttestNumber value=2<%If Form_DEF_EnableAttestNumber = 2 Then%> checked<%End If%>></td><td>ʹ�������������֤��</td>
          		</tr>
				<tr><td colspan=4><b>3.����ȫ����֤�빦��(���������û�ע�ᣬ��¼�������Ȳ���)</b></td></tr>
          		<tr>
          		<td><input class=fmchkbox type=radio name=Form_DEF_EnableAttestNumber value=3<%If Form_DEF_EnableAttestNumber = 3 Then%> checked<%End If%>></td><td>ʹ��ASPJPEG���</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_EnableAttestNumber value=4<%If Form_DEF_EnableAttestNumber = 4 Then%> checked<%End If%>></td><td>ʹ�������������֤��</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>��֤�룲</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_AttestNumberPoints" maxlength="12" size="10" value="<%=htmlencode(Form_DEF_AttestNumberPoints)%>"><span class=note>(����ͨ������Ҫ��֤�빦��ʱ�����û�<%=DEF_PointsName(0)%>���ڴ�ֵʱ����ʹ����֤�룮������Ϊ0ʱĬ�ϼ���ֵ��Ч)</span></td>
		</tr>
		<tr>
			<td class=tdbox>ǩ������</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_EnableUnderWrite value=0<%If Form_DEF_EnableUnderWrite = 0 Then%> checked<%End If%>></td><td>��ֹʹ��ǩ��</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_EnableUnderWrite value=1<%If Form_DEF_EnableUnderWrite = 1 Then%> checked<%End If%>></td><td>����ʹ��ǩ��</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td>����ʱ��</td>
			<td><input class=fminpt type="text" name="Form_DEF_NeedOnlineTime" maxlength="10" size="10" value="<%=htmlencode(Form_DEF_NeedOnlineTime)%>"><span class=note>(��ʹĳЩȨ������Ҫ��������ʱ�䣬���緢������Ϊ0��ʾ�����ƣ���λ��)</span></td>
		</tr>
		<tr>
			<td>�ɣ�����</td>
			<td><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_EnableForbidIP value=0<%If Form_DEF_EnableForbidIP = 0 Then%> checked<%End If%>></td><td>�ر�����</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_EnableForbidIP value=1<%If Form_DEF_EnableForbidIP = 1 Then%> checked<%End If%>></td><td>��������</td>
          		<td>&nbsp; (<span class=note>���û��Ҫ���Σɣе�ַ�����Թر��������վ�ٶ�</span>)</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>�������</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_TopAdString" maxlength="4096" size="50" value="<%=htmlencode(Form_DEF_TopAdString)%>"><span class=note>(ʹ��HTML�﷨)</span></td>
		</tr>
		<tr>
			<td class=tdbox>�� �� ��</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_AccessDatabase" maxlength="255" size="20" value="<%=htmlencode(Form_DEF_AccessDatabase)%>"><span class=note>(Access���ݿ�Ĵ��·��������ڸ�Ŀ¼��ǰ�治�ü�/��)</span></td>
		</tr>
		<tr>
			<td class=tdbox>��վ��ҳ</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_SiteHomeUrl" maxlength="255" size="20" value="<%=htmlencode(Form_DEF_SiteHomeUrl)%>"><span class=note>(��ҳ��ַ�����þ���·��������ΪĬ��Ϊ��̳��ҳ)</span></td>
		</tr>
		<tr>
			<td class=tdbox>Ĭ�Ϸ��</td>
			<td class=tdbox>
				<input class=fminpt type="text" id="Form_DEF_DefaultStyle" name="Form_DEF_DefaultStyle" maxlength="8" size="10" value="<%=htmlencode(Form_DEF_DefaultStyle)%>">
				<select name=local onchange="$id('Form_DEF_DefaultStyle').value=this.value;">
				<%Dim N	
				for N = 0 to DEF_BoardStyleStringNum
					Response.Write "<option value=" & N
					If N = DEF_DefaultStyle Then Response.Write " selected"
					Response.Write ">" & DEF_BoardStyleString(N) & "</option>" & VbCrLf
				Next%></select><span class=note>(������վʱĬ�ϵ���ʾ���,��չ�����ֱ����д���)</span></td>
		</tr>
		<tr>
			<td class=tdbox>�� ý ��</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_EnableFlashUBB value=0<%If Form_DEF_EnableFlashUBB = 0 Then%> checked<%End If%>></td><td>��ֹ</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_EnableFlashUBB value=1<%If Form_DEF_EnableFlashUBB = 1 Then%> checked<%End If%>></td><td>����</td>
          		<td>&nbsp; (<span class=note>�Ƿ��������Flash��Real��mp3��asf�ȶ�ý���ļ�UBB��ǩ</span>)</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td>����ͼƬ</td>
			<td><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_EnableImagesUBB value=0<%If Form_DEF_EnableImagesUBB = 0 Then%> checked<%End If%>></td><td>��ֹ</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_EnableImagesUBB value=1<%If Form_DEF_EnableImagesUBB = 1 Then%> checked<%End If%>></td><td>����</td>
          		<td>&nbsp; (<span class=note>�Ƿ�����������ǩ���в���ͼƬ�ļ�</span>)</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>��������</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_AnnounceFontSize" maxlength="100" size="20" value="<%=htmlencode(Form_DEF_AnnounceFontSize)%>">
			<span class=note>(�����������ִ�С����������(��������)��������д12px��14px����д��14px;FONT-FAMILY:����;�� ��ʾ��ʾΪ14���غ�����)</span></td>
		</tr>
		<tr>
			<td class=tdbox>�༭���</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_EditAnnounceDelay" maxlength="8" size="10" value="<%=htmlencode(Form_DEF_EditAnnounceDelay)%>"><span class=note>(�û��ڷ�������ĳ��ʱ���ڱ༭����ӡ�ϱ༭�ۼ�����λ��)</span></td>
		</tr>
		<tr>
			<td class=tdbox>�༭����</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_EditAnnounceExpires" maxlength="8" size="10" value="<%=htmlencode(Form_DEF_EditAnnounceExpires)%>"><span class=note>(�û��ڷ�������ĳ��ʱ��󽫽�ֹ�༭����λ�룬��Ϊ0��ʾһֱ����)</span></td>
		</tr>
		<tr>
			<td class=tdbox>���߻�Ա</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_DisplayOnlineUser value=0<%If Form_DEF_DisplayOnlineUser = 0 Then%> checked<%End If%>></td><td>��ȫ��ֹ</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_DisplayOnlineUser value=1<%If Form_DEF_DisplayOnlineUser = 1 Then%> checked<%End If%>></td><td>������������ʾ</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_DisplayOnlineUser value=2<%If Form_DEF_DisplayOnlineUser = 2 Then%> checked<%End If%>></td><td>ֱ����ʾ������Ա</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_DisplayOnlineUser value=3<%If Form_DEF_DisplayOnlineUser = 3 Then%> checked<%End If%>></td><td>��ҳֱ����ʾ�����������ʾ</td>
          		</tr></table>
          		&nbsp;(<span class=note>�Ƿ�������ʾ���߻�Ա����</span>)</td>
		</tr>
		<tr>
			<td class=tdbox>��������</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_EnableSpecialTopic value=0<%If Form_DEF_EnableSpecialTopic = 0 Then%> checked<%End If%>></td><td>��ֹ</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_EnableSpecialTopic value=1<%If Form_DEF_EnableSpecialTopic = 1 Then%> checked<%End If%>></td><td>����</td>
          		<td>&nbsp; (<span class=note>�Ƿ�������ظ��ɼ����͹���������ֹ�Ļ��������а����ֹ����(��ʹ�����������)</span>)</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>������Ŀ</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_UBBiconNumber" maxlength="8" size="2" value="<%=htmlencode(Form_DEF_UBBiconNumber)%>"><span class=note>(����ѡ��ı���ͼƬ�ĸ�������Ϊ0��ʾ��ֹʹ�ò������)</span></td>
		</tr>
		<tr>
			<td class=tdbox>�� �� վ</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_EnableDelAnnounce value=0<%If Form_DEF_EnableDelAnnounce = 0 Then%> checked<%End If%>></td><td>����</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_EnableDelAnnounce value=1<%If Form_DEF_EnableDelAnnounce = 1 Then%> checked<%End If%>></td><td>��ֹ</td>
          		<td>&nbsp; (<span class=note>��������վ��������ֱ��ɾ����������(�ظ�����Ȼ��������ɾ��)������ת�Ƶ�����վ���棬ע��Ҫ�ȴ�������վ��飬��������444��</span>)</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>�����趨</td>
			<td class=tdbox>
				<table border=0 cellpadding=1 cellspacing=0>
				<tr>
					<td>&nbsp;���</td>
					<td>&nbsp;����</td>
					<td>&nbsp;Ĭ������</td>
				</td><%
			For n = 0 to Ubound(DEF_PointsName)
				%>
				<tr>
					<td>&nbsp;<%=Right(" " & N,2)%></td>
					<td>&nbsp;<input class=fminpt type="text" name="Form_DEF_PointsName<%=N%>" maxlength="18" size="20" value="<%=htmlencode(Form_DEF_PointsName(n))%>"></td>
					<td>&nbsp;<%=DEF_PointsNameBak(N)%></td>
				</td>
				<%
			Next
			%>
				</table></td>
		</tr>
		<tr>
			<td class=tdbox>�ظ�����</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_EnableMakeTopAnc value=0<%If Form_DEF_EnableMakeTopAnc = 0 Then%> checked<%End If%>></td><td>��</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_EnableMakeTopAnc value=1<%If Form_DEF_EnableMakeTopAnc = 1 Then%> checked<%End If%>></td><td>��</td>
          		<td>&nbsp; (<span class=note>�ظ������ʱ�Ƿ������ᵽ������ǰλ��</span>)</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>�ģ»���</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_EnableDatabaseCache value=0<%If Form_DEF_EnableDatabaseCache = 0 Then%> checked<%End If%>></td><td>��</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_EnableDatabaseCache value=1<%If Form_DEF_EnableDatabaseCache = 1 Then%> checked<%End If%>></td><td>��</td>
          		<td>&nbsp; (<span class=note>�Ƿ��������ݿ����ӻ���</span>)</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>д����</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_WriteEventSpace" maxlength="8" size="2" value="<%=htmlencode(Form_DEF_WriteEventSpace)%>"><span class=note>(����һЩд�붯���ļ���������޸�������֤�ȣ��������1-5��֮�䣬0��ʾ�����ƣ�����������ֵ������Ч��ֹ���д��������Ӷ��ﵽ����������Ӳ�̵�Ŀ��)</span></td>
		</tr>
		<tr>
			<td class=tdbox>������ʾ</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_EnableTreeView value=0<%If Form_DEF_EnableTreeView = 0 Then%> checked<%End If%>></td><td>��ֹʹ������</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_EnableTreeView value=1<%If Form_DEF_EnableTreeView = 1 Then%> checked<%End If%>></td><td>����ʹ������</td>
          		<td>&nbsp; (<span class=note>��ֹʹ�����ͽṹ���������</span>)</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>�ظ���¼</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_RepeatLoginTimeOut" maxlength="8" size="8" value="<%=htmlencode(Form_DEF_RepeatLoginTimeOut)%>"><span class=note>(�˲����������÷�ֹ�ظ���¼��һ�˺Ŷ����õ������ĳ�˵�¼�����������޷��ٽ��е�¼�����0��������߳�ʱ������Ч����������ֵΪ300-1800[5-30]���ӣ���λ��)</span></td>
		</tr>
		<tr>
			<td class=tdbox>FSO���</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_FSOString" maxlength="50" size="25" value="<%=htmlencode(Form_DEF_FSOString)%>">
			<br><span class=note>(�Զ���FSO��������ַ�����Ĭ��ΪScripting.FileSystemObject����ɿձ�ʾ����FSO)</span></td>
		</tr>
		<tr>
			<td class=tdbox>ʱ������</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_Now" maxlength="50" size="25" value="<%=htmlencode(Form_DEF_Now)%>"><br><span class=note>(��λ�����ӣ�����Ϊ������ʱ�����ָ�����ӣ�����Ϊ������ʱ���ȥָ������)</span></td>
		</tr>
		<tr>
			<td class=tdbox>�б�߶�</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_LineHeight" maxlength="8" size="8" value="<%=htmlencode(Form_DEF_LineHeight)%>"><span class=note>(��ʾ������,�Լ�����Ϣ���б���и߶�)</span></td>
		</tr>
		<tr>
			<td class=tdbox>ע���ļ�</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_RegisterFile" maxlength="50" size="25" value="<%=htmlencode(Form_DEF_RegisterFile)%>">
			<br><span class=note>(�Զ���ע���û�ʱʹ�õ��ļ���[UserĿ¼����]Ĭ��ΪNewUser.asp����ռ��и��ļ���Ȩ�����Զ�����������Ϊ.asp��չ��)</span></td>
		</tr>
		<tr>
			<td class=tdbox>�������</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_LimitTitle value=0<%If Form_DEF_LimitTitle = 0 Then%> checked<%End If%>></td><td>��ֹ</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_LimitTitle value=1<%If Form_DEF_LimitTitle = 1 Then%> checked<%End If%>></td><td>����</td>
          		<td>&nbsp; (<span class=note>����һЩ���ư�������ӱ����Ƿ�����������ʾ��������Ϊ������ֻ��ʾ�ܵ�����</span>)</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>���ظ�����Կ</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_DownKey" maxlength="50" size="20" value="<%=htmlencode(Form_DEF_DownKey)%>">
			<span class=note>(���ظ�����Ҫ����֤�ַ���)</span></td>
		</tr>
		<tr>
			<td class=tdbox>��վ�ײ���Ϣ</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_BottomInfo" maxlength="500" size="55" value="<%=htmlencode(Form_DEF_BottomInfo)%>">
			<span class=note>(��վ�ײ���Ϣ���,����ICP��Ϣ ֧��HTML)</span></td>
		</tr>
		<tr>
			<td class=tdbox>����ˢ�¼��</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_UpdateInterval" maxlength="8" size="8" value="<%=htmlencode(Form_DEF_UpdateInterval)%>"><span class=note>(��̳�����ļ�ˢ�µļ��ʱ�� ��λ��)</span></td>
		</tr>
		<tr>
			<td class=tdbox>��վĬ��������Ϣ</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_GBL_Description" maxlength="255" size="55" value="<%=htmlencode(Form_DEF_GBL_Description)%>">
			<span class=note>��վĬ�ϵ������ͷ����Description����</span></td>
		</tr>
		
		<tr>
			<td class=tdbox width=80>��������������</td>
			<td class=tdbox valign=top>
				<ul><%
				for n = 0 to Ubound(DEF_Sideparameter_String,1)
					If instr(DEF_Sideparameter_String(n),"<span") = 0 Then%>
					<li><span class="grayfont"><%
					If n < 9 Then Response.Write "0"
					Response.Write n+1%></span><input type="checkbox" class=fmchkbox name="SideLimit<%=n+1%>" value="1"<%
					If instr(DEF_Sideparameter_String(n),"<span") Then Response.Write " disabled=""disabled"""
					If GetBinarybit(Form_DEF_Sideparameter,n+1) = 1 Then
						Response.Write " checked>"
					Else
						Response.Write ">"
					End If%><%=DEF_Sideparameter_String(n)%></li>
					<%
					End If
				Next%></ul></td>
		</tr>
		
		<tr>
		<td class=tdbox>&nbsp;</td>
		<td class=tdbox>
			<input type=submit name=�ύ value=�ύ class=fmbtn>
			<input type=reset name=ȡ�� value=ȡ�� class=fmbtn>
		</td>
		</tr>
		</table>
		<%

End Function

Function GetDefaultValue

	Dim Rs
	Set Rs = LDExeCute(sql_select("Select SavePoints from LeadBBS_SiteInfo",1),0)
	If Rs.Eof Then
		Form_SavePoints = 0
	Else
		Form_SavePoints = cCur(Rs(0))
	End If
	Form_DEF_BBS_Name = DEF_BBS_Name
	Form_DEF_ManageDir = DEF_ManageDir
	Form_DEF_BBS_DarkColor = DEF_BBS_DarkColor
	Form_DEF_BBS_LightDarkColor = DEF_BBS_LightDarkColor
	Form_DEF_BBS_Color = DEF_BBS_Color
	Form_DEF_BBS_LightColor = DEF_BBS_LightColor
	Form_DEF_BBS_LightestColor = DEF_BBS_LightestColor
	Form_DEF_BBS_TableHeadColor = DEF_BBS_TableHeadColor

	Form_DEF_BBS_MaxLayer = DEF_BBS_MaxLayer
	Form_DEF_UsedDataBase = DEF_UsedDataBase
	Form_DEF_BBS_SearchMode = DEF_BBS_SearchMode

	Form_DEF_BBS_AnnouncePoints = DEF_BBS_AnnouncePoints
	Form_DEF_BBS_PrizeAnnouncePoints = DEF_BBS_PrizeAnnouncePoints
	Form_DEF_BBS_MakeGoodAnnouncePoints = DEF_BBS_MakeGoodAnnouncePoints
	Form_DEF_BBS_MaxTopAnnounce = DEF_BBS_MaxTopAnnounce
	Form_DEF_BBS_MaxAllTopAnnounce = DEF_BBS_MaxAllTopAnnounce
	Form_DEF_BBS_DisplayTopicLength = DEF_BBS_DisplayTopicLength
	Form_DEF_BBS_ScreenWidth = DEF_BBS_ScreenWidth
	Form_DEF_BBS_LeftTDWidth = DEF_BBS_LeftTDWidth
	Form_DEF_MasterCookies = DEF_MasterCookies
	Form_DEF_SiteNameString = DEF_SiteNameString
	Form_DEF_SupervisorUserName = DEF_SupervisorUserName
	Form_DEF_MaxTextLength = DEF_MaxTextLength

	Form_DEF_MaxListNum = DEF_MaxListNum
	Form_DEF_TopicContentMaxListNum = DEF_TopicContentMaxListNum
	Form_DEF_MaxJumpPageNum = DEF_MaxJumpPageNum
	Form_DEF_DisplayJumpPageNum = DEF_DisplayJumpPageNum
	Form_DEF_MaxBoardMastNum = DEF_MaxBoardMastNum
	Form_DEF_EnableUserHidden = DEF_EnableUserHidden
	Form_DEF_VOTE_MaxNum = DEF_VOTE_MaxNum
	Form_DEF_MaxLoginTimes = DEF_MaxLoginTimes
	Form_DEF_EnableUpload = DEF_EnableUpload
	Form_DEF_EnableGFL = DEF_EnableGFL
	Form_DEF_UserOnlineTimeOut = DEF_UserOnlineTimeOut
	Form_DEF_faceMaxNum = DEF_faceMaxNum
	Form_DEF_AllDefineFace = DEF_AllDefineFace
	Form_DEF_AllFaceMaxWidth = DEF_AllFaceMaxWidth
	Form_DEF_BBS_EmailMode = DEF_BBS_EmailMode
	Form_DEF_EnableAttestNumber = DEF_EnableAttestNumber
	Form_DEF_AttestNumberPoints = DEF_AttestNumberPoints
	Form_DEF_EnableUnderWrite = DEF_EnableUnderWrite
	Form_DEF_NeedOnlineTime = DEF_NeedOnlineTime
	Form_DEF_EnableForbidIP = DEF_EnableForbidIP
	Form_DEF_TopAdString = DEF_TopAdString
	Form_DEF_RestSpaceTime = DEF_RestSpaceTime
	Form_DEF_LoginSpaceTime = DEF_LoginSpaceTime
	Form_DEF_AccessDatabase = DEF_AccessDatabase
	Form_DEF_SiteHomeUrl = DEF_SiteHomeUrl
	Form_DEF_DefaultStyle = DEF_DefaultStyle
	Form_DEF_EnableFlashUBB = DEF_EnableFlashUBB
	Form_DEF_EnableImagesUBB = DEF_EnableImagesUBB
	Form_DEF_AnnounceFontSize = DEF_AnnounceFontSize
	Form_DEF_EditAnnounceDelay = DEF_EditAnnounceDelay
	Form_DEF_DisplayOnlineUser = DEF_DisplayOnlineUser
	Form_DEF_EnableSpecialTopic = DEF_EnableSpecialTopic
	Form_DEF_UBBiconNumber = DEF_UBBiconNumber
	Form_DEF_EnableDelAnnounce = DEF_EnableDelAnnounce
	Form_DEF_LimitTitle = DEF_LimitTitle
	Form_DEF_DownKey = DEF_DownKey	
	Form_DEF_UpdateInterval = DEF_UpdateInterval
	Form_DEF_BottomInfo = DEF_BottomInfo
	Form_DEF_GBL_Description = DEF_GBL_Description
	Form_DEF_Sideparameter = DEF_Sideparameter
	
	Dim N
	For n = 0 to Ubound(Form_DEF_PointsName)
		Form_DEF_PointsName(n) = DEF_PointsName(n)
	Next

	Form_DEF_EnableMakeTopAnc = DEF_EnableMakeTopAnc
	Form_DEF_EnableDatabaseCache = DEF_EnableDatabaseCache
	Form_DEF_WriteEventSpace = DEF_WriteEventSpace
	Form_DEF_EnableTreeView = DEF_EnableTreeView
	Form_DEF_EditAnnounceExpires = DEF_EditAnnounceExpires
	Form_DEF_RepeatLoginTimeOut = DEF_RepeatLoginTimeOut
	Form_DEF_FSOString = DEF_FSOString
	Form_DEF_Now = DateDiff("n",now,DEF_Now)
	Form_DEF_LineHeight = DEF_LineHeight
	Form_DEF_RegisterFile = DEF_RegisterFile

End Function

Function GetFormValue

	Form_DEF_ManageDir = Trim(Request.Form("Form_DEF_ManageDir"))
	Form_DEF_BBS_Name = Trim(Request.Form("Form_DEF_BBS_Name"))
	Form_DEF_BBS_DarkColor = Trim(Request.Form("Form_DEF_BBS_DarkColor"))
	Form_DEF_BBS_LightDarkColor = Trim(Request.Form("Form_DEF_BBS_LightDarkColor"))
	Form_DEF_BBS_Color = Trim(Request.Form("Form_DEF_BBS_Color"))
	Form_DEF_BBS_LightColor = Trim(Request.Form("Form_DEF_BBS_LightColor"))
	Form_DEF_BBS_LightestColor = Trim(Request.Form("Form_DEF_BBS_LightestColor"))
	Form_DEF_BBS_TableHeadColor = Trim(Request.Form("Form_DEF_BBS_TableHeadColor"))

	Form_DEF_BBS_MaxLayer = Trim(Request.Form("Form_DEF_BBS_MaxLayer"))
	Form_DEF_UsedDataBase = Trim(Request.Form("Form_DEF_UsedDataBase"))
	Form_DEF_BBS_SearchMode = Trim(Request.Form("Form_DEF_BBS_SearchMode"))

	Form_DEF_BBS_AnnouncePoints = Trim(Request.Form("Form_DEF_BBS_AnnouncePoints"))
	Form_DEF_BBS_PrizeAnnouncePoints = Trim(Request.Form("Form_DEF_BBS_PrizeAnnouncePoints"))
	Form_DEF_BBS_MakeGoodAnnouncePoints = Trim(Request.Form("Form_DEF_BBS_MakeGoodAnnouncePoints"))
	Form_DEF_BBS_MaxTopAnnounce = Trim(Request.Form("Form_DEF_BBS_MaxTopAnnounce"))
	Form_DEF_BBS_MaxAllTopAnnounce = Trim(Request.Form("Form_DEF_BBS_MaxAllTopAnnounce"))
	Form_DEF_BBS_DisplayTopicLength = Trim(Request.Form("Form_DEF_BBS_DisplayTopicLength"))
	Form_DEF_BBS_ScreenWidth = Trim(Request.Form("Form_DEF_BBS_ScreenWidth"))
	Form_DEF_BBS_LeftTDWidth = Trim(Request.Form("Form_DEF_BBS_LeftTDWidth"))
	Form_DEF_MasterCookies = Trim(Request.Form("Form_DEF_MasterCookies"))
	Form_DEF_SiteNameString = Trim(Request.Form("Form_DEF_SiteNameString"))
	Form_DEF_SupervisorUserName = Trim(Request.Form("Form_DEF_SupervisorUserName"))
	Form_DEF_MaxTextLength = Trim(Request.Form("Form_DEF_MaxTextLength"))

	Form_DEF_MaxListNum = Trim(Request.Form("Form_DEF_MaxListNum"))
	Form_DEF_TopicContentMaxListNum = Trim(Request.Form("Form_DEF_TopicContentMaxListNum"))
	Form_DEF_MaxJumpPageNum = Trim(Request.Form("Form_DEF_MaxJumpPageNum"))
	Form_DEF_DisplayJumpPageNum = Trim(Request.Form("Form_DEF_DisplayJumpPageNum"))
	Form_DEF_MaxBoardMastNum = Trim(Request.Form("Form_DEF_MaxBoardMastNum"))
	Form_DEF_EnableUserHidden = Trim(Request.Form("Form_DEF_EnableUserHidden"))
	Form_DEF_VOTE_MaxNum = Trim(Request.Form("Form_DEF_VOTE_MaxNum"))
	Form_DEF_MaxLoginTimes = Trim(Request.Form("Form_DEF_MaxLoginTimes"))
	Form_DEF_EnableUpload = Trim(Request.Form("Form_DEF_EnableUpload"))
	Form_DEF_EnableGFL = Trim(Request.Form("Form_DEF_EnableGFL"))
	Form_DEF_UserOnlineTimeOut = Trim(Request.Form("Form_DEF_UserOnlineTimeOut"))
	Form_DEF_faceMaxNum = Trim(Request.Form("Form_DEF_faceMaxNum"))
	Form_DEF_AllDefineFace = Trim(Request.Form("Form_DEF_AllDefineFace"))
	Form_DEF_AllFaceMaxWidth = Trim(Request.Form("Form_DEF_AllFaceMaxWidth"))
	Form_DEF_BBS_EmailMode = Trim(Request.Form("Form_DEF_BBS_EmailMode"))
	Form_DEF_EnableAttestNumber = Trim(Request.Form("Form_DEF_EnableAttestNumber"))
	Form_DEF_AttestNumberPoints = Trim(Request.Form("Form_DEF_AttestNumberPoints"))
	Form_DEF_EnableUnderWrite = Trim(Request.Form("Form_DEF_EnableUnderWrite"))
	Form_DEF_NeedOnlineTime = Trim(Request.Form("Form_DEF_NeedOnlineTime"))
	Form_DEF_EnableForbidIP = Trim(Request.Form("Form_DEF_EnableForbidIP"))
	Form_DEF_TopAdString = Trim(Request.Form("Form_DEF_TopAdString"))
	Form_DEF_RestSpaceTime = Trim(Request.Form("Form_DEF_RestSpaceTime"))
	Form_DEF_LoginSpaceTime = Trim(Request.Form("Form_DEF_LoginSpaceTime"))
	Form_DEF_AccessDatabase = Trim(Request.Form("Form_DEF_AccessDatabase"))
	Form_DEF_SiteHomeUrl = Trim(Request.Form("Form_DEF_SiteHomeUrl"))
	Form_DEF_DefaultStyle = Trim(Request.Form("Form_DEF_DefaultStyle"))	
	Form_DEF_EnableFlashUBB = Trim(Request.Form("Form_DEF_EnableFlashUBB"))	
	Form_DEF_EnableImagesUBB = Trim(Request.Form("Form_DEF_EnableImagesUBB"))
	Form_DEF_AnnounceFontSize = Trim(Request.Form("Form_DEF_AnnounceFontSize"))
	Form_DEF_EditAnnounceDelay = Trim(Request.Form("Form_DEF_EditAnnounceDelay"))
	Form_DEF_DisplayOnlineUser = Trim(Request.Form("Form_DEF_DisplayOnlineUser"))
	Form_DEF_EnableSpecialTopic = Trim(Request.Form("Form_DEF_EnableSpecialTopic"))
	Form_DEF_UBBiconNumber = Trim(Request.Form("Form_DEF_UBBiconNumber"))
	Form_DEF_EnableDelAnnounce = Trim(Request.Form("Form_DEF_EnableDelAnnounce"))
	Form_DEF_LimitTitle = Trim(Request.Form("Form_DEF_LimitTitle"))
	Form_DEF_DownKey = Left(Trim(Request.Form("Form_DEF_DownKey")),50)	
	Form_DEF_UpdateInterval = Left(Trim(Request.Form("Form_DEF_UpdateInterval")),50)
	Form_DEF_BottomInfo = Left(Request.Form("Form_DEF_BottomInfo"),500)
	Form_DEF_GBL_Description = Left(Trim(Request.Form("Form_DEF_GBL_Description")),255)
	
	Dim N
	For n = 0 to Ubound(DEF_PointsName)
		Form_DEF_PointsName(n) = Trim(Request.Form("Form_DEF_PointsName" & N))
	Next
	
	
	Dim Temp2,TempN
	Form_DEF_Sideparameter = 0
	Temp2 = 1
	For TempN = 0 to Ubound(DEF_Sideparameter_String,1)
		N = Request("SideLimit" & TempN+1)
		If N <> "1" Then N = "0"
		If N = "1" Then Form_DEF_Sideparameter = Form_DEF_Sideparameter+cCur(Temp2)
		Temp2 = Temp2*2
	Next

	Form_DEF_EnableMakeTopAnc = Trim(Request.Form("Form_DEF_EnableMakeTopAnc"))
	Form_DEF_EnableDatabaseCache = Trim(Request.Form("Form_DEF_EnableDatabaseCache"))
	Form_DEF_WriteEventSpace = Trim(Request.Form("Form_DEF_WriteEventSpace"))
	Form_DEF_EnableTreeView = Trim(Request.Form("Form_DEF_EnableTreeView"))
	Form_DEF_EditAnnounceExpires = Trim(Request.Form("Form_DEF_EditAnnounceExpires"))
	Form_DEF_RepeatLoginTimeOut = Trim(Request.Form("Form_DEF_RepeatLoginTimeOut"))
	Form_DEF_FSOString = Trim(Request.Form("Form_DEF_FSOString"))
	Form_DEF_Now = Trim(Request.Form("Form_DEF_Now"))
	Form_SavePoints = Left(Trim(Request.Form("Form_SavePoints")),14)
	Form_DEF_LineHeight = Trim(Request.Form("Form_DEF_LineHeight"))
	Form_DEF_RegisterFile = Trim(Request.Form("Form_DEF_RegisterFile"))

	If isNumeric(Form_SavePoints) = 0 Then Form_SavePoints = 0
	Form_SavePoints = Fix(cCur(Form_SavePoints))
	If Form_SavePoints < 0 Then Form_SavePoints = 0

	If inStr(Form_DEF_ManageDir,"%") Then GBL_CHK_TempStr = "����Ŀ¼���ܰ����ٷֺ�<br>" & VbCrLf
	If inStr(Form_DEF_BBS_Name,"%") Then GBL_CHK_TempStr = "��̳���Ʋ��ܰ����ٷֺ�<br>" & VbCrLf
	If inStr(Form_DEF_SiteHomeUrl,"%") Then GBL_CHK_TempStr = "��վ��ҳ���ܰ����ٷֺ�<br>" & VbCrLf
	If inStr(Form_DEF_BBS_DarkColor,"%") Then GBL_CHK_TempStr = "�� �� ɫ���ܰ����ٷֺ�<br>" & VbCrLf
	If inStr(Form_DEF_BBS_LightDarkColor,"%") Then GBL_CHK_TempStr = "�� �� ɫ���ܰ����ٷֺ�<br>" & VbCrLf
	If inStr(Form_DEF_BBS_Color,"%") Then GBL_CHK_TempStr = "��̳��ɫ���ܰ����ٷֺ�<br>" & VbCrLf
	If inStr(Form_DEF_BBS_LightColor,"%") Then GBL_CHK_TempStr = "�� �� ɫ���ܰ����ٷֺ�<br>" & VbCrLf
	If inStr(Form_DEF_BBS_LightestColor,"%") Then GBL_CHK_TempStr = "�� �� ɫ���ܰ����ٷֺ�<br>" & VbCrLf
	If inStr(Form_DEF_BBS_TableHeadColor,"%") Then GBL_CHK_TempStr = "���ͷɫ���ܰ����ٷֺ�<br>" & VbCrLf
	If isNumeric(Form_DEF_BBS_MaxLayer) = 0 Then GBL_CHK_TempStr = "�ظ���������Ϊ����<br>" & VbCrLf

	If isNumeric(Form_DEF_UsedDataBase) = 0 Then GBL_CHK_TempStr = "�� �� �����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_BBS_SearchMode) = 0 Then GBL_CHK_TempStr = "����ģʽ����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_BBS_AnnouncePoints) = 0 Then GBL_CHK_TempStr = "������������Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_BBS_PrizeAnnouncePoints) = 0 Then GBL_CHK_TempStr = "ɾ���ͷ�����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_BBS_MakeGoodAnnouncePoints) = 0 Then GBL_CHK_TempStr = "������������Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_BBS_MaxTopAnnounce) = 0 Then GBL_CHK_TempStr = "��ඥ������Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_BBS_MaxAllTopAnnounce) = 0 Then GBL_CHK_TempStr = "����ܹ̱���Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_BBS_DisplayTopicLength) = 0 Then GBL_CHK_TempStr = "���ⳤ�ȱ���Ϊ����<br>" & VbCrLf
	If inStr(Form_DEF_BBS_LeftTDWidth,"%") Then GBL_CHK_TempStr = "��̳��Ȳ��ܰ����ٷֺ�<br>" & VbCrLf
	If inStr(Form_DEF_MasterCookies,"%") Then GBL_CHK_TempStr = "Cookies���ܰ����ٷֺ�<br>" & VbCrLf
	If inStr(Form_DEF_SiteNameString,"%") Then GBL_CHK_TempStr = "��վ���Ʋ��ܰ����ٷֺ�<br>" & VbCrLf
	If inStr(Form_DEF_SupervisorUserName,"""") or inStr(Form_DEF_SupervisorUserName,"%") Then GBL_CHK_TempStr = "�� �� Ա���ܰ��������Ż�ٷֺ�<br>" & VbCrLf
	If isNumeric(Form_DEF_MaxTextLength) = 0 Then GBL_CHK_TempStr = "���ݳ��ȱ���Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_MaxListNum) = 0 Then GBL_CHK_TempStr = "��ʾ��¼��������Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_TopicContentMaxListNum) = 0 Then GBL_CHK_TempStr = "��ʾ������������Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_MaxJumpPageNum) = 0 Then GBL_CHK_TempStr = "��תҳ����������Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_DisplayJumpPageNum) = 0 Then GBL_CHK_TempStr = "��ʾ��ת��������Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_MaxBoardMastNum) = 0 Then GBL_CHK_TempStr = "����������������Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_EnableUserHidden) = 0 Then GBL_CHK_TempStr = "�������ñ���Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_VOTE_MaxNum) = 0 Then GBL_CHK_TempStr = "ͶƱ��Ŀ����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_MaxLoginTimes) = 0 Then GBL_CHK_TempStr = "��¼��������Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_RestSpaceTime) = 0 Then GBL_CHK_TempStr = "�����������Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_LoginSpaceTime) = 0 Then GBL_CHK_TempStr = "��¼�������Ϊ����<br>" & VbCrLf

	If isNumeric(Form_DEF_EnableUpload) = 0 Then GBL_CHK_TempStr = "�ϴ�Ȩ�ޱ���Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_EnableGFL) = 0 Then GBL_CHK_TempStr = "ͼ������Ƿ��������Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_UserOnlineTimeOut) = 0 Then GBL_CHK_TempStr = "���߳�ʱ����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_faceMaxNum) = 0 Then GBL_CHK_TempStr = "ͷ���������Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_AllDefineFace) = 0 Then GBL_CHK_TempStr = "�Զ�ͷ�����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_AllFaceMaxWidth) = 0 Then GBL_CHK_TempStr = "ͷ���С����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_BBS_EmailMode) = 0 Then GBL_CHK_TempStr = "�ʼ����ñ���Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_EnableAttestNumber) = 0 Then GBL_CHK_TempStr = "�� ֤ ����ʾ��ʽ����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_AttestNumberPoints) = 0 Then GBL_CHK_TempStr = "��֤�룲����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_EnableUnderWrite) = 0 Then GBL_CHK_TempStr = "ǩ��������ʾ��ʽ����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_NeedOnlineTime) = 0 Then GBL_CHK_TempStr = "����ʱ�����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_EnableForbidIP) = 0 Then GBL_CHK_TempStr = "�ɣ����α���Ϊ����<br>" & VbCrLf
	If inStr(Form_DEF_TopAdString,"%") Then GBL_CHK_TempStr = "������治�ܰ����ٷֺ�<br>" & VbCrLf
	If inStr(Form_DEF_AccessDatabase,"%") Then GBL_CHK_TempStr = "�� �� �����Ӳ��ܰ����ٷֺ�<br>" & VbCrLf
	If isNumeric(Form_DEF_DefaultStyle) = 0 Then GBL_CHK_TempStr = "Ĭ�Ϸ�����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_EnableFlashUBB) = 0 Then GBL_CHK_TempStr = "�� ý �����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_EnableImagesUBB) = 0 Then GBL_CHK_TempStr = "����ͼƬ����Ϊ����<br>" & VbCrLf
	If inStr(Form_DEF_AnnounceFontSize,"%") Then GBL_CHK_TempStr = "�������岻�ܰ����ٷֺ�<br>" & VbCrLf
	If isNumeric(Form_DEF_EditAnnounceDelay) = 0 Then GBL_CHK_TempStr = "�༭�������Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_DisplayOnlineUser) = 0 Then GBL_CHK_TempStr = "���߻�Ա��ʾ��ʽ����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_EnableSpecialTopic) = 0 Then GBL_CHK_TempStr = "�������ӱ���Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_UBBiconNumber) = 0 Then GBL_CHK_TempStr = "��������������Ϊ����<br>" & VbCrLf
	Form_DEF_UBBiconNumber = Fix(cCur(Form_DEF_UBBiconNumber))
	If Form_DEF_UBBiconNumber > 9999 Then Form_DEF_UBBiconNumber = 9999
	If isNumeric(Form_DEF_EnableDelAnnounce) = 0 Then GBL_CHK_TempStr = "�� �� վ�Ƿ��������Ϊ����<br>" & VbCrLf
	For n = 0 to Ubound(DEF_PointsName)
		If inStr(Form_DEF_PointsName(n),"%") Then
			GBL_CHK_TempStr = "��" & N & "�����ƶ����ﲻ�ܰ����ٷֺ�<br>" & VbCrLf
		End If
	Next
	If isNumeric(Form_DEF_EnableMakeTopAnc) = 0 Then GBL_CHK_TempStr = "�ظ���������Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_EnableDatabaseCache) = 0 Then GBL_CHK_TempStr = "�ģ»������Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_WriteEventSpace) = 0 Then GBL_CHK_TempStr = "д��������Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_EnableTreeView) = 0 Then GBL_CHK_TempStr = "������ʾ����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_EditAnnounceExpires) = 0 Then GBL_CHK_TempStr = "�༭���ڱ���Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_RepeatLoginTimeOut) = 0 Then GBL_CHK_TempStr = "�ظ���¼ʱ�����Ϊ����<br>" & VbCrLf
	If inStr(Form_DEF_FSOString,"%") Then GBL_CHK_TempStr = "FSO������Ʋ��ܰ����ٷֺ�<br>" & VbCrLf
	If isNumeric(Form_DEF_Now) = 0 Then GBL_CHK_TempStr = "ʱ�����ñ���Ϊ����<br>" & VbCrLf	
	If isNumeric(Form_DEF_LineHeight) = 0 Then GBL_CHK_TempStr = "�б�߶ȱ���Ϊ����<br>" & VbCrLf
	If isNumeric(DEF_UpdateInterval) = 0 Then GBL_CHK_TempStr = "����ˢ�¼������Ϊ����<br>" & VbCrLf

	Form_DEF_RegisterFile = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Form_DEF_RegisterFile,"""",""),"?",""),"/",""),"\",""),"*",""),":",""),"<",""),">",""),"|","")
	If LCase(Right(Form_DEF_RegisterFile,4)) <> ".asp" Then GBL_CHK_TempStr = "ע���ļ����ƴ��󣬱�����.asp��Ϊ��չ��!<br>" & VbCrLf
	If isNumeric(Form_DEF_LimitTitle) = 0 Then GBL_CHK_TempStr = "������ܱ���Ϊ����<br>" & VbCrLf
	If inStr(DEF_DownKey,"""") or inStr(DEF_DownKey,"%") Then GBL_CHK_TempStr = "���ظ�����Կ���ܰ����ٷֺ�<br>" & VbCrLf
	If isNumeric(DEF_UpdateInterval) = 0 Then GBL_CHK_TempStr = "����ˢ�¼������Ϊ����<br>" & VbCrLf
	If inStr(Form_DEF_BottomInfo,"%") Then GBL_CHK_TempStr = "�ײ���Ϣ���ܰ����ٷֺ�<br>" & VbCrLf
	If inStr(Form_DEF_GBL_Description,"%") Then GBL_CHK_TempStr = "ͷ��Description��Ϣ���ܰ����ٷֺ�<br>" & VbCrLf

End Function

Function ReplaceStr(str)

	ReplaceStr = Replace(Str,"""","""""")

End Function

Function MakeDataBaseLinkFile

	Dim n,TempStr
	TempStr = ""
	TempStr = TempStr & chr(60) & "%@ LANGUAGE=VBScript CodePage=936%" & chr(62) & VbCrLf
	TempStr = TempStr & chr(60) & "%Option Explicit" & VbCrLf
	TempStr = TempStr & "Response.Charset = ""gb2312""" & VbCrLf
	TempStr = TempStr & "Session.CodePage=936" & VbCrLf
	TempStr = TempStr & "Response.Buffer = True" & VbCrLf
	TempStr = TempStr & "Const DEF_ManageDir = """ & Form_DEF_ManageDir & """" & VbCrLf
	TempStr = TempStr & VbCrLf

	TempStr = TempStr & "If isNumeric(application(DEF_MasterCookies & " & chr(34) & "SiteEnableFlagzoieiu" & chr(34) & ")) = 0 Then" & VbCrLf
	TempStr = TempStr & "	Application.Lock" & VbCrLf
	TempStr = TempStr & "	application(DEF_MasterCookies & " & chr(34) & "SiteEnableFlagzoieiu" & chr(34) & ") = 1" & VbCrLf
	TempStr = TempStr & "	Application.UnLock" & VbCrLf
	TempStr = TempStr & "End If"  &VbCrLf

	TempStr = TempStr & "If application(DEF_MasterCookies & " & chr(34) & "SiteEnableFlagzoieiu" & chr(34) & ") = 0 and application(DEF_MasterCookies & " & chr(34) & "SiteDisbleWhyszoieiu" & chr(34) & ")<>" & chr(34) & chr(34) & " and inStr(Replace(Lcase(Request.ServerVariables(" & chr(34) & "URL" & chr(34) & "))," & chr(34) & "\" & chr(34) & "," & chr(34) & "/" & chr(34) & ")," & chr(34) & "/"" & DEF_ManageDir & ""/" & chr(34) & ") = 0 Then" & VbCrLf
	TempStr = TempStr & "	Response.Write application(DEF_MasterCookies & " & chr(34) & "SiteDisbleWhyszoieiu" & chr(34) & ")" & VbCrLf
	TempStr = TempStr & "	Response.End" & VbCrLf
	TempStr = TempStr & "End If" & VbCrLf
	TempStr = TempStr & VbCrLf
	TempStr = TempStr & "Dim DEF_BBS_HomeUrl,DEF_SiteHomeUrl" & VbCrLf
	TempStr = TempStr & "const DEF_BBS_Name=" & Chr(34) & ReplaceStr(Form_DEF_BBS_Name) & Chr(34) & VbCrLf

	TempStr = TempStr & "DEF_BBS_HomeUrl = " & Chr(34) & Chr(34) & VbCrLf
	TempStr = TempStr & "DEF_SiteHomeUrl = " & Chr(34) & ReplaceStr(Form_DEF_SiteHomeUrl) & Chr(34) & VbCrLf
	TempStr = TempStr & "const DEF_BBS_DarkColor = " & Chr(34) & Form_DEF_BBS_DarkColor & Chr(34) & VbCrLf
	TempStr = TempStr & "const DEF_BBS_LightDarkColor = " & Chr(34) & Form_DEF_BBS_LightDarkColor & Chr(34) & VbCrLf
	TempStr = TempStr & "const DEF_BBS_Color = " & Chr(34) & Form_DEF_BBS_Color & Chr(34) & VbCrLf
	TempStr = TempStr & "const DEF_BBS_LightColor = " & Chr(34) & Form_DEF_BBS_LightColor & Chr(34) & VbCrLf
	TempStr = TempStr & "const DEF_BBS_LightestColor = " & Chr(34) & Form_DEF_BBS_LightestColor & chr(34) & VbCrLf
	TempStr = TempStr & "const DEF_BBS_TableHeadColor = " & Chr(34) & Form_DEF_BBS_TableHeadColor & Chr(34) & VbCrLf
	TempStr = TempStr & "const DEF_BBS_MaxLayer = " & Form_DEF_BBS_MaxLayer & VbCrLf
	TempStr = TempStr & "const DEF_UsedDataBase = " & Form_DEF_UsedDataBase & VbCrLf
	TempStr = TempStr & "const DEF_BBS_SearchMode = " & Form_DEF_BBS_SearchMode & VbCrLf

	TempStr = TempStr & "const DEF_BBS_TOPMinID = 99999999990000" & VbCrLf
	TempStr = TempStr & "const DEF_BBS_AnnouncePoints = " & Form_DEF_BBS_AnnouncePoints & VbCrLf
	TempStr = TempStr & "const DEF_BBS_PrizeAnnouncePoints = " & Form_DEF_BBS_PrizeAnnouncePoints & VbCrLf
	TempStr = TempStr & "const DEF_BBS_MakeGoodAnnouncePoints = " & Form_DEF_BBS_MakeGoodAnnouncePoints & VbCrLf
	TempStr = TempStr & "const DEF_BBS_MaxTopAnnounce = " & Form_DEF_BBS_MaxTopAnnounce & VbCrLf
	TempStr = TempStr & "const DEF_BBS_MaxAllTopAnnounce = " & Form_DEF_BBS_MaxAllTopAnnounce & VbCrLf

	TempStr = TempStr & "Dim DEF_BBS_DisplayTopicLength,DEF_BBS_ScreenWidth" & VbCrLf
	TempStr = TempStr & "DEF_BBS_DisplayTopicLength = " & Form_DEF_BBS_DisplayTopicLength & VbCrLf

	TempStr = TempStr & "DEF_BBS_ScreenWidth = " & Chr(34) & ReplaceStr(Form_DEF_BBS_ScreenWidth) & Chr(34) & VbCrLf
	TempStr = TempStr & "const DEF_BBS_LeftTDWidth = " & Chr(34) & ReplaceStr(Form_DEF_BBS_LeftTDWidth) & Chr(34) & VbCrLf

	TempStr = TempStr & "const DEF_MasterCookies = " & Chr(34) & ReplaceStr(Form_DEF_MasterCookies) & Chr(34) & VbCrLf
	TempStr = TempStr & "const DEF_SiteNameString = " & Chr(34) & ReplaceStr(Form_DEF_SiteNameString) & Chr(34) & VbCrLf
	TempStr = TempStr & "const DEF_SupervisorUserName = " & Chr(34) & ReplaceStr(Form_DEF_SupervisorUserName) & Chr(34) & VbCrLf
	TempStr = TempStr & "const DEF_MaxTextLength = " & Form_DEF_MaxTextLength & VbCrLf

	TempStr = TempStr & "Dim DEF_MaxListNum" & VbCrLf
	TempStr = TempStr & "DEF_MaxListNum = " & Form_DEF_MaxListNum & VbCrLf
	TempStr = TempStr & "Const DEF_TopicContentMaxListNum = " & Form_DEF_TopicContentMaxListNum & VbCrLf
	TempStr = TempStr & "Const DEF_MaxJumpPageNum = " & Form_DEF_MaxJumpPageNum & VbCrLf
	TempStr = TempStr & "Const DEF_DisplayJumpPageNum = " & Form_DEF_DisplayJumpPageNum & VbCrLf

	TempStr = TempStr & "const DEF_MaxBoardMastNum = " & Form_DEF_MaxBoardMastNum & VbCrLf

	TempStr = TempStr & "const DEF_EnableUserHidden = " & Form_DEF_EnableUserHidden & VbCrLf
	TempStr = TempStr & "const DEF_VOTE_MaxNum = " & Form_DEF_VOTE_MaxNum & VbCrLf

	TempStr = TempStr & "const DEF_MaxLoginTimes = " & Form_DEF_MaxLoginTimes & VbCrLf
	TempStr = TempStr & "const DEF_RestSpaceTime = " & Form_DEF_RestSpaceTime & VbCrLf
	TempStr = TempStr & "const DEF_LoginSpaceTime = " & Form_DEF_LoginSpaceTime & VbCrLf

	TempStr = TempStr & "const DEF_EnableUpload = " & Form_DEF_EnableUpload & VbCrLf
	TempStr = TempStr & "const DEF_EnableGFL = " & Form_DEF_EnableGFL & VbCrLf
	TempStr = TempStr & "const DEF_UserOnlineTimeOut = " & Form_DEF_UserOnlineTimeOut & VbCrLf
	TempStr = TempStr & "const DEF_faceMaxNum = " & Form_DEF_faceMaxNum & VbCrLf
	TempStr = TempStr & "const DEF_AllDefineFace = " & Form_DEF_AllDefineFace & VbCrLf
	TempStr = TempStr & "const DEF_AllFaceMaxWidth = " & Form_DEF_AllFaceMaxWidth & VbCrLf

	TempStr = TempStr & "const DEF_BBS_EmailMode = " & Form_DEF_BBS_EmailMode & VbCrLf
	TempStr = TempStr & "Const DEF_EnableAttestNumber = " & Form_DEF_EnableAttestNumber & VbCrLf
	TempStr = TempStr & "Const DEF_AttestNumberPoints = " & Form_DEF_AttestNumberPoints & VbCrLf

	TempStr = TempStr & "Dim DEF_BoardStyleString,DEF_BoardStyleStringNum" & VbCrLf

	TempStr = TempStr & "DEF_BoardStyleString = Array("
	For n = 0 to DEF_BoardStyleStringNum
		If n = 0 Then
			TempStr = TempStr & """" & DEF_BoardStyleString(n) & """"
		Else
			TempStr = TempStr & ",""" & DEF_BoardStyleString(n) & """"
		End If
	Next
	TempStr = TempStr & ")" & VbCrLf
	TempStr = TempStr & "DEF_BoardStyleStringNum = Ubound(DEF_BoardStyleString,1)" & VbCrLf

	TempStr = TempStr & "Const DEF_EnableUnderWrite = " & Form_DEF_EnableUnderWrite & VbCrLf
	TempStr = TempStr & "Const DEF_NeedOnlineTime = " & Form_DEF_NeedOnlineTime & VbCrLf
	TempStr = TempStr & "Const DEF_EnableForbidIP = " & Form_DEF_EnableForbidIP & VbCrLf

	TempStr = TempStr & "Const DEF_TopAdString = " & Chr(34) & ReplaceStr(Form_DEF_TopAdString) & Chr(34) & VbCrLf
	TempStr = TempStr & "Const DEF_AccessDatabase = " & Chr(34) & ReplaceStr(Form_DEF_AccessDatabase) & Chr(34) & VbCrLf

	TempStr = TempStr & "Const DEF_DefaultStyle = " & Form_DEF_DefaultStyle & VbCrLf
	TempStr = TempStr & "Const DEF_EnableFlashUBB = " & Form_DEF_EnableFlashUBB & VbCrLf
	TempStr = TempStr & "Const DEF_EnableImagesUBB = " & Form_DEF_EnableImagesUBB & VbCrLf
	TempStr = TempStr & "Const DEF_AnnounceFontSize = " & Chr(34) & ReplaceStr(Form_DEF_AnnounceFontSize) & Chr(34) & VbCrLf
	TempStr = TempStr & "Const DEF_EditAnnounceDelay = " & Form_DEF_EditAnnounceDelay & VbCrLf
	TempStr = TempStr & "Const DEF_DisplayOnlineUser = " & Form_DEF_DisplayOnlineUser & VbCrLf
	TempStr = TempStr & "Const DEF_EnableSpecialTopic = " & Form_DEF_EnableSpecialTopic & VbCrLf
	TempStr = TempStr & "Const DEF_UBBiconNumber = " & Form_DEF_UBBiconNumber & VbCrLf
	TempStr = TempStr & "Const DEF_EnableDelAnnounce = " & Form_DEF_EnableDelAnnounce & VbCrLf
	TempStr = TempStr & "Dim DEF_PointsName" & VbCrLf
	TempStr = TempStr & "DEF_PointsName = Array("
	For n = 0 to Ubound(DEF_PointsName)
		If n = 0 Then
			TempStr = TempStr & """" & Form_DEF_PointsName(n) & """"
		Else
			TempStr = TempStr & ",""" & Form_DEF_PointsName(n) & """"
		End If
	Next
	TempStr = TempStr & ")" & VbCrLf
	TempStr = TempStr & "Const DEF_EnableMakeTopAnc = " & Form_DEF_EnableMakeTopAnc & VbCrLf
	TempStr = TempStr & "Const DEF_EnableDatabaseCache = " & Form_DEF_EnableDatabaseCache & VbCrLf
	TempStr = TempStr & "Const DEF_WriteEventSpace = " & Form_DEF_WriteEventSpace & VbCrLf
	TempStr = TempStr & "Const DEF_EnableTreeView = " & Form_DEF_EnableTreeView & VbCrLf
	TempStr = TempStr & "Const DEF_EditAnnounceExpires = " & Form_DEF_EditAnnounceExpires & VbCrLf
	TempStr = TempStr & "Const DEF_RepeatLoginTimeOut = " & Form_DEF_RepeatLoginTimeOut & VbCrLf
	TempStr = TempStr & "Const DEF_FSOString = " & Chr(34) & ReplaceStr(Form_DEF_FSOString) & Chr(34) & VbCrLf
	TempStr = TempStr & "Dim DEF_Now,DEF_Version" & VbCrLf
	If Form_DEF_Now = 0 Then
		TempStr = TempStr & "DEF_Now = now" & VbCrLf
	Else
		TempStr = TempStr & "DEF_Now = DateAdd(""n""," & Form_DEF_Now & ",now)" & VbCrLf
	End If
	TempStr = TempStr & "DEF_Version = " & Chr(34) & ReplaceStr(DEF_Version) & chr(34) & VbCrLf
	TempStr = TempStr & "Const DEF_LineHeight = " & Form_DEF_LineHeight & VbCrLf
	TempStr = TempStr & "Const DEF_RegisterFile = " & Chr(34) & ReplaceStr(Form_DEF_RegisterFile) & Chr(34) & VbCrLf
	TempStr = TempStr & "Const DEF_LimitTitle = " & Form_DEF_LimitTitle & VbCrLf

	TempStr = TempStr & "Const DEF_DownKey = " & Chr(34) & ReplaceStr(Form_DEF_DownKey) & Chr(34) & VbCrLf
	TempStr = TempStr & "Const DEF_UpdateInterval = " & Form_DEF_UpdateInterval & VbCrLf
	TempStr = TempStr & "Const DEF_BottomInfo = " & Chr(34) & ReplaceStr(Form_DEF_BottomInfo) & Chr(34) & VbCrLf
	TempStr = TempStr & "Dim DEF_GBL_Description" & VbCrLf
	TempStr = TempStr & "DEF_GBL_Description = " & Chr(34) & ReplaceStr(Form_DEF_GBL_Description) & Chr(34) & VbCrLf
	TempStr = TempStr & "Const DEF_Sideparameter = " & ReplaceStr(Form_DEF_Sideparameter) & VbCrLf

	TempStr = TempStr & "%" & chr(62) & VbCrLf

	ADODB_SaveToFile TempStr,"../../inc/BBSSetup.asp"

	CALL Update_InsertSetupRID(1051,"inc/BBSSetup.asp",0,TempStr," and ClassNum=" & 0)
	
	If GBL_CHK_TempStr = "" Then
		Response.Write "<br><span class=greenfont>2.�ɹ�������ã�</span>"
	Else
		%><%=GBL_CHK_TempStr%><br>��������֧������д���ļ����ܣ���ʹ��FTP�ȹ��ܣ���<span Class=redfont>inc/BBSSetup.asp</span>�ļ��滻�ɿ�������(ע�ⱸ��)<p>
		<textarea name="fileContent" cols="80" rows="30" class=fmtxtra><%=Server.htmlencode(TempStr)%></textarea><%
		GBL_CHK_TempStr = ""
	End If
	CALL LDExeCute("Update LeadBBS_SiteInfo Set SavePoints=" & Form_SavePoints,1)
	RennameRegisterFile DEF_RegisterFile,Form_DEF_RegisterFile

End Function

Function RennameRegisterFile(path,NewPath)

	If DEF_FSOString = "" or path = NewPath Then Exit Function
	on error resume next
	Dim fs
	Set fs = Server.CreateObject(DEF_FSOString)
	If err <> 0 Then
		Err.Clear
		Set fs = Nothing
		Response.Write "<p>��������֧��FSO��Ӳ���ϵ�ע���ļ���δ���ģ�"
		RennameRegisterFile = 0
		Exit Function
	End If

	If Not fs.FileExists(Server.Mappath(DEF_BBS_HomeUrl & "User/" & path)) Then
		Set fs = Nothing
		Response.Write "<p>Ӳ���ϵ�ԭ���ļ�" & path & "�����ڣ�������ע���ļ���ʧ�ܣ����¼ftp��飡"
		RennameRegisterFile = 0
		Exit Function
	End If

	If fs.FileExists(Server.Mappath(DEF_BBS_HomeUrl & "User/" & NewPath)) Then
		Set fs = Nothing
		Response.Write "<p>Ӳ���ϵ�Ŀ�������ļ�" & NewPath & "�Ѿ����ڣ�������ע���ļ���ʧ�ܣ����¼ftp��飬��ѡ�������ļ�����"
		RennameRegisterFile = 0
		Exit Function
	End If
	
	fs.MoveFile Server.Mappath(DEF_BBS_HomeUrl & "User/" & path),Server.Mappath(DEF_BBS_HomeUrl & "User/" & NewPath)
	If err <> 0 Then
		Err.Clear
		Set fs = Nothing
		Response.Write "<p>Ӳ���ϵ�ע���ļ���������ʧ�ܣ����¼ftp�ֶ����ģ�"
		RennameRegisterFile = 0
		Exit Function
	End If
	Set fs = Nothing
         
End Function%>