<!-- #include file=../inc/BBSsetup.asp -->
<!-- #include file=../inc/Upload_Setup.asp -->
<!-- #include file=../inc/Board_Popfun.asp -->
<!-- #include file=../User/inc/UserTopic.asp -->
<!-- #include file=inc/Upload_fun.asp -->
<%
DEF_BBS_HomeUrl = "../"

Sub Main

	'GBL_CHK_PWdFlag = 0
	initDatabase
	GBL_CHK_TempStr = ""
	BBS_SiteHead DEF_SiteNameString & " - ��̳����",0,"<span class=navigate_string_step>��̳����</span>"
	UpdateOnlineUserAtInfo 0,"��̳����"
	If GBL_UserID < 1 Then
		GBL_CHK_TempStr = "ֻ��ע���û����ܲ鿴��̳������"
	Else
		If GBL_CHK_OnlineTime < DEF_NeedOnlineTime Then GBL_CHK_TempStr = "����ʱ��(" & DEF_PointsName(4) & ")���㣬ֻ������ʱ�䳬��" & DEF_NeedOnlineTime & "����û�����ʹ�ô˹��ܡ�"
	End If
	
	'If GetBinarybit(GBL_CHK_UserLimit,11) = 1 and (GetBinarybit(GBL_CHK_UserLimit,10) = 1 or GetBinarybit(GBL_CHK_UserLimit,8) = 1) Then
	'If CheckSupervisorUserName = 1 or (GetBinarybit(GBL_CHK_UserLimit,10) = 1 or GetBinarybit(GBL_CHK_UserLimit,8) = 1) or GetBinarybit(GBL_CHK_UserLimit,14) = 1 Then
	'Else
		'GBL_CHK_TempStr = "ȫ�������鿴����Ϊֻ��������鿴��"
	'End If

	UserTopicTopInfo("forum")
	If GBL_CHK_TempStr = "" Then
		CALL Upload_List(0,0,"../Search/UploadList.asp?ttsID=3",1)
	Else
		Response.Write "<div class=alert>" & GBL_CHK_TempStr & "</div>"
	End If
	closeDataBase
	UserTopicBottomInfo
	SiteBottom

End Sub

Main
%>