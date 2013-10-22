<!-- #include file=../inc/BBSsetup.asp -->
<!-- #include file=../inc/Board_Popfun.asp -->
<!-- #include file=../User/inc/UserTopic.asp -->
<!-- #include file=inc/List_fun.asp -->
<%
DEF_BBS_HomeUrl = "../"

Dim GBL_NoneLimitFlag

Sub Main

	GBL_CHK_PWdFlag = 0
	initDatabase
	GBL_CHK_TempStr = ""
	BBS_SiteHead DEF_SiteNameString & " - 论坛帖子",0,"<span class=navigate_string_step>论坛帖子</span>"
	UpdateOnlineUserAtInfo 0,"论坛帖子"
	
	GBL_NoneLimitFlag = CheckSupervisorUserName  '管理员无限制

	UserTopicTopInfo("forum")
	If GBL_CHK_TempStr = "" Then
		GBL_CHK_TempStr = ""
		DisplayAnnouncesSplitPages
	Else
		Response.Write "<div class=alert>" & GBL_CHK_TempStr & "</div>"
	End If
	closeDataBase
	UserTopicBottomInfo
	SiteBottom

End Sub

Main
%>