<%@ LANGUAGE=VBScript CodePage=936%>
<%Option Explicit
Response.Charset = "gb2312"
Session.CodePage=936
Response.Buffer = True
Const DEF_ManageDir = "manage"
Response.Redirect "install/default.asp"
If isNumeric(application(DEF_MasterCookies & "SiteEnableFlagzoieiu")) = 0 Then
	Application.Lock
	application(DEF_MasterCookies & "SiteEnableFlagzoieiu") = 1
	Application.UnLock
End If
If application(DEF_MasterCookies & "SiteEnableFlagzoieiu") = 0 and application(DEF_MasterCookies & "SiteDisbleWhyszoieiu")<>"" and inStr(Replace(Lcase(Request.ServerVariables("URL")),"\","/"),"/" & DEF_ManageDir & "/") = 0 Then
	Response.Write application(DEF_MasterCookies & "SiteDisbleWhyszoieiu")
	Response.End
End If

Dim DEF_BBS_HomeUrl,DEF_SiteHomeUrl
const DEF_BBS_Name="LeadBBS"
DEF_BBS_HomeUrl = ""
DEF_SiteHomeUrl = "/Boards.asp"
const DEF_BBS_DarkColor = "#cccccc"
const DEF_BBS_LightDarkColor = "#666699"
const DEF_BBS_Color = "#e0e0e0"
const DEF_BBS_LightColor = "eeeeee"
const DEF_BBS_LightestColor = "#f7f7f7"
const DEF_BBS_TableHeadColor = "#EEEEF3"
const DEF_BBS_MaxLayer = 10
const DEF_UsedDataBase = 2
const DEF_BBS_SearchMode = 0
const DEF_BBS_TOPMinID = 99999999990000
const DEF_BBS_AnnouncePoints = 1
const DEF_BBS_PrizeAnnouncePoints = 3
const DEF_BBS_MakeGoodAnnouncePoints = 5
const DEF_BBS_MaxTopAnnounce = 9
const DEF_BBS_MaxAllTopAnnounce = 3
Dim DEF_BBS_DisplayTopicLength,DEF_BBS_ScreenWidth
DEF_BBS_DisplayTopicLength = 56
DEF_BBS_ScreenWidth = "770"
const DEF_BBS_LeftTDWidth = "180"
const DEF_MasterCookies = "ld"
const DEF_SiteNameString = ""
const DEF_SupervisorUserName = ",Admin,"
const DEF_MaxTextLength = 51200
Dim DEF_MaxListNum
DEF_MaxListNum = 32
Const DEF_TopicContentMaxListNum = 12
Const DEF_MaxJumpPageNum = 8000
Const DEF_DisplayJumpPageNum = 4
const DEF_MaxBoardMastNum = 10
const DEF_EnableUserHidden = 1
const DEF_VOTE_MaxNum = 25
const DEF_MaxLoginTimes = 5
const DEF_RestSpaceTime = 1
const DEF_LoginSpaceTime = 600
const DEF_EnableUpload = 1
const DEF_EnableGFL = 0
const DEF_UserOnlineTimeOut = 7199
const DEF_faceMaxNum = 254
const DEF_AllDefineFace = 2
const DEF_AllFaceMaxWidth = 120
const DEF_BBS_EmailMode = 0
Const DEF_EnableAttestNumber = 0
Const DEF_AttestNumberPoints = 0
Dim DEF_BoardStyleString,DEF_BoardStyleStringNum
DEF_BoardStyleString = Array("蓝","绿","黄","2012")
DEF_BoardStyleStringNum = Ubound(DEF_BoardStyleString,1)
Const DEF_EnableUnderWrite = 1
Const DEF_NeedOnlineTime = 0
Const DEF_EnableForbidIP = 0
Const DEF_TopAdString = "<a href=""http://idc.leadbbs.com/"" target=""_blank""><img src=""/images/temp/banner17.gif"" width=""468"" height=""60"" alt=""空间广告"" /></a>"
Const DEF_AccessDatabase = ""
Const DEF_DefaultStyle = 1001
Const DEF_EnableFlashUBB = 1
Const DEF_EnableImagesUBB = 1
Const DEF_AnnounceFontSize = "14px;"
Const DEF_EditAnnounceDelay = 300
Const DEF_DisplayOnlineUser = 1
Const DEF_EnableSpecialTopic = 1
Const DEF_UBBiconNumber = 100
Const DEF_EnableDelAnnounce = 0
Dim DEF_PointsName
DEF_PointsName = Array("积分","财富","声望","等级","经验","认证会员","总版主","区版主","论坛版主","门派","专业用户")
Const DEF_EnableMakeTopAnc = 1
Const DEF_EnableDatabaseCache = 0
Const DEF_WriteEventSpace = 2
Const DEF_EnableTreeView = 0
Const DEF_EditAnnounceExpires = 0
Const DEF_RepeatLoginTimeOut = 0
Const DEF_FSOString = "Scripting.FileSystemObject"
Dim DEF_Now,DEF_Version
DEF_Now = now
DEF_Version = "LeadBBS 7.0"
Const DEF_LineHeight = 27
Const DEF_RegisterFile = "Register.asp"
Const DEF_LimitTitle = 1
Const DEF_DownKey = "uNv8A3pLefMn"
Const DEF_UpdateInterval = 300
Const DEF_BottomInfo = " "
Dim DEF_GBL_Description
DEF_GBL_Description = "LeadBBS论坛"
Const DEF_Sideparameter = 0
%>
