<!-- #include file=../inc/BBSsetup.asp -->
<!-- #include file=../inc/User_Setup.asp -->
<!-- #include file=../inc/Board_Popfun.asp -->
<!-- #include file=inc/UserTopic.asp -->
<!-- #include file=../inc/Limit_Fun.asp -->
<!-- #include file=inc/Mail_fun.asp -->
<!-- #include file=inc/UserGetPass_Fun.asp -->
<!-- #include file=inc/UserActive_Fun.asp -->
<%
Response.Expires = 0 
Response.ExpiresAbsolute = DEF_Now - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private" 
Response.CacheControl = "no-cache"

DEF_BBS_HomeUrl = "../"

Main

Sub Main

	initDatabase
	GBL_CHK_TempStr = ""
	
	BBS_SiteHead DEF_SiteNameString & " - УмТыевЛи",0,"<span class=navigate_string_step>УмТыевЛи</span>"

	Boards_Body_Head("")

	%>
	<div class=alertbox>
	<%
	Main_GetPass
	%>
	</div>
	<%
	
	Boards_Body_Bottom
	closeDataBase
	SiteBottom

End Sub

Sub Main_GetPass

	Select Case Left(Request("act"),10)
		Case "active"
			Dim UserActive
			Set UserActive = New User_UserActive
			UserActive.DisplayActive
			Set UserActive = Nothing
		Case else
			Dim UserGetPass
			Set UserGetPass = New User_GetPass
			UserGetPass.GetPass
			Set UserGetPass = Nothing
	End Select

End Sub%>