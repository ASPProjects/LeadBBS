<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Board_Popfun.asp -->
<!--#include file="oauth.asp"-->
<%
DEF_BBS_HomeUrl = "../../"

Sub Main

	If GetBinarybit(DEF_Sideparameter,10) = 0 Then
		Response.Write "Error code:0x0000ff."
		Exit Sub
	End If
	If apiKey = "" or secretKey = "" Then
		Response.Write "Error code:0x0000fe."
		Exit Sub
	End If
	If apiKey = "" or secretKey = "" Then
		Response.Write "Error code:0x0000fe."
		Exit Sub
	End If
	
	dim qc
	If Request("code")<>"" Then
	
	
	
		SET qc = New QqConnet
		If Session("Access_Token") = "" Then
			dim CheckLogin
			CheckLogin=qc.CheckLogin()
			If CheckLogin=False Then
				Response.Write("µÇÂ¼Ê§°Ü£¡")
				Response.End()
			Else
				Session("Access_Token")=qc.GetAccess_Token()
			End If
		End If
		dim UserInfo
		Session("Openid")=qc.Getopenid()
		UserInfo=qc.GetUserInfo()
		dim nickname,sex,icon
		'Response.Write(userinfo)
			nickname = qc.GetUserName(UserInfo)(0)
			sex = qc.GetUserName(UserInfo)(1)
			icon = "http://qzapp.qlogo.cn/qzapp/"&qc.APP_ID&"/"&Session("Openid")&"/30"
		Set qc=Nothing
		   ' Response.Write(Session("State"))

		Response.Cookies(DEF_MasterCookies & "_AppInfo")="1," & Session("Access_Token") & "," & Session("Openid")
			Response.Cookies(DEF_MasterCookies&"_AppInfo").Expires = DateAdd("d",DEF_Now,365)
			Response.Cookies(DEF_MasterCookies&"_AppInfo").Domain = DEF_AbsolutHome
		'Response.Write("êÇ³Æ£º"&nickname&"<br />")
		'Response.Write("sex£º"&sex&"<br />")
		'Response.Write("Í·Ïñ£º"&icon&"<br />")
		'Response.Write("openid£º"&Session("Openid")&"<br />")
		'Response.Write("Access_Token£º"&Session("Access_Token")&"<br />")
		Set qc=Nothing
		Dim UserID
		OpenDatabase
		UserID = App_CheckAppid(1,Session("Openid"))
		If UserID > 0 Then
			App_Login(UserID)
		Else
			Response.Cookies(DEF_MasterCookies)("User") = CodeCookie(LeftTrue("QQ_" & nickname,20))
			Response.Cookies(DEF_MasterCookies).Expires = DateAdd("d",DEF_Now,365)
			Response.Cookies(DEF_MasterCookies).Domain = DEF_AbsolutHome
			Response.Cookies(DEF_MasterCookies&"_apptype") = "1"
			Response.Cookies(DEF_MasterCookies&"_apptype").Expires = DateAdd("d",DEF_Now,365)
			Response.Cookies(DEF_MasterCookies&"_apptype").Domain = DEF_AbsolutHome
		End If
		CloseDatabase
		Response.Redirect DEF_BBS_HomeUrl
	Else 
		Dim url
		Session("Code")=""
		Session("Openid")=""
		Session("Access_Token")=""
		Session("Openid")=""
		SET qc = New QqConnet
		Session("State")=qc.MakeRandNum()
		url = qc.GetAuthorization_Code()
		Set qc=Nothing
		Response.Redirect(url)
	End If 

End Sub


Function App_CheckAppid(AppType,appid)

	Dim Rs
	Set Rs = LDExeCute(sql_select("Select UserID from LeadBBS_AppLogin where appType=" & Replace(appType,"'","''") & " and appid='" & Replace(appid,"'","''") & "'",1),0)
	If Rs.Eof Then
		App_CheckAppid = 0
	Else
		App_CheckAppid = cCur(Rs(0))
	End if
	Rs.Close
	Set Rs = Nothing

End Function

Sub App_Login(UID)

	'Response.Write "..........."
	Dim Rs
	Set Rs = LDExeCute(sql_select("Select ID,UserName,Pass from LeadBBS_User where ID=" & UID,1),0)
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		Exit Sub
	End If
	GBL_CHK_User = Rs(1)
	GBL_CHK_Pass = Rs(2)
	dontRequestFormFlag = "AppLogin"
	GBL_CheckPassDoneFlag = 0
	GBL_CHK_Flag = 1
	checkPass

End Sub

Main
%>