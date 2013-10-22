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
	
	dim qc,ldqq
	If Request("code") <> "" Then
		SET qc = New QqConnet
		set ldqq = new leadbbs_forQQ
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
		Session("Openid") = qc.Getopenid()
		UserInfo = qc.GetUserInfo()
		dim nickname,sex,icon
			nickname = qc.GetUserName(UserInfo)(0)
		if nickname = "" then nickname = "QQÓÃ»§"
			sex = qc.GetUserName(UserInfo)(1)
			icon = "http://qzapp.qlogo.cn/qzapp/"&qc.APP_ID&"/"&Session("Openid")&"/30"
		Set qc = Nothing

		Response.Cookies(DEF_MasterCookies & "_AppInfo")="1," & Session("Access_Token") & "," & Session("Openid")
			Response.Cookies(DEF_MasterCookies&"_AppInfo").Expires = DateAdd("d",DEF_Now,365)
			Response.Cookies(DEF_MasterCookies&"_AppInfo").Domain = DEF_AbsolutHome
		Set qc = Nothing
		Dim UserID
		OpenDatabase
		UserID = ldqq.App_CheckAppid(1,Session("Openid"),Session("Access_Token"))
		If UserID > 0 Then
			ldqq.App_Login(UserID)
		Else
			if ldqq.App_BindExist(nickname,icon) = 0 then
				Response.Cookies(DEF_MasterCookies)("User") = CodeCookie(LeftTrue("QQ_" & nickname,20))
				Response.Cookies(DEF_MasterCookies).Expires = DateAdd("d",DEF_Now,365)
				Response.Cookies(DEF_MasterCookies).Domain = DEF_AbsolutHome
				Response.Cookies(DEF_MasterCookies&"_apptype") = "1"
				Response.Cookies(DEF_MasterCookies&"_apptype").Expires = DateAdd("d",DEF_Now,365)
				Response.Cookies(DEF_MasterCookies&"_apptype").Domain = DEF_AbsolutHome
			end if
		End If
		CloseDatabase
		Response.Redirect DEF_BBS_HomeUrl
	Else
		Dim url
		Session("Code") = ""
		Session("Openid") = ""
		Session("Access_Token") = ""
		Session("Openid") = ""
		SET qc = New QqConnet
		Session("State") = qc.MakeRandNum()
		url = qc.GetAuthorization_Code()
		Set qc = Nothing
		Response.Redirect(url)
	End If

End Sub

Class leadbbs_forQQ

	Private Token,ExpiresTime,Retention1,OpenID

	Private Sub Class_Initialize

		Token = ""
		ExpiresTime = 0
		Retention1 = ""
		OpenID = ""

	End Sub

	Public Function App_CheckAppid(AppType,appid,myToken)

		Dim Rs
		Set Rs = LDExeCute(sql_select("Select UserID,Token,ExpiresTime,Retention1 from LeadBBS_AppLogin where appType=" & Replace(appType,"'","''") & " and appid='" & Replace(appid,"'","''") & "'",1),0)
		If Rs.Eof Then
			App_CheckAppid = 0
			OpenID = Session("Openid")
			Token = Session("Access_Token")
		Else
			App_CheckAppid = cCur(Rs(0))
			Token = Rs(1)
			ExpiresTime = Rs(2)
			Retention1 = Rs(3)
			OpenID = appid
			if Token <> myToken then
				call LDExeCute("Update LeadBBS_AppLogin set Token='" & Replace(myToken,"'","''") & "' where appType=" & Replace(appType,"'","''") & " and appid='" & Replace(appid,"'","''") & "'",1)
				Token = myToken
			end if
		End if
		Rs.Close
		Set Rs = Nothing

	End Function

	Public Sub App_Login(UID)

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
	
	Public function App_BindExist(nickname,icon)
	
		If GBL_UserID > 0 Then
			App_BindExist = 1
			CALL LDExeCute("insert into LeadBBS_AppLogin(UserID,appid,GuestName,appType,ndatetime,IPAddress,Token) values(" & GBL_UserID & ",'" & Replace(OpenID,"'","''") & "','" & Replace(nickname,"'","''") & "',1," & GetTimeValue(DEF_Now) & ",'" & Replace(GBL_IPAddress,"'","''") & "','" & Replace(Token,"'","''") & "')",1)
		Else
			dim sql,userName,N,ExistFlag
			ExistFlag = 1
			App_BindExist = 0
			For N = 0 to 1000
				Randomize
				userName = "QQ#" & Mid(GetTimeValue(DEF_Now),3,6) & (Fix(Rnd*99999)+1)
				If CheckUserNameExist(userName) = 0 then
					ExistFlag = 0
					exit for
				End If
			Next
			
			If ExistFlag = 0 Then
				Dim width
				width = 30
				If right(icon,3) = "/30" then
					icon = mid(icon,1,len(icon)-3) & "/100"
					width = 100
				end if
				sql = "insert into leadbbs_User(UserName,Mail,Address,Sex,ICQ,OICQ,Userphoto,Homepage,Underwrite," &_
					"PrintUnderwrite,Pass,birthday,NongLiBirth,ApplyTime,IP,UserLevel,Officer,Points,Sessionid,Online," &_
					"Prevtime,Answer,Question,LastDoingTime,LastWriteTime,UserLimit,FaceUrl,FaceWidth,FaceHeight,LastAnnounceID) values(" &_
					"'" & userName & "','','','',0,0,0,'',''," &_
					"'','" & md5(rnd*99999999999+Timer) & "',0,0," & GetTimeValue(DEF_Now) & ",'',0,'',0,0,0," &_
					"" & GetTimeValue(DEF_Now) & ",'',''," & GetTimeValue(DEF_Now) & "," & GetTimeValue(DEF_Now) & ",0,'" & Replace(icon,"'","''") & "'," & width & "," & width & ",0" &_
					")"
				CALL LDExeCute(sql,1)
				Dim uid
				uid = GetUserID(userName)
				if uid > 0 then
					CALL LDExeCute("insert into LeadBBS_AppLogin(UserID,appid,GuestName,appType,ndatetime,IPAddress,Token) values(" & uid & ",'" & Replace(OpenID,"'","''") & "','" & Replace(nickname,"'","''") & "',1," & GetTimeValue(DEF_Now) & ",'" & Replace(GBL_IPAddress,"'","''") & "','" & Replace(Token,"'","''") & "')",1)
				end if
				App_BindExist = 1
				
				CALL LDExeCute("Update LeadBBS_SiteInfo Set UserCount=UserCount+1",1)
				UpdateStatisticDataInfo 1,1,1
				UpdateStatisticDataInfo userName,12,0
			End If
		End If
	
	End function
	
	Private Function CheckUserNameExist(n)

		Dim Rs
		Set Rs = LDExeCute(sql_select("Select ID from LeadBBS_User where UserName='" & Replace(n,"'","''") & "'",1),0)
		If Rs.Eof Then
			CheckUserNameExist = 0
		Else
			CheckUserNameExist = 1
		End if
		Rs.Close
		Set Rs = Nothing

	End Function
	
	Private Function GetUserID(UserName)

		Dim Rs
		Set Rs = LDExeCute(sql_select("Select ID from LeadBBS_User where UserName='" & Replace(UserName,"'","''") & "'",1),0)
		If Rs.Eof Then
			GetUserID = 0
		Else
			GetUserID = ccur(Rs(0))
		End if
		Rs.Close
		Set Rs = Nothing

	End Function

End Class

Main
%>