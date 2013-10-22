<!-- #include file=../inc/BBSsetup.asp -->
<!-- #include file=../inc/User_Setup.asp -->
<!-- #include file=../inc/Board_Popfun.asp -->
<!-- #include file=inc/UserTopic.asp -->
<%
Response.Expires = 0 
Response.ExpiresAbsolute = DEF_Now - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private" 
Response.CacheControl = "no-cache"

DEF_BBS_HomeUrl = "../"
Dim AjaxFlag

Main

Sub Main

	If Request("AjaxFlag") = "1" Then
		AjaxFlag = 1
	Else
		AjaxFlag = 0
	End If
	Select Case Request("action")
		Case "logout":
			Main_Logout
		Case "hidden":
			Main_Hidden
		Case "err":
			BBS_SiteHead DEF_SiteNameString & " - Error",0,"<span class=navigate_string_step>Error</span>"
			Boards_Body_Head("")
			Global_ErrMsg(Request.QueryString("err"))
			Boards_Body_Bottom
			SiteBottom
		Case Else
			Main_login	
	End Select

End Sub

Sub Main_login

	OpenDatabase
	GBL_UserID = CheckPass
	
	If AjaxFlag = 0 Then
		BBS_SiteHead DEF_SiteNameString & " - ��¼",0,"<span class=navigate_string_step>��¼</span>"
		
		UpdateOnlineUserAtInfo GBL_board_ID,"��¼"
	
		Boards_Body_Head("")
		%>
	<div class='alertbox fire'>
		<%
	Else%>
	<div class="ajaxbox">
	<%
	End If
	If GBL_CHK_Flag=1 and Request("R")<>"Yes" Then		
		If CheckWriteEventSpace = 0 Then
			Processor_LoginMsg "���Ĳ�����Ƶ��(��¼̫Ƶ)�Ժ�����!","login_title",""
		Else		
			UpdateUserLevel
			LoginAccuessFul
		End If
	Else
		If Request("submitflag")="" Then
			DisplayLoginForm("")
		Else
			If AjaxFlag = 1 Then
				If GBL_CHK_TempStr = "" Then GBL_CHK_TempStr = "��¼��Ϣ���󣬻��ǵ�½ʧ�ܴ������ࡣ"
				Processor_LoginMsg GBL_CHK_TempStr,"login_title","submit_disable($id('login_form'),1);"
			Else
				DisplayLoginForm(GBL_CHK_TempStr)
			End If
		End If
	End If
	closeDataBase
	%>
	</div>
	<%
	If AjaxFlag = 0 Then
		Boards_Body_Bottom
		SiteBottom
	End If

End Sub

Sub LoginAccuessFul

	Dim u
	u = filterUrlstr(Request("u"))
	If u = "" Then u = "../Boards.asp"
	
	If AjaxFlag = 1 Then
		Processor_LoginMsg "<div class=""ajaxbox""><div class=""title"">���Ѿ��ɹ���¼����ҳ���Ժ󽫷����Զ�ˢ�£�</div><div class=""value2""><a href=""" & u & """>��Ҳ���Ե���˴�����ˢ�¡�</a></div></div>","anc_delbody","setTimeout(""document.location.href='" & u & "'"",1000);"
		Exit Sub
	End If
	
	%>

	���Ѿ��ɹ���¼����ҳ�潫��5����Զ�����������ҳ�����Լ���ѡ�����²�����</b>
	<p>
	- <a href=<%=DEF_BBS_HomeUrl%>>������̳��ҳ</a>
	<br><br>
	<table border="0" cellspacing="0" cellpadding="0">
	<tr><td>-&nbsp;</td><td>
	<!-- #include file=../inc/IncHtm/BoardJump.asp -->
	</td></tr></table>
	<%
	If u <> "" Then
	%><br>
	- ����<a href="<%=htmlencode(u)%>"><%=htmlencode(u)%></a><%
	End If%>
	
	<script language=javascript>
				function a_topage()
				{
					document.location.href = "<%
	
	If u <> "" Then
		Response.Write htmlencode(u)
	Else
		u = "../Boards.asp"
		Response.Write u
	End If
	
	Response.Clear
	CloseDatabase
	Response.Redirect u%>"; 
				}
				setTimeout("a_topage()",1000);
				</script>
	<br>
	<%If DEF_RepeatLoginTimeOut > 0 and DEF_RepeatLoginTimeOut < DEF_UserOnlineTimeOut Then
		Response.write "<br>ע�⣺��̳�ѿ������ظ���¼���ܣ���¼�������˽���Ȩʹ�������˺�"
	End If
    
End Sub

Sub Main_Logout

	initDatabase
	If Request.Form("sure")="1" Then
		Dim UserID
		Dim Rs
		Set Rs = LDExeCute(sql_select("Select UserID from LeadBBS_onlineUser where sessionID=" & session.sessionID,1),0)
		If Rs.Eof Then
			UserID = GBL_UserID
		Else
			UserID = cCur(Rs("UserID"))
		End If
		Rs.Close
		Set Rs = Nothing

		If Request.Form("clearck") = "1" Then
			Dim Cookie,Key
			For Each Cookie in Request.Cookies
				If Request.Cookies(Cookie).HasKeys =false then
					Response.Cookies(Cookie) = ""
					Response.Cookies(Cookie).Expires = Date - 1
					Response.Cookies(Cookie).Domain = DEF_AbsolutHome
				Else
					For Each Key in Request.Cookies(Cookie)
						Response.Cookies(Cookie)(Key) = ""
					Next
					Response.Cookies(Cookie).Expires = Date - 1
					Response.Cookies(Cookie).Domain = DEF_AbsolutHome
				End If
			Next
			GBL_AppType = ""
			Pub_ClearCookie
		Else
			GBL_AppType = ""
			Pub_ClearCookie
			'Response.Cookies(DEF_MasterCookies & "_" & GBL_UserID).Expires = Date - 1
			'Response.Cookies(DEF_MasterCookies & "_" & GBL_UserID).Domain = DEF_AbsolutHome
			'Response.Cookies(DEF_MasterCookies & "Time").Expires = Date - 1
			'Response.Cookies(DEF_MasterCookies & "Time").Domain = DEF_AbsolutHome
			'Response.Cookies(DEF_MasterCookies & "style").Expires = Date - 1
			'Response.Cookies(DEF_MasterCookies & "style").Domain = DEF_AbsolutHome
		End If
		
		
		UpdateOnlineUserInfo(" from LeadBBS_onlineUser where sessionID=" & session.sessionID)
		If GBL_UserID > 0 Then UpdateOnlineUserInfo(" from LeadBBS_onlineUser where UserID=" & GBL_UserID)
		SetActiveUserCount
		session.abandon
		If UserID > 0 Then
			CALL LDExeCute("Update LeadBBS_User set LastDoingTime=" & GetTimeValue(DateAdd("s", 0-DEF_UserOnlineTimeOut, DEF_Now)) & " where ID=" & UserID,1)
		End If
		closeDatabase
		
		Dim u
		u = filterUrlstr(Request("u"))
		If u = "" Then
			u = filterUrlstr(Lcase(Request.ServerVariables("HTTP_REFERER")))
			
			Dim HomeUrl
			HomeUrl = "http://"&Request.ServerVariables("server_name")
			If Request.ServerVariables("SERVER_PORT") <> "80" Then HomeUrl = HomeUrl & ":" & Request.ServerVariables("SERVER_PORT")
			
			If Left(u,Len(HomeUrl)) <> HomeUrl Then u = ""
			If inStr(u,"/user/login.asp") > 0 Then u = ""
		End If
		If u = "" Then u = DEF_BBS_HomeUrl & "Boards.asp"
		If AjaxFlag = 0 Then 
			Response.Redirect DEF_BBS_HomeUrl & "Boards.asp"
		Else
			Processor_LoginMsg "<div class=""ajaxbox""><div class=""title"">���Ѿ��ɹ��˳�����ҳ���Ժ��Զ�ˢ�¡�</div><div class=""value2""><a href=""" & u & """>����˴�����ˢ�¡�</a></div></div>","anc_delbody",""
		End If
	Else
		BBS_SiteHead DEF_SiteNameString & " - �˳�",0,"<span class=navigate_string_step>�˳�</span>"
		Boards_Body_Head("")
		%>
		<div class='alertbox fire'>
		<form name=DellClientForm action=Login.asp?action=logout method=post>
			<input type=hidden name=sure value="1">
			<div class=title>��ȷ���˳�, ��Ҫ�����밴ȷ��.</div>
			<div class=value2><input class=fmchkbox type="checkbox" name="clearck" value="1" checked>ͬʱ��ձ�վCOOKIE��Ϣ</div>
			<br>
			<div class=value2><input type=submit value=ȷ�� class="fmbtn btn_2"></div>
		</form>
		<%
		closeDataBase
		Boards_Body_Bottom
		SiteBottom
		If GBL_ShowBottomSure = 1 Then Response.Write GBL_SiteBottomString
	End If

End Sub

Sub Main_Hidden

	initDatabase

	GBL_CHK_TempStr = ""
	
	Dim ShowFlagString
	If GBL_CHK_ShowFlag = 1 Then
		ShowFlagString = "����"
	Else
		ShowFlagString = "����"
	End If
	If AjaxFlag = 0 Then BBS_SiteHead DEF_SiteNameString & " - " & ShowFlagString,0,"<span class=navigate_string_step>" & ShowFlagString & "</span>"
	
	If GBL_UserID=0 Then GBL_CHK_TempStr = GBL_CHK_TempStr & "��û�е�¼!" & VbCrLf
	
	Dim u
	u = filterUrlstr(Request("u"))
	If u = "" Then
		u = filterUrlstr(Lcase(Request.ServerVariables("HTTP_REFERER")))
		
		Dim HomeUrl
		HomeUrl = "http://"&Request.ServerVariables("server_name")
		If Request.ServerVariables("SERVER_PORT") <> "80" Then HomeUrl = HomeUrl & ":" & Request.ServerVariables("SERVER_PORT")
		
		If Left(u,Len(HomeUrl)) <> HomeUrl Then u = ""
		If inStr(u,"/user/login.asp") > 0 Then u = ""
	End If
	If u = "" Then u = DEF_BBS_HomeUrl & "Boards.asp"
	
	If AjaxFlag = 0 Then
		Boards_Body_Head("")
		%>
		<div class='alertbox fire'>
		<%
	Else%>
		<div class="ajaxbox">
	<%End If
	
	If Request.Form("sure") <> "1" Then
		%>
		<form name=DellClientForm action=Login.asp?action=hidden method=post>
			<input type=hidden name=sure value="1">
			<div class=title>�밴ȷ��������</div>
			<input type=hidden value="<%Response.Write htmlencode(u)%>" name=u>
			<div class=value2><input type=submit value=ȷ�� class="fmbtn btn_2">
		</form>
		<%
	Else
		If DEF_EnableUserHidden = 1 Then
			If GBL_CHK_Flag=1 Then
				If ShowFlagString = "����" Then
					CALL LDExeCute("Update LeadBBS_User Set ShowFlag=1 where ID=" & GBL_UserID,1)
					UpdateSessionValue 3,1,0
					CALL LDExeCute("Update LeadBBS_OnlineUser Set HiddenFlag=0,UserName='�����û�' where UserID=" & GBL_UserID,1)
				Else
					CALL LDExeCute("Update LeadBBS_User Set ShowFlag=0 where ID=" & GBL_UserID,1)
					UpdateSessionValue 3,0,0
					CALL LDExeCute("Update LeadBBS_OnlineUser Set HiddenFlag=" & GBL_CHK_UserLimit & ",UserName='" & Replace(GBL_CHK_User,"'","''") & "' where UserID=" & GBL_UserID,1)
				End If
				
				If AjaxFlag = 1 Then
					Processor_LoginMsg "<div class=""title"">���Ѿ��ɹ�" & ShowFlagString & "����ҳ���Ժ��Զ�ˢ�¡�</div><div class=""value2""><a href=""" & u & """>��Ҳ���Ե���˴�����ˢ�¡�</a></div>","anc_delbody",""
				Else
					Response.Write "<p>���Ѿ��ɹ�" & ShowFlagString
					If u <> "" Then Response.Redirect u
				End If
			Else
				If Request("submitflag")="" Then
					DisplayLoginForm("���ȵ�¼")
				Else
					DisplayLoginForm(GBL_CHK_TempStr)
				End If
			End If
		Else%>
			<div class=alert>
				��̳�Ѿ���ֹʹ��������
			</div>
		<%End If
	End If
	%>
	</div>
	<%
	closeDataBase
	If AjaxFlag = 0 Then
		Boards_Body_Bottom
		SiteBottom
	End If

End Sub

Rem ���û���¼ʱ����Ҫ����һЩ��Ϣ����������¼ʱ���
Function UpdateUserLevel

	Dim Temp_N,IP,SessionID,Prevtime
	Dim Rs
	IP = GBL_IPAddress
	SessionID = session.sessionid
	Prevtime = GetTimeValue(DEF_Now)

	If GBL_CHK_ShowFlag = 1 and DEF_EnableUserHidden = 1 Then
		Temp_N = "�����û�"
	Else
		If GBL_UserID > 0 Then
			Temp_N = GBL_CHK_User
		Else
			Temp_N = ""
		End If
	End If
	Dim OnlineID,CountFlag,TmpSessionID,OnlineUserID,tmp,i
	
	If GBL_CHK_ShowFlag = 1 and DEF_EnableUserHidden = 1 Then
		i = "�����û�"
		tmp = 0
	Else
		If GBL_UserID > 0 Then
			i = GBL_CHK_User
			tmp = GBL_CHK_UserLimit
		Else
			I = ""
			tmp = 0
		End If
	End If
	Set Rs = LDExeCute(sql_select("select id,SessionID from LeadBBS_onlineUser where UserID=" & GBL_UserID,2),0)
	If Not Rs.Eof Then
		OnlineID = cCur(Rs(0))
		TmpSessionID = Rs(1)
		Rs.MoveNext
		If Not Rs.Eof Then
			CountFlag = 1
		Else
			CountFlag = 0
		End If
		Rs.Close
		Set Rs = Nothing
		If CountFlag = 1 Then
			UpdateOnlineUserInfo("from LeadBBS_onlineUser where UserID=" & GBL_UserID & " and ID<>" & OnlineID)
			SetActiveUserCount
		End if

		If GBL_UserID > 0 and OnlineID > 0 Then
			CALL LDExeCute("Update LeadBBS_onlineUser set sessionID=" & cCur(SessionID) & ",HiddenFlag=" & GBL_CHK_UserLimit & ",UserID=" & GBL_UserID & ",UserName='" & Replace(Temp_N,"'","''") & "',LastDoingTime=" & GetTimeValue(DEF_Now) & ",IP='" & GBL_IPAddress & "' where ID=" & OnlineID,1)
			UpdateSessionValue 17,GBL_IPAddress,0
			UpdateSessionValue 18,GetTimeValue(DEF_Now),0
		End If
	Else
		Rs.Close
		Set Rs = Nothing
		Set Rs = LDExeCute(sql_select("select UserID,LastDoingTime,sessionID,ID from LeadBBS_onlineUser where IP='" & Replace(GBL_IPAddress,"'","''") & "' and UserID=0",1),0)
		If Rs.Eof Then
			Rs.close
			Set Rs = Nothing
			UpdateOnlineUserInfo("from LeadBBS_onlineUser where SessionID=" & cCur(SessionID))
			CALL LDExeCute("insert into LeadBBS_onlineUser(SessionID,UserID,LastDoingTime,IP,StartTime,AtBoardID,AtUrl,AtInfo,Browser,System,UserName,HiddenFlag,LastRndNumber) values(" & cCur(SessionID) & "," & cCur(GBL_UserID) & "," & GetTimeValue(DEF_Now) & ",'" & GBL_IPAddress & "'," & GetTimeValue(DEF_Now) & ",0,'" & Replace(Left(Request.Servervariables("SCRIPT_NAME") & "?" & Request.QueryString,255),"'","''") & "','����ҳ��','" & Left(Replace(GetSBInfo(1),"'","''"),30) & "','" & Left(Replace(GetSBInfo(2),"'","''"),30) & "','" & Replace(i,"'","''") & "'," & cCur(tmp) & "," & (Fix(Timer) mod 9999) & ")",1)

			If GBL_CHK_User <> "" and GBL_UserID > 0 and CheckSupervisorUserName = 1 Then
				CALL LDExeCute("insert into LeadBBS_Log(LogType,LogTime,LogInfo,UserName,IP,BoardID) Values(51," & GetTimeValue(DEF_Now) & ",'����Ա��¼��̳.','" & Replace(Replace(htmlencode(Left(GBL_CHK_User,14)),"\","\\"),"'","''") & "','" & Replace(GBL_IPAddress,"'","''") & "'," & GBL_Board_ID & ")",1)
			End If

			Application.Lock
			application(DEF_MasterCookies & "ActiveUsers") = application(DEF_MasterCookies & "ActiveUsers") + 1
			Application.UnLock
			OnlineID = 0
		Else
			OnlineID = cCur(Rs(3))
			Rs.Close
			Set Rs = Nothing	
			If GBL_UserID > 0 and OnlineID > 0 Then
				CALL LDExeCute("Update LeadBBS_onlineUser set sessionID=" & cCur(SessionID) & ",HiddenFlag=" & GBL_CHK_UserLimit & ",UserID=" & GBL_UserID & ",UserName='" & Replace(Temp_N,"'","''") & "',LastDoingTime=" & GetTimeValue(DEF_Now) & ",IP='" & GBL_IPAddress & "' where ID=" & OnlineID,1)
			End If
		End If
	End If

	CALL LDExeCute("Update LeadBBS_User set Prevtime=" & Prevtime & ",LastDoingTime=" & GetTimeValue(DEF_Now) & ",LastWriteTime=" & GetTimeValue(DEF_Now) & " where id=" & GBL_UserID,1)
	UpdateSessionValue 11,Prevtime,0
	UpdateSessionValue 13,GetTimeValue(DEF_Now),0
	UpdateUserLevel = 1
	GBL_CHK_Flag = 1	

End Function
%>