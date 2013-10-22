<%Class User_GetPass

	Private SendEmail,SendPass,Question,Answer,SendAnswer,SendUser,SendQuestion,SendPassword1,SendPassword2

	Public Sub GetPass
	
		VierGetPassForm

	End Sub

	Private Sub VierGetPassForm
	
		SendUser = Trim(Request.Form("SendUser"))
		If DEF_User_GetPassMode = 0 Then
			Response.Write "<div class=alert>论坛已经关闭密码找回功能。</div>" & VbCrLf
			Exit Sub
		End If
		If DEF_BBS_EmailMode = 0 and SendUser  <> "" and DEF_User_GetPassMode = 2 Then
			Response.Write "<div class=alert>论坛禁止发送邮件，密码找回功能不能使用。</div>" & VbCrLf
			Exit Sub
		End If
		If Request.Form("act") = "getpass" Then
			SendAnswer = Left(Request.Form("SendAnswer"),20)
			SendQuestion = Left(Request.Form("SendQuestion"),20)
			SendPassword2 = Left(Request.Form("SendPassword2"),14)
			SendPassword1 = Left(Request.Form("SendPassword1"),14)
	
			If Len(SendUser) > 30 Then
				Response.Write "<div class=alert>错误: 用户名太长.</div>" & VbCrLf
				Exit Sub
			End If
					
			SQL = "Select LastDoingTime from LeadBBS_OnlineUser where SessionID=" & Session.SessionID
			Set Rs = LDExeCute(SQL,0)
			If Rs.Eof Then
				GBL_CHK_LastWriteTime = 0
				Rs.Close
				Set Rs = Nothing
				Response.Write "<div class=alert>请在线2分钟以上后再进行密码找回!</div>" & VbCrLf
				Exit Sub
			Else
				GBL_CHK_LastWriteTime = DateDiff("s",RestoreTime(Rs(0)),DEF_Now)
				If GBL_CHK_LastWriteTime < 0 Then GBL_CHK_LastWriteTime = DEF_WriteEventSpace + 1
			End If
	
			If CheckWriteEventSpace = 0 Then
				Response.Write "<div class=alert>您的操作过频，请稍候再作提交!</div>" & VbCrLf
				DisplaySubmitForm
				Exit Sub
			End If
			Dim Rs,SQL
			SQL = "Select mail,Pass,Question,Answer,UserLimit from LeadBBS_User where UserName='" & Replace(SendUser,"'","''") & "'"
			Set Rs = LDExeCute(SQL,0)
			If Rs.Eof Then
				Rs.Close
				Set Rs = Nothing
				Response.Write "<div class=alert>错误: 不存在的用户名.</div>" & VbCrLf
				Exit Sub
			End If
	
			SendEmail = Rs(0)
			SendPass = Rs(1)
			Question = Rs(2)
			Answer = Rs(3)
			GBL_CHK_UserLimit = Rs(4)
			Rs.Close
			Set Rs = Nothing
			If SendEmail = "" or isNull(SendEmail) Then
				Response.Write "<div class=alert>错误: 此用户注册时未提供Email，无法找回密码.</div>" & VbCrLf
				Exit Sub
			End If
	
			SQL = GBL_CHK_User
			GBL_CHK_User = SendUser
			CheckisBoardMaster
			If Len(Answer) < 32 or GBL_BoardMasterFlag >= 4 or (GBL_CHK_User <> "" and inStr(GBL_CHK_User,",") = 0 and inStr(LCase(DEF_SupervisorUserName),"," & LCase(GBL_CHK_User) & ",") > 0) Then
				Response.Write "<div class=alert>错误: 此用户未申请密码保护，无法使用找回密码恢复功能。<br>请联系管理员设定密码保护。</div>" & VbCrLf
				Exit Sub
			End If
			GBL_CHK_User = SQL
	
			If SendAnswer = "" and SendQuestion = "" and SendPassword1 = "" and SendPassword2 = "" Then
				DisplaySubmitForm
			Else
				Dim NumCheck
				NumCheck = CheckRndNumber	
				Randomize
				Session(DEF_MasterCookies & "RndNum") = Fix(Rnd*9999)+1
				If NumCheck = 0 Then
					Response.Write "<div class=alert>验证码填写错误!</div>" & VbCrLf
					DisplaySubmitForm
					Exit Sub
				End If
	
				If Len(SendPassword2) < DEF_UserShortestPassword or Len(SendPassword1) < DEF_UserShortestPassword or SendPassword1 <> SendPassword2 Then
					Response.Write "<div class=alert>新的密码不能少于4位，并且新密码与验证密码必须相同。</div>" & VbCrLf
					SendQuestion = ""
					DisplaySubmitForm
					Exit Sub
				End If
			
				If MD5(SendAnswer) <> Answer and Mid(MD5(SendAnswer),9,16) <> Answer Then
					Response.Write "<div class=alert>密码的提示答案填写错误，无法找回账号，或请联系管理员!</div>" & VbCrLf
					SendQuestion = ""
					DisplaySubmitForm
					CALL LDExeCute("Update LeadBBS_OnlineUser Set LastDoingTime=" & GetTimeValue(DEF_Now) & " where SessionID=" & Session.SessionID,1)
					UpdateSessionValue 18,GetTimeValue(DEF_Now),0
					Exit Sub
				End If
				
				Dim NewSendPass
				SendPass = MD5(SendPassword2)
				CALL LDExeCute("Update LeadBBS_User Set Pass='" & Replace(SendPass,"'","''") & "' where UserName='" & Replace(SendUser,"'","''") & "'",1)
				If Lcase(GBL_CHK_User) = Lcase(SendUser) Then UpdateSessionValue 7,SendPass,0
				Response.Write "<div class=alert>密码已经成功更改，请使用新的密码<a href=""Login.asp?User=" & urlencode(SendUser) & """>登录</a>您的账号!</div>" & VbCrLf
				If DEF_User_GetPassMode = 2 and GetBinarybit(GBL_CHK_UserLimit,1) = 1 Then
					SQL = "Select BoardID from LeadBBS_SpecialUser where UserID=" & GBL_UserID
					Set Rs = LDExeCute(SQL,0)
					If Rs.Eof Then
						Rs.Close
						Set Rs = Nothing
						Response.Write "<div class=alert>此用户无法由用户进行激活，停止发送邮件，请联系管理员.</div>" & VbCrLf
						Exit Sub
					End If
					SQL = cCur(Rs(0))
					Rs.Close
					Set Rs = Nothing
					SendGetPassMail SendUser,SendEmail,SendPassword2,SQL
					Response.Write "<div class='alert greenfont'>同时，新的密码及激活码已经发送到您的注册邮箱。</div>" & VbCrLf
				End If
				CALL LDExeCute("Update LeadBBS_OnlineUser Set LastDoingTime=" & GetTimeValue(DEF_Now) & " where SessionID=" & Session.SessionID,1)
				UpdateSessionValue 18,GetTimeValue(DEF_Now),0
			End If
		Else
			DisplaySubmitForm
		End If
	
	End Sub
	
	Private Sub DisplaySubmitForm
	
		If Question = "" Then Question = SendQuestion
		If SendAnswer = "" and SendUser = "" and SendQuestion = "" Then%>
	<div class=title>请输入您的用户名。</div>
	<form action=<%=DEF_BBS_HomeUrl%>User/UserGetPass.asp method="post" onSubmit="submit_disable(this);">
		用户名: <input name=SendUser type=text maxlength=20 size=22 value="<%
		If GBL_CHK_user = "" or isNull(GBL_CHK_user) Then
			Response.Write htmlencode(Request("user"))
		Else
			Response.Write htmlencode(GBL_CHK_user)
		End If%>" class=fminpt><br>
		<input type=hidden value="getpass" name=act><br>
		<input type=submit value="取回密码" class="fmbtn btn_3"> <input type=reset value="取消" class="fmbtn btn_2">
	</form>
	<br>
	<div class=value2>注意： 版主以上用户及未设立密码保护的用户不支持找回密码
	</div>
		<%
		Else
			'If SendAnswer = "" and SendQuestion = "" and SendPassword1 = "" and SendPassword2 = "" Then%>
		<script language="javascript">	
		var ValidationPassed = true;
		function submitonce(theform)
		{
			if(theform.sendanswer.value=="")
			{
				alert("请输入你的提示答案!\n");
				ValidationPassed = false;
				theform.sendanswer.focus();
				return;
			}
			if(theform.sendpassword1.value=="")
			{
				alert("请输入你的密码!\n");
				ValidationPassed = false;
				theform.sendpassword1.focus();
				return;
			}
	
			if(theform.sendpassword2.value=="")
			{
				alert("请输入你的验证密码！\n");
				ValidationPassed = false;
				theform.sendpassword2.focus();
				return;
			}
	
			if(theform.sendpassword1.value!=theform.sendpassword2.value)
			{
				alert("你的两次密码输入不相同！\n");
				ValidationPassed = false;
				theform.sendpassword1.focus();
				return;
			}
			ValidationPassed = true;
			submit_disable(theform);
		}
		</script>
	<div class=title>
	请输入您的用户名及相关信息。</div>
	<form action=<%=DEF_BBS_HomeUrl%>User/UserGetPass.asp method="post" onSubmit="submitonce(this);return ValidationPassed;">
		<div class="value2">
		用户名称：<input name=SendUser type=text maxlength=20 size=22 value="<%=htmlencode(SendUser)%>" class="fminpt input_2">
		</div>
		<input type=hidden value="getpass" name=act>
		<div class="value2">
		密码提示：<input name=sendquestion value="<%=htmlencode(Question)%>" maxlength=14 size=22 readonly class="fminpt input_2">
		</div>
		<div class="value2">
		提示答案：<input name=sendanswer type=text maxlength=20 size=22 value="<%=htmlencode(SendAnswer)%>" class="fminpt input_2">
		</div>
		<div class="value2">
		新的密码：<input name=sendpassword1 type=password maxlength=14 size=22 value="<%=htmlencode(SendPassword1)%>" class="fminpt input_2">
		</div>
		<div class="value2">
		验证密码：<input name=sendpassword2 type=password maxlength=14 size=22 value="<%=htmlencode(SendPassword2)%>" class="fminpt input_2">
		</div>
		<%If DEF_EnableAttestNumber > 0 Then%>
			<div class="value2">验证码：<%
			displayVerifycode%></div><%
		End If%>
		<br />
		<input type=submit value="取回密码" class="fmbtn btn_3"> <input type=reset value="取消" class="fmbtn btn_2">
	</form>
		<%
			'End If
		End If
	
	End Sub
	
	Private Sub SendGetPassMail(Form_UserName,Form_Mail,pass,ActiveCode)
	
		Dim HomeUrl
		HomeUrl = "http://"&Request.ServerVariables("server_name")
		If Request.ServerVariables("SERVER_PORT") <> "80" Then HomeUrl = HomeUrl & ":" & Request.ServerVariables("SERVER_PORT")
		HomeUrl = Lcase(HomeUrl & Request.Servervariables("SCRIPT_NAME"))
		HomeUrl = Replace(HomeUrl,"user/usergetpass.asp","")
	
		Dim MailBody,Topic,TextBody
		Topic = "您在" & DEF_SiteNameString & "的密码找回"
		MailBody = "<html>"
		TextBody = ""
		MailBody = MailBody & "<title>账号信息</title>"
		MailBody = MailBody & "<BODY>"
		MailBody = MailBody & "<table BORDER=0 WIDTH=95% ALIGN=CENTER><TBODY><tr>"
		MailBody = MailBody & "<TD valign=MIDDLE ALIGN=TOP><HR WIDTH=100% SIZE=1>"
		TextBody = TextBody & "------------------------------------------" & VbCrLf
		MailBody = MailBody & VbCrLf & htmlencode(Form_UserName)&"，您好：<br><br>"
		TextBody = TextBody & htmlencode(Form_UserName)&"，您好：" & VbCrLf & VbCrLf
		MailBody = MailBody & "您在本论坛使用了密码找回，下面是您的账号信息！<br><br>"
		TextBody = TextBody & "您在本论坛使用了密码找回，下面是您的账号信息！" & VbCrLf & VbCrLf
		MailBody = MailBody & "用户名："&htmlencode(Form_UserName)&"<br>"
		TextBody = TextBody & "用户名："&htmlencode(Form_UserName) & VbCrLf
		MailBody = MailBody & "密　码：" & pass & "<br>"
		TextBody = TextBody & "密　码：" & pass & VbCrLf
		If ActiveCode <> "" Then
			MailBody = MailBody & "激活码：" & ActiveCode & "<br>"
			TextBody = TextBody & "激活码：" & ActiveCode & VbCrLf
		End If
		MailBody = MailBody & "<br><br>"
		MailBody = MailBody & "<CENTER><font COLOR=RED><a href=""" & HomeUrl & """>欢迎光临论坛！</a></font>"
		MailBody = MailBody & "</td></tr></table><br><HR WIDTH=95% SIZE=1>"
		MailBody = MailBody & "<p ALIGN=CENTER>" & DEF_SiteNameString & " <a href=http://www.leadbbs.com target=_blank class=NavColor>" & DEF_Version & "</a></P>"
		TextBody = TextBody & VbCrLf & DEF_BBS_HomeUrl & VbCrLf
		TextBody = TextBody & "------------------------------------------" & VbCrLf
		MailBody = MailBody & "</body>"
		MailBody = MailBody & "</html>"
		Select Case DEF_BBS_EmailMode
			Case 1: If SendEasyMail(Form_Mail,Topic,MailBody,TextBody) = 1 Then
						Response.Write "<br><br>资料成功发送到您的注册邮箱！"
					Else
						Response.Write "<br><br>论坛未正确设置邮件发送，资料发送失败！"
					End If
			Case 2: If SendJmail(Form_Mail,Topic,MailBody) = 1 Then
						Response.Write "<br><br>资料成功发送到您的注册邮箱！"
					Else
						Response.Write "<br><br>论坛未正确设置邮件发送，资料发送失败2！"
					End If
			Case 3: If SendCDOMail(Form_Mail,Topic,TextBody) = 1 Then
						Response.Write "<br><br>资料成功发送到您的注册邮箱！"
					Else
						Response.Write "<br><br>论坛未正确设置邮件发送，资料发送失败！"
					End If
			Case Else: 
		End Select
	
	End Sub

End Class%>