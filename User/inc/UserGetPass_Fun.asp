<%Class User_GetPass

	Private SendEmail,SendPass,Question,Answer,SendAnswer,SendUser,SendQuestion,SendPassword1,SendPassword2

	Public Sub GetPass
	
		VierGetPassForm

	End Sub

	Private Sub VierGetPassForm
	
		SendUser = Trim(Request.Form("SendUser"))
		If DEF_User_GetPassMode = 0 Then
			Response.Write "<div class=alert>��̳�Ѿ��ر������һع��ܡ�</div>" & VbCrLf
			Exit Sub
		End If
		If DEF_BBS_EmailMode = 0 and SendUser  <> "" and DEF_User_GetPassMode = 2 Then
			Response.Write "<div class=alert>��̳��ֹ�����ʼ��������һع��ܲ���ʹ�á�</div>" & VbCrLf
			Exit Sub
		End If
		If Request.Form("act") = "getpass" Then
			SendAnswer = Left(Request.Form("SendAnswer"),20)
			SendQuestion = Left(Request.Form("SendQuestion"),20)
			SendPassword2 = Left(Request.Form("SendPassword2"),14)
			SendPassword1 = Left(Request.Form("SendPassword1"),14)
	
			If Len(SendUser) > 30 Then
				Response.Write "<div class=alert>����: �û���̫��.</div>" & VbCrLf
				Exit Sub
			End If
					
			SQL = "Select LastDoingTime from LeadBBS_OnlineUser where SessionID=" & Session.SessionID
			Set Rs = LDExeCute(SQL,0)
			If Rs.Eof Then
				GBL_CHK_LastWriteTime = 0
				Rs.Close
				Set Rs = Nothing
				Response.Write "<div class=alert>������2�������Ϻ��ٽ��������һ�!</div>" & VbCrLf
				Exit Sub
			Else
				GBL_CHK_LastWriteTime = DateDiff("s",RestoreTime(Rs(0)),DEF_Now)
				If GBL_CHK_LastWriteTime < 0 Then GBL_CHK_LastWriteTime = DEF_WriteEventSpace + 1
			End If
	
			If CheckWriteEventSpace = 0 Then
				Response.Write "<div class=alert>���Ĳ�����Ƶ�����Ժ������ύ!</div>" & VbCrLf
				DisplaySubmitForm
				Exit Sub
			End If
			Dim Rs,SQL
			SQL = "Select mail,Pass,Question,Answer,UserLimit from LeadBBS_User where UserName='" & Replace(SendUser,"'","''") & "'"
			Set Rs = LDExeCute(SQL,0)
			If Rs.Eof Then
				Rs.Close
				Set Rs = Nothing
				Response.Write "<div class=alert>����: �����ڵ��û���.</div>" & VbCrLf
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
				Response.Write "<div class=alert>����: ���û�ע��ʱδ�ṩEmail���޷��һ�����.</div>" & VbCrLf
				Exit Sub
			End If
	
			SQL = GBL_CHK_User
			GBL_CHK_User = SendUser
			CheckisBoardMaster
			If Len(Answer) < 32 or GBL_BoardMasterFlag >= 4 or (GBL_CHK_User <> "" and inStr(GBL_CHK_User,",") = 0 and inStr(LCase(DEF_SupervisorUserName),"," & LCase(GBL_CHK_User) & ",") > 0) Then
				Response.Write "<div class=alert>����: ���û�δ�������뱣�����޷�ʹ���һ�����ָ����ܡ�<br>����ϵ����Ա�趨���뱣����</div>" & VbCrLf
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
					Response.Write "<div class=alert>��֤����д����!</div>" & VbCrLf
					DisplaySubmitForm
					Exit Sub
				End If
	
				If Len(SendPassword2) < DEF_UserShortestPassword or Len(SendPassword1) < DEF_UserShortestPassword or SendPassword1 <> SendPassword2 Then
					Response.Write "<div class=alert>�µ����벻������4λ����������������֤���������ͬ��</div>" & VbCrLf
					SendQuestion = ""
					DisplaySubmitForm
					Exit Sub
				End If
			
				If MD5(SendAnswer) <> Answer and Mid(MD5(SendAnswer),9,16) <> Answer Then
					Response.Write "<div class=alert>�������ʾ����д�����޷��һ��˺ţ�������ϵ����Ա!</div>" & VbCrLf
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
				Response.Write "<div class=alert>�����Ѿ��ɹ����ģ���ʹ���µ�����<a href=""Login.asp?User=" & urlencode(SendUser) & """>��¼</a>�����˺�!</div>" & VbCrLf
				If DEF_User_GetPassMode = 2 and GetBinarybit(GBL_CHK_UserLimit,1) = 1 Then
					SQL = "Select BoardID from LeadBBS_SpecialUser where UserID=" & GBL_UserID
					Set Rs = LDExeCute(SQL,0)
					If Rs.Eof Then
						Rs.Close
						Set Rs = Nothing
						Response.Write "<div class=alert>���û��޷����û����м��ֹͣ�����ʼ�������ϵ����Ա.</div>" & VbCrLf
						Exit Sub
					End If
					SQL = cCur(Rs(0))
					Rs.Close
					Set Rs = Nothing
					SendGetPassMail SendUser,SendEmail,SendPassword2,SQL
					Response.Write "<div class='alert greenfont'>ͬʱ���µ����뼰�������Ѿ����͵�����ע�����䡣</div>" & VbCrLf
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
	<div class=title>�����������û�����</div>
	<form action=<%=DEF_BBS_HomeUrl%>User/UserGetPass.asp method="post" onSubmit="submit_disable(this);">
		�û���: <input name=SendUser type=text maxlength=20 size=22 value="<%
		If GBL_CHK_user = "" or isNull(GBL_CHK_user) Then
			Response.Write htmlencode(Request("user"))
		Else
			Response.Write htmlencode(GBL_CHK_user)
		End If%>" class=fminpt><br>
		<input type=hidden value="getpass" name=act><br>
		<input type=submit value="ȡ������" class="fmbtn btn_3"> <input type=reset value="ȡ��" class="fmbtn btn_2">
	</form>
	<br>
	<div class=value2>ע�⣺ ���������û���δ�������뱣�����û���֧���һ�����
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
				alert("�����������ʾ��!\n");
				ValidationPassed = false;
				theform.sendanswer.focus();
				return;
			}
			if(theform.sendpassword1.value=="")
			{
				alert("�������������!\n");
				ValidationPassed = false;
				theform.sendpassword1.focus();
				return;
			}
	
			if(theform.sendpassword2.value=="")
			{
				alert("�����������֤���룡\n");
				ValidationPassed = false;
				theform.sendpassword2.focus();
				return;
			}
	
			if(theform.sendpassword1.value!=theform.sendpassword2.value)
			{
				alert("��������������벻��ͬ��\n");
				ValidationPassed = false;
				theform.sendpassword1.focus();
				return;
			}
			ValidationPassed = true;
			submit_disable(theform);
		}
		</script>
	<div class=title>
	�����������û����������Ϣ��</div>
	<form action=<%=DEF_BBS_HomeUrl%>User/UserGetPass.asp method="post" onSubmit="submitonce(this);return ValidationPassed;">
		<div class="value2">
		�û����ƣ�<input name=SendUser type=text maxlength=20 size=22 value="<%=htmlencode(SendUser)%>" class="fminpt input_2">
		</div>
		<input type=hidden value="getpass" name=act>
		<div class="value2">
		������ʾ��<input name=sendquestion value="<%=htmlencode(Question)%>" maxlength=14 size=22 readonly class="fminpt input_2">
		</div>
		<div class="value2">
		��ʾ�𰸣�<input name=sendanswer type=text maxlength=20 size=22 value="<%=htmlencode(SendAnswer)%>" class="fminpt input_2">
		</div>
		<div class="value2">
		�µ����룺<input name=sendpassword1 type=password maxlength=14 size=22 value="<%=htmlencode(SendPassword1)%>" class="fminpt input_2">
		</div>
		<div class="value2">
		��֤���룺<input name=sendpassword2 type=password maxlength=14 size=22 value="<%=htmlencode(SendPassword2)%>" class="fminpt input_2">
		</div>
		<%If DEF_EnableAttestNumber > 0 Then%>
			<div class="value2">��֤�룺<%
			displayVerifycode%></div><%
		End If%>
		<br />
		<input type=submit value="ȡ������" class="fmbtn btn_3"> <input type=reset value="ȡ��" class="fmbtn btn_2">
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
		Topic = "����" & DEF_SiteNameString & "�������һ�"
		MailBody = "<html>"
		TextBody = ""
		MailBody = MailBody & "<title>�˺���Ϣ</title>"
		MailBody = MailBody & "<BODY>"
		MailBody = MailBody & "<table BORDER=0 WIDTH=95% ALIGN=CENTER><TBODY><tr>"
		MailBody = MailBody & "<TD valign=MIDDLE ALIGN=TOP><HR WIDTH=100% SIZE=1>"
		TextBody = TextBody & "------------------------------------------" & VbCrLf
		MailBody = MailBody & VbCrLf & htmlencode(Form_UserName)&"�����ã�<br><br>"
		TextBody = TextBody & htmlencode(Form_UserName)&"�����ã�" & VbCrLf & VbCrLf
		MailBody = MailBody & "���ڱ���̳ʹ���������һأ������������˺���Ϣ��<br><br>"
		TextBody = TextBody & "���ڱ���̳ʹ���������һأ������������˺���Ϣ��" & VbCrLf & VbCrLf
		MailBody = MailBody & "�û�����"&htmlencode(Form_UserName)&"<br>"
		TextBody = TextBody & "�û�����"&htmlencode(Form_UserName) & VbCrLf
		MailBody = MailBody & "�ܡ��룺" & pass & "<br>"
		TextBody = TextBody & "�ܡ��룺" & pass & VbCrLf
		If ActiveCode <> "" Then
			MailBody = MailBody & "�����룺" & ActiveCode & "<br>"
			TextBody = TextBody & "�����룺" & ActiveCode & VbCrLf
		End If
		MailBody = MailBody & "<br><br>"
		MailBody = MailBody & "<CENTER><font COLOR=RED><a href=""" & HomeUrl & """>��ӭ������̳��</a></font>"
		MailBody = MailBody & "</td></tr></table><br><HR WIDTH=95% SIZE=1>"
		MailBody = MailBody & "<p ALIGN=CENTER>" & DEF_SiteNameString & " <a href=http://www.leadbbs.com target=_blank class=NavColor>" & DEF_Version & "</a></P>"
		TextBody = TextBody & VbCrLf & DEF_BBS_HomeUrl & VbCrLf
		TextBody = TextBody & "------------------------------------------" & VbCrLf
		MailBody = MailBody & "</body>"
		MailBody = MailBody & "</html>"
		Select Case DEF_BBS_EmailMode
			Case 1: If SendEasyMail(Form_Mail,Topic,MailBody,TextBody) = 1 Then
						Response.Write "<br><br>���ϳɹ����͵�����ע�����䣡"
					Else
						Response.Write "<br><br>��̳δ��ȷ�����ʼ����ͣ����Ϸ���ʧ�ܣ�"
					End If
			Case 2: If SendJmail(Form_Mail,Topic,MailBody) = 1 Then
						Response.Write "<br><br>���ϳɹ����͵�����ע�����䣡"
					Else
						Response.Write "<br><br>��̳δ��ȷ�����ʼ����ͣ����Ϸ���ʧ��2��"
					End If
			Case 3: If SendCDOMail(Form_Mail,Topic,TextBody) = 1 Then
						Response.Write "<br><br>���ϳɹ����͵�����ע�����䣡"
					Else
						Response.Write "<br><br>��̳δ��ȷ�����ʼ����ͣ����Ϸ���ʧ�ܣ�"
					End If
			Case Else: 
		End Select
	
	End Sub

End Class%>