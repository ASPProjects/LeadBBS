<%Class User_UserActive

	Private AttestNumber

	Private Sub Class_Initialize
	
		AttestNumber = Left(Request.Form("AttestNumber"),14)
		If isNumeric(AttestNumber) = False Then AttestNumber = 0
		AttestNumber = Fix(cCur(AttestNumber))
	
	End Sub

	Public Sub DisplayActive

		If GetBinarybit(GBL_CHK_UserLimit,1) = 0 and GBL_CHK_UserLimit <> "" and GBL_CHK_TempStr = "" and GBL_UserID > 0 Then
			Response.Write "<div class=alert>���û��Ѿ���������ٽ��м��������</div>" & VbCrLf
			Exit Sub
		End If
		If Request.Form("act") = "active" and GBL_CHK_TempStr = "" Then
			If GBL_CHK_Flag = 1 and GBL_UserID > 0 Then
				If GetBinarybit(GBL_CHK_UserLimit,1) = 0 and GBL_CHK_UserLimit <> "" Then
					Response.Write "<div class=alert>���û��Ѿ���������ٽ��м��������</div>" & VbCrLf
				Else
					Dim Rs,SQL
					SQL = "Select BoardID from LeadBBS_SpecialUser where UserID=" & GBL_UserID
					Set Rs = LDExeCute(SQL,0)
					If Rs.Eof Then
						Rs.Close
						Set Rs = Nothing
						Response.Write "<div class=alert>���û��޷����û����м������ϵ����Ա.</div>" & VbCrLf
						Exit Sub
					End If
					SQL = cCur(Rs(0))
					Rs.Close
					Set Rs = Nothing
					If SQL < 1 Then
						Response.Write "<div class=alert>���û��޷����û����м������ϵ����Ա.</div>" & VbCrLf
					Else
						If AttestNumber = SQL Then
							CALL LDExeCute("Delete from LeadBBS_SpecialUser where UserID=" & GBL_UserID,1)
							CALL LDExeCute("Update LeadBBS_User Set UserLimit=" & SetBinarybit(GBL_CHK_UserLimit,1,0) & " where ID=" & GBL_UserID,1)
							UpdateSessionValue 2,SetBinarybit(GBL_CHK_UserLimit,1,0),0
							Response.Write "<div class='alert greenfont'>�û��ɹ����<a href=Login.asp?User=" & urlencode(GBL_CHK_User) & "&Relogin=Yes>�����µ�¼������̳</a>.</div>" & VbCrLf
						Else
							Randomize
							CALL LDExeCute("Update LeadBBS_User Set Prevtime=" & GetTimeValue(DEF_Now) & ",Login_IP='" & Replace(GBL_IPAddress,"'","''") & "',Login_lastpass='" & Fix(rnd*99999) & "',Login_falsenum=Login_falsenum+1 Where ID=" & GBL_UserID,1)
							UpdateSessionValue 11,GetTimeValue(DEF_Now),0
							UpdateSessionValue 7,Fix(rnd*99999),0
							UpdateSessionValue 8,1,1
							Response.Write "<div class=alert>���������ע���û�����ʧ�ܣ�</div>"
							VierForm
						End If
					End If
				End If
			Else
				If GBL_CHK_TempStr = "" Then
					If GBL_CHK_User = "" or GBL_CHK_Pass = "" Then GBL_CHK_TempStr = "�û���������������д����"
				End If
				'Response.Write "<div class=alert>" & GBL_CHK_TempStr & "</div>"
				VierForm
			End If
		Else
			VierForm
		End If
	
	End Sub
	
	Private Sub VierForm%>

	<div class='alert redfont'><%=GBL_CHK_TempStr%></div>
	<div class=title>��������Ҫ������˺ţ����뼰�����롣</div>
	<form action=<%=DEF_BBS_HomeUrl%>User/UserGetPass.asp method="post" onSubmit="submit_disable(this);">
		<div class="value2">�û����� <input name=User type=text maxlength=20 size=22 value="<%
		If GBL_CHK_user = "" or isNull(GBL_CHK_user) Then
			Response.Write htmlencode(Request("user"))
		Else
			Response.Write htmlencode(GBL_CHK_user)
		End If%>" class='fminpt input_2'></div>
		<input name=act type=hidden value="active">
		<div class="value2">�ܡ��룺 <input name=pass type=password maxlength=20 size=22 value="<%'=htmlencode(GBL_CHK_Pass)%>" class='fminpt input_2'>
		</div>
		<div class="value2">
		�����룺 <input name=AttestNumber type=text maxlength=10 size=22 value="<%If AttestNumber > 0 Then Response.Write AttestNumber%>" class="fminpt input_2">
		</div>
		<br /><input type=submit value="�����û�" class="fmbtn btn_3"> <input type=reset value="ȡ��" class="fmbtn btn_2">
	</form>
	<br />
	<div class=title>˵����</div>
	<ul><li>��������˺���Ҫ�����������������������ݣ�����������˺Ű�ť��</li>
	<li>��������ע��ʱ���Ѿ����͵��������䣬������ȷ�ļ�������ܼ��������˺š�</li>
	<li>ĳЩ�˺�ֻ���ɹ���Ա���ܼ�����ô˹��ܽ���Ȼ�޷����</li>
	</ul>
	<%
		If DEF_User_GetPassMode = 2 Then
			Response.Write "<br><a href=UserGetPass.asp><font color=red class=redfont><b>����������δ�յ��������ż�������ʹ�������һع���Ҫ���ٴη��ͣ�</b></font></a>"
		End If
	
	End Sub
	
End Class%>