<!-- #include file=../../../inc/BBSsetup.asp -->
<!-- #include file=../../../inc/Board_Popfun.asp -->
<!-- #include file=../../../inc/Limit_fun.asp -->
<!-- #include file=../inc/BoardMaster_Fun.asp -->
<!-- #include file=../../../User/inc/Fun_SendMessage.asp -->
<%
DEF_BBS_HomeUrl = "../../../"
initDatabase
GBL_CHK_TempStr = ""

SiteHead(DEF_SiteNameString & " - " & DEF_PointsName(6) & "����")

UserTopicTopInfo
DisplayUserNavigate("����µ������û�")%>
<br><br><%If GBL_CHK_Flag=1 and BDM_isBoardMasterFlag = 1 and BDM_SpecialPopedomFlag = 1 Then
	LoginAccuessFul
Else%>
	<table width=96%>
	<tr>
	<td>
	<%
	If Request("submitflag")="" Then
		Response.Write "<br><b>���ȵ�¼</b>"
	Else
		Response.Write "<br><p align=left><font color=ff0000 class=redfont><b>" & GBL_CHK_TempStr & "</b></font>"
	End If
	DisplayLoginForm
	Response.Write "</p>"%>
	</td>
	</tr>
	</table>
<%End If
UserTopicBottomInfo
closeDataBase
SiteBottom
If GBL_ShowBottomSure = 1 Then Response.Write GBL_SiteBottomString

Dim GBL_UserName,GBL_Assort,GBL_ndatetime,GBL_WhyString,GBL_ExpiresTime
GBL_ExpiresTime = -1
Dim GBL_UserName_UserLimit,GBL_UserName_UserID

Function LoginAccuessFul

	GBL_UserName = Left(Request.Form("GBL_UserName"),20)
	GBL_ndatetime = GetTimeValue(DEF_Now)
	GBL_Assort = Left(Request("GBL_Assort"),14)
	GBL_WhyString = Left(Request.Form("GBL_WhyString"),100)
	GBL_ExpiresTime = Left(Request.Form("GBL_ExpiresTime"),14)

	If isNumeric(GBL_ExpiresTime) = 0 Then GBL_ExpiresTime = -1
	GBL_ExpiresTime = fix(cCur(GBL_ExpiresTime))
	If GBL_ExpiresTime < 0 or GBL_ExpiresTime > 30 Then GBL_ExpiresTime = -1

	If isNumeric(GBL_Assort) = 0 Then GBL_Assort = -1
	GBL_Assort = fix(cCur(GBL_Assort))
	',0-��֤��Ա,1-����,2-�ܰ���,3-���λ�Ա,4-���Ի�Ա,5-���޸Ļ�Ա,6-����ʽ��Ա
	If GBL_Assort <> 3 and GBL_Assort <> 4 and GBL_Assort <> 5 and GBL_Assort <> 6 Then
		GBL_Assort = -1
	End If

	If Request.Form("submitflag") <> "" Then
		CheckNewIP
		If GBL_CHK_TempStr = "" Then
			SaveNewIP
			If CheckSupervisorUserName = 0 Then
				CALL LDExeCute("Update LeadBBS_User Set LastWriteTime=" & GetTimeValue(DEF_Now) & " where ID=" & GBL_UserID,1)
			End If
			Response.Write GBL_CHK_TempStr
		Else
			DisplayNewIPForm
		End If
	Else
		DisplayNewIPForm
	End If

End Function

Function SaveNewIP

	Dim SQL,Rs,Number
	SQL = sql_select("Select ID from LeadBBS_SpecialUser where Assort=" & GBL_Assort & " and UserID=" & GBL_UserName_UserID,1)
	Set Rs = LDExeCute(SQL,0)
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
	Else
		Rs.Close
		Set Rs = Nothing
		CALL LDExeCute("Delete from LeadBBS_SpecialUser Where Assort=" & GBL_Assort & " and UserID=" & GBL_UserName_UserID,1)
		GBL_CHK_TempStr = "<br><br><font color=008800 class=greenfont>�����ݿ��д���һЩ����Ӧ���Ѿ��ɹ��޸���<br>" & VbCrLf
	End If
	
	SQL = "Insert Into LeadBBS_SpecialUser(UserID,UserName,BoardID,Assort,ndatetime,ExpiresTime,WhyString) Values(" & GBL_UserName_UserID & ",'" & Replace(GBL_UserName,"'","''") & "',0," & GBL_Assort & "," & GBL_ndatetime & "," & GBL_ExpiresTime & ",'" & Replace(GBL_WhyString,"'","''") & "')"
	CALL LDExeCute(SQL,0)
	GBL_CHK_TempStr = "<font color=008800 class=greenfont>�����ɹ���ɣ���ӳɹ�,�����Ѿ�֪ͨ��Ա��<br>" & VbCrLf

End Function

Function CheckNewIP

	If CheckWriteEventSpace = 0 Then
		Response.Write "<b><font color=Red Class=redfont>���Ĳ�����Ƶ�����Ժ������ύ!</font></b> <br>" & VbCrLf
		Exit Function
	End If
	If GBL_Assort <> 3 and GBL_Assort <> 4 and GBL_Assort <> 5 and GBL_Assort <> 6 Then
		GBL_CHK_TempStr = "���󣺻�Ա����ѡ���������ȷѡ��"
		Exit function
	End If
	
	If isNumeric(GBL_ExpiresTime) = 0 Then GBL_ExpiresTime = -1
	GBL_ExpiresTime = fix(cCur(GBL_ExpiresTime))
	If GBL_ExpiresTime < 0 or GBL_ExpiresTime > 30 Then GBL_ExpiresTime = -1
	If GBL_ExpiresTime = -1 Then
		GBL_CHK_TempStr = "������������ѡ���������ȷѡ��"
		Exit function
	End If

	If GBL_UserName = "" Then
		GBL_CHK_TempStr = "��������д�û�����"
		Exit function
	End If
		
	If CheckUserNameExist(GBL_UserName) = 0 Then
		Exit function
	End If

	If GBL_ExpiresTime > 0 Then
		GBL_ExpiresTime = GetTimeValue(DateAdd("d",GBL_ExpiresTime,DEF_Now))
	Else
		GBL_ExpiresTime = 0
	End If

End Function

Function DisplayNewIPForm

	Dim N
	If GBL_CHK_TempStr <> "" Then Response.Write "<p><font color=ff0000 class=redfont><b>" & GBL_CHK_TempStr & "</b></font></p>"%>
		  ���������������Ϣ
          <form action=NewSpecialUser.asp method=post id=fobform name=fobform>
          	�� �� ����<input name=GBL_UserName value="<%=htmlencode(GBL_UserName)%>" class=fminpt><br>
          	<input name=submitflag type=hidden value="LKOkxk2">
          	����ѡ��<select name=GBL_Assort>
          				<option value=-1>==��ѡ��==</option>
          				<option value=3<%If GBL_Assort = 3 Then Response.Write " selected"%>>�����û��ѷ��������</option>
          				<option value=4<%If GBL_Assort = 4 Then Response.Write " selected"%>>��ֹ�û�����������</option>
          				<option value=5<%If GBL_Assort = 5 Then Response.Write " selected"%>>��ֹ�û��޸����Ӻ���������</option>
          				<option value=6<%If GBL_Assort = 6 Then Response.Write " selected"%>>ǿ���û���Ϊδ�����û�</option>
          			</select><br>
          	��Чʱ�䣺<select name=GBL_ExpiresTime>
          					<%For N = 1 to 30
          						If N = GBL_ExpiresTime Then
          							Response.Write "<option value=" & N & " selected>��Ч��" & Right("0" & N,2) & "��</option>"
          						Else
          							Response.Write "<option value=" & N & ">��Ч��" & Right("0" & N,2) & "��</option>"
          						End If
          					Next%>
          					<option value=0<%If GBL_ExpiresTime = 0 Then Response.Write " Selected"%>>������Ч</option>
          				</select>
          				<br>
          	ԭ��ע����<input name=GBL_WhyString MaxLength=100 size=30 value="<%=htmlencode(GBL_WhyString)%>" class=fminpt>
          	<select onchange="document.fobform.GBL_WhyString.value=this.value;">
          		<option value="">=====һЩ����ԭ����ѡ��=====</option>
          		<option value="��������ɫ������">��������ɫ������</option>
          		<option value="����̳���ж��⹥��(�ڿ���Ϊ)">����̳���ж��⹥��(�ڿ���Ϊ)</option>
          		<option value="��ͣ�Ķ����ˮ��ע�����û�">��ͣ�Ķ����ˮ��ע�����û�</option>
          		<option value="�û����ֲ�����Ҫ��">�û����ֲ�����Ҫ��</option>
          		<option value="������̳�������в�����">������̳�������в�����</option>
          	</select>
          	<br><br>
          	<input type=submit value="�ύ" class=fmbtn> <input type=reset value="ȡ��" class=fmbtn></form>
          	<br>
          	<p>
<%End Function

Rem ���ĳ�û����Ƿ����
Function CheckUserNameExist(UserName)

	If UserName <> "" and inStr(UserName,",") = 0 and inStr(Lcase(DEF_SupervisorUserName),"," & Lcase(UserName) & ",") > 0 Then
		'��������ͬ����ʾ��Ϊ���Է�����Ա���ֱ�й©��ʵ��Ӧ����ʾ����Ա���ܱ�����
		GBL_CHK_TempStr = "�����û���" & htmlencode(UserName) & "�����ڣ�"
		CheckUserNameExist = 0
		Exit Function
	End If
	Dim Rs
	Set Rs = LDExeCute(sql_select("Select ID,UserLimit,UserName from LeadBBS_User where UserName='" & Replace(UserName,"'","''") & "'",1),0)
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		CheckUserNameExist = 0
		GBL_UserName_UserLimit = 0
		GBL_CHK_TempStr = "�����û���" & htmlencode(UserName) & "�����ڣ�"
		Exit Function
	Else
		GBL_UserName_UserLimit = cCur(Rs(1))
		GBL_UserName_UserID = cCur(Rs(0))
		GBL_UserName = Rs(2)
	End if
	Rs.Close
	Set Rs = Nothing
	',0-��֤��Ա,1-����,2-�ܰ���,3-���λ�Ա,4-���Ի�Ա,5-���޸Ļ�Ա,6-����ʽ��Ա
	Dim TmpStr
	Select Case GBL_Assort
		'Case 0: 
		'		If GetBinarybit(GBL_UserName_UserLimit,2) = 1 Then
		'			GBL_CHK_TempStr = "�����û�" & htmlencode(UserName) & "�Ѿ���" & DEF_PointsName(5) & "�������ظ���ӣ�"
		'			CheckUserNameExist = 0
		'			Exit Function
		'		Else
		'			GBL_UserName_UserLimit = SetBinaryBit(GBL_UserName_UserLimit,2,1)
		'		End If
		Case 3:
				If GetBinarybit(GBL_UserName_UserLimit,7) = 1 Then
					GBL_CHK_TempStr = "�����û�" & htmlencode(UserName) & "�ķ������ݼ�ǩ���Ѿ������Σ������ظ���ӣ�"
					CheckUserNameExist = 0
					Exit Function
				Else
					GBL_UserName_UserLimit = SetBinaryBit(GBL_UserName_UserLimit,7,1)
					TmpStr = "�������з��������Ѿ�������."
				End If
		Case 4:
				If GetBinarybit(GBL_UserName_UserLimit,3) = 1 Then
					GBL_CHK_TempStr = "�����û�" & htmlencode(UserName) & "�Ѿ������Լ����Ͷ���Ϣ�������ظ���ӣ�"
					CheckUserNameExist = 0
					Exit Function
				Else
					GBL_UserName_UserLimit = SetBinaryBit(GBL_UserName_UserLimit,3,1)
					TmpStr = "���Ѿ������Է���."
				End If
		Case 5:
				If GetBinarybit(GBL_UserName_UserLimit,4) = 1 Then
					GBL_CHK_TempStr = "�����û�" & htmlencode(UserName) & "�Ѿ�����ֹ�޸����Ӽ��������ϣ������ظ���ӣ�"
					CheckUserNameExist = 0
					Exit Function
				Else
					GBL_UserName_UserLimit = SetBinaryBit(GBL_UserName_UserLimit,4,1)
					TmpStr = "���Ѿ��������޸�."
				End If
		Case 6:
				If GetBinarybit(GBL_UserName_UserLimit,1) = 1 Then
					GBL_CHK_TempStr = "�����û�" & htmlencode(UserName) & "�Ѿ�����δ����״̬�������ظ���ӣ�"
					CheckUserNameExist = 0
					Exit Function
				Else
					GBL_UserName_UserLimit = SetBinaryBit(GBL_UserName_UserLimit,1,1)
					TmpStr = "��Ŀǰ����δ����."
				End If
		Case Else:
				GBL_CHK_TempStr = "�����û�" & htmlencode(UserName) & "�Ѿ�����δ����״̬�������ظ���ӣ�"
				CheckUserNameExist = 0
				Exit Function
	End Select
	If GBL_ExpiresTime > 0 Then
		Rs = GetTimeValue(DateAdd("d",GBL_ExpiresTime,DEF_Now))
	Else
		Rs = "������Ч"
	End If
	SendNewMessage GBL_CHK_User,UserName,"��̳���ţ�����Ȩ�޷����ı�֪ͨ","[color=blue]����Ȩ���������Ա�����������仯[/color][hr]" & VbCrLf &_
	"[b]����ԭ��[/b]" & GBL_WhyString & VbCrLf & _
	"[b]��Чֱ����[/b]" & Rs & VbCrLf & _
	"[b]���������[/b]" & TmpStr & VbCrLf,GBL_IPAddress
	GBL_CHK_TempStr = ""
	CALL LDExeCute("Update LeadBBS_User Set UserLimit=" & GBL_UserName_UserLimit & " where ID=" & GBL_UserName_UserID,1)
	CheckUserNameExist = 1

End Function%>