<%
Dim LimitBoardStringData,LimitBoardStringDataNum
Rem 1.������,2.ֻ��Էǰ���,3.ֻ��Էǰ���,4.ֻ��Էǰ���,5.ֻ��Էǰ���,6,������,7.������,8.���԰���,9.������̳-���δ��¼�û�,10��������,11.������,12.�������,13.�˰������ӷ�����Ҫ��֤,14.�������Ӱ����ظ����빺����,15.ֻ��רҵ�û�����,16.Ĭ�ϱ༭ģʽ: 0ΪĬ���趨ֵ(����������ָ��). 1Ϊ��Ĭ���趨ֵ(����������ָ��)��ͬ�ı༭ģʽ 17.�Ƿ�ظ����� 18.ֱ����ʾר��,19.�Ӱ����Լ��ʾ.20.�Ӱ�����ʾ�ڵͲ�,21.�鵵�Ƿ��ֹ(1.��ֹ),22.��ʾ�����Ա���(��ǰ̨������ʾ),23.�Ƿ����ѡ��ר��
LimitBoardStringData = Array("ֻ�е�¼�û����ܷ���","ֻ��" & DEF_PointsName(5) & "����","��ֹ����������","�������޸���̳����","������ɾ����̳����","��ֹ�ظ�����","ֻ��" & DEF_PointsName(8) & "���Ͽ���","������ת������","������̳","��������" & DEF_PointsName(8) & "��������","��������" & DEF_PointsName(8) & "�ظ�����","��Ϊ������̳","������Ҫ��˲�����ʾ","��������������","ֻ��" & DEF_PointsName(10) & "����","�༭ģʽ(��ѡ��ʾ����̳����������ָ����Ĭ�ϱ༭��ʽ��ͬ)","�ظ�����(��ѡ��ʾ����̳����������ָ����Ĭ�������෴)","ֱ����ʾר����","�Ӱ����Լ��ʾ","�Ӱ����õͲ���ʾ","��ֹ�鵵","��ʾ��˵�ֱ����ʾ","��������ѡ��ר��")
LimitBoardStringDataNum = Ubound(LimitBoardStringData,1)

Dim LimitUserStringData,LimitUserStringDataNum
Rem 1.������,2.������,3.������,4.������,����ͬʱ�����޸��Լ�����,5.���԰���,6.ֻ��԰�����Ч,7.������,8.�Ƿ�����̳����,9.�Ƿ��������ת�����ӵ�������̳,10.�Ƿ����ܰ���,11.��Ϊ�����ܰ���,12.������ܰ���,13������,14.�Ƿ�������,15.רҵ�û�,16.����HTML.�κ��û�����Ч 17.������,�κ��û���Ч 18��������ӣ��ܰ�������)
LimitUserStringData = Array("δ�����û�",DEF_PointsName(5),"��ֹ���Ժͷ��Ͷ���Ϣ","��ֹ�޸ĸ������Ϻ���������","��ֹɾ������","��ֹ��������","���з�������",DEF_PointsName(8),"��ֹת������",DEF_PointsName(6),"ɾ���ϴ�����","����Ȩ��","�����պ��Ѷ���Ϣ",DEF_PointsName(7),"�Ƿ�" & DEF_PointsName(10),"����HTML��ֱ�Ӳ���ý��","��ֹ������ʾ����Ϣ","רְ���Ա/��������")
LimitUserStringDataNum = Ubound(LimitUserStringData,1)
Dim GBL_BoardMasterFlag
GBL_BoardMasterFlag = 0

Sub CheckisBoardMaster

	'If GBL_CheckPassDoneFlag = 0 Then CheckPass
	'6-�������
	If CheckSupervisorUserName = 1 Then
		GBL_BoardMasterFlag = 9 '����Ա
		Exit Sub
	End If
	If GetBinarybit(GBL_CHK_UserLimit,10) = 1 Then
		GBL_BoardMasterFlag = 7 '�ܰ���
		Exit Sub
	End If
	If GetBinarybit(GBL_CHK_UserLimit,14) = 1 Then
		If GBL_Board_MasterList = "?LeadBBS?" or inStr("," & GBL_Board_AssortMaster & ",","," & GBL_CHK_User & ",") > 0 Then
			GBL_BoardMasterFlag = 6 '��������
			Exit Sub
		Else
			GBL_BoardMasterFlag = 4 '������,���Ǳ���
		End If
	End If
	If GetBinarybit(GBL_CHK_UserLimit,8) = 1 Then
		If GBL_Board_MasterList = "?LeadBBS?" or inStr("," & GBL_Board_MasterList & ",","," & GBL_CHK_User & ",") > 0 Then
			GBL_BoardMasterFlag = 5 '�������
			Exit Sub
		Else
			GBL_BoardMasterFlag = 4 '����
		End If
	End If
	If GBL_BoardMasterFlag >= 4 Then Exit Sub
	If GetBinarybit(GBL_CHK_UserLimit,2) = 1 Then
		GBL_BoardMasterFlag = 2 '��֤�û�
	Else
		GBL_BoardMasterFlag = 0 '�ǰ���
	End If

End Sub

Function CheckBoardReAnnounceLimit

	If GetBinarybit(GBL_Board_BoardLimit,12) = 1 Then
		GBL_CHK_TempStr = "�˰������ڷ�����̳��������˲�����" & VbCrLf
		CheckBoardReAnnounceLimit = 0
	ElseIf GetBinarybit(GBL_Board_BoardLimit,6) = 1 Then
		GBL_CHK_TempStr = "�˰���" & LimitBoardStringData(5) & "��" & VbCrLf
		CheckBoardReAnnounceLimit = 0
	ElseIf GetBinarybit(GBL_Board_BoardLimit,11) = 1 and GBL_BoardMasterFlag < 5 Then
		GBL_CHK_TempStr = "�˰���" & LimitBoardStringData(10) & "��" & VbCrLf
		CheckBoardReAnnounceLimit = 0
	End If
	CheckBoardReAnnounceLimit = 1

End Function

Function CheckBoardAnnounceLimit

	If GetBinarybit(GBL_Board_BoardLimit,12) = 1 Then
		GBL_CHK_TempStr = "�˰������ڷ�����̳��������˲�����" & VbCrLf
		CheckBoardAnnounceLimit = 0
	ElseIf GetBinarybit(GBL_Board_BoardLimit,3) = 1 Then
		GBL_CHK_TempStr = "�˰���" & LimitBoardStringData(2) & "��" & VbCrLf
		CheckBoardAnnounceLimit = 0
	ElseIf GetBinarybit(GBL_Board_BoardLimit,10) = 1 and GBL_BoardMasterFlag < 5 Then
		GBL_CHK_TempStr = "�˰���" & LimitBoardStringData(9) & "��" & VbCrLf
		CheckBoardAnnounceLimit = 0
	End If
	CheckBoardAnnounceLimit = 1

End Function

Function CheckUserAnnounceLimit

	If GetBinarybit(GBL_CHK_UserLimit,7) = 1 Then
		GBL_CHK_TempStr = "������" & LimitUserStringData(2) & "�У����س�����Щ������" & VbCrLf
		CheckUserAnnounceLimit = 0
		Exit Function
	End If
	If GetBinarybit(GBL_CHK_UserLimit,1) = 1 Then
		GBL_CHK_TempStr = "��Ŀǰ����" & LimitUserStringData(0) & "״̬������<a href=""" & DEF_BBS_HomeUrl & "User/UserGetPass.asp?act=active"">����</a>��ȴ�������Ա��ˡ�" & VbCrLf
		CheckUserAnnounceLimit = 0
		Exit Function
	End If
	If GetBinarybit(GBL_CHK_UserLimit,3) = 1 Then
		GBL_CHK_TempStr = "���Ѿ���" & LimitUserStringData(2) & "��ͶƱ�Ȳ�����" & VbCrLf
		CheckUserAnnounceLimit = 0
		Exit Function
	End If
	CheckUserAnnounceLimit = 1

End Function

Function CheckBoardModifyLimit

	If GetBinarybit(GBL_Board_BoardLimit,4) = 1 Then
		GBL_CHK_TempStr = "�˰���" & LimitBoardStringData(3) & "��" & VbCrLf
		CheckBoardModifyLimit = 0
		Exit Function
	End If
	CheckBoardModifyLimit = 1

End Function

Function CheckUserModifyLimit

	If GetBinarybit(GBL_CHK_UserLimit,4) = 1 Then
		GBL_CHK_TempStr = "���Ѿ���" & LimitUserStringData(3) & "��" & VbCrLf
		CheckUserModifyLimit = 0
		Exit Function
	End If
	CheckUserModifyLimit = 1

End Function

Function GetBinaryString(Number)

	Dim Temp1,Temp2,TempN
	Temp2 = Number
	Temp1 = ""
	For TempN = BinaryDataNum+1 to 1 step -1
		If Temp2 >= BinaryData(TempN-1) Then
			Temp1 = Temp1 & "1"
			Temp2 = Temp2 - BinaryData(TempN-1)
		Else
			Temp1 = Temp1 & "0"
		End If
	Next
	GetBinaryString = Temp1

End Function

Function SetBinarybit(Number,bit,value)

	Dim Temp
	Temp = GetBinarybit(Number,bit)

	If Temp = value Then
		SetBinarybit = Number
	ElseIf Temp = 1 and  value = 0 Then
		SetBinarybit = cCur(Number) - BinaryData(Bit-1)
	ElseIf Temp = 0 and  value = 1 Then
		SetBinarybit = cCur(Number) + BinaryData(Bit-1)
	End If

End Function

Sub CheckAccessLimit_TimeLimit

	If GBL_Board_ID < 1 Then Exit Sub
	If (GBL_Board_StartTime <> "000000" or GBL_Board_EndTime <> "000000")  Then
		Dim T1,t2,t3
		t1 = int(Mid(GBL_Board_StartTime,1,2))
		t2 = int(Mid(GBL_Board_EndTime,1,2))
		If t1 <> 0 or t2 <> 0 Then
			t3 = hour(DEF_Now)
			If t2 >= t1 Then
				if t3 >= t1 and t3 <=t2 Then
					GBL_CHK_TempStr = "�˰���ÿ�� " & t1 & ":00 �� " & t2 & ":59  ��ʱ�ر�,����ʱ��" & DEF_Now & "��" & VbCrLf
					Exit Sub
				End If
			Else
				If (t3 >= t1 and t3 <=23) or (t3 >=0 and t3 <=t2) Then
					GBL_CHK_TempStr = "�˰���ÿ�� " & t1 & ":00 �� ����" & t2 & ":59  ��ʱ�ر�,����ʱ��" & DEF_Now & "��" & VbCrLf
					Exit Sub
				End If
			End If
		End If
		t1 = int(Mid(GBL_Board_StartTime,3,2))
		t2 = int(Mid(GBL_Board_EndTime,3,2))
		If t1 <> 0 or t2 <> 0 Then
			t3 = weekday(DEF_Now,2)
			If t2 >= t1 Then
				if t3 >= t1 and t3 <=t2 Then
					GBL_CHK_TempStr = "�˰���ÿ�� " & t1 & " - " & t2 & " �ر���,����������" & t3  & "��" & VbCrLf
					Exit Sub
				End If
			Else
				If (t3 >= t1 and t3 <=7) or (t3 >=1 and t3 <=t2) Then
					GBL_CHK_TempStr = "�˰���ÿ��" & t1 & "�����գ���һ����" & t2 & " �ر���,����������" & t3  & "��" & VbCrLf
					Exit Sub
				End If
			End If
		End If
		t1 = int(Mid(GBL_Board_StartTime,5,2))
		t2 = int(Mid(GBL_Board_EndTime,5,2))
		If t1 <> 0 or t2 <> 0 Then
			t3 = day(DEF_Now)
			If t2 >= t1 Then
				if t3 >= t1 and t3 <=t2 Then
					GBL_CHK_TempStr = "�˰���ÿ�� " & t1 & "�� - " & t2 & "�� �ر���,������" & t3  & "�š�" & VbCrLf
					Exit Sub
				End If
			Else
				If (t3 >= t1 and t3 <=31) or (t3 >=1 and t3 <=t2) Then
					GBL_CHK_TempStr = "�˰���ÿ�� " & t1 & "�ŵ��µף�һ�ŵ�" & t2 & "�� �ر���,������" & t3  & "�š�" & VbCrLf
					Exit Sub
				End If
			End If
		End If
	End If

End Sub

Sub CheckAccessLimit

	Dim Temp
	If GBL_Board_ID < 1 Then Exit Sub
	If GBL_UserID > 0 and CheckSupervisorUserName = 1 Then Exit Sub
	
	If GBL_Board_OtherLimit > 0 Then
		If GBL_Board_OtherLimit < 100 Then
			Temp = 0
		Else
			Temp = cCur(Left(GBL_Board_OtherLimit,Len(GBL_Board_OtherLimit)-2))
		End If
		Select Case CCur(Right(GBL_Board_OtherLimit,2))
			Case 1: If GBL_CHK_Points < Temp Then GBL_CHK_TempStr = "���" & DEF_PointsName(0) & "ֵ���㣬���ʴ˰�����Ҫ" & Temp & DEF_PointsName(0) & "��" & VbCrLf
			Case 2: If (GBL_CHK_OnlineTime/60) < Temp Then GBL_CHK_TempStr = "���" & DEF_PointsName(4) & "ֵ���㣬���ʴ˰�����Ҫ" & Temp & DEF_PointsName(4) & "ֵ��" & VbCrLf
			Case 3: If GBL_CHK_CharmPoint < Temp Then GBL_CHK_TempStr = "���" & DEF_PointsName(1) & "ֵ���㣬���ʴ˰�����Ҫ" & Temp & DEF_PointsName(1) & "ֵ��" & VbCrLf
			Case 4: If GBL_CHK_CachetValue < Temp Then GBL_CHK_TempStr = "���" & DEF_PointsName(2) & "ֵ���㣬���ʴ˰�����Ҫ" & Temp & DEF_PointsName(2) & "ֵ��" & VbCrLf
			Case 5: If isArray(GBL_UDT) Then
					If inStr(GBL_UDT(19),"," & Cstr(Temp) & ",") = 0 Then
						GBL_CHK_TempStr = "�˰���ֻ�����ض�" & DEF_PointsName(9) & "[���" & Temp & "]���ʡ�" & VbCrLf
					End If
				Else
					 GBL_CHK_TempStr = "���ʴ˰�����" & DEF_PointsName(9) & "���ơ�" & VbCrLf
				End If
		End Select
		If GBL_CHK_TempStr <> "" Then Exit Sub
	End If

	If GetBinarybit(GBL_Board_BoardLimit,7) = 1 Then
		If GBL_BoardMasterFlag < 4 or GBL_UserID < 1 Then
			GBL_CHK_TempStr = "�˰���" & LimitBoardStringData(6) & "��" & VbCrLf
			Exit Sub
		End If
	End If

	If GBL_CHK_GuestFlag = 1 and GetBinarybit(GBL_Board_BoardLimit,1) = 1 and GBL_CHK_Flag = 0 Then
		GBL_CHK_TempStr = "�˰���" & LimitBoardStringData(0) & "������<a href=""" & DEF_BBS_HomeUrl & "User/Login.asp?u=" & urlencode(Request.Servervariables("SCRIPT_NAME") & "?" & Request.QueryString) & """>��¼</a>��<a href=""" & DEF_BBS_HomeUrl & "User/" & DEF_RegisterFile & """>ע��</a>���û���" & VbCrLf
		Exit Sub
	End If

	If GetBinarybit(GBL_Board_BoardLimit,2) = 1 Then
		If GetBinarybit(GBL_CHK_UserLimit,2) = 0 or GBL_UserID < 1 Then
			GBL_CHK_TempStr = "�˰���" & LimitBoardStringData(1) & "��" & VbCrLf
			Exit Sub
		End If
	End If

	If GetBinarybit(GBL_Board_BoardLimit,15) = 1 Then
		If GetBinarybit(GBL_CHK_UserLimit,15) = 0 or GBL_UserID < 1 Then
			GBL_CHK_TempStr = "�˰���" & LimitBoardStringData(14) & "��" & VbCrLf
			Exit Sub
		End If
	End If

	If GBL_CHK_TempStr <> "" Then Exit Sub
	If GBL_Board_HiddenFlag = 2 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "�˰����Ѿ��ر�,��ֹ�����" & VbCrLf
		Exit Sub
	End If

	If GBL_Board_ForumPass <> "" Then
		If GBL_UserID < 1 Then
			GBL_CHK_TempStr = GBL_CHK_TempStr & "�������û���ݵ�¼!" & VbCrLf
			Exit sub
		End If

		If CheckWriteEventSpace = 0 Then
			GBL_CHK_TempStr = "���Ĳ�����Ƶ�����Ժ�����!" & VbCrLf
			Exit sub
		End If
		If GBL_Board_ForumPass <> DecodeCookie(Left(Request.Cookies(DEF_MasterCookies & "_" & GBL_UserID)("Board_" & GBL_board_ID),255)) Then
			GBL_CHK_TempStr = GBL_CHK_TempStr & " " & VbCrLf
				%>
				<div class="alertbox">
				<%
				Dim ForumPass
				If Request("submitflag") <> "" Then
					ForumPass = Request.form("ForumPass")
					Dim NumCheck
					NumCheck = CheckRndNumber
					If ForumPass = GBL_Board_ForumPass and NumCheck = 1 Then
						Response.Cookies(DEF_MasterCookies & "_" & GBL_UserID)("Board_" & GBL_board_ID) = CodeCookie(ForumPass)
						Response.Cookies(DEF_MasterCookies & "_" & GBL_UserID).Expires = DEF_Now + 365
						Response.Cookies(DEF_MasterCookies & "_" & GBL_UserID).Domain = DEF_AbsolutHome
						Response.Write "<span class=""title greenfont"">��¼�ɹ�</span>"
						Response.Write "<br /><br />-- ���� <a href=""http://" & Request.ServerVariables("SERVER_NAME")&Request.ServerVariables("URL") & "?" & Request.QueryString & """>" & htmlencode(Request.ServerVariables("SERVER_NAME")&Request.ServerVariables("URL")) & "</a>" & VbCrLf
					Else
						If NumCheck = 0 Then
							Response.Write "<span class=""alert redfont"">��֤����д����!</span>" & VbCrLf
						Else
							Response.Write "<span class=""alert redfont"">�����������!</span>" & VbCrLf
						End If
						Call LDExeCute("Update LeadBBS_User Set LastWriteTime=" & GetTimeValue(DEF_Now) & " where ID=" & GBL_UserID,1)
						DisplayPassWordLoginForm
					End If
				Else
					Response.Write "<span class=""title"">����̳Ϊ������̳����������ȷ����֤��Ϣ��</span>" & VbCrLf
					DisplayPassWordLoginForm
				End If
				%>
			<%
				Exit Sub
		End If
	End If

End Sub

Function CheckRndNumber
	If DEF_EnableAttestNumber = 0 Then
		CheckRndNumber = 1
		Exit Function
	End If

	Dim RndNumber
	RndNumber = Left(Session(DEF_MasterCookies & "RndNum") & "",4)
	If RndNumber = "" Then
		Randomize
		RndNumber = Fix(Rnd*9999)+1
		Session(DEF_MasterCookies & "RndNum") = RndNumber
	End If

	Dim ForumNumber
	If dontRequestFormFlag = "" Then
		ForumNumber = Left(Request.form("ForumNumber"),4)
	Else
		ForumNumber = Left(GetFormData("ForumNumber"),4)
	End If
	If LCase(RndNumber) = LCase(ForumNumber) Then
		CheckRndNumber = 1
	Else
		CheckRndNumber = 0
	End If

End Function

Sub DisplayPassWordLoginForm

	Dim Temp
	Temp = Request.ServerVariables("URL")
	Temp = StrReverse(Temp)
	Temp = Replace(Temp,"\","/")
	if Instr(Temp,"/") > 0 Then Temp = Left(Temp,Instr(Temp,"/")-1)
	Temp = StrReverse(Temp)
	%>
	<form action="<%=Temp%>?<%=Request.QueryString%>" method="post">
		<div class=value2>�ܡ��룺 <input name="ForumPass" type="password" maxlength="20" size="20" value="<%=htmlencode(Request("ForumPass"))%>" class="fminpt input_2" />
		</div><%If DEF_EnableAttestNumber > 0 Then%>
		<div class=value2>��֤�룺 <%
			displayVerifycode
		End If%>
		</div>
		<input name="submitflag" type="hidden" value="ddddls-+++" />
		<div class=value2>
		<input type="submit" value="��¼" class="fmbtn btn_2"> <input type="reset" value="ȡ��" class="fmbtn btn_2" />
		</div>
	</form>
	<%

End Sub%>

<%
Sub displayVerifycode

	Dim Url
	Url = filterUrlstr(Left(Request.QueryString("dir"),100))
	if Url = "" and dontRequestFormFlag = "" then
		Url = filterUrlstr(Left(Request.form("dir"),100))
	end if
	If Url = "" Then
		Url = DEF_BBS_HomeUrl
	End If
%>
		<input name="ForumNumber" id="ForumNumber" maxlength="4" value="<%=htmlencode(Session(DEF_MasterCookies & "RndNum_par") & "")%>" onfocus="verify_load(0,'<%=url%>');" class="fminpt input_1" />
		<img src="<%=Url%>images/blank.gif" id="verifycode" align="middle" onclick="verify_load(1,'<%=url%>');" /> 
		<a href="javascript:;" id=verify_click onclick="this.style.display='none';verify_load(1,'<%=url%>');return false;">�����ʾ��֤��</a>
		<noscript>     
		<div class="verifycode"><img src="<%=Url%>User/number.asp?r=1" id="verifycode" class="verifycode" align="middle" onclick="verify_load(1,'<%=url%>');" /></div>
		</noscript>
<%End Sub%>