<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Board_popfun.asp -->
<!-- #include file=../../User/inc/UserTopic.asp -->
<%
DEF_BBS_homeUrl="../../"
Const PLUG_LeadCard_Length = 14 '���ų���
Const PLUG_LeadCard_ChangeValueLimit = 10 'ת��ʱҪ����˺ŵ��߲Ƹ�

Main

Sub Main

	InitDatabase
	BBS_SiteHead DEF_SiteNameString & " - LeadCard",0,"<span class=navigate_string_step>LeadCard</span>"
	
	UserTopicTopInfo("plug")
	
	If GBL_CHK_User = "" then
		Response.write "<div class=alert>��û��ʹ��LeadCard��Ȩ�ޣ����ȵ�½����ע��Ϊ��̳��Ա��</div>"
	Else
		Main_LeadCard
	End If
	
	CloseDatabase
	UserTopicBottomInfo
	SiteBottom

End Sub

Sub Main_LeadCard

	Dim SupervisorFlag
	SupervisorFlag = CheckSupervisorUserName
%>
	<div class=title>LeadCard ����״̬</div>
	<div class=value2>
	<%=DEF_PointsName(1) & "��" & GBL_CHK_CharmPoint%>
	 / <%=DEF_PointsName(0) & "��" & GBL_CHK_Points%>
	 / <%=DEF_PointsName(2) & "��" & GBL_CHK_CachetValue%></div>

	<%If Request.Form("submitflag") = "1" Then
		Select Case Request("act")
			Case "0": LeadCard_InputValue
			Case "1": If SupervisorFlag = 1 Then LeadCard_Made
			Case "11": LeadCard_ChangeValue
		End Select
	Else%>
		<br>
		<div class=title>��Ƭ��ֵ</div>
		<div class=value2>
		1.<a href=#none onclick="$id('CardUser2').value=$id('CardUser').value='<%=htmlencode(GBL_CHK_User)%>';">���Լ���ֵ</a>
		2.<a href=#none onclick="$id('CardUser2').value=$id('CardUser').value='';">�����ѳ�ֵ</a>
		</div>
		<form method=post action=Default.asp name=cardform id=cardform onSubmit="submit_disable(this);">
		<input type=hidden value=1 name=submitflag>
		<input type=hidden value=0 name=act>
		<div class=value2>��ֵ���ţ�<input maxlength=20 id=CardID name=CardID value="<%=Left(Request("CardID"),16)%>" size="20" class='fminpt input_3'>
		<%=PLUG_LeadCard_Length%>λ����
		</div>
		<div class=value2>�����˺ţ�<input maxlength=20 name=CardUser value="<%=Left(Request("CardUser"),20)%>" size="20" class='fminpt input_2'>
		��дҪ������˺�
		</div>
		<div class=value2>�ظ��˺ţ�<input maxlength=20 name=CardUser2 value="<%=Left(Request("CardUser2"),20)%>" size="20" class='fminpt input_2'>
		ȷ�ϳ�ֵ�˺�
		</div>
		<div class=value2>
		<input name=submit2 type=submit value="������ֵ" class='fmbtn btn_3'>
		</div>
		</form>

		<br>
		<div class=title>����<%=DEF_PointsName(1)%>ת��</div>
		
		<%If PLUG_LeadCard_ChangeValueLimit > 0 Then
			Response.Write "<div class=value2>�˻��������" & PLUG_LeadCard_ChangeValueLimit & " ��</div>"
		End If%>
		<div class=value2>����ת�˵����<%
		If GBL_CHK_CharmPoint - PLUG_LeadCard_ChangeValueLimit > 0 Then
			Response.Write GBL_CHK_CharmPoint-PLUG_LeadCard_ChangeValueLimit
		Else
			Response.Write "0"
		End If%> ��
		</div>
		<form method=post action=Default.asp name=cardform id=cardform onSubmit="submit_disable(this);">
		<input type=hidden value=1 name=submitflag>
		<input type=hidden value=11 name=act>
		<div class=value2>ת�����<input maxlength=20 name=ChangeValue value="<%=Left(Request("ChangeValue"),16)%>" size="20" class='fminpt input_2'>
		</div>
		<div class=value2>��Ҫת���<%=DEF_PointsName(1)%>��ֵ����һ������
		</div>
		<div class=value2>ת���˺ţ�<input maxlength=20 name=CardUser value="<%=Left(Request("CardUser"),20)%>" size="20" class='fminpt input_2'>
		��дҪת����˺�
		<div class=value2>
		�ظ��˺ţ�<input maxlength=20 name=CardUser2 value="<%=Left(Request("CardUser2"),20)%>" size="20" class='fminpt input_2'>
		ȷ��ת���˺�
		</div>
		<div class=value2>
		<input name=submit2 type=submit value="����ת��" class='fmbtn btn_3'>
		</div>
		</form>
		<%
		If SupervisorFlag = 1 Then%>
		<br><div class=title>����Ա�������³�ֵ��</div>
		<%LeadCard_MakeForm%>
		<br>
		<div class=title>����Ա����ֵ���б�</a></div><br>
		<ol>
		<li><a href=Default.asp>���100�ų�ֵ��</a>
		<li><a href=Default.asp?T=1>���100��<%=DEF_PointsName(0)%>��</a>
		<li><a href=Default.asp?T=2>���100��<%=DEF_PointsName(1)%>��</a>
		<li><a href=Default.asp?T=3>���100��<%=DEF_PointsName(2)%>��</a>
		<li><a href=Default.asp?T=4>���100��<%=DEF_PointsName(4)%>��</a>
		</ol><%
			LeadCard_List
		End If
	End If%>
<%

End Sub

Sub LeadCard_MakeForm

%>
		<form method=post action=Default.asp name=madeform id=madeform onSubmit="submit_disable(this);">
		<input type=hidden value=1 name=submitflag>
		<input type=hidden value=1 name=act>
		<div class=value2>����������<input maxlength=20 name=CardNum value="<%=Left(Request("CardNum"),16)%>" size="20" class='fminpt input_2'>
		�����³�ֵ������ һ�����1000��
		</div>
		<div class=value2>��ֵ���ͣ�<select name=CardType class=TBBG9>
			<option value=-1>��ѡ��
			<option value=1><%=DEF_PointsName(0)%>��(����<%=DEF_PointsName(0)%>)
			<option value=2><%=DEF_PointsName(1)%>��(����<%=DEF_PointsName(1)%>)
			<option value=3><%=DEF_PointsName(2)%>��(����<%=DEF_PointsName(2)%>)
			<option value=4><%=DEF_PointsName(4)%>��(����<%=DEF_PointsName(4)%>)
			<select> ��ͬ���͵ĳ�ֵ������Ӧ��ͬ��ֵ����
		</div>
		<div class=value2>
		��ֵ������<select name=CardPoints class=TBBG9>
			<option value=-1>��ѡ��
			<option value=1>1
			<option value=2>2
			<option value=5>5
			<option value=10>10
			<option value=25>25
			<option value=50>50
			<option value=100>100
			<option value=500>500
			<option value=1000>1000
			<option value=10000>10000
			</select> ��ֵ����ֵ��ɻ�ȡ����
		</div>
		<div class=value2>
		����ʱ�䣺<select name=ExpiresDate class=TBBG9>
			<option value=-1>��ѡ��
			<option value=1>1��
			<option value=7>1��
			<option value=30>1��
			<option value=120>3����
			<option value=365>1��
			<option value=3650>10��
			</select> �ڳ���ʱ���δ��Ŀ�Ƭ����ʧЧ
		</div>
		<div class=value2>
		<input name=submit2 type=submit value="��������" class='fmbtn btn_3'>
		</div>
		</form>
<%

End Sub

Function GBLFUN_Clng(Str)

	Dim Tmp
	Tmp = Left(Str & "",14)
	If isNumeric(Tmp) = 0 Then Tmp = 0
	Tmp = Fix(cCur(Tmp))
	GBLFUN_Clng = Tmp

End Function
	

Sub LeadCard_ChangeValue

	Dim ChangeValue,CardUser,CardUser2
	ChangeValue = Left(Request.Form("ChangeValue"),14)
	CardUser = Left(Trim(Request.Form("CardUser")),20)
	CardUser2 = Left(Trim(Request.Form("CardUser2")),20)

	ChangeValue = GBLFUN_Clng(ChangeValue)
	If ChangeValue < 0 or ChangeValue > 10000 Then
		LeadCard_Err "ת�˴�����������ȷ����ֵ��"
		Exit Sub
	End If
	
	If LCase(CardUser) <> LCase(CardUser2) Then
		LeadCard_Err "ת�˴������������ת���û�������ͬ��"
		Exit Sub
	End If
	
	If Trim(LCase(GBL_CHK_User)) = Trim(LCase(CardUser)) Then
		LeadCard_Err "ת�˴��󣺲���ת�˸��Լ���"
		Exit Sub
	End If
	
	Dim CardUserID
	CardUserID = CheckUserNameExist(CardUser)
	If CardUserID = 0 Then
		LeadCard_Err "ת�˴��󣺲������û� " & htmlencode(CardUser) & "��"
		Exit Sub
	End If

	Dim Rs
	Set Rs = Con.ExeCute(sql_select("Select CharmPoint from LeadBBS_User Where ID=" & GBL_UserID,1),0)
	If Rs.Eof Then
		LeadCard_Err "ת�˴������ȵ�¼��"
		Rs.Close
		Set Rs = Nothing
		Exit Sub
	End If
	GBL_CHK_CharmPoint = cCur(Rs(0))
	Rs.Close
	Set Rs = Nothing

	If ChangeValue > GBL_CHK_CharmPoint Then
		LeadCard_Err "ת�˴�������" & DEF_PointsName(1) & "���㣮"
		Exit Sub
	End If

	If GBL_CHK_CharmPoint <= PLUG_LeadCard_ChangeValueLimit or ChangeValue > (GBL_CHK_CharmPoint - PLUG_LeadCard_ChangeValueLimit) Then
		LeadCard_Err "ת�˴�������" & DEF_PointsName(1) & "δ������ת�˵���ֵ�򳬳�������ת�˵���ֵ��"
		Exit Sub
	End If

	Con.ExeCute("Update LeadBBS_User Set CharmPoint=CharmPoint-" & ChangeValue & " Where ID=" & GBL_UserID)
	Con.ExeCute("Update LeadBBS_User Set CharmPoint=CharmPoint+" & ChangeValue & " Where UserName='" & Replace(CardUser,"'","''") & "'")
	
	UpdateSessionValue 15,0-ChangeValue,1
	
	LeadCard_Done "ת�˳ɹ���ʾ���ɹ�ת���˺�<u>" & htmlencode(CardUser) & "</u>������" & DEF_PointsName(1) & "<u>" & ChangeValue & "</u>��"

End Sub

Sub LeadCard_InputValue

	Dim CardID,CardUser,CardUser2
	CardID = Left(Request.Form("CardID"),14)
	CardUser = Left(Trim(Request.Form("CardUser")),20)
	CardUser2 = Left(Trim(Request.Form("CardUser2")),20)

	CardID = GBLFUN_Clng(CardID)
	If Len(Cstr(CardID)) <> PLUG_LeadCard_Length Then
		LeadCard_Err "��ֵ���󣺿��Ŵ����޷���ɳ�ֵ��"
		Exit Sub
	End If
	
	If LCase(CardUser) <> LCase(CardUser2) Then
		LeadCard_Err "��ֵ������������ĳ�ֵ�û���ͬ���޷���ɳ�ֵ��"
		Exit Sub
	End If
	
	Dim CardUserID
	CardUserID = CheckUserNameExist(CardUser)
	If CardUserID = 0 Then
		LeadCard_Err "��ֵ���󣺲������û� " & htmlencode(CardUser) & "��"
		Exit Sub
	End If

	Dim Rs
	Dim CardType,ExpiresDate,CardPoints
	Set Rs = Con.ExeCute(sql_select("Select CardType,ExpiresDate,CardPoints from LeadBBS_Plug_Card Where CardID=" & CardID,1),0)
	If Rs.Eof Then
		LeadCard_Err "��ֵ���󣺿��� " & CardID & " �����ڻ��ѱ���ֵ��"
		Rs.Close
		Set Rs = Nothing
		Exit Sub
	End If
	CardType = Rs(0)
	ExpiresDate = Rs(1)
	CardPoints = Rs(2)
	Rs.Close
	Set Rs = Nothing

	If cCur(Left(GetTimeValue(DEF_Now),8)) > ExpiresDate Then
		LeadCard_Err "��ֵ���󣺿��� " & CardID & " �ѵ������ϣ�"
		Exit Sub
	End If

	Dim TypeCol,TypeStr
	Select Case CardType
	Case 1: TypeStr = DEF_PointsName(0)
		TypeCol = "Points"
	Case 2: TypeStr = DEF_PointsName(1)
		TypeCol = "CharmPoint"
	Case 3: TypeStr = DEF_PointsName(2)
		TypeCol = "CachetValue"
	Case 4: TypeStr = DEF_PointsName(4)
		TypeCol = "OnlineTime"
		CardPoints = CardPoints * 60
	Case Else: 
		LeadCard_Err "��ֵ���󣺿��� " & CardID & " ��������Ԥ֪�Ĵ�������ϵ����Ա�����"
		Exit Sub
	End Select
	
	Con.ExeCute("Update LeadBBS_User Set " & TypeCol & "=" & TypeCol & "+" & CardPoints & " Where ID=" & CardUserID)
	Con.ExeCute("Delete from LeadBBS_Plug_Card Where CardID=" & CardID)
	If CardType = 4 Then CardPoints = CardPoints / 60
	LeadCard_Done "��ֵ�ɹ���ʾ���ɹ�Ϊ�˺�<u>" & htmlencode(CardUser) & "</u>����" & TypeStr & "�������Ƶ���<u>" & CardPoints & "</u>��"
	

End Sub

Sub LeadCard_Made

	Dim CardNum,CardType,CardPoints,ExpiresDate
	CardNum = Request.Form("CardNum")
	CardType = Request.Form("CardType")
	CardPoints = Request.Form("CardPoints")
	ExpiresDate = Request.Form("ExpiresDate")
	
	CardNum = GBLFUN_Clng(CardNum)
	If CardNum < 0 or CardNum > 1000 Then
		LeadCard_Err "�����³�ֵ��������������������1-1000��"
		Exit Sub
	End If

	CardType = GBLFUN_Clng(CardType)
	If CardType < 1 or CardType > 4 Then
		LeadCard_Err "�����³�ֵ��������ѡ���ֵ�����ͣ�"
		Exit Sub
	End If

	CardPoints = GBLFUN_Clng(CardPoints)
	If CardPoints < 1 or CardPoints > 10000 Then
		LeadCard_Err "�����³�ֵ��������ѡ����ȷ�ĳ�ֵ��������"
		Exit Sub
	End If

	ExpiresDate = GBLFUN_Clng(ExpiresDate)
	If ExpiresDate < 1 or ExpiresDate > 3650 Then
		LeadCard_Err "�����³�ֵ����������ȷѡ���ֵ���������ڣ�"
		Exit Sub
	End If

	%>
	<br>
	<table cellpadding=0 cellspacing=0 class=table_in>
	<tr class=tbinhead>
	<td><div class=value>����</div></td>
	<td><div class=value>����</div></td>
	<td><div class=value>����</div></td>
	<td><div class=value>��������</div></td>
	</tr>
	<%
	Dim TypeStr
	
	Select Case CardType
	Case 1: TypeStr = DEF_PointsName(0)
	Case 2: TypeStr = DEF_PointsName(1)
	Case 3: TypeStr = DEF_PointsName(2)
	Case 4: TypeStr = DEF_PointsName(4)
	End Select

	Dim ExpiresDateTmp
	ExpiresDateTmp = Left(GetTimeValue(DateAdd("d",ExpiresDate,DEF_Now)),8)

	Dim CardID,N,Rs,Num
	Num = 0
	For N = 1 To CardNum
		Randomize
		CardID = Fix(Rnd*99999999999999)
		
		If Len(CardID) >= PLUG_LeadCard_Length Then
			Set Rs = Con.ExeCute(sql_select("Select CardID From LeadBBS_Plug_Card Where CardID=" & CardID,1),0)
			If Rs.Eof Then
				Rs.Close
				Set Rs = Nothing
				Num = Num + 1
				Con.ExeCute("Insert Into LeadBBS_Plug_Card(CardID,CardType,ExpiresDate,CardPoints) Values(" & CardID & "," & CardType & "," & ExpiresDateTmp & "," & CardPoints & ")")
				Response.Write "<tr><td class=tdbox>" & CardID & "</td>"
				Response.Write "<td class=tdbox>" & TypeStr & "��</td>"
				Response.Write "<td class=tdbox>" & CardPoints & "</td>"
				Response.Write "<td class=tdbox>" & ExpiresDate & "��</td></tr>" & VbCrLf
			Else
				Rs.Close
				Set Rs = Nothing
			End If
		End If
	Next
	%>
	</table>
	<%
	LeadCard_Done "��ֵ���ɹ����ɣ�	����" & Num & "�ţ�"

End Sub

Sub LeadCard_Err(str)

	Response.Write "<div class=alert>" & Str & "</div>"
	Response.Write "<a href=Default.asp>[����������]</a>"

End Sub

Sub LeadCard_Done(str)

	Response.Write "<div class='alert greenfont'>" & Str & "</div>"
	Response.Write "<a href=Default.asp>[����������]</a>"

End Sub

Rem ���ĳ�û����Ƿ����
Function CheckUserNameExist(UserName)

	Dim Rs
	Set Rs = LDExeCute(sql_select("Select ID from LeadBBS_User where UserName='" & Replace(UserName,"'","''") & "'",1),0)
	If Rs.Eof Then
		CheckUserNameExist = 0
	Else
		CheckUserNameExist = cCur(Rs(0))
	End if
	Rs.Close
	Set Rs = Nothing

End Function

Sub LeadCard_List

	Dim Rs,GetData
	Dim T
	T = Left(Request("T"),14)
	T = GBLFUN_Clng(T)
	If T < 1 or T > 4 Then T = 0
	If T = 0 Then
		Set Rs = Con.ExeCute(sql_select("Select ID,CardID,CardType,ExpiresDate,CardPoints From LeadBBS_Plug_Card",100),0)
	Else
		Set Rs = Con.ExeCute(sql_select("Select ID,CardID,CardType,ExpiresDate,CardPoints From LeadBBS_Plug_Card where CardType=" & T,100),0)
	End If
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		LeadCard_Done "û�з��ϲ�ѯ�����ĳ�ֵ����"
		Exit Sub
	End If
	
	GetData = Rs.GetRows(100)
	Rs.Close
	Set Rs = Nothing
		
	%>
	<table cellpadding=0 cellspacing=0 class=table_in>
	<tr class=tbinhead>
	<td><div class=value>���</div></td>
	<td><div class=value>����</div></td>
	<td><div class=value>����</div></td>
	<td><div class=value>����</div></td>
	<td><div class=value>����</div></td>
	</tr>
	<%
	Dim TypeStr

	Dim N,CardNum
	
	CardNum = Ubound(GetData,2)
	For N = 0 To CardNum	
		Select Case GetData(2,N)
		Case 1: TypeStr = DEF_PointsName(0)
		Case 2: TypeStr = DEF_PointsName(1)
		Case 3: TypeStr = DEF_PointsName(2)
		Case 4: TypeStr = DEF_PointsName(4)
		End Select

		Response.Write "<tr><td class=tdbox>" & GetData(0,N) & "</td>"
		Response.Write "<td class=tdbox>" & GetData(1,N) & "</td>"
		Response.Write "<td class=tdbox>" & TypeStr & "��</td>"
		Response.Write "<td class=tdbox>" & GetData(4,N) & "</td>"
		Response.Write "<td class=tdbox>" & GetData(3,N) & "</td></tr>" & VbCrLf
	Next
	%>
	</table>
	<%

End Sub%>