<!-- #include file=../../../inc/BBSsetup.asp -->
<!-- #include file=../../../inc/Board_popfun.asp -->
<!-- #include file=../../../inc/Upload_Setup.asp -->
<!-- #include file=../../../inc/Limit_fun.asp -->
<!-- #include file=../inc/BoardMaster_Fun.asp -->
<!-- #include file=../../../User/inc/Fun_SendMessage.asp -->
<%
DEF_BBS_HomeUrl = "../../../"
initDatabase
GBL_CHK_TempStr = ""
CheckisBoardMasterFlag

BBS_SiteHead DEF_SiteNameString & " - ע�����û�",0,"<span class=navigate_string_step>" & DEF_PointsName(6) & "����</span>"

Dim LMT_Action

rem for special user
Dim GBL_UserName,GBL_Assort,GBL_ndatetime,GBL_WhyString,GBL_ExpiresTime
GBL_ExpiresTime = -1
Dim GBL_UserName_UserLimit,GBL_UserName_UserID

rem for fob ip
Dim GBL_IPStart,GBL_IPEnd
Dim GBL_AnnounceID,GBL_MessageID

rem for modifyuser
Dim GBL_ModifyMode,GBL_UserName_FaceUrl
Dim GBL_UserName_UnderWrite,GBL_UserName_UserTitle
GBL_ModifyMode = 0


If GBL_CHK_Flag=1 and BDM_isBoardMasterFlag = 1 and BDM_SpecialPopedomFlag = 1 Then
	LMT_Action = Request("action")
	Select Case LMT_Action
		Case "specialuser"
			Select Case Left(Request("GBL_Assort"),14)
				Case "4"
					UserTopicTopInfo(4)
				Case "5"
					UserTopicTopInfo(5)
				Case Else
					UserTopicTopInfo(3)
			End Select
			NewSpecialUser
		Case "fobip"
			UserTopicTopInfo(6)
			DisplayNewForbidIP
		Case "modifyuser"
			UserTopicTopInfo(7)
			DisplayModifyUser
		Case "clear"
			UserTopicTopInfo(10)
			View_ClearExpiresInfo
		Case Else
			LMT_Action = ""
			UserTopicTopInfo(2)
			SpecialUserBrowser
	End Select
Else
	UserTopicTopInfo(0)
	If Request("submitflag")="" Then
		DisplayLoginForm("���ȵ�¼")
	Else
		DisplayLoginForm("<span class=""redfont"">" & GBL_CHK_TempStr & "</span>")
	End If
End If
UserTopicBottomInfo
closeDataBase
SiteBottom
If GBL_ShowBottomSure = 1 Then Response.Write GBL_SiteBottomString

Function SpecialUserBrowser

	GBL_CHK_TempStr=""
	Dim Rs,SQL
	Dim UpDownPageFlag
	UpDownPageFlag = Request("UpDownPageFlag")
	
	Dim Assort
	Assort = Left(Request.QueryString("Assort"),14)
	If isNumeric(Assort) = 0 Then Assort = 3
	Assort = Fix(cCur(Assort))
	If Assort < 3 or Assort > 6 then Assort = 3

	Dim Start,key
	'Dim recordCount
	'recordCount=0
	
	Dim SQLendString

	Start = Left(Trim(Request("Start")),14)
	Start = start
	If isNumeric(Start)=0 or Start="" Then Start=0
	Start = cCur(Start)
	key = Request.Form("key")
	If key="" Then key = Request("key")

	Dim SQLCountString,whereFlag
	SQLendString=""
	SQLendString = " where T1.Assort=" & Assort
	whereFlag = 1

	If key<>"" Then
		If whereFlag = 1 Then
			SQLendString = SQLendString & " and T1.UserName like'" & Replace(key,"'","''") & "%'"
		Else
			SQLendString = SQLendString & " where T1.UserName like'" & Replace(key,"'","''") & "%'"
			whereFlag = 1
		End If
	End If
	SQLCountString = SQLendString
	If UpDownPageFlag = "1" and Start>0 then
		If whereFlag = 1 Then
			SQLendString = SQLendString & " and T1.ID<" & Start
		Else
			SQLendString = SQLendString & " where T1.ID<" & Start
			whereFlag = 1
		End If
	Else
		If whereFlag = 1 Then
			SQLendString = SQLendString & " and T1.ID>" & Start
		Else
			SQLendString = SQLendString & " where T1.ID>" & Start
			whereFlag = 1
		End If
	end If
	
	If UpDownPageFlag = "1" then
		'If DEF_IDFocusFlag<> 2 Then SQLendString = SQLendString & " Order by  T1.ID DESC"
		SQLendString = SQLendString & " Order by  T1.ID DESC"
	Else
		'If DEF_IDFocusFlag<> 1 Then SQLendString = SQLendString & " Order by  T1.ID ASC"
		SQLendString = SQLendString & " Order by  T1.ID ASC"
	End If
	
	Dim MaxRecordID,MinRecordID
	MaxRecordID = 0

	SQL = "select Max(T1.id) from LeadBBS_SpecialUser as T1 " & SQLCountString
	Set Rs = LDExeCute(SQL,0)
	
	If not Rs.Eof Then
		If Rs(0) <> "" Then
			MaxRecordID = cCur(Rs(0))
		Else
			MaxRecordID = 0
		End If
	End If
	Rs.Close
	Set Rs = Nothing
	
	SQL = "select Min(id) from LeadBBS_SpecialUser as T1 " & SQLCountString
	Set Rs = LDExeCute(SQL,0)

	If not Rs.Eof Then
		If Rs(0) <> "" Then
			MinRecordID = cCur(Rs(0))
		else
			MinRecordID = 0
		end If
	End If

	Rs.Close
	Set Rs = Nothing

	Dim FirstID,LastID

	SQL = sql_select("select T1.ID,T1.UserID,T1.UserName,T1.ndatetime,T1.Assort,t2.BoardName,T1.BoardID,T1.WhyString,T1.ExpiresTime from LeadBBS_SpecialUser as T1 left join LeadBBS_Boards as T2 on T1.BoardID=T2.BoardID" & SQLendString,DEF_MaxListNum)

	Set Rs = LDExeCute(SQL,0)
	Dim Num
	Dim GetData
	If Not rs.Eof Then
		GetData = Rs.GetRows(-1)
		Num = Ubound(GetData,2) + 1
	Else
		Num = 0
	End If
	Rs.close
	Set Rs = Nothing
	
	
	Dim i,N,DoStr

	DoStr = LimitUserManage_NavInfo(Assort)

	If Num>0 Then
		i=1
	
		Dim MinN,MaxN,StepValue
		SQL = ubound(getdata,2)
		If UpDownPageFlag = "1" then
			MinN = SQL
			MaxN = 0
			StepValue = -1
		Else
			MinN = 0
			MaxN = SQL
			StepValue = 1
		End If
		
		LastID = cCur(GetData(0,MaxN))
		FirstID = cCur(GetData(0,MinN))
		
		
		Dim EndwriteQueryString,PageSplictString
		EndwriteQueryString = "?Assort=" & Assort
		If key<>"" Then EndwriteQueryString = EndwriteQueryString & "&key=" & urlencode(key)
	
		PageSplictString = PageSplictString & "<div class=j_page>"
		if FirstID>MinRecordID and FirstID<>0 then
			PageSplictString = PageSplictString & "<a href=LimitUserManage.asp" & EndwriteQueryString & "&Start=0&SubmitFlag=3829EwoqIaNfoG>��ҳ</a> " & VbCrLf
		else
			'PageSplictString = PageSplictString & "<font color=999999 class=grayfont>��ҳ</font> " & VbCrLf
		end if
	
		if FirstID > MinRecordID and FirstID<>0 then
			PageSplictString = PageSplictString & " <a href=LimitUserManage.asp" & EndwriteQueryString & "&Start=" & FirstID & "&UpDownPageFlag=1&SubmitFlag=3829EwoqIaNfoG>��ҳ</a> " & VbCrLf
		else
			'PageSplictString = PageSplictString & " <font color=999999 class=grayfont>��ҳ</font> " & VbCrLf
		end if
	
		if LastID<MaxRecordID and LastID<>0 then
			PageSplictString = PageSplictString & " <a href=LimitUserManage.asp" & EndwriteQueryString & "&Start=" & LastID & "&SubmitFlag=3829EwoqIaNfoG>��ҳ</a> " & VbCrLf
		else
			'PageSplictString = PageSplictString & " <font color=999999 class=grayfont>��ҳ</font> " & VbCrLf
		end if
	
		if LastID < MaxRecordID and LastID<>0 then
			PageSplictString = PageSplictString & " <a href=LimitUserManage.asp" & EndwriteQueryString & "&Start=" & MaxRecordID+1 & "&UpDownPageFlag=1&SubmitFlag=3829EwoqIaNfoG>βҳ</a> " & VbCrLf
		else
			'PageSplictString = PageSplictString & " <font color=999999 class=grayfont>βҳ</font> " & VbCrLf
		end if
		'PageSplictString = PageSplictString & "��<b>" & recordCount & "</b>����Ϣ"
		'If (recordCount mod DEF_MaxListNum)=0 Then
		'	PageSplictString = PageSplictString & " ��<b>" & clng(recordCount/DEF_MaxListNum) & "</b>ҳ"
		'Else
		'	If recordCount>=DEF_MaxListNum Then
		'		PageSplictString = PageSplictString & " ��<b>" & clng(recordCount/DEF_MaxListNum)+1 & "</b>ҳ"
		'	Else
		'		PageSplictString = PageSplictString & " ��<b>1</b>ҳ"
		'	End If
		'End If
		'PageSplictString = PageSplictString & " ÿҳ<b>" & DEF_MaxListNum & "</b>����¼"
		PageSplictString = PageSplictString & "</div>"
		Dim ColN
		ColN = 6
		If Assort <> 1 and Assort <> 6 Then ColN = 5
		%>
		<script language="JavaScript" type="text/javascript">
		function kill(killID)
		{
			window.open('DelSpecialUser.asp?'+killID,'','width=450,height=37,scrollbars=auto,status=no');
		}
		</script>
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class=table_in>
		<tr class=tbinhead>		
		<td width=50%>
		<form action=LimitUserManage.asp?assort=<%=assort%> method=post>
		�û�����<input size=6 name=key value="<%=htmlencode(key)%>" class="fminpt input_1"> <input type=submit name=submit value=���� class="fmbtn btn_1"></form>
		</td>
		<td align=right width=50%>
		<div class=value><%=PageSplictString%></div>
		</td></tr></table>
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class=table_in>

		<tr class=tbinhead>
			<td width=64><div class=value>ID</div></td>
			<td width=122><div class=value>����</div></td>
			<td width=82><div class=value>����ʱ��</div></td>
			<td width=64><div class=value>����</div></td><%If Assort = 1 Then%>
			<td width=104><div class=value>����</div></td><%End If
			If Assort = 6 Then%>
			<td width=80><div class=value>������</div></td><%End If%>
			<td><div class=value>˵������Чʱ��</div></td>
		</tr>
<%
		for n= MinN to MaxN Step StepValue
			%>
		<tr>
			<td class=tdbox width=48><%=GetData(0,n)%></td>
			<td class=tdbox>
				<a href=<%=DEF_BBS_HomeUrl%>User/LookUserInfo.asp?id=<%=GetData(1,n)%>><%=htmlencode(GetData(2,n))%></a>
				<a href='javascript:kill("GBL_UserName=<%=GetData(2,n)%>&GBL_Assort=<%=Assort%>");'><font color=008800 class=greenfont><%=DoStr%></font></a></td>
			<td class=tdbox><%=RestoreTime(Left(GetData(3,n),8))%></td>
			<td class=tdbox><%Select Case GetData(4,n)
				Case 0: Response.Write DEF_PointsName(5)
				Case 1: Response.Write "����"
				Case 2: Response.Write DEF_PointsName(6)
				Case 3: Response.Write "���η���"
				Case 4: Response.Write "��ֹ����"
				Case 5: Response.Write "��ֹ�޸�"
				Case 6: Response.Write "�ȴ���֤"
				End Select%></td><%If Assort = 1 Then%>
			<td class=tdbox><a href=../ForumBoard/ForumBoardModify.asp?GBL_ModifyID=<%=GetData(6,n)%>><%=GetData(5,n)%></a></td><%End If
			If Assort = 6 Then
				If cCur(GetData(6,n)) = 0 Then
					Response.Write "<td width=80 class=tdbox>��</td>"
				Else%>
			<td class=tdbox><a href=../../User/UserGetPass.asp?act=active&user=<%=htmlencode(GetData(2,n))%>><%=GetData(6,n)%></a></td><%
				End If
			End If%>
			<td class=tdbox><%
			If GetData(7,n) <> "" Then Response.Write htmlencode(GetData(7,n)) & "<br>"
			If cCur(GetData(8,n)) > 0 Then
				Response.Write "<font color=gray class=grayfont>���ڣ�" & RestoreTime(GetData(8,n))
			Else
				Response.Write "<font color=gray class=grayfont>������Ч"
			End If%>	</td>
                    </tr><%
			i=i+1
			if i>DEF_MaxListNum then exit for
		next
%>
                  </table>
		<%=PageSplictString%>
	<%
	Else
		Response.Write "<br>" & GBL_CHK_TempStr & "		<p>������ؼ�¼��" & VbCrLf
	End If

End Function


Function LimitUserManage_NavInfo(Assort)

	Dim DoStr
	DoStr = "����"

	Response.Write "<div class='user_item_nav fire'><ul>"
	Response.Write "<li><div class=name>�����û�����</div></li>"
	If Assort = 3 Then
		DoStr = "���"
		Response.Write "	<li><div class=navactive><span>���η���</span></div></li>"
	Else
		Response.Write "	<li><a href=LimitUserManage.asp?assort=3>���η���</a></li>"
	End If

	If Assort = 4 Then
		DoStr = "���"
		Response.Write "	<li><div class=navactive>��ֹ����</div></li>"
	Else
		Response.Write "	<li><a href=LimitUserManage.asp?assort=4>��ֹ����</a></li>"
	End If

	If Assort = 5 Then
		DoStr = "���"
		Response.Write "	<li><div class=navactive>��ֹ�޸�</div></li>"
	Else
		Response.Write "	<li><a href=LimitUserManage.asp?assort=5>��ֹ�޸�</a></li>"
	End If

	If Assort = 6 Then
		DoStr = "����"
		Response.Write "	<li><div class=navactive>δ�����û�</div></li>"
	Else
		Response.Write "	<li><a href=LimitUserManage.asp?assort=6>δ�����û�</a></li>"
	End If

	Response.Write "</ul></div>"
	LimitUserManage_NavInfo = DoStr

End Function


Function NewSpecialUser

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
	',0-��֤��Ա,1-����,2-�ܰ���,3-�����û�,4-�����û�,5-���޸��û�,6-����ʽ�û�
	If GBL_Assort <> 3 and GBL_Assort <> 4 and GBL_Assort <> 5 and GBL_Assort <> 6 Then
		GBL_Assort = -1
	End If

	If Request.Form("submitflag") <> "" Then
		CheckNewSpecialUser
		If GBL_CHK_TempStr = "" Then
			SaveNewSpecialUser
			If CheckSupervisorUserName = 0 Then
				CALL LDExeCute("Update LeadBBS_User Set LastWriteTime=" & GetTimeValue(DEF_Now) & " where ID=" & GBL_UserID,1)
			End If
			Response.Write GBL_CHK_TempStr
		Else
			DisplayNewSpecialUserForm
		End If
	Else
		DisplayNewSpecialUserForm
	End If

End Function

Function SaveNewSpecialUser

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
	GBL_CHK_TempStr = "<font color=008800 class=greenfont>�����ɹ���ɣ���ӳɹ�,�����Ѿ�֪ͨ�û���<br>" & VbCrLf

End Function

Function CheckNewSpecialUser

	If CheckWriteEventSpace = 0 Then
		Response.Write "<b><font color=Red Class=redfont>���Ĳ�����Ƶ�����Ժ������ύ!</font></b> <br>" & VbCrLf
		Exit Function
	End If
	If GBL_Assort <> 3 and GBL_Assort <> 4 and GBL_Assort <> 5 and GBL_Assort <> 6 Then
		GBL_CHK_TempStr = "������ʾ���û�����ѡ���������ȷѡ��"
		Exit function
	End If
	
	If isNumeric(GBL_ExpiresTime) = 0 Then GBL_ExpiresTime = -1
	GBL_ExpiresTime = fix(cCur(GBL_ExpiresTime))
	If GBL_ExpiresTime < 0 or GBL_ExpiresTime > 30 Then GBL_ExpiresTime = -1
	If GBL_ExpiresTime = -1 Then
		GBL_CHK_TempStr = "������ʾ����������ѡ���������ȷѡ��"
		Exit function
	End If

	If GBL_UserName = "" Then
		GBL_CHK_TempStr = "������ʾ������д�û�����"
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

Function DisplayNewSpecialUserForm

	Dim N
	If GBL_CHK_TempStr <> "" Then Response.Write "<p><font color=ff0000 class=redfont><b>" & GBL_CHK_TempStr & "</b></font></p>"%>
	<div class="title">�û�Ȩ�޲�����</div>
          <form action=LimitUserManage.asp method=post id=fobform name=fobform>
          	<div class="value2">��д�û���<input name=GBL_UserName value="<%=htmlencode(GBL_UserName)%>" class=fminpt></div>
          	<input name=submitflag type=hidden value="LKOkxk2">
          	<input name=action type=hidden value="specialuser">
          	<div class="value2">
          	����ѡ��<select name=GBL_Assort>
          				<option value=-1>==��ѡ��==</option>
          				<option value=3<%If GBL_Assort = 3 Then Response.Write " selected"%>>�����û��ѷ��������</option>
          				<option value=4<%If GBL_Assort = 4 Then Response.Write " selected"%>>��ֹ�û�����������</option>
          				<option value=5<%If GBL_Assort = 5 Then Response.Write " selected"%>>��ֹ�û��޸����Ӻ͸�������</option>
          				<option value=6<%If GBL_Assort = 6 Then Response.Write " selected"%>>ǿ���û���Ϊδ�����û�</option>
          			</select>
          	</div>
          	<div class="value2">
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
		</div>
		<div class="value2">
          	ԭ��ע����<input name=GBL_WhyString MaxLength=100 size=30 value="<%=htmlencode(GBL_WhyString)%>" class=fminpt>
          	<select onchange="document.fobform.GBL_WhyString.value=this.value;">
          		<option value="">=====һЩ����ԭ����ѡ��=====</option>
          		<option value="��������Υ��">��������Υ��</option>
          		<option value="����̳���ж��⹥��">����̳���ж��⹥��</option>
          		<option value="�����ˮ">�����ˮ</option>
          		<option value="�û����ֲ�����Ҫ��">�û����ֲ�����Ҫ��</option>
          		<option value="������̳����">������̳����</option>
          	</select>
          	</div>
          	<div class="value2">
          	<input type=submit value="�ύ" class="fmbtn btn_2"> <input type=reset value="ȡ��" class="fmbtn btn_2">
          	</div></form>
          	<p>
          	<div class="title">ע�ͣ�</div>
          	<div class="value2">
          	<ol>
          	<li>�����û��ѷ�������ݣ��˲��������θ��û����е���̳��������</li>
          	<li>��ֹ�û����������۾����˲�������ֹ���û����Ͷ���Ϣ������ͶƱ���������ӵȹ���</li>
          	<li>��ֹ�û��޸����Ӻ͸������ϣ��˲�������ֹ���û��޸��Ѿ�����������Ӽ���������</li>
          	<li>ǿ���û���Ϊδ�����û������û����³�Ϊδ�����û�����ֻ�й�����Ա�������¼���</li>
          	</ol>
          	</div>
<%End Function

Rem ���ĳ�û����Ƿ����
Function CheckUserNameExist(UserName)

	If UserName <> "" and inStr(UserName,",") = 0 and inStr(Lcase(DEF_SupervisorUserName),"," & Lcase(UserName) & ",") > 0 Then
		'��������ͬ����ʾ��Ϊ���Է�����Ա���ֱ�й©��ʵ��Ӧ����ʾ����Ա���ܱ�����
		GBL_CHK_TempStr = "������ʾ���û���" & htmlencode(UserName) & "�����ڣ�"
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
		GBL_CHK_TempStr = "������ʾ���û���" & htmlencode(UserName) & "�����ڣ�"
		Exit Function
	Else
		GBL_UserName_UserLimit = cCur(Rs(1))
		GBL_UserName_UserID = cCur(Rs(0))
		GBL_UserName = Rs(2)
	End if
	Rs.Close
	Set Rs = Nothing
	',0-��֤��Ա,1-����,2-�ܰ���,3-�����û�,4-�����û�,5-���޸��û�,6-����ʽ�û�
	Dim TmpStr
	Select Case GBL_Assort
		'Case 0: 
		'		If GetBinarybit(GBL_UserName_UserLimit,2) = 1 Then
		'			GBL_CHK_TempStr = "������ʾ���û�" & htmlencode(UserName) & "�Ѿ���" & DEF_PointsName(5) & "�������ظ���ӣ�"
		'			CheckUserNameExist = 0
		'			Exit Function
		'		Else
		'			GBL_UserName_UserLimit = SetBinaryBit(GBL_UserName_UserLimit,2,1)
		'		End If
		Case 3:
				If GetBinarybit(GBL_UserName_UserLimit,7) = 1 Then
					GBL_CHK_TempStr = "������ʾ���û�" & htmlencode(UserName) & "�ķ������ݼ�ǩ���Ѿ������Σ������ظ���ӣ�"
					CheckUserNameExist = 0
					Exit Function
				Else
					GBL_UserName_UserLimit = SetBinaryBit(GBL_UserName_UserLimit,7,1)
					TmpStr = "�������з��������Ѿ�������."
				End If
		Case 4:
				If GetBinarybit(GBL_UserName_UserLimit,3) = 1 Then
					GBL_CHK_TempStr = "������ʾ���û�" & htmlencode(UserName) & "�Ѿ������Լ����Ͷ���Ϣ�������ظ���ӣ�"
					CheckUserNameExist = 0
					Exit Function
				Else
					GBL_UserName_UserLimit = SetBinaryBit(GBL_UserName_UserLimit,3,1)
					TmpStr = "���Ѿ������Է���."
				End If
		Case 5:
				If GetBinarybit(GBL_UserName_UserLimit,4) = 1 Then
					GBL_CHK_TempStr = "������ʾ���û�" & htmlencode(UserName) & "�Ѿ�����ֹ�޸����Ӽ��������ϣ������ظ���ӣ�"
					CheckUserNameExist = 0
					Exit Function
				Else
					GBL_UserName_UserLimit = SetBinaryBit(GBL_UserName_UserLimit,4,1)
					TmpStr = "���Ѿ��������޸�."
				End If
		Case 6:
				If GetBinarybit(GBL_UserName_UserLimit,1) = 1 Then
					GBL_CHK_TempStr = "������ʾ���û�" & htmlencode(UserName) & "�Ѿ�����δ����״̬�������ظ���ӣ�"
					CheckUserNameExist = 0
					Exit Function
				Else
					GBL_UserName_UserLimit = SetBinaryBit(GBL_UserName_UserLimit,1,1)
					TmpStr = "��Ŀǰ����δ����."
				End If
		Case Else:
				GBL_CHK_TempStr = "������ʾ���û�" & htmlencode(UserName) & "�Ѿ�����δ����״̬�������ظ���ӣ�"
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

End Function

rem fob ip
Function DisplayNewForbidIP

	If DEF_EnableForbidIP = 10 Then
		Response.Write "<div class=""title redfont"">ϵͳ�Ѿ���ֹ����IP���ܣ���Ҫ����IP��ַ����ϵ����Ա������</div>"
		Exit Function
	End If
	GBL_UserName = Trim(Left(Request.Form("GBL_UserName"),14))
	GBL_AnnounceID = Left(Request.Form("GBL_AnnounceID"),14)
	GBL_MessageID = Left(Request.Form("GBL_MessageID"),14)
	
	If GBL_MessageID <> "" Then
	ElseIf GBL_AnnounceID <> "" Then
	ElseIf GBL_UserName <> "" Then
		'CheckUserIPInfo(GBL_UserName)
	Else
		'GBL_IPStart = Request.Form("GBL_IPStart")
		'GBL_IPEnd = Request.Form("GBL_IPEnd")
	End If
	GBL_ExpiresTime = Left(Request.Form("GBL_ExpiresTime"),14)
	GBL_WhyString = Left(Request.Form("GBL_WhyString"),100)
	
	If isNumeric(GBL_ExpiresTime) = 0 Then GBL_ExpiresTime = -1
	GBL_ExpiresTime = fix(cCur(GBL_ExpiresTime))
	If GBL_ExpiresTime < 0 or GBL_ExpiresTime > 30 Then GBL_ExpiresTime = -1

	If Request.Form("submitflag") <> "" Then
		CheckNewIP
		If GBL_CHK_TempStr = "" Then
			SaveNewIP
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
	GBL_IPEnd = Right("000000000000" & cStr(GBL_IPEnd),12)
	GBL_IPStart = Right("000000000000" & cStr(GBL_IPStart),12)
	Number = (Left(GBL_IPEnd,3) * 256 * 256 * 256 + Mid(GBL_IPEnd,4,3) * 256 * 256 + Mid(GBL_IPEnd,7,3) * 256 + Mid(GBL_IPEnd,10,3))-(Left(GBL_IPStart,3) * 256 * 256 * 256 + Mid(GBL_IPStart,4,3) * 256 * 256 + Mid(GBL_IPStart,7,3) * 256 + Mid(GBL_IPStart,10,3)) + 1
	SQL = sql_select("Select ID from LeadBBS_ForbidIP where IPStart<=" & GBL_IPStart & " and IPEnd>=" & GBL_IPEnd,1)
	Set Rs = LDExeCute(SQL,0)
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		SQL = "Insert Into LeadBBS_ForbidIP(IPStart,IPEnd,IPNumber,ExpiresTime,WhyString) Values(" & GBL_IPStart & "," & GBL_IPEnd & "," & Number & "," & GBL_ExpiresTime & ",'" & Replace(GBL_WhyString,"'","''") & "')"
		CALL LDExeCute(SQL,0)
		GBL_CHK_TempStr = "<font color=008800 class=greenfont>�ɹ����δ�IP��,����" & Number & "��!<br>" & VbCrLf
		'GBL_CHK_TempStr = GBL_CHK_TempStr & "��ʼIP��ַ��" & GBL_IPStart & "<br>" & VbCrLf
		'GBL_CHK_TempStr = GBL_CHK_TempStr & "��ֹIP��ַ��" & GBL_IPEnd & "</font><br>" & VbCrLf
	Else
		Rs.Close
		Set Rs = Nothing
		GBL_CHK_TempStr = "<font color=ff0000 class=redfont>������ʾ����IP��ַ���Ѿ��������б���,�����ظ����!</font><br>" & VbCrLf
	End If
	If CheckSupervisorUserName = 0 Then
		CALL LDExeCute("Update LeadBBS_User Set LastWriteTime=" & GetTimeValue(DEF_Now) & " where ID=" & GBL_UserID,1)
	End If

End Function

Function CheckNewIP

	If CheckWriteEventSpace = 0 Then
		Response.Write "<b><span Class=redfont>���Ĳ�����Ƶ�����Ժ������ύ!</span></b>" & VbCrLf
		Exit Function
	End If
	If GBL_MessageID <> "" or Request.Form("submitflag") = "LKOkxk4" Then
		If CheckMessageID(GBL_MessageID) = 0 Then
			Exit Function
		End If
	ElseIf GBL_AnnounceID <> "" or Request.Form("submitflag") = "LKOkxk3" Then
		If CheckAnnounceID(GBL_AnnounceID) = 0 Then
			Exit Function
		End If
	ElseIf GBL_UserName <> "" or Request.Form("submitflag") = "LKOkxk2" Then
		If CheckUserIPInfo(GBL_UserName) = 0 Then
			Exit Function
		End If
	End If
	Dim Tmp_IPStart,Tmp_IPEnd
	Tmp_IPStart = FormatIPaddress(GBL_IPStart)
	Tmp_IPEnd = FormatIPaddress(GBL_IPEnd)

	If isNumeric(GBL_ExpiresTime) = 0 Then GBL_ExpiresTime = -1
	GBL_ExpiresTime = fix(cCur(GBL_ExpiresTime))
	If GBL_ExpiresTime < 0 or GBL_ExpiresTime > 30 Then GBL_ExpiresTime = -1
	If GBL_ExpiresTime = -1 Then
		GBL_CHK_TempStr = "������ʾ����������ѡ���������ȷѡ�񣬿����Ǵ��û�IP��ַ�����Ϲ滮��"
		Exit function
	End If

	If Len(Tmp_IPStart) <> 15 Then
		GBL_CHK_TempStr = "������ʾ����ʼ�ɣе�ַ���󣬿����Ǵ��û�IP��ַ�����Ϲ滮"
		Exit function
	End If

	If Len(Tmp_IPStart) <> 15 Then
		GBL_CHK_TempStr = "������ʾ����ֹ�ɣе�ַ���󣬿����Ǵ��û�IP��ַ�����Ϲ滮"
		Exit function
	End If
	
	Dim NewGBL_IPStart,NewGBL_IPEnd
	NewGBL_IPStart = Left(Replace(Tmp_IPStart,".",""),14)
	NewGBL_IPEnd = Left(Replace(Tmp_IPEnd,".",""),14)
	If isNumeric(NewGBL_IPStart) = 0 Then
		GBL_CHK_TempStr = "������ʾ����ʼ�ɣе�ַ���󣬱��������֣������Ǵ��û�IP��ַ�����Ϲ滮"
		Exit function
	End If
	
	If isNumeric(NewGBL_IPEnd) = 0 Then
		GBL_CHK_TempStr = "������ʾ����ֹ�ɣе�ַ���󣬱��������֣������Ǵ��û�IP��ַ�����Ϲ滮"
		Exit function
	End If
	
	NewGBL_IPStart = cCur(NewGBL_IPStart)
	NewGBL_IPEnd = cCur(NewGBL_IPEnd)
	If NewGBL_IPStart > NewGBL_IPEnd Then
		GBL_CHK_TempStr = "������ʾ����ֹ�ɣе�ַ���ܱ���ʼ�ɣе�ַС�������Ǵ��û�IP��ַ�����Ϲ滮"
		Exit function
	End If
	
	If NewGBL_IPStart > 255255255255 Then
		GBL_CHK_TempStr = "������ʾ����ʼ�ɣе�ַ�������IP��ַΪ255.255.255.255�������Ǵ��û�IP��ַ�����Ϲ滮"
		Exit function
	End If
	If NewGBL_IPEnd > 255255255255 Then
		GBL_CHK_TempStr = "������ʾ����ֹ�ɣе�ַ�������IP��ַΪ255.255.255.255�������Ǵ��û�IP��ַ�����Ϲ滮"
		Exit function
	End If

	GBL_IPStart = NewGBL_IPStart
	GBL_IPEnd = NewGBL_IPEnd
	If GBL_ExpiresTime > 0 Then
		GBL_ExpiresTime = GetTimeValue(DateAdd("d",GBL_ExpiresTime,DEF_Now))
	Else
		GBL_ExpiresTime = 0
	End If

End Function

Function DisplayNewIPForm

	Dim N
	If GBL_CHK_TempStr <> "" Then Response.Write "<p><font color=ff0000 class=redfont><b>" & GBL_CHK_TempStr & "</b></font></p>"%>

			<%If Request.Form("submitflag") = "LKOkxk2" or Request.Form("submitflag") = "" Then%>
	<div class=title>
	���������û��������Σ�������Ҫ���Σɣе�ַ�������û�����
	</div>
	<form action=LimitUserManage.asp method=post id=fobform name=fobform>
		<div class="value2">
			���ߵ��û�����<input name=GBL_UserName value="<%=htmlencode(GBL_UserName)%>" class=fminpt><br>
			<input name=submitflag type=hidden value="LKOkxk2">
			<input name=action type=hidden value="fobip">
		</div>
		<div class="value2">
			����ʱ��ѡ��<select name=GBL_ExpiresTime>
						<%For N = 1 to 30
							If N = GBL_ExpiresTime Then
								Response.Write "<option value=" & N & " selected>����" & Right("0" & N,2) & "��</option>"
							Else
								Response.Write "<option value=" & N & ">����" & Right("0" & N,2) & "��</option>"
							End If
						Next%>
						<option value=0<%If GBL_ExpiresTime = 0 Then Response.Write " Selected"%>>��������</option>
					</select>
		</div>
		<div class="value2">
			����ԭ��ע����<input name=GBL_WhyString MaxLength=100 size=30 value="<%=htmlencode(GBL_WhyString)%>" class=fminpt>
			<select onchange="document.fobform.GBL_WhyString.value=this.value;">
				<option value="">===һЩ����ԭ����ѡ��===</option>
				<option value="��������ɫ������">��������ɫ������</option>
				<option value="����̳���ж��⹥��(�ڿ���Ϊ)">����̳���ж��⹥��(�ڿ���Ϊ)</option>
				<option value="��ͣ�Ķ����ˮ��ע�����û�">��ͣ�Ķ����ˮ��ע�����û�</option>
			</select>
		</div>
		<div class="value2">
			<input type=submit value="�ύ" class="fmbtn btn_2"> <input type=reset value="ȡ��" class="fmbtn btn_2">
		</div>
		</form>
		
		<div class="title">
		��ʾ��
		</div>
		<div class="value2"><span class=grayfont>�˲���ֻ�Ե�ǰ���ߵ��û��Ż���Ч</span>
		</div>
		<%End If%>

		<%If Request.Form("submitflag") = "LKOkxk3" or Request.Form("submitflag") = "" Then%>
		<br>
		<hr class=splitline>
		<div class="title">
		���ݷ������������Σ�����ĳ�û����������ӵı��
		</div>
          	<form action=LimitUserManage.asp method=post id=fobform name=fobform>
          	<div class="value2">
			��̳���ӱ�ţ�<input name=GBL_AnnounceID value="<%=htmlencode(GBL_AnnounceID)%>" class=fminpt>
		</div>
			<input name=submitflag type=hidden value="LKOkxk3">
			<input name=action type=hidden value="fobip">
		<div class="value2">
			����ʱ��ѡ��<select name=GBL_ExpiresTime>
						<%For N = 1 to 30
							If N = GBL_ExpiresTime Then
								Response.Write "<option value=" & N & " selected>����" & Right("0" & N,2) & "��</option>"
							Else
								Response.Write "<option value=" & N & ">����" & Right("0" & N,2) & "��</option>"
							End If
						Next%>
						<option value=0<%If GBL_ExpiresTime = 0 Then Response.Write " Selected"%>>��������</option>
					</select>
		</div>
		<div class="value2">
			����ԭ��ע����<input name=GBL_WhyString MaxLength=100 size=30 value="<%=htmlencode(GBL_WhyString)%>" class=fminpt>
			<select onchange="document.fobform.GBL_WhyString.value=this.value;">
				<option value="">===һЩ����ԭ����ѡ��===</option>
				<option value="��������ɫ������">��������ɫ������</option>
				<option value="����̳���ж��⹥��(�ڿ���Ϊ)">����̳���ж��⹥��(�ڿ���Ϊ)</option>
				<option value="��ͣ�Ķ����ˮ��ע�����û�">��ͣ�Ķ����ˮ��ע�����û�</option>
			</select>
		</div>
		<div class="value2">
			<input type=submit value="�ύ" class="fmbtn btn_2"> <input type=reset value="ȡ��" class="fmbtn btn_2">
		</div></form>
		
		<div class="title">
		��ʾ��
		</div>
		<div class="value2"><span class=grayfont>���ӵı�ţ��ڰ����б��У�����������ǰ���ͼ���Ͽ�����ʾ����������ڲ鿴��������ʱ������������������ϣ�������ʾ��������ظ����ı��</span>
		</div><%End If%>
			

		<%If Request.Form("submitflag") = "LKOkxk4" or Request.Form("submitflag") = "" Then%>
		<br>
		<hr class=splitline>
		<div class="title">���ݶ���Ϣ��������Σ�����ĳ�û������Ͷ���Ϣ�ı��
		</div>
			<form action=LimitUserManage.asp method=post id=fobform name=fobform>
		<div class="value2">
			����Ϣ�ı�ţ�<input name=GBL_MessageID value="<%=htmlencode(GBL_MessageID)%>" class=fminpt>
		</div>
			<input name=submitflag type=hidden value="LKOkxk4">
			<input name=action type=hidden value="fobip">
		<div class="value2">
			����ʱ��ѡ��<select name=GBL_ExpiresTime>
						<%For N = 1 to 30
							If N = GBL_ExpiresTime Then
								Response.Write "<option value=" & N & " selected>����" & Right("0" & N,2) & "��</option>"
							Else
								Response.Write "<option value=" & N & ">����" & Right("0" & N,2) & "��</option>"
							End If
						Next%>
						<option value=0<%If GBL_ExpiresTime = 0 Then Response.Write " Selected"%>>��������</option>
					</select>
		</div>
		<div class="value2">
			����ԭ��˵����<input name=GBL_WhyString MaxLength=100 size=30 value="<%=htmlencode(GBL_WhyString)%>" class=fminpt>
			<select onchange="document.fobform.GBL_WhyString.value=this.value;">
				<option value="">===һЩ����ԭ����ѡ��===</option>
				<option value="��������ɫ������">��������ɫ������</option>
				<option value="����̳���ж��⹥��(�ڿ���Ϊ)">����̳���ж��⹥��(�ڿ���Ϊ)</option>
				<option value="��ͣ�Ķ����ˮ��ע�����û�">��ͣ�Ķ����ˮ��ע�����û�</option>
			</select>
		</div>
		<div class="value2">
			<input type=submit value="�ύ" class="fmbtn btn_2"> <input type=reset value="ȡ��" class="fmbtn btn_2">
		</div></form>
		<div class="title">
		��ʾ��
		</div>
		<div class="value2"><span class=grayfont>����Ϣ��ſ����ڲ鿴�ռ����б�����ʾ</span>
		</div>
		<%End If%>

<%End Function


Function FormatIPaddress(KIP)

	Dim IP
	IP = KIP
	Rem ��ȥ���׵Ŀյ㣬����ʽ����XXX.XXX.XXX.XXX
	Dim Temp1,Temp2,TempN,Temp
	IP = Trim(IP & "")
	If inStr(IP,".") = 0 or Len(IP) = "" Then
		FormatIPaddress = IP
		Exit Function
	End if
	
	Temp1 = Split(IP,".")
	IP = ""
	Temp2 = Ubound(Temp1,1)
	
	TempN = 0
	do while IP = ""
		If Temp1(TempN) <> "" Then
			if IsNumeric(Temp1(TempN)) Then Temp1(TempN) = cStr(cCur(Temp1(TempN)))
			If Len(Temp1(TempN)) < 3 Then
				IP = string(3-len(Temp1(TempN)),"0") & Temp1(TempN)
			else
				IP = Temp1(TempN)
			End If
			TempN = TempN + 1
			Exit Do
		Else
			TempN = TempN + 1
		End If
		If TempN > Temp2 Then Exit do
	Loop
	
	For Temp = TempN to Temp2
		If Temp1(TempN) <> "" Then
			If isNumeric(Temp1(TempN)) = 0 Then
				FormatIPaddress = ""
				Exit Function
			End If
			Temp1(TempN) = Fix(cCur(Temp1(TempN)))
			If Temp1(TempN) < 0 or Temp1(TempN) > 255 Then
				FormatIPaddress = ""
				Exit Function
			End If
			if IsNumeric(Temp1(TempN)) Then Temp1(TempN) = cStr(cCur(Temp1(TempN)))
			If Len(Temp1(TempN)) < 3 Then
				IP = IP & "." & string(3-len(Temp1(TempN)),"0") & Temp1(TempN)
			else
				IP = IP & "." & Temp1(TempN)
			End If
		End If
		TempN = TempN + 1
	Next
	FormatIPaddress = IP
	Rem ���ص�IP��ַ�պ���15λ���������15���ַ����Ǵ�����Ч��IP��ַ

End Function


Function CheckUserIPInfo(UserName)

	If UserName <> "" and inStr(UserName,",") = 0 and inStr(Lcase(DEF_SupervisorUserName),"," & Lcase(UserName) & ",") > 0 Then
		'��������ͬ����ʾ��Ϊ���Է�����Ա���ֱ�й©��ʵ��Ӧ����ʾ����Ա���ܱ�����
		GBL_CHK_TempStr = "������ʾ���û���" & htmlencode(UserName) & "�����ڣ�"
		CheckUserIPInfo = 0
		Exit Function
	End If

	Dim Rs
	Set Rs = LDExeCute(sql_select("Select ID,UserName from LeadBBS_User where UserName='" & Replace(UserName,"'","''") & "'",1),0)
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		CheckUserIPInfo = 0
		GBL_CHK_TempStr = "������ʾ���û���" & htmlencode(UserName) & "�����ڣ�"
		Exit Function
	Else
		GBL_UserName_UserID = cCur(Rs(0))
		GBL_UserName = Rs(1)
	End if
	Rs.Close
	Set Rs = Nothing
	
	Set Rs = LDExeCute(sql_select("Select IP from LeadBBS_OnlineUser where UserID=" & GBL_UserName_UserID,1),0)
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		CheckUserIPInfo = 0
		GBL_CHK_TempStr = "������ʾ���û���" & htmlencode(UserName) & "Ŀǰ�����ߣ��޷�������Σ���ʹ�������ķ�ʽ�����Σ�"
		Exit Function
	Else
		GBL_IPStart = Rs(0)
		GBL_IPEnd = GBL_IPStart
		Rs.Close
		Set Rs = Nothing
	End if
		
	CheckUserIPInfo = 1

End Function

Rem ���ĳ����
Function CheckAnnounceID(AnnounceID)

	If isNumeric(AnnounceID) = False Then
		GBL_CHK_TempStr = "������ʾ�����Ӳ������ڻ���Ȩ���Σ�"
		CheckAnnounceID = 0
		Exit Function
	End If
	AnnounceID = Fix(cCur(AnnounceID))
	If AnnounceID < 1 Then
		GBL_CHK_TempStr = "������ʾ�����" & htmlencode(AnnounceID) & "�����Ӳ������ڻ���Ȩ���Σ�"
		CheckAnnounceID = 0
		Exit Function
	End If
	
	Dim Rs
	Set Rs = LDExeCute(sql_select("Select IPAddress,UserName from LeadBBS_Announce where ID=" & AnnounceID,1),0)
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		CheckAnnounceID = 0
		GBL_CHK_TempStr = "������ʾ�����" & htmlencode(AnnounceID) & "�����Ӳ������ڻ���Ȩ���Σ�"
		Exit Function
	Else
		GBL_IPStart = Rs(0)
		GBL_IPEnd = GBL_IPStart
		GBL_UserName = Rs(1)
	End if
	Rs.Close
	Set Rs = Nothing

	If GBL_UserName <> "" and inStr(GBL_UserName,",") = 0 and inStr(Lcase(DEF_SupervisorUserName),"," & Lcase(GBL_UserName) & ",") > 0 Then
		'��������ͬ����ʾ��Ϊ���Է�����Ա���ֱ�й©��ʵ��Ӧ����ʾ����Ա���ܱ�����
		GBL_CHK_TempStr = "������ʾ�����" & htmlencode(AnnounceID) & "�����Ӳ������ڻ���Ȩ���Σ�"
		CheckAnnounceID = 0
		Exit Function
	End If
	CheckAnnounceID = 1

End Function


Rem ���ĳ����
Function CheckMessageID(MessageID)

	If isNumeric(MessageID) = False Then
		GBL_CHK_TempStr = "������ʾ������Ϣ�������ڻ���Ȩ���Σ�"
		CheckMessageID = 0
		Exit Function
	End If
	MessageID = Fix(cCur(MessageID))
	If MessageID < 1 Then
		GBL_CHK_TempStr = "������ʾ�����" & htmlencode(MessageID) & "�Ķ���Ϣ�������ڻ���Ȩ���Σ�"
		CheckMessageID = 0
		Exit Function
	End If
	
	Dim Rs
	Set Rs = LDExeCute(sql_select("Select IP,FromUser from LeadBBS_InfoBox where ID=" & MessageID,1),0)
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		CheckMessageID = 0
		GBL_CHK_TempStr = "������ʾ�����" & htmlencode(MessageID) & "�Ķ���Ϣ�������ڻ���Ȩ���Σ�"
		Exit Function
	Else
		GBL_IPStart = Rs(0)
		GBL_IPEnd = GBL_IPStart
		GBL_UserName = Rs(1)
	End if
	Rs.Close
	Set Rs = Nothing

	If GBL_UserName <> "" and inStr(GBL_UserName,",") = 0 and inStr(Lcase(DEF_SupervisorUserName),"," & Lcase(GBL_UserName) & ",") > 0 Then
		'��������ͬ����ʾ��Ϊ���Է�����Ա���ֱ�й©��ʵ��Ӧ����ʾ����Ա���ܱ�����
		GBL_CHK_TempStr = "������ʾ�����" & htmlencode(MessageID) & "�Ķ���Ϣ�������ڻ���Ȩ���Σ�"
		CheckMessageID = 0
		Exit Function
	End If
	CheckMessageID = 1

End Function

rem modifyuser

Function DisplayModifyUser

	If Request.Form("submitflag") <> "" Then
		CheckModifyUserForm
		If GBL_CHK_TempStr = "" Then
			ModifyUser
			Response.Write GBL_CHK_TempStr
			GBL_CHK_TempStr = ""
			DisplayModifyUserForm
		Else
			DisplayModifyUserForm
		End If
	Else
		DisplayModifyUserForm
	End If

End Function

Function ModifyUser

	Response.Write "<p><b>��ʼ����û�<u>" & htmlencode(GBL_UserName) & "</u>���������ϣ�</b></p>" & VbCrLf
	If inStr(GBL_ModifyMode,",1,") Then
		If GBL_UserName_FaceUrl & "" = "" Then
			Response.Write "<p><font color=Red class=redfont>�������ͷ�� ���û�ͷ���Ѿ���Ĭ��ͷ���Թ�������</font></p>"
			DeleteUploadFace(GBL_UserName_UserID)
		Else
			CALL LDExeCute("Update LeadBBS_User Set FaceUrl='' where ID=" & GBL_UserName_UserID,1)
			Response.Write "<p><font color=Green class=greenfont>�������ͷ�� �ɹ������</font></p>"
			DeleteUploadFace(GBL_UserName_UserID)
		End If
	End If

	If inStr(GBL_ModifyMode,",2,") Then
		If GBL_UserName_UnderWrite & "" = "" Then
			Response.Write "<p><font color=Red class=redfont>����û�ǩ���� ���û���ǩ�����ݣ��Թ�������</font></p>"
		Else
			CALL LDExeCute("Update LeadBBS_User Set UnderWrite='',PrintUnderWrite='' where ID=" & GBL_UserName_UserID,1)
			Response.Write "<p><font color=Green class=greenfont>����û�ǩ���� �ɹ������</font></p>"
		End If
	End If

	If inStr(GBL_ModifyMode,",3,") Then
		If GBL_UserName_UserTitle & "" = "" Then
			Response.Write "<p><font color=Red class=redfont>����û�ͷ�Σ� ���û���ͷ�Σ��Թ�������</font></p>"
		Else
			CALL LDExeCute("Update LeadBBS_User Set UserTitle='' where ID=" & GBL_UserName_UserID,1)
			Response.Write "<p><font color=Green class=greenfont>����û�ͷ�Σ� �ɹ������</font></p>"
		End If
	End If

	If CheckSupervisorUserName = 0 Then
		CALL LDExeCute("Update LeadBBS_User Set LastWriteTime=" & GetTimeValue(DEF_Now) & " where ID=" & GBL_UserID,1)
	End If

End Function

Function CheckModifyUserForm

	If CheckWriteEventSpace = 0 Then
		Response.Write "<b><font color=Red Class=redfont>���Ĳ�����Ƶ�����Ժ������ύ!</font></b> <br>" & VbCrLf
		Exit Function
	End If
	
	GBL_ModifyMode = Replace("," & Left(Request.Form("GBL_ModifyMode"),10) & ","," ","")
	GBL_UserName = Left(Request.Form("GBL_UserName"),20)
	If isNumeric(Replace(GBL_ModifyMode,",","")) = 0 Then
		GBL_CHK_TempStr = "������ʾ������ѡ��ѡ�����"
		Exit Function
	End If

	If GBL_UserName = "" Then
		GBL_CHK_TempStr = "������ʾ���������û�����"
		Exit Function
	End If
	
	If CheckModifyUserNameExist(GBL_UserName) = 0 Then
		GBL_CHK_TempStr = "������ʾ���û��������ڣ�"
		Exit Function
	End If

End Function

Function DisplayModifyUserForm

	If GBL_CHK_TempStr <> "" Then Response.Write "<div class=""title redfont"">" & GBL_CHK_TempStr & "</div>"
	If Request.Form("submitflag") = "LKOkxk2" or Request.Form("submitflag") = "" Then%>
	<div class="title">�����û�����</div>
	<form action=LimitUserManage.asp method=post id=fobform name=fobform>
	<div class="value2">
		�� �� ���� <input name=GBL_UserName value="<%=htmlencode(GBL_UserName)%>" class=fminpt>
	</div>
		<input name=submitflag type=hidden value="LKOkxk2">
		<input name=action type=hidden value="modifyuser">
	<div class="value2">
		ѡ�������<input name=GBL_ModifyMode value=1<%If inStr(GBL_ModifyMode,",1,") Then Response.Write " checked"%> type=checkbox>�������ͷ��
		<input name=GBL_ModifyMode value=2<%If inStr(GBL_ModifyMode,",2,") Then Response.Write " checked"%> type=checkbox>����û�ǩ��
		<input name=GBL_ModifyMode value=3<%If inStr(GBL_ModifyMode,",3,") Then Response.Write " checked"%> type=checkbox>����û�ͷ��
	</div>
	<div class="value2">
		<input type=submit value="�ύ" class="fmbtn btn_2"> <input type=reset value="ȡ��" class="fmbtn btn_2">
	</div>
	</form>
	<br>
	<div class="title">��ʾ��</div>
	<ol>
	<li>����û�����ͷ��󣬴��û�ͷ��ָ�Ϊ��̳���е�Ĭ��ͷ��</li>
	<li>����û�ǩ������ʹָ�����û�ǩ������ȫ���Ƴ�</li>
	<li>����û�ͷ�ν���ʹָ�����û�ͷ����ȡ��</li>
	<li>ĳЩ�ض��û����ϲ������޸�</li>
	</ol>
	<%End If%>

<%End Function

Rem ���ĳ�û����Ƿ����
Function CheckModifyUserNameExist(UserName)

	If UserName <> "" and inStr(UserName,",") = 0 and inStr(Lcase(DEF_SupervisorUserName),"," & Lcase(UserName) & ",") > 0 Then
		'��������ͬ����ʾ��Ϊ���Է�����Ա���ֱ�й©��ʵ��Ӧ����ʾ����Ա���ܱ�����
		GBL_CHK_TempStr = "������ʾ���û���" & htmlencode(UserName) & "�����ڣ�"
		CheckModifyUserNameExist = 0
		Exit Function
	End If

	Dim Rs
	Set Rs = LDExeCute(sql_select("Select ID,UserName,FaceUrl,UnderWrite,UserTitle from LeadBBS_User where UserName='" & Replace(UserName,"'","''") & "'",1),0)
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		CheckModifyUserNameExist = 0
		GBL_CHK_TempStr = "������ʾ���û���" & htmlencode(UserName) & "�����ڣ�"
		Exit Function
	Else
		GBL_UserName_UserID = cCur(Rs(0))
		GBL_UserName = Rs(1)
		GBL_UserName_FaceUrl = Rs(2)
		GBL_UserName_UnderWrite = Rs(3)
		GBL_UserName_UserTitle = Rs(4)
	End if
	Rs.Close
	Set Rs = Nothing
		
	CheckModifyUserNameExist = 1

End Function


Function DeleteUploadFace(DelUserID)

	If DEF_FSOString = "" Then
		Response.Write "<p><span class=redfont>��̳��֧������ɾ���ļ����Թ��ϴ�ͷ��ɾ����</span>"
		Exit Function
	End If
	Dim SQL,Rs
	SQL = "Select ID,PhotoDir,SPhotoDir from LeadBBS_UserFace where UserID=" & DelUserID & " order by ID ASC"
	Set Rs = LDExeCute(SQL,0)
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		Response.Write "<p><b><span class=redfont>�û����ϴ�ͷ���Թ�ɾ��!</span></b>"
	Else
		If Rs("PhotoDir") <> "" Then DeleteFiles(Server.Mappath(DEF_BBS_HomeUrl & DEF_BBS_UploadPhotoUrl & "face/" & Rs("PhotoDir")))
		If Rs("SPhotoDir") <> "" Then DeleteFiles(Server.Mappath(DEF_BBS_HomeUrl & DEF_BBS_UploadPhotoUrl & "face/" & Rs("SPhotoDir")))
		Rs.Close
		Set Rs = Nothing
		CALL LDExeCute("Delete from LeadBBS_UserFace where UserID=" & DelUserID,1)
		Response.Write "<p><b><span class=greenfont>����û��ϴ�ͷ���ɾ��!</span></b>"
	End If

End Function

Function DeleteFiles(path)

	'on error resume next
	Dim fs
	Set fs=Server.CreateObject(DEF_FSOString)
	If fs.FileExists(path) then
		fs.DeleteFile path,True
		DeleteFiles = 1
	Else
		DeleteFiles = 0
	End If
    Set fs=nothing
         
End Function

rem clear
Sub View_ClearExpiresInfo

	If Request.Form("DeleteSure")="E72ksiOkw2" Then
		If DeleteForbidIPandUser = 1 Then
			Response.Write "<p><font color=008800 class=greenfont><b>�Ѿ��ɹ�������е��ڵ������û������εģɣе�ַ��</b></font></p>"
		else
			Response.Write "<p><font color=ff0000 class=redfont><b>" & GBL_CHK_TempStr & "</b></font></p>"
		End If
	Else
		%>
		<form action=LimitUserManage.asp method=post>
		<div class="title">�����������û�������IP��ַ</div>
		<div class="value2">
		<span class=redfont>ȷ����Ϣ��������<%=year(DEF_Now)%>��<%=month(DEF_Now)%>��<%=day(DEF_Now)%>���˶������������֮ǰ�ѵ��ڵ���Ϣ���������£�<span>
		</div>
		<ol>
		<li>��������ε�IP��ַ</li>
		<li>��������η������ݵ��û�</li>
		<li>��������Ե��û�</li>
		<li>�������ֹ�޸ĵ��û�</li>
		<li>�ָ������˵�<%=DEF_PointsName(5)%>����ͨ�û�״̬</li>
		<li>����ڼ�����Ч���ѹ�����δ�����ע���û�</li>
		</ol>
		<input type=hidden name=DeleteSure value="E72ksiOkw2">
		<input type=hidden name=action value="clear">
		<div class="value2">
		<input type=submit value=��ʼ���� class="fmbtn btn_3">
		</div>
		</form>
	<%End If

End Sub


Function DeleteForbidIPandUser

	Server.ScriptTimeOut = 6000
	'If UserName <> "" and inStr(UserName,",") = 0 and inStr(Lcase(DEF_SupervisorUserName),"," & Lcase(UserName) & ",") > 0 Then
	'	GBL_CHK_TempStr = "������ʾ���û���" & htmlencode(UserName) & "�����ڣ�"
	'	DeleteForbidIPandUser = 0
	'	Exit Function
	'End If
	
	Response.Write "<div class=title>������ɣ�</div>"
	Dim ExpiresTime
	ExpiresTime = GetTimeValue(year(DEF_Now) & "-" & Month(DEF_Now) & "-" & Day(DEF_Now))
	Dim Rs
	Set Rs = LDExeCute("Select T2.ID,T2.UserLimit,T2.UserName,T1.Assort from LeadBBS_SpecialUser as T1 Left join LeadBBS_User As T2 on T1.UserID=T2.ID where T1.ExpiresTime>0 and T1.ExpiresTime<" & ExpiresTime,0)
	If Rs.Eof Then
		DeleteForbidIPandUser = 1
		Response.Write "<div class=value2>���κε��ڵ������û�������Ҫ���£�</div>"
	End If
	Dim GBL_UserName_UserID,GBL_UserName_UserLimit,GBL_UserName,GBL_Assort
	Do while Not Rs.Eof
		GBL_UserName_UserLimit = cCur(Rs(1))
		GBL_UserName_UserID = cCur(Rs(0))
		GBL_UserName = Rs(2)
		GBL_Assort = cCur(Rs(3))
		
		',0-��֤��Ա,1-����,2-�ܰ���,3-�����û�,4-�����û�,5-���޸��û�,6-����ʽ�û�
		Select Case GBL_Assort
			Case 0:
					If GetBinarybit(GBL_UserName_UserLimit,2) = 1 Then
						Response.Write "<div class=value2>�û�" & htmlencode(GBL_UserName) & "�Ѿ����" & DEF_PointsName(5) & "״̬��</div>"
						GBL_UserName_UserLimit = SetBinaryBit(GBL_UserName_UserLimit,2,0)
						CALL LDExeCute("Update LeadBBS_User Set UserLimit=" & GBL_UserName_UserLimit & " where ID=" & GBL_UserName_UserID,1)
						CALL LDExeCute("Delete from LeadBBS_SpecialUser Where Assort=" & GBL_Assort & " and UserID=" & GBL_UserName_UserID,1)
					End If
			Case 3:
					If GetBinarybit(GBL_UserName_UserLimit,7) = 1 Then
						Response.Write "<div class=value2>�û�" & htmlencode(GBL_UserName) & "�Ѿ�������η������ݼ�ǩ����</div>"
						GBL_UserName_UserLimit = SetBinaryBit(GBL_UserName_UserLimit,7,0)
						CALL LDExeCute("Update LeadBBS_User Set UserLimit=" & GBL_UserName_UserLimit & " where ID=" & GBL_UserName_UserID,1)
						CALL LDExeCute("Delete from LeadBBS_SpecialUser Where Assort=" & GBL_Assort & " and UserID=" & GBL_UserName_UserID,1)
					End If
			Case 4:
					If GetBinarybit(GBL_UserName_UserLimit,3) = 1 Then
						Response.Write "<div class=value2>�û�" & htmlencode(GBL_UserName) & "�Ѿ�������Լ����Ͷ���Ϣ��</div>"
						GBL_UserName_UserLimit = SetBinaryBit(GBL_UserName_UserLimit,3,0)
						CALL LDExeCute("Update LeadBBS_User Set UserLimit=" & GBL_UserName_UserLimit & " where ID=" & GBL_UserName_UserID,1)
						CALL LDExeCute("Delete from LeadBBS_SpecialUser Where Assort=" & GBL_Assort & " and UserID=" & GBL_UserName_UserID,1)
					End If
			Case 5:
					If GetBinarybit(GBL_UserName_UserLimit,4) = 1 Then
						Response.Write "<div class=value2>�û�" & htmlencode(GBL_UserName) & "�Ѿ������ֹ�޸����Ӽ��������ϣ�</div>"
						GBL_UserName_UserLimit = SetBinaryBit(GBL_UserName_UserLimit,4,0)
						CALL LDExeCute("Update LeadBBS_User Set UserLimit=" & GBL_UserName_UserLimit & " where ID=" & GBL_UserName_UserID,1)
						CALL LDExeCute("Delete from LeadBBS_SpecialUser Where Assort=" & GBL_Assort & " and UserID=" & GBL_UserName_UserID,1)
					End If
			Case 6:
					If GetBinarybit(GBL_UserName_UserLimit,1) = 1 Then
						Response.Write "<div class=value2>δ�����û�" & htmlencode(GBL_UserName) & "�Ѿ����ɹ�ɾ����</div>"
						GBL_UserName_UserLimit = SetBinaryBit(GBL_UserName_UserLimit,1,0)
						CALL LDExeCute("delete from LeadBBS_User where ID=" & GBL_UserName_UserID,1)
						CALL LDExeCute("Update LeadBBS_SiteInfo Set UserCount=UserCount-1",1)
						UpdateStatisticDataInfo -1,1,1
						CALL LDExeCute("Delete from LeadBBS_SpecialUser Where Assort=" & GBL_Assort & " and UserID=" & GBL_UserName_UserID,1)
					End If
			Case Else:
		End Select
		Rs.MoveNext
	Loop
	Rs.Close
	Set Rs = Nothing
	Response.Write "<div class=value2><span Class=greenfont>���������û�������ɣ�</span></div>"
	Set Rs = LDExeCute("Delete From LeadBBS_ForbidIP where ExpiresTime>0 and ExpiresTime<" & ExpiresTime,0)
	Response.Write "<div class=value2><span class=greenfont>�������ڵı����Σɣе�ַ�Ѿ��ɹ���ɣ�</span></div>"
	DeleteForbidIPandUser = 1

End Function
%>