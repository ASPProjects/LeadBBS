<%
Const DEF_AllowPunish = 1 '�Ƿ�������ͨ�û��ͷ������û���1.�����ձ��û��ͷ������û���������ֹ
Const DEF_AllowOpinionNum = 3 '������ͨ�û����۴��� 0,��ֹ,-1 �������� >0 ָ������
Const DEF_MasterNolimit = 0 '����������Ա���۴����Ƿ����ޣ����������ޣ���������ͬ��ͨ�û���
Const DEF_AllowBoardMasterCachetValue = 1 '�Ƿ������������������1.�� 0.��
Dim GBL_GoodFlag,ALL_FirstRootID,ALL_LastRootID,LMT_UserID
Dim Form_OpinionUser,Form_OpinionWhys,Form_OpinionNum,Old_Form_OpinionNum,Form_OpinionStr,Form_OpitionType,Form_CachetNum,Form_CharmNum,Form_PointsNum,Form_OpinionCount
Dim MakeGood_Title,MakeGood_User
Dim MakeGood_Level
MakeGood_Level = 0

Function CheckMakeGoodSure

	If GetBinarybit(GBL_CHK_UserLimit,6) = 1 Then
		Processor_ErrMsg "����Ȩ�޲��㣡" & VbCrLf
		CheckMakeGoodSure = 0
		MakeGood_Level = 0
		Exit Function
	End if
	
	Dim Rs,SQL
	SQL = sql_select("Select BoardID,UserID,GoodFlag,Opinion,Title,UserName,TitleStyle from LeadBBS_Announce where id=" & LMT_AncID,1)
	Set Rs = LDExeCute(SQL,0)
	If Rs.Eof Then
		Processor_ErrMsg "����Ҫ����������ID�����ڣ�" & VbCrLf
		Rs.Close
		Set Rs = Nothing
		CheckMakeGoodSure = 0
		MakeGood_Level = 0
		Exit Function
	End if

	GBL_Board_ID = Rs(0)
	LMT_UserID = cCur(Rs(1))
	GBL_GoodFlag = ccur(Rs(2))
	Form_OpinionStr = Rs(3)
	MakeGood_Title = KillHTMLLabel(DisplayAnnounceTitle(Rs(4),Rs(6)))
	MakeGood_User = Rs(5)
	Rs.Close
	Set Rs = Nothing

	If Len(Replace(Form_OpinionStr,"|","")) = Len(Form_OpinionStr) - 2 Then
		Form_OpinionStr = Split(Form_OpinionStr,"|")
		Form_OpinionUser = Form_OpinionStr(0)
		Form_OpinionNum = Form_OpinionStr(1)
		If isNumeric(Form_OpinionNum) = 0 Then Form_OpinionNum = 0
		Form_OpinionNum = Fix(cCur(Form_OpinionNum))
		Old_Form_OpinionNum = Form_OpinionNum
		Form_OpinionWhys = Form_OpinionStr(2)
		Form_OpitionType = 2
	ElseIf Len(Replace(Form_OpinionStr,"|","")) = Len(Form_OpinionStr) - 3 Then
		Form_OpinionStr = Split(Form_OpinionStr,"|")
		If isNumeric(Form_OpinionStr(0)) = 0 or Form_OpinionStr(0) = "" Then Form_OpinionStr(0) = 0
		If isNumeric(Form_OpinionStr(1)) = 0 or Form_OpinionStr(1) = "" Then Form_OpinionStr(1) = 0
		If isNumeric(Form_OpinionStr(2)) = 0 or Form_OpinionStr(2) = "" Then Form_OpinionStr(2) = 0
		If isNumeric(Form_OpinionStr(3)) = 0 or Form_OpinionStr(3) = "" Then Form_OpinionStr(3) = 0
		Form_PointsNum = cCur(Form_OpinionStr(0))
		Form_CachetNum = cCur(Form_OpinionStr(1))
		Form_CharmNum = cCur(Form_OpinionStr(2))
		Form_OpinionCount = cCur(Form_OpinionStr(3))
		Form_OpitionType = 3
	Else
		Form_OpinionUser = ""
		Form_OpinionWhys = ""
		Form_OpinionNum = 0
		Old_Form_OpinionNum = 0
		Form_OpitionType = 0
	End If

	Dim Temp
	Temp = Application(DEF_MasterCookies & "BoardInfo" & GBL_Board_ID)
	If isArray(Temp) = False Then
		ReloadBoardInfo(GBL_Board_ID)
		Temp = Application(DEF_MasterCookies & "BoardInfo" & GBL_Board_ID)
	End If
	If isArray(Temp) = False Then
		Processor_ErrMsg "������̳������������ϵ����Ա��" & VbCrLf
		CheckMakeGoodSure = 0
		MakeGood_Level = 0
		Set Rs = Nothing
		Exit Function
	End If
	
	GBL_Board_MasterList = Temp(10,0)
	ALL_FirstRootID = Temp(33,0)
	ALL_LastRootID = Temp(34,0)
	
	CheckisBoardMaster
	If CheckSupervisorUserName = 1 Then
		CheckMakeGoodSure = 1
		MakeGood_Level = 3
	ElseIf GBL_UserID >= 1 and (GBL_BoardMasterFlag >= 5 and GetBinarybit(GBL_CHK_UserLimit,6) = 0) Then
		CheckMakeGoodSure = 1
		MakeGood_Level = 2
	ElseIf GBL_UserID >= 1 and CheckUserAnnounceLimit = 1 Then
		GBL_CHK_TempStr = ""
		CheckMakeGoodSure = 1
		MakeGood_Level = 1
	Else
		CheckMakeGoodSure = 0
		MakeGood_Level = 0
		Processor_ErrMsg "����Ȩ�޲���(UsrLMT)��"
	End If

End Function

Function Opinion_CheckAllowOpinion

	If DEF_AllowOpinionNum = -1 Then
		Opinion_CheckAllowOpinion = 1
		Exit Function
	End If

	If DEF_AllowOpinionNum = 0 Then
		Opinion_CheckAllowOpinion = 0
		Exit Function
	End If
	
	Dim Rs
	Set Rs = LDExeCute("Select count(*) from LeadBBS_Opinion where AnnounceID=" & LMT_AncID  & " and UserName='" & Replace(GBL_CHK_User,"'","''") & "'",0)
	If Rs.Eof Then
		Opinion_CheckAllowOpinion = 1
	Else
		Dim Tmp
		Tmp = Rs(0)
		If isNumeric(Tmp) = 0 or Tmp & "" = "" Then Tmp = 0
		Tmp = cCur(Tmp)
		If Tmp >= DEF_AllowOpinionNum Then
			Opinion_CheckAllowOpinion = 0
		Else
			Opinion_CheckAllowOpinion = 1
		End If
	End If
	Rs.Close
	Set Rs = Nothing

End Function

Sub DisplayMakeGoodAnnounce

	If MakeGood_Level = 1 or (Request.Form("Form_GoodType") = "2" and DEF_MasterNolimit = 0) Then
		If Opinion_CheckAllowOpinion = 0 Then
			Processor_ErrMsg "<span class=redfont>���Դ������۴����ѳ�����������,��������ֵ����.</span>"
			Exit Sub
		End If
	End If
	If Request.Form("Form_GoodType") = "2" Then
		OpinionAnnounce
		Exit Sub
	End If
	If Request.Form("SureFlag")="1" and MakeGood_Level >= 2 Then
		If CheckWriteEventSpace = 0 Then
			Processor_ErrMsg "<span class=redfont>���Ĳ�����Ƶ�����Ժ�ˢ�����ԣ�</span>"
			Exit Sub
		End If

		Dim Rs,SQL
		Set Rs = Server.CreateObject("ADODB.RecordSet")
		SQL = sql_select("Select UserID,GoodFlag,BoardID,ParentID from LeadBBS_Announce where id=" & LMT_AncID,1)
		Set Rs = LDExeCute(SQL,0)
		If Rs.Eof Then
			Processor_ErrMsg "����δѡ��Ҫ���������ӣ�" & VbCrLf
			Rs.Close
			Set Rs = Nothing
			Exit Sub
		'ElseIf cCur(Rs("ParentID")) <> 0 Then
		'	Processor_ErrMsg "����,���������ӱ�������������" & VbCrLf
		'	Rs.Close
		'	Set Rs = Nothing
		'	Exit Sub
		End if
		
		Dim UserID,GoodFlag,BoardID
		UserID = Rs(0)
		GoodFlag = ccur(Rs(1))
		BoardID = Rs(2)
		If GoodFlag = 1 Then
			Rs.Close
			Set Rs = Nothing
			CALL LDExeCute("Update LeadBBS_Announce Set GoodFlag = 0 where id=" & LMT_AncID,1)
			If DEF_UsedDataBase = 1 Then CALL LDExeCute("Update LeadBBS_Topic Set GoodFlag = 0 where id=" & LMT_AncID,1)
			If inStr(application(DEF_MasterCookies & "TopAncList"),"," & LMT_AncID & ",") Then
				UpdateAnnounceApplicationInfo LMT_AncID,12,0,0,0
			Else
				If inStr(application(DEF_MasterCookies & "TopAncList" & GBL_Board_BoardAssort),"," & LMT_AncID & ",") Then UpdateAnnounceApplicationInfo LMT_AncID,12,0,0,GBL_Board_BoardAssort
			End if
			
			Processor_Done "�������Ѿ��Ǿ���������ȡ��������" & VbCrLf
			CALL LDExeCute("Update LeadBBS_User set Points=Points-" & DEF_BBS_MakeGoodAnnouncePoints & ",AnnounceGood=AnnounceGood-1 Where ID =" & UserID,1)
			UpdateBoardAnnounceNum Application(DEF_MasterCookies & "BoardInfo" & BoardID)(28,0),0,0,0,-1
			
			If LMT_UserID > 0 and (LMT_Prc_MsgFlag = 2 or Request.Form("SendMessage") = "1") Then SendNewMessage Prc_User,MakeGood_User,"��̳���ţ�����ȡ������֪ͨ","[color=blue]������������ӱ�ȡ������[/color]" & VbCrLf & VbCrLf &_
				"[b]���ڰ��棺[/b][url=../b/b.asp?b=" & GBL_Board_ID & "]" & htmlencode(KillHTMLLabel(GBL_Board_BoardName)) & "[/url]" & VbCrLf & _
				"[b]������Ա��[/b]" & htmlencode(GBL_CHK_User) & VbCrLf & _
				"[b]����ԭ��[/b]" & htmlencode(Left(Request.Form("SendWhys"),24)) & VbCrLf & _
				"[b]������ӣ�[/b][url=../a/a.asp?b=" & GBL_Board_ID & "&id=" & LMT_AncID & "]" & htmlencode(MakeGood_Title) & "[/url]",GBL_IPAddress
		Else
			Rs.Close
			Set Rs = Nothing
			CALL LDExeCute("Update LeadBBS_Announce Set GoodFlag = 1 where id=" & LMT_AncID,1)
			If DEF_UsedDataBase = 1 Then CALL LDExeCute("Update LeadBBS_Topic Set GoodFlag = 1 where id=" & LMT_AncID,1)
			If inStr(application(DEF_MasterCookies & "TopAncList"),"," & LMT_AncID & ",") Then
				UpdateAnnounceApplicationInfo LMT_AncID,12,1,0,0
			Else
				If inStr(application(DEF_MasterCookies & "TopAncList" & GBL_Board_BoardAssort),"," & LMT_AncID & ",") Then UpdateAnnounceApplicationInfo LMT_AncID,12,1,0,GBL_Board_BoardAssort
			End If
			Processor_Done "�ɹ�������̳���ӣ�"
			CALL LDExeCute("Update LeadBBS_User set Points=Points+" & DEF_BBS_MakeGoodAnnouncePoints & ",AnnounceGood=AnnounceGood+1 Where ID =" & UserID,1)
			CALL LDExeCute("Update LeadBBS_Boards set GoodNum=GoodNum+1 Where BoardID =" & BoardID,1)
			UpdateBoardAnnounceNum Application(DEF_MasterCookies & "BoardInfo" & BoardID)(28,0),0,0,0,1
			CALL LDExeCute("insert into LeadBBS_Log(LogType,LogTime,LogInfo,UserName,IP,BoardID) Values(102," & GetTimeValue(DEF_Now) & ",'�ɹ����� ������ӣ�������" & GBL_Board_ID & "���ӱ��" & LMT_AncID & " ���߱��:" & LMT_UserID & "��','" & Replace(Replace(htmlencode(Left(GBL_CHK_User,14)),"\","\\"),"'","''") & "','" & Replace(GBL_IPAddress,"'","''") & "'," & GBL_Board_ID & ")",1)
			
			If LMT_UserID > 0 and (LMT_Prc_MsgFlag = 2 or Request.Form("SendMessage") = "1") Then SendNewMessage Prc_User,MakeGood_User,"��̳���ţ����Ӿ���֪ͨ","[color=blue]������������ӱ���Ϊ����[/color]" & VbCrLf & VbCrLf &_
				"[b]���ڰ��棺[/b][url=../b/b.asp?b=" & GBL_Board_ID & "]" & htmlencode(KillHTMLLabel(GBL_Board_BoardName)) & "[/url]" & VbCrLf & _
				"[b]������Ա��[/b]" & htmlencode(GBL_CHK_User) & VbCrLf & _
				"[b]����ԭ��[/b]" & htmlencode(Left(Request.Form("SendWhys"),24)) & VbCrLf & _
				"[b]������ӣ�[/b][url=../a/a.asp?b=" & GBL_Board_ID & "&id=" & LMT_AncID & "]" & htmlencode(MakeGood_Title) & "[/url]",GBL_IPAddress
		End if
		If LMT_AncID >= ALL_FirstRootID  Then UpdateBoardApplicationInfo GBL_board_ID,0,33
		If LMT_AncID <= ALL_LastRootID  Then UpdateBoardApplicationInfo GBL_board_ID,0,34
		If CheckSupervisorUserName = 0 Then
			CALL LDExeCute("Update LeadBBS_User Set LastWriteTime=" & GetTimeValue(DEF_Now) & " where ID=" & GBL_UserID,1)
			UpdateSessionValue 13,GetTimeValue(DEF_Now),0
		End If
	Else
		Dim N
		Processor_Head
		%>
		<form name=DellClientForm id=DellClientForm action=Processor.asp?action=MakeGood&b=<%=GBL_Board_ID%>&ID=<%=LMT_AncID%> onSubmit="submit_disable(this);" method="post"<%
		If AjaxFlag = 1 Then
			Response.Write " target=""hidden_frame"""
		End If
		%>>
			<input type=hidden name=SureFlag value="1">
			<input type=hidden name=JsFlag value="1">
			<input type=hidden name=AjaxFlag value="<%=AjaxFlag%>">
			<input type=hidden name=ID value="<%=LMT_AncID%>">
			<input type=hidden name=BoardID value="<%=GBL_Board_ID%>">
			<%
		If DEF_BBS_PrizeAnnouncePoints <> 0 Then%>
			<table><tr><td>
			<div class=title>��ѡ��Ҫ���еĲ���</div>
			<%If MakeGood_Level >= 2 Then%>
			<hr class=splitline>
					<input class=fmchkbox type=radio name=Form_GoodType value=0 checked><span onclick="$id('DellClientForm').Form_GoodType[0].checked=1;" style="cursor:pointer"><%
					If GBL_GoodFlag = 1 Then%>ȡ��<%
					End If%>�������Ϊ<font color=ff0000 class=redfont><%=LMT_AncID%></font>������</span>
					<div class=value2>
					<font color=Gray class=grayfont>��������ɽ����ӹ�����˾���ר�������������̳����ר����<br>��Ӱ������û���Ӧ����</font>
					</div>
			<%End If
			If LMT_UserID > 0 Then  '�ο������޷�����%>
			<hr class=splitline>
				<input class=fmchkbox type=radio name=Form_GoodType id=Opinion2 value=2<%If MakeGood_Level = 1 Then Response.Write " checked=checked"%>><span onclick="$id('Opinion2').checked=1;" style="cursor:pointer">�����ӽ�������</span>
				<div class=value2>
				�����<input maxlength=24 onfocus="$id('Opinion2').checked=1;" name=Form_OpinionWhys value="" size="24" class='fminpt input_3'>
				<span class=grayfont>���24�ֽ�</span>
				</div>
				
				
			<%If MakeGood_Level >= 1 Then%>
				<div class=value2>
				��<%=DEF_PointsName(0)%>��<select name=Form_AddPoints onchange="$id('Opinion2').checked=1;">
				<%
					If MakeGood_Level >=2 Then Response.Write "<option value=0 selected>-----</option>" & VbCrLf
					For N = 1 to abs(DEF_BBS_PrizeAnnouncePoints)
						Response.Write "<option value=" & N & ">���� " & DEF_PointsName(0) & " + " & N & "</option>" & VbCrLf
					Next
					If MakeGood_Level >=2 or DEF_AllowPunish = 1 Then
						For N = abs(DEF_BBS_PrizeAnnouncePoints) to 1 Step -1
							Response.Write "<option value=-" & N & ">���� " & DEF_PointsName(0) & " - " & N & "</option>" & VbCrLf
						Next
					End If
				%>
				</select>
				</div><%If MakeGood_Level = 1 Then%>
						<div class=value2><span class=grayfont>ע�⣺���۷�������������<%=DEF_PointsName(0)%>�п۳�(��������ֵ)<%
						If DEF_AllowOpinionNum > 0 Then Response.Write ", <br>�����������" & DEF_AllowOpinionNum & "��"%></span>
						</div><%
					End If%>
			<%
			End If

			If MakeGood_Level >= 2 and DEF_AllowBoardMasterCachetValue = 1 Then%>
				<div class=value2>
				��<%=DEF_PointsName(2)%>��<select name=Form_OpinionNum onchange="$id('Opinion2').checked=1;">
				<%
					For N = abs(DEF_BBS_PrizeAnnouncePoints) to 1 Step -1
						Response.Write "<option value=-" & N
						Response.Write ">���� " & DEF_PointsName(2) & " - " & N & "</option>" & VbCrLf
					Next
					Response.Write "<option value=0 selected=selected>-----</option>" & VbCrLf
					For N = 1 to abs(DEF_BBS_PrizeAnnouncePoints)
						Response.Write "<option value=" & N
						Response.Write ">���� " & DEF_PointsName(2) & " + " & N & "</option>" & VbCrLf
					Next
				%>
				</select>
				</div>
			<%End If
				
			If MakeGood_Level >= 3 Then '����Ա�ɽ��вƸ�����%>
				
				<div class=value2>
				<%=DEF_PointsName(1)%>���֣�<select name=Form_AddCharm onchange="$id('Opinion2').checked=1;">
				<%
					For N = abs(DEF_BBS_PrizeAnnouncePoints) to 1 Step -1
						Response.Write "<option value=-" & N & ">���� " & DEF_PointsName(1) & " - " & N & "</option>" & VbCrLf
					Next
					Response.Write "<option value=0 selected>�޽���</option>" & VbCrLf
					For N = 1 to abs(DEF_BBS_PrizeAnnouncePoints)
						Response.Write "<option value=" & N & ">���� " & DEF_PointsName(1) & " + " & N & "</option>" & VbCrLf
					Next
				%>
				</select>
				</div>
			<%End If
			End If 'end for guest%>
			</td></tr></table>
			<%
		Else%><b>ȷ��Ҫ<%
			If GBL_GoodFlag = 1 Then%>ȡ��<%
			End If%>�������Ϊ<font color=ff0000 class=redfont><%=LMT_AncID%></font>��������</b>
			<br><%
		End If
		
		If MakeGood_Level >= 2 and LMT_UserID > 0 Then Processor_MsgForm%>
		<br>
		<div class=value2>
		<input type=submit value=ȷ�� class="fmbtn btn_2">
		</div>
		</form>
		<%Processor_Bottom
	End If

End Sub

Sub Opinion_Update(PointsNum,CachetNum,CharmNum,OpinionStr,UserName)

	If PointsNum = 0 and CachetNum = 0 and CharmNum = 0 and OpinionStr = "" Then Exit Sub
	
	Dim OpinionTime
	OpinionTime = GetTimeValue(DEF_Now)
	If Form_OpitionType = 2 Then
		Dim Tmp_Form_OpinionUser,Tmp_Form_OpinionWhys
		
		Tmp_Form_OpinionUser = Form_OpinionStr(0)
		Form_CachetNum = Form_OpinionStr(1)
		If isNumeric(Form_CachetNum) = 0 Then Form_CachetNum = 0
		Form_CachetNum = Fix(cCur(Form_CachetNum))
		Tmp_Form_OpinionWhys = Form_OpinionStr(2)
		Dim OldType
		If Form_CachetNum > 0 Then
			OldType = 1
		Else
			OldType = 0
		End If
		CALL LdExeCute("insert into LeadBBS_Opinion(AnnounceID,UserName,Opinion,NumType,Num,IP,Ndatetime) Values(" & LMT_AncID & ",'" & Replace(Tmp_Form_OpinionUser,"'","''") & "','" & Replace(Tmp_Form_OpinionWhys,"'","''") & "'," & OldType & "," & Form_CachetNum & ",'" & Replace(GBL_IPAddress,"'","''") & "'," & OpinionTime & ")",1)
		
		Form_CachetNum = Form_CachetNum + CachetNum
		Form_CharmNum = CharmNum
		Form_PointsNum = PointsNum
		Form_OpinionCount = 1
	ElseIf Form_OpitionType = 3 Then
		Form_PointsNum = cCur(Form_OpinionStr(0))
		Form_CachetNum = cCur(Form_OpinionStr(1))
		Form_CharmNum = cCur(Form_OpinionStr(2))
		Form_OpinionCount = cCur(Form_OpinionStr(3))
		Form_CachetNum = Form_CachetNum + CachetNum
		Form_CharmNum = Form_CharmNum + CharmNum
		Form_PointsNum = Form_PointsNum + PointsNum
	Else
		Form_CachetNum = CachetNum
		Form_CharmNum = CharmNum
		Form_PointsNum = PointsNum
		Form_OpinionCount = 0
	End If
	If PointsNum <> 0 Then Form_OpinionCount = Form_OpinionCount + 1
	If CachetNum <> 0 Then Form_OpinionCount = Form_OpinionCount + 1
	If CharmNum <> 0 Then Form_OpinionCount = Form_OpinionCount + 1
	Dim UserSQLStr,Tmp
	UserSQLStr = ""
	Tmp = ""
	If PointsNum <> 0 Then
		If UserSQLStr = "" Then
			UserSQLStr = "Points=Points+" & PointsNum
		Else
			UserSQLStr = UserSQLStr & ",Points=Points+" & PointsNum
		End If
		Tmp = DEF_PointsName(0) & " " & PointsNum
		If MakeGood_Level = 1 Then CALL LDExeCute("Update LeadBBS_User Set Points=Points-" & Abs(PointsNum) & " where ID=" & GBL_UserID,1)
		Free_UDT
		CALL LDExeCute("Update LeadBBS_SiteInfo Set SavePoints=SavePoints-" & Abs(PointsNum),1)
		CALL LdExeCute("insert into LeadBBS_Opinion(AnnounceID,UserName,Opinion,NumType,Num,IP,Ndatetime) Values(" & LMT_AncID & ",'" & Replace(UserName,"'","''") & "','" & Replace(OpinionStr,"'","''") & "',0," & PointsNum & ",'" & Replace(GBL_IPAddress,"'","''") & "'," & OpinionTime & ")",1)
	End If
	If CachetNum <> 0 Then
		If UserSQLStr = "" Then
			UserSQLStr = "CachetValue=CachetValue+" & CachetNum
		Else
			UserSQLStr = UserSQLStr & ",CachetValue=CachetValue+" & CachetNum
		End If
		Tmp = Tmp & " " & DEF_PointsName(2) & " " & CachetNum
		CALL LdExeCute("insert into LeadBBS_Opinion(AnnounceID,UserName,Opinion,NumType,Num,IP,Ndatetime) Values(" & LMT_AncID & ",'" & Replace(UserName,"'","''") & "','" & Replace(OpinionStr,"'","''") & "',1," & CachetNum & ",'" & Replace(GBL_IPAddress,"'","''") & "'," & OpinionTime & ")",1)
	End If
	If CharmNum <> 0 Then
		If UserSQLStr = "" Then
			UserSQLStr = "CharmPoint=CharmPoint+" & CharmNum
		Else
			UserSQLStr = UserSQLStr & ",CharmPoint=CharmPoint+" & CharmNum
		End If
		Tmp = Tmp & " " & DEF_PointsName(1) & " " & CharmNum
		CALL LdExeCute("insert into LeadBBS_Opinion(AnnounceID,UserName,Opinion,NumType,Num,IP,Ndatetime) Values(" & LMT_AncID & ",'" & Replace(UserName,"'","''") & "','" & Replace(OpinionStr,"'","''") & "',2," & CharmNum & ",'" & Replace(GBL_IPAddress,"'","''") & "'," & OpinionTime & ")",1)
	End If
	If UserSQLStr <> "" Then CALL LDExeCute("Update LeadBBS_User Set " & UserSQLStr & " where ID=" & LMT_UserID,1)
	Dim Form_Opinion
	Form_Opinion = Form_PointsNum & "|" & Form_CachetNum & "|" & Form_CharmNum & "|" & Form_OpinionCount
	CALL LDExeCute("Update LeadBBS_Announce Set Opinion='" & Replace(Form_Opinion,"'","''") & "' where ID=" & LMT_AncID,1)

	CALL LDExeCute("insert into LeadBBS_Log(LogType,LogTime,LogInfo,UserName,IP,BoardID) Values(105," & OpinionTime & ",'������ӣ�������" & GBL_Board_ID & "���ӱ��" & LMT_AncID & " ���߱��:" & LMT_UserID & " ����:" & Replace(Form_Opinion,"'","''") & "��','" & Replace(Replace(htmlencode(Left(GBL_CHK_User,14)),"\","\\"),"'","''") & "','" & Replace(GBL_IPAddress,"'","''") & "'," & GBL_Board_ID & ")",1)

	If LMT_UserID > 0 and (LMT_Prc_MsgFlag = 2 or Request.Form("SendMessage") = "1") Then SendNewMessage Prc_User,MakeGood_User,"��̳���ţ���������֪ͨ","[color=blue]��������������ܵ�����Ӱ��[/color]" & VbCrLf & VbCrLf &_
		"[b]���棺[/b][url=../b/b.asp?b=" & GBL_Board_ID & "]" & htmlencode(KillHTMLLabel(GBL_Board_BoardName)) & "[/url]" & VbCrLf & _
		"[b]������Ա��[/b]" & htmlencode(GBL_CHK_User) & VbCrLf & _
		"[b]ԭ��[/b]" & htmlencode(Left(Request.Form("SendWhys"),24)) & VbCrLf & _
		"[b]���[/b]" & htmlencode(OpinionStr) & VbCrLf & _
		"[b]���֣�[/b]" & Tmp & VbCrLf & _
		"[b]���ӣ�[/b][url=../a/a.asp?b=" & GBL_Board_ID & "&id=" & LMT_AncID & "]" & htmlencode(MakeGood_Title) & "[/url]",GBL_IPAddress

End Sub

Sub OpinionAnnounce

	If LMT_UserID = 0 Then
		Processor_ErrMsg "<span class=redfont>�ο������޷�����.</span>"
		Exit Sub
	End If
	Dim Form_Opinion
	Form_OpinionWhys = Request.Form("Form_OpinionWhys")

	If StrLength(Form_OpinionWhys) > 24 Then
		Processor_ErrMsg "����ʧ�ܣ������ʲ��ܳ���24���ֽڣ�"
		Exit Sub
	End If

	If inStr(Form_OpinionWhys,"|") or inStr(Form_OpinionWhys,"<") or inStr(Form_OpinionWhys,"""") or inStr(Form_OpinionWhys,"script") Then
		Processor_ErrMsg "����ʧ�ܣ������ʲ��ܺ���|&lt;""script���ַ��򵥴ʣ�"
		Exit Sub
	End If
	
	If LMT_UserID = GBL_UserID Then
		Processor_ErrMsg "���������Լ���������ӣ�"
		Exit Sub
	End If
	
	Form_OpinionUser = GBL_CHK_User

	If MakeGood_Level >= 2 and DEF_AllowBoardMasterCachetValue = 1 Then
		Form_OpinionNum = Request.Form("Form_OpinionNum")
		If isNumeric(Form_OpinionNum) = 0 then Form_OpinionNum = 0
		Form_OpinionNum = Fix(cCur(Form_OpinionNum))
		If Form_OpinionNum < 0-Abs(DEF_BBS_PrizeAnnouncePoints) or Form_OpinionNum > Abs(DEF_BBS_PrizeAnnouncePoints) Then Form_OpinionNum = 0
	Else
		Form_OpinionNum = 0
	End If
	
	Dim Form_AddPoints	
	If MakeGood_Level >= 1 Then
		Form_AddPoints = Request.Form("Form_AddPoints")
		If isNumeric(Form_AddPoints) = 0 then Form_AddPoints = 0
		Form_AddPoints = Fix(cCur(Form_AddPoints))
		If Form_AddPoints < 0-Abs(DEF_BBS_PrizeAnnouncePoints) or Form_AddPoints > Abs(DEF_BBS_PrizeAnnouncePoints) Then Form_AddPoints = 0
		If DEF_AllowPunish = 1 Then
			If MakeGood_Level = 1 and Form_AddPoints = 0 Then
				Processor_Done "�ɹ����������û�" & DEF_PointsName(0) & "���κθĶ���"
				Exit Sub
			End If
		Else
			If MakeGood_Level = 1 and Form_AddPoints < 0 Then Form_AddPoints = 0
			If MakeGood_Level = 1 and Form_AddPoints < 1 Then
				Processor_Done "�ɹ����������û�" & DEF_PointsName(0) & "���κθĶ���"
				Exit Sub
			End If
		End If
		Free_UDT
		CheckPass
		If MakeGood_Level = 1 and GBL_CHK_Points < Abs(Form_AddPoints) Then
			Processor_Done "����" & DEF_PointsName(0) & "����, �޷���ɴ˴����ۣ�"
			Exit Sub
		End If
	Else
		Form_AddPoints = 0
	End If
	
	Dim Form_AddCharm
	If MakeGood_Level >= 3 Then
		Form_AddCharm = Request.Form("Form_AddCharm")
		If isNumeric(Form_AddCharm) = 0 then Form_AddCharm = 0
		Form_AddCharm = Fix(cCur(Form_AddCharm))
		If Form_AddCharm < 0-Abs(DEF_BBS_PrizeAnnouncePoints) or Form_AddCharm > Abs(DEF_BBS_PrizeAnnouncePoints) Then Form_AddCharm = 0
	Else
		Form_AddCharm = 0
	End If

	If Form_AddPoints = 0 and Form_OpinionNum = 0 and Form_AddCharm = 0 Then
		Processor_ErrMsg "��Ч������, ���۱���ѡ�����֣�"
		Exit Sub
	End If
	CALL Opinion_Update(Form_AddPoints,Form_OpinionNum,Form_AddCharm,Form_OpinionWhys,Form_OpinionUser)
	Processor_Done "�ɹ���ɶ����ӵ����ۣ�����¼����־��" & VbCrLf

End Sub
%>