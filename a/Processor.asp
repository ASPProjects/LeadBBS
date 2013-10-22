<!-- #include file=../inc/BBSsetup.asp -->
<!-- #include file=../inc/Board_Popfun.asp -->
<!-- #include file=../inc/Upload_Setup.asp -->
<!-- #include file=../inc/Limit_Fun.asp -->
<!-- #include file=inc/MakeAnnounceTop.asp -->
<!-- #include file=inc/AllTopAnc.asp -->
<!-- #include file=inc/TopAnc.asp -->
<!-- #include file=inc/MoveAnnounce.asp -->
<!-- #include file=inc/RepairAnnounce.asp -->
<!-- #include file=inc/TypeAnnounce.asp -->
<!-- #include file=inc/DelAnnounce.asp -->
<!-- #include file=inc/MakeGoodAnnounce.asp -->
<!-- #include file=inc/AddFriend.asp -->
<!-- #include file=../User/inc/Fun_SendMessage.asp -->
<!-- #include file=inc/DelUpload_Fun.asp -->
<%
Const LMT_MaxCollectAnnounce = 500 '��������ղ���������
Const LMT_Prc_anonymity = 1 '�������Ƿ���������Ϣ֪ͨ�û��� 0 ����Ϊϵͳ 1 ԭ������
Const LMT_Prc_MsgFlag = 2 '����Ա�Ƿ����Ϣ֪ͨ�û��� 0 Ĭ��ѡ��Ϊ��֪ͨ,����ѡ���Ƿ�֪ͨ 1 Ĭ�϶���Ϣ֪ͨ,Ҳ��ѡ���Ƿ�֪ͨ 2.ǿ�ƶ���Ϣ֪ͨ,�����Ƿ�֪ͨ
DEF_BBS_HomeUrl = "../"

Dim Prc_User

Dim LMT_AncID,GoodAssort,Form_ParentID,Form_RootID,Form_RootIDBak,AjaxFlag
Dim RootStr,AllTopFlag,Part,PartStr,Action_Str

Function LoginAccuessFul

	Action_Str = Request("Action")

	If AjaxFlag = 0 Then BBS_SiteHead DEF_SiteNameString & " - " & KillHTMLLabel(GBL_Board_BoardName),GBL_board_ID,"<span class=navigate_string_step>����</span>"
	
	If AjaxFlag = 0 Then Boards_Body_Head("")
	
	If Action_Str <> "Del" and Action_Str <> "Move" Then
		LMT_AncID = Left(Request("ID"),14)
		If isNumeric(LMT_AncID) = 0 or inStr(LMT_AncID,",") > 0 or LMT_AncID = "" Then LMT_AncID = 0
		LMT_AncID = Fix(cCur(LMT_AncID))
	Else
		LMT_AncID = Request("ID")
		If InStr(LMT_AncID,",") > 0 Then
			Dim TmpMsg,i
			TmpMsg = Split(LMT_AncID,",")
			If Ubound(TmpMsg,1) >= DEF_MaxListNum and Ubound(TmpMsg,1) >= DEF_TopicContentMaxListNum Then
				LMT_AncID = 0
			Else
				LMT_AncID = ""
				For i = 0 to Ubound(TmpMsg,1)
					If isNumeric(TmpMsg(i)) = 0 Then
						LMT_AncID = 0
						Exit For
					Else
						If LMT_AncID = "" Then
							LMT_AncID = Fix(cCur(TmpMsg(i)))
						Else
							LMT_AncID = LMT_AncID & "," & Fix(cCur(TmpMsg(i)))
						End If
					End If
				Next
			End If
			If LMT_AncID = "" Then LMT_AncID = 0
		Else
			If isNumeric(LMT_AncID) = 0 or LMT_AncID = "" Then LMT_AncID = 0
			LMT_AncID = Fix(cCur(LMT_AncID))
		End If
	End If
	
	If GBL_CHK_Flag=1 Then
		If LMT_Prc_anonymity = 1 Then
			Prc_User = "[LeadBBS]"
		Else
			Prc_User = GBL_CHK_User
		End If
		Select Case Action_Str
			Case "Collect": If CheckIsCanCollSure = 1 Then DisplayCollectAnnounce
			Case "Top": If CheckTopSure = 1 Then DisplayMakeTopAnnounce
			Case "Repair": If CheckRepairSure = 1 Then DisplayRepairAnnounce
			Case "mirror","Move": If CheckMoveSure = 1 Then DisplayMoveAnnounce
			Case "TopAnc": If CheckTopAncSure = 1 Then DisplayTopAncAnnounce
			Case "AllTopAnc": If CheckAllTopAncSure = 1 Then DisplayAllTopAncAnnounce
			Case "TypeSet": If CheckTypeSetSure = 1 Then DisplayTypeSetAnnounce
			Case "Del": If CheckDelSure = 1 Then DisplayDelAnnounce
			Case "MakeGood": If CheckMakeGoodSure = 1 Then DisplayMakeGoodAnnounce
			Case "AddFriend": If CheckAddFriendSure = 1 Then DisplayAddFriend
			Case Else: Processor_ErrMsg "δѡ��������"
		End Select
	Else
		If Request("submitflag") = "" Then
			Processor_ErrMsg "���ȵ�¼��"
		Else
			Processor_ErrMsg GBL_CHK_TempStr
		End If
	End If

End Function

Function CheckSure

	If LMT_AncID = 0 Then
		Processor_ErrMsg "����ѡ��Ҫ�����ļ�¼��" & VbCrLf
		CheckSure = 0
		Exit Function
	End if

	Dim Rs,SQL,TmpArr,TmpID
	If inStr(LMT_AncID,",") Then
		TmpArr = Split(LMT_AncID,",")
		TmpID = TmpArr(0)
	Else
		TmpID = LMT_AncID
	End If
	SQL = sql_select("Select BoardID,GoodAssort,ParentID,UserID,RootID,RootIDBak from LeadBBS_Announce where id=" & TmpID,1)
	Set Rs = LDExeCute(SQL,0)
	If Rs.Eof Then
		Processor_ErrMsg "ѡ��ļ�¼�Ѳ����ڡ�" & VbCrLf
		Rs.Close
		Set Rs = Nothing
		CheckSure = 0
		Exit Function
	End if

	GBL_Board_ID = Rs("BoardID")
	GoodAssort = cCur(Rs("GoodAssort"))
	Form_ParentID = cCur(Rs("ParentID"))
	Form_RootID = cCur(Rs("RootID"))
	Form_RootIDBak = cCur(Rs("RootIDBak"))
	Rs.Close
	Set Rs = Nothing

	Dim Temp
	Temp = Application(DEF_MasterCookies & "BoardInfo" & GBL_Board_ID)
	If isArray(Temp) = False Then
		ReloadBoardInfo(GBL_Board_ID)
		Temp = Application(DEF_MasterCookies & "BoardInfo" & GBL_Board_ID)
	End If
	If isArray(Temp) = False Then
		Processor_ErrMsg "��̳������������ϵ����Ա��" & VbCrLf
		CheckSure = 0
		Set Rs = Nothing
	End If
	GBL_Board_BoardName = Temp(0,0)
	GBL_Board_MasterList = Temp(10,0)
	GBL_Board_BoardLimit = Temp(9,0)
	GBL_Board_BoardAssort = cCur(Temp(1,0))
	GBL_Board_AssortMaster = Temp(35,0)
	CheckSure = 1
	
End Function

Function CheckTopSure

	If CheckSure = 0 Then Exit Function
	
	If Form_ParentID <> 0 Then
		Processor_ErrMsg "Ҫ��������ӱ���Ϊ��������"
		CheckTopSure = 0
		Exit Function
	End if

	CheckisBoardMaster
	If GBL_UserID >= 1 and (GBL_BoardMasterFlag >= 5 and GetBinarybit(GBL_CHK_UserLimit,4) = 0) Then
		CheckTopSure = 1
	Else
		CheckTopSure = 0
		Processor_ErrMsg "����Ȩ�޲��㣮"
	End If

End Function


Function DisplayMakeTopAnnounce

	If Request.Form("SureFlag")="1" Then
		CALL MakeAnnounceTop(LMT_AncID,"")
		If CheckSupervisorUserName = 0 Then
			CALL LDExeCute("Update LeadBBS_User Set LastWriteTime=" & GetTimeValue(DEF_Now) & " where ID=" & GBL_UserID,1)
			UpdateSessionValue 13,GetTimeValue(DEF_Now),0
		End If
		Processor_Done("�ɹ�������̳���ӣ�")
	Else
		Processor_form "Top","����"
	End If

End Function

Rem �ղ�����
Function CheckIsCanCollSure

	If CheckSure = 0 Then Exit Function
	If Form_ParentID <> 0 Then
		Processor_ErrMsg "Ҫ��������ӱ���Ϊ��������"
		CheckIsCanCollSure = 0
		Exit Function
	End if

	CheckisBoardMaster
	CheckAccessLimit
	If GBL_CHK_TempStr <> "" or GetBinarybit(GBL_CHK_UserLimit,1) = 1 Then
		Processor_ErrMsg "����Ȩ�޲���." & VbCrLf
		CheckIsCanCollSure = 0
		Exit Function
	End If
	
	If CheckWriteEventSpace = 0 Then
		Processor_ErrMsg "<b><font color=Red Class=redfont>���Ĳ�����Ƶ�����Ժ�����!</font></b> <br>" & VbCrLf
		CheckIsCanCollSure = 0
		Exit Function
	End If
	CheckIsCanCollSure = 1

End Function

Sub DisplayCollectAnnounce

	If LMT_AncID = 0 Then
		Processor_ErrMsg "����Ҫ�ղص����ⲻ���ڣ�" & VbCrLf
		Exit Sub
	End if
	If Request.Form("SureFlag")="1" Then
		Dim Rs,SQL
		SQL = "Select count(*) from LeadBBS_CollectAnc where UserID=" & GBL_UserID
		Set Rs = LDExeCute(SQL,0)
		If Rs.Eof Then
			SQL = 0
		Else
			SQL = Rs(0)
			If IsNull(SQL) Then SQL = 0
			SQL = cCur(SQL)
		End If
		Rs.Close
		Set Rs = Nothing

		If SQL > LMT_MaxCollectAnnounce Then
			Processor_ErrMsg "�������ղ������Ѿ�����" & LMT_MaxCollectAnnounce & "�����������ղأ�" & VbCrLf
			Exit Sub
		End if
		
		SQL = sql_select("Select ID from LeadBBS_CollectAnc where AnnounceID=" & LMT_AncID & " and UserID=" & GBL_UserID,1)
		Set Rs = LDExeCute(SQL,0)
		If Not Rs.Eof Then
			Processor_ErrMsg "���Ѿ��ղع�������"
			//Processor_ErrMsg "<div id=collect_msg>���Ѿ��ղع������� <a href=""javascript:p_url = '" & DEF_BBS_HomeUrl & "User/DeleteMessage.asp';" & VbCrLf & "p_para='AjaxFlag=1&FriendFlag=2&DeleteSureFlag=dk9@dl9s92lw_SWxl&MessageID=';" & VbCrLf & "p_command = '$id(\'collect_msg\').innerHTML=tmp';" & VbCrLf & "p_type = 1;" & VbCrLf & "p_once(" & Rs(0) & ");"">�������ɾ�����ղء�</a></div>" & VbCrLf
			Rs.Close
			Set Rs = Nothing
			Exit Sub
		End If
		Rs.Close
		Set Rs = Nothing

		Processor_Done("����Ҫ�ղص������ѳɹ�������ղ��б�")
		CALL LDExeCute(" insert into LeadBBS_CollectAnc(AnnounceID,UserID) Values(" & LMT_AncID & "," & GBL_UserID & ")",1)

		Set Rs = Nothing
		If CheckSupervisorUserName = 0 Then
			CALL LDExeCute("Update LeadBBS_User Set LastWriteTime=" & GetTimeValue(DEF_Now) & " where ID=" & GBL_UserID,1)
			UpdateSessionValue 13,GetTimeValue(DEF_Now),0
		End If
	Else
		Processor_form "Collect","�ղ�"
	End If

End Sub

Sub Processor_MsgForm

	%>
	<div class=value2>
	<input class=fmchkbox type="checkbox" name=SendMessage value=1<%
	Select Case LMT_Prc_MsgFlag
		Case 0
			Response.Write ""
		Case 1
			Response.Write " checked"
		Case 2
			Response.Write " checked disabled"
	End Select%> onclick="if(this.checked){$id('SendWhyshidden').style.display='block';}else{$id('SendWhyshidden').style.display='none';}">����Ϣ֪ͨ������
	</div>
	<div class=value2>
	<span id=SendWhyshidden name=SendWhyshidden<%If LMT_Prc_MsgFlag = 0 Then Response.Write " style=display:none"%>> 
	����ԭ�� <input maxlength=24 name=SendWhys id=SendWhys value="" size="15" class='fminpt input_3'>
	<select name=swys onchange="$id('SendWhys').value=this.value;">
	<option value="">--ѡ��ԭ��--
	<option value="����Υ��">����Υ��
	<option value="�����">�����
	<option value="�����ˮ">�����ˮ
	<option value="�ظ���">�ظ���
	<option value="">
	<option value="����ԭ��">����ԭ��
	<option value="��ʾ��ͬ">��ʾ��ͬ
	<option value="����">����
	<option value="���а���">���а���
	</select>
	</span>
	</div>
	<%

End Sub

Sub Processor_form(Action,Str)

	Processor_Head
	%>
	<form action=<%=DEF_BBS_HomeUrl%>a/Processor.asp?action=<%=Action%>&b=<%=GBL_Board_ID%>&ID=<%=LMT_AncID%> onSubmit="submit_disable(this);" method="post"<%
	If AjaxFlag = 1 Then
		Response.Write " target=""hidden_frame"""
	End If
	%>>
		<input type=hidden name=SureFlag value="1">
		<input type=hidden name=JsFlag value="1">
		<input type=hidden name=AjaxFlag value="<%=AjaxFlag%>">
		<input type=hidden name=ID value="<%=LMT_AncID%>">
		<input type=hidden name=BoardID value="<%=GBL_Board_ID%>">
		<div class=value2><%=Str%>���ӣ�<%
		If inStr(LMT_AncID,",") Then
			Response.Write "<b>��" & Len(LMT_AncID)-Len(Replace(LMT_AncID,",","")) + 1 & "����¼</b>"
		Else
			Response.Write "<b>��1����¼</b>"
		End If
		If Action = "Del" Then%></div>
		<%
			Processor_MsgForm
		End If
		%>
		<br><div class=value2><input type=submit value=ȷ�� class="fmbtn btn_2"></div>
	</form>
	<%
	Processor_Bottom

End Sub

Sub Processor_Head

	If AjaxFlag = 1 Then
	%>
	<div class="ajaxbox">
	<%
	Else
	%>
	<div class="alertbox">
	<%
	End If

End Sub

Sub Processor_Bottom

%>
	</div>
<%

End Sub

Sub Processor_Done(Str)

	Processor_Head
	Dim Fresh,url
	Select Case Action_Str
		Case "Del"
			Fresh = 1
			url = "a.asp?b=" & GBL_Board_ID & "&id=" & Form_RootIDBak
		Case "Move"
			Fresh = 1
			url = DEF_BBS_HomeUrl & "b/b.asp?b=" & GBL_Board_ID
		Case Else
			Fresh = 0
			url = ""
	End Select
	If Action_Str = "Del" or Action_Str = "Move" Then
		Fresh = 1
	Else
		Fresh = 0
	End If
	If AjaxFlag = 1 and Request.Form("JsFlag")="1" Then%>
	<script>parent.layer_outmsg("anc_delbody","<div class=\"ajaxbox\"><div class='value2 greenfont'><%=Replace(Replace(Replace(Str,"\","\\"),"""","\"""),VbCrLf,"")%></div></div>","<%=url%>");</script>
	<%
	Else
		Response.Write "<div class='value2 greenfont'><b>" & Str & "</b></div>"
	End If
	If AjaxFlag = 0 Then
		Response.Write "<ul>"
		If LMT_AncID > 0 Then Response.Write "<li><a href=a.asp?b=" & GBL_Board_ID & "&ID=" & LMT_AncID & ">���ص�ǰ����</a></li>"
		If GBL_Board_ID > 0 Then Response.Write "<li><a href=../b/b.asp?b=" & GBL_Board_ID & ">���ص�ǰ����</a></li>"
		Response.Write "<ul><br>"
	End If
	Processor_Bottom

End Sub

Sub ChangeGoodAssort(ID,ID2)

	If ID = ID2 Then Exit Sub
	Dim TArray,N,Num,NN,ExitN
	TArray = Application(DEF_MasterCookies & "BoardInfo" & GBL_Board_ID & "_TI")
	If isArray(TArray) = False Then
		'ChangeGoodAssort = 0
		Exit Sub
	End If
	Num = Ubound(TArray,2)
	NN = 0
	ExitN = 2
	If ID = 0 or ID2 = 0 Then ExitN = 1
	For N = 0 To Num
		If ID = cCur(TArray(0,N)) Then
			If cCur(TArray(2,N)) = -1 Then
				TArray(3,N) = 0
				TArray(4,N) = 0
			Else
				TArray(2,N) = cCur(TArray(2,N)) - 1
				TArray(3,N) = 0
				TArray(4,N) = 0
			End If
			NN = NN + 1
			If NN >= ExitN Then Exit For
		End If
		If ID2 = cCur(TArray(0,N)) Then
			If cCur(TArray(2,N)) <> 0 Then
				If cCur(TArray(2,N)) = -1 Then
					TArray(2,N) = 1
					TArray(3,N) = 0
					TArray(4,N) = 0
				Else
					TArray(2,N) = cCur(TArray(2,N)) + 1
					TArray(3,N) = 0
					TArray(4,N) = 0
				End If
			End If
			TArray(2,N) = 0
			NN = NN + 1
			If NN >= ExitN Then Exit For
		End If
	Next
	If NN > 0 Then
		Application.Lock
		Application(DEF_MasterCookies & "BoardInfo" & GBL_Board_ID & "_TI") = TArray
		Application.UnLock
	End If

	If NN < 2 Then
		TArray = Application(DEF_MasterCookies & "BoardInfo" & 0 & "_TI")
		If isArray(TArray) = False Then
			'ChangeGoodAssort = 0
			Exit Sub
		End If
		Num = Ubound(TArray,2)
		NN = 0
		ExitN = 2
		If ID = 0 or ID2 = 0 Then ExitN = 1
		For N = 0 To Num
			If ID = cCur(TArray(0,N)) Then
				If cCur(TArray(2,N)) = -1 Then
					TArray(3,N) = 0
					TArray(4,N) = 0
				Else
					TArray(2,N) = cCur(TArray(2,N)) - 1
					TArray(3,N) = 0
					TArray(4,N) = 0
				End If
				NN = NN + 1
				If NN >= ExitN Then Exit For
			End If
			If ID2 = cCur(TArray(0,N)) Then
				If cCur(TArray(2,N)) <> 0 Then
					If cCur(TArray(2,N)) = -1 Then
						TArray(2,N) = 1
						TArray(3,N) = 0
						TArray(4,N) = 0
					Else
						TArray(2,N) = cCur(TArray(2,N)) + 1
						TArray(3,N) = 0
						TArray(4,N) = 0
					End If
				End If
				TArray(2,N) = 0
				NN = NN + 1
				If NN >= ExitN Then Exit For
			End If
		Next
		If NN > 0 Then
			Application.Lock
			Application(DEF_MasterCookies & "BoardInfo" & 0 & "_TI") = TArray
			Application.UnLock
		End If
	End If

End Sub

Function CheckGoodAssort(ID)

	Dim Rs
	If ID = 0 Then
		CheckGoodAssort = 1
		Exit Function
	End If
	Set Rs = LDExeCute(sql_select("Select BoardID from LeadBBS_GoodAssort where ID=" & ID,1),0)
	If Rs.Eof Then
		CheckGoodAssort = 0
	Else
		If cCur(Rs(0)) <> GBL_Board_ID and cCur(Rs(0)) <> 0 Then
			CheckGoodAssort = 0
		Else
			CheckGoodAssort = 1
		End If
	End if
	Rs.Close
	Set Rs = Nothing

End Function

Function UpdateBoardAnnounceNum(BoardList,TopicNum,AnnounceNum,TodayAnnounce,GoodNum)

	Dim SQL,N
	If BoardList = "" or (TopicNum = 0 and AnnounceNum = 0 and TodayAnnounce = 0 and GoodNum = 0) Then Exit Function
	SQL = "Update LeadBBS_Boards Set AnnounceNum=AnnounceNum+" & AnnounceNum & ",AnnounceNum_All=AnnounceNum_All+" & AnnounceNum
	If TopicNum <> 0 Then SQL = SQL & ",TopicNum=TopicNum+" & TopicNum
	If TopicNum <> 0 Then SQL = SQL & ",TopicNum_All=TopicNum_All+" & TopicNum
	If TodayAnnounce <> 0 Then  SQL = SQL & ",TodayAnnounce=TodayAnnounce+" & TodayAnnounce
	If TodayAnnounce <> 0 Then  SQL = SQL & ",TodayAnnounce_All=TodayAnnounce_All+" & TodayAnnounce
	If GoodNum <> 0 Then  SQL = SQL & ",GoodNum=GoodNum+" & GoodNum
	If GoodNum <> 0 Then  SQL = SQL & ",GoodNum_All=GoodNum_All+" & GoodNum
	SQL = SQL & " where BoardID in(" & BoardList & ")"
	CALL LDExeCute(SQL,1)
	BoardList = Split(BoardList,",")
	SQL = Ubound(BoardList,1)
	Dim Temp
	For N = 0 To SQL
		Temp = Application(DEF_MasterCookies & "BoardInfo" & BoardList(N))
		If isArray(Temp) = False Then
			ReloadBoardInfo(BoardList(N))
			Temp = Application(DEF_MasterCookies & "BoardInfo" & BoardList(N))
		End If
		If isArray(Temp) = True Then
			If TopicNum <> 0 Then
				UpdateBoardApplicationInfo BoardList(N),cCur(Temp(5,0))+TopicNum,5
				UpdateBoardApplicationInfo BoardList(N),cCur(Temp(29,0))+TopicNum,29
			End If
			If AnnounceNum <> 0 Then
				UpdateBoardApplicationInfo BoardList(N),cCur(Temp(6,0))+AnnounceNum,6
				UpdateBoardApplicationInfo BoardList(N),cCur(Temp(30,0))+AnnounceNum,30
			End If
			If TodayAnnounce <> 0 Then
				UpdateBoardApplicationInfo BoardList(N),cCur(Temp(18,0))+TodayAnnounce,18
				UpdateBoardApplicationInfo BoardList(N),cCur(Temp(31,0))+TodayAnnounce,31
			End If
			If GoodNum <> 0 Then
				UpdateBoardApplicationInfo BoardList(N),cCur(Temp(13,0))+GoodNum,13
				UpdateBoardApplicationInfo BoardList(N),cCur(Temp(32,0))+GoodNum,32
			End If
		End If
	Next
	'28,T1.ParentBoardStr,29.TopicNum_All,30.AnnounceNum_All,31.TodayAnnounce_All,32.GoodNum_All

End Function

Function DeleteAllTopData(MoveID)

	Dim Rs,Num
	Dim GetDataTop
	GetDataTop = application(DEF_MasterCookies & "TopAnc")
	If isArray(GetDataTop) = False Then
		If GetDataTop <> "yes" Then
			ReloadTopAnnounceInfo(0)
			GetDataTop = application(DEF_MasterCookies & "TopAnc")
		End If
	End If

	If isArray(GetDataTop) = False Then
		Num = 0
	Else
		Num = Ubound(GetDataTop,2) + 1
	End If

	For Rs = 1 to Num
		If cCur(GetDataTop(0,Rs-1)) = MoveID Then
			CALL LDExeCute("Delete from LeadBBS_TopAnnounce Where RootID=" & MoveID,1)
			Exit Function
		End If
	Next

	GetDataTop = application(DEF_MasterCookies & "TopAnc" & GBL_Board_BoardAssort)
	If isArray(GetDataTop) = False Then
		If GetDataTop <> "yes" Then
			ReloadTopAnnounceInfo(GBL_Board_BoardAssort)
			GetDataTop = application(DEF_MasterCookies & "TopAnc" & GBL_Board_BoardAssort)
		End If
	End If

	If isArray(GetDataTop) = False Then
		Num = 0
	Else
		Num = Ubound(GetDataTop,2) + 1
	End If

	For Rs = 1 to Num
		If cCur(GetDataTop(0,Rs-1)) = MoveID Then
			CALL LDExeCute("Delete from LeadBBS_TopAnnounce Where RootID=" & MoveID,1)
			Exit Function
		End If
	Next

End Function

Function UpdateBoardAnnounceNum(BoardList,TopicNum,AnnounceNum,TodayAnnounce,GoodNum)

	Dim SQL,N
	If BoardList = "" or (TopicNum = 0 and AnnounceNum = 0 and TodayAnnounce = 0 and GoodNum = 0) Then Exit Function
	SQL = "Update LeadBBS_Boards Set AnnounceNum=AnnounceNum+" & AnnounceNum & ",AnnounceNum_All=AnnounceNum_All+" & AnnounceNum
	If TopicNum <> 0 Then SQL = SQL & ",TopicNum=TopicNum+" & TopicNum
	If TopicNum <> 0 Then SQL = SQL & ",TopicNum_All=TopicNum_All+" & TopicNum
	If TodayAnnounce <> 0 Then  SQL = SQL & ",TodayAnnounce=TodayAnnounce+" & TodayAnnounce
	If TodayAnnounce <> 0 Then  SQL = SQL & ",TodayAnnounce_All=TodayAnnounce_All+" & TodayAnnounce
	If GoodNum <> 0 Then  SQL = SQL & ",GoodNum=GoodNum+" & GoodNum
	If GoodNum <> 0 Then  SQL = SQL & ",GoodNum_All=GoodNum_All+" & GoodNum
	SQL = SQL & " where BoardID in(" & BoardList & ")"
	CALL LDExeCute(SQL,1)
	BoardList = Split(BoardList,",")
	SQL = Ubound(BoardList,1)
	Dim Temp
	For N = 0 To SQL
		Temp = Application(DEF_MasterCookies & "BoardInfo" & BoardList(N))
		If isArray(Temp) = False Then
			ReloadBoardInfo(BoardList(N))
			Temp = Application(DEF_MasterCookies & "BoardInfo" & BoardList(N))
		End If
		If isArray(Temp) = True Then
			If TopicNum <> 0 Then
				UpdateBoardApplicationInfo BoardList(N),cCur(Temp(5,0))+TopicNum,5
				UpdateBoardApplicationInfo BoardList(N),cCur(Temp(29,0))+TopicNum,29
			End If
			If AnnounceNum <> 0 Then
				UpdateBoardApplicationInfo BoardList(N),cCur(Temp(6,0))+AnnounceNum,6
				UpdateBoardApplicationInfo BoardList(N),cCur(Temp(30,0))+AnnounceNum,30
			End If
			If TodayAnnounce <> 0 Then
				UpdateBoardApplicationInfo BoardList(N),cCur(Temp(18,0))+TodayAnnounce,18
				UpdateBoardApplicationInfo BoardList(N),cCur(Temp(31,0))+TodayAnnounce,31
			End If
			If GoodNum <> 0 Then
				UpdateBoardApplicationInfo BoardList(N),cCur(Temp(13,0))+GoodNum,13
				UpdateBoardApplicationInfo BoardList(N),cCur(Temp(32,0))+GoodNum,32
			End If
		End If
	Next
	'28,T1.ParentBoardStr,29.TopicNum_All,30.AnnounceNum_All,31.TodayAnnounce_All,32.GoodNum_All

End Function

Sub Processor_ErrMsg(str)

	If AjaxFlag = 0 Then
		Global_ErrMsg(str)
	Else
		If AjaxFlag = 1 and Request.Form("JsFlag")="1" Then%>
		<script>parent.layer_outmsg("anc_delbody","<div class=\"ajaxbox\"><b>��ʾ��Ϣ</b>��<span class=\"redfont\"><%=Replace(Replace(Replace(Str,"\","\\"),"""","\"""),VbCrLf,"\n")%></span></div>");</script>
		<%
		Else%>
	<div class="ajaxbox">
	<b>��ʾ��Ϣ</b>��
		<span class="redfont">
			<%=Str%>
		</span>
	</div>
	<%	End If
	End If

End Sub

Sub Main

	GBL_CHK_PWdFlag = 1
	GBL_CHK_GuestFlag = 0

	If Request("AjaxFlag") = "1" Then
		AjaxFlag = 1
	Else
		AjaxFlag = 0
	End If

	initDatabase
	LoginAccuessFul
	closeDataBase
	If AjaxFlag = 0 Then
		Boards_Body_Bottom
		If GBL_ShowBottomSure = 0 Then GBL_SiteBottomString = ""
		SiteBottom
	End If

End Sub

Main
%>