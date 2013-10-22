<!-- #include file=../../inc/ubbCode.asp -->
<%
Dim DoingFlag,Form_NotReplay,Form_TitleStyle,Form_Title,Form_UserLimit,Form_AncUserID,Form_AncUserName

Function CheckTypeSetSure

	Dim UserID
	If LMT_AncID = 0 Then
		Processor_ErrMsg "�������ṩҪ�Զ��Ű�����ӵ�ID��" & VbCrLf
		CheckTypeSetSure = 0
		Exit Function
	End if
	Dim Rs,SQL
	SQL = sql_select("Select TA.BoardID,TA.UserID,TA.NotReplay,TA.TitleStyle,TA.ParentID,TA.RootIDBAK,TA.Title,TU.UserLimit,TA.UserID,TA.UserName from LeadBBS_Announce as TA left join LeadBBS_User as TU on TA.UserID=TU.ID where TA.id=" & LMT_AncID,1)
	Set Rs = LDExeCute(SQL,0)
	If Rs.Eof Then
		Processor_ErrMsg "����δѡ��Ҫ��������ӣ�" & VbCrLf
		Rs.Close
		Set Rs = Nothing
		CheckTypeSetSure = 0
		Exit Function
	End if

	GBL_Board_ID = Rs(0)
	UserID = cCur(Rs(1))
	Form_NotReplay = Rs(2)
	Form_TitleStyle = Rs(3)
	Form_ParentID = cCur(Rs(4))
	Form_RootIDBAK = cCur(Rs(5))
	Form_Title = Rs(6)
	Form_UserLimit = cCur("0" & Rs(7))
	Form_AncUserID = cCur(Rs(8))
	Form_AncUserName = Rs(9)
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
		CheckTypeSetSure = 0
		Set Rs = Nothing
	End If
	GBL_Board_BoardAssort = cCur(Temp(1,0))
	GBL_Board_MasterList = Temp(10,0)
	
	CheckisBoardMaster
	If GBL_UserID >= 1 and (GBL_BoardMasterFlag >= 5 and GetBinarybit(GBL_CHK_UserLimit,4) = 0) Then
		CheckTypeSetSure = 1
		DoingFlag = Request.Form("DoingFlag")
		If DoingFlag <> "0" and DoingFlag <> "1" and DoingFlag <> "2" and DoingFlag <> "3" Then
			DoingFlag = 0
		Else
			If GBL_BoardMasterFlag < 7 and DoingFlag = "3" Then DoingFlag = 0
		End If
		DoingFlag = cCur(DoingFlag)
	Else
		DoingFlag = 0
		If (UserID = GBL_UserID) Then
			CheckTypeSetSure = 1
		Else
			CheckTypeSetSure = 0
			Processor_ErrMsg "����Ȩ�޲��㣡"
		End If
	End If

End Function

Sub DisplayTypeSetAnnounce

	If LMT_AncID = 0 Then
		Processor_ErrMsg "�������ṩҪ�Զ��Ű�����ӵ�ID��" & VbCrLf
		Exit Sub
	End if
	If Request.Form("SureFlag")="1" Then
		If CheckWriteEventSpace = 0 Then
			Processor_ErrMsg "<font color=red class=redfont>���Ĳ�����Ƶ�����Ժ�ˢ�����ԣ�</font>"
			Exit Sub
		End If
		Select Case DoingFlag
			Case 1:	If Form_NotReplay = 0 Then
						Form_NotReplay = 1
						Processor_Done "�ɹ��������ӡ�"
					Else
						Form_NotReplay = 0
						Processor_Done "�ɹ�������ӽ�����"
					End If
					CALL LDExeCute("Update LeadBBS_Announce Set NotReplay=" & Form_NotReplay & " where ID=" & LMT_AncID,1)
					If DEF_UsedDataBase = 1 Then CALL LDExeCute("Update LeadBBS_Topic Set NotReplay=" & Form_NotReplay & " where ID=" & LMT_AncID,1)
			Case 2:	If Form_TitleStyle >= 60 Then	
					Form_TitleStyle = Form_TitleStyle - 60
					If inStr(application(DEF_MasterCookies & "TopAncList"),"," & LMT_AncID & ",") Then
						UpdateAnnounceApplicationInfo LMT_AncID,2,Form_Title,0,0
						UpdateAnnounceApplicationInfo LMT_AncID,16,Form_TitleStyle,0,0
					Else
						If inStr(application(DEF_MasterCookies & "TopAncList" & GBL_Board_BoardAssort),"," & LMT_AncID & ",") Then
							UpdateAnnounceApplicationInfo LMT_AncID,2,Form_Title,0,GBL_Board_BoardAssort
							UpdateAnnounceApplicationInfo LMT_AncID,16,Form_TitleStyle,0,GBL_Board_BoardAssort
						End If
					End If
					If Form_TitleStyle = 1 Then Form_Title = KillHTMLLabel(Form_Title)
					Processor_Done "���ӳɹ�ͨ����˲�����"
				Else
					Form_TitleStyle = Form_TitleStyle + 60
					Form_Title = "���������..."
					If inStr(application(DEF_MasterCookies & "TopAncList"),"," & LMT_AncID & ",") Then
						UpdateAnnounceApplicationInfo LMT_AncID,2,Form_Title,0,0
						UpdateAnnounceApplicationInfo LMT_AncID,16,Form_TitleStyle,0,0
					Else
						If inStr(application(DEF_MasterCookies & "TopAncList" & GBL_Board_BoardAssort),"," & LMT_AncID & ",") Then
							UpdateAnnounceApplicationInfo LMT_AncID,2,Form_Title,0,GBL_Board_BoardAssort
							UpdateAnnounceApplicationInfo LMT_AncID,16,Form_TitleStyle,0,GBL_Board_BoardAssort
						End If
					End If
					Processor_Done "���ӹرճɹ���"
				End If
				CALL LDExeCute("Update LeadBBS_Announce Set TitleStyle=" & Form_TitleStyle & " where ID=" & LMT_AncID,1)
				If DEF_UsedDataBase = 1 Then CALL LDExeCute("Update LeadBBS_Topic Set TitleStyle=" & Form_TitleStyle & " where ID=" & LMT_AncID,1)
				If Form_ParentID = 0 Then UpdateBoardLastAnnounce
			Case 3:
				If Form_AncUserID > 0 and inStr(LCase(DEF_SupervisorUserName),"," & LCase(Form_AncUserName) & ",") = 0 Then
					If Form_TitleStyle <> 30 Then
						If GetBinarybit(Form_UserLimit,3) = 1 or GetBinarybit(Form_UserLimit,7) = 1 Then
							Processor_Done "���û��ѱ����Ի����η��ԣ�����Ҫ�ظ�������"
						Else
							Form_UserLimit = SetBinarybit(Form_UserLimit,3,1)
							CALL UpdateSpecialUserTable2(Form_UserLimit,Form_AncUserID,Form_AncUserName,3,4)
							CALL LDExeCute("Update LeadBBS_User Set UserLimit=" & Form_UserLimit & " where ID=" & Form_AncUserID,1)
							CALL LDExeCute("Update LeadBBS_Announce Set TitleStyle=30,OtherInfo='������" & Replace(GBL_CHK_User,"'","''") & "��" & RestoreTime(GetTimeValue(DEF_Now)) & "��ǲ������û���' where ID=" & LMT_AncID,1)
							If DEF_UsedDataBase = 1 Then CALL LDExeCute("Update LeadBBS_Topic Set TitleStyle=30 where ID=" & LMT_AncID,1)
							Processor_Done "�ɹ������û���" & htmlencode(Form_AncUserName) & "������Ǵ�����"
						End If
					Else
						If GetBinarybit(Form_UserLimit,3) = 1 Then
							Form_UserLimit = SetBinarybit(Form_UserLimit,3,0)
							CALL UpdateSpecialUserTable2(Form_UserLimit,Form_AncUserID,Form_AncUserName,3,4)
							CALL LDExeCute("Update LeadBBS_User Set UserLimit=" & Form_UserLimit & " where ID=" & Form_AncUserID,1)
						End If
						CALL LDExeCute("Update LeadBBS_Announce Set TitleStyle=0,OtherInfo='' where ID=" & LMT_AncID,1)
						If DEF_UsedDataBase = 1 Then CALL LDExeCute("Update LeadBBS_Topic Set TitleStyle=0 where ID=" & LMT_AncID,1)
						Processor_Done "�ɹ���������û���" & htmlencode(Form_AncUserName) & "���������Ǵ�����"
					End If
				Else
					Processor_Done "���㹻Ȩ�ޣ���������ֹ��"
				End If
			Case Else:
				GBL_CHK_TempStr = ""
				ReMakeIDDoc(LMT_AncID)
				If GBL_CHK_TempStr <> "" Then
					Processor_ErrMsg GBL_CHK_TempStr & VbCrLf
				Else
					Processor_Done "�ɹ�����������Զ��Ű档"
				End If
		End Select
	Else
		Processor_Head
		%>
		<form name=DellClientForm action=Processor.asp?Action=TypeSet&b=<%=GBL_Board_ID%> onSubmit="submit_disable(this);" method="post"<%
	If AjaxFlag = 1 Then
		Response.Write " target=""hidden_frame"""
	End If
	%>>
			<input type=hidden name=SureFlag value="1">
			<input type=hidden name=JsFlag value="1">
			<input type=hidden name=AjaxFlag value="<%=AjaxFlag%>">
			<input type=hidden name=ID value="<%=LMT_AncID%>">
			<input type=hidden name=BoardID value="<%=GBL_Board_ID%>">
			<div class="value2">
			<%If GBL_UserID >= 1 and (GBL_BoardMasterFlag >= 5 and GetBinarybit(GBL_CHK_UserLimit,4) = 0) Then%>
			<b>��ѡ�������</b>
			<input type=radio class=fmchkbox name=DoingFlag value=0 checked>�Զ��Ű�
			<input type=radio class=fmchkbox name=DoingFlag value=1><%If Form_NotReplay = 0 Then%>��������<%Else%>�������<%End If%>
			<input type=radio class=fmchkbox name=DoingFlag value=2><%If Form_TitleStyle >= 60 Then%>ͨ�����<%Else%>���δ���<%End If%>
			<%Else%>
			<b>ȷ��Ҫ�Զ��Ű���Ϊ<font color=ff0000 class=redfont><%=LMT_AncID%></font>������������</b>
			<%End If
			If GBL_BoardMasterFlag >= 7 Then%>
			<input type=radio class=fmchkbox name=DoingFlag value=3><%If Form_TitleStyle = 30 Then%>�������<%Else%>��������Դ��û�<%End If
			End If%>
			</div>
			<p><input type=submit value=ȷ�� class="fmbtn btn_2">
		</form>
		<%Processor_Bottom
	End If

End Sub

Function ResumeCode(Tstr)

	Dim str
	str = Tstr
	Str = Replace(str," &nbsp; &nbsp; &nbsp;",chr(9))
	Str = Replace(str,"<br>" & "&nbsp;",VbCrLf & " ")
	Str = Replace(str,"<br>" & "&nbsp;",VbCrLf & " ")
	Str = Replace(str,"<br>" & VbCrLf,VbCrLf)
	Str = Replace(str,"<br>" & VbCrLf,VbCrLf)
	Str = Replace(str,"<br>",VbCrLf)
	Str = Replace(str,"<br>",VbCrLf)
	Str = Replace(str,"&nbsp;"," ")
	str = Replace(str,"&gt;",">")
	Str = Replace(str,"&lt;","<")
	Str = Replace(str,"&quot;","""")
	ResumeCode = Str

End Function

Function ReMakeIDDoc(ID)

	Dim Rs,htmlflag,Content
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	Set Rs = LDExeCute(sql_select("Select Content,htmlflag from LeadBBS_Announce where ID=" & ID,1),0)
	If Rs.Eof Then
		ReMakeIDDoc = 0
		Rs.Close
		Set Rs = Nothing
		GBL_CHK_TempStr = GBL_CHK_TempStr & "�Ҳ��������ӣ�<br>" & VbCrLf
		Exit Function
	Else
		htmlflag = Rs("htmlflag")
		If htmlflag <> 0 and htmlflag <> 2 and htmlflag <> 3 then
			ReMakeIDDoc = 0
			Rs.Close
			Set Rs = Nothing
			GBL_CHK_TempStr = GBL_CHK_TempStr & "����������Ͳ��Ǵ��ı���UBB��ʽ�����ܽ����Զ��Ű棡<br>" & VbCrLf
			Exit Function
		End if
		
		Dim NewTemp,N,I,moredataArray,tm,Tmp,fullstop,tm2,splitflag
		Content = Rs("Content")		
		Rs.Close
		Set Rs = Nothing
		
		If htmlflag = 2 and inStr(Content,"[CODE]") > 0 Then
			GBL_CHK_TempStr = "������д����ǩ[CODE]���Զ��Ű���ȡ����<br>" & VbCrLf
			Exit Function
		End If
		
		NewTemp = Content
		fullstop = "|?|��|.|��|��|;|��|!|:|��|��|`|`��|��|��|""|��|"
		If isNull(NewTemp) or NewTemp = "" Then
		Else
			If htmlflag = 2 Then NewTemp = Replace(NewTemp,"[p]","[P]")
			moredataArray = split(NewTemp,VbCrLf)
			I = Ubound(moredataArray,1)
			NewTemp = ""
			moredataArray(0) = Trim(moredataArray(0))
			If Left(moredataArray(0),1) <> "��" and (htmlflag <> 2 or Left(moredataArray(0),3) <> "[P]") Then
				If Left(moredataArray(0),2) <> "����" Then
					moredataArray(0) = "����" & moredataArray(0)
				Else
					moredataArray(0) = "��" & moredataArray(0)
				End If
			End If
			NewTemp = moredataArray(0)
			splitflag = 0
			For N = 0 to I-1
				do While right(moredataArray(N),1) = "��" or right(moredataArray(N),1) = " "
					moredataArray(N) = left(moredataArray(N),len(moredataArray(N))-1)
				loop
				tm = clearUbbcode(Trim(Replace(moredataArray(N),chr(9),"      ")))
				Tmp = right(tm,1)
				If inStr(fullstop,"|" & Tmp & "|") or splitflag = 1 Then
					If inStr(fullstop,"|" & Tmp & "|") Then splitflag = 1
					tm2 = clearUbbcode(Trim(Replace(moredataArray(N+1),chr(9),"      ")))
					do While right(tm2,1) = "��" or right(tm2,1) = " "
						tm2 = left(tm2,len(tm2)-1)
					loop
					if tm2 <> "" Then
						splitflag = 0
						moredataArray(N+1) = Trim(moredataArray(N+1))
						If Left(moredataArray(N+1),1) <> "��" and (htmlflag <> 2 or Left(moredataArray(N+1),3) <> "[P]") Then
							If Left(moredataArray(N+1),2) <> "����" Then
								moredataArray(N+1) = "����" & moredataArray(N+1)
							Else
								moredataArray(N+1) = "��" & moredataArray(N+1)
							End If
						End If
						If inStr(fullstop,"|" & Tmp & "|") Then NewTemp = NewTemp & VbCrLf
						If tm="" or isNull(tm) Then
							NewTemp = NewTemp & VbCrLf & moredataArray(N+1)
						Else
							NewTemp = Rtrim(NewTemp) & VbCrLf & moredataArray(N+1)
						End If
					Else
						NewTemp = NewTemp & VbCrLf & moredataArray(N+1)
					End If
				Else
					tm = Left(moredataArray(N+1),1)
					If tm <> " " and tm <> "��" and tm <> chr(9) and tm <> "" and len(moredataArray(N)) > 25 Then
						NewTemp = NewTemp & moredataArray(N+1)
					Else
						moredataArray(N+1) = Trim(moredataArray(N+1))
						If Left(moredataArray(N+1),1) <> "��" and (htmlflag <> 2 or Left(moredataArray(N+1),3) <> "[P]") Then
							If Left(moredataArray(N+1),2) <> "����" Then
								moredataArray(N+1) = "����" & moredataArray(N+1)
							Else
								moredataArray(N+1) = "��" & moredataArray(N+1)
							End If
						End If
						NewTemp = NewTemp & VbCrLf & moredataArray(N+1)
					End If
				End If
			Next
			
			'split [p]
			If htmlflag = 2 Then
				moredataArray = split(Replace(NewTemp,"[p]","[P]"),"[P]")
				I = Ubound(moredataArray,1)
				NewTemp = ""
				Dim addflag
				For N = 0 to I
					tm = clearUbbcode(Replace(moredataArray(N),chr(9),"      "))
					Tmp = left(tm,2)
					
					addflag = 0
					If Replace(Replace(tm,"��","")," ","") <> "" Then
						If Tmp <> "����" and Tmp <> "  " and Tmp <> "�� " and Tmp <> " ��" Then
							If N = 0 Then
								NewTemp = NewTemp & "����" & moredataArray(N)
							Else
								NewTemp = NewTemp & "[P]����" & moredataArray(N)
							End If
							addflag = 1
						End If
					End If
					
					If addflag = 0 Then
						If N = 0 Then
							NewTemp = NewTemp & moredataArray(N)
						Else
							NewTemp = NewTemp & "[P]" & moredataArray(N)
						End If
					End If
				Next
			End If

			If htmlflag = 2 Then
				'���ֲ���
			ElseIf htmlflag = 1 and CheckSupervisorUserName = 1 and GBL_UserID > 0 Then
				'���ֲ���
			Else
				'���ֲ���
			End If
			If NewTemp <> Content Then
				CALL LDExeCute("Update LeadBBS_Announce Set Content='" & Replace(NewTemp,"'","''") & "',htmlflag=" & htmlflag & " where ID=" & ID,1)
			End If
		End If
		If CheckSupervisorUserName = 0 Then
			CALL LDExeCute("Update LeadBBS_User Set LastWriteTime=" & GetTimeValue(DEF_Now) & " where ID=" & GBL_UserID,1)
			UpdateSessionValue 13,GetTimeValue(DEF_Now),0
		End If
		ReMakeIDDoc = 1
	End if

End Function

sub UpdateBoardLastAnnounce

	Dim Rs,SQL
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	Dim LastAnnounceID
	LastAnnounceID = cCur(Application(DEF_MasterCookies & "BoardInfo" & GBL_Board_ID)(19,0))

	If LastAnnounceID = LMT_AncID or LastAnnounceID = Form_RootIDBAK Then
		CALL LDExeCute("Update LeadBBS_Boards Set LastTopicName='" & Replace(Form_Title,"'","''") & "' where BoardID=" & GBL_Board_ID,1)
		UpdateBoardApplicationInfo GBL_Board_ID,Form_Title,20
	End If

End sub

Sub UpdateAnnounceApplicationInfo(AncID,IndexN,Value,tp,tid)

	Dim GetDataTop,AllTopNum,N,Str
	If tid = 0 Then
		Str = ""
	Else
		Str = tid
	End if
	AllTopNum = -1
	GetDataTop = Application(DEF_MasterCookies & "TopAnc" & Str)
	If isArray(GetDataTop) = False Then
		'If GetDataTop <> "yes" Then ReloadTopAnnounceInfo(tid)
		Exit Sub
	Else
		AllTopNum = Ubound(GetDataTop,2)
	End If

	For N = 0 to AllTopNum
		If cCur(AncID) = cCur(GetDataTop(0,N)) Then
			If tp = 1 Then
				GetDataTop(IndexN,N) = cCur(GetDataTop(IndexN,N)) + Value
			Else
				GetDataTop(IndexN,N) = Value
			End If
			Application.Lock
			Application(DEF_MasterCookies & "TopAnc" & Str) = GetDataTop
			Application.UnLock
			Exit Sub
		End If
	Next

End Sub

Sub UpdateSpecialUserTable2(UserLimit,UserID,UserName,N,assort)

	Dim Rs
	Dim Flag
	
	Rem ��֤�û�
	Flag = GetBinarybit(UserLimit,N)
	If Flag = 0 Then
		CALL LDExeCute("Delete from LeadBBS_SpecialUser where Assort=" & assort & " and UserID=" & UserID,1)
	Else
		Set Rs = LDExeCute(sql_select("Select ID from LeadBBS_SpecialUser Where Assort=" & assort & " and UserID=" & UserID,1),0)
		If Rs.Eof Then
			Rs.Close
			Set Rs = Nothing
			CALL LDExeCute("insert into LeadBBS_SpecialUser(UserID,UserName,BoardID,Assort,ndatetime) values(" & UserID & ",'" & Replace(UserName,"'","''") & "',0," & assort & "," & GetTimeValue(DEF_Now) & ")",1)
		Else
			Rs.Close
			Set Rs = Nothing
		End If
	End If

End Sub
%>