<%
Const ViewSelectListFlag = 1'0 �Ʊ����ʾ 1.�ո����������¼���+�Ŵ���
Dim GBL_AssortID,GBL_AssortName
Dim GBL_BoardID,GBL_BoardAssort,GBL_BoardName,GBL_BoardIntro,GBL_LastWriter,GBL_LastWriteTime,GBL_TopicNum
Dim GBL_AnnounceNum,GBL_BoardManage,GBL_ForumPass,GBL_HiddenFlag,GBL_BoardLimit,GBL_MasterList,GBL_OrderID,GBL_OrderID_Old
Dim GBL_MasterList_Old,GBL_BoardID_Old,GBL_BoardStyle,GBL_StartTime,GBL_EndTime
Dim GBL_BoardImgUrl,GBL_BoardImgWidth,GBL_BoardImgHeight,GBL_BoardHead,GBL_BoardBottom
Dim GBL_ParentBoard,GBL_LowerBoard,GBL_OtherLimit,GBL_OtherLimit_Part1,GBL_OtherLimit_Part2
Dim GBL_BoardAssort_Old,GBL_ParentBoard_Old
Dim GBL_LowerBoardTemp

GBL_OrderID_Old = 0
GBL_ParentBoard = 0
GBL_LowerBoard = ""

Dim GBL_GetData
GBL_BoardID_Old = 0
GBL_BoardImgUrl = ""
GBL_BoardImgWidth = 0
GBL_BoardImgHeight = 0

Dim GBL_LimitHourStart,GBL_LimitHourEnd
GBL_LimitHourStart = 0
GBL_LimitHourEnd = 0

Dim GBL_LimitWeekEnd,GBL_LimitWeekStart
GBL_LimitWeekEnd = 0
GBL_LimitWeekStart = 0

Dim GBL_LimitMonthStart,GBL_LimitMonthEnd
GBL_LimitMonthStart = 0
GBL_LimitMonthEnd = 0

Dim GBL_HiddenFlagData,GBL_HiddenFlagNum
GBL_HiddenFlagData = Array("������ʾ","��̳�б�������","�ر���̳")
GBL_HiddenFlagNum = Ubound(GBL_HiddenFlagData,1)

GBL_LastWriteTime = GetTimeValue(DEF_Now)
GBL_TopicNum = 0
GBL_AnnounceNum = 0

Rem ������֤
Function CheckFormForumBoardData

	Dim GBL_MasterListArray,GBL_MasterList_OldD
	GBL_MasterListArray = Split(GBL_MasterList,",")
	GBL_MasterList_OldD = GBL_MasterList

	If isNumeric(GBL_BoardID) = False Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ����Ϊ��̳����IDָ��һ������0�����֣��������������ַ���<br>" & VbCrLf
		CheckFormForumBoardData = 0
		Exit Function
	End If

	GBL_BoardID = cCur(GBL_BoardID)
	If GBL_BoardID > 2147479999 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ��̳����ID��д̫��<br>" & VbCrLf
		CheckFormForumBoardData = 0
		Exit Function
	End If

	If GBL_BoardID < 1 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ����Ϊ��̳����IDָ��һ������0�����֡�<br>" & VbCrLf
		CheckFormForumBoardData = 0
		Exit Function
	End If

	If isNumeric(GBL_OrderID) = False Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ����Ϊ��̳��������˳��ָ��һ������0�����֣��������������ַ���<br>" & VbCrLf
		CheckFormForumBoardData = 0
		Exit Function
	End If

	GBL_OrderID = cCur(GBL_OrderID)
	If GBL_OrderID < 0 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ����Ϊ��̳����˳��ָ��һ������0�����֡�<br>" & VbCrLf
		CheckFormForumBoardData = 0
		Exit Function
	End If
	

	If isNumeric(GBL_BoardAssort) = 0 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ����Ϊ��̳����IDָ��һ������0�����֣��������������ַ���<br>" & VbCrLf
		CheckFormForumBoardData = 0
		Exit Function
	End If

	GBL_BoardAssort = cCur(GBL_BoardAssort)
	If GBL_BoardAssort < 1 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ����Ϊ��̳����IDָ��һ������0�����֡�<br>" & VbCrLf
		CheckFormForumBoardData = 0
		Exit Function
	End If

	If len(GBL_BoardName)<1 or GBL_BoardName = "" Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ��̳���������Ǳ�����<br>" & VbCrLf
		CheckFormForumBoardData = 0
		Exit Function
	End If

	If strLength(GBL_BoardName) > 250 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ��̳�������Ƴ��Ȳ��ܳ���250���ַ�<br>" & VbCrLf
		CheckFormForumBoardData = 0
		Exit Function
	End If
	
	If inStr(LCase(GBL_BoardName),"""") > 0 or inStr(LCase(GBL_BoardName),"<script") > 0 or inStr(LCase(GBL_BoardName),"<\script") > 0 or inStr(LCase(GBL_BoardName),"</script") > 0 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ��̳�������Ʋ��������js���������룬������ʹ��˫����<br>" & VbCrLf
		CheckFormForumBoardData = 0
		Exit Function
	End If
	
	Dim FobWords,TempN,TempURL,Temp2
	FobWords = Array(91,92,304,305,430,431,437,438,12460,12461,12462,12463,12464,12465,12466,12467,12468,12469,12470,12471,12472,12473,12474,12475,12476,12477,12478,12479,12480,12481,12482,12483,12485,12486,12487,12488,12489,12490,12496,12497,12498,12499,12500,12501,12502,12503,12504,12505,12506,12507,12508,12509,12510,12532,12533,65339,65340)
	Temp2 = Ubound(FobWords,1)
	For TempN = 1 to Temp2
		If inStr(GBL_BoardName,ChrW(FobWords(TempN))) > 0 Then
			'GBL_CHK_TempStr = GBL_CHK_TempStr & "��̳�����е��ַ�<u>" & ChrW(FobWords(TempN)) & "</u>���ڷǷ��ַ�!<br>"
			'GBL_CHK_Flag = 0
			'Exit Function
		End If
	Next

	If strLength(GBL_BoardIntro) > 500 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ��̳������̳������������Ȳ��ܳ���500���ַ�<br>" & VbCrLf
		CheckFormForumBoardData = 0
		Exit Function
	End If
	
	If inStr(GBL_BoardIntro,"<script") > 0 or inStr(GBL_BoardIntro,"<\script") > 0 or inStr(GBL_BoardIntro,"</script") > 0 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: �����鲻�������js����������<br>" & VbCrLf
		CheckFormForumBoardData = 0
		Exit Function
	End If

	If strLength(GBL_BoardIntro) > 500 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ��̳������̳������������Ȳ��ܳ���500���ַ�<br>" & VbCrLf
		CheckFormForumBoardData = 0
		Exit Function
	End If
		

	If Len(GBL_LastWriter) > 20 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ��̳������󷢱��������߳��Ȳ��ܳ���20����<br>" & VbCrLf
		CheckFormForumBoardData = 0
		Exit Function
	End If

	If isDate(RestoreTime(GBL_LastWriteTime)) = False and GBL_LastWriteTime <> 0 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ��̳������󷢱����ӵ�ʱ�����������ڸ�ʽ��<br>" & VbCrLf
		CheckFormForumBoardData = 0
		Exit Function
	End If

	If isNumeric(GBL_TopicNum) = 0 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ����ӵ�е�����������������һ�����֣��������������ַ���<br>" & VbCrLf
		CheckFormForumBoardData = 0
		Exit Function
	End If
	GBL_TopicNum = cCur(GBL_TopicNum)

	If isNumeric(GBL_AnnounceNum) = 0 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ����ӵ�е�������������һ�����֣��������������ַ���<br>" & VbCrLf
		CheckFormForumBoardData = 0
		Exit Function
	End If
	GBL_AnnounceNum = cCur(GBL_AnnounceNum)
	
	If Len(GBL_ForumPass) > 20 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ��̳�������벻�ܳ���20λ��<br>" & VbCrLf
		CheckFormForumBoardData = 0
		Exit Function
	End If

	If StrLength(GBL_BoardImgUrl) > 255 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ��̳ͼƬurl̫�������ܳ���255�ֽڡ�<br>" & VbCrLf
		CheckFormForumBoardData = 0
		Exit Function
	End If
	
	If IsNumeric(GBL_BoardImgWidth) = 0 or isNull(GBL_BoardImgWidth) Then GBL_BoardImgWidth = 0
	GBL_BoardImgWidth = Fix(cCur(GBL_BoardImgWidth))
	If GBL_BoardImgWidth < 0 Then GBL_BoardImgWidth = 0
	If GBL_BoardImgWidth > 200 Then GBL_BoardImgWidth = 200
	
	If IsNumeric(GBL_BoardImgHeight) = 0 or isNull(GBL_BoardImgHeight) Then GBL_BoardImgHeight = 0
	GBL_BoardImgHeight = Fix(cCur(GBL_BoardImgHeight))
	If GBL_BoardImgHeight < 0 Then GBL_BoardImgHeight = 0
	If GBL_BoardImgHeight > 200 Then GBL_BoardImgHeight = 200

	Dim Temp1
	GBL_BoardLimit = 0
	Temp2 = 1
	For TempN = 0 to LimitBoardStringDataNum
		Temp1 = Request("Limit" & TempN+1)
		If Temp1 <> "1" Then Temp1 = "0"
		If Temp1 = "1" Then GBL_BoardLimit = GBL_BoardLimit+cCur(Temp2)
		Temp2 = Temp2*2
	Next

	If isNumeric(GBL_HiddenFlag) = 0 or GBL_HiddenFlag = "" or inStr(GBL_HiddenFlag,",") > 0 then GBL_HiddenFlag = 0
	GBL_HiddenFlag = cCur(GBL_HiddenFlag)
	If GBL_HiddenFlag < 0 or GBL_HiddenFlag > GBL_HiddenFlagNum Then GBL_HiddenFlag = 0
	
	
	If isNumeric(GBL_BoardStyle) = 0 or GBL_BoardStyle = "" or inStr(GBL_BoardStyle,",") > 0 then GBL_BoardStyle = 0
	GBL_BoardStyle = cCur(GBL_BoardStyle)
	If GBL_BoardStyle < 0 or GBL_BoardStyle > DEF_BoardStyleStringNum Then GBL_BoardStyle = 0
	
	If isNumeric(GBL_LimitWeekStart) = 0 or GBL_LimitWeekStart = "" or inStr(GBL_LimitWeekStart,",") > 0 then GBL_LimitWeekStart = 0
	GBL_LimitWeekStart = cCur(GBL_LimitWeekStart)
	If GBL_LimitWeekStart < 0 or GBL_LimitWeekStart > 7 Then GBL_LimitWeekStart = 0	
	If isNumeric(GBL_LimitWeekEnd) = 0 or GBL_LimitWeekEnd = "" or inStr(GBL_LimitWeekEnd,",") > 0 then GBL_LimitWeekEnd = 0
	GBL_LimitWeekEnd = cCur(GBL_LimitWeekEnd)
	If GBL_LimitWeekEnd < 0 or GBL_LimitWeekEnd > 7 Then GBL_LimitWeekEnd = 0
	'If GBL_LimitWeekEnd < GBL_LimitWeekStart Then
	'	GBL_LimitWeekEnd = 0
	'	GBL_LimitWeekStart = 0
	'End If
		
	If isNumeric(GBL_LimitMonthStart) = 0 or GBL_LimitMonthStart = "" or inStr(GBL_LimitMonthStart,",") > 0 then GBL_LimitMonthStart = 0
	GBL_LimitMonthStart = cCur(GBL_LimitMonthStart)
	If GBL_LimitMonthStart < 0 or GBL_LimitMonthStart > 31 Then GBL_LimitMonthStart = 0	
	If isNumeric(GBL_LimitMonthEnd) = 0 or GBL_LimitMonthEnd = "" or inStr(GBL_LimitMonthEnd,",") > 0 then GBL_LimitMonthEnd = 0
	GBL_LimitMonthEnd = cCur(GBL_LimitMonthEnd)
	If GBL_LimitMonthEnd < 0 or GBL_LimitMonthEnd > 31 Then GBL_LimitMonthEnd = 0
	'If GBL_LimitMonthEnd < GBL_LimitMonthStart Then
	'	GBL_LimitMonthEnd = 0
	'	GBL_LimitMonthStart = 0
	'End If
		
	If isNumeric(GBL_LimitHourStart) = 0 or GBL_LimitHourStart = "" or inStr(GBL_LimitHourStart,",") > 0 then GBL_LimitHourStart = 0
	GBL_LimitHourStart = cCur(GBL_LimitHourStart)
	If GBL_LimitHourStart < 0 or GBL_LimitHourStart > 23 Then GBL_LimitHourStart = 0	
	If isNumeric(GBL_LimitHourEnd) = 0 or GBL_LimitHourEnd = "" or inStr(GBL_LimitHourEnd,",") > 0 then GBL_LimitHourEnd = 0
	GBL_LimitHourEnd = cCur(GBL_LimitHourEnd)
	If GBL_LimitHourEnd < 0 or GBL_LimitHourEnd > 23 Then GBL_LimitHourEnd = 0
	'If GBL_LimitHourEnd < GBL_LimitHourStart Then
	'	GBL_LimitHourEnd = 0
	'	GBL_LimitHourStart = 0
	'End If

	If isNumeric(GBL_ParentBoard) = 0 Then
		GBL_CHK_TempStr = "�����ϼ�����ָ���������ϼ���������д����0"
		CheckFormForumBoardData = 0
		Exit Function
	End If

	GBL_ParentBoard = Fix(cCur(GBL_ParentBoard))

	Dim TempName

	If Ubound(GBL_MasterListArray,1) = 0 and GBL_MasterList = "?LeadBBS?" Then

	Else
		GBL_MasterList = ""
		If Ubound(GBL_MasterListArray,1) > DEF_MaxBoardMastNum - 1 Then
			GBL_CHK_TempStr = "���󣬰������ֻ������" & DEF_MaxBoardMastNum & "��"
			CheckFormForumBoardData = 0
			GBL_MasterList = GBL_MasterList_OldD
			Exit Function
		End if
	
		For TempN = 0 to Ubound(GBL_MasterListArray,1)
			If Trim(GBL_MasterListArray(TempN)) <> "" Then
				TempName = CheckUserNameExist(GBL_MasterListArray(TempN))
				If TempName = "" Then
					GBL_CHK_TempStr = "Error: " & DEF_PointsName(8) & "�б�����û�" & htmlencode(GBL_MasterListArray(TempN)) & "�����ڣ�<br>" & VbCrLf
					CheckFormForumBoardData = 0
					GBL_MasterList = GBL_MasterList_OldD
					Exit Function
				Else
					GBL_MasterList = GBL_MasterList & "," & TempName
				End If
			End If
		Next
		If Left(GBL_MasterList,1) = "," Then GBL_MasterList = Mid(GBL_MasterList,2)
	End If

	If isNumeric(GBL_OtherLimit_Part1) = 0 Then GBL_OtherLimit_Part1 = 0
	GBL_OtherLimit_Part1 = Fix(cCur(GBL_OtherLimit_Part1))

	If isNumeric(GBL_OtherLimit_Part2) = 0 Then GBL_OtherLimit_Part2 = 0
	GBL_OtherLimit_Part2 = Fix(cCur(GBL_OtherLimit_Part2))
	If GBL_OtherLimit_Part2 < 0 or GBL_OtherLimit_Part2 > 999999999999 Then 
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: �������������ֵ���󣬱�������㣡<br>" & VbCrLf
		CheckFormForumBoardData = 0
		Exit Function
	End If
	
	If GBL_OtherLimit_Part1 = 5 Then
		GBL_OtherLimit_Part2 = Left(Request("GBL_UserOfficerString"),14)
		If isNumeric(GBL_OtherLimit_Part2) = 0 Then GBL_OtherLimit_Part2 = 0
		GBL_OtherLimit_Part2 = cCur(Fix(GBL_OtherLimit_Part2))
		If GBL_OtherLimit_Part2 < 0 or GBL_OtherLimit_Part2 > DEF_UserOfficerNum Then GBL_OtherLimit_Part2 = 0
	End If

	If GBL_OtherLimit_Part1 <= 0 or GBL_OtherLimit_Part1 > 5 Then
		GBL_OtherLimit_Part1 = 0
		GBL_OtherLimit_Part2 = 0
		GBL_OtherLimit = 0
	Else
		GBL_OtherLimit = cCur(GBL_OtherLimit_Part2 & Right("0" & GBL_OtherLimit_Part1,2))
	End If

	CheckFormForumBoardData = 1

	GBL_StartTime = Right("0" & GBL_LimitHourStart,2) & Right("0" & GBL_LimitWeekStart,2) & Right("0" & GBL_LimitMonthStart,2)
	GBL_EndTime = Right("0" & GBL_LimitHourEnd,2) & Right("0" & GBL_LimitWeekEnd,2) & Right("0" & GBL_LimitMonthEnd,2)

End Function

Rem ���ĳ�û����Ƿ����
Function CheckUserNameExist(UserName)

	Dim Rs
	Set Rs = LDExeCute(sql_select("Select UserName from LeadBBS_User where UserName='" & Replace(UserName,"'","''") & "'",1),0)
	If Rs.Eof Then
		CheckUserNameExist = ""
	Else
		CheckUserNameExist = Rs(0)
	End if
	Rs.Close
	Set Rs = Nothing

End Function

Rem ���ĳ����ID�Ƿ����
Function CheckForumAssortIDExist(AssortID)

	Dim Rs
	Set Rs = LDExeCute(sql_select("Select AssortID from LeadBBS_Assort where AssortID=" & AssortID,1),0)
	If Rs.Eof Then
		CheckForumAssortIDExist = 0
	Else
		CheckForumAssortIDExist = 1
	End if
	Rs.Close
	Set Rs = Nothing

End Function

Rem ���ĳ����ID�Ƿ����
Function CheckForumBoardIDExist(BoardID)

	Dim Rs
	Set Rs = LDExeCute(sql_select("Select BoardID from LeadBBS_Boards where BoardID=" & BoardID,1),0)
	If Rs.Eof Then
		CheckForumBoardIDExist = 0
	Else
		CheckForumBoardIDExist = 1
	End if
	Rs.Close
	Set Rs = Nothing

End Function

Rem ���ĳ���������Ƿ����
Function CheckForumBoardNameExist(BoardName)

	Dim Rs
	Set Rs = LDExeCute(sql_select("Select BoardID from LeadBBS_Boards where BoardName='" & Replace(BoardName,"'","''") & "'",1),0)
	If Rs.Eof Then
		CheckForumBoardNameExist = 0
	Else
		CheckForumBoardNameExist = cCur(rs(0))
	End if
	Rs.Close
	Set Rs = Nothing

End Function

Rem ɾ��ĳ����
Function DeleteForumBoard(BoardID)

	Dim Rs,ParentBoard,MasterList
	Set Rs = LDExeCute(sql_select("Select ParentBoard,LowerBoard,MasterList from LeadBBS_Boards where BoardID=" & BoardID,1),0)
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ��̳����ID��" & BoardID & "������!<br>" & VbCrLf
		DeleteForumBoard = 0
		Exit Function
	Else
		If Rs(1) & "" <> "" Then
			GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: �˰���ӵ���ӣ��¼�����̳������ɾ��!<br>" & VbCrLf
			Rs.Close
			Set Rs = Nothing
			Exit Function
		End if
		ParentBoard = cCur(Rs(0))
		MasterList = Rs(2)
		Rs.Close
		Set Rs = Nothing
		Set Rs = LDExeCute(sql_select("Select ID from LeadBBS_Announce where BoardID=" & BoardID,1),0)
		If Not Rs.Eof Then
			GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: �˰����»������Ӵ��ڣ��������ɾ������!<br>" & VbCrLf
			DeleteForumBoard = 0
			Rs.Close
			Set Rs = Nothing
			Exit Function
		End If
		Rs.Close
		Set Rs = Nothing
		CALL LDExeCute("delete from LeadBBS_GoodAssort where BoardID=" & BoardID,1)
		CALL LDExeCute("delete from LeadBBS_Boards where BoardID=" & BoardID,1)
		UpdateMasterList MasterList,0
		DeleteForumBoard = 1
	End if
	ReloadBoardInfo(BoardID)
	If ParentBoard > 0 Then
		UpdateParentBoard_LowerBoardColumn2(ParentBoard)
		ReloadBoardInfo(ParentBoard)
	End If
	ReloadBoardListData

	MakeBoardList "BoardJump.asp","b.asp"
	MakeBoardList "BoardJump2.asp","b2.asp"
	MakeBoardList_For_MoveAnnounce

End Function

Rem ����ĳ����
Function InsertForumBoard

	If CheckForumAssortIDExist(GBL_BoardAssort) = 0 Then
		InsertForumBoard = 0
		GBL_CHK_TempStr = GBL_CHK_TempStr & "�������ڵķ���ID��" & GBL_BoardAssort & "������!<br>" & VbCrLf
		Exit Function
	End If

	If CheckForumBoardIDExist(GBL_BoardID) = 1 Then
		InsertForumBoard = 0
		GBL_CHK_TempStr = GBL_CHK_TempStr & "����ID��" & GBL_BoardID & "�Ѿ�����!<br>" & VbCrLf
		Exit Function
	End If

	If CheckForumBoardNameExist(GBL_BoardName) = 1 Then
		InsertForumBoard = 0
		GBL_CHK_TempStr = GBL_CHK_TempStr & "��������" & htmlencode(GBL_BoardName) & "�Ѿ�����!<br>" & VbCrLf
		Exit Function
	End If

	CALL LDExeCute("insert into LeadBBS_Boards(BoardID,BoardAssort,BoardName,BoardIntro,LastWriter," &_
			"LastWriteTime,TopicNum,AnnounceNum,ForumPass,HiddenFlag,MasterList,BoardLimit,BoardStyle,StartTime,EndTime,BoardHead,BoardBottom,BoardImgUrl,BoardImgWidth,BoardImgHeight,ParentBoard,LowerBoard,ParentBoardStr,BoardLevel,OtherLimit) values(" &_
			GBL_BoardID & "," & Replace(GBL_BoardAssort,"'","''") & ",'" & Replace(GBL_BoardName,"'","''") & "','" & Replace(GBL_BoardIntro,"'","''") & "','" & Replace(GBL_LastWriter,"'","''") & "'," &_
			GBL_LastWriteTime & "," & GBL_TopicNum & "," & GBL_AnnounceNum & ",'" & Replace(GBL_ForumPass,"'","''") & "'" & _
			"," & GBL_HiddenFlag & ",'" & Replace(GBL_MasterList,"'","''") & "'," & GBL_BoardLimit & "," & GBL_BoardStyle & "," & GBL_StartTime & "," & GBL_EndTime & ",'','','',0,0," & GBL_ParentBoard & ",'" & Replace(GBL_LowerBoard,"'","''") & "'," & GBL_BoardID & ",1," & GBL_OtherLimit & ")",1)

	If GBL_MasterList <> "?LeadBBS?" Then UpdateMasterList GBL_MasterList,1

	ReloadBoardInfo(GBL_BoardID)
	ReloadBoardListData

	MakeBoardList "BoardJump.asp","b.asp"
	MakeBoardList "BoardJump2.asp","b2.asp"
	MakeBoardList_For_MoveAnnounce

	InsertForumBoard = 1

End Function

Rem �õ�ĳ������Ϣ
Function GetForumBoardData(BoardID)

	Dim Rs
	Set Rs = LDExeCute("Select * from LeadBBS_Boards Where BoardID = " & BoardID,0)
	If Rs.Eof Then
		GetForumBoardData = 0
		Rs.Close
		Set Rs = Nothing
		Exit Function
	Else
		GBL_GetData = Rs.GetRows(-1)
		Rs.Close
		Set Rs = Nothing
		GetForumBoardData = 1
		Exit Function
	End If

End Function

Rem ����ĳ����
Function UpdateForumBoard
	
	If isNumeric(GBL_MODIFYID) = 0 or GBL_MODIFYID = "" Then GBL_MODIFYID = 0
	GBL_MODIFYID = cCur(GBL_MODIFYID)
	If GBL_MODIFYID = 0 or GBL_MODIFYID<1 then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: Ҫ�޸ĵİ��治���ڣ�<br>" & VbCrLf
		GBL_CHK_Flag = 0
		UpdateForumBoard = 0
		Exit Function
	End If

	If GetForumBoardData(GBL_MODIFYID) = 0 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ����ID��" & GBL_BoardID & "�������޷�����޸ġ�<br>" & VbCrLf
		GBL_CHK_Flag = 0
		UpdateForumBoard = 0
		Exit Function
	End If

	If cCur(GBL_GetData(0,0))<>GBL_BoardID and CheckForumBoardIDExist(GBL_BoardID) = 1 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ����ID��" & GBL_BoardID & "�Ѿ����ڣ���ʹ������ID�š�<br>" & VbCrLf
		GBL_CHK_Flag = 0
		UpdateForumBoard = 0
		Exit Function
	End If
	
	If GBL_ParentBoard > 0 Then
		If CheckForumBoardIDExist(GBL_ParentBoard) = 0 Then
			GBL_CHK_TempStr = "�ϼ�������" & GBL_ParentBoard & "�����ڣ�����ȷ��д��<br>" & VbCrLf
			GBL_CHK_Flag = 0
			UpdateForumBoard = 0
			Exit Function
		End If
	End If
			
	If CheckBoardRelation(GBL_ParentBoard,GBL_BoardID) = 0 Then
		UpdateForumBoard = 0
		Exit Function
	End if

	If UpdateParentBoard_LowerBoardColumn(GBL_ParentBoard,GBL_BoardID) = 0 Then
		UpdateForumBoard = 0
		Exit Function
	End if
	
	Dim Temp
	Temp = CheckForumBoardNameExist(GBL_BoardName)
	'If Temp<>0 and Temp<>cCur(GBL_GetData(0,0)) Then
	'	GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: �Ѿ���������Ϊ<b>" & htmlencode(GBL_BoardName) & "</b>�İ���<br>" & VbCrLf
	'	GBL_CHK_Flag = 0
	'	UpdateForumBoard = 0
	'	Exit Function
	'End If

	If GBL_MasterList_Old <> "?LeadBBS?" and GBL_MasterList_Old <> GBL_MasterList Then UpdateMasterList GBL_MasterList_Old,0

	CALL LDExeCute("Update LeadBBS_Boards Set BoardAssort=" & GBL_BoardAssort & ",BoardName='" & Replace(GBL_BoardName,"'","''") & "',BoardIntro='" & Replace(GBL_BoardIntro,"'","''") & "'" &_
	",LastWriter='" & Replace(GBL_LastWriter,"'","''") & "',LastWriteTime=" & GBL_LastWriteTime & ",TopicNum=" & GBL_TopicNum & ",AnnounceNum=" & GBL_AnnounceNum & ",ForumPass='" & Replace(GBL_ForumPass,"'","''") & "',HiddenFlag=" & GBL_HiddenFlag & _
	",MasterList='" & Replace(GBL_MasterList,"'","''") & "'" & _
	",BoardLimit=" & GBL_BoardLimit & _
	",OrderID=" & GBL_OrderID & _
	",BoardStyle=" & GBL_BoardStyle & _
	",StartTime=" & GBL_StartTime & _
	",EndTime=" & GBL_EndTime & _
	",BoardHead='" & Replace(GBL_BoardHead,"'","''") & "'" &_
	",BoardBottom='" & Replace(GBL_BoardBottom,"'","''") & "'" &_
	",BoardImgUrl='" & Replace(GBL_BoardImgUrl,"'","''") & "'" &_
	",BoardImgWidth=" & GBL_BoardImgWidth & _
	",BoardImgHeight=" & GBL_BoardImgHeight & _
	",ParentBoard=" & GBL_ParentBoard & _
	",OtherLimit=" & GBL_OtherLimit & _
	" where BoardID=" & GBL_GetData(0,0),1)

	If GBL_ParentBoard <> GBL_ParentBoard_Old or GBL_OrderID_Old <> GBL_OrderID Then
		UpdateParentBoard_LowerBoardColumn2(GBL_ParentBoard_Old)
		UpdateParentBoard_LowerBoardColumn2(GBL_ParentBoard)
		UpdateParentBoardStrColumn GBL_ParentBoard_Old,GBL_ParentBoard,cCur(GBL_GetData(0,0))
	End If

	
	If GBL_MasterList <> "?LeadBBS?" and GBL_MasterList_Old <> GBL_MasterList Then UpdateMasterList GBL_MasterList,1

	ReloadBoardInfo(GBL_GetData(0,0))
	ReloadBoardListData

	MakeBoardList "BoardJump.asp","b.asp"
	MakeBoardList "BoardJump2.asp","b2.asp"
	MakeBoardList_For_MoveAnnounce

	UpdateForumBoard = 1

End Function

Function UpdateMasterList(MasterList,Flag)

	Rem ���¸�����̳�û�����״̬
	Dim TA,N

	TA = Split(MasterList,",")
	For N = 0 to Ubound(TA,1)
		If TA(N) <> "" Then SetUserMastFlag TA(N),Flag
	Next

End Function

Rem ����ĳ�û��Ƿ����
Function SetUserMastFlag(UserName,Fla)

	Dim Flag
	Flag = Fla
	If Flag <> 1 and Flag <> 0 Then Flag = 0
	Fla = Flag
	Dim Rs,Temp,SQL
	If Flag = 0 Then
		SQL = sql_select("Select BoardID from LeadBBS_Boards where BoardID<>" & GBL_BoardID_Old & " and (MasterList='" & Replace(UserName,"'","''") & "' or MasterList like'" & Replace(UserName,"'","''") & ",%' or MasterList like'%," & Replace(UserName,"'","''") & "' or MasterList like'%," & Replace(UserName,"'","''") & ",%')",1)
		Set Rs = LDExeCute(SQL,0)
		If Rs.Eof Then
			Flag = 0
		Else
			Flag = 1
		End If
		Rs.Close
		Set Rs = Nothing
	End if

	Dim Tmp
	Set Rs = LDExeCute(sql_select("Select UserLimit,ID from LeadBBS_User where UserName='" & Replace(UserName,"'","''") & "'",1),0)
	If Not Rs.Eof Then
		Temp = Rs(0)
		Tmp = Rs(1)
		If isNull(Temp) Then Temp = 0
		Temp = SetBinarybit(Temp,8,Flag)
		SetUserMastFlag = 1
		Rs.Close
		Set Rs = Nothing
		CALL LDExeCute("Update LeadBBS_User Set UserLimit=" & Temp & " where UserName='" & Replace(UserName,"'","''") & "'",1)
		If Fla = 0 Then
			CALL LDExeCute("Delete from LeadBBS_SpecialUser where Assort=1 and UserID=" & Tmp & " and BoardID=" & GBL_BoardID,1)
		Else
			Set Rs = LDExeCute(sql_select("Select ID from LeadBBS_SpecialUser Where Assort=1 and UserID=" & Tmp & " and BoardID=" & GBL_BoardID,1),0)
			If Rs.Eof Then
				Rs.Close
				Set Rs = Nothing
				CALL LDExeCute("insert into LeadBBS_SpecialUser(UserID,UserName,BoardID,Assort,ndatetime) values(" & Tmp & ",'" & Replace(UserName,"'","''") & "'," & GBL_BoardID & ",1," & GetTimeValue(DEF_Now) & ")",1)
			Else
				Rs.Close
				Set Rs = Nothing
			End If
		End If
	Else
		SetUserMastFlag = 0
		Rs.Close
		Set Rs = Nothing
	End if

End Function

Rem �жϸ�����̳�뵱ǰ��̳�Ĺ�ϵ
Function CheckBoardRelation(ParentBoard,BoardID)

	If ParentBoard = 0 Then
		CheckBoardRelation = 1
		Exit Function
	End If

	Dim Rs,BoardAssort
	Set Rs = LDExeCute(sql_select("Select BoardAssort from LeadBBS_Boards where BoardID=" & ParentBoard,1),0)
	If Rs.Eof Then
		CheckBoardRelation = 0
		Rs.close
		Set Rs = Nothing
		GBL_CHK_TempStr = GBL_CHK_TempStr & "��������ID��" & ParentBoard & "�����ڣ�����ȷ��д��<br>" & VbCrLf
		Exit Function
	Else
		BoardAssort = cCur(Rs(0))
	End if
	Rs.Close
	Set Rs = Nothing
	
	If BoardAssort <> GBL_BoardAssort Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "��ǰ�������ڷ�������븸�����汣��һ�£�����ȷ��д��<br>" & VbCrLf
		CheckBoardRelation = 0
		Exit Function
	End If

	If GBL_LowerBoard & "" <> "" and GBL_BoardAssort_Old <> GBL_BoardAssort Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "�˰�������¼����棬���Խ�ֹ�޸��������ࡣ<br>" & VbCrLf
		CheckBoardRelation = 0
		Exit Function
	End If
	
	Dim ParentBoardTemp
	ParentBoardTemp = ParentBoard
	Dim N
	For N = 1 to 20
		If ParentBoardTemp = 0 then Exit for
		If ParentBoardTemp = GBL_BoardID Then
			GBL_CHK_TempStr = GBL_CHK_TempStr & "������ָ̳�����󣬵�ǰ�����Ѿ��������ϼ���̳���ϼ���ǰ���档<br>" & VbCrLf
			CheckBoardRelation = 0
			Exit Function
		End If
		Set Rs = LDExeCute(sql_select("Select ParentBoard from LeadBBS_Boards where BoardID=" & ParentBoardTemp,1),0)
		If Rs.Eof Then
			Rs.Close
			Set Rs = Nothing
			Exit For
		Else
			ParentBoardTemp = cCur(Rs(0))
			Rs.Close
			Set Rs = Nothing
		End If
	Next
	Set Rs = Nothing
	CheckBoardRelation = 1

End Function

Rem ����Ƿ񻹿����޸�Ϊ����������Ӱ���
Function UpdateParentBoard_LowerBoardColumn(ParentBoard,BoardID)

	If ParentBoard = GBL_ParentBoard_Old Then
		UpdateParentBoard_LowerBoardColumn = 1
		Exit Function
	End If
	Dim Rs,ParentBoardStr
	Set Rs = LDExeCute(sql_select("Select LowerBoard,ParentBoardStr from LeadBBS_Boards where BoardID=" & BoardID,1),0)
	If Rs.Eof Then
		UpdateParentBoard_LowerBoardColumn = 0
		Rs.close
		Set Rs = Nothing
		GBL_CHK_TempStr = GBL_CHK_TempStr & "��������ID��" & ParentBoard & "�����ڣ�����ȷ��д��<br>" & VbCrLf
		Exit Function
	Else
		GBL_LowerBoardTemp = Rs(0)
		ParentBoardStr = Rs(1)
	End if
	Rs.Close
	Set Rs = Nothing
	Rem ������Ļ����Ŀ�����Ĳ�������
	If Len(ParentBoardStr & "," & BoardID) > 55 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "��̳�������������������ĸ������棬�޸�ʧ�ܡ�<br>" & VbCrLf
		UpdateParentBoard_LowerBoardColumn = 0
		Exit Function
	End If
	
	If inStr("," & GBL_LowerBoardTemp & ",","," & BoardID & ",") Then
	Else
		If Len(GBL_LowerBoardTemp & "," & BoardID) > 255 Then
			GBL_CHK_TempStr = GBL_CHK_TempStr & "ѡ��ĸ������Ѿ��ﵽ������������̳��Ŀ���޸�ʧ�ܡ�<br>" & VbCrLf
			UpdateParentBoard_LowerBoardColumn = 0
			Exit Function
		End If
	End If
	UpdateParentBoard_LowerBoardColumn = 1

End Function

Rem ���¸���������Ӱ�������
Function UpdateParentBoard_LowerBoardColumn2(ParentBoard)

	If ParentBoard < 1 Then Exit Function
	Dim Rs,Temp
	Set Rs = LDExeCute("Select BoardID from LeadBBS_Boards where ParentBoard=" & ParentBoard & " and HiddenFlag = 0 order by BoardAssort,OrderID ASC",0)
	If Rs.Eof Then
		Temp = ""
	Else
		Temp = Rs(0)
		Rs.MoveNext
		Do while Not Rs.Eof
			Temp = Temp & "," & Rs(0)
			Rs.MoveNext
		Loop
	End If
	Rs.Close
	Set Rs = Nothing
	CALL LDExeCute("Update LeadBBS_Boards Set LowerBoard='" & Replace(Temp,"'","''") & "' Where BoardID=" & ParentBoard,1)
	ReloadBoardInfo(ParentBoard)

End Function

Rem ���¸�����������
Function UpdateParentBoardStrColumn(ParentOld,ParentNew,BoardID)

	If ParentOld = ParentNew Then Exit Function
	Dim Rs,SQL
	Dim ParentBoardStrOld,ParentBoardStrNew,Level
	If ParentOld = 0 Then
		ParentBoardStrOld = BoardID
	Else
		SQL = sql_select("Select ParentBoardStr from LeadBBS_Boards where BoardID=" & ParentOld,1)
		Set Rs = LDExeCute(SQL,0)
		If Rs.Eof Then
			ParentBoardStrOld = BoardID
		Else
			ParentBoardStrOld = Rs(0)
			If isNull(ParentBoardStrOld) or ParentBoardStrOld = "" Then
				ParentBoardStrOld = BoardID
			Else
				ParentBoardStrOld = ParentBoardStrOld & "," & BoardID
			End If
		End If
		Rs.Close
		Set Rs = Nothing
	End If

	If ParentNew = 0 Then
		ParentBoardStrNew = BoardID
	Else
		SQL = sql_select("Select ParentBoardStr from LeadBBS_Boards where BoardID=" & ParentNew,1)
		Set Rs = LDExeCute(SQL,0)
		If Rs.Eof Then
			ParentBoardStrNew = BoardID
		Else
			ParentBoardStrNew = Rs(0)
			If isNull(ParentBoardStrNew) or ParentBoardStrNew = "" Then
				ParentBoardStrNew = BoardID
			Else
				ParentBoardStrNew = ParentBoardStrNew & "," & BoardID
			End If
		End If
		Rs.Close
		Set Rs = Nothing
	End If

	Dim Temp,Temp1
	SQL = "Select BoardID,ParentBoardStr from LeadBBS_Boards where ParentBoardStr like'" & Replace(ParentBoardStrOld,"'","''") & ",%'"
	Set Rs = LDExeCute(SQL,0)
	Do While Not Rs.Eof
		Temp = cCur(Rs(0))
		Temp1 = Rs(1)
		Temp1 = Replace("a" & Temp1,"a" & ParentBoardStrOld & ",",ParentBoardStrNew & ",")
		Level = Ubound(Split(Temp1,","),1) + 1
		CALL LDExeCute("Update LeadBBS_Boards Set ParentBoardStr='" & Replace(Temp1,"'","''") & "',BoardLevel=" & Level & " where BoardID=" & Temp,1)
		If Temp <> BoardID Then ReloadBoardInfo(Temp)
		Rs.MoveNext
	Loop
	Rs.Close
	Set Rs = Nothing
	Level = Ubound(Split(ParentBoardStrNew,","),1) + 1
	CALL LDExeCute("Update LeadBBS_Boards Set ParentBoardStr='" & Replace(ParentBoardStrNew,"'","''") & "',BoardLevel=" & Level & " where BoardID=" & BoardID & " and ParentBoardStr='" & Replace(ParentBoardStrOld,"'","''") & "'",1)

End Function

Function MakeBoardList(savefile,filename)

	Dim Rs,GetData,BoardNum
	Set Rs = LDExeCute("Select BoardID,BoardAssort,BoardName,BoardIntro,LastWriter,LastWriteTime,TopicNum,AnnounceNum,ForumPass,HiddenFlag,LastAnnounceID,LastTopicName,MasterList,BoardLimit,LeadBBS_Assort.AssortID,LeadBBS_Assort.AssortName,LowerBoard from LeadBBS_Boards left join LeadBBS_Assort on LeadBBS_Assort.AssortID=LeadBBS_Boards.BoardAssort where LeadBBS_Boards.ParentBoard=0 and LeadBBS_Boards.HiddenFlag = 0 order by LeadBBS_Boards.BoardAssort,LeadBBS_Boards.OrderID ASC",0)

	If Not Rs.Eof Then
		GetData = Rs.GetRows(-1)
		BoardNum = Ubound(GetData,2)
	Else
		BoardNum = -1
	End If
	Rs.Close
	Set Rs = Nothing
	
	'on error resume next
	Dim TempStr
	TempStr = ""

	Dim N,WriteStr,LastFlag
If savefile <> "BoardJump2.asp" Then
	TempStr = TempStr & "	" & Chr(60) & "script type=""text/javascript"">" & VbCrLf
	TempStr = TempStr & "	<!--" & VbCrLf
	TempStr = TempStr & "	function surfto1(list)" & VbCrLf
	TempStr = TempStr & "	{" & VbCrLf
	TempStr = TempStr & "		var myindex1  = list.selectedIndex;" & VbCrLf
	TempStr = TempStr & "		if (myindex1 != 0)" & VbCrLf
	TempStr = TempStr & "		{" & VbCrLf
	TempStr = TempStr & "			var URL = ""../"" + list.options[list.selectedIndex].value;" & VbCrLf
	TempStr = TempStr & "			this.location.href = URL; " & VbCrLf
	TempStr = TempStr & "			target = '_self';" & VbCrLf
	TempStr = TempStr & "		}" & VbCrLf
	TempStr = TempStr & "	}" & VbCrLf
	TempStr = TempStr & "	-->" & VbCrLf
	TempStr = TempStr & "	" & Chr(60) & "/script>" & VbCrLf
	TempStr = TempStr & "	<select name=""jumpto"" onchange=""surfto1(this)"" style=""width:100px;"">" & VbCrLf
	TempStr = TempStr & "		<option value=""Boards.asp"">�л����桭</option>" & VbCrLf
	TempStr = TempStr & "		<option value=""Boards.asp"">��̳��ҳ</option>" & VbCrLf

	If BoardNum = -1 Then
	Else
		CurrentAssosrt = -1183
		LastAssosrt = cCur(GetData(1,BoardNum))
		For N = 0 to BoardNum
			WriteStr = ""
			If CurrentAssosrt<>cCur(GetData(1,N)) Then
				CurrentAssosrt = cCur(GetData(1,N))
				If LastAssosrt = CurrentAssosrt Then
					WriteStr = "����"
				Else
					WriteStr = "����"
				End If
				If ViewSelectListFlag = 1 Then WriteStr = "��"
				TempStr = TempStr & "		<option value=""Boards.asp?Assort=" & GetData(14,N) & """>" & WriteStr & KillHTMLLabel(GetData(15,N) & "") & "</option>" & VbCrLf
			End If
			
			If N >= BoardNum Then
				If LastAssosrt = CurrentAssosrt Then
					If GetData(16,n) & ""  = "" Then
						WriteStr = "����"
					Else
						WriteStr = "����"
					End if
				Else
					WriteStr = "����"
				End If

				If ViewSelectListFlag = 1 Then
					WriteStr = "��"
					If GetData(16,n) & "" <> "" Then WriteStr = "����"
				End If
			Else
				If CurrentAssosrt<>cCur(GetData(1,N+1)) Then
					If LastAssosrt = CurrentAssosrt Then
						WriteStr = "����"
					Else
						WriteStr = "����"
					End If
				Else
					If LastAssosrt = CurrentAssosrt Then
						WriteStr = "����"
					Else
						WriteStr = "����"
					End If
				End If
				If ViewSelectListFlag = 1 Then
					WriteStr = "��"
					If GetData(16,n) & "" <> "" Then WriteStr = "����"
				End If
			End If
			WriteStr = WriteStr & KillHTMLLabel(GetData(2,N))
			If StrLength(WriteStr) > 21 Then
				WriteStr = LeftTrue(WriteStr,18) & "..."
			End If
			TempStr = TempStr & "		<option value=""b/" & filename & "?B=" & GetData(0,N) & """>" & WriteStr & "" & "</option>" & VbCrLf
			GBL_LowBoardString = ""
			GBL_LoopN = 0
			GetLowBoardString GetData(16,n),filename
			If GBL_LowBoardString <> "" Then TempStr = TempStr & GBL_LowBoardString
		Next
	End If

	TempStr = TempStr & "	</select>" & VbCrLf
Else	
	If BoardNum = -1 Then
	Else
		CurrentAssosrt = -1183
		LastAssosrt = cCur(GetData(1,BoardNum))
		For N = 0 to BoardNum
			WriteStr = ""
			If CurrentAssosrt<>cCur(GetData(1,N)) Then
				if CurrentAssosrt <> -1183 Then TempStr = TempStr & "</ul>"
				CurrentAssosrt = cCur(GetData(1,N))
				TempStr = TempStr & "<div class=""swap_collapse"" onclick=""swap_view('master_part_" & CurrentAssosrt & "',this);""><span>" & KillHTMLLabel(GetData(15,N) & "") & "</span></div>"
				TempStr = TempStr & "<ul id=""master_part_" & CurrentAssosrt & """>"
			End If
			TempStr = TempStr & "<li><a href=""javascript:url_to(" & GetData(0,N) & ");"">" & KillHTMLLabel(GetData(2,N)) & "</A></li>"
		Next
		if CurrentAssosrt <> -1183 Then TempStr = TempStr & "</ul>"
	End If
end if
	
	ADODB_SaveToFile TempStr,"../../inc/IncHtm/" & savefile & ""
	If GBL_CHK_TempStr = "" Then
		Response.Write "<br><font color=Green class=greenfont>2.�ɹ������ļ�../../inc/IncHtm/" & savefile & "��</font>"
	Else
		%><p><%=GBL_CHK_TempStr%><br>��������֧������д���ļ����ܣ���ʹ��FTP�ȹ��ܣ�<br>��<font color=Red Class=redfont>inc/IncHtm/<%=savefile%></font>�ļ��滻���¿�������(ע�ⱸ��)<p>
		<textarea name="fileContent" cols="80" rows="20" class=fmtxtra><%=Server.htmlencode(TempStr)%></textarea><%
		GBL_CHK_TempStr = ""
	End If

End Function

Dim GBL_LowBoardString,GBL_LoopN
Dim LastAssosrt,CurrentAssosrt
GBL_LoopN = 0

Function GetLowBoardString(LowBoardStr,filename)

	If LowBoardStr = "" or isNull(LowBoardStr) or GBL_LoopN > 100 Then Exit Function
	GBL_LoopN = GBL_LoopN + 1
	Dim BoardNum,LowArray,N
	LowArray = Split(LowBoardStr,",")
	BoardNum = Ubound(LowArray,1)

	Dim Temp
	Dim WriteStr
	For N = 0 to BoardNum
		Temp = Application(DEF_MasterCookies & "BoardInfo" & LowArray(N))
		If isArray(Temp) = False Then
			ReloadBoardInfo(LowArray(N))
			Temp = Application(DEF_MasterCookies & "BoardInfo" & LowArray(N))
		End If
		If isArray(Temp) = True Then
			If Temp(8,0) = 0 Then
				If N >= BoardNum Then
					If LastAssosrt = CurrentAssosrt Then
						WriteStr = "��" & String(GBL_LoopN, "��") & "��"
					Else
						WriteStr = "��" & String(GBL_LoopN, "��") & "��"
					End If
				Else
					If LastAssosrt = CurrentAssosrt Then
						WriteStr = "����"
					Else
						WriteStr = "��" & String(GBL_LoopN, "��") & "��"
					End If
				End If
				If ViewSelectListFlag = 1 Then
					If Temp(27,0) & "" <> "" Then
						WriteStr = String(GBL_LoopN + 1, "��") & "��"
					Else
						WriteStr = String(GBL_LoopN + 1, "��")
					End If
				End If
				'WriteStr = String(GBL_LoopN, "��") & WriteStr
				WriteStr = WriteStr & KillHTMLLabel(Temp(0,0))
				If StrLength(WriteStr) > 21 Then
					WriteStr = LeftTrue(WriteStr,18) & "..."
				End If
				GBL_LowBoardString = GBL_LowBoardString & "		<option value=""b/" & filename & "?B=" & LowArray(N) & """>" & WriteStr & "" & VbCrLf
				GetLowBoardString Temp(27,0),filename
			End If
		End If
	Next
		
	GBL_LoopN = GBL_LoopN - 1
	
End Function


Function GetLowBoardString_Move(LowBoardStr)

	If LowBoardStr = "" or isNull(LowBoardStr) or GBL_LoopN > 100 Then Exit Function
	GBL_LoopN = GBL_LoopN + 1
	Dim BoardNum,LowArray,N
	LowArray = Split(LowBoardStr,",")
	BoardNum = Ubound(LowArray,1)

	Dim Temp
	Dim WriteStr
	For N = 0 to BoardNum
		Temp = Application(DEF_MasterCookies & "BoardInfo" & LowArray(N))
		If isArray(Temp) = False Then
			ReloadBoardInfo(LowArray(N))
			Temp = Application(DEF_MasterCookies & "BoardInfo" & LowArray(N))
		End If
		If isArray(Temp) = True Then
			If Temp(8,0) = 0 Then
				If N >= BoardNum Then
					If LastAssosrt = CurrentAssosrt Then
						WriteStr = "��" & String(GBL_LoopN, "��") & "��"
					Else
						WriteStr = "��" & String(GBL_LoopN, "��") & "��"
					End If
				Else
					If LastAssosrt = CurrentAssosrt Then
						WriteStr = "����"
					Else
						WriteStr = "��" & String(GBL_LoopN, "��") & "��"
					End If
				End If
				If ViewSelectListFlag = 1 Then
					If Temp(27,0) & "" <> "" Then
						WriteStr = String(GBL_LoopN + 1, "��") & "��"
					Else
						WriteStr = String(GBL_LoopN + 1, "��")
					End If
				End If
				'WriteStr = String(GBL_LoopN, "��") & WriteStr
				WriteStr = WriteStr & KillHTMLLabel(Temp(0,0))
				If StrLength(WriteStr) > 21 Then
					WriteStr = LeftTrue(WriteStr,18) & "..."
				End If
				GBL_LowBoardString = GBL_LowBoardString & "		<option value=" & LowArray(N) & ">" & WriteStr & "" & VbCrLf
				GetLowBoardString_Move Temp(27,0)
			End If
		End If
	Next
		
	GBL_LoopN = GBL_LoopN - 1
	
End Function

Function MakeBoardList_For_MoveAnnounce

	Dim Rs,GetData,BoardNum
	Set Rs = LDExeCute("Select BoardID,BoardAssort,BoardName,BoardIntro,LastWriter,LastWriteTime,TopicNum,AnnounceNum,ForumPass,HiddenFlag,LastAnnounceID,LastTopicName,MasterList,BoardLimit,LeadBBS_Assort.AssortID,LeadBBS_Assort.AssortName,LowerBoard from LeadBBS_Boards left join LeadBBS_Assort on LeadBBS_Assort.AssortID=LeadBBS_Boards.BoardAssort where LeadBBS_Boards.ParentBoard=0 and LeadBBS_Boards.HiddenFlag = 0 order by LeadBBS_Boards.BoardAssort,LeadBBS_Boards.OrderID ASC",0)
	If Not Rs.Eof Then
		GetData = Rs.GetRows(-1)
		BoardNum = Ubound(GetData,2)
	Else
		BoardNum = -1
	End If
	Rs.Close
	Set Rs = Nothing

	'on error resume next
	Dim TempStr
	TempStr = ""

	TempStr = TempStr & "	<select name=""BoardID2"">" & VbCrLf
	TempStr = TempStr & "		<option value=""0"">ѡ����桭</option>" & VbCrLf

	If BoardNum = -1 Then
	Else
		Dim CurrentAssosrt,N
		CurrentAssosrt = -1183
		Dim LastAssosrt,WriteStr
		LastAssosrt = cCur(GetData(1,BoardNum))
		Dim LastFlag
		For N = 0 to BoardNum
			WriteStr = ""
			If CurrentAssosrt<>cCur(GetData(1,N)) Then
				CurrentAssosrt = cCur(GetData(1,N))
				If LastAssosrt = CurrentAssosrt Then
					WriteStr = "����"
				Else
					WriteStr = "����"
				End If
				If ViewSelectListFlag = 1 Then WriteStr = "��"
				TempStr = TempStr & "		<option value=""0"">" & WriteStr & KillHTMLLabel(GetData(15,N)) & "" & VbCrLf
			End If
			If N >= BoardNum Then
				If LastAssosrt = CurrentAssosrt Then
					If GetData(16,n) & ""  = "" Then
						WriteStr = "����"
					Else
						WriteStr = "����"
					End if
				Else
					WriteStr = "����"
				End If

				If ViewSelectListFlag = 1 Then
					WriteStr = "��"
					If GetData(16,n) & "" <> "" Then WriteStr = "����"
				End If
			Else
				If CurrentAssosrt<>cCur(GetData(1,N+1)) Then
					If LastAssosrt = CurrentAssosrt Then
						WriteStr = "����"
					Else
						WriteStr = "����"
					End If
				Else
					If LastAssosrt = CurrentAssosrt Then
						WriteStr = "����"
					Else
						WriteStr = "����"
					End If
				End If
				If ViewSelectListFlag = 1 Then
					WriteStr = "��"
					If GetData(16,n) & "" <> "" Then WriteStr = "����"
				End If
			End If
			WriteStr = WriteStr & KillHTMLLabel(GetData(2,N))
			If StrLength(WriteStr) > 21 Then
				WriteStr = LeftTrue(WriteStr,18) & "..."
			End If
			TempStr = TempStr & "		<option value=" & GetData(0,N) & ">" & WriteStr & "" & VbCrLf
			GBL_LowBoardString = ""
			GBL_LoopN = 0
			GetLowBoardString_Move GetData(16,n)
			If GBL_LowBoardString <> "" Then TempStr = TempStr & GBL_LowBoardString
			
		Next
	End If

	TempStr = TempStr & "	</select>" & VbCrLf

	ADODB_SaveToFile TempStr,"../../inc/IncHtm/BoardForMoveList.asp"
	If GBL_CHK_TempStr = "" Then
		Response.Write "<br><font color=Green class=greenfont>2.�ɹ������ļ�inc/IncHtm/BoardForMoveList.asp��</font>"
	Else
		%><p>��������֧������д���ļ����ܣ���ʹ��FTP�ȹ��ܣ�<br>��<font color=Red Class=redfont>inc/IncHtm/BoardForMoveList.asp</font>�ļ��滻�ɿ�������(ע�ⱸ��)<p>
		<textarea name="fileContent" cols="80" rows="20" class=fmtxtra><%=Server.htmlencode(TempStr)%></textarea><%
		GBL_CHK_TempStr = ""
	End If

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
%>