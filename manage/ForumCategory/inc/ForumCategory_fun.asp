<%
Dim GBL_AssortID,GBL_AssortName,GBL_AssortMaster,GBL_AssortLimit,GBL_GetData
Dim GBL_AssortID_Old
GBL_AssortID_Old = 1

Dim LimitAssortStringData,LimitAssortStringDataNum
LimitAssortStringData = Array("��Լ��ʾ����")
LimitAssortStringDataNum = Ubound(LimitAssortStringData,1)

Rem ������֤
Function CheckFormForumCateGoryData

	If isNumeric(GBL_AssortID) = 0 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ����Ϊ��̳����IDָ��һ������0�����֣��������������ַ���<br>" & VbCrLf
		CheckFormForumCateGoryData = 0
		Exit Function
	End If

	GBL_AssortID = cCur(GBL_AssortID)
	If GBL_AssortID > 2147479999 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ��̳����ID��д̫��<br>" & VbCrLf
		CheckFormForumBoardData = 0
		Exit Function
	End If
	If GBL_AssortID < 1 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ����Ϊ��̳����IDָ��һ������0�����֡�<br>" & VbCrLf
		CheckFormForumCateGoryData = 0
		Exit Function
	End If

	If len(GBL_AssortName)<1 or GBL_AssortName = "" Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ��̳���������Ǳ�����<br>" & VbCrLf
		CheckFormForumCateGoryData = 0
		Exit Function
	End If
	
	If inStr(LCase(GBL_AssortName),"""") > 0 or inStr(GBL_AssortName,"<script") > 0 or inStr(GBL_AssortName,"<\script") > 0 or inStr(GBL_AssortName,"</script") > 0 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ��̳�������Ʋ��������js���������룬������ʹ��˫����<br>" & VbCrLf
		CheckFormForumCateGoryData = 0
		Exit Function
	End If

	If strLength(GBL_AssortName) > 250 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: �������Ƴ��Ȳ��ܳ���250���ַ�<br>" & VbCrLf
		CheckFormForumCateGoryData = 0
		Exit Function
	End If

	If strLength(GBL_AssortMaster) > 250 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ��������������Ȳ��ܳ���250���ַ�<br>" & VbCrLf
		CheckFormForumCateGoryData = 0
		Exit Function
	End If
	
	Dim GBL_AssortMasterArray,GBL_AssortMaster_OldD
	GBL_AssortMasterArray = Split(GBL_AssortMaster,",")
	GBL_AssortMaster_OldD = GBL_AssortMaster
	Dim TempN,TempName
	If Ubound(GBL_AssortMasterArray,1) = 0 and GBL_AssortMaster = "?LeadBBS?" Then
		
	Else
		GBL_AssortMaster = ""
		If Ubound(GBL_AssortMasterArray,1) > DEF_MaxBoardMastNum - 1 Then
			GBL_CHK_TempStr = "����" & DEF_PointsName(7) & "���ֻ������" & DEF_MaxBoardMastNum & "��"
			CheckFormForumCateGoryData = 0
			GBL_AssortMaster = GBL_AssortMaster_OldD
			Exit Function
		End if

		For TempN = 0 to Ubound(GBL_AssortMasterArray,1)
			If Trim(GBL_AssortMasterArray(TempN)) <> "" Then
				TempName = CheckUserNameExist(GBL_AssortMasterArray(TempN))
				If TempName = "" Then
					GBL_CHK_TempStr = "Error: " & DEF_PointsName(8) & "�б�����û�" & htmlencode(GBL_AssortMasterArray(TempN)) & "�����ڣ���<br>" & VbCrLf
					CheckFormForumCateGoryData = 0
					GBL_AssortMaster = GBL_AssortMaster_OldD
					Exit Function
				Else
					GBL_AssortMaster = GBL_AssortMaster & "," & TempName
				End If
			End If
		Next
		If Left(GBL_AssortMaster,1) = "," Then GBL_AssortMaster = Mid(GBL_AssortMaster,2)
	End If

	CheckFormForumCateGoryData = 1

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

Rem ���ĳ���������Ƿ����
Function CheckForumAssortNameExist(AssortName)

	Dim Rs
	Set Rs = LDExeCute(sql_select("Select AssortID from LeadBBS_Assort where AssortName='" & Replace(AssortName,"'","''") & "'",1),0)
	If Rs.Eof Then
		CheckForumAssortNameExist = 0
	Else
		CheckForumAssortNameExist = cCur(rs(0))
	End if
	Rs.Close
	Set Rs = Nothing

End Function

Rem ɾ��ĳ����
Function DeleteForumAssort(AssortID)

	Dim Rs,AssortMaster
	Set Rs = LDExeCute(sql_select("Select AssortID,AssortMaster from LeadBBS_Assort where AssortID=" & AssortID,1),0)
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ��̳����ID��" & AssortID & "������!<br>" & VbCrLf
		DeleteForumAssort = 0
		Exit Function
	Else
		AssortMaster = Rs(1)
		Rs.Close
		Set Rs = Nothing
		Set Rs = LDExeCute(sql_select("Select BoardAssort from LeadBBS_Boards where BoardAssort=" & AssortID,1),0)
		If Not Rs.Eof Then
			GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: �˷����»��а�����ڣ��������ɾ������!<br>" & VbCrLf
			DeleteForumAssort = 0
			Rs.Close
			Set Rs = Nothing
			Exit Function
		End If
		Rs.Close
		Set Rs = Nothing

		GBL_AssortID = AssortID
		UpdateAssortMasterList AssortMaster,0
		CALL LDExeCute("delete from LeadBBS_Assort where AssortID=" & AssortID,1)
		DeleteForumAssort = 1
	End if

End Function

Rem ����ĳ����
Function InsertForumAssort

	If CheckForumAssortIDExist(GBL_AssortID) = 1 Then
		InsertForumAssort = 0
		GBL_CHK_TempStr = GBL_CHK_TempStr & "����ID��" & GBL_AssortID & "�Ѿ�����!<br>" & VbCrLf
		Exit Function
	End If
	
	If CheckForumAssortNameExist(GBL_AssortName) = 1 Then
		InsertForumAssort = 0
		GBL_CHK_TempStr = GBL_CHK_TempStr & "�������ƺ�" & htmlencode(GBL_AssortName) & "�Ѿ�����!<br>" & VbCrLf
		Exit Function
	End If

	CALL LDExeCute("insert into LeadBBS_Assort(AssortID,AssortName,AssortMaster) values(" &_
			GBL_AssortID & ",'" & Replace(GBL_AssortName,"'","''") & "','" & Replace(GBL_AssortMaster,"'","''") & "')",1)

	GBL_AssortID_Old = GBL_AssortID
	UpdateAssortMasterList GBL_AssortMaster,1
	InsertForumAssort = 1

End Function

Rem �õ�ĳ������Ϣ
Function GetForumAssortData(AssortID)

	Dim Rs
	Set Rs = LDExeCute("Select AssortID,AssortName,AssortMaster,AssortLimit from LeadBBS_Assort Where AssortID = " & AssortID,0)
	If Rs.Eof Then
		GetForumAssortData = 0
		Rs.Close
		Set Rs = Nothing
		Exit Function
	Else
		GBL_GetData = Rs.GetRows(-1)
		Rs.Close
		Set Rs = Nothing
		GetForumAssortData = 1
		Exit Function
	End If

End Function

Rem ����ĳ����
Function UpdateForumAssort
	
	If isNumeric(GBL_MODIFYID) = 0 or GBL_MODIFYID = "" Then GBL_MODIFYID = 0
	GBL_MODIFYID = cCur(GBL_MODIFYID)
	If GBL_MODIFYID = 0 or GBL_MODIFYID<1 then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: Ҫ�޸ĵķ��಻���ڣ�<br>" & VbCrLf
		GBL_CHK_Flag = 0
		UpdateForumAssort = 0
		Exit Function
	End If

	If GetForumAssortData(GBL_MODIFYID) = 0 Then
		GBL_CHK_Flag = 0
		UpdateForumAssort = 0
		Exit Function
	End If

	If cCur(GBL_GetData(0,0))<>GBL_AssortID and CheckForumAssortIDExist(GBL_AssortID) = 1 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ����ID��" & GBL_AssortID & "�Ѿ����ڣ���ʹ������ID�š�<br>" & VbCrLf
		GBL_CHK_Flag = 0
		UpdateForumAssort = 0
		Exit Function
	End If
	Dim Temp
	Temp = CheckForumAssortNameExist(GBL_AssortName)
	If Temp<>0 and Temp<>cCur(GBL_GetData(0,0)) Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ͬ���������Ѿ���������Ϊ<b>" & htmlencode(GBL_AssortName) & "</b>�ķ���<br>" & VbCrLf
		GBL_CHK_Flag = 0
		UpdateForumAssort = 0
		Exit Function
	End If

	If GBL_AssortID <> cCur(GBL_GetData(0,0)) Then
		CALL LDExeCute("Update LeadBBS_Boards Set BoardAssort=" & GBL_AssortID & " where BoardAssort=" & GBL_GetData(0,0),1)
	End If

	GBL_AssortID_Old = cCur(GBL_GetData(0,0))
	UpdateAssortMasterList GBL_GetData(2,0),0
	CALL LDExeCute("Update LeadBBS_Assort Set AssortID=" & GBL_AssortID & ",AssortName='" & Replace(GBL_AssortName,"'","''") & "',AssortMaster='" & Replace(GBL_AssortMaster,"'","''") & "',AssortLimit=" & GBL_AssortLimit & " where AssortID=" & GBL_GetData(0,0),1)
	GBL_AssortID_Old = GBL_AssortID
	UpdateAssortMasterList GBL_AssortMaster,1
	UpdateForumAssort = 1
	ReloadBoardApplicationInfo

End Function

Function ReloadBoardApplicationInfo

	Dim Rs,SQL,GetData
	SQL = "Select BoardID from LeadBBS_Boards Where BoardAssort=" & GBL_AssortID
	Set Rs = LDExeCute(SQL,0)
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		Exit Function
	End If
	GetData = Rs.GetRows(-1)
	Rs.Close
	Set Rs = Nothing
	SQL = Ubound(GetData,2)
	Dim N
	For N = 0 to SQL
		ReloadBoardInfo(GetData(0,n))
	Next

End Function

Function UpdateAssortMasterList(AssortMaster,Flag)

	Rem ���¸�����̳�û�����״̬
	Dim TA,N

	TA = Split(AssortMaster,",")
	For N = 0 to Ubound(TA,1)
		If TA(N) <> "" Then SetUserAssortMastFlag TA(N),Flag
	Next

End Function


Rem ����ĳ�û��Ƿ����
Function SetUserAssortMastFlag(UserName,Fla)

	Dim Flag
	Flag = Fla
	If Flag <> 1 and Flag <> 0 Then Flag = 0
	Fla = Flag
	Dim Rs,Temp,SQL
	If Flag = 0 Then
		SQL = sql_select("Select AssortID from LeadBBS_Assort where AssortID<>" & GBL_AssortID_Old & " and (AssortMaster='" & Replace(UserName,"'","''") & "' or AssortMaster like'" & Replace(UserName,"'","''") & ",%' or AssortMaster like'%," & Replace(UserName,"'","''") & "' or AssortMaster like'%," & Replace(UserName,"'","''") & ",%')",1)
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
		Temp = SetBinarybit(Temp,14,Flag)
		Rs.Close
		Set Rs = Nothing
		SetUserAssortMastFlag = 1
		CALL LDExeCute("Update LeadBBS_User Set UserLimit=" & Temp & " where UserName='" & Replace(UserName,"'","''") & "'",1)
		If Fla = 0 Then
			CALL LDExeCute("Delete from LeadBBS_SpecialUser where Assort=7 and UserID=" & Tmp & " and BoardID=" & GBL_AssortID,1)
		Else
			Set Rs = LDExeCute(sql_select("Select ID from LeadBBS_SpecialUser Where Assort=7 and UserID=" & Tmp & " and BoardID=" & GBL_AssortID,1),0)
			If Rs.Eof Then
				Rs.Close
				Set Rs = Nothing
				CALL LDExeCute("insert into LeadBBS_SpecialUser(UserID,UserName,BoardID,Assort,ndatetime) values(" & Tmp & ",'" & Replace(UserName,"'","''") & "'," & GBL_AssortID & ",7," & GetTimeValue(DEF_Now) & ")",1)
			Else
				Rs.Close
				Set Rs = Nothing
			End If
		End If
	Else
		SetUserAssortMastFlag = 0
		Rs.Close
		Set Rs = Nothing
	End if

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