<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Upload_Setup.asp -->
<!-- #include file=../../inc/Board_Popfun.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<%
server.scriptTimeOut = 9999
DEF_BBS_HomeUrl = "../../"
InitDatabase
Dim GBL_ID,Form_ID

Dim DelUserID
DelUserID = Left(Request("DelUserID"),14)
If isNumeric(DelUserID) = 0 or DelUserID = "" or InStr(DelUserID,",") > 0 Then DelUserID = 0
DelUserID = Fix(cCur(DelUserID))

If DelUserID < 0 Then DelUserID = 0

GBL_CHK_TempStr=""
GBL_ID = checkSupervisorPass
Form_ID = GBL_ID
Form_ID = cCur(Form_ID)
If Form_ID < 0 Then Form_ID = 0
If Form_ID = 0 or GBL_CHK_Flag = 0 Then
	GBL_CHK_TempStr = GBL_CHK_TempStr & "��û�е�¼<br>" & VbCrLf
End If

Dim UserName

Manage_sitehead DEF_SiteNameString & " - ����Ա",""
frame_TopInfo

	Dim Rs,SQL,SQLendString,ClearFlag
	If GBL_CHK_TempStr = "" Then
		If GBL_UserID<1 or CheckSupervisorUserName = 0 Then
			SQL = sql_select("Select ID,UserName from LeadBBS_User where ID=-111",1)
		Else
			SQL = sql_select("Select ID,UserName from LeadBBS_User where id=" & DelUserID,1)
		End If
		Set Rs = LDExeCute(SQL,0)
		If Rs.Eof Then
			Response.Write "�Ҳ�����¼��<br>" & VbCrLf
			Rs.Close
			Set Rs = Nothing
		Else
			Dim dflag
			dflag = Left(Request("dflag"),10)
			If Request.Form("DeleteSureFlag")="dk9@dl9s92lw_SWxl" Then
				UserName = Rs(1)
				Rs.Close
				Set Rs = Nothing
				DeleteUserInfo DelUserID,UserName
			Else
				UserName = Rs(1)
				Rs.Close
				Set Rs = Nothing
				%>
				<form name=DellClientForm action=DelUserAllAnnounce.asp method=post>
				<input type=hidden name=DeleteSureFlag value="dk9@dl9s92lw_SWxl">
				<input type=hidden name=DelUserID value="<%=htmlencode(DelUserID)%>">
				<input type=hidden name=dflag value="<%=htmlencode(dflag)%>">
				<%If dflag <> "onlyupload" Then%>
				<b>�˲��������棬ȷ��Ҫɾ���û�<span class=redfont><%=htmlencode(UserName)%></span>���������������
				<br>
				ע�⣬��������<span class=redfont>������̳�б��޸�</span>�Ĺ������Ա�֤��������ͳ��׼ȷ��</b>
				<%Else%>
				<b>�˲��������棬ȷ��Ҫɾ���û�<span class=redfont><%=htmlencode(UserName)%></span>�������ϴ�����ô��</b>
				<%End If%>
				<br><label>
				<input class=fmchkbox type="checkbox" name="delusersure" value="yes" checked=checked>ͬʱɾ���û�</label>
				<p><input type=submit value=ȷ�� class=fmbtn>
				<input type=button value=��ɾ onclick="javascript:window.close();" class=fmbtn>
				</form>
				<%
			End If
		End If
	Else
		Response.Write "<div class=alert>" & GBL_CHK_TempStr & "</div>"
	End If

closeDataBase
frame_BottomInfo
Manage_Sitebottom("none")

Function DeleteUserInfo(DelUserID,UserName)

	Dim Rs,SQL
	Dim NowID,EndFlag
	NowID = 0
	EndFlag = 0

	Response.Write "����(��Ҫʱ����ܻ�ܳ�����ȴ��������)����"
	Response.Flush
	Dim TempNum
	Dim Tmp,Tmp2
	
	If DEF_FSOString = "" Then
		Response.Write " <font color=Red class=redfont>��֧��FSO���Թ�����ɾ����</font>"
	Else
		Do while EndFlag = 0
			SQL = sql_select("Select ID,PhotoDir,SPhotoDir from LeadBBS_Upload where UserID=" & DelUserID & " and ID>" & NowID & " order by ID ASC",100)
			Set Rs = LDExeCute(SQL,0)
			If Rs.Eof Then
				EndFlag = 1
				Rs.Close
				Set Rs = Nothing
			Else
				TempNum = 0
				Do while Not Rs.Eof
					If Rs("PhotoDir") <> "" Then DeleteFiles(Server.Mappath(Replace(Replace(DEF_BBS_HomeUrl & DEF_BBS_UploadPhotoUrl & Rs("PhotoDir"),"/","\"),"\\","\")))
					If Rs("SPhotoDir") <> "" Then DeleteFiles(Server.Mappath(Replace(Replace(DEF_BBS_HomeUrl & DEF_BBS_UploadPhotoUrl & Rs("SPhotoDir"),"/","\"),"\\","\")))
					TempNum = TempNum + 1
					NowID = Rs(0)
					Rs.MoveNext
				Loop
				Response.Write "��"
				Rs.Close
				Set Rs = Nothing
				Response.Flush
				CALL LDExeCute("Delete from LeadBBS_Upload where UserID=" & DelUserID & " and ID<=" & NowID,1)
				CALL LDExeCute("update LeadBBS_SiteInfo set UploadNum=UploadNum-" & TempNum,1)
				CALL LDExeCute("Update LeadBBS_User Set UploadNum=UploadNum-" & TempNum & " where id=" & DelUserID,1)
			End If
		Loop
		
		SQL = "Select ID,PhotoDir,SPhotoDir from LeadBBS_UserFace where UserID=" & DelUserID & " order by ID ASC"
		Set Rs = LDExeCute(SQL,0)
		If Rs.Eof Then
			Rs.Close
			Set Rs = Nothing
		Else
			Tmp = lcase(replace(replace(Rs("PhotoDir"),"/","\"),"\\","\"))
			If inStr(Tmp,"\face\") Then
				Tmp = mid(Tmp,instrrev(Tmp,"\face\")+6)
			End If
			If Tmp <> "" Then DeleteFiles(Server.Mappath(Replace(Replace(DEF_BBS_HomeUrl & DEF_BBS_UploadPhotoUrl & "face\" & Tmp,"/","\"),"\\","\")))
			
			Tmp = lcase(replace(replace(Rs("SPhotoDir"),"/","\"),"\\","\"))
			If inStr(Tmp,"\face\") Then
				Tmp = mid(Tmp,instrrev(Tmp,"\face\")+6)
			End If
			If Tmp <> "" Then DeleteFiles(Server.Mappath(Replace(Replace(DEF_BBS_HomeUrl & DEF_BBS_UploadPhotoUrl & "face\" & Tmp,"/","\"),"\\","\")))
			Response.Write "��"
			Rs.Close
			Set Rs = Nothing
			Response.Flush
			CALL LDExeCute("Delete from LeadBBS_UserFace where UserID=" & DelUserID,1)
		End If
	End If

	If Left(Request("dflag"),10) = "onlyupload" Then
		Response.Write " <b><font color=green class=greenfont>����û�����ɾ��!</font></b>"
		Exit Function
	End If

	NowID = 0
	EndFlag = 0
	Do while EndFlag = 0
		SQL = sql_select("Select ID from LeadBBS_FriendUser where UserID=" & DelUserID & " and ID>" & NowID & " order by ID ASC",100)
		Set Rs = LDExeCute(SQL,0)
		If Rs.Eof Then
			EndFlag = 1
			Rs.Close
			Set Rs = Nothing
		Else
			Do While Not Rs.Eof
				NowID = Rs(0)
				Rs.MoveNext
			Loop
			Response.Write "��"
			Rs.Close
			Set Rs = Nothing
			Response.Flush
			CALL LDExeCute("Delete from LeadBBS_FriendUser where UserID=" & DelUserID & " and ID<=" & NowID,1)
		End If
	Loop

	NowID = 0
	EndFlag = 0

	Response.Write "��"
	Do while EndFlag = 0
		SQL = sql_select("Select ID from LeadBBS_FriendUser where FriendUserID=" & DelUserID & " and ID>" & NowID & " order by ID ASC",100)
		Set Rs = LDExeCute(SQL,0)
		If Rs.Eof Then
			EndFlag = 1
			Rs.Close
			Set Rs = Nothing
		Else
			Do While Not Rs.Eof
				NowID = Rs(0)
				Rs.MoveNext
			Loop
			Response.Write "��"
			Rs.Close
			Set Rs = Nothing
			Response.Flush
			CALL LDExeCute("Delete from LeadBBS_FriendUser where FriendUserID=" & DelUserID & " and ID<=" & NowID,1)
		End If
	Loop

	NowID = 0
	EndFlag = 0

	Response.Write "��"
	Do while EndFlag = 0
		SQL = sql_select("Select ID from LeadBBS_InfoBox where FromUser='" & Replace(UserName,"'","''") & "' and ID>" & NowID & " order by ID ASC",100)
		Set Rs = LDExeCute(SQL,0)
		If Rs.Eof Then
			EndFlag = 1
			Rs.Close
			Set Rs = Nothing
		Else
			Do While Not Rs.Eof
				NowID = Rs(0)
				Rs.MoveNext
			Loop
			Response.Write "��"
			Rs.Close
			Set Rs = Nothing
			Response.Flush
			CALL LDExeCute("Delete from LeadBBS_InfoBox where FromUser='" & Replace(UserName,"'","''") & "' and ID<=" & NowID,1)
		End If
	Loop

	NowID = 0
	EndFlag = 0

	Response.Write "��"
	Do while EndFlag = 0
		SQL = sql_select("Select ID from LeadBBS_InfoBox where ToUser='" & Replace(UserName,"'","''") & "' and ID>" & NowID & " order by ID ASC",100)
		Set Rs = LDExeCute(SQL,0)
		If Rs.Eof Then
			EndFlag = 1
			Rs.Close
			Set Rs = Nothing
		Else
			Do while Not Rs.Eof
				NowID = Rs(0)
				Rs.MoveNext
			Loop
			Response.Write "��"
			Rs.Close
			Set Rs = Nothing
			Response.Flush
			'CALL LDExeCute("Delete from LeadBBS_InfoBox where ToUser='" & Replace(UserName,"'","''") & "' and ID<=" & NowID,1)
		End If
	Loop

	NowID = 0
	EndFlag = 0
	Dim GetData,N

	Response.Write "��"
	Do while EndFlag = 0
		SQL = sql_select("Select ID from LeadBBS_Announce where UserID=" & DelUserID & " and ID>" & NowID & " order by ID ASC",100)
		Set Rs = LDExeCute(SQL,0)
		If Rs.Eof Then
			EndFlag = 1
			Rs.Close
			Set Rs = Nothing
		Else
			GetData = Rs.GetRows(-1)
			Rs.Close
			Set Rs = Nothing
			SQL = Ubound(GetData,2)
			For N = 0 to SQL
				NowID = cCur(GetData(0,n))
				DelAnnounce(NowID)
				CALL LDExeCute("delete from LeadBBS_Announce Where ID=" & NowID,1)
			Next
			Response.Write "��"
			Response.Flush
		End If
	Loop
	
	CALL LDExeCute("Delete from LeadBBS_Assessor where UserName='" & Replace(UserName,"'","''") & "'",1)
	Response.Write "��"

	ReloadTopAnnounceInfo(0)
	ReloadOtherTopAnnounce
	ReloadStatisticData
	
	If Request("delusersure") = "yes" Then DeleteUser(DelUserID)
	Response.Write " <b><font color=green class=greenfont>���!</font></b>"

End Function

Function DelAnnounce(DelID)

	Dim Rs,SQL
	SQL = sql_select("Select BoardID,ParentID,UserID,RootID,Layer,TopicType,ndatetime,GoodFlag,RootIDBak from LeadBBS_Announce where id=" & DelID,100)
	Set Rs = LDExeCute(SQL,0)
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		Exit Function
	End if
	Dim BoardID,ParentID,UserID,RootID,Layer,TopicType,ndatetime,GoodFlag,RootIDBak
	BoardID = cCur(Rs("BoardID"))
	ParentID = cCur(Rs("ParentID"))
	UserID = cCur(Rs("UserID"))
	RootID = cCur(Rs("RootID"))
	Layer = cCur(Rs("Layer"))
	TopicType = Rs("TopicType")
	ndatetime = RestoreTime(Rs("ndatetime"))
	If isTrueDate(ndatetime) = 0 Then ndatetime = DEF_Now
	If IsNull(TopicType) Then TopicType = 0
	GoodFlag = Rs("GoodFlag")
	RootIDBak = cCur(Rs("RootIDBak"))
	Rs.Close
	Set Rs = Nothing
	
	Dim RootIDTopicType
	If ParentID > 0 Then
		select case DEF_UsedDataBase
			case 0,2:
				SQL = sql_select("Select TopicType from LeadBBS_Announce where ParentID=0 and RootIDBak=" & RootIDBak,1)
			case Else
				SQL = sql_select("Select TopicType from LeadBBS_Topic where ID=" & RootIDBak,1)
		End select
		Set Rs = LDExeCute(SQL,0)
		If Not Rs.Eof Then
			RootIDTopicType = Rs(0)
			If isNull(RootIDTopicType) Then RootIDTopicType = 0
		Else
			RootIDTopicType = 0
		End If
		Rs.Close
	Else
		RootIDTopicType = 0
	End If
	Set Rs = Nothing
	
	Dim todayAnnounce,GoodNum,GoodNum2
	todayAnnounce = 0
	
	If Day(ndatetime) = Day(DEF_Now) and Year(ndatetime) = year(DEF_Now) and Month(ndatetime) = Month(DEF_Now) Then todayAnnounce = 1
	If GoodFlag = 1 Then
		GoodNum = 1
	Else
		GoodNum = 0
	End If
	GoodNum2 = 0

	If ParentID > 0 Then
		CALL LDExeCute("Delete from LeadBBS_Announce where id=" & DelID,1)
		CALL LDExeCute("Delete from LeadBBS_Opinion Where AnnounceID=" & DelID,1)
		Dim TmpMaxID,TmpMinID
		SQL = "select max(ID) from LeadBBS_Announce where RootIDBak=" & RootIDBak
		Set Rs = LDExeCute(SQL,0)
		If Rs.Eof Then
			TmpMaxID = 0
		Else
			TmpMaxID = Rs(0)
			If isNull(TmpMaxID) then TmpMaxID = 0
		End If
		Rs.Close
		Set Rs = Nothing
		SQL = "select min(ID) from LeadBBS_Announce where RootIDBak=" & RootIDBak
		Set Rs = LDExeCute(SQL,0)
		If Rs.Eof Then
			TmpMinID = 0
		Else
			TmpMinID = Rs(0)
			If isNull(TmpMinID) then TmpMinID = 0
		End If
		Rs.Close
		Set Rs = Nothing
		If ParentID = 0 Then
			CALL LDExeCute("Update LeadBBS_Boards Set TopicNum=TopicNum-1,AnnounceNum=AnnounceNum-1,todayAnnounce=todayAnnounce-" & todayAnnounce & ",GoodNum=GoodNum-" & GoodNum & " where BoardID=" & BoardID,1)
		Else
			SQL = sql_select("Select ID,ndatetime,UserName,ParentID from LeadBBS_Announce where id=" & TmpMaxID,1)
			Set Rs = LDExeCute(SQL,0)
			Dim LastTime,LastUser
			If Not Rs.Eof Then
				LastTime = Rs("ndatetime")
				If cCur(Rs("ParentID")) = 0 Then
					LastUser = ""
				Else
					LastUser = Rs("UserName")
				End If
			Else
				LastTime = GetTimeValue(DEF_Now)
			End If
			Rs.Close
			Set Rs = Nothing
			CALL LDExeCute("Update LeadBBS_Boards Set AnnounceNum=AnnounceNum-1,todayAnnounce=todayAnnounce-" & todayAnnounce & ",GoodNum=GoodNum-" & GoodNum & " where BoardID=" & BoardID,1)
			If Layer>2 Then			
				select case DEF_UsedDataBase
					case 0,2:
						CALL LDExeCute("Update LeadBBS_Announce Set ChildNum=ChildNum-1,LastTime=" & LastTime & ",LastUser='" & Replace(LastUser,"'","''") & "',LastInfo='' where ParentID=0 and RootIDBak=" & RootIDBak,1)
						CALL LDExeCute("Update LeadBBS_Announce Set ChildNum=ChildNum-1,LastTime=" & LastTime & ",LastUser='" & Replace(LastUser,"'","''") & "' where id=" & ParentID,1)
					case Else
						CALL LDExeCute("Update LeadBBS_Announce Set ChildNum=ChildNum-1,LastTime=" & LastTime & ",LastUser='" & Replace(LastUser,"'","''") & "',LastInfo='' where ID=" & RootIDBak,1)
						CALL LDExeCute("Update LeadBBS_Announce Set ChildNum=ChildNum-1,LastTime=" & LastTime & ",LastUser='" & Replace(LastUser,"'","''") & "' where id=" & ParentID,1)
						CALL LDExeCute("Update LeadBBS_Topic Set ChildNum=ChildNum-1,LastTime=" & LastTime & ",LastUser='" & Replace(LastUser,"'","''") & "',LastInfo='' where ID=" & RootIDBak,1)
						CALL LDExeCute("Update LeadBBS_Topic Set ChildNum=ChildNum-1,LastTime=" & LastTime & ",LastUser='" & Replace(LastUser,"'","''") & "' where id=" & ParentID,1)
				End select
			Else
				CALL LDExeCute("Update LeadBBS_Announce Set ChildNum=ChildNum-1,LastTime=" & LastTime & ",LastUser='" & Replace(LastUser,"'","''") & "' where id=" & RootIDBak,1)
				If DEF_UsedDataBase = 1 Then CALL LDExeCute("Update LeadBBS_Topic Set ChildNum=ChildNum-1,LastTime=" & LastTime & ",LastUser='" & Replace(LastUser,"'","''") & "',LastInfo='' where id=" & RootIDBak,1)
			End If				
		End If
		CALL LDExeCute("Update LeadBBS_User set AnnounceNum=AnnounceNum-1,AnnounceGood=AnnounceGood-" & GoodNum & " Where ID =" & UserID,1)
		'UpdateBoardValue(BoardID)
		Rem ����MaxRootID

		select case DEF_UsedDataBase
			case 0,2:
				SQL = "Update LeadBBS_Announce set RootMaxID=" & TmpMaxID & _
						",RootMinID=" & TmpMinID &_
						" where ParentID=0 and RootIDBak=" & RootIDBak
				CALL LDExeCute(SQL,1)
			case Else
				SQL = "Update LeadBBS_Announce set RootMaxID=" & TmpMaxID & _
						",RootMinID=" & TmpMinID &_
						" where ID=" & RootIDBak
				CALL LDExeCute(SQL,1)
				SQL = "Update LeadBBS_Topic set RootMaxID=" & TmpMaxID & _
						",RootMinID=" & TmpMinID &_
						" where ID=" & RootIDBak
				CALL LDExeCute(SQL,1)
		End select
	Else
		Dim LoopFlag,NowID,GetData,N
		LoopFlag = 1
		NowID = 0

		GoodNum2 = 0
		Dim TmpTopicNum,TmpAnnounceNum,GoodNum3
		TmpAnnounceNum = 0
		TmpTopicNum = 0
		todayAnnounce = 0
		Do While LoopFlag = 1
			Rem �������������ɾ��������ɾ������������ֹ������������
			SQL = sql_select("Select ID,BoardID,ParentID,UserID,ndatetime,GoodFlag,Opinion from LeadBBS_Announce where RootIDBak=" & RootIDBak & " and id>" & NowID & " order by ID",100)
			Set Rs = LDExeCute(SQL,0)
			If Not Rs.Eof Then
				GetData = Rs.GetRows(100)
				Rs.Close
				Set Rs = Nothing
				SQL = Ubound(GetData,2)
				For N = 0 to SQL
					If GetData(5,n) = 1 Then
						GoodNum2 = GoodNum2 + 1
						GoodNum3 = 1
					Else
						GoodNum3 = 0
					End If
					If cCur(GetData(2,N)) = 0 Then
						TmpAnnounceNum = TmpAnnounceNum + 1
						TmpTopicNum = TmpTopicNum + 1
					Else
						TmpAnnounceNum = TmpAnnounceNum + 1
					End If
					GetData(4,n) = RestoreTime(GetData(4,n))
					If isTrueDate(GetData(4,n)) = 1 Then
						If Day(GetData(4,n)) = Day(DEF_Now) and Year(GetData(4,n)) = year(DEF_Now) and Month(GetData(4,n)) = Month(DEF_Now) Then todayAnnounce = todayAnnounce + 1
					End If
					If GetData(6,n) & "" <> "" Then CALL LDExeCute("Delete from LeadBBS_Opinion Where AnnounceID=" & GetData(0,n),1)
					If cCur(GetData(2,N)) = 0 Then
						CALL LDExeCute("Update LeadBBS_User set AnnounceNum=AnnounceNum-1,AnnounceTopic=AnnounceTopic-1,AnnounceGood=AnnounceGood-" & GoodNum3 & " Where ID =" & GetData(3,N),1)
					Else
						CALL LDExeCute("Update LeadBBS_User set AnnounceNum=AnnounceNum-1,AnnounceGood=AnnounceGood-" & GoodNum3 & " Where ID =" & GetData(3,N),1)
					End If
					NowID = cCur(GetData(0,N))
				Next
				CALL LDExeCute("Delete from LeadBBS_Announce where RootIDBak=" & RootIDBak & " and id<=" & NowID,1)
			Else
				LoopFlag = 0
				Rs.Close
				Set Rs = Nothing
			End If
		Loop
		If DEF_UsedDataBase = 1 Then CALL LDExeCute("Delete from LeadBBS_Topic where ID=" & RootIDBak,1)
		If TopicType = 80 or TopicType = 54 or TopicType = 114 Then
			CALL LDExeCute("Delete from LeadBBS_VoteUser Where AnnounceID=" & DelID,1)
			CALL LDExeCute("Delete from LeadBBS_Opinion Where AnnounceID=" & DelID,1)
			If TopicType = 80 Then
				CALL LDExeCute("Delete from LeadBBS_VoteItem Where AnnounceID=" & DelID,1)
			End If
		End If
		CALL LDExeCute("Update LeadBBS_Boards Set TopicNum=TopicNum-" & TmpTopicNum & ",AnnounceNum=AnnounceNum-" & TmpAnnounceNum & ",todayAnnounce=todayAnnounce-" & todayAnnounce & ",GoodNum=GoodNum-" & GoodNum2 & " where BoardID=" & BoardID,1)
	End If

End Function

Function DeleteFiles(path)

	If DEF_FSOString = "" Then Exit Function
	on error resume next
	Dim fs
	Set fs = Server.CreateObject(DEF_FSOString)
	If err <> 0 Then
		Err.Clear
		'Response.Write "<br>��������֧��FSO��Ӳ���ļ�δɾ����"
		Exit Function
	End If
	If fs.FileExists(path) Then
		fs.DeleteFile path,True
		DeleteFiles = 1
	Else
		DeleteFiles = 0
	End If
	Set fs = Nothing
         
End Function

Sub ReloadOtherTopAnnounce

	Dim Rs,SQL,GetData,N
	Set Rs = LDExeCute("Select AssortID from LeadBBS_Assort",0)
	If Not Rs.Eof Then
		GetData = Rs.GetRows(-1)
		Rs.Close
		Set Rs = Nothing
		SQL = Ubound(GetData,2)
		For N = 0 to SQL
			ReloadTopAnnounceInfo(cCur(GetData(0,n)))
		Next
	Else
		Rs.Close
		Set Rs = Nothing
	End If

End Sub

Sub ReloadTopAnnounceInfo(TID)

	Dim Rs,GetDataTop,TIDStr
	If TID = 0 Then
		TIDStr = ""
	Else
		TIDStr = TID
	End If
	Set Rs = LDExeCute("Select RootID,BoardID from LeadBBS_TopAnnounce where TopType=" & TID,0)
	If Rs.Eof Then
		Application.Lock
		application(DEF_MasterCookies & "TopAnc" & TIDStr) = "yes"
		application(DEF_MasterCookies & "TopAncList" & TIDStr) = ""
		Application.UnLock
		Rs.Close
		Set Rs = Nothing
		Exit Sub
	Else
		GetDataTop = Rs.GetRows(-1)
		Rs.close
		Set Rs = Nothing
	End If
	
	Dim Temp,N
	Temp = ""
	If cCur(GetDataTop(0,0)) > 0 Then Temp = GetDataTop(0,0)
	For N = 1 to Ubound(GetDataTop,2)
		If cCur(GetDataTop(0,N)) > 0 Then Temp = Temp & "," & GetDataTop(0,N)
	Next
	If Left(Temp,1) = "," Then Temp = Mid(Temp,2)
	If cStr(Temp) <> "" Then
		Set Rs = LDExeCute(sql_select("select id,ChildNum,Title,FaceIcon,LastTime,Hits,Length,UserName,UserID,RootID,LastUser,NotReplay,GoodFlag,BoardID,TopicType,PollNum,TitleStyle,LastInfo,ndatetime,GoodAssort,NeedValue from LeadBBS_Announce where ParentID=0 and RootIDBak in(" & Temp & ")",Ubound(GetDataTop,2)+1),0)
		If Not Rs.Eof Then
			GetDataTop = Rs.GetRows(-1)
			Rs.Close
			Set Rs = Nothing
			Application.Lock
			application(DEF_MasterCookies & "TopAnc" & TIDStr) = GetDataTop
			Application.UnLock
			Application.Lock
			application(DEF_MasterCookies & "TopAncList" & TIDStr) = "," & Temp & ","
			Application.UnLock
		Else
			Rs.Close
			Set Rs = Nothing
			Application.Lock
			application(DEF_MasterCookies & "TopAnc" & TIDStr) = "yes"
			Application.UnLock
		End If
	Else
		Application.Lock
		application(DEF_MasterCookies & "TopAnc" & TIDStr) = "yes"
		Application.UnLock
	End If

End Sub

rem ɾ��ĳ�û�
Function DeleteUser(ID)

	Dim Rs
	Set Rs = LDExeCute(sql_select("Select ID,UserName from LeadBBS_User where ID=" & ID,1),0)
	If Rs.Eof Then
		DeleteUser = 0
		Rs.Close
		Set Rs = Nothing
		GBL_CHK_TempStr = GBL_CHK_TempStr & "�Ҳ������û���<br>" & VbCrLf
	Else
		GBL_CHK_User = Rs("UserName")
		If CheckSupervisorUserName = 1 Then
			GBL_CHK_TempStr = GBL_CHK_TempStr & "��������Ա����ɾ����<br>" & VbCrLf
			Rs.Close
			Set Rs = Nothing
			DeleteUser = 0
			Exit Function
		End If
		Rs.Close
		Set Rs = Nothing
		CALL LDExeCute("delete from LeadBBS_SpecialUser where UserID=" & ID,1)
		CALL LDExeCute("delete from LeadBBS_User where ID=" & ID,1)
		
		
		Response.Write "<br><p><font color=008800 class=greenfont><b>�Ѿ��ɹ�ɾ��IDΪ" & ID & "���û���</b></font></p>"
		CALL LDExeCute("Update LeadBBS_SiteInfo Set UserCount=UserCount-1",1)
		UpdateStatisticDataInfo -1,1,1
		DeleteUser = 1
	End if

End Function
%>