<%
Function MakeAnnounceTop(AnnounceID,morestr)

	Dim Rs,SQL,BoardID,RootID,MaxRootID,RootIDBak
	SQL = sql_select("Select ID,RootID,BoardID,RootIDBak from LeadBBS_Announce where ID=" & AnnounceID,1)
	Set Rs = LDExeCute(SQL,0)
	If Not Rs.Eof Then
		RootID = cCur(Rs(1))
		BoardID = cCur(Rs(2))
		RootIDBak = cCur(Rs(3))
		Rs.Close
		Set Rs = Nothing
	Else
		Rs.Close
		Set Rs = Nothing
		Exit function
	End If

	If RootID<DEF_BBS_TOPMinID Then
		select case DEF_UsedDataBase
			case 0:
	 			CALL LDExeCute("Update LeadBBS_Announce Set RootID=(Select Max(RootID) from LeadBBS_Announce where ParentID=0 and BoardID=" & BoardID & " and RootID<" & DEF_BBS_TOPMinID & ")+1" & morestr & " where ParentID=0 and RootIDBak=" & RootIDBak,1)
	 		case 2:
	 			CALL LDExeCute("Update LeadBBS_Announce Set RootID=(select t.rootid from(Select Max(RootID) as rootid from LeadBBS_Announce where ParentID=0 and BoardID=" & BoardID & " and RootID<" & DEF_BBS_TOPMinID & ") as t)+1" & morestr & " where ParentID=0 and RootIDBak=" & RootIDBak,1)
	 		case Else
				Set Rs = LDExeCute("Select Max(RootID) from LeadBBS_Topic where BoardID=" & BoardID & " and RootID<" & DEF_BBS_TOPMinID,0)
				If Rs.Eof Then
					MaxRootID = 1
				Else
					MaxRootID = Rs(0)
					If isNull(MaxRootID) or MaxRootID="" Then
						MaxRootID=1
					End If
					MaxRootID = cCur(MaxRootID)
				End If
				Rs.Close
				Set Rs = Nothing
			 	CALL LDExeCute("Update LeadBBS_Announce Set RootID=" & MaxRootID+1 & morestr & " where ID=" & RootIDBak,1)
			 	CALL LDExeCute("Update LeadBBS_Topic Set RootID=" & MaxRootID+1 & morestr & " where ID=" & RootIDBak,1)
	 	End select
	End If
	'UpdateBoardValue(BoardID)

End Function

Function UpdateBoardValue(BoardID)

	Dim Rs,SQL
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	Dim N,AllMinRootID,AllMaxRootID
	select case DEF_UsedDataBase
		case 0,2:
			Set Rs = LDExeCute("select Min(RootID) from LeadBBS_Announce where ParentID=0 and BoardID=" & BoardID,0)
		case Else
			Set Rs = LDExeCute("select Min(RootID) from LeadBBS_Topic where BoardID=" & BoardID,0)
	End select
	If Rs.Eof Then
		AllMinRootID = 0
	Else
		AllMinRootID = Rs(0)
		If isNull(AllMinRootID) Then AllMinRootID = 0
		AllMinRootID = cCur(AllMinRootID)
	End If
	Rs.Close
	Set Rs = Nothing

	select case DEF_UsedDataBase
		case 0,2:
			Set Rs = LDExeCute("select Max(RootID) from LeadBBS_Announce where ParentID=0 and BoardID=" & BoardID,0)
		case Else
			Set Rs = LDExeCute("select Max(RootID) from LeadBBS_Topic where BoardID=" & BoardID,0)
	End select
	If Rs.Eof Then
		AllMaxRootID = 0
	Else
		AllMaxRootID = Rs(0)
		If isNull(AllMaxRootID) Then AllMaxRootID = 0
		AllMaxRootID = cCur(AllMaxRootID)
	End If
	Rs.Close
	Set Rs = Nothing
	
	Dim LastAnnounceID,LastTopicName,LastWriter,LastTime
	select case DEF_UsedDataBase
		case 0,2:
			Set Rs = LDExeCute(sql_select("Select ID,title,LastUser,UserName,LastTime,TitleStyle from LeadBBS_Announce where ParentID = 0 and BoardID=" & BoardID & " and RootID<" & DEF_BBS_TOPMinID & " order By RootID DESC",1),0)
		case Else
			Set Rs = LDExeCute(sql_select("Select ID,title,LastUser,UserName,LastTime,TitleStyle from LeadBBS_Topic where BoardID=" & BoardID & " and RootID<" & DEF_BBS_TOPMinID & " order By RootID DESC",1),0)
	End select
	If Rs.Eof Then
		LastAnnounceID = 0
		LastTopicName = ""
		LastTime = 0
		LastWriter = ""
	Else
		LastAnnounceID = Rs(0)
		LastTopicName = Rs(1)
		LastWriter = Rs(2)
		If LastWriter = "" or isNull(LastWriter) Then LastWriter = Rs(3)
		LastTime = cCur(Rs(4))
		If isNull(LastTime) then LastTime = 0
		If Rs(5) = 1 Then LastTopicName = KillHTMLLabel(LastTopicName)
	End If
	Rs.Close
	Set Rs = Nothing
	CALL LDExeCute("Update LeadBBS_Boards Set AllMinRootID=" & AllMinRootID & ",AllMaxRootID=" & AllMaxRootID & ",LastAnnounceID=" & LastAnnounceID & ",LastTopicName='" & Replace(LastTopicName,"'","''") & "',LastWriter='" & Replace(LastWriter,"'","''") & "',LastWriteTime=" & LastTime & " where boardID=" & BoardID,1)
	ReloadBoardInfo(BoardID)
	If isArray(application(DEF_MasterCookies & "TopAnc")) Then ReloadTopAnnounceInfo(0)
	

End Function

Sub ReloadTopAnnounceInfo(TID)

	If isNumeric(TID) = 0 Then Exit Sub
	TID = Fix(cCur(TID))
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
		select case DEF_UsedDataBase
			case 0,2:
				Set Rs = LDExeCute(sql_select("select id,ChildNum,Title,FaceIcon,LastTime,Hits,Length,UserName,UserID,RootID,LastUser,NotReplay,GoodFlag,BoardID,TopicType,PollNum,TitleStyle,LastInfo,ndatetime,GoodAssort,NeedValue,NeedValue from LeadBBS_Announce where ParentID=0 and RootIDBak in(" & Temp & ")",Ubound(GetDataTop,2)+1),0)
			case Else
				Set Rs = LDExeCute(sql_select("select id,ChildNum,Title,FaceIcon,LastTime,Hits,Length,UserName,UserID,RootID,LastUser,NotReplay,GoodFlag,BoardID,TopicType,PollNum,TitleStyle,LastInfo,ndatetime,GoodAssort,NeedValue,NeedValue from LeadBBS_Topic where ID in(" & Temp & ")",Ubound(GetDataTop,2)+1),0)
		End select
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
%>