<%
Class Small_List

Private RootID

Sub DisplayAnnouncesSplit

	RootID = Left(Request.Form("id"),14)
	If isNumeric(RootID) = False Then RootID = 0
	RootID = Fix(cCur(RootID))
	
	Dim TopicSortID,Temp1
	Dim SQLEndString,Upflag,WhereFlag
	TopicSortID = Left(Request.Form("TID"),14)
	If isNumeric(TopicSortID) = 0 Then TopicSortID = 0
	TopicSortID = cCur(TopicSortID)
	If TopicSortID < 0 Then TopicSortID = 0
	Dim Rs,SQL
	Dim ALL_FirstTopicSortID,ALL_LastTopicSortID

	Dim ALL_Count
	SQL = sql_select("Select ParentID,ChildNum,RootMaxID,RootMinID,BoardID from LeadBBS_Announce where ID=" & RootID,1)
	Set Rs = LDExeCute(SQL,0)
	If Not Rs.Eof Then
		If cCur(Rs(0)) <> 0 Then
			Rs.Close
			Set Rs = Nothing
			Exit Sub
		End If
		ALL_Count = cCur(Rs(1))
		ALL_FirstTopicSortID = cCur(Rs(3))
		ALL_LastTopicSortID = cCur(Rs(2))
		GBL_Board_ID = cCur(Rs(4))
	Else
		ALL_Count = 0
		Rs.Close
		Set Rs = Nothing
		Exit Sub
	End If
	Rs.Close
	Set Rs = Nothing


	Dim TArray,ForumPass,BoardLimit,OtherLimit,HiddenFlag
	TArray = Application(DEF_MasterCookies & "BoardInfo" & GBL_Board_ID)
	If isArray(TArray) = False Then
		ReloadBoardInfo(GBL_Board_ID)
		TArray = Application(DEF_MasterCookies & "BoardInfo" & GBL_Board_ID)
	End If
	If isArray(TArray) = False Then
			Exit Sub
	Else
		ForumPass = TArray(7,0)
		BoardLimit = cCur(TArray(9,0))
		OtherLimit = cCur(TArray(36,0))
		HiddenFlag = TArray(8,0)
	End If
	If GBL_CheckLimitTitle(ForumPass,BoardLimit,OtherLimit,HiddenFlag) = 1 Then
		Response.Write "<ul><li>限制版面已禁用此功能.</li></ul>"
		Exit Sub
	End If

	WhereFlag = 0

	Upflag = Request.Form("Upflag")
	If Upflag<>"1" and Upflag<>"0" Then Upflag="0"
	If Upflag="0" and TopicSortID>=ALL_LastTopicSortID Then TopicSortID = ALL_LastTopicSortID-1
	If Upflag="1" and TopicSortID<=ALL_FirstTopicSortID Then TopicSortID = ALL_FirstTopicSortID+1
	Dim HaveTopicSortIDFlag
	Dim FirstTopicSortID,LastTopicSortID
	Dim GetData
	
	If RootID<>0 Then
		If ALL_Count<DEF_MaxListNum Then
			If DEF_EnableTreeView = 1 then
				SQLEndString = " Where boardid=" & GBL_board_ID & " and RootIDBak=" & RootID & " order by ID ASC"
			Else
				SQLEndString = " Where RootIDBak=" & RootID & " order by ID ASC"
			End If
		Else
			If Upflag="0" Then
				If WhereFlag = 1 Then
					SQLEndString = " where boardid=" & GBL_board_ID
					WhereFlag = 1
					SQLEndString = SQLEndString & " and (RootIDBak=" & RootID & " and ID>" & TopicSortID & ")"
				Else
					SQLEndString = SQLEndString & " Where (RootIDBak=" & RootID & " and ID>" & TopicSortID & ")"
					WhereFlag = 1
				End If
			Else
				If WhereFlag = 1 Then
					SQLEndString = " where boardid=" & GBL_board_ID
					WhereFlag = 1
					SQLEndString = SQLEndString & " and (RootIDBak=" & RootID & " and ID<" & TopicSortID & ")"
				Else
					SQLEndString = SQLEndString & " where (RootIDBak=" & RootID & " and ID<" & TopicSortID & ")"
					WhereFlag = 1
				End If
			End If

			If Upflag="0" Then
				SQLEndString = SQLEndString & " order by ID ASC"
			Else
				SQLEndString = SQLEndString & " order by ID DESC"
			End If
		End If
		SQL = sql_select("Select id,Title,ndatetime,UserName,UserID,TitleStyle from LeadBBS_Announce " & SQLEndString,DEF_MaxListNum)
		Set Rs = LDExeCute(SQL,0)
		If Not Rs.Eof Then
			HaveTopicSortIDFlag = 1
			GetData = Rs.GetRows(-1)
		Else
			HaveTopicSortIDFlag = 0
		End If
		Rs.Close
		Set Rs = Nothing
	Else
		HaveTopicSortIDFlag = 0
	End If

	If HaveTopicSortIDFlag = 1 Then
		Temp1 = Ubound(GetData,2)
		FirstTopicSortID = cCur(GetData(0,0))
		LastTopicSortID = cCur(GetData(0,Temp1))

		If FirstTopicSortID>LastTopicSortID Then
			SQL = LastTopicSortID
			LastTopicSortID = FirstTopicSortID
			FirstTopicSortID = SQL
		End If
	Else
		Temp1 = 0
	End If

	If ALL_Count>=DEF_MaxListNum Then
	SQL = "&B=" & GBL_board_ID
	If RootID <> 0 Then SQL = SQL & "&ID=" & RootID
	Dim PageSplitString
	PageSplitString = "<ul><li>"
	If FirstTopicSortID<=ALL_FirstTopicSortID Then
		PageSplitString = PageSplitString & "首页"
		PageSplitString = PageSplitString & " 上页"
	Else
		PageSplitString = PageSplitString & "<a href=#no onclick='getAJAX(""b.asp"",""ol=3" & SQL & """,""Lead" & RootID & """);'>首页</a>"
		PageSplitString = PageSplitString & " <a href=#no onclick='getAJAX(""b.asp"",""ol=3" & SQL & "&RootID=" & RootID & "&TID=" & FirstTopicSortID & "&Upflag=1"",""Lead" & RootID & """);'>上页</a>"
	End If

	If LastTopicSortID>=ALL_LastTopicSortID Then
		PageSplitString = PageSplitString & " 下页"
		PageSplitString = PageSplitString & " 尾页"
	Else
		PageSplitString = PageSplitString & " <a href=#no onclick='getAJAX(""b.asp"",""ol=3" & SQL & "&RootID=" & RootID & "&TID=" & LastTopicSortID & """,""Lead" & RootID & """);'>下页</a>"
		PageSplitString = PageSplitString & " <a href=#no onclick='getAJAX(""b.asp"",""ol=3" & SQL & "&Upflag=1&TID=" & ALL_LastTopicSortID+1 & """,""Lead" & RootID & """);'>尾页</a>"
	End If
	Rs = Temp1
	If HaveTopicSortIDFlag = 1 Then Rs = Rs+1
	PageSplitString = PageSplitString & " 相关帖共有<b>" & ALL_Count &"</b>帖 此页<b>" & Rs & "</b>帖 每页<b>" & DEF_MaxListNum & "</b>帖 执行时间" & fix(abs(CDBL(Timer)*1000 - DEF_PageExeTime1*1000)) & "毫秒"
	PageSplitString = PageSplitString & "</li></ul>"
	End If

	Dim For1,For2,StepValue
	If HaveTopicSortIDFlag = 1 Then
		Response.Write "<ul>"
		If Upflag="0" Then
			For1 = 0
			For2 = Temp1
			StepValue = 1
		Else
			For1 = Temp1
			For2 = 0
			StepValue = -1
		End If
		DisplaySmallAnnounceData For1,For2,StepValue,GetData
		Response.Write "</ul>" & VbCrLf
		Response.Write PageSplitString & VbCrLf
	End If

	Set Rs = Nothing

End Sub

Private Sub DisplaySmallAnnounceData(For1,For2,StepValue,GetData)

	Dim N,Temp
	For N = For1 to For2 Step StepValue
		If RootID <> cCur(GetData(0,N)) Then
			Response.Write "<li>" & VbCrLf	
			If N = For2 Then
				Response.Write "└" & VbCrLf
			Else
				Response.Write "├" & VbCrLf
			End If
			Response.Write "<a href=" & DEF_BBS_HomeUrl & "a/a.asp?B=" & GBL_Board_ID & "&ID=" & GetData(0,N) & "&re=1>" & VbCrLf
		End If
		If GetData(5,n) <> 1 and Len(GetData(1,N))>DEF_BBS_DisplayTopicLength Then GetData(1,N) = Left(GetData(1,N),DEF_BBS_DisplayTopicLength-3) & "..."
		If GetData(5,n) <> 1 Then GetData(1,n) = Replace(GetData(1,n) & "","<","&lt;")
		GetData(1,n) = DisplayAnnounceTitle(GetData(1,n),GetData(5,n))
		If GetData(5,n) >=60 Then
			GetData(1,n) = "帖子等待审核中..."
			GetData(5,n) = 1
		End If
		GetData(1,n) = Replace(Replace(GetData(1,N) & "","\","\\"),"""","\""")
		If Left(GetData(1,n),3) = "re:" Then GetData(1,n) = Mid(GetData(1,n),4)
		Response.Write GetData(1,n) & VbCrLf
		If RootID <> cCur(GetData(0,N)) Then
			Response.Write "</a>" & VbCrLf
		End If
		If cCur(GetData(4,N)) > 0 Then
			Response.Write " [<a href=" & DEF_BBS_HomeUrl & "User/LookUserInfo.asp?ID=" & GetData(4,N) & ">" & GetData(3,N) & "</a> " & VbCrLf
		Else
			Response.Write " [" & GetData(3,N) & " " & VbCrLf
		End If
		Temp = RestoreTime(GetData(2,N))
		If DateDiff("d",Temp,DEF_Now)<1 Then
			Response.Write " <font color=Red class=redfont>" & VbCrLf
			Response.Write Temp & "</font>]</li>" & VbCrLf
		Else
			Response.Write Temp & "]</li>" & VbCrLf
		End If
	Next

End Sub

End Class%>