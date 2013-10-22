<%
Class Mini_Board

	Public Sub List

		If GBL_Board_ID < 1 or GBL_CheckLimitContent(GBL_Board_ForumPass,GBL_Board_BoardLimit,GBL_Board_OtherLimit,GBL_Board_HiddenFlag) = 1 Then
			Response.Write "<ul>查询错误，可能原因：<ol><li>版面限制查看。</li><li>需要的版面无法查找。</li></ol></ul>"
			Exit Sub
		End If
		Dim Rs,SQL,Temp
		Dim SQLEndString,WhereFlag
		WhereFlag = 1
		select case DEF_UsedDataBase
			case 0,2:
				SQLEndString = " where ta.ParentID=0 and ta.boardid=" & GBL_board_ID
			case Else
				SQLEndString = " where ta.boardid=" & GBL_board_ID
		End select
	
		Dim ALL_Count
		ALL_Count = GBL_Board_TopicNum
	
		Dim RootID,Temp1,Temp2
		Dim Upflag
	
		RootID = M_Par.R
	
		Dim LastNum,LastNumBak
		LastNum = 0

		LastNumBak = (ALL_Count mod DEF_MaxListNum)
		If LastNumBak = 0 Then LastNumBak = DEF_MaxListNum
	
		Upflag = M_Par.UpFlag
		If Upflag = 1 Then
			LastNum = M_Par.Num
			If LastNum <> 0 Then
				LastNum = LastNumBak
			End If
		End If
	
		Dim Page,MaxPage,JMPage
		Page = M_Par.P
		MaxPage = Fix(All_Count / DEF_MaxListNum)
		If (All_Count mod DEF_MaxListNum)<>0 Then MaxPage = MaxPage + 1
		MaxPage = MaxPage - 1
		If Page > MaxPage or LastNum > 0 Then
			Page = MaxPage
		End If
	
		JMPage = M_Par.q
		If JMPage > DEF_MaxJumpPageNum Then JMPage = 0
		
		Dim JMPRootID
		JMPRootID = M_Par.r
		
		If JMPage > Maxpage or Maxpage < 0 Then JMPage = 0
		If Upflag="0" and JMPage+Page > MaxPage Then JMPage = 0
		If Upflag="1" and JMPage+Page < 0 Then JMPage = 0
		If JMPRootID > GBL_Board_AllMaxRootID Then JMPage = 0
		If JMPRootID < GBL_Board_AllMinRootID Then JMPage = 0
	
		If Upflag="1" Then
			Page = Page - JMPage
		Else
			Page = Page + JMPage
		End If
	
		If Page = 0 Then '开启此项则当页数为0时即忽略一切信息的返回首页
			JMPage = 0
			JMPRootID = 0
			LastNum = 0
			Upflag = "0"
			RootID = 0
		End If
	
		Dim HaveRootIDFlag
		Dim FirstRootID,LastRootID
	
		If Temp1+1<DEF_MaxListNum and All_Count > 0 Then
			If Upflag="0" Then
				If JMPage > 0 Then
					If JMPRootID<>0 Then
						If WhereFlag = 1 Then
							SQLEndString = SQLEndString & " And ta.RootID<" & JMPRootID
						Else
							SQLEndString = SQLEndString & " Where ta.RootID<" & JMPRootID
							WhereFlag = 1
						End If
					End If
				Else
					If RootID<>0 Then
						If WhereFlag = 1 Then
							SQLEndString = SQLEndString & " And ta.RootID<" & RootID
						Else
							SQLEndString = SQLEndString & " Where ta.RootID<" & RootID
							WhereFlag = 1
						End If
					End If
				End If
			Else
				If JMPage > 0 Then
					If WhereFlag = 1 Then
						SQLEndString = SQLEndString & " And ta.RootID>" & JMPRootID
					Else
						SQLEndString = SQLEndString & " Where ta.RootID>" & JMPRootID
						WhereFlag = 1
					End If
				Else
					If RootID>=GBL_Board_AllMaxRootID Then
						RootID = GBL_Board_AllMaxRootID-1
						'TopicSortID = GBL_Board_AllMaxRootID + 1
					
						If RootID<>0 Then
							If WhereFlag = 1 Then
								SQLEndString = SQLEndString & " And ta.RootID>" & RootID
							Else
								SQLEndString = SQLEndString & " Where ta.RootID>" & RootID
								WhereFlag = 1
							End If
						End If
					Else
						If RootID<>0 Then
							If WhereFlag = 1 Then
								SQLEndString = SQLEndString & " And ta.RootID>" & RootID
							Else
								SQLEndString = SQLEndString & " Where ta.RootID>" & RootID
								WhereFlag = 1
							End If
						End If
					End If
				End If
			End If
	
			Dim NoPage
			NoPage = 0
	
			If Page < 0 or (Page > MaxPage and MaxPage>=(DEF_MaxJumpPageNum-1)) or (Page > (DEF_MaxJumpPageNum-1) and Page<(MaxPage-DEF_MaxJumpPageNum+1)) Then NoPage = 1
			If (RootID > 0 or LastNum>0 or NoPage = 1) and JMPage < 1 Then
				If Upflag="0" Then
					SQLEndString = SQLEndString & " order by ta.RootID DESC"
				Else
					SQLEndString = SQLEndString & " order by ta.RootID ASC"
				End If
				If LastNum > 0 Then
					Temp = LastNum
				Else
					Temp = DEF_MaxListNum
				End If
			Else
				If JMPage > 0 Then
					If Upflag="0" Then
						SQLEndString = SQLEndString & " order by ta.RootID DESC"
						Upflag="0"
						Temp = (JMPage-1) * DEF_MaxListNum + DEF_MaxListNum
					Else
						SQLEndString = SQLEndString & " order by ta.RootID ASC"
						Upflag="1"
						Temp = (JMPage-1) * DEF_MaxListNum + DEF_MaxListNum
					End If
				Else
					If Page < DEF_MaxJumpPageNum Then
						SQLEndString = SQLEndString & " order by ta.RootID DESC"
						Upflag="0"
						Temp = Page * DEF_MaxListNum + DEF_MaxListNum
					Else
						SQLEndString = SQLEndString & " order by ta.RootID ASC"
						Upflag="1"
						Temp = (MaxPage-Page) * DEF_MaxListNum + LastNumBak
					End If
				End If
			End If
	
			Dim FirstRootID_2,LastRootID_2
			Dim GetData_2,movenum

'新代码开始
If (DEF_UsedDataBase = 0 or DEF_UsedDatabase) and Temp>1000 Then

	select case DEF_UsedDataBase
		case 0:
			SQL = sql_select("select ta.RootID from LeadBBS_Announce as ta " & SQLEndString,Temp)
			Set Rs = LDExeCute(SQL,0)
	
			If LastNum = "" or isNull(LastNum) then LastNum = 0
			LastNum = cCur(LastNum)
			If (RootID = 0 and Page >= 1 and LastNum = 0 and NoPage = 0) or JMPage > 0 Then
				If Not Rs.Eof Then
					If JMPage > 0 Then
						If Upflag="0" Then
							Rs.Move (JMPage-1)* DEF_MaxListNum
						Else
							If Page < MaxPage Then
								Rs.Move (JMPage-2) * DEF_MaxListNum + DEF_MaxListNum
							End If
						End If
					Else
						If Page < DEF_MaxJumpPageNum Then
							Rs.Move Page * DEF_MaxListNum
						Else
							If Page < MaxPage Then
								Rs.Move (MaxPage-Page-1) * DEF_MaxListNum + LastNumBak
							End If
						End If
					End If
				End If
			End If
		case 2:
			movenum = 0
	
			If LastNum = "" or isNull(LastNum) then LastNum = 0
			LastNum = cCur(LastNum)
			If (RootID = 0 and Page >= 1 and LastNum = 0 and NoPage = 0) or JMPage > 0 Then
					If JMPage > 0 Then
						If Upflag="0" Then
							movenum = (JMPage-1)* DEF_MaxListNum
						Else
							If Page < MaxPage Then
								movenum = (JMPage-2) * DEF_MaxListNum + DEF_MaxListNum
							End If
						End If
					Else
						If Page < DEF_MaxJumpPageNum Then
							movenum = Page * DEF_MaxListNum
						Else
							If Page < MaxPage Then
								movenum = (MaxPage-Page-1) * DEF_MaxListNum + LastNumBak
							End If
						End If
					End If
			End If
			SQL = sql_select("select ta.RootID from LeadBBS_Announce as ta " & SQLEndString,movenum & "," & Temp)
			Set Rs = LDExeCute(SQL,0)
		end select

	Dim Cur_RootID,ID
	If Not Rs.Eof Then
		Cur_RootID = Rs(0)
	Else
		Cur_RootID = 0
	End If
	Rs.Close
	Set Rs = Nothing
	Dim SQLEndString_J
	SQLEndString_J = Replace(SQLEndString,">" & JMPRootID & " ",">=" & Cur_RootID & " ")
	SQLEndString_J = Replace(SQLEndString_J,"<" & JMPRootID & " ","<=" & Cur_RootID & " ")
	SQLEndString_J = Replace(SQLEndString_J,">" & RootID & " ",">=" & Cur_RootID & " ")
	SQLEndString_J = Replace(SQLEndString_J,"<" & RootID & " ","<=" & Cur_RootID & " ")
	SQL = sql_select("select ta.id,ta.ChildNum,ta.Title,ta.UserName,ta.TitleStyle,ta.RootID,tb.ForumPass,tb.BoardLimit,tb.OtherLimit,tb.HiddenFlag from LeadBBS_Announce as ta Left Join LeadBBS_Boards as tb on ta.BoardID=tb.BoardID " & SQLEndString_J,DEF_MaxListNum)
	Set Rs = LDExeCute(SQL,0)
Else
'新代码结束
	select case DEF_UsedDataBase
		case 0,1:	
			select case DEF_UsedDataBase
				case 0,2:
					SQL = sql_select("select ta.id,ta.ChildNum,ta.Title,ta.UserName,ta.TitleStyle,ta.RootID,tb.ForumPass,tb.BoardLimit,tb.OtherLimit,tb.HiddenFlag from LeadBBS_Announce as ta Left Join LeadBBS_Boards as tb on ta.BoardID=tb.BoardID " & SQLEndString,Temp)
				case Else
					SQL = sql_select("select ta.id,ta.ChildNum,ta.Title,ta.UserName,ta.TitleStyle,ta.RootID,tb.ForumPass,tb.BoardLimit,tb.OtherLimit,tb.HiddenFlag from LeadBBS_Topic as ta Left Join LeadBBS_Boards as tb on ta.BoardID=tb.BoardID " & SQLEndString,Temp)
			End select
			Set Rs = LDExeCute(SQL,0)
	
			If LastNum = "" or isNull(LastNum) then LastNum = 0
			LastNum = cCur(LastNum)
			If (RootID = 0 and Page >= 1 and LastNum = 0 and NoPage = 0) or JMPage > 0 Then
				If Not Rs.Eof Then
					If JMPage > 0 Then
						If Upflag="0" Then
							Rs.Move (JMPage-1)* DEF_MaxListNum
						Else
							If Page < MaxPage Then
								Rs.Move (JMPage-2) * DEF_MaxListNum + DEF_MaxListNum
							End If
						End If
					Else
						If Page < DEF_MaxJumpPageNum Then
							Rs.Move Page * DEF_MaxListNum
						Else
							If Page < MaxPage Then
								Rs.Move (MaxPage-Page-1) * DEF_MaxListNum + LastNumBak
							End If
						End If
					End If
				End If
			End If
		case 2:
			movenum = 0	
			If LastNum = "" or isNull(LastNum) then LastNum = 0
			LastNum = cCur(LastNum)
			If (RootID = 0 and Page >= 1 and LastNum = 0 and NoPage = 0) or JMPage > 0 Then
					If JMPage > 0 Then
						If Upflag="0" Then
							movenum = (JMPage-1)* DEF_MaxListNum
						Else
							If Page < MaxPage Then
								movenum = (JMPage-2) * DEF_MaxListNum + DEF_MaxListNum
							End If
						End If
					Else
						If Page < DEF_MaxJumpPageNum Then
							movenum = Page * DEF_MaxListNum
						Else
							If Page < MaxPage Then
								movenum = (MaxPage-Page-1) * DEF_MaxListNum + LastNumBak
							End If
						End If
					End If
			End If
			SQL = sql_select("select ta.id,ta.ChildNum,ta.Title,ta.UserName,ta.TitleStyle,ta.RootID,tb.ForumPass,tb.BoardLimit,tb.OtherLimit,tb.HiddenFlag from LeadBBS_Announce as ta Left Join LeadBBS_Boards as tb on ta.BoardID=tb.BoardID " & SQLEndString,movenum & "," & Temp)
			Set Rs = LDExeCute(SQL,0)
		end select
'新代码开始
End If
'新代码结束
			If Not Rs.Eof Then
				HaveRootIDFlag = 1
				GetData_2 = Rs.GetRows(DEF_MaxListNum)
			Else
				HaveRootIDFlag = 0
			End If
			Rs.Close
			Set Rs = Nothing
		
			If HaveRootIDFlag = 1 Then
				Temp2 = Ubound(GetData_2,2)
				FirstRootID_2 = cCur(GetData_2(5,0))
				LastRootID_2 = cCur(GetData_2(5,Temp2))
					
				If FirstRootID_2<LastRootID_2 Then
					SQL = FirstRootID_2
					FirstRootID_2 = LastRootID_2
					LastRootID_2 = SQL
				End If
		
				LastRootID = LastRootID_2
				FirstRootID = FirstRootID_2
			Else
				Temp2 = 0
			End If
		Else
			HaveRootIDFlag = 0
		End If
		
		Dim N
	
		Dim PageSplitString,PageSplitString2
		PageSplitString = "<tr height=25 class=TBBG9><td colspan=5><table width=100% border=0 cellspacing=0 cellpadding=0><tr><td><img src=../images/null.gif width=2 height=2><br><img src=../images/null.gif width=2 height=2>"
		If FirstRootID >= GBL_Board_AllMaxRootID Then
			PageSplitString = PageSplitString & "<font color=888888 class=grayfont face=webdings title=首页>9</font>"
			PageSplitString = PageSplitString & "<font color=888888 class=grayfont face=webdings title=上页>7</font>"
		Else
			PageSplitString = PageSplitString & "<a href=Default.asp?" & M_Par.GetPar(GBL_Board_ID,0,0,0,0,0,0,"b") & "><font face=webdings title=首页>9</font></a>"
			PageSplitString = PageSplitString & "<a href=Default.asp?" & M_Par.GetPar(GBL_Board_ID,0,FirstRootID,1,Page-1,0,0,"b") & "><font face=webdings title=上页>7</font></a>"
		End If
		
		Dim DN
		DN = DEF_DisplayJumpPageNum
		Dim For1,For2,StepValue,DotFlag
		DotFlag = 0
		PageSplitString = PageSplitString & " "
	
		If MaxPage > 0 Then
			For1 = Page - DN
			For2 = Page + DN
			If For1 < 0 Then
				For1 = 0
			End If
			If For2 >= MaxPage Then For2 = MaxPage
			If For2 > MaxPage Then For2 = MaxPage
			If For2 - For1 < DEF_DisplayJumpPageNum*2 and For1 > 1 and For2 > For1 Then For1 = For1 - (DEF_DisplayJumpPageNum*2 - ((For2 - For1) + 1))
			
			If For1 < 0 Then
				For1 = 0
			ElseIf For1 > 0 Then
				PageSplitString = PageSplitString & " ..."
			End If
	
			If For2 - For1 < DEF_DisplayJumpPageNum*2 and For2 < MaxPage and For2 > For1 Then For2 = For2 + (DEF_DisplayJumpPageNum*2 - ((For2 - For1) + 1))
			If For2 >= MaxPage Then For2 = MaxPage
	
			If For2 > 998 Then
				If Page > For1 Then For1 = For1 + 1
				If Page > For1 Then For1 = For1 + 1
				If Page < For2 - DN + 1 Then For2 = For2 - 1
			End If
			'M_Par.BoardID,M_Par.ID,M_Par.R,M_Par.UpFlag,M_Par.p,M_Par.q,M_Par.Num,"h"
			For N = For1 to For2
				If N <> For1 Then PageSplitString = PageSplitString & " "
				If N = Page Then
					PageSplitString = PageSplitString & "<font color=Red class=redfont>" & N + 1 & "</font>"
				Else
					If (N-Page) > 0 Then
						PageSplitString = PageSplitString & "<a href=Default.asp?" & M_Par.GetPar(GBL_Board_ID,0,LastRootID,0,Page,N-Page,0,"b") & ">" & N + 1 & "</a>"
					Else
						PageSplitString = PageSplitString & "<a href=Default.asp?" & M_Par.GetPar(GBL_Board_ID,0,FirstRootID,1,Page,Page-N,0,"b") & ">" & N + 1 & "</a>"
					End If
				End If
				DotFlag = 2
			Next
			If For2 < MaxPage Then
				PageSplitString = PageSplitString & "..."
				DotFlag = 1
			End If
		Else
			PageSplitString = PageSplitString & " <font color=888888 class=grayfont>1</font> "
		End If
	
		PageSplitString = PageSplitString & " "
	
		If LastRootID <= GBL_Board_AllMinRootID Then
			PageSplitString = PageSplitString & " <font color=888888 class=grayfont face=webdings title=下页>8</font>"
			PageSplitString = PageSplitString & "<font color=888888 class=grayfont face=webdings title=尾页>:</font>"
		Else
			PageSplitString = PageSplitString & " <a href=Default.asp?" & M_Par.GetPar(GBL_Board_ID,0,LastRootID,0,Page+1,0,0,"b") & "><font face=webdings title=下页>8</font></a>"
			PageSplitString = PageSplitString & "<a href=Default.asp?" & M_Par.GetPar(GBL_Board_ID,0,0,1,MaxPage,0,0,"b") & "><font face=webdings title=尾页>:</font></a>"
		End If
	
		Rs = Temp1
		Rs = Temp2+Rs
		If HaveRootIDFlag = 1 Then Rs = Rs+1
	
		if MaxPage < 1 Then MaxPage = 0
		'if MaxPage > 0 Then MaxPage = MaxPage + 1
		PageSplitString = PageSplitString & " 共" & ALL_Count &"主题 第" & Page+1 & "/" & MaxPage+1 & "页 每页" & DEF_MaxListNum & "条<td align=right><img src=../images/null.gif width=2 height=2><br>"
		PageSplitString2 = "</td></tr></table></td></tr>"
		For1 = 0
		For2 = 0
		If Rs < DEF_MaxListNum and GBL_ShowBottomSure = 0 Then GBL_SiteBottomString = ""
		
		Response.Write PageSplitString
		Response.Write PageSplitString2
		
		If Upflag="0" Then
			If HaveRootIDFlag = 1 Then
				If Upflag="0" Then
					For1 = 0
					For2 = Temp2
					StepValue = 1
				Else
					For1 = Temp2
					For2 = 0
					StepValue = -1
				End If
				DisplayAnnounceData For1,For2,StepValue,GetData_2
			End If
		Else
	
			If HaveRootIDFlag = 1 Then
				If Upflag="0" Then
					For1 = 0
					For2 = Temp2
					StepValue = 1
				Else
					For1 = Temp2
					For2 = 0
					StepValue = -1
				End If
				DisplayAnnounceData For1,For2,StepValue,GetData_2
			End If
		End If
		Response.Write PageSplitString
		Response.Write PageSplitString2
	
	End Sub
	
	Private Sub DisplayAnnounceData(For1,For2,StepValue,GetData)
	
		'id,ChildNum,Title,UserName,TitleStyle
		Response.Write "<ol>" & VbCrLf
		Dim N,Temp,Temp1
		For N = For1 to For2 Step StepValue
			If GBL_CheckLimitTitle(GetData(6,n),GetData(7,n),GetData(8,n),GetData(9,n)) = 1 Then
				GetData(2,n) = "此帖子标题已设置为隐藏"
				GetData(3,N) = ""
			Else
				GetData(3,N) = GetData(3,N) & ","
				If GetData(4,n) = 1 Then
					GetData(2,n) = KillHTMLLabel(GetData(2,n))
				Else
					GetData(2,n) = HtmlEncode(GetData(2,n))
				End If
			End If
			If GetData(4,n) >=60 Then
				GetData(2,n) = "<font color=gray class=grayfont>帖子等待审核中...</font>"
				GetData(4,n) = 1
			End If
			Response.Write "<li><a href=Default.asp?" & M_Par.GetPar(M_Par.BoardID,GetData(0,N),0,0,0,0,0,"a") & ">"
			Response.Write GetData(2,n)
			Response.Write "</a> (" & HtmlEncode(GetData(3,N)) & GetData(1,N) & "回复)</li>" & VbCrLf
		Next
		Response.Write "</ol>" & VbCrLf
	
	End Sub

End Class
%>