<%
Class Mini_Announce

	Public LMT_TopicName
	Private Form_ParentID,Form_RootMaxID,Form_RootMinID,Form_ChildNum,LMT_RootIDBak,page
	Private Flag

	Public Sub GetTopicInfo
	
		Form_ParentID = M_Par.ID
		If Form_ParentID = 0 Then Exit Sub
		Dim Rs,SQL,Form_TopicType,Form_NeedValue,LMT_TopicTitleStyle
		Flag = 0
	
		SQL = sql_select("Select ta.Title,ta.RootMaxID,ta.RootMinID,ta.Hits,ta.ChildNum,ta.ID,ta.TitleStyle,ta.ParentID,tb.ForumPass,tb.BoardLimit,tb.OtherLimit,tb.HiddenFlag from LeadBBS_Announce as TA left Join LeadBBS_Boards as TB on ta.boardid=tb.boardid where ta.ID=" & Form_ParentID,1)
		Set Rs = LDExeCute(SQL,0)
		If Rs.Eof Then
			Rs.Close
			Set Rs = Nothing
			Flag = 1
			Exit Sub
		Else
			LMT_TopicName = Rs(0)
			Form_RootMaxID = cCur(Rs(1))
			Form_RootMinID = cCur(Rs(2))
			Form_ChildNum = cCur(Rs(4))
			LMT_RootIDBak = cCur(Rs(5))
			LMT_TopicTitleStyle = Rs(6)
			If cCur(Rs(7)) > 0 Then Flag = 1
			If GBL_CheckLimitContent(Rs(8),Rs(9),Rs(10),Rs(11)) = 1 Then Flag = 1
			Rs.Close
			Set Rs = Nothing
			
			If LMT_TopicTitleStyle = 1 Then
				LMT_TopicName = KillHTMLLabel(LMT_TopicName)
			Else
				LMT_TopicName = HtmlEncode(LMT_TopicName)
			End If
		End If
	
	End Sub
	
	Public Sub DisplayTopic
	
		If Flag = 1 Then
			Response.Write "<ul>查询错误，可能原因：<ol><li>主题限制查看。</li><li>需要的帖子无法查找。</li></ol></ul>"
			Exit Sub
		End If
		Dim Rs,SQL
		Dim ALL_FirstID,ALL_LastID
		ALL_FirstID = Form_RootMaxID
		ALL_LastID = Form_RootMinID
	
		Dim ALL_Count
		ALL_Count = Form_ChildNum + 1
	
		Dim LMT_First,Temp1,Temp2
		Dim SQLEndString,Upflag,WhereFlag
		WhereFlag = 0
	
		LMT_First = M_Par.AFirst
		
		Dim LastNum,LastNumBak
		LastNum = 0
	
		LastNumBak = (ALL_Count mod DEF_TopicContentMaxListNum)
		If LastNumBak = 0 Then LastNumBak = DEF_TopicContentMaxListNum
		
		Upflag = M_Par.UpFlag
		If Upflag = 1 Then
			LastNum = M_Par.Num
			If LastNum <> 0 Then
				LastNum = LastNumBak
			End If
		End If
	
		Dim MaxPage
		Page = M_Par.P
		MaxPage = Fix(All_Count / DEF_TopicContentMaxListNum)
		If (All_Count mod DEF_TopicContentMaxListNum)<>0 Then MaxPage = MaxPage + 1
		MaxPage = MaxPage - 1
		If Page > MaxPage or LastNum > 0 Then
			Page = MaxPage
		End If
	
		Dim JMPage
		JMPage = M_Par.q
		If JMPage > DEF_MaxJumpPageNum Then JMPage = 0
		
		Dim JMPRootID
		JMPRootID = M_Par.r
		
		If JMPage > Maxpage or Maxpage < 0 Then JMPage = 0
		If Upflag="0" and JMPage+Page > MaxPage Then JMPage = 0
		If Upflag="1" and JMPage+Page < 0 Then JMPage = 0
		If JMPRootID > ALL_FirstID Then JMPage = 0
		If JMPRootID < ALL_LastID Then JMPage = 0
		
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
			LMT_First = 0
		End If
	
		Dim HaveIDFlag
		Dim LastID,FirstID
	
		If Temp1+1<DEF_TopicContentMaxListNum Then
	
			SQLEndString = " where T1.RootIDBak=" & LMT_RootIDBak
			WhereFlag = 1
			If Upflag="0" Then
				If JMPage > 0 Then
					If JMPRootID<>0 Then
						If WhereFlag = 1 Then
							SQLEndString = SQLEndString & " And T1.ID>" & JMPRootID
						Else
							SQLEndString = SQLEndString & " Where T1.ID>" & JMPRootID
							WhereFlag = 1
						End If
					End If
				Else
					If LMT_First<>0 Then
						If WhereFlag = 1 Then
							SQLEndString = SQLEndString & " And T1.ID>" & LMT_First
						Else
							SQLEndString = SQLEndString & " Where T1.ID>" & LMT_First
							WhereFlag = 1
						End If
					End If
				End If
			Else
				If JMPage > 0 Then
					If WhereFlag = 1 Then
						SQLEndString = SQLEndString & " And T1.ID<" & JMPRootID
					Else
						SQLEndString = SQLEndString & " Where T1.ID<" & JMPRootID
						WhereFlag = 1
					End If
				Else
					If LMT_First>=ALL_FirstID Then
						If LMT_First<>0 Then
							If WhereFlag = 1 Then
								SQLEndString = SQLEndString & " And T1.ID<" & LMT_First
							Else
								SQLEndString = SQLEndString & " Where T1.ID<" & LMT_First
								WhereFlag = 1
							End If
						End If
					Else
						If LMT_First<>0 Then
							If WhereFlag = 1 Then
								SQLEndString = SQLEndString & " And T1.ID<" & LMT_First
							Else
								SQLEndString = SQLEndString & " Where T1.ID<" & LMT_First
								WhereFlag = 1
							End If
						End If
					End If
				End If
			End If
	
			Dim NoPage
			NoPage = 0
			
			If Page < 0 or (Page > MaxPage and MaxPage>=(DEF_MaxJumpPageNum-1)) or (Page > (DEF_MaxJumpPageNum-1) and Page<(MaxPage-DEF_MaxJumpPageNum+1)) Then NoPage = 1
			If (LMT_First > 0 or LastNum>0 or NoPage = 1) and JMPage < 1 Then
				If Upflag="0" Then
					SQLEndString = SQLEndString & " order by T1.ID ASC"
				Else
					SQLEndString = SQLEndString & " order by T1.ID DESC"
				End If
				If LastNum > 0 Then
					SQL = LastNum
				Else
					SQL = DEF_TopicContentMaxListNum
				End If
			Else
				If JMPage > 0 Then
					If Upflag="0" Then
						SQLEndString = SQLEndString & " order by T1.ID ASC"
						Upflag="0"
						SQL = (JMPage-1) * DEF_TopicContentMaxListNum + DEF_TopicContentMaxListNum
					Else
						SQLEndString = SQLEndString & " order by T1.ID DESC"
						Upflag="1"
						SQL = (JMPage-1) * DEF_TopicContentMaxListNum + DEF_TopicContentMaxListNum
					End If
				Else
					If Page < DEF_MaxJumpPageNum Then
						SQLEndString = SQLEndString & " order by T1.ID ASC"
						Upflag="0"
						SQL = Page * DEF_TopicContentMaxListNum + DEF_TopicContentMaxListNum
					Else
						SQLEndString = SQLEndString & " order by T1.ID DESC"
						Upflag="1"
						SQL = (MaxPage-Page) * DEF_TopicContentMaxListNum + LastNumBak
					End If
				End If
			End If
	
			Dim FirstID_2,LastID_2
			Dim GetData_2,MoveNum

'新代码开始
If (DEF_UsedDataBase = 0 or DEF_UsedDatabase = 2) and SQL>1000 Then
		Dim TmpSQL
	select case DEF_UsedDataBase
	case 0:
		TmpSQL = sql_select("Select T1.ID from LeadBBS_Announce as T1 " & SQLEndString,sql)
		Set Rs = LDExeCute(TmpSQL,0)
		If (LMT_First = 0 and LastNum = 0 and Page >= 1 and NoPage = 0) or JMPage > 0 Then
			If Not Rs.Eof Then
				If JMPage > 0 Then
					If Upflag="0" Then
						Rs.Move (JMPage-1)* DEF_TopicContentMaxListNum
					Else
						If Page < MaxPage Then
							Rs.Move (JMPage-2) * DEF_TopicContentMaxListNum + DEF_TopicContentMaxListNum
						End If
					End If
				Else
					If Page < DEF_MaxJumpPageNum Then
						Rs.Move Page * DEF_TopicContentMaxListNum
					Else
						If Page < MaxPage Then
							Rs.Move (MaxPage-Page-1) * DEF_TopicContentMaxListNum + LastNumBak
						End If
					End If
				End If
			End If
		End If
	case 2:
		movenum = 0
		If (LMT_First = 0 and LastNum = 0 and Page >= 1 and NoPage = 0) or JMPage > 0 Then
				If JMPage > 0 Then
					If Upflag="0" Then
						movenum = (JMPage-1)* DEF_TopicContentMaxListNum
					Else
						If Page < MaxPage Then
							movenum = (JMPage-2) * DEF_TopicContentMaxListNum + DEF_TopicContentMaxListNum
						End If
					End If
				Else
					If Page < DEF_MaxJumpPageNum Then
						movenum = Page * DEF_TopicContentMaxListNum
					Else
						If Page < MaxPage Then
							movenum = (MaxPage-Page-1) * DEF_TopicContentMaxListNum + LastNumBak
						End If
					End If
				End If
		End If
		TmpSQL = sql_select("Select T1.ID from LeadBBS_Announce as T1 " & SQLEndString,movenum & "," & sql)
		Set Rs = LDExeCute(TmpSQL,0)
	end select
	Dim Cur_RootID
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
	SQLEndString_J = Replace(SQLEndString_J,">" & LMT_First & " ",">=" & Cur_RootID & " ")
	SQLEndString_J = Replace(SQLEndString_J,"<" & LMT_First & " ","<=" & Cur_RootID & " ")
	TmpSQL = sql_select("Select T1.ID,T1.ParentID,T1.BoardID,T1.ChildNum,T1.Title,T1.Content,T1.ndatetime,T1.UserName,T1.HTMLFlag,T2.UserLimit,T1.TopicType,T1.TitleStyle from LeadBBS_Announce as T1 left join LeadBBS_User as T2 on T2.Id=T1.Userid " & SQLEndString_J,DEF_MaxListNum)
	Set Rs = LDExeCute(TmpSQL,0)
Else
'新代码结束
	select case DEF_UsedDataBase
		case 0,1:
			SQL = sql_select("Select T1.ID,T1.ParentID,T1.BoardID,T1.ChildNum,T1.Title,T1.Content,T1.ndatetime,T1.UserName,T1.HTMLFlag,T2.UserLimit,T1.TopicType,T1.TitleStyle from LeadBBS_Announce as T1 left join LeadBBS_User as T2 on T2.Id=T1.Userid " & SQLEndString,sql)
			Set Rs = LDExeCute(SQL,0)
			If (LMT_First = 0 and LastNum = 0 and Page >= 1 and NoPage = 0) or JMPage > 0 Then
				If Not Rs.Eof Then
					If JMPage > 0 Then
						If Upflag="0" Then
							Rs.Move (JMPage-1)* DEF_TopicContentMaxListNum
						Else
							If Page < MaxPage Then
								Rs.Move (JMPage-2) * DEF_TopicContentMaxListNum + DEF_TopicContentMaxListNum
							End If
						End If
					Else
						If Page < DEF_MaxJumpPageNum Then
							Rs.Move Page * DEF_TopicContentMaxListNum
						Else
							If Page < MaxPage Then
								Rs.Move (MaxPage-Page-1) * DEF_TopicContentMaxListNum + LastNumBak
							End If
						End If
					End If
				End If
			End If
		case 2:	
			movenum = 0
			If (LMT_First = 0 and LastNum = 0 and Page >= 1 and NoPage = 0) or JMPage > 0 Then
					If JMPage > 0 Then
						If Upflag="0" Then
							movenum = (JMPage-1)* DEF_TopicContentMaxListNum
						Else
							If Page < MaxPage Then
								movenum = (JMPage-2) * DEF_TopicContentMaxListNum + DEF_TopicContentMaxListNum
							End If
						End If
					Else
						If Page < DEF_MaxJumpPageNum Then
							movenum = Page * DEF_TopicContentMaxListNum
						Else
							If Page < MaxPage Then
								movenum = (MaxPage-Page-1) * DEF_TopicContentMaxListNum + LastNumBak
							End If
						End If
					End If
			End If
			SQL = sql_select("Select T1.ID,T1.ParentID,T1.BoardID,T1.ChildNum,T1.Title,T1.Content,T1.ndatetime,T1.UserName,T1.HTMLFlag,T2.UserLimit,T1.TopicType,T1.TitleStyle from LeadBBS_Announce as T1 left join LeadBBS_User as T2 on T2.Id=T1.Userid " & SQLEndString,movenum & "," & sql)
			Set Rs = LDExeCute(SQL,0)
		end select
'新代码开始
End If
'新代码结束
			If Not Rs.Eof Then
				HaveIDFlag = 1
				GetData_2 = Rs.GetRows(DEF_TopicContentMaxListNum)
			Else
				HaveIDFlag = 0
			End If
			Rs.Close
			Set Rs = Nothing
	
			If HaveIDFlag = 1 Then
				Temp2 = Ubound(GetData_2,2)
				FirstID_2 = cCur(GetData_2(0,0))
				LastID_2 = cCur(GetData_2(0,Temp2))
		
				If FirstID_2<LastID_2 Then
					SQL = FirstID_2
					FirstID_2 = LastID_2
					LastID_2 = SQL
				End If
		
				LastID = LastID_2
				FirstID = FirstID_2
			Else
				Temp2 = 0
			End If
		Else
			HaveIDFlag = 0
		End If
	
		SQL = "?B=" & GBL_board_ID & "&ID=" & Form_parentID
		Dim PageSplitString,PageSplitString2
		PageSplitString = "&nbsp;"
		If LastID <= All_LastID Then
			PageSplitString = PageSplitString & "<font color=888888><font face=webdings title=首页>9</font></font>"
			PageSplitString = PageSplitString & "<font color=888888><font face=webdings title=上页>7</font></font>"
		Else
			PageSplitString = PageSplitString & "<a href=Default.asp?" & M_Par.GetPar(GBL_board_ID,Form_ParentID,0,0,0,0,0,"a") & "><font face=webdings title=首页>9</font></a>"
			PageSplitString = PageSplitString & "<a href=Default.asp?" & M_Par.GetPar(GBL_board_ID,Form_ParentID,0,1,Page-1,0,0,"a-" & LastID) & "><font face=webdings title=上页>7</font></a>"
		End If
	
		Dim DN,N
		DN = DEF_DisplayJumpPageNum
		Dim For1,For2,StepValue,DotFlag
		DotFlag = 0
		PageSplitString = PageSplitString & " "
	
		If MaxPage > 0 Then
			For1 = Page - DN
			For2 = Page + DN
			If For1 < 0 Then
				For1 = 0
			ElseIf For1 > 0 Then
				PageSplitString = PageSplitString & " ..."
			End If
			If For2 >= MaxPage Then
				For2 = MaxPage
			End If
			If For2 > MaxPage Then For2 = MaxPage
			For N = For1 to For2
				If N <> For1 Then PageSplitString = PageSplitString & " "
				If N = Page Then
					PageSplitString = PageSplitString & "<font color=Red class=redfont>" & N + 1 & "</font>"
				Else
					If (N-Page) > 0 Then
						PageSplitString = PageSplitString & "<a href=Default.asp?" & M_Par.GetPar(GBL_board_ID,Form_ParentID,FirstID,0,Page,N-Page,0,"a") & ">" & N + 1 & "</a>"
					Else
						PageSplitString = PageSplitString & "<a href=Default.asp?" & M_Par.GetPar(GBL_board_ID,Form_ParentID,LastID,1,Page,Page-N,0,"a") & ">" & N + 1 & "</a>"
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
		
		If FirstID >= All_FirstID Then
			PageSplitString = PageSplitString & " <font color=888888><font face=webdings title=下页>8</font></font>"
			PageSplitString = PageSplitString & "<font color=888888><font face=webdings title=尾页>:</font></font>"
		Else
			PageSplitString = PageSplitString & " <a href=Default.asp?" & M_Par.GetPar(GBL_board_ID,Form_ParentID,0,0,Page+1,0,0,"a-" & FirstID) & "><font face=webdings title=下页>8</font></a>"
			PageSplitString = PageSplitString & "<a href=Default.asp?" & M_Par.GetPar(GBL_board_ID,Form_ParentID,0,1,MaxPage,0,1,"a") & "><font face=webdings title=尾页>:</font></a>"
		End If
		Rs = Temp1
		Rs = Temp2+Rs
	
		If HaveIDFlag = 1 Then Rs = Rs+1
		If Rs < 3 and GBL_ShowBottomSure = 0 then GBL_SiteBottomString = ""
		PageSplitString = PageSplitString & " 此主题共有" & ALL_Count &"帖 此页" & Rs & "帖 每页" & DEF_TopicContentMaxListNum & "帖<td align=right>"
		'<img src=" & DEF_BBS_HomeUrl & "images/null.gif width=2 height=2><br>返回<a href=""" & DEF_BBS_HomeUrl & "b/b.asp?B=" & GBL_Board_ID & """>" & GBL_Board_BoardName & "</a>
		PageSplitString2 = "&nbsp;"
	
		For1 = 0
		For2 = 0
		
		If HaveIDFlag = 1 Then
			If Upflag="0" Then
				For1 = 0
				For2 = Temp2
				StepValue = 1
			Else
				For1 = Temp2
				For2 = 0
				StepValue = -1
			End If
			
			If cCur(GetData_2(1,For1)) = 0 and (GetData_2(8,For1) = 0 or GetData_2(8,For1) = 2) Then GetData_2(5,For1) = PrintTrueText(GetData_2(5,For1))
			DisplayAnnounce For1,For2,StepValue,GetData_2
		End If
		Response.Write PageSplitString
		Response.Write PageSplitString2
	
	End Sub
	
	Private Sub DisplayAnnounce(For1,For2,StepValue,GetData)
	
		Dim Temp,N
	
		Temp = LCase(Request.ServerVariables("server_name"))
		If inStr(Temp,".") <> inStrRev(Temp,".") Then Temp = Mid(Temp,inStr(Temp,".") + 1)
		%>
		<script src="../a/inc/leadcode.js"></script>
		<script language=javascript>
		var GBL_domain="<%=Temp%>";
		HU="<%=DEF_BBS_HomeUrl%>";
		var DEF_DownKey="<%=UrlEncode(DEF_DownKey)%>";</script><%
	
		Temp = LCase(Request.ServerVariables("server_name"))
		If inStr(Temp,".") <> inStrRev(Temp,".") Then Temp = Mid(Temp,inStr(Temp,".") + 1)
	
		For N = For1 to For2 Step StepValue
			If isNull(GetData(9,n)) Then GetData(9,n) = 0
	
			If (GetData(8,n) = 0 or GetData(8,n) = 2) and cCur(GetData(1,n)) <> 0 Then GetData(5,n) = PrintTrueText(GetData(5,n))
	
			If GetData(11,n) >=60 Then GetData(5,n) = GetFobStr("此帖有待管理人员审核才能查看")
			If GetBinarybit(GetData(9,n),7) = 1 Then
				Response.Write "<br><span style=""line-height:15pt;"">" & GetFobStr("该用户发言已经被屏蔽") & "</span><br>"
			Else
				If GetData(10,n) > 0 and cCur(GetData(1,n)) = 0 and GetData(10,n) <> 80 Then
					If GetData(10,n) > 0 Then GetData(5,n) = GetFobStr("此帖内容已经加密，要查看请点击完整模式")
				End If
				If Lcase(Left(GetData(4,n),3)) <> "re:" or cCur(GetData(1,n)) = 0 Then Response.Write "<b>" & DisplayAnnounceTitle(GetData(4,n),GetData(11,n)) & "</b>" & VbCrLf
				Response.Write "<font color=gray class=grayfont>" & HtmlEncode(GetData(7,N)) & "," & RestoreTime(GetData(6,N)) & "</font><br><br>" & VbCrLf
				If DEF_AnnounceFontSize <> "0" then Response.Write "<span style=font-size:" & DEF_AnnounceFontSize & ">"
				If GetData(8,n) <> 2 Then
					Response.Write GetData(5,n)
				Else
					Response.Write "<span id=Content" & GetData(0,n) & ">"
					Response.Write GetData(5,n)
					Response.Write "</span><script language=javascript>" & VbCrLf & "<!--" & VbCrLf & "leadcode('Content" & GetData(0,n) & "');" & VbCrLf & "//-->" & VbCrLf & "</script>"
				End If
		
				If DEF_AnnounceFontSize <> "0" then Response.Write "</span>"
			End If
			Response.Write "<hr size=1>"
		Next
	
	End Sub
	
	Private Function PrintTrueText(tempString)
	
		If tempString<>"" Then
			PrintTrueText=Replace(Replace(Replace(Replace(Replace(htmlEncode(tempString),VbCrLf & " ","<br>" & "&nbsp;"),VbCrLf,"<br>" & VbCrLf),"   "," &nbsp; "),"  "," &nbsp;"),chr(9)," &nbsp; &nbsp; &nbsp;")
	
			If Left(PrintTrueText,1) = chr(32) Then
				PrintTrueText = "&nbsp;" & Mid(PrintTrueText,2)
			End If
		Else
			PrintTrueText=""
		End If
	
	End Function
	
	
	Private Function GetFobStr(Str)
	
		GetFobStr = "<font color=888888 class=grayfont>……………………………………………………隐藏内容…<br>" & _
					"<font color=blue class=bluefont>" & Str & "</font><br>" & _
					"…………………………………………………………………</font><br>"
	
	End Function

End Class
%>