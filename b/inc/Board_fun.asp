<!-- #include file=../../inc/Templet/HTML/Normal_1.asp -->
<%
Dim LMT_UrlEndString

Sub DisplayAnnouncesSplitPages

	Dim Temp
	If GBL_BoardMasterFlag >= 5 Then%>
	<script src="<%=DEF_BBS_HomeUrl%>inc/js/p_list.js?ver=20090601.2" type="text/javascript"></script>
	<%End If%>
	<script type="text/javascript">
	<!--
	function Show2(obj,obj2,id)
	{
		if ($id(obj).style.display!='block')
		{
			$id(obj).style.display="block";
			$id((obj2)).src="../images/<%=GBL_DefineImage%>Expand.gif";
			if($id(obj).innerHTML=="")
			{
				$id(obj).innerHTML="下载中...";getAJAX("b.asp","ol=3&id=" + id,"Lead" + id);
			}
		}else{
			$id(obj).style.display="none";
			$id((obj2)).src="../images/<%=GBL_DefineImage%>clsExpand.gif";
		}
	}
	<%If GBL_BoardMasterFlag >= 5 Then
		If GBL_Board_ID <> 444 and DEF_EnableDelAnnounce = 0 Then
			Temp = "Move&b=" & GBL_Board_ID & "&BoardID2=444"
		Else
			Temp = "Del&b=" & GBL_Board_ID
		End If
	%>
	function a_command(cstr,obj,action)
	{
		layer_view(cstr,obj,'','','anc_delbody','<%=DEF_BBS_HomeUrl%>a/Processor.asp','',1,'AjaxFlag=1&action=' + action,1);return(false);
	}
	function delbody_view(obj)
	{
		layer_create("anc_msgbody");
		$id('anc_msgbody').innerHTML="<div class=ajaxbox>已选择 <b id=layer_selectnum>" + p_getnum() + "</b> 条记录：<br>请选择操作：<b><a href=\"javascript:;\" onclick=\"a_command('删除帖子',$id('" + obj.id + "'),'<%=Temp%>&ID='+p_getselected());\">批量删除</a>, <a href=\"javascript:;\" onclick=\"a_command('转移帖子',$id('" + obj.id + "'),'<%Response.Write "Move&b=" & GBL_Board_ID & ""%>&ID='+p_getselected());\">批量转移</a></b><br><input class=\"fmchkbox\" type=\"checkbox\" name=\"selmsg\" id=\"selmsg\" value=\"1\" onclick=\"achoose();\" />选择全部</div>";
		layer_view('',obj,'','','anc_msgbody','','',0,'',0,20);
	}
	<%End If%>
	-->
	</script>
<%
	If EFlag >=0 Then
		DisplayAnnouncesSplitPages_Elist
	Else
		DisplayAnnouncesSplitPages_List
	End If

End Sub

Function DisplayAnnouncesSplitPages_List

	Dim Rs,SQL,Temp
	Dim SQLEndString,WhereFlag
	Dim ALL_Count
	Dim RootID,Temp1,Temp2
	Dim Upflag
	Dim LastNum,LastNumBak
	Dim Page,MaxPage,JMPage
	Dim JMPRootID
	Dim HaveRootIDFlag
	Dim FirstRootID,LastRootID
	Dim NoPage
	Dim FirstRootID_2,LastRootID_2
	Dim GetData_2
	Dim GetDataTop,AllTopNum,N
	Dim GetDataPartTop,PartTopNum
	Dim PageSplitString
	Dim DN
	Dim For1,For2,StepValue,DotFlag
	Dim TArray,Num
	Dim BoardListClass
	Dim JumpOnly
	JumpOnly = 1
	WhereFlag = 1

	select case DEF_UsedDataBase
		case 0,2:
			SQLEndString = " where ParentID=0 and boardid=" & GBL_board_ID
		case Else
			SQLEndString = " where boardid=" & GBL_board_ID
	End select

	ALL_Count = GBL_Board_TopicNum


	RootID = Left(Request.QueryString("RootID"),14)
	If isNumeric(RootID)=0 Then RootID=0
	RootID = cCur(RootID)
	If RootID > 0 Then JumpOnly = 0

	LastNum = 0
	
	LastNumBak = (ALL_Count mod DEF_MaxListNum)
	If LastNumBak = 0 Then LastNumBak = DEF_MaxListNum

	Upflag = Request.QueryString("Upflag")
	If Upflag <> "" Then JumpOnly = 0
	If Upflag<>"1" and Upflag<>"0" Then Upflag="0"
	If Upflag = "1" Then
		LastNum = Request.QueryString("Num")
		If LastNum <> "" Then
			LastNum = LastNumBak
		End If
	End If


	JMPage = Left(Request.QueryString("q"),14)
	If JMPage = "" Then JumpOnly = 0
	
	Page = Left(Request.QueryString("p"),14)
	If isNumeric(Page) = 0 or inStr(Page,".") > 0 Then Page = 0
	Page = cCur(Page)
	MaxPage = Fix(All_Count / DEF_MaxListNum)
	If (All_Count mod DEF_MaxListNum)<>0 Then MaxPage = MaxPage + 1
	If JumpOnly = 0 Then MaxPage = MaxPage - 1
	If Page > MaxPage or LastNum > 0 Then
		Page = MaxPage
	End If

	If isNumeric(JMPage) = 0 or inStr(JMPage,".") > 0 Then JMPage = 0
	JMPage = Fix(cCur(JMPage))
	If JMPage > DEF_MaxJumpPageNum+1 Then JMPage = 0
	
	JMPRootID = Left(Request.QueryString("r"),14)
	If isNumeric(JMPRootID)=0 Then JMPRootID=0
	JMPRootID = Fix(cCur(JMPRootID))
	
	If JMPage > Maxpage+1 or Maxpage < 0 Then JMPage = 0
	If Upflag="0" and JMPage+Page > MaxPage Then JMPage = 0
	If Upflag="1" and JMPage+Page < 0 Then JMPage = 0
	If JMPRootID > GBL_Board_AllMaxRootID+1 and JumpOnly = 0 Then JMPage = 0
	If JMPRootID < GBL_Board_AllMinRootID-1 and JumpOnly = 0 Then JMPage = 0


	If RootID > 0 Then LMT_UrlEndString = LMT_UrlEndString & "&amp;RootID=" & RootID
	If Page > 0 Then LMT_UrlEndString = LMT_UrlEndString & "&amp;p=" & Page	
	If Upflag = "1" Then LMT_UrlEndString = LMT_UrlEndString & "&amp;Upflag=" & Upflag	
	If LastNum <> 0 Then LMT_UrlEndString = LMT_UrlEndString & "&amp;Num=1"
	If JMPage > 0 Then LMT_UrlEndString = LMT_UrlEndString & "&amp;q=" & JMPage
	If JMPRootID > 0 Then LMT_UrlEndString = LMT_UrlEndString & "&amp;r=" & JMPRootID
	
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


	If Temp1+1<DEF_MaxListNum and All_Count > 0 Then
		If Upflag="0" Then
			If JMPage > 0 Then
				If JMPRootID<>0 Then
					If WhereFlag = 1 Then
						SQLEndString = SQLEndString & " And RootID<" & JMPRootID
					Else
						SQLEndString = SQLEndString & " Where RootID<" & JMPRootID
						WhereFlag = 1
					End If
				End If
			Else
				If RootID<>0 Then
					If WhereFlag = 1 Then
						SQLEndString = SQLEndString & " And RootID<" & RootID
					Else
						SQLEndString = SQLEndString & " Where RootID<" & RootID
						WhereFlag = 1
					End If
				End If
			End If
		Else
			If JMPage > 0 Then
				If WhereFlag = 1 Then
					SQLEndString = SQLEndString & " And RootID>" & JMPRootID
				Else
					SQLEndString = SQLEndString & " Where RootID>" & JMPRootID
					WhereFlag = 1
				End If
			Else
				If RootID>=GBL_Board_AllMaxRootID Then
					RootID = GBL_Board_AllMaxRootID-1
					'TopicSortID = GBL_Board_AllMaxRootID + 1
				
					If RootID<>0 Then
						If WhereFlag = 1 Then
							SQLEndString = SQLEndString & " And RootID>" & RootID
						Else
							SQLEndString = SQLEndString & " Where RootID>" & RootID
							WhereFlag = 1
						End If
					End If
				Else
					If RootID<>0 Then
						If WhereFlag = 1 Then
							SQLEndString = SQLEndString & " And RootID>" & RootID
						Else
							SQLEndString = SQLEndString & " Where RootID>" & RootID
							WhereFlag = 1
						End If
					End If
				End If
			End If
		End If

		NoPage = 0

		If Page < 0 or (Page > MaxPage and MaxPage>=(DEF_MaxJumpPageNum-1)) or (Page > (DEF_MaxJumpPageNum-1) and Page<(MaxPage-DEF_MaxJumpPageNum+1)) Then NoPage = 1
		If (RootID > 0 or LastNum>0 or NoPage = 1) and JMPage < 1 Then
			If Upflag="0" Then
				SQLEndString = SQLEndString & " order by RootID DESC"
			Else
				SQLEndString = SQLEndString & " order by RootID ASC"
			End If
			If LastNum > 0 Then
				Temp = LastNum
			Else
				Temp = DEF_MaxListNum
			End If
		Else
			If JMPage > 0 Then
				If Upflag="0" Then
					SQLEndString = SQLEndString & " order by RootID DESC"
					Upflag="0"
					Temp = (JMPage-1) * DEF_MaxListNum + DEF_MaxListNum
				Else
					SQLEndString = SQLEndString & " order by RootID ASC"
					Upflag="1"
					Temp = (JMPage-1) * DEF_MaxListNum + DEF_MaxListNum
				End If
			Else
				If Page < DEF_MaxJumpPageNum Then
					SQLEndString = SQLEndString & " order by RootID DESC"
					Upflag="0"
					Temp = Page * DEF_MaxListNum + DEF_MaxListNum
				Else
					SQLEndString = SQLEndString & " order by RootID ASC"
					Upflag="1"
					Temp = (MaxPage-Page) * DEF_MaxListNum + LastNumBak
				End If
			End If
		End If
		dim forindex
		forindex = get_index("IX_LeadBBS_Announce_ParentID")

	Dim moveNum
'新代码开始
If (DEF_UsedDataBase = 0 or DEF_UsedDataBase = 2) and Temp>1000 Then
	select case DEF_UsedDataBase
	case 0:
		SQL = sql_select("select RootID from LeadBBS_Announce " & SQLEndString,Temp)
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
		MoveNum = 0
		If LastNum = "" or isNull(LastNum) then LastNum = 0
		LastNum = cCur(LastNum)
		If (RootID = 0 and Page >= 1 and LastNum = 0 and NoPage = 0) or JMPage > 0 Then
				If JMPage > 0 Then
					If Upflag="0" Then
						MoveNum = (JMPage-1)* DEF_MaxListNum
					Else
						If Page < MaxPage Then
							MoveNum = (JMPage-2) * DEF_MaxListNum + DEF_MaxListNum
						End If
					End If
				Else
					If Page < DEF_MaxJumpPageNum Then
						MoveNum = Page * DEF_MaxListNum
					Else
						If Page < MaxPage Then
							MoveNum = (MaxPage-Page-1) * DEF_MaxListNum + LastNumBak
						End If
					End If
				End If
		End If
		SQL = sql_select("select RootID from LeadBBS_Announce " & SQLEndString,MoveNum & "," & Temp)
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
	SQLEndString_J = SQLEndString
	
	If JumpOnly = 1 and inStr(SQLEndString_J," order by RootID DESC") Then
		SQLEndString_J = Replace(SQLEndString_J," order by RootID DESC"," and RootID<=" & Cur_RootID & " order by RootID DESC")
	Else
		SQLEndString_J = Replace(SQLEndString_J,">" & JMPRootID & " ",">=" & Cur_RootID & " ")
		SQLEndString_J = Replace(SQLEndString_J,"<" & JMPRootID & " ","<=" & Cur_RootID & " ")
		SQLEndString_J = Replace(SQLEndString_J,">" & RootID & " ",">=" & Cur_RootID & " ")
		SQLEndString_J = Replace(SQLEndString_J,"<" & RootID & " ","<=" & Cur_RootID & " ")
	End If
	SQL = sql_select("select id,ChildNum,Title,FaceIcon,LastTime,Hits,Length,UserName,UserID,RootID,LastUser,NotReplay,GoodFlag,BoardID,TopicType,PollNum,TitleStyle,LastInfo,ndatetime,GoodAssort,NeedValue from LeadBBS_Announce " & forindex & "" & SQLEndString_J,DEF_MaxListNum)
	Set Rs = LDExeCute(SQL,0)
Else
'新代码结束
		select case DEF_UsedDataBase
			case 0:
				SQL = sql_select("select id,ChildNum,Title,FaceIcon,LastTime,Hits,Length,UserName,UserID,RootID,LastUser,NotReplay,GoodFlag,BoardID,TopicType,PollNum,TitleStyle,LastInfo,ndatetime,GoodAssort,NeedValue from LeadBBS_Announce " & forindex & "" & SQLEndString,Temp)
			case 2:
				MoveNum = 0
				If (RootID = 0 and Page >= 1 and LastNum = 0 and NoPage = 0) or JMPage > 0 Then
						If JMPage > 0 Then
							If Upflag="0" Then
								MoveNum = (JMPage-1)* DEF_MaxListNum
							Else
								If Page < MaxPage Then
									MoveNum = (JMPage-2) * DEF_MaxListNum + DEF_MaxListNum
								End If
							End If
						Else
							If Page < DEF_MaxJumpPageNum Then
								MoveNum = Page * DEF_MaxListNum
							Else
								If Page < MaxPage Then
									MoveNum = (MaxPage-Page-1) * DEF_MaxListNum + LastNumBak
								End If
							End If
						End If
				End If
				SQL = sql_select("select id,ChildNum,Title,FaceIcon,LastTime,Hits,Length,UserName,UserID,RootID,LastUser,NotReplay,GoodFlag,BoardID,TopicType,PollNum,TitleStyle,LastInfo,ndatetime,GoodAssort,NeedValue from LeadBBS_Announce " & forindex & "" & SQLEndString,MoveNum & "," & Temp)
			case Else
				SQL = sql_select("select id,ChildNum,Title,FaceIcon,LastTime,Hits,Length,UserName,UserID,RootID,LastUser,NotReplay,GoodFlag,BoardID,TopicType,PollNum,TitleStyle,LastInfo,ndatetime,GoodAssort,NeedValue from LeadBBS_Topic " & SQLEndString,Temp)
		End select
		Set Rs = LDExeCute(SQL,0)

		If LastNum = "" or isNull(LastNum) then LastNum = 0
		LastNum = cCur(LastNum)
		if DEF_UsedDataBase <> 2 then
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
		end if
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
			FirstRootID_2 = cCur(GetData_2(9,0))
			LastRootID_2 = cCur(GetData_2(9,Temp2))
				
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
	
	AllTopNum = -1
	If FirstRootID >= GBL_Board_AllMaxRootID Then
		GetDataTop = application(DEF_MasterCookies & "TopAnc")
		If isArray(GetDataTop) = False Then
			If GetDataTop & "" <> "yes" Then
				ReloadTopAnnounceInfo(0)
				GetDataTop = application(DEF_MasterCookies & "TopAnc")
			End If
		End If
		If isArray(GetDataTop) Then
			GetDataTop = application(DEF_MasterCookies & "TopAnc")
			AllTopNum = Ubound(GetDataTop,2)
		End If
	End If
	
	PartTopNum = -1
	If FirstRootID >= GBL_Board_AllMaxRootID Then
		GetDataPartTop = application(DEF_MasterCookies & "TopAnc" & GBL_Board_BoardAssort)
		If isArray(GetDataPartTop) = False Then
			If GetDataPartTop & "" <> "yes" Then
				ReloadTopAnnounceInfo(GBL_Board_BoardAssort)
				GetDataPartTop = application(DEF_MasterCookies & "TopAnc" & GBL_Board_BoardAssort)
			End If
		End If
		If isArray(GetDataPartTop) Then
			GetDataPartTop = application(DEF_MasterCookies & "TopAnc" & GBL_Board_BoardAssort)
			PartTopNum = Ubound(GetDataPartTop,2)
		End If
	End If

	Dim RewriteFlag,RewriteFile,RewriteStr
	If JumpOnly = 1 and GetBinarybit(DEF_Sideparameter,16) = 1 Then
		RewriteFlag = 1
		RewriteFile = "forum"
	Else
		RewriteFlag = 0
		RewriteFile = "b.asp"
	End If
	If RewriteFlag = 0 Then
		SQL = "?B=" & GBL_board_ID
	Else
		SQL = "-" & GBL_board_ID
	End If
	PageSplitString = "<div class=""j_page"">"
	If JumpOnly = 1 and MaxPage > 0 Then
		MaxPage = MaxPage - 1
	End If
	If JumpOnly = 1 and Page > 0 Then
		Page = Page - 1
	End If
	If FirstRootID >= GBL_Board_AllMaxRootID and Page = 0 Then
	Else
		If RewriteFlag = 0 Then
			RewriteStr = "&amp;p=0"
		Else
			RewriteStr = "-1.html"
		End If
		PageSplitString = PageSplitString & "<a href=""" & RewriteFile & "" & SQL & RewriteStr & """>1</a>"
		
		if page <> 1 Then
			If RewriteFlag = 0 Then
				RewriteStr = "&amp;RootID=" & FirstRootID & "&amp;Upflag=1&amp;p=" & Page-1
			Else
				RewriteStr = "-" & Page & ".html"
			End If
			PageSplitString = PageSplitString & "<a href=""" & RewriteFile & "" & SQL & RewriteStr & """>上页"
			If (Page - DEF_DisplayJumpPageNum) > 0 Then PageSplitString = PageSplitString & "…"
			PageSplitString = PageSplitString & "</a>"
		End If
	End If
	
	DN = DEF_DisplayJumpPageNum
	DotFlag = 0

	If MaxPage > 0 Then
		For1 = Page - DN
		For2 = Page + DN
		If For1 < 0 Then
			For1 = 0
		End If
		If For2 >= MaxPage Then For2 = MaxPage
		'忽略 If For2 - For1 < DEF_MaxJumpPageNum and For1 > 1 and For2 > For1 Then For1 = For1 - (DEF_MaxJumpPageNum - ((For2 - For1) + 1))
		If For2 - For1 < DEF_DisplayJumpPageNum*2 and For1 > 1 and For2 > For1 Then For1 = For1 - (DEF_DisplayJumpPageNum*2 - ((For2 - For1) + 1))
		
		If For1 < 0 Then
			For1 = 0
		'ElseIf For1 > 0 Then
		'	PageSplitString = PageSplitString & "<b>…</b>"
		End If

		'忽略 If For2 - For1 < DEF_MaxJumpPageNum and For2 < MaxPage and For2 > For1 Then For2 = For2 + (DEF_MaxJumpPageNum - ((For2 - For1) + 1))
		If For2 - For1 < DEF_DisplayJumpPageNum*2 and For2 < MaxPage and For2 > For1 Then For2 = For2 + (DEF_DisplayJumpPageNum*2 - ((For2 - For1) + 1))
		If For2 >= MaxPage Then For2 = MaxPage

		If For2 > 998 Then
			If Page > For1 Then For1 = For1 + 1
			If Page > For1 Then For1 = For1 + 1
			If Page < For2 - DN + 1 Then For2 = For2 - 1
		End If
		For N = For1 to For2
			If N = Page Then
				PageSplitString = PageSplitString & "<b>" & N + 1 & "</b>"
			Else
				If N <> MaxPage and N <> 0 Then
					If (N-Page) > 0 Then
						If RewriteFlag = 0 Then
							RewriteStr = "&amp;r=" & LastRootID & "&amp;Upflag=0&amp;p=" & page & "&amp;q=" & N - Page
						Else
							RewriteStr = "-" & N + 1 & ".html"
						End If
						PageSplitString = PageSplitString & "<a href=""" & RewriteFile & "" & SQL & RewriteStr & """>" & N + 1 & "</a>"
					Else
						If RewriteFlag = 0 Then
							RewriteStr = "&amp;r=" & FirstRootID & "&amp;Upflag=1&amp;p=" & page & "&amp;q=" & Page-N
						Else
							RewriteStr = "-" & N + 1 & ".html"
						End If
						PageSplitString = PageSplitString & "<a href=""" & RewriteFile & "" & SQL & RewriteStr & """>" & N + 1 & "</a>"
					End If
				End If
			End If
			DotFlag = 2
		Next
		If For2 < MaxPage Then
			'PageSplitString = PageSplitString & "<b>…</b>"
			DotFlag = 1
		End If
	Else
		PageSplitString = PageSplitString & " <b>1</b> "
	End If

	If LastRootID <= GBL_Board_AllMinRootID Then
	Else
		If page <> MaxPage-1 Then
			If RewriteFlag = 0 Then
				RewriteStr = "&amp;RootID=" & LastRootID & "&amp;p=" & Page+1
			Else
				RewriteStr = "-" & Page + 2 & ".html"
			End If
			PageSplitString = PageSplitString & "<a href=""" & RewriteFile & "" & SQL & RewriteStr & """>"
			If (Page + DN) < MaxPage Then PageSplitString = PageSplitString & "…"
			PageSplitString = PageSplitString & "下页</a>"
		End If
		If RewriteFlag = 0 Then
			RewriteStr = "&amp;Upflag=1&amp;Num=1&amp;p=" & MaxPage
		Else
			RewriteStr = "-" & MaxPage + 1 & ".html"
		End If
		PageSplitString = PageSplitString & "<a href=""" & RewriteFile & "" & SQL & RewriteStr & """>" & MaxPage + 1 & "</a>"
	End If

	Rs = Temp1
	Rs = Temp2+Rs
	If HaveRootIDFlag = 1 Then Rs = Rs+1

	If AllTopNum <> -1 Then
		ALL_Count = ALL_Count + AllTopNum + 1
		Rs =  Rs + AllTopNum + 1
	End If
	If PartTopNum <> -1 Then
		ALL_Count = ALL_Count + PartTopNum + 1
		Rs =  Rs + PartTopNum + 1
	End If
	if MaxPage < 1 Then MaxPage = 0
	'if MaxPage > 0 Then MaxPage = MaxPage + 1
	'PageSplitString = PageSplitString & " 共" & ALL_Count &"主题 第" & Page+1 & "/" & MaxPage+1 & "页 每页" & DEF_MaxListNum & "条"
	
	If MaxPage > DEF_DisplayJumpPageNum*2 Then PageSplitString = PageSplitString & "<input type=""text"" title=""输入页数,按Enter键跳转。"" size=""2"" onkeydown=""javascript:if(event.keyCode==13){location='b.asp?b=" & GBL_Board_ID & "&amp;r=" & GBL_Board_AllMaxRootID+1 & "&amp;p=-1&Upflag=0&amp;q='+(parseInt(this.value))+'';return false;}"">"
	
	PageSplitString = PageSplitString & "</div>"

	For1 = 0
	For2 = 0
	If Rs < DEF_MaxListNum and GBL_ShowBottomSure = 0 Then GBL_SiteBottomString = ""


	CALL B_DisplaySplitPageString(PageSplitString,"b_box_none")

	Global_TableHead
	Set BoardListClass = New BoardList_HTML_Class
	%>
	<div class="contentbox">
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class="tablebox table_options<%
		If BoardListClass.CFlag = 1 Then Response.Write "_sim"
		%>" id="table_options">
	<%
	BoardListClass.Showhead
	If AllTopNum <> -1 Then DisplayAnnounceData_HTML 0,AllTopNum,1,GetDataTop,1,BoardListClass
	If PartTopNum <> -1 Then DisplayAnnounceData_HTML 0,PartTopNum,1,GetDataPartTop,2,BoardListClass
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
			DisplayAnnounceData_HTML For1,For2,StepValue,GetData_2,0,BoardListClass
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
			DisplayAnnounceData_HTML For1,For2,StepValue,GetData_2,0,BoardListClass
		End If
	End If
	Set BoardListClass = Nothing
	%>
		</table>
	</div>
	<%Global_TableBottom
	CALL B_DisplaySplitPageString(PageSplitString,"b_box_none2")

End Function

Function DisplayAnnouncesSplitPages_Elist

	Dim Rs,SQL,Temp,TempArray
	Dim ALL_FirstRootID,ALL_LastRootID
	Dim ALL_Count
	Dim SQLEndString,WhereFlag
	Dim RootID,Temp1,Temp2
	Dim Upflag
	Dim LastNum,LastNumBak
	Dim Page,MaxPage,JMPage
	Dim JMPRootID
	Dim HaveRootIDFlag
	Dim FirstRootID,LastRootID
	Dim NoPage
	Dim FirstRootID_2,LastRootID_2
	Dim GetData_2
	Dim N
	Dim PageSplitString
	Dim DN
	Dim For1,For2,StepValue,DotFlag
	Dim TArray,Num
	Dim BoardListClass
	If EFlag = 1 and EID = 0 Then
		'GBL_SiteBottomString = ""
		'Exit Function
	End If

	If EFlag = 1 Then
		TempArray = Application(DEF_MasterCookies & "BoardInfo" & GBL_board_ID & "_TI")
		If isArray(TempArray) Then
			If EID = 0 Then EID = cCur(TempArray(0,0))
		Else
			TempArray = Application(DEF_MasterCookies & "BoardInfo" & GBL_board_ID)
			EID = 0
		End If
	Else
		TempArray = Application(DEF_MasterCookies & "BoardInfo" & GBL_board_ID)
	End if
	If isArray(TempArray) = False Then Exit Function

	WhereFlag = 1

	If EID > 0 Then
		SQLEndString = " where GoodAssort=" & EID
	Else
		SQLEndString = " where GoodFlag=1 and boardid=" & GBL_board_ID
	End If
	ALL_Count = 0
	If EID > 0 Then
		'If Ubound(TempArray,2) = 0 or Ubound(TempArray,2) < EIndex Then Exit Function
		ALL_Count = cCur(TempArray(2,EIndex))
		If ALL_Count <= 0 Then
			select case DEF_UsedDataBase
				case 0,2:
					Set Rs = LDExeCute("Select Count(*) from LeadBBS_Announce where GoodAssort=" & EID,0)
				case Else
					Set Rs = LDExeCute("Select Count(*) from LeadBBS_Topic where GoodAssort=" & EID,0)
			End select
			If Rs.Eof Then
				ALL_Count = -1
			Else
				ALL_Count = Rs(0)
				If isNull(ALL_Count) Then ALL_Count = 0
				ALL_Count = cCur(ALL_Count)
				If ALL_Count = 0 Then ALL_Count = -1
			End If
			Rs.Close
			Set Rs = Nothing
			TempArray(2,EIndex) = ALL_Count
			Application.Lock
			Application(DEF_MasterCookies & "BoardInfo" & GBL_board_ID & "_TI") = TempArray
			Application.UnLock
		End If
	Else
		ALL_Count = GBL_Board_GoodNum
	End If

	If All_Count > 0 Then
		If EID > 0 Then
			ALL_FirstRootID = ccur(TempArray(3,EIndex))
			ALL_LastRootID = ccur(TempArray(4,EIndex))
		Else
			ALL_FirstRootID = ccur(TempArray(33,0))
			ALL_LastRootID = ccur(TempArray(34,0))
		End If
		If ALL_FirstRootID = 0 Then
			select case DEF_UsedDataBase
				case 0,2:
					SQL = "Select Max(ID) from LeadBBS_Announce " & SQLEndString
				case Else
					SQL = "Select Max(ID) from LeadBBS_Topic " & SQLEndString
			End select
			Set Rs = LDExeCute(SQL,0)
			If Not Rs.Eof Then
				ALL_FirstRootID = Rs(0)
				If isNull(ALL_FirstRootID) Then ALL_FirstRootID = 0
				ALL_FirstRootID = cCur(ALL_FirstRootID)
				If EID > 0 Then
					TempArray(3,EIndex) = ALL_FirstRootID
					Application.Lock
					Application(DEF_MasterCookies & "BoardInfo" & GBL_board_ID & "_TI") = TempArray
					Application.UnLock
				Else
					UpdateBoardApplicationInfo GBL_board_ID,ALL_FirstRootID,33
				End If
			Else
				ALL_FirstRootID = 0
			End If
			Rs.Close
			Set Rs = Nothing
		End If
		If ALL_LastRootID = 0 Then
			select case DEF_UsedDataBase
				case 0,2:
					SQL = "Select Min(ID) from LeadBBS_Announce " & SQLEndString
				case Else
					SQL = "Select Min(ID) from LeadBBS_Topic " & SQLEndString
			End select
			Set Rs = LDExeCute(SQL,0)
			If Not Rs.Eof Then
				ALL_LastRootID = Rs(0)
				If isNull(ALL_LastRootID) Then ALL_LastRootID = 0
				ALL_LastRootID = cCur(ALL_LastRootID)
				If EID > 0 Then
					TempArray(4,EIndex) = ALL_LastRootID
					Application.Lock
					Application(DEF_MasterCookies & "BoardInfo" & GBL_board_ID & "_TI") = TempArray
					Application.UnLock
				Else
					UpdateBoardApplicationInfo GBL_board_ID,ALL_LastRootID,34
				End If
			Else
				ALL_LastRootID = 0
			End If
			Rs.Close
			Set Rs = Nothing
		End If
	Else
		ALL_FirstRootID = 0
		ALL_LastRootID = 0
	End If

	RootID = Left(Request.QueryString("RootID"),14)
	If isNumeric(RootID)=0 Then RootID=0
	RootID = cCur(RootID)

	LastNum = 0
	
	LastNumBak = (ALL_Count mod DEF_MaxListNum)
	If LastNumBak = 0 Then LastNumBak = DEF_MaxListNum

	Upflag = Request.QueryString("Upflag")
	If Upflag<>"1" and Upflag<>"0" Then Upflag="0"
	If Upflag = "1" Then
		LastNum = Request.QueryString("Num")
		If LastNum <> "" Then
			LastNum = LastNumBak
		End If
	End If

	Page = Left(Request.QueryString("p"),14)
	If isNumeric(Page) = 0 or inStr(Page,".") > 0 Then Page = 0
	Page = cCur(Page)
	MaxPage = Fix(All_Count / DEF_MaxListNum)
	If (All_Count mod DEF_MaxListNum)<>0 Then MaxPage = MaxPage + 1
	MaxPage = MaxPage - 1
	If Page > MaxPage or LastNum > 0 Then
		Page = MaxPage
	End If

	JMPage = Left(Request.QueryString("q"),14)
	If isNumeric(JMPage) = 0 or inStr(JMPage,".") > 0 Then JMPage = 0
	JMPage = Fix(cCur(JMPage))
	If JMPage > DEF_MaxJumpPageNum Then JMPage = 0
	
	JMPRootID = Left(Request.QueryString("r"),14)
	If isNumeric(JMPRootID)=0 Then JMPRootID=0
	JMPRootID = Fix(cCur(JMPRootID))
	
	If JMPage > Maxpage or Maxpage < 0 Then JMPage = 0
	If Upflag="0" and JMPage+Page > MaxPage Then JMPage = 0
	If Upflag="1" and JMPage+Page < 0 Then JMPage = 0
	If JMPRootID > ALL_FirstRootID Then JMPage = 0
	If JMPRootID < ALL_LastRootID Then JMPage = 0

	LMT_UrlEndString = "&amp;E=" & EFlag
	If EID > 0 Then LMT_UrlEndString = LMT_UrlEndString & "&amp;EID=" & EID
	If RootID > 0 Then LMT_UrlEndString = LMT_UrlEndString & "&amp;RootID=" & RootID
	If Page > 0 Then LMT_UrlEndString = LMT_UrlEndString & "&amp;p=" & Page	
	If Upflag = "1" Then LMT_UrlEndString = LMT_UrlEndString & "&amp;Upflag=" & Upflag	
	If LastNum <> 0 Then LMT_UrlEndString = LMT_UrlEndString & "&amp;Num=1"
	If JMPage > 0 Then LMT_UrlEndString = LMT_UrlEndString & "&amp;q=" & JMPage
	If JMPRootID > 0 Then LMT_UrlEndString = LMT_UrlEndString & "&amp;r=" & JMPRootID

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


	If Temp1+1<DEF_MaxListNum and All_Count > 0 Then
		If Upflag="0" Then
			If JMPage > 0 Then
				If JMPRootID<>0 Then
					If WhereFlag = 1 Then
						SQLEndString = SQLEndString & " And ID<" & JMPRootID
					Else
						SQLEndString = SQLEndString & " Where ID<" & JMPRootID
						WhereFlag = 1
					End If
				End If
			Else
				If RootID<>0 Then
					If WhereFlag = 1 Then
						SQLEndString = SQLEndString & " And ID<" & RootID
					Else
						SQLEndString = SQLEndString & " Where ID<" & RootID
						WhereFlag = 1
					End If
				End If
			End If
		Else
			If JMPage > 0 Then
				If WhereFlag = 1 Then
					SQLEndString = SQLEndString & " And ID>" & JMPRootID
				Else
					SQLEndString = SQLEndString & " Where ID>" & JMPRootID
					WhereFlag = 1
				End If
			Else
				If RootID>=ALL_FirstRootID Then
					RootID = ALL_FirstRootID-1
					'TopicSortID = ALL_FirstRootID + 1
				
					If RootID<>0 Then
						If WhereFlag = 1 Then
							SQLEndString = SQLEndString & " And ID>" & RootID
						Else
							SQLEndString = SQLEndString & " Where ID>" & RootID
							WhereFlag = 1
						End If
					End If
				Else
					If RootID<>0 Then
						If WhereFlag = 1 Then
							SQLEndString = SQLEndString & " And ID>" & RootID
						Else
							SQLEndString = SQLEndString & " Where ID>" & RootID
							WhereFlag = 1
						End If
					End If
				End If
			End If
		End If

		NoPage = 0

		If Page < 0 or (Page > MaxPage and MaxPage>=(DEF_MaxJumpPageNum-1)) or (Page > (DEF_MaxJumpPageNum-1) and Page<(MaxPage-DEF_MaxJumpPageNum+1)) Then NoPage = 1
		If (RootID > 0 or LastNum>0 or NoPage = 1) and JMPage < 1 Then
			If Upflag="0" Then
				SQLEndString = SQLEndString & " order by ID DESC"
			Else
				SQLEndString = SQLEndString & " order by ID ASC"
			End If
			If LastNum > 0 Then
				Temp = LastNum
			Else
				Temp = DEF_MaxListNum
			End If
		Else
			If JMPage > 0 Then
				If Upflag="0" Then
					SQLEndString = SQLEndString & " order by ID DESC"
					Upflag="0"
					Temp = (JMPage-1) * DEF_MaxListNum + DEF_MaxListNum
				Else
					SQLEndString = SQLEndString & " order by ID ASC"
					Upflag="1"
					Temp = (JMPage-1) * DEF_MaxListNum + DEF_MaxListNum
				End If
			Else
				If Page < DEF_MaxJumpPageNum Then
					SQLEndString = SQLEndString & " order by ID DESC"
					Upflag="0"
					Temp = Page * DEF_MaxListNum + DEF_MaxListNum
				Else
					SQLEndString = SQLEndString & " order by ID ASC"
					Upflag="1"
					Temp = (MaxPage-Page) * DEF_MaxListNum + LastNumBak
				End If
			End If
		End If

'新代码开始
dim movenum
dim forindex
forindex = get_index("IX_LeadBBS_Announce_ParentID")
If (DEF_UsedDataBase = 0 or DEF_UsedDataBase = 2) and Temp>10 Then

	select case DEF_UsedDataBase
	case 0:
		SQL = sql_select("select id from LeadBBS_Announce " & SQLEndString,Temp)
	case 2:
		movenum = 0
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
		SQL = sql_select("select id from LeadBBS_Announce " & SQLEndString,movenum & "," & Temp)
	end select
	Set Rs = LDExeCute(SQL,0)
	
	
		If LastNum = "" or isNull(LastNum) then LastNum = 0
		LastNum = cCur(LastNum)
	if DEF_UsedDataBase = 0 then
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
	end if
	
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
	SQLEndString_J = Replace(SQLEndString_J,">" & RootID & " ",">=" & Cur_RootID & " ")
	SQLEndString_J = Replace(SQLEndString_J,"<" & RootID & " ","<=" & Cur_RootID & " ")
	SQL = sql_select("select id,ChildNum,Title,FaceIcon,LastTime,Hits,Length,UserName,UserID,RootID,LastUser,NotReplay,GoodFlag,BoardID,TopicType,PollNum,TitleStyle,LastInfo,ndatetime,GoodAssort,NeedValue from LeadBBS_Announce " & SQLEndString_J,DEF_MaxListNum)
	Set Rs = LDExeCute(SQL,0)
Else

'新代码结束
		select case DEF_UsedDataBase
			case 0,2:
				SQL = sql_select("select id,ChildNum,Title,FaceIcon,LastTime,Hits,Length,UserName,UserID,RootID,LastUser,NotReplay,GoodFlag,BoardID,TopicType,PollNum,TitleStyle,LastInfo,ndatetime,GoodAssort,NeedValue from LeadBBS_Announce " & forindex & "" & SQLEndString,Temp)
			case Else
				SQL = sql_select("select id,ChildNum,Title,FaceIcon,LastTime,Hits,Length,UserName,UserID,RootID,LastUser,NotReplay,GoodFlag,BoardID,TopicType,PollNum,TitleStyle,LastInfo,ndatetime,GoodAssort,NeedValue from LeadBBS_Topic " & SQLEndString,Temp)
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
			FirstRootID_2 = cCur(GetData_2(0,0))
			LastRootID_2 = cCur(GetData_2(0,Temp2))
				
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
	

	SQL = "?B=" & GBL_board_ID
	SQL = SQL & "&amp;E=" & EFlag
	If EID > 0 Then SQL = SQL & "&amp;EID=" & EID




	PageSplitString = "<div class=""j_page"">"
	If FirstRootID >= ALL_FirstRootID and Page = 0 Then
	Else
		PageSplitString = PageSplitString & "<a href=""b.asp" & SQL & "&amp;p=0"">1</a>"
		
		if page <> 1 Then
			PageSplitString = PageSplitString & "<a href=""b.asp" & SQL & "&amp;RootID=" & FirstRootID & "&amp;Upflag=1&amp;p=" & Page-1 & """>上页"
			If (Page - DEF_DisplayJumpPageNum) > 0 Then PageSplitString = PageSplitString & "…"
			PageSplitString = PageSplitString & "</a>"
		End If
	End If
	
	DN = DEF_DisplayJumpPageNum
	DotFlag = 0

	If MaxPage > 0 Then
		For1 = Page - DN
		For2 = Page + DN
		If For1 < 0 Then
			For1 = 0
		End If
		If For2 >= MaxPage Then For2 = MaxPage
		'If For2 - For1 < DEF_MaxJumpPageNum and For1 > 1 and For2 > For1 Then For1 = For1 - (DEF_MaxJumpPageNum - ((For2 - For1) + 1))
		If For2 - For1 < DEF_DisplayJumpPageNum*2 and For1 > 1 and For2 > For1 Then For1 = For1 - (DEF_DisplayJumpPageNum*2 - ((For2 - For1) + 1))
		
		If For1 < 0 Then
			For1 = 0
		'ElseIf For1 > 0 Then
		'	PageSplitString = PageSplitString & "<b>…</b>"
		End If

		'If For2 - For1 < DEF_MaxJumpPageNum and For2 < MaxPage and For2 > For1 Then For2 = For2 + (DEF_MaxJumpPageNum - ((For2 - For1) + 1))
		If For2 - For1 < DEF_DisplayJumpPageNum*2 and For2 < MaxPage and For2 > For1 Then For2 = For2 + (DEF_DisplayJumpPageNum*2 - ((For2 - For1) + 1))
		If For2 >= MaxPage Then For2 = MaxPage

		If For2 > 998 Then
			If Page > For1 Then For1 = For1 + 1
			If Page > For1 Then For1 = For1 + 1
			If Page < For2 - DN + 1 Then For2 = For2 - 1
		End If
		For N = For1 to For2
			If N = Page Then
				PageSplitString = PageSplitString & "<b>" & N + 1 & "</b>"
			Else
				If N <> MaxPage Then
					If N <> 0 Then
						If (N-Page) > 0 Then
							PageSplitString = PageSplitString & "<a href=""b.asp" & SQL & "&amp;r=" & LastRootID & "&amp;Upflag=0&amp;p=" & page & "&amp;q=" & N - Page & """>" & N + 1 & "</a>"
						Else
							PageSplitString = PageSplitString & "<a href=""b.asp" & SQL & "&amp;r=" & FirstRootID & "&amp;Upflag=1&amp;p=" & page & "&amp;q=" & Page-N & """>" & N + 1 & "</a>"
						End If
					End If
				End If
			End If
			DotFlag = 2
		Next
		If For2 < MaxPage Then
			'PageSplitString = PageSplitString & "<b>…</b>"
			DotFlag = 1
		End If
	Else
		PageSplitString = PageSplitString & " <b>1</b> "
	End If

	If LastRootID <= ALL_LastRootID Then
	Else
		If page <> MaxPage-1 Then
			PageSplitString = PageSplitString & "<a href=""b.asp" & SQL & "&amp;RootID=" & LastRootID & "&amp;p=" & Page+1 & """>"
			If (Page + DN) < MaxPage Then PageSplitString = PageSplitString & "…"
			PageSplitString = PageSplitString & "下页</a>"
		End If
		PageSplitString = PageSplitString & "<a href=""b.asp" & SQL & "&amp;Upflag=1&amp;Num=1&amp;p=" & MaxPage & """>" & MaxPage + 1 & "</a>"
	End If

	Rs = Temp1
	Rs = Temp2+Rs
	If HaveRootIDFlag = 1 Then Rs = Rs+1

	if MaxPage < 1 Then MaxPage = 0
	'if MaxPage > 0 Then MaxPage = MaxPage + 1
	If ALL_Count < 0 Then ALL_Count = 0
	
	'PageSplitString = PageSplitString & " 共" & ALL_Count &"帖 第" & Page+1 & "/" & MaxPage+1 & "页 每页" & DEF_MaxListNum & "条"
	PageSplitString = PageSplitString & "</div>"

	For1 = 0
	For2 = 0
	If Rs < DEF_MaxListNum and GBL_ShowBottomSure = 0 Then GBL_SiteBottomString = ""

	CALL B_DisplaySplitPageString(PageSplitString,"b_box_none")

	Global_TableHead%>
	<div class="contentbox">
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class="tablebox">
	<%
	Set BoardListClass = New BoardList_HTML_Class
	BoardListClass.Showhead
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
			DisplayAnnounceData_HTML For1,For2,StepValue,GetData_2,0,BoardListClass
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
			DisplayAnnounceData_HTML For1,For2,StepValue,GetData_2,0,BoardListClass
		End If
	End If
	Set BoardListClass = Nothing
	%>
		</table>
	</div>
	<%Global_TableBottom
	CALL B_DisplaySplitPageString(PageSplitString,"b_box_none2")

End Function

Sub DisplayAnnounceData_HTML(For1,For2,StepValue,GetData,AllFlag,obj)

	Dim N,Temp,Temp1
	For N = For1 to For2 Step StepValue
		If AllFlag = 0 or GBL_Board_ID <> cCur(GetData(13,N)) Then
			'GetData(2,n) = Replace(GetData(2,n),"&#60","&lt;")
			If GetData(16,n) <> 1 Then GetData(2,n) = Replace(GetData(2,n) & "","<","&lt;")
			If GetData(16,n) >=60 Then
				GetData(2,n) = "<span class=""grayfont"">帖子等待审核中...</span>"
				GetData(16,n) = 1
			End If
			GetData(17,n) = Replace(Replace(Replace(GetData(17,N) & "","<","&lt;"),chr(13),""),chr(10),"")
			CALL obj.leadbbs(AllFlag,GetData(0,N),GetData(1,N),GetData(2,N),GetData(3,N),GetData(4,N),GetData(5,N),GetData(6,N),Replace(GetData(7,N),"<","&lt;"),GetData(8,N),GetData(9,N),Replace(GetData(10,N),"<","&lt;"),GetData(11,n),GetData(12,N),GetData(13,N),GetData(14,N),GetData(15,N),GetData(16,N),GetData(17,n),GetData(18,N),GetData(19,N),GetData(20,N))
		End If
	Next

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
		select case DEF_UsedDataBase
			case 0,2:
				Set Rs = LDExeCute(sql_select("select id,ChildNum,Title,FaceIcon,LastTime,Hits,Length,UserName,UserID,RootID,LastUser,NotReplay,GoodFlag,BoardID,TopicType,PollNum,TitleStyle,LastInfo,ndatetime,GoodAssort,NeedValue from LeadBBS_Announce " & get_index("IX_LeadBBS_Announce_RootIDBak2") & "where ParentID=0 and RootIDBak in(" & Temp & ") order by ID DESC",Ubound(GetDataTop,2)+1),0)
			case Else
				Set Rs = LDExeCute(sql_select("select id,ChildNum,Title,FaceIcon,LastTime,Hits,Length,UserName,UserID,RootID,LastUser,NotReplay,GoodFlag,BoardID,TopicType,PollNum,TitleStyle,LastInfo,ndatetime,GoodAssort,NeedValue from LeadBBS_Topic where ID in(" & Temp & ") order by ID DESC",Ubound(GetDataTop,2)+1),0)
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

Sub B_DisplaySplitPageString(PageSplitString,css)
%>
	<div class="<%=css%> fire">
		<div class="a_post_image">
			<div class="layer_item">
				<a href="../a/a2.asp?B=<%=GBL_board_ID%>" class="b_post_link"><img src="../images/blank.gif" class="b_post" /></a>
				<div class="layer_iteminfo">
					<ul class="menu_list"><li><a href="../a/a2.asp?B=<%=GBL_board_ID%>">发表新主题</a></li>
					<li><a href="../a/a2.asp?B=<%=GBL_board_ID%>&amp;VoteFlag=yes">发起投票</a></li>
					</ul>
				</div>
			</div>
		</div>
	<%=PageSplitString%>
	</div>
<%
End Sub
%>