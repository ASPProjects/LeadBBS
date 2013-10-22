<!-- #include file=../inc/BBSsetup.asp -->
<!-- #include file=../inc/Board_Popfun.asp -->
<!-- #include file=../inc/Limit_Fun.asp -->
<!-- #include file=inc/Board_fun.asp -->
<!-- #include file=../inc/Fun/ViewOnline_fun.asp -->
<!-- #include file=inc/SmallList.asp -->
<!-- #include file=../inc/Templet/HTML/Normal_0.asp -->
<!-- #include file=../inc/Fun/VierAnc_Fun.asp -->
<!-- #include file=../inc/IncHtm/Boards_Side.asp -->
<!-- #include file=../inc/IncHtm/Boards_Side_Setup2.asp -->
<%
DEF_BBS_HomeUrl = "../"
Dim Boards_dis_assortStr

Dim EFlag,EString,EUrlString,EID,EName,ENumber,EIndex
ENumber = 0
EIndex = 0

Sub DisplayBoard_HTML_MastList(s,num,flag)

	If "?LeadBBS?" = s Then
		Response.Write "全体" & DEF_PointsName(8)
	Else
		If s = "" or s = null Then
			Response.Write flag & "：无"
			Exit Sub
		End If
		Dim ss,n,m
		ss = Split(s,",")
		m = Ubound(ss,1)
		If m >= num Then
			%><%=flag%>：<%
		Else%>
			<%=flag%>：<%
		End If
		For n = 0 to m
			If n >= num Then Exit For
			If n > 0 Then Response.Write ", "
			Response.Write "<a href=""" & DEF_BBS_HomeUrl & "User/LookUserinfo.asp?name=" & urlEncode(ss(n)) & """"
			Response.Write ">" & HtmlEncode(ss(n)) & "</a>"
		Next
		If n >= num and n <= m Then
			%>
				<div class="layer_item" style="display:inline"><span class="layer_item_title"><em>...</em></span>
				<div class="layer_iteminfo">
				<ul class="menu_list">
					<%
			Response.Write "<li><b>更多" & flag & "</b></li>"
			Dim t
			t = n
			For n = t to m
				Response.Write "<li><a href=""" & DEF_BBS_HomeUrl & "User/LookUserinfo.asp?name=" & urlEncode(ss(n)) & """"
				Response.Write ">" & HtmlEncode(ss(n)) & "</a></li>"
			Next
			%>
				</ul>
				</div>
			</div><%
		End If
	End If

End Sub

Function LoginAccuessFul

	GBL_CHK_TempStr = ""
	If GBL_board_ID = 0 and EFlag = 0 Then
		Global_ErrMsg "论坛不存在此版面，请返回首页重新访问。" & VbCrLf
		GBL_SiteBottomString = ""
		Exit Function
	End If
	If GBL_CHK_TempStr<> "" then
		Global_ErrMsg GBL_CHK_TempStr
		GBL_SiteBottomString = ""
	Else
		DisplayAnnouncesSplitPages
	End If

End Function

Function GetActiveUserNumber(BoardID)

	If GBL_board_ID < 1 Then
		GetActiveUserNumber = 0
		Exit Function
	End If
	Dim Rs,tmp
	If isArray(Application(DEF_MasterCookies & "BoardInfo" & BoardID)) = True Then
		Rs = Application(DEF_MasterCookies & "BDOL" & BoardID)
		If Rs >= 0 and Rs <= cCur(Application(DEF_MasterCookies & "ActiveUsers")) Then 
			GetActiveUserNumber = Rs
			Exit Function
		End If
	Else
		GetActiveUserNumber = 0
		Exit Function
	End If
	Set Rs = LDExeCute("select count(*) from LeadBBS_onlineUser where AtBoardID=" & BoardID,0)
	If Rs.Eof Then
		tmp = 0
	Else
		tmp = Rs(0)
		If isNull(tmp) Then tmp = 0
		tmp = cCur(tmp)
	End If
	Rs.Close
	Set Rs = Nothing
	GetActiveUserNumber = tmp
	If isArray(Application(DEF_MasterCookies & "BoardInfo" & BoardID)) = True Then
		Application.Lock
		Application(DEF_MasterCookies & "BDOL" & BoardID) = tmp
		Application.UnLock
	End if

End Function

Sub DisplayBoard_HTML(BoardNum,Blist)

	Dim BoardID,ForumPass,GetData
	Dim N
	Dim BoardClass
	Set BoardClass = New DisplayBoard_HTML_Class
	Dim ShowFlag
	ShowFlag = 0

	Dim CloseAssort,OpenAssort
	CloseAssort = Request.Cookies(DEF_MasterCookies & "clsassort")
	OpenAssort = Request.Cookies(DEF_MasterCookies & "openassort")
	Boards_dis_assortStr = Request.Cookies(DEF_MasterCookies & "dis_assort")
	For N = 0 to BoardNum
		BoardID = Blist(n)
		GetData = Application(DEF_MasterCookies & "BoardInfo" & BoardID)
		If isArray(GetData) = False Then
			ReloadBoardInfo(BoardID)
			GetData = Application(DEF_MasterCookies & "BoardInfo" & BoardID)
		End If

		ForumPass = GetData(7,0)
		If ForumPass <> "" Then ForumPass = "leadbbs"
		GetData(9,0) = cCur(GetData(9,0))
		If GBL_CheckLimitTitle(ForumPass,GetData(9,0),GetData(36,0),GetData(8,0)) = 1 Then
			GetData(20,0) = "已设置为隐藏"
			GetData(3,0) = ""
		End If
		GetData(1,0) = GBL_Board_ID
		GetData(14,0) = GBL_Board_BoardName
		If inStr(OpenAssort,",b" & GBL_Board_ID & ",") > 0 or ((GetBinarybit(GBL_Board_BoardLimit,19) = 0 and (inStr(CloseAssort,",b" & GBL_Board_ID) & ",") = 0)) Then
		'If GetBinarybit(GBL_Board_BoardLimit,19) = 0 Then
			If GetData(8,0) = 0 Then
				If ShowFlag = 0 Then
					Global_TableHead
					ShowFlag = 1
				End If
				CALL BoardClass.DisplayBoard_HTML_Fun(BoardID,GetData(1,0),GetData(0,0),GetData(2,0),GetData(3,0),GetData(4,0),GetData(29,0),GetData(30,0),ForumPass,GetData(19,0),Replace(GetData(20,0),"<","&lt;"),GetData(10,0),GetData(9,0),GetData(14,0),GetData(31,0),GetData(32,0),GetData(21,0),GetData(22,0),GetData(23,0),0,GetData(27,0),Replace(GetData(35,0),"<","&lt;"))
			End If
		Else
			
			If ShowFlag = 0 Then
				Global_TableHead
				ShowFlag = 1
			End If
			CALL BoardClass.DisplayBoard_HTML_Fun_Simple(BoardID,GetData(1,0),GetData(0,0),GetData(2,0),GetData(3,0),GetData(4,0),GetData(29,0),GetData(30,0),ForumPass,GetData(19,0),Replace(GetData(20,0),"<","&lt;"),GetData(10,0),GetData(9,0),GetData(14,0),GetData(31,0),GetData(32,0),GetData(21,0),GetData(22,0),GetData(23,0),0,GetData(27,0),Replace(GetData(35,0),"<","&lt;"))
			'CALL BoardClass.DisplayBoard_HTML_Fun_Simple(BoardID,GetData(1,0),GetData(0,0),GetData(14,0),GetData(18,0),GetData(5,0),GetData(6,0))
		End If
	Next
	BoardClass.DisplayBoard_HTML_Fill
	Set BoardClass = Nothing
	If ShowFlag = 1 Then
		Response.Write "</table></div></div>"
		Global_TableBottom
	End If

End Sub

Sub Boards_CloseAssort

	%>
	<script src="../inc/js/boardlist.js" type="text/javascript"></script>
	<%

End Sub

Sub DisplayBoard(Blist)

	If Blist & "" = "" Then Exit Sub
	Blist = Split(Blist,",")
	If Ubound(Blist,1) < 0 then Exit Sub

	Dim BoardNum
	BoardNum = Ubound(Blist,1)

	If BoardNum = -1 Then
	Else
		Boards_CloseAssort
		CALL DisplayBoard_HTML(BoardNum,Blist)
	End If

End Sub

Sub b_DisplayBoard

	Dim Page,JMPage
	Page = Left(Request.QueryString("p"),14)
	If isNumeric(Page) = 0 or inStr(Page,".") > 0 Then Page = 0
	Page = cCur(Page)
			
	JMPage = Left(Request.QueryString("q"),14)
	If isNumeric(JMPage) = 0 or inStr(JMPage,".") > 0 Then JMPage = 0
	JMPage = Fix(cCur(JMPage))
	If JMPage > DEF_MaxJumpPageNum Then JMPage = 0
			
	If Request.QueryString("Upflag")="1" Then
		Page = Page - JMPage
	Else
		Page = Page + JMPage
	End If

	If GetBinarybit(GBL_Board_BoardLimit,12) = 1 or (Page <= 1) Then
		If isArray(Application(DEF_MasterCookies & "BoardInfo" & GBL_board_ID)) Then DisplayBoard(Application(DEF_MasterCookies & "BoardInfo" & GBL_board_ID)(27,0))
	End If

End Sub

Sub B_Main

	If GetBinarybit(DEF_Sideparameter,17) = 1 Then
	%>
	<script language="JavaScript" type="text/javascript">
	function forum_opt_init()
	{
		var cur="<%=GBL_Board_BoardAssort%>";
		$(".boardnavlist>.user_itemlist ul").hide();
		$("#master_part_" + cur).show();
		$(".swap_collapse").toggleClass("swap_open");
		$("#master_part_" + cur).prev().attr("class","swap_collapse");
	}
	function swap_view(str,sobj)
	{
		$(".swap_collapse").toggleClass("swap_open");
		sobj.className = "swap_collapse";
		$(".boardnavlist>.user_itemlist ul").hide();
		$("#"+str).show();
	}
	function url_to(id)
	{<%if GetBinarybit(DEF_Sideparameter,16) = 0 then%>
		document.location="<%=DEF_BBS_HomeUrl%>b/b.asp?b="+id;
		<%Else%>
		document.location="<%=DEF_BBS_HomeUrl%>b/forum-"+id+"-1.html";
		<%end if%>
	}
	</script>
	<div class="boardnavlist">
		<div class="user_itemlist">
			<div class="navtitle" oncontextmenu="$(this).parent().parent().hide();return false;">版块导航</div>
			<!-- #include file=../inc/incHtm/BoardJump2.asp -->
		</div>
	</div>
	<script>
	forum_opt_init();
	</script>
	<div class="boardnavlist_sider">
	<%
	End If
	If GBL_CHK_TempStr = "" Then
		UpdateOnlineUserAtInfo GBL_board_ID,GBL_Board_BoardName & " " & EString
		If GetBinarybit(GBL_Board_BoardLimit,20) <> 1 Then
			If GBL_B_SubBoard_Flag = 0 Then b_DisplayBoard
		End If
		If GetBinarybit(GBL_Board_BoardLimit,12) = 1 Then
			If GBL_ShowBottomSure = 0 Then GBL_SiteBottomString = ""
		Else
	%>
		<%	If GBL_Board_ID > 0 or EFlag > 0 Then
					%>
		<div class="b_box fire">
			<div class="b_box_nav">
				<ul>
			<%
					Response.Write "<li>"
					If EFlag < 0 Then
						Response.Write "<b>全部</b>"
					Else
						Response.Write "<a href=""b.asp?B=" & GBL_board_ID & """>全部</a>"
					End If
					Response.Write "</li>"
					If GBL_Board_GoodNum > 0 Then
						Response.Write "<li>"
						If EFlag = 0 Then
							Response.Write "<b>精华帖</b>"
						Else
							Response.Write "<a href=""b.asp?B=" & GBL_board_ID & "&amp;E=0"">精华帖</a>"
						End If
						Response.Write "</li>"
					End If
					If isArray(Application(DEF_MasterCookies & "BoardInfo" & GBL_Board_ID & "_TI")) = False Then
						If Application(DEF_MasterCookies & "BoardInfo" & GBL_Board_ID & "_TI") & "" <> "yes" Then ReloadTopicAssort(GBL_Board_ID)
					End If
					If isArray(Application(DEF_MasterCookies & "BoardInfo" & GBL_Board_ID & "_TI")) Then
						%><li><a href="b.asp?B=<%=GBL_board_ID%>&amp;E=1" onclick="ShowOnline('followAssort','swap_assort',2);return false;">
							<span class="swap_ol<%If GetBinarybit(GBL_Board_BoardLimit,18) = 0 Then Response.Write "_close"%>" id="swap_assort"><%
						If EFlag > 0 Then
							Response.Write "<b>专题</b>"
						Else
							Response.Write "专题"
						End If
						%></span></a></li><%
					End If
				%><li><a href="b.asp?B=<%=GBL_board_ID%>&amp;E=1" onclick="ShowOnline('follow0','swap_ol',1);return false;">
					<span class="swap_ol<%If DEF_DisplayOnlineUser = 1 or DEF_DisplayOnlineUser = 3 Then Response.Write "_close"%>" id="swap_ol">在线<%=GetActiveUserNumber(GBL_Board_ID)%>人</span></a></li>
					<li>
					主题: <%=GBL_Board_TopicNum%> / 帖子: <%=GBL_Board_AnnounceNum%></li></ul>
			</div>
			<div class="b_anc_master">
				<%DisplayBoard_HTML_MastList GBL_Board_MasterList,3,DEF_PointsName(8) %>
			</div>
		</div>
		<%
			End If%>
	<script type="text/javascript" language="JavaScript">
	<!--
	function ShowOnline(obj,swap,ol){
		if ($id(obj).style.display!='block'){
			$id(obj).style.display="block";
			if($id(obj).innerHTML=="loading...")
			{
				$id(obj).innerHTML = layer_loadstr;
				getAJAX("b.asp","ol=" + ol + "&b=<%=GBL_Board_ID%>",obj);
			}
			$id(swap).className = "swap_ol";
			}else{
			$id(obj).style.display="none";
			$id(swap).className = "swap_ol_close";
		}
	}
	-->
	</script>
			<%If GetBinarybit(GBL_Board_BoardLimit,18) = 1 Then%>
				<div class="b_box fire" id="followAssort" style="display: block">
			          <%DisplayTopicAssort%>
				</div>
			<%
			Else%>
				<div class="b_box fire" id="followAssort" style="display: none">loading...</div>
			<%
			End If
			If DEF_DisplayOnlineUser = 1 or DEF_DisplayOnlineUser = 3 Then%>
				<div class="b_box fire" id="follow0" style="display: none">loading...</div>
			<%End If%>
			<%If DEF_DisplayOnlineUser = 2 Then%>
				<div class="b_box fire" id="follow0" style="display: block">
			          <%DisplayUserOnline GBL_Board_ID,"../"%>
				</div><%
			End If%>
			<%
			LoginAccuessFul
		End If
	Else
		Global_ErrMsg GBL_CHK_TempStr
	End If
	
	If GBL_CHK_TempStr = "" and GetBinarybit(GBL_Board_BoardLimit,20) = 1 Then
		b_DisplayBoard
	End If
	
	If GetBinarybit(DEF_Sideparameter,16) = 1 Then
	%>
	</div>
	<%
	End If

End Sub

Sub Main

	GBL_CHK_PWdFlag = 0
	GBL_CHK_GuestFlag = 0
	initDatabase
	CheckisBoardMaster
	GBL_CHK_TempStr = ""
	Select Case Request.Form("ol")
		Case "3"
		If GBL_CHK_TempStr = "" Then
			Dim SmallList
			Set SmallList = New Small_List
			SmallList.DisplayAnnouncesSplit
			Set SmallList = Nothing
		End If
		CloseDataBase
		Exit Sub
	Case "1"
		GetStyleInfo
		If GBL_CHK_TempStr = "" Then DisplayUserOnline GBL_Board_ID,"../"
		CloseDataBase
		Exit Sub
	Case "2"
		GetStyleInfo
		If GBL_CHK_TempStr = "" Then DisplayTopicAssort
		CloseDataBase
		Exit Sub
	Case "side":
		Boars_Side_Box("")
		CloseDatabase
		Exit Sub
	End Select

	EFlag = Request.QueryString("E")
	If EFlag = "1" Then
		EFlag = 1
		EString = "专题"
		EUrlString = "&E=1"
	ElseIf EFlag = "0" Then
		EFlag = 0
		EString = "精华帖"
		EUrlString = ""
	Else
		EFlag = -1
		EString = ""
		EUrlString = ""
	End If

	EID = Left(Request.QueryString("EID"),14)
	If isNumeric(EID) = 0 Then EID = 0
	EID = Fix(cCur(EID))
	If EID < 1 Then EID = 0
	If EID > 0 Then EName = GetEName(EID)
	If EName = "" Then EID = 0
	If EID > 0 Then
		EString = "<span class=""navigate_string_step""><a href=""b.asp?B=" & GBL_Board_ID & "&amp;E=1"">专题</a></span><span class=""navigate_string_step"">" & EName & "</span>"
	Else
		If EString <> "" Then EString = "<span class=""navigate_string_step"">" & EString & "</span>"
	End If


	Dim SideFlag,SideNomal
	SideFlag = GetBinarybit(DEF_Sideparameter,3)
	SideNomal = GetBinarybit(DEF_Sideparameter,4)
	SideFlag = Cstr(SideFlag)
	If SideFlag = "0" Then
		SideFlag = "1"
	Else
		SideFlag = "0"
	End If
	GBL_SideFlag = Cstr(SideFlag) & Cstr(SideNomal)

	DEF_GBL_Description = KillHTMLLabel(EString & " " & GBL_Board_BoardName) & " " & DEF_SiteNameString
	BBS_SiteHead DEF_SiteNameString & " - " & KillHTMLLabel(GBL_Board_BoardName),GBL_board_ID,EString

	CheckAccessLimit
	If SideFlag = 1 or GBL_CHK_TempStr <> "" Then
		Boards_Body_Head("")
	Else
		Boards_Body_Head("request" & SideNomal)
	End If


	B_Main

	CloseDataBase
	Boards_Body_Bottom
	If GBL_CHK_TempStr <> "" Then
		If GBL_ShowBottomSure = 0 Then GBL_SiteBottomString = ""
	End If
	SiteBottom

End Sub

Function GetEName(ID)

	Dim TArray,N,Num
	TArray = Application(DEF_MasterCookies & "BoardInfo" & GBL_Board_ID & "_TI")
	If isArray(TArray) = False Then
		ReloadTopicAssort(GBL_Board_ID)
		Exit Function
	End If
	Num = Ubound(TArray,2)
	For N = 0 To Num
		If ID = cCur(TArray(0,N)) Then
			EIndex = N
			EName = TArray(1,n)
			GetEName = EName
			ENumber = cCur(TArray(2,n))
			Exit For
		End If
	Next

End Function

Sub ReloadTopicAssort(BoardID)

	Dim Rs
	Set Rs = LDExeCute("select ID,AssortName,0,0,0 from LeadBBS_GoodAssort where BoardID=" & BoardID & " Order by BoardID,OrderID",0)
	If Not Rs.Eof Then
		Application.Lock
		Application(DEF_MasterCookies & "BoardInfo" & BoardID & "_TI") = Rs.GetRows(-1)
		Application.UnLock
	Else
		Application.Lock
		Set Application(DEF_MasterCookies & "BoardInfo" & BoardID & "_TI") = Nothing
		Application(DEF_MasterCookies & "BoardInfo" & BoardID & "_TI") = "yes"
		Application.UnLock
	End If
	Rs.Close
	Set Rs = Nothing

End Sub

Sub DisplayTopicAssort

	Dim TArray,N,Num,M
	TArray = Application(DEF_MasterCookies & "BoardInfo" & GBL_Board_ID & "_TI")
	If isArray(TArray) = False Then Exit Sub
	Num = Ubound(TArray,2)
	If N < 0 Then Exit Sub
	%>
	<div class="b_assortlist">
	<ul>
	<%
	For N = 0 To Num
		%><li>
			<%
			If EID = TArray(0,N) Then
				Response.Write "<b>"
				Response.Write TArray(1,N)
				Response.Write "</b>"
			Else
				%><a href="b.asp?B=<%=GBL_Board_ID%>&amp;E=1&amp;EID=<%=TArray(0,N)%>"><%=TArray(1,N)%></a>
			<%End If%></li>
		<%
	Next%>
	</ul>
	</div>
	<%

End Sub

Sub Boars_Side_Box_MakeFile(side)

	If side <> "_close" Then	
		Response.Write SideBoard_GetContent
	End If

End Sub

Sub Boars_Side_Box(side)

	Boars_Side_Box_MakeFile(side)

End Sub

Main
%>