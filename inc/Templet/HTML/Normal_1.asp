<%
Class BoardList_HTML_Class

Private Temp,Temp1,Temp2,T1,T2,B_Now,GN,Num,Index

Private GDI,DEF_BTMI,LMT_URLStr,DEF_TCMLN,DEF_DTL,allflg,RewriteFlag,RewriteStr

Public CFlag

Private Sub Class_Initialize

	Dim CloseAssort,OpenAssort
	CloseAssort = Request.Cookies(DEF_MasterCookies & "clsassort")
	OpenAssort = Request.Cookies(DEF_MasterCookies & "openassort")
	
	If inStr(OpenAssort,",foption,") > 0 or (GetBinarybit(DEF_Sideparameter,19) = 1 and inStr(CloseAssort,",foption,") = 0) Then
		CFlag = 1
	Else
		CFlag = 0
	End If

	T2 = ""
	Dim TArray,N
	B_Now = Left(GetTimeValue(DEF_Now),8)
	TArray = Application(DEF_MasterCookies & "BoardInfo" & GBL_Board_ID & "_TI")
	If isArray(TArray) = False Then
		Num = 0
		Redim GN(0,2)
		GN(0,0) = 0
		GN(0,1) = ""
		GN(0,2) = ""
	Else
		Num = Ubound(TArray,2)
		Redim GN(Num,2)
		For N = 0 To Num
			GN(N,0) = TArray(0,N)
			GN(N,1) = TArray(1,N)
			GN(N,2) = KillHTMLLabel(TArray(1,N))
		Next
	End If
	
	GDI = GBL_DefineImage
	DEF_BTMI = DEF_BBS_TOPMinID
	LMT_URLStr = LMT_UrlEndString
	DEF_TCMLN = DEF_TopicContentMaxListNum
	DEF_DTL = DEF_BBS_DisplayTopicLength
	allflg = 0
	Index = 0
	If GetBinarybit(DEF_Sideparameter,16) = 0 Then
		RewriteFlag = 0
	Else
		RewriteFlag = 1
		If LMT_URLStr <> "" Then
			If inStr(LMT_URLStr,"&q=") or inStr(LMT_URLStr,"&amp;q=") Then
				If Ubound(Split(LMT_URLStr,"=")) <> 1 Then RewriteFlag = 0
			Else
				RewriteFlag = 0
			End If
		End If
		If RewriteFlag = 1 and LMT_URLStr <> "" Then
			LMT_URLStr = Replace(LMT_URLStr,"&q=","-")
			LMT_URLStr = Replace(LMT_URLStr,"&amp;q=","-")
			If LMT_URLStr = "-1" then LMT_URLStr = ""
		End If
	End If

End Sub

Public Sub Showhead

	Response.Write "<tr class=""tbhead"">"
	%>	<td width="30" class="forum_options<%If CFlag = 1 Then Response.Write "_sim"%>" title="基本/详细" onclick="forum_options(this);">&nbsp;
	<script src="<%=DEF_BBS_HomeUrl%>inc/js/boardlist.js" type="text/javascript"></script>
	<script>
	var foption=<%
			If CFlag = 1 Then
				Response.Write "1"
			Else
				Response.Write "0"
			End If%>;
	function forum_options(j)
	{
		j.className=(j.className=='forum_options')?'forum_options_sim':'forum_options';
		if(j.className=='forum_options')
		{$id('table_options').className='tablebox table_options';
		$("#forum_linehead").attr("colspan","3");}
		else
		{$id('table_options').className='tablebox table_options_sim';
		$("#forum_linehead").attr("colspan","1");}
		foption=(foption==0)?1:0;
		LD.blist.assort_click('foption',foption,"none");
	}
	</script>
	</td>
	<%Response.Write "	<td><div class=""value"">主题</div></td>"
	Response.Write "	<td width=""110"" class=""author""><div class=""value"">作者/回复人</div></td>"
	Response.Write "	<td width=""74"" class=""hits""><div class=""value"">回复/阅读</div></td>"
	'Response.Write "	<td width=""110"" class=""reply""><div class=""value"">最后更新</div></td>"
	Response.Write "</tr>"

End Sub

Public Sub leadbbs(AllFlag,TopicID,ChildNum,Title,FaceIcon,G6,G7,G8,G9,G10,G11,G13,G14,G15,G16,G17,G18,G19,G20,G21,G22,NeedValue)

	G11 = cCur(G11)
	G18 = cCur(G18)
	G10 = cCur(G10)
	G7 = cCur(G7)
	If allflg = 1 and G11 < DEF_BTMI and AllFlag <> -1 and EFlag < 0 Then
		Response.Write "<tr class=""tbhead2""><td><div class=""value"">&nbsp;</div></td><td colspan="""
		If CFlag = 1 Then
			Response.Write "1"
		Else
			Response.Write "3"
		End If
		Response.Write """ id=""forum_linehead""><div class=""value"">普通主题</div></td></tr>"
		allflg = 2
	End If
	%>
	<tr class="b_list" onmouseover="this.className='b_list_active';" onmouseout="this.className='b_list';"><td align="center" class="tdcontent">
	<%
	If AllFlag = 1 or AllFlag = 2 Then
		If AllFlag = 1 Then
			Temp = "alltop"
			Temp2 = "总置顶"
		Else
			Temp = "parttop"
			Temp2 = "区置顶"
		End If
		allflg = 1
	Else
		If G11 >= DEF_BTMI and AllFlag <> -1 and EFlag < 0 Then
			Temp = "intop"
			Temp2 = "版面置顶"
			allflg = 1
		Else
			If G17 = 80 Then
				'If G18 >= 20 Then
				'	Temp = "vthot"
				'	Temp2 = "热门投票"
				'Else
					Temp = "vt"
					Temp2 = "投票"
				'End If
			Else
				If G14 = 1 Then
					Temp = "lock"
					Temp2 = "锁定"
				ElseIf ChildNum >= 20 Then
					Temp = "hot"
					Temp2 = "热门帖"
				Else
					If Left(G21,8) = B_Now Then
						Temp = "tpcnew"
						Temp2 = "新帖"
					Else
						Temp = "tpc"
						Temp2 = "topic"
					End If
				End If
			End If
		End If
	End If
	If RewriteFlag = 0 Then
		RewriteStr = "a.asp?B=" & G16 & "&amp;ID=" & TopicID & LMT_URLStr
	Else
		RewriteStr = "topic-" & G16 & "-" & TopicID & "-1" & LMT_URLStr & ".html"
	End If
	Response.Write "<a href=""../a/" & RewriteStr & """ title=""" & G8 & """ target=""_blank"">"
	Response.Write "<img src=""../images/state/" & GDI & Temp & ".gif"" title=""" & Temp2 & """ alt=""" & Temp2 & """ /></a>"
	Response.Write "</td><td class=""tdcontent""><span class=""word-break-all"">"

	If ChildNum > 0 Then
		Response.Write "<img src=""../images/" & GDI & "clsExpand.gif"" id=""LeadImg" & TopicID & """ class=""b_getlist absmiddle"" onclick=""Show2('Lead" & TopicID & "','LeadImg" & TopicID & "'," & TopicID & ")"" style=""cursor: pointer"" alt=""展开/收起"" />"
	Else
		Response.Write "<img src=""../images/" & GDI & "Expand_blank.gif"" class=""b_getlist absmiddle"" />"
	End If

	If G8 > 1024 Then
		G8 = "主题内容：" & Fix(G8/1024) & "KB"
	Else
		G8 = "主题内容：" & G8 & "字节"
	End If
	If G17 <> 39 and G20 <> "" Then G8 = G8 & VbCrLf & "最后回复：" & HtmlEncode(G20)
	
	G8 = G8 & VbCrLf & "发表时间：" & ConvertSimTimeString(Mid(G21,1,4) & "-" & Mid(G21,5,2) & "-" & Mid(G21,7,2) & " " & Mid(G21,9,2) & ":" & Mid(G21,11,2))

	Temp1 = Fix((ChildNum+1)/DEF_TCMLN)
	If Temp1 < ((ChildNum+1)/DEF_TCMLN) Then Temp1 = Temp1 + 1
	If DEF_DTL < 255 Then '255长度忽略判断 完整显示
		If ChildNum >= DEF_TCMLN Then
			Temp = DEF_DTL - Len(Temp1 & "") - 3
		Else
			Temp = DEF_DTL
		End If
	
		If ccur(G15) = 1 Then Temp = Temp - 3
		If G17 <> 80 and G17 <> 54 and G17 <> 114 and G17 <> 49 and G17 <> 109 and G18 > 0 Then Temp = Temp - 2
	End If
	If G22 > 0 Then
		T2 = ""
		For T1 = 0 To Num
			If GN(T1,0) = G22 Then
				If DEF_DTL < 255 Then
					T2 = StrLength(GN(T1,2))
					If T2 <= 14 Then
						Temp = Temp - (T2 + 2)
						T2 = GN(T1,1)
					Else
						Temp = Temp - 18
						T2 = LeftTrue(GN(T1,2),11) & "..."
					End If
				Else
					T2 = GN(T1,1)
				End If
				Exit For
			End If
		Next
	Else
		T2 = ""
	End If

If DEF_DTL < 255 Then '255长度忽略判断 完整显示
	If G19 = 1 Then
		If StrLength(Title) > Temp Then Title = LeftTrueHTML(Title,Temp-3)
	Else
		If Strlength(Title) > Temp Then Title = LeftTrue(Title,Temp-3) & "..."
	End If
End If

	If G9 = "[LeadBBS]" Then G9 = "系统"
	Dim old_TopicID
	old_TopicID = TopicID
	If G17 = 39 Then
		Temp = TopicID
		TopicID = NeedValue
		NeedValue = Temp
		Temp = G16
		If isNumeric(G20) = 0 Then G20 = 0
		G16 = cCur(G20)
		G20 = Temp
	End If

	If GBL_BoardMasterFlag >= 5 and ((GBL_Board_ID = cCur(G16) and G17 <> 39) or (GBL_Board_ID = G20 and G17 = 39)) Then
	%>
	<span class="layerico"><input class="fmchkbox" type="checkbox" name="ids" id="ids<%=Index%>" value="<%
		If G17 = 39 Then
			Response.Write NeedValue
		Else
			Response.Write TopicID
		End If%>" onclick="delbody_view(this);" /></span><%
		Index = Index + 1
	End If

	If FaceIcon > 0 Then Response.Write "<img src=""../images/" & GBL_DefineImage & "bf/face" & FaceIcon & ".gif"" class=""absmiddle"" alt=""表情"" /> "
	If G17 <> 80 and G17 <> 54 and G17 <> 114 and G17 <> 49 and G17 <> 109 and G18 > 0 Then Response.Write "<img src=""../images/" & GDI & "TC/" & G18 & ".gif"" class=""absmiddle"" alt="""" />"
	If T2 <> "" Then Response.Write "<a href=""../b/b.asp?B=" & G16 & "&amp;E=1&amp;EID=" & G22 & """><span class=""subjectfont"">[" & T2 & "]</span></a>"


	If RewriteFlag = 0 Then
		RewriteStr = "a.asp?B=" & G16 & "&amp;ID=" & TopicID & LMT_URLStr
	Else
		RewriteStr = "topic-" & G16 & "-" & TopicID & "-1" & LMT_URLStr & ".html"
	End If
	Response.Write "<a href=""../a/" & RewriteStr & """ title=""" & G8 & """ class=""visit"">"
	If G19 = 0 Then
		Response.Write HtmlEncode(Title)
	Else
		Response.Write DisplayAnnounceTitle(Title,G19)
	End If
	Response.Write "</a></span>"

	If ChildNum >= DEF_TCMLN Then
		If RewriteFlag = 0 Then
			CALL pagesplit("../a/a.asp?b=" & G16 & "&amp;id=" & TopicID & "",Temp1)
		Else
			CALL pagesplit("../a/topic-" & G16 & "-" & TopicID & "",Temp1)
		End If
		'Response.Write " <a href=""../a/a.asp?B=" & G16 & "&amp;ID=" & TopicID & "&amp;AUpflag=1&amp;ANum=1" & LMT_URLStr & """ class=""page"">[" & Temp1 & "]</a>"
	End If
	If RewriteFlag = 0 Then
		RewriteStr = "a.asp?B=" & G20 & "&amp;ID=" & NeedValue & LMT_URLStr
	Else
		RewriteStr = "topic-" & G20 & "-" & NeedValue & "-1" & LMT_URLStr & ".html"
	End If
	If G17 = 39 Then Response.Write " <a href=""../a/" & RewriteStr & """><span class=""subjectfont""><span class=""grayfont"">[镜像]</span></span></a>"
	
	If Left(G21,8) = B_Now or Left(G6,8) = B_Now Then Response.Write "<img src=""../images/new.gif"" class=""absmiddle"" alt=""新更新"" />"

	If ccur(G15) = 1 Then Response.Write "<img src=""../images/" & GDI & "jh1.GIF"" title=""精华帖子"" class=""absmiddle"" alt=""精华"" />"

	'If G17 <> 39 and G20 <> "" Then Response.Write "<br /><span class=""grayfont note"">" & HtmlEncode(G20) & "</span>"
	Response.Write "<div id=""Lead" & old_TopicID & """ class=""b_smalllist"" style=""display: none""></div>"
	Response.Write "</td><td class=""tdcontent author"">" & VbCrLf
	If G10 > 0 Then
		Response.Write "<a href=""../User/LookUserInfo.asp?ID=" & G10 & """ class=""postuser"">" & HtmlEncode(G9) & "</a>"
	Else
		Response.Write "<span class=""postuser"">" & HtmlEncode(G9) & "</span>"
	End If
	Response.Write "<br />"
	If ChildNum = 0 Then
		Response.Write "<span class=""lastuser"">- - -</span>"
	Else
		If G13 = "" or G13 = null Then
			If G17 = 39 Then
				Response.Write G9
			Else
				Response.Write "<a href=""../User/LookUserInfo.asp?ID=" & G10 & """ class=""lastuser"">" & HtmlEncode(G9) & "</a>"
			End If
		Else
			If G13 <> "游客" Then
				Response.Write "<a href=""../User/LookUserInfo.asp?name=" & urlEncode(G13) & """ class=""lastuser"">" & HtmlEncode(G13) & "</a>"
			Else
				Response.Write "<span class=""lastuser"">" & HtmlEncode(G13) & "</span>"
			End If
		End If
	End If
	Response.Write "</td><td class=""tdcontent hits"">" & VbCrLf

	If G18 = null Then G18 = 0
	If G17 = 80 Then
		Response.Write "共" & G18 & "票"
	Else
		If G7 >= 10000 Then G7 = "<span class=""bluefont"" title=""人气:" & G7 & """>" & Fix(G7/10000) & "</span>万"
		If ChildNum >= 10000 Then ChildNum = "<span class=""bluefont"" title=""回帖:" & ChildNum & """>" & Fix(ChildNum/10000) & "</span>万"
		Response.Write "<em>" & ChildNum & "/" & G7 & "</em>"
	End If
	If Left(G6,8) = B_Now Then
		G6 = "<span class=""redfont"">" & ConvertSimTimeString(Mid(G6,1,4) & "-" & Mid(G6,5,2) & "-" & Mid(G6,7,2) & " " & Mid(G6,9,2) & ":" & Mid(G6,11,2)) & "</span>"
	Else
		G6 = ConvertSimTimeString(Mid(G6,1,4) & "-" & Mid(G6,5,2) & "-" & Mid(G6,7,2) & " " & Mid(G6,9,2) & ":" & Mid(G6,11,2))
	End If

	Response.Write "<br /><em title=""最后更新"">" & G6 & "</em>"
	Response.Write "</td>"

	'Response.Write "<td class=""tdcontent reply""></td></tr>"
	'Response.Write "<tr><td></td><td colspan=""3""><div id=""Lead" & old_TopicID & """ class=""b_smalllist"" style=""display: none""></div></td></tr>" & VbCrLf

End Sub

Private Sub pagesplit(url,max)

	'CALL pagesplit("../a/a.asp?b=" & G16 & "&amp;id=" & TopicID & "",Temp1)
	if max < 2 then exit sub
	dim n
	Response.Write "<span class=""spage""><span class=""ps""> [ </span><span class=""page"">"
	for n = 2 to 4
		if n > max then exit for
		If RewriteFlag = 0 Then
			Response.Write " <a href=""" & url & "&Aq=" & n & LMT_URLStr & """>" & n & "</a>"
		Else
			Response.Write " <a href=""" & url & "-" & n & LMT_URLStr & ".html"">" & n & "</a>"
		End If
	next
	if max > 5 then Response.Write "..."
	if max > 4 then
		If RewriteFlag = 0 Then
			Response.Write " <a href=""" & url & "&amp;AUpflag=1&amp;ANum=1" & LMT_URLStr & """>" & max & "</a>"
		Else
			Response.Write " <a href=""" & url & "-" & max & LMT_URLStr & ".html"">" & max & "</a>"
		End If
	End If
	Response.Write " </span><span class=""pe""> ] </span></span>"
		
End Sub

End Class
%>