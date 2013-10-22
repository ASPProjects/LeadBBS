<%
Const DEF_HubbleDatabase = "Data Source=127.0.0.1;Initial Catalog=big;"
Const DEF_HubbleTable = "big"
Const DEF_HubbleHightlightColumn = "title"
Const DEF_HubbleMaxPage = 1000

Class hubblesearch_class

	private keywords,key_board,key_userid,key_user,key_mode
	private curPage,MaxPage,recCount,executeTime,PageSize,GetData,RewriteFlag
	private key_order,key_onlytitle,key_onlygood
	
	Private Sub Class_Initialize
		
		If GetBinarybit(DEF_Sideparameter,16) = 0 Then
			RewriteFlag = 0
		Else
			RewriteFlag = 1
		End If

		PageSize = DEF_TopicContentMaxListNum
		keywords = search_get("key",255)
		key_board = search_get("key_board",20)
		if key_board = "" then key_board = search_get("BoardID2",20)
		key_userid = search_get("key_userid",20)
		key_user = search_get("key_user",32)
		key_mode = search_get("key_mode",1)
		curPage = search_get("page",5)
		
		key_order = search_get("key_order",1)
		key_onlytitle = search_get("key_onlytitle",1)
		key_onlygood = search_get("key_onlygood",1)
				
		if isNumeric(key_board) = 0 then key_board = 0
		key_board = fix(ccur(key_board))
		if isNumeric(key_userid) = 0 then key_userid = 0
		key_userid = fix(ccur(key_userid))
		if key_mode <> "0" and key_mode <> "1" and key_mode <> "2" then key_mode = 0
		key_mode = ccur(key_mode)
		if isNumeric(curPage) = 0 then curPage = 0
		curPage = fix(ccur(curPage))
		if curPage < 1 then curPage = 1
		
		If key_order = "1" Then
			key_order = 1
		ElseIf key_order = "2" Then
			key_order = 2
		Else
			key_order = 0
		End If
		
		If key_onlytitle = "1" Then
			key_onlytitle = 1
		Else
			key_onlytitle = 0
		End If
		
		If key_onlygood = "1" Then
			key_onlygood = 1
		Else
			key_onlygood = 0
		End If
		
		search_getUserInfo
		
		search_hubbleform
		
		search_hubblesearch
	
	end sub
	
	private function search_get(name,max)
	
		dim str
		str = left(request.querystring(name),max)
		if str = "" then str = left(request.form(name),max)
		search_get = str
		
	end function
	
	private sub search_getUserInfo
	
		dim rs
		if key_userid <=0 then key_userid = 0
		if key_userid > 0 and key_user = "" then
			set rs = ldexecute("select username from leadbbs_user where id=" & key_userid,0)
			if not rs.eof then
				key_user = rs(0)
			else
				key_userid = 0
			end if
			rs.close
			set rs = nothing
		end if
		if key_userid = 0 and key_user <> "" then
			set rs = ldexecute("select id from leadbbs_user where username='" & replace(key_user,"'","''") & "'",0)
			if not rs.eof then
				key_userid = ccur(rs(0))
			else
				key_user = ""
			end if
			rs.close
			set rs = nothing
		end if
		
	end sub
	
	private sub search_hubbleform
	
		dim viewed
		if key_mode > 0 or key_board>0 or key_userid>0 or key_order > 0 or key_onlytitle>0 or key_onlygood>0 then
			viewed = 1
		else
			viewed = 0
		end if
		%>
	<script language=javascript>
	var ValidationPassed = true;
	function submitonce(theform)
	{	
		
		if(theform.key.value=="")
		{
			alert("������Ҫ���������ݣ�\n");
			ValidationPassed = false;
			theform.key.focus();
			return;
		}
		else
		{ValidationPassed = true;
		}
		submit_disable(theform);
	}
	
	function delbody_view(obj)
	{
		layer_create("anc_msgbody");
		$id('anc_msgbody').innerHTML="<div class=ajaxbox><ul><li>����������������������ͬ�ִ�֮����һ���ո�������������������У�<span class=grayfont>����</span>�����Һ��У�<span class=grayfont>������</span>���������룢<span class=grayfont>���� ������</span>��</li><li>Ĭ����෵��<%=DEF_HubbleMaxPage%>ҳ��¼��������ʽѡ��Ϊ�����򣢾���ͻ��ҳ�����ƣ�</li></ul></div>";
		layer_view('',obj,'','','anc_msgbody','','',0,'',0,0);
	}
	</script>
			<form name="searchform" id="searchform" action=Search.asp onSubmit="submitonce(this);return ValidationPassed;">
			
				<div class=value2>
				<input value="<%=htmlencode(keywords)%>" type="text" name=key size=22 maxlength=255 class='fminpt input_3 searchkey'><input name=submit2 type=submit value="����" class="fmbtn btn_2 searchbtn"><%
				If viewed = 0 then%>
				<a href=javascript:; class=greenfont onclick="this.style.display='none';$id('search_advance').style.display='block';">�߼�����</a><%
				end if%>
				<span class="layerico">
				<a href=javascript:; class=greenfont onclick="delbody_view(this);">
				��������
				</a>
				</span>
				</div>
				
				<div id="search_advance"<%
				If viewed = 0 then response.write " style=""display:none;"""
				%>>
					<div class=value2>������Χ��
						<input name=key_mode class=fmchkbox type=radio value=0<%
						If key_mode = 0 Then Response.Write " checked"
						%>>ȫ��
						<input name=key_mode class=fmchkbox type=radio value=1<%
						If key_mode = 1 Then Response.Write " checked"
						%>>������
						<input name=mode class=fmchkbox type=radio value=2<%
						If key_mode = 2 Then Response.Write " checked"%>>������
						
					</div>
					<div class=value2 id="search_boardlist">
					�ض���飺<!-- #include file=../../inc/incHTM/BoardForMoveList.asp -->	
					<label>
					<input class="fmchkbox" type="checkbox" name="key_onlytitle" value="1"<%
						If key_onlytitle=1 then
						%> checked="checked"<%
						End If%> />������
					</label>
					<label>
					<input class="fmchkbox" type="checkbox" name="key_onlygood" value="1"<%
						If key_onlygood=1 then
						%> checked="checked"<%
						End If%> />�޾���
					</label>								
					</div>
					<script language=javascript>
						function search_initboard()
						{
							var child=$id("search_boardlist").getElementsByTagName("SELECT");
							if(child[0])
							{
								SelectItemByValue(child[0], "<%=key_board%>");
							}
						}
						search_initboard();
					</script>
					<div class=value2>
					ָ���û���
					<input value="<%=htmlencode(key_user)%>" type="text" name=key_user size=30 maxlength=255 class='fminpt input_2'>
					�� �û�ID��
					<input value="<%if key_userid > 0 then response.write key_userid%>" type="text" name=key_userid size=30 maxlength=255 class='fminpt input_1'>
					</div>
					<div class=value2>
					����ʽ��
					
						<input name=key_order class=fmchkbox type=radio value=0<%
						If key_order = 0 Then Response.Write " checked"
						%>>��������
						<input name=key_order class=fmchkbox type=radio value=1<%
						If key_order = 1 Then Response.Write " checked"
						%>>��������
						<input name=key_order class=fmchkbox type=radio value=2<%
						If key_order = 2 Then Response.Write " checked"%>>����
					</div>
				</div>
			</form>
		<%
	
	end sub
	
	private sub search_hubblesearch
	
		if keywords = "" then exit sub
		'curPage,MaxPage,recCount,executeTime
		
		'select between @begin to @end * from bbs where title contains @matchString" + sortstr + " order by id desc
		dim sql,matchString,OrderStr
		matchString = ""
		if key_mode = 1 then
			matchString = matchString & " title contains @matchString"
		elseif key_mode = 2 then
			matchString = matchString & " content contains @matchString"
		else
			matchString = matchString & " (title contains @matchString or content contains @matchString)"
		end if
		
		if key_onlytitle > 0 then
			matchString = matchString & " and parentid=0"
		end if
		
		if key_onlygood > 0 then
			matchString = matchString & " and goodflag>0"
		end if
		
		if key_board > 0 then
			matchString = matchString & " and boardid=" & key_board
		end if
		
		if key_userid > 0 then
			matchString = matchString & " and userid=" & key_userid
		end if
		
		If key_order = 1 Then
			OrderStr = " order by id asc"
		elseif key_order = 2 Then
			OrderStr = ""
		else
			OrderStr = " order by id desc"
		end if
		
		
		If key_order <> 2 and curPage > DEF_HubbleMaxPage Then curPage = DEF_HubbleMaxPage
		
		sql = "select between " & PageSize*(curPage-1) & " to " & PageSize*curPage-1 & " ID,Title,Content,BoardID,BoardName,UserID,UserName,ndatetime,BoardLimit,OtherLimit,HiddenFlag,ForumPass,ParentID,ChildNum,Hits,GoodFlag,TopicType,NeedValue,TitleStyle,GoodAssort,AssortName from " & DEF_HubbleTable & " where" & matchString & OrderStr
		'response.write sql
		
		dim hubbleObj		
		dim outstr,recNum
		Set hubbleObj = CreateObject("leadbbs.forhubble")
		GetData = hubbleObj.GetData(DEF_HubbleDatabase,sql,keywords,"-1,1,2",DEF_HubbleTable,DEF_HubbleHightlightColumn,256,outstr)
		set hubbleObj = Nothing
		GBL_DBNum = GBL_DBNum + 1
		recNum = Ubound(GetData)
		
		if recNum < 0 then
			%>
			<div class="title">û���ҵ��� ��<span class=redfont><%=htmlencode(keywords)%></span>�� ��ص����ݡ�
			</div>
			<div class="value2">���鳢�������ʻ��������ʡ�
			</div>
			<%
			exit sub
		end if
		outstr = Split(outstr,"|")
		If Ubound(outstr) > 1 then
			recCount = outstr(0)
			if isNumeric(recCount) = 0 then recCount = 0
			recCount = ccur(recCount)
			executeTime = outstr(1)
			MaxPage = fix(recCount/PageSize)
			If (MaxPage mod PageSize) > 0 then MaxPage = MaxPage + 1
			If key_order <> 2 and MaxPage > DEF_HubbleMaxPage Then MaxPage = DEF_HubbleMaxPage
		end if
		set hubbleObj = nothing
		
		Dim url
		url = "search.asp?key=" & urlencode(keywords)
		if key_mode > 0 then url = url & "&key_mode=" & key_mode
		if key_board > 0 then url = url & "&key_board=" & key_board
		if key_userid > 0 then url = url & "&key_userid=" & key_userid
		if key_user <> "" then url = url & "&key_user=" & urlencode(key_user)
		if key_order > 0 then url = url & "&key_order=" & key_order
		if key_onlytitle > 0 then url = url & "&key_onlytitle=" & key_onlytitle
		if key_onlygood > 0 then url = url & "&key_onlygood=" & key_onlygood
		call pagelist(url,MaxPage,curPage,"Total " & recCount & " results (" & executeTime & " ms)","")
		
		call viewGetData(0,recNum,1)
		
		call pagelist(url,MaxPage,curPage,"Total " & recCount & " results (" & executeTime & " ms)","")
	
	end sub
	
	Private function pagelist(url,num,curp,result,ajaxobj)
	
		dim n
		%>
		<div class="clear"></div>
		<div class="searchresult">
		<div class="j_page">
		<%
		if curp > 1 then%>
		<a href="<%=url%>&page=<%=curp-1%>"<%if ajaxobj <> "" Then
				%> onclick="getAJAX(this.href+'&AjaxFlag=1&jsflag=1','','<%=ajaxobj%>');return(false);"<%
			end if%> class="previous">Previous</a>
		<%
		end if
		
		for n = curp-4 to curp+4
			if n >=1 and n <= num then
				if n <> curp then
		%><a href="<%=url%>&page=<%=n%>"<%
			if ajaxobj <> "" Then
				%> onclick="getAJAX(this.href+'&AjaxFlag=1&jsflag=1','','<%=ajaxobj%>');return(false);"<%
			end if
		%>><%
		if curp-4 > 1 and n= curp-4 then Response.Write "..."
		if curp+4 < num  and n = curp+4 then
			Response.Write n & "..."
		else
			response.Write n
		end if
		%></a><%
				else
		%><b class=curpage><%=n%></b><%
				end if
			end if
		next
		
		
		if curp < num then%>
		<a href="<%=url%>&page=<%=curp+1%>"<%if ajaxobj <> "" Then
				%> onclick="getAJAX(this.href+'&AjaxFlag=1&jsflag=1','','<%=ajaxobj%>');return(false);"<%
			end if%> class="next">Next</a>
		<%
		end if
		
		if result <> "" then
			%>
			<b class="total"><%=result%></b>
		<%end if
		%>
		</div>
		</div>
		<%
	
	end function
	
	private Sub viewGetData(For1,For2,AllFlag)
	
		Dim N,Temp,Temp1,ForumPass,BoardLimit,OtherLimit,HiddenFlag,BoardName
		'ID,Title,Content,BoardID,BoardName,UserID,UserName,ndatetime,BoardLimit,OtherLimit,HiddenFlag,ForumPass,ParentID,ChildNum,Hits,GoodFlag,TopicType,NeedValue,TitleStyle,GoodAssort,AssortName
		dim ID,Title,Content,BoardID,UserID,UserName,ndatetime,ParentID,ChildNum,Hits,GoodFlag,TopicType,NeedValue,TitleStyle,GoodAssort,AssortName
		
		dim re
		set re = New RegExp
		re.Global = True
		re.IgnoreCase = True
		LMT_WidthStr = "100%"
		%>
		<div class="search_content">
		<%
		For N = For1 to For2
			ID = cCur(GetData(N,0))
			Title = GetData(N,1)
			Content = GetData(N,2)
			BoardID = cCur(GetData(N,3))
			BoardName = GetData(N,4)
			UserID = cCur(GetData(N,5))
			UserName = GetData(N,6)
			ndatetime = GetData(N,7)
			BoardLimit = GetData(N,8)
			OtherLimit = GetData(N,9)
			HiddenFlag = GetData(N,10)
			ForumPass = GetData(N,11)
			ParentID = cCur(GetData(N,12))
			ChildNum = cCur(GetData(N,13))
			Hits = cCur(GetData(N,14))
			GoodFlag = cCur(GetData(N,15))
			TopicType = cCur(GetData(N,16))
			NeedValue = cCur(GetData(N,17))
			TitleStyle = cCur(GetData(N,18))
			GoodAssort = cCur(GetData(N,19))
			AssortName = GetData(N,20)
			
			Response.Write "<div class=clear></div><hr class=splitline>"
	
			Response.Write "<div class=valuetitle><span class=fontzi><a href="""
			If RewriteFlag = 1 Then
				Response.Write DEF_BBS_HomeUrl & "a/topic-" & BoardID & "-" & ID & "-1.html"
			Else
				Response.Write DEF_BBS_HomeUrl & "a/a.asp?B=" & BoardID & "&ID=" & ID
			End If
			If AllFlag = 1 Then
				Response.Write """ target=_blank class=bluefont>"
			Else
				Response.Write """ class=bluefont>"
			End If
	
			Temp1 = Fix((ChildNum+1)/DEF_TopicContentMaxListNum)
			If ((ChildNum+1) mod DEF_TopicContentMaxListNum) > 0 Then Temp1 = Temp1 + 1

			If GBL_NoneLimitFlag = 0 and GBL_CheckLimitTitle(ForumPass,BoardLimit,OtherLimit,HiddenFlag) = 1 Then
				Title = "<span calss=grayfont>�����ӱ���������Ϊ����</span>"
				TitleStyle = 1
			End If
			If GBL_CheckLimitContent(ForumPass,BoardLimit,OtherLimit,HiddenFlag) = 1 Then Content = "<span calss=grayfont>�����������������ư��棬��������鿴</span>"
	
			If TopicType <> 80 and TopicType <> 0 Then Content = "<span calss=grayfont>�����������������ƣ���������鿴...</span>"
	
			If left(Title,3) = "re:" and Title <> "re:" Then Title = Mid(Title,4)

			If Title = "" Then Title = "����"
			Response.Write Title
			Response.Write "</a></span>"
	
			If ChildNum>=DEF_TopicContentMaxListNum Then
				CALL pagesplit(BoardID,ID,Temp1)
				'Response.Write " [<a href=" & DEF_BBS_HomeUrl & "a/a.asp?B=" & BoardID & "&ID=" & ID & "&AUpflag=1&ANum=1>" & Temp1 & "</b></a>]"
			End If
	
			If ccur(GoodFlag) = 1 Then
				Response.Write "<img src=../images/" & GBL_DefineImage & "jh1.GIF border=0 title=�������� align=absbottom>"
			End If
	
			If BoardName <> "" Then Response.Write " <span class=grayfont>-</span> <a href=../b/b.asp?b=" & BoardID & "><span class=greenfont>" & BoardName & "</span></a>"
			If GoodAssort > 0 Then
				Response.Write " <span class=grayfont>-</span> "
				Response.Write "[<a href=../b/b.asp?B=" & BoardID & "&E=1&EID=" & GoodAssort & "><span class=greenfont>" & AssortName & "</span></a>]"
			end if
			Response.Write "</div>"
	
			Response.Write "<div class=value>"
			Response.Write Content
			response.write "</div>"
			Response.Write "<div class=value2><span class=grayfont>���ߣ�"
			If UserID > 0 Then
				Response.Write "<a href=../User/LookUserInfo.asp?ID=" & UserID & "><span class=greenfont>" & htmlencode(UserName) & "</span></a>"
			Else
				Response.Write htmlencode(UserName)
			End If
	
			Temp = RestoreTime(NDateTime)
			If DateDiff("d",Temp,DEF_Now)<1 Then
				Response.Write " ������ <span class=redfont>" & Temp & "</span>"
			Else
				Response.Write " ������ " & Temp
			End If
			Response.Write "</span></div>"
			If N = For2 Then Response.Write "<hr class=splitline>"
		Next
		Response.Write "</div>" & VbCrLf
	
	End Sub
	
	Private Sub pagesplit(boardid,id,max)

	if max < 2 then exit sub
	dim n
	Response.Write " [  <span class=""page"">"
	for n = 2 to 4
		if n > max then exit for
		If RewriteFlag = 0 Then
			Response.Write " <a href=""" & DEF_BBS_HomeUrl & "a/a.asp?B=" & BoardID & "&ID=" & ID & "&Aq=" & n & """>" & n & "</a>"
		Else
			Response.Write " <a href=""" & DEF_BBS_HomeUrl & "a/topic-" & BoardID & "-" & ID & "-" & n & ".html"">" & n & "</a>"
		End If
	next
	if max > 5 then Response.Write "..."
	if max > 4 then
		If RewriteFlag = 0 Then
			Response.Write " <a href=""" & DEF_BBS_HomeUrl & "a/a.asp?B=" & BoardID & "&ID=" & ID & "Aq=" & max & """>" & max & "</a>"
		Else
			Response.Write " <a href=""" & DEF_BBS_HomeUrl & "a/topic-" & BoardID & "-" & ID & "-" & max & ".html"">" & max & "</a>"
		End If
	End If
	Response.Write "</span> ]"
		
End Sub

end class
%>