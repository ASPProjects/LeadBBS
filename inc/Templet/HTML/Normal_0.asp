<%
Class DisplayBoard_HTML_Class

Private simflag,simn,num,CurrentAssosrt,Flag,AssortMasterNum,BoardMasterNum,Rewritestr,LMT_EnableRewrite

Private Sub Class_Initialize

	simflag=0
	simn=2
	
	num = 0
	CurrentAssosrt = 0
	Flag = 0
	
	AssortMasterNum = 3
	BoardMasterNum = 1
	
	LMT_EnableRewrite = GetBinarybit(DEF_Sideparameter,16)

End Sub

Public Sub DisplayBoard_HTML_Fun_Simple(BoardID,BoardAssort,BoardName,BoardIntro,LastWriter,LastWriteTime,TopicNum,AnnounceNum,ForumPass,LastAnnounceID,LastTopicName,MasterList,BoardLimit,AssortName,TodayAnnounce,GoodNum,BoardImgUrl,BoardImgWidth,BoardImgHeight,onlineUser,LowerBoard,AssortMaster)

	Dim n
	If CurrentAssosrt <> BoardAssort Then
		If num > 0 and simflag = 1 Then
			For n = num to simn - 1
				Response.Write "<td width=""" & Fix(100/simn) & "%"" class=""b_list tdbox"">&nbsp;<br />&nbsp;<br />&nbsp;</td>" & VbCrLf
			Next
			Response.Write "</tr>"
		End If
		CurrentAssosrt = BoardAssort
		If Flag = 1 Then
			Response.Write "</table></div></div>"
		End If
		%>
		<div class="contentbox contentbox_boards">
		<%
		Flag = 1
		num = 0
		%>
			<table width="100%" cellspacing="0" cellpadding="0" class="tablebox_sim">
			<tr class="tbhead">
			<td colspan="<%=simn%>">
			<div class="b_assort">
				<div class="b_assort_title">
				<span class="clicktext" title="�ر�/չ��" onclick="LD.blist.assort_disable('<%=BoardAssort%>');"><a href=javascript:; id="b_assort_img_<%=BoardAssort%>" class="b_assort_close<%If inStr(Boards_dis_assortStr,"," & BoardAssort & ",") Then Response.Write "_swap"%>" alt="�ر�/չ��"></a></span>
				
		<%
		If DEF_BBS_HomeUrl = "" or DEF_BBS_HomeUrl = "./" Then
			Response.Write "<a href=""" & DEF_BBS_HomeUrl & "Boards.asp?Assort=" & BoardAssort & """ id=""b_assort_link_" & BoardAssort & """><b>" & AssortName & "</b></a>"
		Else
			If LMT_EnableRewrite = 0 Then
				RewriteStr = "b.asp?b=" & BoardAssort
			Else
				RewriteStr = "forum-" & BoardAssort & "-1.html"
			End If
			Response.Write "<a href=""" & DEF_BBS_HomeUrl & "b/" & RewriteStr & """><b>" & AssortName & "</b></a> [�Ӱ��] "
			BoardAssort = "b" & BoardAssort
		End If
		%></div>
				<span class="clicktext" title="��Լ/����" onclick="LD.blist.assort_click('<%=BoardAssort%>',1);"><img src="<%=DEF_BBS_HomeUrl%>images/blank.gif" class="b_assort_mini_swap" alt="��Լ/����" /></span>
		<%
		If AssortMaster <> "" and DEF_BBS_HomeUrl = "" Then
			Response.Write "<div class=""b_assort_master"">"
			CALL DisplayBoard_HTML_MastList(AssortMaster,AssortMasterNum,DEF_PointsName(7))
			Response.Write "</div>"
		End If%>
			</div>
			</td></tr>
			</table>
		<div id="b_assort_<%=BoardAssort%>"<%If inStr(Boards_dis_assortStr,"," & BoardAssort & ",") Then Response.Write " style=""display:none"""%>>
			<table width="100%" cellspacing="0" cellpadding="0" class="tablebox_sim">
		<%
	End If
	simflag = 1
	If (num mod simn) = 0 Then
		Response.Write "<tr>"
	End If%>
	<td width="<%=Fix(100/simn)%>%" class="b_list tdbox" onmouseover="this.className=$replace(this.className,'b_list','b_list_active');" onmouseout="this.className=$replace(this.className,'b_list_active','b_list')">
	<%
	TodayAnnounce = cCur(TodayAnnounce)
	If TodayAnnounce > 0  Then
		Response.Write "<div class=""b_new fire"">"
	Else
		Response.Write "<div class=""b_none fire"">"
	End If
			If LMT_EnableRewrite = 0 Then
				RewriteStr = "b.asp?b=" & BoardID
			Else
				RewriteStr = "forum-" & BoardID & "-1.html"
			End If
		%>
	
			<div class="oneline">
				<a href="<%=DEF_BBS_HomeUrl%>b/<%=RewriteStr%>"><b><%=BoardName%></b></a> <%
				If TodayAnnounce>0 Then
					Response.Write " ( ���� <b class=""redfont"">" & TodayAnnounce & "</b> ) "
				End If%>
			</div>
			<%
			
			
	If LastTopicName = "" or LastTopicName = null Then
		Response.Write "<span class=""name"">������</span>"
	Else
		'If StrLength(LastTopicName) > 23 Then LastTopicName = LeftTrue(LastTopicName,23-3) & "..."
		LastTopicName = HtmlEncode(LastTopicName)
		Response.Write "<div class=""oneline"">"
		If LastTopicName = "������Ϊ����" and LastWriter = "" Then
			Response.Write "<span class=""name"">����</span><a>������Ϊ����</a>"
		Else
			If cCur(LastAnnounceID) = 0 Then
				Response.Write "<span class=""name"">����</span><a>" & HtmlEncode(LastTopicName) & "</a>"
			Else
				Response.Write "<span class=""name"">����</span><a href=""" & DEF_BBS_HomeUrl & "a/a.asp?B=" & BoardID & "&amp;ID=" & LastAnnounceID & "&amp;Aupflag=1&amp;Anum=1""><span class=""word-break-all"">" & htmlencode(LastTopicName) & "</span></a>"
			End If
		End If
		Response.Write "</div>"
	End If
			%>
			<div class="oneline"><span class="name">���⣺</span><a><%=TopicNum%>, ����<%=AnnounceNum%></a></div>
		</div>
	</td>
	<%
	num = num + 1
	If num > simn - 1 Then
		num = 0
		Response.Write "</tr>"
	End If

End Sub

Function DisplayBoard_HTML_GetBoardType(ForumPass,BoardLimit)

	Dim Tmp
	If ForumPass <> "" Then
		Tmp = "������̳"
	Else
		If GetBinarybit(BoardLimit,7) = 1 Then
			Tmp = DEF_PointsName(8) & "ר��"
		Else
			If GetBinarybit(BoardLimit,4) = 1 and GetBinarybit(BoardLimit,3) = 1 and GetBinarybit(BoardLimit,6) = 1 Then
				Tmp = "ֻ����̳"
			Else
				If GetBinarybit(BoardLimit,2) = 1 Then
					Tmp = "��ʽ�û���̳"
				Else
					If GetBinarybit(BoardLimit,1) = 1 Then
						Tmp = "ע����֤��̳"
					Else
						If GetBinarybit(BoardLimit,9) = 1 Then
							Tmp = "������̳"
						Else
							Tmp = "������̳"
						End If
					End If
				End If
			End If
		End If
	End If
	DisplayBoard_HTML_GetBoardType = Tmp

End Function

Public Sub DisplayBoard_HTML_Fun(BoardID,BoardAssort,BoardName,BoardIntro,LastWriter,LastWriteTime,TopicNum,AnnounceNum,ForumPass,LastAnnounceID,LastTopicName,MasterList,BoardLimit,AssortName,TodayAnnounce,GoodNum,BoardImgUrl,BoardImgWidth,BoardImgHeight,onlineUser,LowerBoard,AssortMaster)

	Dim n,BoardType
	Temp = 0
	If  CurrentAssosrt <> BoardAssort Then
		If num > 0 and simflag = 1 Then
			For n = num to simn -1
				Response.Write "<td width=""" & Fix(100/simn) & "%"" class=""b_list tdbox"">&nbsp;<br />&nbsp;<br />&nbsp;</td></tr>" & VbCrLf
			Next
		End If
		CurrentAssosrt = BoardAssort
		If Flag = 1 Then
			Response.Write "</table></div></div>"
		End If
		%>
		<div class="contentbox contentbox_boards">
		<%
		Flag = 1
		%>
		<table width="100%" cellspacing="0" cellpadding="0" class="tablebox">
		<tr class="tbhead">
			<td>
				<div class="b_assort">
				<div class="b_assort_title">
				<span class="clicktext" title="�ر�/չ��" onclick="LD.blist.assort_disable('<%=BoardAssort%>');"><a href=javascript:; id="b_assort_img_<%=BoardAssort%>" class="b_assort_close<%If inStr(Boards_dis_assortStr,"," & BoardAssort & ",") Then Response.Write "_swap"%>" alt="�ر�/չ��"></a></span>
		<%
		If DEF_BBS_HomeUrl = "" or DEF_BBS_HomeUrl = "./" Then
			Response.Write "<a href=""" & DEF_BBS_HomeUrl & "Boards.asp?Assort=" & BoardAssort & """ id=""b_assort_link_" & BoardAssort & """><b>" & AssortName & "</b></a>"
		Else
			If LMT_EnableRewrite = 0 Then
				RewriteStr = "b.asp?b=" & BoardAssort
			Else
				RewriteStr = "forum-" & BoardAssort & "-1.html"
			End If
			Response.Write "<a href=""" & DEF_BBS_HomeUrl & "b/" & RewriteStr & """><b>" & AssortName & "</b></a> [�Ӱ��] "
			BoardAssort = "b" & BoardAssort
		End If%></div>
				
				<span class="clicktext" title="��Լ/����" onclick="LD.blist.assort_click('<%=BoardAssort%>',0);"><img src="<%=DEF_BBS_HomeUrl%>images/blank.gif" class="b_assort_mini" alt="��Լ/����" /></span>
		<%
		If AssortMaster <> "" and DEF_BBS_HomeUrl = "" Then
			Response.Write "<div class=""b_assort_master"">"
			CALL DisplayBoard_HTML_MastList(AssortMaster,AssortMasterNum,DEF_PointsName(7))
			Response.Write "</div>"
		End If%>
			</div>
			</td></tr>
			</table>
		<div id="b_assort_<%=BoardAssort%>"<%If inStr(Boards_dis_assortStr,"," & BoardAssort & ",") Then Response.Write " style=""display:none"""%>>
			<table width="100%" cellspacing="0" cellpadding="0" class="tablebox">
		<%
	End If
	simflag = 0%>
	<tr class="b_list" onmouseover="this.className='b_list_active';" onmouseout="this.className='b_list';">
		<td valign="top" class="tdbox">
	<%
	TodayAnnounce = cCur(TodayAnnounce)
	If TodayAnnounce > 0  Then
		Response.Write "<div class=""b_new fire"">"
	Else
		Response.Write "<div class=""b_none fire"">"
	End If

	If LMT_EnableRewrite = 0 Then
		RewriteStr = "b.asp?b=" & BoardID
	Else
		RewriteStr = "forum-" & BoardID & "-1.html"
	End If
	If BoardImgUrl & "" <> "" Then
		Response.Write "<a href=""" & DEF_BBS_HomeUrl & "b/" & RewriteStr & """><img src=""" & BoardImgUrl & """ width=""" & BoardImgWidth & """ height=""" & BoardImgHeight & """ class=""b_list_img"" alt="""" /></a>"
	End If
	%>
		<a href="<%=DEF_BBS_HomeUrl%>b/<%=RewriteStr%>"><b><%=BoardName%></b></a>
		
	<a href="<%=DEF_BBS_HomeUrl%>a/a2.asp?B=<%=BoardID%>">
		<img src="<%=DEF_BBS_HomeUrl%>images/<%=GBL_DefineImage%>BoardTopic/post.gif" title="��������" alt="��������" /></a>
	<a href="<%=DEF_BBS_HomeUrl%>b/b.asp?B=<%=BoardID%>&amp;E=0">
		<img src="<%=DEF_BBS_HomeUrl%>images/<%=GBL_DefineImage%>BoardTopic/elist.gif" title="�鿴����" alt="�鿴����" /></a>
		<br />
		<span><%=BoardIntro%></span>
		<br />
			<div class="b_board_master">
			<%CALL DisplayBoard_HTML_MastList(MasterList,BoardMasterNum,DEF_PointsName(8))%>
			</div>
	</div>
	</td>
	<td width="100" valign="top" class="tdbox">
		<span class="name">����</span> <%=TopicNum%><br />
		<span class="name">����</span> <b<%If TodayAnnounce>0 Then Response.Write " class=""redfont"""%>><%=TodayAnnounce%></b><br />
		<span class="name">����</span> <%=AnnounceNum%>
	</td>
	<td align="left" width="222" valign="top" class="tdbox">
	<%
	Dim Temp
	If LastTopicName = "" or LastTopicName = null Then
		Response.Write "<span class=""name"">������</span>"
	Else
		'If StrLength(LastTopicName) > 31 Then LastTopicName = LeftTrue(LastTopicName,31-3) & "..."
		LastTopicName = HtmlEncode(LastTopicName)
		Response.Write "<div class=""oneline"">"
		If LastTopicName = "������Ϊ����" and LastWriter = "" Then
			Response.Write "<span class=""name"">����</span>������Ϊ����"
			Temp = 1
		Else
			If cCur(LastAnnounceID) = 0 Then
				Response.Write "<span class=""name"">����</span><a>" & HtmlEncode(LastTopicName) & "</a>"
			Else
				Response.Write "<span class=""name"">����</span><a href=""" & DEF_BBS_HomeUrl & "a/a.asp?B=" & BoardID & "&amp;ID=" & LastAnnounceID & "&amp;Aupflag=1&amp;Anum=1""><span class=""word-break-all"">" & htmlencode(LastTopicName) & "</span></a>"
			End If
		End If
		Response.Write "</div>"
	End If
	Response.Write "<div class=""oneline"">"
	If LastWriter = "" and Temp <> 1 Then
		Response.Write "<span class=""name"">���ߣ�</span>"
		LastWriter = "��"
		Response.Write LastWriter
		Response.Write "</div>"
	Else
		Response.Write "<span class=""name"">���ߣ�</span>"
		If LastWriter <> "�ο�" Then
			If Temp = 1 Then
				Response.Write "����"
			Else
				Response.Write "<a href=""" & DEF_BBS_HomeUrl & "User/LookUserInfo.asp?Name=" & urlEncode(LastWriter) & """>" & HtmlEncode(LastWriter) & "</a>"
			End If
		Else
			Response.Write HtmlEncode(LastWriter)
		End If
		Response.Write "</div>"
		If Len(LastWriteTime) = 14 Then
			LastWriteTime = RestoreTime(LastWriteTime)
			LastWriteTime = ConvertTimeString(LastWriteTime)
		Else
			LastWriteTime = "��"
		End If%>
			<div class="oneline"><span class="name">ʱ�䣺</span><a><%=LastWriteTime%></a></div>
		<%
	End If%>
	</td>
	</tr>
	<%

End Sub

Public Sub DisplayBoard_HTML_Fill

	Dim n
	If num > 0 and simflag = 1 Then
		For n = num to simn - 1
			Response.Write "<td width=""" & Fix(100/simn) & "%"" class=""b_list tdbox"">&nbsp;<br />&nbsp;<br />&nbsp;</td>" & VbCrLf
		Next
		Response.Write "</tr>"
	End If

End Sub

Private Sub DisplayBoard_HTML_MastList(s,num,flag)

	If "?LeadBBS?" = s Then
		Response.Write "ȫ��" & DEF_PointsName(8)
	Else
		If s = "" or s = null Then
			Response.Write flag & "����"
			Exit Sub
		End If
		Dim ss,n,m
		ss = Split(s,",")
		m = Ubound(ss,1)
		If m >= num Then
			%><%=flag%>��<%
		Else%>
			<%=flag%>��<%
		End If
		For n = 0 to m
			If n >= num Then Exit For
			If n > 0 Then Response.Write ", "
			Response.Write "<a href=""" & DEF_BBS_HomeUrl & "User/LookUserinfo.asp?name=" & urlEncode(ss(n)) & """"
			Response.Write ">" & HtmlEncode(ss(n)) & "</a>"
		Next
		If n >= num and n <= m Then
			%> <div class="layer_item" style="display:inline"><span class="layer_item_title name"><em>...</em></span>
				<div class="layer_iteminfo">
				<ul class="menu_list">
					<%
			Response.Write "<li><b>����" & flag & "</b></li>"
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

End Class
%>