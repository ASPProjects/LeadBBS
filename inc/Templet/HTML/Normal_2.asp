<%
		i = i + 1
		Dim tmp1,tmp2,Tmp3,MasterF
		MasterF = 0
		tmp1 = ""
		Dim OnlineFlag
		GetData(15,n) = cCur(GetData(15,n))
		If GetData(15,n) > 0 Then
			If GetBinarybit(GetData(37,n),10) = 1 Then
				tmp1 = "<span class=""name"">ְ��</span><span class=""bluefont value"">" & DEF_PointsName(6) & "</span>"
				MasterF = 1
			ElseIf GetBinarybit(GetData(37,n),14) = 1 Then
				tmp1 = "<span class=""name"">ְ��</span><span class=""bluefont value"">" & DEF_PointsName(7) & "</span>"
				MasterF = 1
			ElseIf GetBinarybit(GetData(37,n),8) = 1 Then
				tmp1 = "<span class=""name"">ְ��</span><span class=""bluefont value"">" & DEF_PointsName(8) & "</span>"
				MasterF = 1
			ElseIf GetBinarybit(GetData(37,n),2) = 1 Then
				tmp1 = "<span class=""name"">��Ա</span><span  class=""greenfont value"">" & DEF_PointsName(5) & "</span>"
			End If
		End If%>
	<div id="anc_table_div_<%=GetData(0,n)%>" class="anc_table_div">
	<table width="100%" border="0" cellspacing="0" cellpadding="0" class="tablebox anc_table<%
		if MasterF = 1 then
			response.write " post_master"
		end if%>">
	<tr>
		<td width="<%=DEF_BBS_LeftTDWidth%>" class="tdleft a_userinfotd">
			<div class="a_author"><%
		
		If GetData(14,n) = "[LeadBBS]" Then GetData(14,n) = "ϵͳ"
		If GetData(15,n) > 0 Then
			Response.Write "<a href=""" & DEF_BBS_HomeUrl & "User/LookUserInfo.asp?ID=" & GetData(15,n) & """ class=""uname""><b>" & htmlencode(GetData(14,n)) & "</b></a>"
		Else
			Response.Write htmlencode(GetData(14,n))
		End If
		%>
		
			</div>
		<%
		

		If (ccur(GetData(42,n)) = 1) and DEF_EnableUserHidden = 1 Then
			OnlineFlag = 0
		Else
			OnlineFlag = DateDiff("s",RestoreTime(GetData(32,n)), DEF_Now)
			If OnlineFlag < 0 or OnlineFlag > DEF_UserOnlineTimeOut Then
				OnlineFlag = 0
			Else
				OnlineFlag = 1
			End If
		End If


		If GetData(14,n) = "�ο�" and GetData(15,n) = 0 Then OnlineFlag = 1
		Response.Write "<img src=""" & DEF_BBS_HomeUrl & "images/" & GBL_DefineImage & "sxmg/"
		Select Case GetData(31,n)
		Case "Ů"
			If OnlineFlag = 1 Then
				Response.Write "FeMale.gif"" title=""��Ů����"
			Else
				Response.Write "OfFeMale.gif"" title=""��Ů����"
			End If
		Case "��"
			If OnlineFlag = 1 Then
				Response.Write "Male.gif"" title=""˧�磬��"
			Else
				Response.Write "OfMale.gif"" title=""˧�磬��"
			End If
		Case Else
			If OnlineFlag = 1 Then
				Response.Write "Male.gif"" title=""��"
			Else
				Response.Write "OfMale.gif"" title=""��"
			End If
		End Select
		Response.Write "��"" class=""a_seximg"" alt=""�������"" />"
		
		%>
		</td>
		<td class="tdright a_ancinfotd">
			<div class="a_ancinfobox fire">
			<div class="a_ancinfo">
				<ul>
				<%If cCur(GetData(9,n)) > 0 Then Response.Write "<li><img src=""../images/" & GBL_DefineImage & "bf/face" & GetData(9,n) & ".gif"" align=""absmiddle"" alt=""����"" /></li>"%>
				<li><em><%=ConvertTimeString(RestoreTime(GetData(10,n)))%></em></li>
				<%If CheckSupervisorUserName = 1 Then%><li><em><%=GetData(19,n)%></em></li><%End If%>
				<%If cCur(GetData(1,n)) = 0 Then%><li>�Ķ���<%=GetData(12,n)%>��</li><%End If%>
				<%


		If A_NotReplay <> 1 Then Response.Write "<li><a href=""a2.asp?b=" & GBL_board_ID & "&amp;ID=" & GetData(0,n) & "&amp;submitflag=first&amp;repost=1"" title=""���ûظ�����"">����</a></li>"

		
		If GBL_CHK_User <> "" Then
			If GetBinarybit(GBL_CHK_UserLimit,6) = 0 Then
			%><li class="layerico"><a href="Processor.asp?action=MakeGood&amp;b=<%=GBL_Board_ID%>&amp;ID=<%=GetData(0,n)%>" onclick="return(a_command('��������',this,'MakeGood&b=<%=GBL_Board_ID%>&ID=<%=GetData(0,n)%>'));" title="�������ֻ򾫻�">����</a></li><%
			End If
		End If

		If GBL_CHK_User <> "" and (GBL_BoardMasterFlag >= 5 or GetData(14,n) = GBL_CHK_User) Then
			Response.Write "<li><a href=""Editannounce.asp?b=" & GBL_board_ID & "&amp;ID=" & GetData(0,n) & """ title=""�༭��������"">�༭</a></li>"
			%><li><a href="Processor.asp?action=TypeSet&b=<%=GBL_Board_ID%>&ID=<%=GetData(0,n)%>" onclick="return(a_command('���ӣ��ۺϹ���',this,'TypeSet&b=<%=GBL_Board_ID%>&ID=<%=GetData(0,n)%>'));" title="����ĵ�����������">����</a></li><%
		End If

		If GBL_CHK_User <> "" and GBL_BoardMasterFlag >= 5 Then
			If GetBinarybit(GBL_Board_BoardLimit,5) = 0 and GetBinarybit(GBL_CHK_UserLimit,5) = 0 Then
				%><li class="layerico"><%
				If cCur(GetData(1,n)) > 0 Then%><input class="fmchkbox" type="checkbox" name="ids" id="ids<%=Index%>" value="<%=GetData(0,n)%>" onclick="delbody_view(this);" /><%
					Tmp3 = "Del&b=" & GBL_Board_ID & "&ID=" & GetData(0,n)
				Else%><%
					If GBL_Board_ID <> 444 and DEF_EnableDelAnnounce = 0 Then
						Tmp3 = "Move&b=" & GBL_Board_ID & "&ID=" & GetData(0,n) & "&BoardID2=444"
					Else
						Tmp3 = "Del&b=" & GBL_Board_ID & "&ID=" & GetData(0,n)
					End If
				End If
				Index = Index + 1
				%><a href="Processor.asp?action=<%=Tmp3%>" onclick="return(a_command('ɾ������',this,'<%=Tmp3%>'));" title="ɾ������">ɾ��</a></li><%
			End If
		End If
		
		%></ul>
		</div>
		<%

		Response.Write "<div class=""a_floor""><a name=F" & GetData(0,n) & "></a><span class=layerico><span class=""clicktext"" oncontextmenu=""copyClipboard('text',$id('Content" & GetData(0,n) & "').innerHTML,'�������ݳɹ�������������!','" & DEF_BBS_HomeUrl & "',this);return(false);"" onclick=""var clipdata='';if(event.ctrlKey){clipdata=$id('Content" & GetData(0,n) & "').innerHTML;}else{clipdata=$id('Content" & GetData(0,n) & "').innerText;};copyClipboard('text',clipdata,'�������ݳɹ�������������!','" & DEF_BBS_HomeUrl & "',this);"" title=""���������������(�һ���ctrl+�������Դ��)"">"
		If cCur(GetData(1,n)) = 0 Then
			Response.Write "<b>¥��</b>"
		Else
			Response.Write "��<b>" & page*DEF_TopicContentMaxListNum+i & "</b>¥"
		End If
		Response.Write "</span></span></div>"
		%>
		</div>
		</td>
	</tr>
	<tr>
		<td width="<%=DEF_BBS_LeftTDWidth%>" valign="top" class="tdleft" rowspan="2" onmouseover="swap_ancinfo(this,1)" onmouseout="swap_ancinfo(this,0)">
			<div class="a_userinfo">
			<ul class="info_one"><%
		Response.Write "<li><img class=""a_faceimg"" src="
		If GetData(35,n) > DEF_AllFaceMaxWidth Then GetData(35,n) = DEF_AllFaceMaxWidth
		If GetData(36,n) > DEF_AllFaceMaxWidth*2 Then GetData(36,n) = DEF_AllFaceMaxWidth
		If DEF_AllDefineFace <> 0 and GetData(34,n) <> "" Then
			Response.Write chr(34) & htmlencode(GetData(34,n)) & chr(34) & " width=""" & GetData(35,n) & """ height=""" & GetData(36,n) & """"
		Else
			Response.Write """" & DEF_BBS_HomeUrl & "images/face/" & string(4-len(cstr(GetData(22,n))),"0")&GetData(22,n) & ".gif"""
		End If
		Response.Write " alt=""ͷ��"" /></li>"
		If (Not isNull(GetData(44,n))) and GetData(44,n) <> "" Then Response.write "<li><span class=""grayfont"">" & htmlencode(GetData(44,n)) & "</span></li>"

		'If len(GetData(29,n))=14 Then
			'GetData(29,n) = RestoreTime(Left(GetData(29,n),8))
			'If isTrueDate(GetData(29,n)) Then
			'	Response.Write " " & Constellation(GetData(29,n))
			'End If
			'If Len(GetData(41,n)) = 14 Then
			'	GetData(41,n) = ReStoreTime(GetData(41,n))
			'	If isTrueDate(GetData(41,n)) Then Response.Write " " & DisplayBirthAnimal(year(GetData(41,n)))
			'End If
		'End If

		Response.Write "<li class=""level""><img src=""" & DEF_BBS_HomeUrl & "images/" & GBL_DefineImage & "lvstar/level" & GetData(23,N) & ".gif"" class=""a_levelimg"" title=""���� " & DEF_UserLevelString(GetData(23,N)) & """ alt=""����"" /></li>"

		Response.Write "</ul>"
		
		
		Response.Write "<ul class=""info_3""><li><span>"
		If GetData(15,n) > 0 Then Response.Write "<a href=""" & DEF_BBS_HomeUrl & "User/SendMessage.asp?SdM_ToUser=" & urlencode(GetData(14,n)) & """ onclick=""return(sendprivatemsg(this,'" & DEF_BBS_HomeUrl & "'));""><img src=""../images/" & GBL_DefineImage & "message.GIF"" alt=""��" & htmlencode(GetData(14,n)) & "����Ϣ"" class=""absmiddle"" /></a>"
		If trim(GetData(24,n))<>"" Then
			If Left(lcase(GetData(24,n)),4)<>"http" Then GetData(24,n) = "http://" & GetData(24,n)
			Response.Write " <a href=""" & htmlencode(GetData(24,n)) & """ target=""_blank""><img src=""../images/" & GBL_DefineImage & "home.gif"" alt=""���û���ҳ"" class=""absmiddle"" /></a>"
		End if
		
		'If GetData(20,n)<>"" and (ccur(GetData(43,n)) = 1) Then
		'	Response.Write "<a href=""mailto:" & GetData(20,n) & """ target=""_blank""><img src=""../images/" & GBL_DefineImage & "mail.gif"" alt=""�����û����ʼ�"" class=""absmiddle"" /></a>"
		'End if
		
		'If isNull(GetData(21,n)) or GetData(21,n)="" Then GetData(21,n)=0
		'If cCur(GetData(21,n))>=10000 Then Response.Write "<a href=""http://search.tencent.com/cgi-bin/friend/user_show_info?ln=" & GetData(21,n) & """ target=""_blank""><img src=""../images/" & GBL_DefineImage & "oicq.gif"" title=""�鿴QQ����"" class=""absmiddle"" /></a>"
		If GBL_CHK_User <> "" and GetData(15,n) > 0 Then
			%><a href="Processor.asp?action=AddFriend&FriendName=<%=UrlEncode(GetData(14,n))%>&b=<%=GBL_Board_ID%>&ID=<%=GetData(0,n)%>" onclick="return(a_msg(this,'<%Response.Write "AddFriend&FriendName=" & UrlEncode(GetData(14,n))%>&SureFlag=1'));"><img src="../images/<%=GBL_DefineImage%>friend.gif" alt="��Ϊ����" class="absmiddle" /></a><%
		End If
		Response.Write "</span></li></ul>"
		
		Response.Write "<ul class=""info_two"">"
		If GetData(15,n) > 0 Then
			If GetData(27,n)<>"0" and GetData(27,n) <> "" Then
				Response.write "<li><span class=""name"">" & DEF_PointsName(9) & "</span><span class=""value"">" & Replace(DisplayOfficerString(GetData(27,n)),",","</span></li><li><span class=""name"">&nbsp;</span><span class=""value"">") & "</span></li>"
			End If
			If tmp1 <> "" Then Response.Write "<li>" & tmp1 & "</li>"
			'Response.write "<li><span class=""name"">����</span><span class=""value"">" & DEF_UserLevelString(GetData(23,N)) & "</span></li>"

			GetData(47,n) = cCur(GetData(47,n))
			If GetData(47,n) <> 0 Then
				If GetData(47,n) < 0 Then
					GetData(47,n) = GetData(47,n)
				Else
					GetData(47,n) = "<span class=""bluefont value"">��" & GetData(47,n) & "</span>"
				End If
				Response.write "<li><span class=""name"">" & DEF_PointsName(2) & "</span><span class=""value"">" & GetData(47,n) & "</span></li>"
			End If
			If cCur(GetData(48,n)) <> 0 Then Response.Write "<li><span class=""name"">" & DEF_PointsName(1) & "</span><span class=""value"">" & GetData(48,n) & "</span></li>"

			Response.Write "<li><span class=""name"">" & DEF_PointsName(0) & "</span><span class=""value"">" & GetData(26,n) & "</span></li>"
			Response.Write "<li><span class=""name"">" & DEF_PointsName(4) & "</span><span class=""value"">" & CLng(cCur(GetData(28,n))/60) & "</span></li>"
			Response.Write "<li><span class=""name"">����</span><span class=""value"">" & GetData(33,n) & "</span></li>"
			Response.Write "<li><span class=""name"">ע��</span><span class=""value"">" & Mid(RestoreTime(GetData(30,n)),1,10) & "</span></li>"
		End If
		Response.Write "</ul>"
		%>
		</div>
		
		</td>
		<td class="tdright a_topiccontent" valign="top">