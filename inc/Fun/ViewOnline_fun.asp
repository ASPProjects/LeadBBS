<%
Function DisplayUserOnline(BoardID,url)

	Dim Lines
	Lines = 5

	Dim Time1
	Time1 = Timer
	Const LMT_Max_OnlineUserList = 197
	Dim ActiveUsers
	ActiveUsers = 0
	If BoardID > 0 and isArray(Application(DEF_MasterCookies & "BoardInfo" & BoardID)) = True Then ActiveUsers = cCur(Application(DEF_MasterCookies & "BDOL" & BoardID))

	Response.Write "<div class=""b_list_box""><table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""95%"">"
	Dim Temp
	Temp = 0
	If ActiveUsers < Num + 1 Then ActiveUsers = Num + 1

	Dim Rs,SQL,m
	m = Left("0" & Request.Form("n"),20)
	If isNumeric(n) = False Then m = 0
	m = cCur(m)
	
	Dim usr
	usr = Left(Request.Form("usr") & "",20)
	If BoardID = 0 Then
		If usr <> "1" Then
			usr = "0"
			if m > 0 Then rs = " where id>=" & m
			SQL = sql_select("Select ID,UserName,HiddenFlag from LeadBBS_onlineUser" & rs & " Order by ID ASC",LMT_Max_OnlineUserList + 1)
		Else
			SQL = sql_select("Select ID,UserName,HiddenFlag from LeadBBS_onlineUser where UserID > 0",LMT_Max_OnlineUserList + 1)
		End If
	Else
		if m > 0 Then rs = " and id>=" & m
		SQL = sql_select("Select ID,UserName,HiddenFlag from LeadBBS_onlineUser where AtBoardID=" & boardID &" and UserID>0" & rs & " Order by ID ASC",LMT_Max_OnlineUserList + 1)
	End If

	Set Rs = LDExeCute(SQL,0)
	Dim Num
	Dim GetData
	If Not rs.Eof Then
		GetData = Rs.GetRows(-1)
		Num = Ubound(GetData,2)
	Else
		Num = -1
	End If
	Rs.close
	Set Rs = Nothing
	
	dim n,rguser,vuser,nxtflag
	rguser = 0
	vuser = 0
	nxtflag = 0
	n = 1
	If Num >= 0 Then
		Dim i
		
		Dim lineNum,ModNum,LineIndex
		lineNum = Fix(Num / Lines)
		ModNum = (Num mod Lines)
		LineIndex = 1
		Response.Write "<tr>"
		For i = 0 To Num
			If n = 1 Then Response.Write "<td width=""" & fix(100/Lines) & "%"" valign=""top""><ul class=""list_break"">"
			Response.Write "<li>"
			if isNull(GetData(1,i)) or GetData(1,i) = "" Then
				Response.Write "<span class=""ol_7""><a href=""" & Url & "User/LookUserInfo.asp?OlID=" & GetData(0,i) & "&amp;Evol=more"" target=""_blank"">�ο�</a></span>"
				vuser = vuser + 1
			Else
				rguser = rguser + 1
				If GetData(1,i) = "�����û�" Then
					Response.Write "<span class=""ol_6""><a href=""" & Url & "User/LookUserInfo.asp?OlID=" & GetData(0,i) & "&amp;Evol=more"" target=""_blank"">�����û�</a></span>"
				Else
					GetData(2,i) = cCur(GetData(2,i))
					Response.Write "<span class=""ol_"
					If GetData(2,i) = 0 Then
						Response.Write "5"
					ElseIf GetBinarybit(GetData(2,i),10) = 1 Then
						Response.Write "1"
					ElseIf GetBinarybit(GetData(2,i),14) = 1 Then
						Response.Write "2"
					ElseIf GetBinarybit(GetData(2,i),8) = 1 Then
						Response.Write "3"
					ElseIf GetBinarybit(GetData(2,i),2) Then
						Response.Write "4"
					Else
						Response.Write "5"
					End If
					Response.Write """><a href=""" & Url & "User/LookUserInfo.asp?name=" & urlEncode(GetData(1,i)) & """ target=""_blank"">" & GetData(1,i) & "</a></span>"
				End If
			End If
			Response.Write "</li>"
			n = n + 1
			If (n > lineNum and ModNum < 0) or (N > (lineNum + 1) and ModNum >= 0) Then
				Response.Write "</ul></td>"
				n = 1
				ModNum = ModNum - 1
			End If
		Next
		Response.Write "</tr>"
		If Num + 1 >= LMT_Max_OnlineUserList Then
			nxtflag = cCur(GetData(0,Num))
		End If
	End If

	If ActiveUsers < Num Then ActiveUsers = Num
	Temp = Num
	If isNumeric(Application(DEF_MasterCookies & "ActiveUsers")) = 0 Then
		Application.Lock
		Application(DEF_MasterCookies & "ActiveUsers") = 0
		Application.UnLock
	End If
	If Application(DEF_MasterCookies & "ActiveUsers") < Temp Then
		Application.Lock
		Application(DEF_MasterCookies & "ActiveUsers") = Temp
		Application.UnLock
	End If
	Dim online
	online = cCur(Application(DEF_MasterCookies & "ActiveUsers"))
	If online < (rguser+vuser) Then online = rguser + vuser
	%>
	<tr><td colspan="<%=Lines%>">
	<br />
	<div class="j_page">
	<%If nxtflag  > 0 Then
		%><a href="#bol" onclick="getAJAX('B<%If BoardID = 0 Then Response.Write "oards"%>.asp','ol=1&amp;n=<%=nxtflag%>&amp;b=<%=BoardID%>','follow0');">��һҳ</a><%
	End If
	If m > 0 Then
		%><a href="#bol" onclick="getAJAX('B<%If BoardID = 0 Then Response.Write "oards"%>.asp','ol=1&amp;b=<%=BoardID%>','follow0');">����</a>
	<%
	End If%>
	<%
	If vuser = 0 Then vuser = ActiveUsers - rguser
	If BoardID = 0 Then
		If usr <> "1" Then
			%><a href="#bol" onclick="getAJAX('Boards.asp','ol=1&amp;usr=1','follow0');">ֻ�������û�</a><%
		Else%><a href="#bol" onclick="getAJAX('Boards.asp','ol=1','follow0');">�鿴ȫ������</a>
		<%
		End If
	End If%>
	<a href="#bol" onclick="ShowOnline('follow0','swap_ol');">�ر�</a>
	</div>
	<%
	If BoardID > 0 Then
		Response.Write "��ǰ���湲 " & ActiveUsers & " �����ߣ�"
		Response.Write "��ǰ�б���ע���û��� " & rguser & " �ˣ��ο� " & vuser & " �� "
	Else
		Response.Write "��ǰ�б���ע���û��� " & rguser & " �� "
	End If
	%>
	</td></tr>
	<tr><td colspan="<%=Lines%>">
	<%If BoardID = 0 Then%>
	<br />ͼ�� <span class="ol_1"><%=DEF_PointsName(6)%></span>
	<span class="ol_2"><%=DEF_PointsName(7)%></span>
	<span class="ol_3"><%=DEF_PointsName(8)%></span>
	<span class="ol_4"><%=DEF_PointsName(5)%></span>
	<span class="ol_5">ע���û�</span>
	<span class="ol_6">�����û�</span>
	<span class="ol_7">�ο�</span><%
	End If%>
	</td></tr></table></div>
	<%

End Function

%>