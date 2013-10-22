<!-- #include file=../inc/BBSSetup.asp -->
<!-- #include file=../inc/User_Setup.ASP -->
<!-- #include file=../inc/Board_Popfun.asp -->
<!-- #include file=inc/UserTopic.asp -->
<%
DEF_BBS_HomeUrl = "../"
'DEF_MaxListNum = 100
If CheckSupervisorUserName = 0 Then
	GBL_CHK_PWdFlag = 0
	OpenDatabase
Else
	initDatabase
End If


Dim AtBoardID,AtBoardName,AtBoardIDCount
GetAtBoardIDInfo

If AtBoardID = 0 Then
	BBS_SiteHead DEF_SiteNameString & " - �����û�",0,"<span class=navigate_string_step>�����û�</span>"
Else
	BBS_SiteHead DEF_SiteNameString & " - �����û�",0,"<span class=navigate_string_step><a href=userOnline.asp>�����û�</a></span><span class=navigate_string_step>" & AtBoardName & " ����</span>"
End If
UpdateOnlineUserAtInfo GBL_board_ID,"�鿴�����û�"


UserTopicTopInfo("forum")
		
DisplayUserOnline
closeDataBase
UserTopicBottomInfo
SiteBottom


Function GetAtBoardIDInfo

	AtBoardID = Left(Request.QueryString("AtBoardID"),14)
	If isNumeric(AtBoardID) = 0 Then AtBoardID = 0
	AtBoardID = cCur(AtBoardID)
	If AtBoardID < 1 Then
		AtBoardID = 0
		Exit Function
	End If

	If isArray(Application(DEF_MasterCookies & "BoardInfo" & AtBoardID)) = False Then
		AtBoardID = 0
	Else
		AtBoardName = "<a href=../b/b.asp?B=" & AtBoardID & ">" & Application(DEF_MasterCookies & "BoardInfo" & AtBoardID)(0,0) & "</a>"
	End If

End Function

Function GetActiveUserNumber(BoardID)

	Dim Rs,tmp
	If isArray(Application(DEF_MasterCookies & "BoardInfo" & BoardID)) = True Then
		Rs = Application(DEF_MasterCookies & "BDOL" & BoardID)
		If Rs > 0 and Rs <= cCur(Application(DEF_MasterCookies & "ActiveUsers")) Then 
			GetActiveUserNumber = Rs
			Exit Function
		End If
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
	End If

End Function

Function DisplayUserOnline

	GBL_CHK_TempStr=""
	Dim Rs,SQL
	Dim UpDownPageFlag
	UpDownPageFlag = Request.QueryString("uf")

	Dim Start,RecordCount
	RecordCount=0
	Dim SQLendString

	Start = Left(Trim(Request.QueryString("Start")),14)
	Start = start
	If isNumeric(Start)=0 or Start="" Then Start=0
	Start = cCur(Start)

	Dim SQLCountString,whereFlag
	'SQLendString=" where T1.UserID>1"
	'whereFlag = 1
	If AtBoardID > 0 Then
		SQLendString=" where T1.AtBoardID=" & AtBoardID
		whereFlag = 1
	End If

	Rem ����Ĵ���ʹĿǰ�ݲ��ṩ���з���˫�ز�ѯ
	
	SQLCountString = SQLendString
	If UpDownPageFlag = "1" and Start>0 then
		If whereFlag = 1 Then
			SQLendString = SQLendString & " and T1.ID<" & Start
		Else
			SQLendString = SQLendString & " where T1.ID<" & Start
			whereFlag = 1
		End If
	Else
		If whereFlag = 1 Then
			SQLendString = SQLendString & " and T1.ID>" & Start
		Else
			SQLendString = SQLendString & " where T1.ID>" & Start
			whereFlag = 1
		End If
	end If

	If UpDownPageFlag = "1" then
		'If DEF_IDFocusFlag<> 2 Then SQLendString = SQLendString & " Order by  T2.ID DESC"
		SQLendString = SQLendString & " Order by  T1.ID DESC"
	Else
		'If DEF_IDFocusFlag<> 1 Then SQLendString = SQLendString & " Order by  T2.ID ASC"
		SQLendString = SQLendString & " Order by  T1.ID ASC"
	End If

	Dim ReCount
	Dim MaxRecordID,MinRecordID
	MaxRecordID = 0
	MinRecordID = 0
	If AtBoardID > 0 Then
		ReCount = GetActiveUserNumber(AtBoardID)
	Else
		ReCount = application(DEF_MasterCookies & "ActiveUsers")
	End If
	RecordCount = Request.QueryString("rc")
	If isNumeric(RecordCount) = 0 Then RecordCount = 0
	RecordCount = Fix(cCur(RecordCount))
	
	MaxRecordID = Request.QueryString("maxr")
	If isNumeric(MaxRecordID) = 0 Then MaxRecordID = -1
	MaxRecordID = Fix(cCur(MaxRecordID))

	MinRecordID = Request.QueryString("minr")
	If isNumeric(MinRecordID) = 0 Then MinRecordID = -1
	MinRecordID = Fix(cCur(MinRecordID))

	If RecordCount <> ReCount or MaxRecordID < 0 or MinRecordID < 0 Then
		RecordCount = ReCount
		SQL = "select Max(T1.id) from LeadBBS_onlineUser as T1 left join LeadBBS_User as T2 on T1.UserID=T2.ID" & SQLCountString
		Set Rs = LDExeCute(SQL,0)
	
		If not Rs.Eof Then
			If Rs(0) <> "" Then
				MaxRecordID = cCur(Rs(0))
			Else
				MaxRecordID = 0
			End If
		End If
		Rs.Close
		Set Rs = Nothing
		
		SQL = "select Min(T1.id) from LeadBBS_onlineUser as T1 left join LeadBBS_User as T2 on T1.UserID=T2.ID" & SQLCountString
		Set Rs = LDExeCute(SQL,0)
	
		If not Rs.Eof Then
			If Rs(0) <> "" Then
				MinRecordID = cCur(Rs(0))
			else
				MinRecordID = 0
			end If
		End If
		Rs.Close
		Set Rs = Nothing
	End If

	Dim FirstID,LastID

	SQL = sql_select("select T1.ID,T2.UserName,T2.Points,T1.IP,T2.OnlineTime,T2.UserLevel,T2.ID,T1.AtUrl,T1.AtInfo,T2.Userphoto,T2.FaceUrl,T2.FaceWidth,T2.FaceHeight,T2.ShowFlag,T1.Browser,T1.System from LeadBBS_onlineUser as T1 left join LeadBBS_User As T2 on T1.UserID=T2.ID" & SQLendString,DEF_MaxListNum)
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
	
	
	Dim i,N
	If Num>=0 Then
		i=1
	
	
	Dim MinN,MaxN,StepValue
	SQL = ubound(getdata,2)
	If UpDownPageFlag = "1" then
		MinN = SQL
		MaxN = 0
		StepValue = -1
	Else
		MinN = 0
		MaxN = SQL
		StepValue = 1
	End If
	
	LastID = cCur(GetData(0,MaxN))
	FirstID = cCur(GetData(0,MinN))

	Dim QueryStr,PageSplictString
	QueryStr = "?rc=" & RecordCount
	If MaxRecordID >= 0 Then QueryStr = QueryStr & "&maxr=" & MaxRecordID
	If MinRecordID >= 0 Then QueryStr = QueryStr & "&minr=" & MinRecordID
	If AtBoardID > 0 Then QueryStr = QueryStr & "&AtBoardID=" & AtBoardID

	PageSplictString = PageSplictString & "<div class=j_page>"
	if FirstID>MinRecordID and FirstID<>0 then
		PageSplictString = PageSplictString & "<a href=UserOnline.asp" & QueryStr & "&Start=0>��ҳ</a> " & VbCrLf
	else
		'PageSplictString = PageSplictString & "��ҳ" & VbCrLf
	end if

	if FirstID > MinRecordID and FirstID<>0 then
		PageSplictString = PageSplictString & " <a href=UserOnline.asp" & QueryStr & "&Start=" & FirstID & "&uf=1>��ҳ</a> " & VbCrLf
	else
		'PageSplictString = PageSplictString & " ��ҳ" & VbCrLf
	end if

	if LastID<MaxRecordID and LastID<>0 then
		PageSplictString = PageSplictString & " <a href=UserOnline.asp" & QueryStr & "&Start=" & LastID & ">��ҳ</a> " & VbCrLf
	else
		'PageSplictString = PageSplictString & " ��ҳ" & VbCrLf
	end if

	if LastID < MaxRecordID and LastID<>0 then
		PageSplictString = PageSplictString & " <a href=UserOnline.asp" & QueryStr & "&Start=" & MaxRecordID+1 & "&uf=1>βҳ</a> " & VbCrLf
	else
		'PageSplictString = PageSplictString & " βҳ" & VbCrLf
	end if
	PageSplictString = PageSplictString & "<b>��" & RecordCount & "����</b>"
	'If (RecordCount mod DEF_MaxListNum)=0 Then
	'	PageSplictString = PageSplictString & " ��<b>" & clng(RecordCount/DEF_MaxListNum) & "</b>ҳ"
	'Else
	'	If RecordCount>=DEF_MaxListNum Then
	'		If (RecordCount mod DEF_MaxListNum) = 0 Then
	'			PageSplictString = PageSplictString & " ��<b>" & clng(RecordCount/DEF_MaxListNum) & "</b>ҳ"
	'		Else
	'			PageSplictString = PageSplictString & " ��<b>" & clng(RecordCount/DEF_MaxListNum)+1 & "</b>ҳ"
	'		End If
	'	Else
	'		PageSplictString = PageSplictString & " ��<b>1</b>ҳ"
	'	End If
	'End If
	'PageSplictString = PageSplictString & " ÿҳ<b>" & DEF_MaxListNum & "</b>��"
	PageSplictString = PageSplictString & "</div>"

	%>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" class=table_in>
				<tr class=tbinhead>
					<td width=<%=DEF_AllFaceMaxWidth+30%>><div class=value>�û�</div></td>
					<td><div class=value>��Ϣ</div></td>
				</tr>
<%
		for n= MinN to MaxN Step StepValue
			If CheckSupervisorUserName = 1 and GBL_UserID > 0 then GetData(13,N) = 0
			%>
                    <tr><%
                    If GetData(1,n) & "" = "" Then%>
			<td class=tdbox>
				<a href=LookUserInfo.asp?OlID=<%=GetData(0,n)%>&Evol=more>�ο�</a></td>
			<td class=tdbox>
				<ul>
				<%If CheckSupervisorUserName = 1 Then%>
                     		<li>IP: <%=GetData(3,n)%></li><%
                     		End If%>
                     		<li>
                     		�������<%=htmlencode(GetData(14,n))%> ����ϵͳ��<%=htmlencode(GetData(15,n))%>
                     		</li>
                     		<li>
                     		����λ��: <a href="<%=GetData(7,N)%>"><%
				If StrLength(GetData(8,N)) > 130 Then
					Response.Write htmlencode(LeftTrue(GetData(8,N),127)) & "..."
				Else
					Response.Write htmlencode(GetData(8,N))
				End If%></a>
				</li>
				</ul>
			<%
			Else%>
			<td class=tdbox>
				<%
				If (ccur(GetData(13,N)) = 1) and DEF_EnableUserHidden = 1 Then
				Else
					If DEF_AllDefineFace = 0 or GetData(10,N) & "" = "" Then
						If GetData(9,N)<>"" and isNumeric(GetData(9,N)) Then
							%><img src=../images/face/<%=string(4-len(cstr(GetData(9,N))),"0")&GetData(9,N)%>.gif align=middle><%
						End If
					Else%>
							<img src="<%=htmlencode(GetData(10,N))%>" align=middle width=<%=GetData(11,N)%> height=<%=GetData(12,N)%>>
					<%End If
				End If
                     		If (ccur(GetData(13,N)) = 1) and DEF_EnableUserHidden = 1 Then
                     			%><a href=LookUserInfo.asp?OlID=<%=GetData(0,n)%>&Evol=more><div class=user>�����û�</div></a><%
                     		Else
                     			%><a href=LookUserInfo.asp?id=<%=GetData(6,n)%>><div class=user><%=htmlencode(GetData(1,n))%></div></a><%
                     		End If%>
                     		
                     	</td>
                     	<td class=tdbox>
				<ul>
				<%If ccur(GetData(13,N)) = 0 or DEF_EnableUserHidden = 0 Then
					%>
					<li>
					<%=DEF_PointsName(0) & ": " & GetData(2,n)%> / 
					<%=DEF_PointsName(4) & ": " & Clng(cCur(GetData(4,n))/60)%> / 
					<%=DEF_PointsName(3) & ": " & DEF_UserLevelString(GetData(5,n))%>
					</li>
					<%
				End If
				If CheckSupervisorUserName = 1 Then%>
					<li>
					IP: <%=GetData(3,n)%>
					</li>
					<%
				End If%>
				<li>�����/����ϵͳ��<%=htmlencode(GetData(14,n))%>/<%=htmlencode(GetData(15,n))%>
				</li>
				<li>��ǰλ��: <a href="<%=htmlencode(GetData(7,N))%>"><%
				If StrLength(GetData(8,N)) > 130 Then
					Response.Write htmlencode(LeftTrue(GetData(8,N),127)) & "..."
				Else
					Response.Write htmlencode(GetData(8,N))
				End If%></a>
				</li>
				</ul><%
			End If%>
                    </tr><%
		Next
%>
		<tr><td colspan=2 class=tdbox>
				<%=PageSplictString%>
		</td></tr>
		</table>
	<%Else%>
		<div class=alert><%
		Response.Write GBL_CHK_TempStr & "		<p>�������û���" & VbCrLf
		%>
		</div>
		<%
	End If

End Function%>