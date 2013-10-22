<!-- #include file=../plug-ins/bbschat/Inc/Chat_Setup.asp -->
<!-- #include file=Str_Fun.asp -->
<!-- #include file=MD5.asp -->
<%
Const OPEN_DEBUG = 1
Const DEBUG_User = "Super"

If inStr(Lcase(Request.Servervariables("SCRIPT_NAME")),"board_popfun.asp") > 0 Then Response.End
Const DEF_Jer = "?ver=2013022715"
Const DEF_mustDefaultStyle = -1

Dim DEF_PageExeTime1,GBL_DBNum,GBL_DBWrite,con,GBL_ConFlag,GBL_CheckPassDoneFlag,GBL_FileDir,GBL_SideFlag,CursorLocation,GBL_HeadResource
DEF_PageExeTime1 = Timer
GBL_DBNum = 0
GBL_DBWrite = 0
GBL_ConFlag = 0
GBL_CheckPassDoneFlag = 0
GBL_SideFlag = ""
CursorLocation = 0

Dim DEF_AbsolutHome
DEF_AbsolutHome = Request.ServerVariables("server_name")
If isNumeric(Replace(DEF_AbsolutHome,".","")) Then
ElseIf Len(Replace(DEF_AbsolutHome,".","")) <= Len(DEF_AbsolutHome) - 2 Then
	DEF_AbsolutHome = Mid(DEF_AbsolutHome,inStr(DEF_AbsolutHome,".") + 1)
End If

Dim dontRequestFormFlag
dontRequestFormFlag = Left(Request.QueryString("dontRequestFormFlag"),1)

Dim GBL_Board_ID
If dontRequestFormFlag = "" Then
	GBL_Board_ID = Request.QueryString("b")
	If GBL_Board_ID = "" Then GBL_Board_ID = Request.Form("b")
	If GBL_Board_ID = "" Then GBL_Board_ID = Request.QueryString("BoardID")
	If GBL_Board_ID = "" Then GBL_Board_ID = Request.Form("BoardID")
Else
	GBL_Board_ID = Request.QueryString("b")
	If GBL_Board_ID = "" Then GBL_Board_ID = Request.QueryString("BoardID")
End If

GBL_Board_ID = Left(GBL_Board_ID,14)
If isNumeric(GBL_Board_ID)=0 Then GBL_Board_ID=0
GBL_Board_ID = Fix(cCur(GBL_Board_ID))
If GBL_Board_ID > 2147479999 Then GBL_Board_ID = 0

Dim GBL_Board_BoardName,GBL_Board_BoardAssort
Dim GBL_Board_TopicNum,GBL_Board_AnnounceNum,GBL_Board_AssortName,GBL_Board_BoardLimit,GBL_Board_GoodNum
Dim GBL_Board_ForumPass,GBL_Board_HiddenFlag,GBL_Board_MasterList,GBL_Board_AllMaxRootID,GBL_Board_AllMinRootID
Dim GBL_Board_BoardStyle,GBL_Board_StartTime,GBL_Board_EndTime,GBL_Board_AssortMaster,GBL_Board_OtherLimit
GBL_Board_BoardStyle = DEF_DefaultStyle
GBL_Board_StartTime = "000000"
GBL_Board_EndTime = "000000"

Dim GBL_UserID
GBL_UserID = 0

Dim GBL_SiteHeadString,GBL_SiteBottomString,GBL_DefineImage,GBL_TableHeadString,GBL_TableBottomString,GBL_ShowBottomSure,GBL_TempletID,GBL_TempletFlag

Dim GBL_CookieTime,GBL_IPAddress
GBL_CookieTime = Request.Cookies(DEF_MasterCookies & "Time")

Dim GBL_UDT
If isArray(Session(DEF_MasterCookies & "UDT")) Then GBL_UDT = Session(DEF_MasterCookies & "UDT")
GetIPAddress

Application.Lock
Application(DEF_MasterCookies & "SitePageCount") = cCur("0" & Application(DEF_MasterCookies & "SitePageCount")) + 1
Application.UnLock

Dim GBL_CHK_Pass,GBL_CHK_User,GBL_CHK_ShowFlag,GBL_CHK_Flag,GBL_CHK_TempStr,GBL_CHK_UserLimit,GBL_CHK_CharmPoint,GBL_CHK_CachetValue
Dim GBL_CHK_Points,GBL_CHK_OnlineTime,GBL_CHK_MessageFlag,GBL_CHK_LastWriteTime,GBL_CHK_LastAnnounceID
Dim GBL_CHK_PWdFlag,GBL_CHK_GuestFlag,GBL_CookiePassFlag
GBL_CHK_PWdFlag = 1
GBL_CHK_GuestFlag = 1
GBL_CHK_LastWriteTime = 0

Dim GBL_AppType
GBL_AppType = "0"


Dim LMT_EnableRewrite
LMT_EnableRewrite = GetBinarybit(DEF_Sideparameter,16)

REM *******Chat Start*******

Sub Chat_Application_OnStart

	Dim Chat_MaxCache
	Chat_MaxCache = 50
	If Application(DEF_MasterCookies & "_Chat_Load") & "" <> "1" Then
		Dim Temp
		Redim Temp(Chat_MaxCache)
		Application.Lock
		Application(DEF_MasterCookies & "_Chat_World") = Temp
		Application(DEF_MasterCookies & "_Chat_World_Index") = 0
		Application(DEF_MasterCookies & "_Chat_Load") = "1"
		Application.UnLock
	End If

End Sub

Sub Chat_Appand_pop(c,Str)

	Chat_Application_OnStart
	Dim Temp,Index
	Temp =  Application(DEF_MasterCookies & "_Chat_World")
	Index = Application(DEF_MasterCookies & "_Chat_World_Index")
	Index = Index + 1
	If Index > Chat_MaxCache - 1 Then Index = 0
	Temp(Index) = c & " " & Str
	Application.Lock
	Application(DEF_MasterCookies & "_Chat_World_Index") = Index
	Application(DEF_MasterCookies & "_Chat_World") = Temp
	Application.UnLock

End Sub

Sub Chat_Appand_Session(Str,User)

	Dim Temp,Index
	Temp =  Application(DEF_MasterCookies & "_Chat_S_Data_" & User)
	If isArray(Temp) = False Then Exit Sub
	Index = Application(DEF_MasterCookies & "_Chat_S_Index_" & User)
	Index = Index + 1
	If Index > Chat_MaxSessionCache - 1 Then Index = 0
	Temp(Index) = "5 " & Str
	Application.Lock
	Application(DEF_MasterCookies & "_Chat_S_Index_" & User) = Index
	Application(DEF_MasterCookies & "_Chat_S_Data_" & User) = Temp
	Application.UnLock

End Sub

Sub Chat_SessionCreate(User)

	If Application(DEF_MasterCookies & "_Chat_Load") = "1" Then
		If isArray(Application(DEF_MasterCookies & "_Chat_S_Data_" & User)) = False Then
			Dim Temp
			Redim Temp(Chat_MaxSessionCache)
			Dim LMT
			If GetBinarybit(GBL_CHK_UserLimit,7) = 1 or GetBinarybit(GBL_CHK_UserLimit,1) = 1 or GetBinarybit(GBL_CHK_UserLimit,3) = 1 Then
				LMT = 1
			Else
				LMT = 0
			End If
			Application.Lock
			Application(DEF_MasterCookies & "_Chat_S_Data_" & User) = Temp '�û���Ϣ����
			Application(DEF_MasterCookies & "_Chat_S_Index_" & User) = 0 '��ǰ�����α�
			Application(DEF_MasterCookies & "_Chat_S_ID_" & User) = cCur(Session.SessionID) '��ǰsessionID
			Application(DEF_MasterCookies & "_Chat_S_LMT_" & User) = LMT '�û�Ȩ��
			Application(DEF_MasterCookies & "_Chat_S_Name_" & User) = User '�û�����
			Application.UnLock
			CALL Chat_Appand_pop(7,User)
		Else
			If Application(DEF_MasterCookies & "_Chat_S_ID_" & User) <> cCur(Session.SessionID) Then
				Application.Lock
				Application(DEF_MasterCookies & "_Chat_S_ID_" & User) = cCur(Session.SessionID)
				Application.UnLock
			End If
		End If
	End If

End Sub

Sub Chat_SessionFree(User)

	If isArray(Application(DEF_MasterCookies & "_Chat_S_Data_" & User)) Then
		If Application(DEF_MasterCookies & "_Chat_S_ID_" & User) = cCur(Session.SessionID) Then 'ֻ�е�ǰ��ȷ�û������ͷ�
			Application.Contents.Remove(DEF_MasterCookies & "_Chat_S_Data_" & User)
			Application.Contents.Remove(DEF_MasterCookies & "_Chat_S_Index_" & User)
			Application.Contents.Remove(DEF_MasterCookies & "_Chat_S_ID_" & User)
			Application.Contents.Remove(DEF_MasterCookies & "_Chat_S_LMT_" & User)
			Application.Contents.Remove(DEF_MasterCookies & "_Chat_S_Name_" & User)
			CALL Chat_Appand_pop(8,User)
		End If
	End If

End Sub
REM *******Chat End*********

Function GBL_CheckLimitTitle(Pass,lmt,otherlmt,HiddenFlag)

	're 0:none limit 1:limited
	If HiddenFlag = 2 Then
		GBL_CheckLimitTitle = 1
		Exit Function
	End If
	If DEF_LimitTitle = 0 Then
		GBL_CheckLimitTitle = 0
	Else
		If isNull(lmt) Then lmt = 0
		If isNull(otherlmt) Then otherlmt = 0
		If cCur(lmt) > 0 or cCur(otherlmt) > 0 Then
			'ȫ����
			'If (Pass <> "" or GetBinarybit(lmt,7) = 1 or GetBinarybit(lmt,2) = 1 or GetBinarybit(lmt,15) = 1 or cCur(otherlmt) > 0) Then
			'��֤���������û�����
			If (Pass <> "" or GetBinarybit(lmt,7) = 1 or GetBinarybit(lmt,2) = 1 or GetBinarybit(lmt,15) = 1) Then
			'If Pass <> "" Then
				GBL_CheckLimitTitle = 1
			Else
				GBL_CheckLimitTitle = 0
			End If
		Else
			GBL_CheckLimitTitle = 0
		End If
	End If

End Function

Function GBL_CheckLimitContent(Pass,lmt,otherlmt,HiddenFlag)

	're 0:none limit 1:limited
	If HiddenFlag = 2 Then
		GBL_CheckLimitContent = 1
		Exit Function
	End If
	If cCur(lmt) > 0 or cCur(otherlmt) > 0 Then
		If (Pass <> "" or GetBinarybit(lmt,7) = 1 or GetBinarybit(lmt,2) = 1 or GetBinarybit(lmt,15) = 1 or cCur(otherlmt) > 0) Then
			GBL_CheckLimitContent = 1
		Else
			GBL_CheckLimitContent = 0
		End If
	Else
		GBL_CheckLimitContent = 0
	End If

End Function

dim sqlstring
Function LDExeCute(sql,flag)

	'flag 0 ��,���ҷ��� 1 д,������ 2 ��,������ 3 д,����
	If GBL_ConFlag = 0 Then Exit Function
	on error resume next
	'If Err Then Err.Clear
	'Response.Write "<P>" & sql
	if OPEN_DEBUG = 1 and DEBUG_User = GBL_CHK_User Then sqlstring = sqlstring & "<li><span class=bluefont>" & FormatNumber(cCur(Timer - DEF_PageExeTime1),4,True) & "</span> " & htmlencode(sql) & "</li>" & VbCrLf
	'Response.Write "<br>sql:" & sql & "<br>Page created in " & FormatNumber(cCur(Timer - DEF_PageExeTime1),4,True) & " seconds width " & GBL_DBNum & " queries."

	If DEF_UsedDataBase = 2 Then
		sql = replace(sql,"\","\\")
	end if
	If flag = 0 or flag = 3 Then
		Set LDExeCute = Con.ExeCute(SQL)
	Else
		Con.ExeCute(SQL)
	End If
	If Err Then
		Response.Write "<p>����SQL���ִ�г�������������ֹ������ϵ�ٷ������</p><p><font color=gray>" & server.htmlencode(SQL) & "</font></P>"
		Response.Write "<p>��������: <font color=red>" & err.description & "</font></p>"
		CloseDatabase
		Response.End
	End If
	

	GBL_DBNum = GBL_DBNum + 1
	If flag = 1 or flag = 3 Then
		If GBL_DBWrite = 0 Then
			If GBL_CheckPassDoneFlag = 0 Then
				If isArray(GBL_UDT) Then
					GBL_CHK_LastWriteTime = DateDiff("s",RestoreTime(GBL_UDT(13)),DEF_Now)
				Else
					GBL_UserID = CheckPass
				End If
			End If
			'CheckWriteEventSpace
		End If
		GBL_DBWrite = GBL_DBWrite + 1
		UpdateLastWriteTime
	End If

End Function

Sub ReloadBoardInfo(ID)

	If GBL_ConFlag = 0 Then Exit Sub
	Dim Rs,GetData
	Set Rs = LDExeCute(sql_select("Select T1.BoardName,T1.BoardAssort,T1.BoardIntro,T1.LastWriter,T1.LastWriteTime,T1.TopicNum,T1.AnnounceNum,T1.ForumPass,T1.HiddenFlag,T1.BoardLimit,T1.MasterList,T1.AllMaxRootID,T1.AllMinRootID,T1.GoodNum,T2.AssortName,T1.BoardStyle,T1.StartTime,T1.EndTime,T1.TodayAnnounce,T1.LastAnnounceID,T1.LastTopicName,T1.BoardImgUrl,T1.BoardImgWidth,T1.BoardImgHeight,T1.BoardHead,T1.BoardBottom,T1.ParentBoard,T1.LowerBoard,T1.ParentBoardStr,T1.TopicNum_All,T1.AnnounceNum_All,T1.TodayAnnounce_All,T1.GoodNum_All,0,0,T2.AssortMaster,T1.OtherLimit,T2.AssortLimit from LeadBBS_Boards as T1 left join LeadBBS_Assort as T2 on t1.BoardAssort=T2.AssortID where T1.BoardID=" & ID,1),0)
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		GBL_Board_ID = 0
		GBL_Board_BoardLimit = 0
	Else
		GetData = Rs.GetRows(1)
		Rs.Close
		Set Rs = Nothing
		Application.Lock			
		Application(DEF_MasterCookies & "BoardInfo" & ID) = GetData
		Application.UnLock
	End If

End Sub

Sub Borad_GetBoardIDValue(ID)

	Dim Temp
	Temp = Application(DEF_MasterCookies & "BoardInfo" & ID)
	If isArray(Temp) = False Then
		ReloadBoardInfo(ID)
		Temp = Application(DEF_MasterCookies & "BoardInfo" & ID)
		If isArray(Temp) = False Then GBL_Board_ID = 0
	End If
	If GBL_Board_ID = 0 Then Exit Sub

	GBL_Board_BoardName = Temp(0,0)
	GBL_Board_BoardAssort = cCur(Temp(1,0))
	'GBL_Board_BoardIntro = Temp(2,0)
	'GBL_Board_LastWriter = Temp(3,0)
	'GBL_Board_LastWriteTime = Temp(4,0)
	GBL_Board_TopicNum = cCur(Temp(5,0))
	GBL_Board_AnnounceNum = cCur(Temp(6,0))
	GBL_Board_AssortName =Temp(14,0)
	GBL_Board_ForumPass = Temp(7,0)
	GBL_Board_HiddenFlag = Temp(8,0)
	GBL_Board_BoardLimit = cCur(Temp(9,0))
	GBL_Board_MasterList = Temp(10,0)
	GBL_Board_AllMaxRootID = cCur(Temp(11,0))
	GBL_Board_AllMinRootID = cCur(Temp(12,0))
	GBL_Board_GoodNum = cCur(Temp(13,0))
	If Temp(15,0) <> 0 Then GBL_Board_BoardStyle = Temp(15,0)
	GBL_Board_StartTime = Right("000000" & Temp(16,0),6)
	GBL_Board_EndTime = Right("000000" & Temp(17,0),6)
	'GBL_Board_TodayAnnounce = Temp(18,0)
	'LastAnnounceID = 19,0
	',T1.LastTopicName= 20,0
	'21,22,23
	GBL_SiteHeadString = Temp(24,0)
	GBL_SiteBottomString = Temp(25,0)
	'26,27,T1.ParentBoard,T1.LowerBoard
	'28,T1.ParentBoardStr,29.TopicNum_All,30.AnnounceNum_All,31.TodayAnnounce_All,32.GoodNum_All
	GBL_Board_AssortMaster = Temp(35,0)
	GBL_Board_OtherLimit = cCur(Temp(36,0))
	'GBL_Board_AssortLimit = cCur(Temp(37,0))

	If GBL_CHK_PWdFlag = 0 and (GBL_Board_BoardLimit <> 0 or GBL_Board_ForumPass <> "" or GBL_Board_OtherLimit > 0) Then
		If GBL_Board_ForumPass <> "" or GBL_Board_OtherLimit > 0 or GetBinarybit(GBL_Board_BoardLimit,2) = 1 or GetBinarybit(GBL_Board_BoardLimit,7) = 1 or GetBinarybit(GBL_Board_BoardLimit,15) = 1 Then
			GBL_CHK_PWdFlag = 1
		ElseIf GBL_CHK_GuestFlag = 1 and GetBinarybit(GBL_Board_BoardLimit,1) = 1 Then
			GBL_CHK_PWdFlag = 1
		End If
	End If

End Sub

Sub BBS_SiteHead(headString,BoardID,Str)

	SiteHead(headString)
	%>
	<div class="body_area_out">
	<%
	If Left(headString,7) <> "       " Then DisplayBBSNavigate BoardID,Str

End Sub

Sub UpdateOnlineUserAtInfo(BoardID,AtInfo)

	If BoardID <> 0 and isArray(Application(DEF_MasterCookies & "BoardInfo" & BoardID)) = False Then
		Exit Sub
	End If

	If GBL_CheckPassDoneFlag = 0 Then
		If isArray(GBL_UDT) Then
			GBL_CHK_LastWriteTime = DateDiff("s",RestoreTime(GBL_UDT(13)),DEF_Now)
		Else
			GBL_UserID = CheckPass
		End If
	End If
	If isArray(GBL_UDT) Then
		If CheckWriteEventSpace = 0 Then Exit Sub
	End If
	Dim AtBoardIDCookie
	AtBoardIDCookie = Left(Request.Cookies(DEF_MasterCookies & "AtBD"),14)
	If isNumeric(AtBoardIDCookie) = False or AtBoardIDCookie = "" Then
		If isArray(GBL_UDT) Then
			Response.Cookies(DEF_MasterCookies & "AtBD") = BoardID
			Response.Cookies(DEF_MasterCookies & "AtBD").Domain = DEF_AbsolutHome
		End If
	Else
		AtBoardIDCookie = cCur(AtBoardIDCookie)
		If isTrueDate(GBL_CookieTime) = 1 Then
			If DateDiff("s",GBL_CookieTime,DEF_Now) < 240 and (AtBoardIDCookie = BoardID or GetBinarybit(DEF_Sideparameter,6) = 0) Then Exit Sub
			If AtBoardIDCookie <> BoardID and isArray(GBL_UDT) Then
				Response.Cookies(DEF_MasterCookies & "AtBD") = BoardID
				Response.Cookies(DEF_MasterCookies & "AtBD").Domain = DEF_AbsolutHome
			End If
		Else
			Response.Cookies(DEF_MasterCookies & "Time") = DEF_Now
			Response.Cookies(DEF_MasterCookies & "Time").Domain = DEF_AbsolutHome
		End If
	End If

	Dim Rs,SQL,OL2
	SQL = sql_select("Select AtBoardID,ID,LastDoingTime from LeadBBS_OnlineUser where SessionID=" & cCur(Session.SessionID),2)
	Set Rs = LDExeCute(SQL,0)
	If Not Rs.Eof Then
		SQL = cCur(Rs(0))
		GBL_CHK_LastWriteTime = DateDiff("s",RestoreTime(cCur(Rs(2))),DEF_Now)
		Rs.MoveNext
		If Not Rs.Eof Then
			OL2 = Rs(1)
		Else
			OL2 = 0
		End If
		Rs.Close
		Set Rs = Nothing
		'Ϊ�οͼ���д�����ж�
		If CheckWriteEventSpace = 0 Then Exit Sub
		If cCur(OL2) > 0 Then
			UpdateOnlineUserInfo("from LeadBBS_onlineUser where ID=" & OL2)
			'Application.Lock
			'Application(DEF_MasterCookies & "ActiveUsers") = cCur("0" & Replace("" & Application(DEF_MasterCookies & "ActiveUsers"),"-","")) - 1
			'Application.UnLock
		End If
		If isArray(GBL_UDT) = False Then
			Response.Cookies(DEF_MasterCookies & "AtBD") = BoardID
			Response.Cookies(DEF_MasterCookies & "AtBD").Domain = DEF_AbsolutHome
		End If
		If SQL <> BoardID Then
			Application.Lock
			If BoardID <> 0 Then Application(DEF_MasterCookies & "BDOL" & BoardID) = Application(DEF_MasterCookies & "BDOL" & BoardID) + 1
			If SQL <> 0 Then Application(DEF_MasterCookies & "BDOL" & SQL) = Application(DEF_MasterCookies & "BDOL" & SQL) - 1
			Application.UnLock
			Call LDExeCute("Update LeadBBS_OnlineUser Set LastDoingTime=" & GetTimeValue(DEF_Now) & ",AtBoardID=" & BoardID & ",AtInfo='" & Replace(Left(KillHTMLLabel(AtInfo),255),"'","''") & "',AtUrl='" & Replace(Left(Request.Servervariables("SCRIPT_NAME") & "?" & Request.QueryString,255),"'","''") & "' where SessionID=" & cCur(Session.SessionID),1)
		Else
			Call LDExeCute("Update LeadBBS_OnlineUser Set LastDoingTime=" & GetTimeValue(DEF_Now) & ",AtInfo='" & Replace(Left(KillHTMLLabel(AtInfo),255),"'","''") & "',AtUrl='" & Replace(Left(Request.Servervariables("SCRIPT_NAME") & "?" & Request.QueryString,255),"'","''") & "' where SessionID=" & cCur(Session.SessionID),1)
		End If
	Else
		Rs.Close
		Set Rs = Nothing
	End If

End Sub

Sub UpdateOnlineUserInfo(Str)

	Dim Rs,SQL,GetData,N,delN
	SQL = "Select AtBoardID " & Str
	Set Rs = LDExeCute(SQL,0)
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		Exit Sub
	End If
	GetData = Rs.GetRows(-1)
	Rs.Close
	Set Rs = Nothing
	
	SQL = 0
	Rs = Ubound(GetData,2)
	For N = 0 to Rs
		SQL = cCur(GetData(0,N))
		If SQL > 0 Then 
			Application.Lock
			Application(DEF_MasterCookies & "BDOL" & SQL) = Application(DEF_MasterCookies & "BDOL" & SQL) - delN
			Application.UnLock
		End If
	Next
	Application.Lock
	Application(DEF_MasterCookies & "ActiveUsers") = Application(DEF_MasterCookies & "ActiveUsers") - (Rs + 1)
	Application.UnLock
	
	Call LDExeCute("Delete " & Str,1)

End Sub

Sub Boards_Body_Head(str)

	Dim side
	side = "_close"
	If left(str,7) = "request" Then
		side = Request.Cookies(DEF_MasterCookies & "_side")
		If inStr(",_left,_right,_close,","," & side & ",") = 0 Then
			if str = "request0" Then
				side = "_right"
			Else
				side = "_close"
			End If
		End If
	Else
		side = "_close"
	End If
	%>
<div class="area">
<div class="main"><%If left(str,7) = "request" Then%>
	<div class="content_side<%=side%>" id="p_side">
		<%Boars_Side_Box(side)%>
	</div><%End If%>
	<div class="content_main<%=side%>">
		<div class="content_main_2<%=side%>">
		<div class="content_main_body">
	<%

End Sub

Sub Boards_Body_Bottom

%>
		</div>
	</div>
	</div>
</div>
</div>
<%

End Sub

Sub DisplayBBSNavigate(BoardID,Str)

	'Global_SmallTableHead
	%>
	
	<div class="navigate_sty_out">
	<div class="area">
		<div class="navigate_sty">
			<div class="navigate_string">
			<%
			'If DEF_SiteHomeUrl = "" Then DEF_SiteHomeUrl = DEF_BBS_HomeUrl & "Boards.asp"
			If GBL_Board_BoardAssort = "" and Str = "" Then
				If DEF_SiteHomeUrl <> "" and DEF_SiteNameString <> "" Then
					Response.Write "<span class=""navigate_string_home"">" & DEF_SiteNameString & "</span>"
					Response.Write "<span class=""navigate_string_step"">" & DEF_BBS_Name & "</span>"
				else
					Response.Write "<span class=""navigate_string_home"">" & DEF_BBS_Name & "</span>"
				end if
			Else
				If DEF_SiteHomeUrl <> "" and DEF_SiteNameString <> "" Then
					Response.Write "<span class=""navigate_string_home""><a href=" & DEF_SiteHomeUrl & ">" & DEF_SiteNameString & "</a></span>"
					Response.Write "<span class=""navigate_string_step""><a href=""" & DEF_BBS_HomeUrl & "Boards.asp"">" & DEF_BBS_Name & "</a></span>"
				else
					Response.Write "<span class=""navigate_string_home""><a href=""" & DEF_BBS_HomeUrl & "Boards.asp"">" & DEF_BBS_Name & "</a></span>"
				end if
				If GBL_Board_BoardName="" Then 
					If GBL_Board_AssortName<>"" Then Response.Write "<span class=""navigate_string_step"">" & GBL_Board_AssortName & "</span>"
				Else
					If GBL_Board_AssortName<>"" Then Response.Write "<span class=""navigate_string_step""><a href=""" & DEF_BBS_HomeUrl & "Boards.asp?Assort=" & GBL_Board_BoardAssort & """>" & GBL_Board_AssortName & "</a></span>"
					If BoardID > 0 Then
						Dim Temp,TempStr,N
						Temp = cCur(Application(DEF_MasterCookies & "BoardInfo" & GBL_Board_ID)(26,0))
						Do While Temp > 0
							If isArray(Application(DEF_MasterCookies & "BoardInfo" & Temp)) = False Then
								ReloadBoardInfo(Temp)
								If isArray(Application(DEF_MasterCookies & "BoardInfo" & Temp)) = False Then Exit Do
							End If
							TempStr = "<span class=""navigate_string_step""><a href=""" & DEF_BBS_HomeUrl & "b/b.asp?B=" & Temp & """>" & Application(DEF_MasterCookies & "BoardInfo" & Temp)(0,0) & "</a></span>" & TempStr
							Temp = cCur(Application(DEF_MasterCookies & "BoardInfo" & Temp)(26,0))
							N = N + 1
							If N > 10 Then Exit Do
						Loop
						Response.Write TempStr
						Response.Write "<span class=""navigate_string_step""><a href=""" & DEF_BBS_HomeUrl & "b/b.asp?B=" & GBL_Board_ID & """>" & GBL_Board_BoardName & "</a></span>"
					End If
				End If
			End If
			Response.write Str%>
		</div>
		<%navigate_sidecontrol%>
	</div>
	</div>
	</div>
	<%

End Sub

Sub navigate_sidecontrol

	If GBL_SideFlag = "" or Left(GBL_SideFlag,1) = "1" Then Exit Sub
%>
		<div class="navigate_sidecontrol">
			<table border="0" cellspacing="0" cellpadding="0" align="center"><tr><%
			Dim side,SideBak
			side = Request.Cookies(DEF_MasterCookies & "_side")
			
			If side = "" and inStr(",_left,_right,_close,","," & side & ",") = 0 Then
				if right(GBL_SideFlag,1) = "0" Then
					side = "_right"
				Else
					side = "_close"
				End If
			End If

			SideBak = side
			If side = "_left" or side = "_close" Then
				side = "_right"
			Else
				side = "_left"
			End If
			%>
			<td valign="top"><div class="<%If SideBak = "_close" Then side ="_close"%>p_side<%=side%>" id="p_side_img"><a href="javascript:void(0)" oncontextmenu="swap_col('close');return false;" onclick="swap_col();" title="�����л�/(�һ��ر�)" class="unsel" hidefocus="true"></a></div>
			</td></tr></table>
		</div>
		
	<script type="text/javascript">
	<!--
	function swap_col(ty)
	{
		var f="_right",f1="_left";
		if($$("content_side_left")[0])
		{
			f="_left";
			f1="_right";
		}
		else if($$("content_side_close")[0])
		{
			f="_close";
			f1="_right";
		}
		else
		{
			if(!$$("content_side_right")[0])return;
		}
		var tmp = $$("content_side" + f)[0];
		if(ty=="close")
		{
			tmp.className = "content_side_close";
			$$("content_main" + f)[0].className = "content_main_close";
			$$("content_main_2" + f)[0].className = "content_main_2_close";
			$id("p_side_img").className = "p_side_close";				
			//$id("p_side_closeimg").style.display = "none";
			LD.Cookie.Add(DEF_MasterCookies + "_side","_close")
		}
		else
		{
			if($id('p_side').innerHTML.length < 30)
			{
				var js = "$(\"#p_side\").html(tmp);"
			<%If GBL_Board_ID < 1 Then%>
				getAJAX("Boards.asp","ol=side",js,1);
			<%Else%>
				getAJAX("<%=DEF_BBS_HomeUrl%>b/b.asp?b=<%=GBL_Board_ID%>","ol=side",js,1);
			<%End If%>
			}
			//$id("p_side_closeimg").style.display = "block";
			$$("content_side"+f)[0].className = "content_side"+f1;
			$$("content_main"+f)[0].className = "content_main"+f1;
			$$("content_main_2"+f)[0].className = "content_main_2"+f1;
			if($id("p_side_img").className=='p_side_left')
				$id("p_side_img").className = "p_side_right";
			else
				$id("p_side_img").className = "p_side_left";
			LD.Cookie.Add(DEF_MasterCookies + "_side",f1)
		}
	}
	-->
	</script>
<%
End Sub

Sub DisplayInfoBoxNavigate

REM *******Chat Start*******
If Chat_EnablePageRequest = 1 Then
%>
<%If GetBinarybit(GBL_CHK_UserLimit,17) = 0 and (GBL_CHK_MessageFlag = 1) Then%><bgsound src=<%=DEF_BBS_HomeUrl%>images/NOTIFY.wav border=0 loop="1"><%End If%>
<a href="<%=DEF_BBS_HomeUrl%>User/MyInfobox.asp" id="c_pub_mes" class="head_privatemsg<%
If (GBL_CHK_MessageFlag = 1) Then Response.Write "_new"
%>" title="����Ϣ��ʾ"><span id="c_pub_mes_txt"><%
If GBL_CHK_MessageFlag <> 1 Then%>
�ռ���<%
Else%>�����µ���Ϣ<%
End If%></span></a>
<script type="text/javascript">
<!--
var c_infoflag = <%
If GBL_CHK_MessageFlag = 1 Then
	Response.Write "1"
Else
	Response.Write "0"
End If
%>;
var c_home="<%=DEF_BBS_HomeUrl%>",c_homeurl="<%=DEF_BBS_HomeUrl%>plug-ins/bbschat/";
var c_User="<%=urlencode(GBL_CHK_User)%>";
-->
</script>
<script src="<%=DEF_BBS_HomeUrl%>plug-ins/bbschat/inc/chat_pubmsg.js<%=DEF_Jer%>" type="text/javascript"></script>
<%
Exit Sub '������Ϣ���˳�
End If
REM *******Chat End*********
	If GBL_CHK_MessageFlag <> 1 Then
		If GBL_CHK_User <> "" Then%>
					<a href="<%=DEF_BBS_HomeUrl%>User/MyInfobox.asp" class="head_privatemsg">�ռ���</a><%
		End If
	Else
		%>
					<%If GetBinarybit(GBL_CHK_UserLimit,17) = 0 Then%><bgsound src="<%=DEF_BBS_HomeUrl%>images/NOTIFY.wav" border="0" loop="1"><%End If%>
					<a href="<%=DEF_BBS_HomeUrl%>User/MyInfobox.asp" class="head_privatemsg<%
If (GBL_CHK_MessageFlag = 1) Then Response.Write "_new"
%>">�����µ���Ϣ</a>
		<%
	End if

End Sub

Sub CloseDatabase

	If GBL_ConFlag = 0 Then Exit Sub
	Application.Lock
	If GBL_DBWrite > 0 Then Application(DEF_MasterCookies & "DBWrite") = cCur("0" & Application(DEF_MasterCookies & "DBWrite")) + GBL_DBWrite
	If GBL_DBNum > 0 Then Application(DEF_MasterCookies & "DBNum") = cCur("0" & Application(DEF_MasterCookies & "DBNum")) + GBL_DBNum
	Application.UnLock
	If Application(DEF_MasterCookies & "DBWrite") > 2999 Then
		Application.Lock
		Application(DEF_MasterCookies & "DBWrite") = 0
		Application(DEF_MasterCookies & "DBNum") = Application(DEF_MasterCookies & "DBNum") + 1
		Application.UnLock
		Call LDExeCute("Update LeadBBS_SiteInfo Set DBWrite=DBWrite+3001",1)
	End If
	If Application(DEF_MasterCookies & "DBNum") > 2999 Then
		Application.Lock
		Application(DEF_MasterCookies & "DBNum") = 0
		Application(DEF_MasterCookies & "DBWrite") = Application(DEF_MasterCookies & "DBWrite") + 1
		Application.UnLock
		Call LDExeCute("Update LeadBBS_SiteInfo Set DBNum=DBNum+3001",1)
	End If
		
	If DEF_EnableDatabaseCache = 1 Then
		Set con = Nothing
	Else
		Con.close
		Set con = Nothing
	End If
	GBL_ConFlag = 0

End Sub

Sub OpenDatabase

	'on error resume next
	If DEF_EnableDatabaseCache = 1 and DEF_UsedDataBase = 1 Then
		If isObject(Application(DEF_MasterCookies & "con")) = False Then
			Set Con = Server.CreateObject("ADODB.Connection")
			Select case DEF_UsedDataBase
				Case 1:
					Con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath(DEF_BBS_HomeUrl & DEF_AccessDatabase)
				Case 2:
					Con.ConnectionString = DEF_AccessDatabase
				Case else
					Con.ConnectionString = DEF_AccessDatabase
			End select
			Con.Open
			If Err Then
				Err.Clear
				Set Con = Nothing
				GBL_CHK_TempStr = "LeadBBS: connect database error."
				Response.Write GBL_CHK_TempStr
				Response.End
			End If
			Application.Lock
			Set Application(DEF_MasterCookies & "con") = con
			Application.UnLock
		Else
			Set con = Application(DEF_MasterCookies & "con")
		End If
	Else
		Set con = Server.CreateObject("ADODB.Connection")
		Select case DEF_UsedDataBase
			Case 1:
				Con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath(DEF_BBS_HomeUrl & DEF_AccessDatabase)
			Case 2:
				Con.ConnectionString = DEF_AccessDatabase
				if CursorLocation = 3 then con.CursorLocation = 3
			case Else
				Con.ConnectionString = DEF_AccessDatabase
		End Select
		Con.Open
		If Err Then
			Err.Clear
			Set Con = Nothing
			GBL_CHK_TempStr = "LeadBBS: connect database error."
			Response.Write GBL_CHK_TempStr
			Response.End
		End If
	End If
	GBL_ConFlag = 1
	If inStr(Application(DEF_MasterCookies & "Version") & "","LeadBBS") = 0 Then ReloadVesion
	DEF_Version = Application(DEF_MasterCookies & "Version")
	If Application(DEF_MasterCookies & "SitePageCount") > 98 Then
		Application.Lock
		Application(DEF_MasterCookies & "SitePageCount") = 0
		Application.UnLock
		Call LDExeCute("Update LeadBBS_SiteInfo Set PageCount=PageCount+99",1)
		UpdateStatisticDataInfo 99,4,1
	End If
	If IsArray(Application(DEF_MasterCookies & "BListAll")) = False Then ReloadBoardListData

End Sub

Sub initDatabase

	OpenDatabase
	CheckUserOnline

End Sub

Sub ReloadBoardListData

	Dim Rs,GetData
	Set Rs = LDExeCute("Select BoardID,BoardAssort from LeadBBS_Boards where ParentBoard=0 and HiddenFlag = 0 order by BoardAssort,OrderID ASC",0)
	If Not Rs.Eof Then
		GetData = Rs.GetRows(-1)
		Application.Lock
		Application(DEF_MasterCookies & "BList") = GetData
		Application.UnLock
	End If
	Rs.Close
	Set Rs = Nothing
	Set Rs = LDExeCute("Select BoardID,BoardAssort from LeadBBS_Boards",0)
	If Not Rs.Eof Then
		GetData = Rs.GetRows(-1)
		Application.Lock
		Application(DEF_MasterCookies & "BListAll") = GetData
		Application.UnLock
	End If
	Rs.Close
	Set Rs = Nothing

End Sub

Sub GetIPAddress

	GBL_IPAddress = Left(Replace(Request.ServerVariables("HTTP_X_FORWARDED_FOR"),"'",""),15)
	If GBL_IPAddress = "" Then GBL_IPAddress = Left(Replace(Request.ServerVariables("REMOTE_ADDR"),"'",""),15)

End Sub

Sub RepairOnlineUser

	Dim GetData	
	GetData = Application(DEF_MasterCookies & "BListAll")
	If isArray(GetData) = False Then
		ReloadBoardListData
		GetData = Application(DEF_MasterCookies & "BListAll")
	End If
	If isArray(GetData) = False Then Exit Sub

	Dim N,m,i,Rs
	m = Ubound(GetData,2)
	Server.ScriptTimeOut = 600
	SetActiveUserCount
	For N = 0 to m
		Set Rs = LDExeCute("Select count(*) from LeadBBS_OnlineUser Where AtBoardID=" & GetData(0,n),0)
		If Rs.Eof Then
			i = 0
		Else
			i = Rs(0)
			If isNull(i) Then i = 0
			i = cCur(i)
		End If
		Rs.Close
		Set Rs = Nothing
		Application.Lock
		Application(DEF_MasterCookies & "BDOL" & GetData(0,n)) = i
		Application.UnLock
	Next

End Sub

function get_index(i)

	select case DEF_UsedDataBase
	case 0: get_index = " with (index(" & i & ")) "
	case 1: get_index = ""
	case 2: get_index = "use index (" & i & ") "
	end select

end function
	
Function sql_select(sql,topn)

 	select Case DEF_UsedDataBase
	Case 2:
		sql_select = sql & " limit " & topn
	case else
		if lcase(left(sql,16)) = "select distinct " then
			sql_select = replace(sql,"select distinct ","select distinct top " & topn &" ",1,1,1)
		else
			sql_select = replace(sql,"select ","select top " & topn &" ",1,1,1)
		end if
	end select

End Function

Sub CheckUserOnline

	Dim Rs,Count,SQL,TmpSessionID,DayFlag

	Dim I
	If isNumeric(Replace(GBL_IPAddress,".","")) = 0 or (Replace(GBL_IPAddress,".","",1,3,0) = Replace(GBL_IPAddress,".","",1,2,0)) Then
		'Response.Write "��̳��ֹ�Ƿ�IP��ַ�߷���"
		'CloseDatabase
		'Response.End
		GBL_IPAddress = "1.1.1.1"
	End If
	Dim NIP,NewIP,tmp
	If DEF_EnableForbidIP = 1 Then
		NIP = GBL_IPAddress
		I = inStr(NIP,".")
		NewIP = Left(NIP,I-1)
		
		NIP = Mid(NIP,I+1)
		I = inStr(NIP,".")
		tmp = Left(NIP,I-1)
		NewIP = NewIP & Right("00" & tmp,3)
		
		NIP = Mid(NIP,I+1)
		I = inStr(NIP,".")
		tmp = Left(NIP,I-1)
		NewIP = NewIP & Right("00" & tmp,3)
		NIP = Mid(NIP,I+1)
		NewIP = NewIP & Right("00" & NIP,3)
		
		SQL = sql_select("Select id from LeadBBS_ForbidIP Where IPStart<=" & NewIP & " and IPEnd>=" & NewIP,1)
		Set Rs = LDExeCute(SQL,0)
		If Not Rs.Eof Then
			Rs.Close
			Set Rs = Nothing
			CloseDatabase
			Response.Write "Forbid IP."
			Response.End
		End If
		Rs.Close
		Set Rs = Nothing
	End If
	If GBL_Board_ID > 0 Then Borad_GetBoardIDValue(GBL_Board_ID)

	If cCur(Session.SessionID) < 1 or isNull(Session.SessionID) or Session.SessionID = "" or isNumeric(Session.SessionID) = 0 Then
		Response.Write "�������֧��Cookie"
		CloseDatabase
		Response.End
	End If

	DayFlag = Application(DEF_MasterCookies & "UserDateChangesoieiu")
	If isNumeric(DayFlag) = False or DayFlag = "" Then
		Application.Lock
		Application(DEF_MasterCookies & "UserDateChangesoieiu") = 0
		Application.UnLock
		DayFlag = 0
	End If

	If DayFlag <> Day(DEF_Now) and DayFlag <> 0 Then
		Application.Lock
		Application(DEF_MasterCookies & "UserDateChangesoieiu") = Day(DEF_Now)
		Application.UnLock
		
		Set Rs = LDExeCute(sql_select("Select YesterDay from LeadBBS_SiteInfo",1),0)
		If Rs.Eof Then
			DayFlag = 0
		Else
			DayFlag = cCur(Rs(0))
			If DayFlag > 0 Then
				If Day(RestoreTime(DayFlag)) = Day(DEF_Now) Then DayFlag = 0
			Else
				DayFlag = -1
			End If
		End If

		If DayFlag <> 0 Then
			Con.CommandTimeout = 600
			Set Rs = LDExeCute("select sum(TodayAnnounce) from LeadBBS_Boards",0)
			If Rs.Eof Then
				SQL = 0
			Else
				SQL = Rs(0)
				If isNull(SQL) Then SQL = 0
				SQL = cCur(SQL)
			End If
			Rs.Close
			Set Rs = Nothing
			Call LDExeCute("Update LeadBBS_SiteInfo set YesterdayAnc=" & SQL & ",YesterDay=" & GetTimeValue(DEF_Now),1)
			Dim TmpData
			TmpData = Application(DEF_MasterCookies & "StatisticData")
			If isArray(TmpData) = False Then
				ReloadStatisticData
				TmpData = Application(DEF_MasterCookies & "StatisticData")
			End If
			If SQL > cCur(TmpData(6,0)) Then
				Call LDExeCute("Update LeadBBS_SiteInfo Set MaxAnnounce=" & SQL & ",MaxAncTime=" & Left(GetTimeValue(DateAdd("d",-1,DEF_Now)),8) & "235959",1)
				UpdateStatisticDataInfo SQL,6,0
				UpdateStatisticDataInfo cCur(Left(GetTimeValue(DateAdd("d",-1,DEF_Now)),8) & "235959"),7,0
			End If
			Call LDExeCute("Update LeadBBS_Boards Set TodayAnnounce=0,TodayAnnounce_All=0",1)
			ReloadStatisticData
			If isArray(Application(DEF_MasterCookies & "BListAll")) = True Then
				SQL = Ubound(Application(DEF_MasterCookies & "BListAll"),2)
				For I = 0 To SQL
					UpdateBoardApplicationInfo Application(DEF_MasterCookies & "BListAll")(0,I),0,18
					UpdateBoardApplicationInfo Application(DEF_MasterCookies & "BListAll")(0,I),0,31
				Next
			End If
			Call LDExeCute("insert into LeadBBS_Log(LogType,LogTime,LogInfo,UserName,IP,BoardID) Values(0," & GetTimeValue(DEF_Now) & ",'��̳�����µ�һ�죬�ɹ����һϵ�и��¡�','" & Replace(Replace(htmlencode(Left(GBL_CHK_User,14)),"\","\\"),"'","''") & "','" & Replace(GBL_IPAddress,"'","''") & "'," & GBL_Board_ID & ")",1)
		End If
	ElseIf DayFlag = 0 then
		DayFlag = Day(DayFlag)
		Application.Lock
		Application(DEF_MasterCookies & "UserDateChangesoieiu") = DayFlag
		Application.UnLock
		RepairOnlineUser
		Call LDExeCute("insert into LeadBBS_Log(LogType,LogTime,LogInfo,UserName,IP,BoardID) Values(0," & GetTimeValue(DEF_Now) & ",'ϵͳ���������Ѿ��ɹ�����ͳ�ƣ�����ԭ������̳��WEB������������','" & Replace(Replace(htmlencode(Left(GBL_CHK_User,14)),"\","\\"),"'","''") & "','" & Replace(GBL_IPAddress,"'","''") & "'," & GBL_Board_ID & ")",1)
	End If

	If GBL_CHK_PWdFlag = 1 Then CheckPass
	If isTrueDate(GBL_CookieTime) = 0 Then
		Response.Cookies(DEF_MasterCookies & "Time") = DEF_Now
		Response.Cookies(DEF_MasterCookies & "Time").Domain = DEF_AbsolutHome
		'Exit sub '�˴������Ļ����������������ʱ��һ�������µ��û���Ϊ�����û�
	Else
		LastDoingTime = DateDiff("s",GBL_CookieTime, DEF_Now)
		If LastDoingTime<0 or LastDoingTime > DEF_UserOnlineTimeOut Then
			LastDoingTime = 0
			Response.Cookies(DEF_MasterCookies & "Time") = DEF_Now
			Response.Cookies(DEF_MasterCookies & "Time").Domain = DEF_AbsolutHome
			Exit sub
		End If
		If LastDoingTime < 240 Then '240��
			Exit sub
		Else
			Response.Cookies(DEF_MasterCookies & "Time") = DEF_Now
			Response.Cookies(DEF_MasterCookies & "Time").Domain = DEF_AbsolutHome
		End If
	End If

	If GBL_CHK_PWdFlag = 0 Then CheckPass
	SQL = cCur(Timer)
	If SQL>(cCur("0" & Application(DEF_MasterCookies & "UserRefreshNum1oieiu"))+DEF_UserOnlineTimeOut) or SQL<cCur("0" & Application(DEF_MasterCookies & "UserRefreshNum1oieiu")) Then
		Application.Lock
		SQL = Application(DEF_MasterCookies & "ActiveUsers")
		Application(DEF_MasterCookies & "UserRefreshNum1oieiu") = Timer
		Application.UnLock
		Con.CommandTimeout = 600
		Server.ScriptTimeOut = 600
		Call LDExeCute("delete from LeadBBS_onlineUser where LastDoingTime<" & GetTimeValue(DateAdd("s", 0-DEF_UserOnlineTimeOut, DEF_Now)),1)
		RepairOnlineUser
		Call LDExeCute("insert into LeadBBS_Log(LogType,LogTime,LogInfo,UserName,IP,BoardID) Values(0," & GetTimeValue(DEF_Now) & ",'ϵͳ����������ʱ���³ɹ�,���ǰ����" & SQL & "��,��" & Application(DEF_MasterCookies & "ActiveUsers") & "�ˡ�','" & Replace(Replace(htmlencode(Left(GBL_CHK_User,14)),"\","\\"),"'","''") & "','" & Replace(GBL_IPAddress,"'","''") & "'," & GBL_Board_ID & ")",1)
		Rem ��ʱ��ɾ���κζ���Ϣ
		Call LDExeCute("delete from LeadBBS_InfoBox where ExpiresDate<" & CLng(Left(GetTimeValue(Now),8)) & " and ExpiresDate>0",1)
	End If

	'If GBL_IPAddress = "61.154.122.50" Then Exit Sub
	If GBL_UserID > 0 Then
		SQL = sql_select("select UserID,LastDoingTime,SessionID,ID from LeadBBS_onlineUser where UserID=" & GBL_UserID,2)
	Else
		SQL = sql_select("select UserID,LastDoingTime,SessionID,ID from LeadBBS_onlineUser where IP='" & Replace(GBL_IPAddress,"'","''") & "' and UserID=0",2)
	End If
	Set Rs = LDExeCute(SQL,0)

	If GBL_CHK_ShowFlag = 1 and DEF_EnableUserHidden = 1 Then
		i = "�����û�"
		tmp = 0
	Else
		If GBL_UserID > 0 Then
			i = GBL_CHK_User
			tmp = GBL_CHK_UserLimit
		Else
			I = ""
			tmp = 0
		End If
	End If
	Dim OL2
	If Not Rs.Eof Then
		Dim LastDoingTime
		SQL = cCur(Rs(0))
		TmpSessionID = cCur(Rs(2))
		NIP = cCur(Rs(3))
		LastDoingTime = DateDiff("s",RestoreTime(Rs(1)), DEF_Now)
		If LastDoingTime < 0 or LastDoingTime > DEF_UserOnlineTimeOut Then LastDoingTime = 0
		Rs.MoveNext
		If Not Rs.Eof Then
			OL2 = Rs(3)
			Count = 2
		Else
			Count = 1
		End If
		Rs.Close
		Set Rs = Nothing
		If LastDoingTime > 240 Then '240�뱣��һ�ξ���
			Call LDExeCute("Update LeadBBS_onlineUser set LastDoingTime=" & GetTimeValue(DEF_Now) & ",IP='" & GBL_IPAddress & "',SessionID=" & cCur(Session.SessionID) & " where ID=" & NIP,1)
			UpdateSessionValue 17,GBL_IPAddress,0
			UpdateSessionValue 18,GetTimeValue(DEF_Now),0

			If GBL_UserID > 0 and SQL > 0 Then
				Call LDExeCute("Update LeadBBS_User set OnlineTime=OnlineTime+" & LastDoingTime & ",LastDoingTime=" & GetTimeValue(DEF_Now) & " where ID=" & GBL_UserID,1)
				UpdateSessionValue 5,LastDoingTime,1
REM *******Chat Start*******
If GBL_UserID > 0 Then
	CALL Chat_Appand_pop(3,"<span onclick=c_sc(this.innerHTML) style=cursor: pointer class=c_name>" & GBL_CHK_User & "</span>��" & DEF_PointsName(4) & "����" & Fix(LastDoingTime/60) & "��!")
End If
REM *******Chat End*********
			End If
			
			Application.Lock
			If isNumeric(Application(DEF_MasterCookies & "SiteOlTime")) = False Then Application(DEF_MasterCookies & "SiteOlTime") = 0
			Application(DEF_MasterCookies & "SiteOlTime") = Application(DEF_MasterCookies & "SiteOlTime") + LastDoingTime
			Application.UnLock
			If Application(DEF_MasterCookies & "SiteOlTime") > 53873 Then '1Сʱ�ౣ��һ��
				Application.Lock
				NewIP = Application(DEF_MasterCookies & "SiteOlTime")
				Application(DEF_MasterCookies & "SiteOlTime") = 0
				Application.UnLock
				Call LDExeCute("Update LeadBBS_SiteInfo Set OnlineTime=OnlineTime+" & NewIP,1)
				UpdateStatisticDataInfo NewIP,0,1
			End If
		End If
		If Count > 1 Then
			UpdateOnlineUserInfo("from LeadBBS_onlineUser where ID=" & OL2)
			Application.Lock
			Application(DEF_MasterCookies & "ActiveUsers") = cCur("0" & Application(DEF_MasterCookies & "ActiveUsers")) - 1
			Application.UnLock
		Else
			If SQL <> GBL_UserID Then
				Call LDExeCute("Update LeadBBS_onlineUser set UserID=" & GBL_UserID & ",HiddenFlag=" & tmp & ",UserName='" & Replace(i,"'","''") & "',LastDoingTime=" & GetTimeValue(DEF_Now) & ",IP='" & GBL_IPAddress & "' where ID=" & NIP,1)
				UpdateSessionValue 17,GBL_IPAddress,0
				UpdateSessionValue 18,GetTimeValue(DEF_Now),0
			ElseIf cCur(Session.SessionID) <> cCur(TmpSessionID) And GBL_UserID> 0 Then
				Call LDExeCute("Update LeadBBS_onlineUser set SessionID=" & cCur(Session.SessionID) & ",HiddenFlag=" & tmp & ",LastDoingTime=" & GetTimeValue(DEF_Now) & ",IP='" & GBL_IPAddress & "' where ID=" & NIP,1)
				UpdateSessionValue 17,GBL_IPAddress,0
				UpdateSessionValue 18,GetTimeValue(DEF_Now),0
			End If
		End If
	Else
		Rs.Close
		Set Rs = Nothing
		UpdateOnlineUserInfo("from LeadBBS_onlineUser where SessionID=" & cCur(Session.SessionID))
		Call LDExeCute("insert into LeadBBS_onlineUser(SessionID,UserID,LastDoingTime,IP,StartTime,AtBoardID,AtUrl,AtInfo,Browser,System,UserName,HiddenFlag,LastRndNumber) values(" & cCur(Session.SessionID) & "," & cCur(GBL_UserID) & "," & GetTimeValue(DEF_Now) & ",'" & GBL_IPAddress & "'," & GetTimeValue(DEF_Now) & "," & GBL_Board_ID & ",'" & Replace(Left(Request.Servervariables("SCRIPT_NAME") & "?" & Request.QueryString,255),"'","''") & "','����ҳ��','" & Left(Replace(GetSBInfo(1),"'","''"),30) & "','" & Left(Replace(GetSBInfo(2),"'","''"),30) & "','" & Replace(i,"'","''") & "'," & cCur(tmp) & "," & (Fix(Timer*1000) mod 9999) & ")",1)

		If GBL_CHK_User <> "" and GBL_UserID > 0 and CheckSupervisorUserName = 1 Then
			Call LDExeCute("insert into LeadBBS_Log(LogType,LogTime,LogInfo,UserName,IP,BoardID) Values(0," & GetTimeValue(DEF_Now) & ",'����Ա��¼��̳.','" & Replace(Replace(htmlencode(Left(GBL_CHK_User,14)),"\","\\"),"'","''") & "','" & Replace(GBL_IPAddress,"'","''") & "'," & GBL_Board_ID & ")",1)
		End If

		Application.Lock
		Application(DEF_MasterCookies & "ActiveUsers") = Application(DEF_MasterCookies & "ActiveUsers") + 1
		Application.UnLock
	End If
	GBL_CHK_TempStr = ""

End Sub

Sub GetStyleInfo

	Dim Temp
	If GBL_SiteHeadString = "" and GBL_SiteBottomString = "" Then
		'��Զ�����治ͬ������� Temp = Left(Request.Cookies(DEF_MasterCookies & "style")("border" & GBL_Board_ID),14)
		Temp = Left(Request.Cookies(DEF_MasterCookies & "style")("border"),14)
		If isNumeric(Temp) = 0 or Temp="" Then Temp = -1
		Temp = cCur(Temp)
		If (Temp >= 0 and Temp <= DEF_BoardStyleStringNum) or Temp >= 1000 Then GBL_Board_BoardStyle = Temp
	End If

	If GBL_Board_BoardStyle > DEF_BoardStyleStringNum and GBL_Board_BoardStyle < 1000 Then GBL_Board_BoardStyle = DEF_BoardStyleStringNum
	GBL_Board_BoardStyle = cCur(GBL_Board_BoardStyle)
	IF DEF_mustDefaultStyle >= 0 Then GBL_Board_BoardStyle = DEF_mustDefaultStyle

	Temp = Application(DEF_MasterCookies & "Style" & GBL_Board_BoardStyle)
	If isArray(Temp) = False Then
		If Temp & "" <> "yes" Then
			If GBL_ConFlag = 1 Then ReloadBoardStyleInfo(GBL_Board_BoardStyle)
			Temp = Application(DEF_MasterCookies & "Style" & GBL_Board_BoardStyle)
		End If
	End If
	If isArray(Temp) = True Then
		DEF_BBS_ScreenWidth = RTrim(Temp(1,0))
		DEF_BBS_DisplayTopicLength = Temp(2,0)
		GBL_DefineImage = ccur(Temp(3,0))
		If GBL_DefineImage = 1 Then
			GBL_DefineImage = "Skin/" & GBL_Board_BoardStyle & "/"
		Else
			GBL_DefineImage = ""
		End If
		If GBL_SiteHeadString = "" Then GBL_SiteHeadString = Temp(4,0)
		If GBL_SiteBottomString = "" Then GBL_SiteBottomString = Temp(5,0)
		GBL_TableHeadString = Temp(6,0)
		GBL_TableBottomString = Temp(7,0)
		GBL_ShowBottomSure = Temp(8,0)
		GBL_TempletID = cCur("0" & Temp(9,0))
		GBL_TempletFlag = cCur("0" & Temp(10,0))
	End If

End Sub

Sub SiteHead(headString)

	Dim Temp
	GetStyleInfo
	%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
	"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"  xml:lang="zh-CN" lang="zh-CN">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
	<meta name="description" content="<%=htmlencode(DEF_GBL_Description)%>" />
	<title>
		<%=Replace(headString,"<","&lt;")%>
	</title>
	<link rel="stylesheet" id="css" type="text/css" href="<%=DEF_BBS_homeUrl%>inc/<%
	If GBL_Board_BoardStyle < 1000 Then
		Response.Write "style" & GBL_Board_BoardStyle
	Else
		Response.Write "css/"
		If GBL_Board_BoardStyle < 10000 Then Response.Write "0"
		Response.Write GBL_Board_BoardStyle
	End If
	%>.css<%=DEF_Jer%>" title="cssfile" />
	<script type="text/javascript">
	<!--
	var DEF_MasterCookies = "<%=htmlencode(DEF_MasterCookies)%>";
	var GBL_Style = "<%=GBL_Board_BoardStyle%>";
	-->
	</script>
	<script src="<%=DEF_BBS_HomeUrl%>inc/js/jquery.js<%=DEF_Jer%>" type="text/javascript"></script>
	<script src="<%=DEF_BBS_HomeUrl%>inc/js/common.js<%=DEF_Jer%>" type="text/javascript"></script>
	<%=GBL_HeadResource%>
</head>
<%If Left(headString,3) = "   " Then
	If Left(headString,5) <> "     " Then Response.Write "<body id=""body"" onkeydown=""if(window.event.keyCode==27)return(false);"">"
Else%>
<body id="body">
<iframe name="hidden_frame" id="hidden_frame" style="display:none"></iframe>
<a name="top"></a><%
Global_SiteHead%>
<div class="head_top_out">
<div class="area">
<div class="head_top">
<div class="head_top_loginform" id="head_top_loginform">
<div>
<%If GBL_CHK_User = "" Then%>
		<ul class="list_line">
		<li><a href="<%=DEF_BBS_HomeUrl%>User/<%=DEF_RegisterFile%>">ע��</a></li>
		<li><a href="<%=DEF_BBS_HomeUrl%>User/Login.asp" onclick="return(pub_command('��¼',this,'anc_delbody','&dir=<%=DEF_BBS_HomeUrl%>'));">��¼</a></li><%
		If GetBinarybit(DEF_Sideparameter,10) = 1 Then%>
		<li>������<a href="<%=DEF_BBS_HomeUrl%>app/qqlogin/login.asp"><img src="<%=DEF_BBS_HomeUrl%>images/app/1.gif" border="0" style="position:absolute;" /><span style="padding-left:18px;">QQ��¼</span></a></li><%
		End If%>
		</ul><%
Else
	If GBL_CHK_User <> "" Then
		Response.Write "<span class='head_hellowords'>"
		Select Case Hour(DEF_Now)
		Case 0,1:Response.Write "��ҹ"
		Case 2,3,4:Response.Write "��ҹ"
		Case 5,6,7:Response.Write "����"
		Case 8,9,10:Response.Write "����"
		Case 11,12:Response.Write "����"
		Case 13,14,15,16,17,18:Response.Write "����"
		Case 19,20:Response.Write "�ƻ�"
		Case 21,22,23:Response.Write "����"
		End Select
	%>�ã�
	</span>
	<%	If GBL_CHK_Pass = "" Then%>
			<%=htmlEncode(GBL_CHK_User)%> <a href="<%=DEF_BBS_HomeUrl%>User/<%=DEF_RegisterFile%>?action=bind" style="position:relative;" title="����Ҫ�󶨻������ʺ���Ϣ."><img src="<%=DEF_BBS_HomeUrl%>images/app/<%=GBL_AppType%>.gif" border="0" style="position:absolute;" /><span style="padding-left:18px;">����/���ʺ�</span></a>
	<%	Else
	%>
	<span class="head_hellouser">
		<a href="<%=DEF_BBS_HomeUrl%>User/LookUserInfo.asp"><%=htmlEncode(GBL_CHK_User)%></a>
	</span><%
			If DEF_EnableUserHidden = 1 Then
				%><span class="layerico"><a href="<%=DEF_BBS_HomeUrl%>User/Login.asp?action=hidden" onclick="return(pub_msg(this,'layer_ajaxmsg','&sure=1','setTimeout(\'document.location.reload();\',2000);'));" class="head_hidden"><%
				If GBL_CHK_ShowFlag = 0 Then
					%>����<%
				Else
					%>����<%
				End If
				%></a></span><%
			End If
			%> <a href="<%=DEF_BBS_HomeUrl%>User/Login.asp?R=Yes&dir=<%=DEF_BBS_HomeUrl%>" onclick="return(pub_command('���µ�¼',this,'anc_delbody','&R=Yes'));" class="head_relogin">�ص�¼</a>
			<%
			DisplayInfoBoxNavigate
		End If
		If GBL_CHK_Flag = 1 or (GBL_CHK_User <> "" and GBL_AppType <> "") Then
			%><a href="<%=DEF_BBS_HomeUrl%>User/login.asp?action=logout" onclick="return(pub_msg(this,'layer_ajaxmsg','&sure=1','setTimeout(\'document.location.reload();\',1000);'));" class="head_logout">�˳�</a><%
		End If
	End If
End If%></div>
</div>
<div class="head_top_string">
<script type="text/javascript">
<!--
function fAddFavorite(Tl, UL)
{
if(document.all)
    window.external.AddFavorite(UL, Tl);
else
    window.sidebar.addPanel(Tl, UL, "");
}
-->
</script>
	<ul class="list_line">
		<li><a href="<%=DEF_SiteHomeUrl%>" target="_top">��վ��ҳ</a></li>
		<li><a href="#33" onclick="fAddFavorite('<%=DEF_SiteNameString%>','http://<%=Request.Servervariables("SERVER_NAME") & Request.Servervariables("SCRIPT_NAME") & "?" & htmlencode(Replace(Request.QueryString,"&","&amp;"))%>')">�����ղ�</a></li>
		<li><a href="<%=DEF_BBS_HomeUrl%>User/help/help.asp">ʹ�ð���</a></li>
		<li><a href="<%=DEF_BBS_HomeUrl%>User/help/about.asp">��ϵ����</a></li>
	</ul>
	</div>
</div>
<%If GBL_SiteHeadString = "" Then%>
<div class="head_banner">
<span class="head_banner_logo"><a href="<%=DEF_BBS_HomeUrl%>"><img src="<%=DEF_BBS_HomeUrl%>images/blank.GIF" alt="������̳��ҳ" /></a></span>
<span class="head_banner_ad"><%=DEF_TopAdString%></span>
</div>
<%End If%>
<div class="head_sty">
<%'Global_SmallTableHead%>
<%
If DEF_BoardStyleStringNum > 0 Then%>
<script type="text/JavaScript">
<!--
function selsty(n,ty)
{
	if(ty==1)
	{
		$id("skin_select").style.display = "none";
		setStyle("<%=DEF_BBS_HomeUrl%>inc/style" + n + ".css","n")
		return;
	}
	var url = "<%=DEF_BBS_HomeUrl%>User/BoardStyle.asp?b=<%=GBL_Board_ID%>&AjaxFlag=1&s=" + n,u="";
	if((parent.document.location + "").toLowerCase().indexOf("frame.asp") != -1)
	{	u = "parent.document.location = \"<%
		Response.Write DEF_BBS_HomeUrl & "Frame.asp?u=http://" & Request.ServerVariables("server_name")
		If Request.ServerVariables("SERVER_PORT") <> "80" Then Response.Write Server.UrlEncode(":" & Request.ServerVariables("SERVER_PORT"))
		Randomize
		Response.Write Server.UrlEncode(Request.Servervariables("SCRIPT_NAME") & "?" & Left(Request.QueryString,200))%>&rnd=<%=Fix(Rnd*1314)
		%>\";";
		getAJAX(url,'SureFlag=E72ksiOkw2',u,1);
	}
	else
	{	u = "document.location.reload();";
		getAJAX(url,'SureFlag=E72ksiOkw2',u,1);
	}
}
-->
</script>
<%End If%>
<div class="munu_nav2 fire">
<%
If GBL_SiteHeadString = "" Then Response.Write ""
%>
<%If GBL_CHK_User <> "" Then%>
<div class="menu_nav">
	<div class="layer_item2">
		<div class="title"><a href="<%=DEF_BBS_HomeUrl%>User/MyInfoBox.asp"><span class="layer_item_title">����ר��</span></a></div>
		<div class="layer_iteminfo2">
			<ul class="menu_list">
			<li><a href="<%=DEF_BBS_HomeUrl%>User/LookUserInfo.asp">������Ϣ</a></li>
			<li><a href="<%=DEF_BBS_HomeUrl%>User/UserModify.asp">�޸��ҵ�����</a></li>
			<li><a href="<%=DEF_BBS_HomeUrl%>User/SendMessage.asp" onclick="return(sendprivatemsg(this,'<%=DEF_BBS_HomeUrl%>'));">���Ͷ���Ϣ</a></li>
			<li><a href="<%=DEF_BBS_HomeUrl%>User/MyInfoBox.asp">�ռ���</a></li>
			<li><a href="<%=DEF_BBS_HomeUrl%>User/LookUserInfo.asp?Evol=bag">�����ղؼ�</a></li>
			<li><a href="<%=DEF_BBS_HomeUrl%>User/LookUserInfo.asp?Evol=g">�ҷ��������</a></li>
			<li><a href="<%=DEF_BBS_HomeUrl%>User/alipay/Payment.asp"><%=DEF_PointsName(1)%>��ֵ</a></li>
			</ul>
		</div>
	</div>
</div>
<%End If%>

<div class="menu_nav">
	<div class="layer_item2">
		<div class="title"><a href="<%=DEF_BBS_HomeUrl%>User/UserTop.asp"><span class="layer_item_title">��̳����</span></a></div>
		<div class="layer_iteminfo2">
			<ul class="menu_list">
			<li><a href="<%=DEF_BBS_HomeUrl%>User/UserTop.asp?S">�������а�</a></li>
			<li><a href="<%=DEF_BBS_HomeUrl%>Search/List.asp?1">��������</a></li>
			<li><a href="<%=DEF_BBS_HomeUrl%>Search/List.asp?2">������</a></li>
			<li><a href="<%=DEF_BBS_HomeUrl%>User/Help/Help.asp">��������</a></li>
			<li><a href="<%=DEF_BBS_HomeUrl%>b/B.asp?E=1">��ר����</a></li>
			<li><a href="<%=DEF_BBS_HomeUrl%>Search/UploadList.asp">�鿴��̳����</a></li>
			<li><a href="<%=DEF_BBS_HomeUrl%>User/Help/About.asp">��̳�����Ŷ�</a></li>
			</ul>
		</div>
	</div>
</div>


<div class="menu_nav">
	<div class="layer_item2">
		<div class="title"><a href="<%=DEF_BBS_HomeUrl%>Plug-Ins/ChineseCode/"><span class="layer_item_title">���</span></a></div>
		<div class="layer_iteminfo2">
			<div class="layer_iteminfo2_box">
			<ul class="menu_list">
			<li><a href="<%=DEF_BBS_HomeUrl%>plug-ins/ChineseCode/">��̳���</a></li>
			<li><a href="<%=DEF_BBS_HomeUrl%>plug-ins/LeadCard/">LeadCard</a></li>
			</ul>
			</div>
		</div>
	</div>
</div>

<%If DEF_BoardStyleStringNum > 0 and DEF_mustDefaultStyle < 0 Then%>
<div class="menu_nav">
	<div class="layer_item2">
<div class="title"><a href="<%=DEF_BBS_HomeUrl%>User/BoardStyle.asp?b=<%=GBL_Board_ID%>"><span class="layer_item_title">ѡ����</span></a></div>
<div class="layer_iteminfo2" id="skin_select">
			<ul class="menu_list"><%

For Temp = 0 to DEF_BoardStyleStringNum
	If GBL_Board_BoardStyle = Temp Then
		Response.Write "<li><a href=""javascript:selsty(" & Temp & ")"">" & DEF_BoardStyleString(Temp) & "��</a></li>"
	Else
		Response.Write "<li><a href=""javascript:selsty(" & Temp & ")"">" & DEF_BoardStyleString(Temp) & "</a></li>"
	End If
Next%>
			<li><a href="<%=DEF_BBS_HomeUrl%>User/BoardStyle.asp?b=<%=GBL_Board_ID%>&dir=<%=DEF_BBS_HomeUrl%>&action=extended" onclick="return(pub_command('ѡ����',this,'anc_delbody',''));">����...</a></li>
			</ul>
		</div>
	</div>
</div>
<%End If%>
	<%
				Response.Write "<div class=""layer_item3"">"

				If CheckSupervisorUserName = 1 Then
					Response.Write "<a href=""" & DEF_BBS_HomeUrl & DEF_ManageDir & "/default.asp"" class=""head_manage""><span class=""head_item_title"">����</span></a>"
				ElseIf GetBinarybit(GBL_CHK_UserLimit,10) = 1 Then
					Response.Write "<a href=""" & DEF_BBS_HomeUrl & "User/BoardMaster/default.asp"" class=""head_manage""><span class=""head_item_title"">����</span></a>"
				End If

				Response.Write "<a href=""" & DEF_BBS_HomeUrl & "Search/Search.asp"" class=""head_search""><span class=""head_item_title"">����</span></a>"%>
			</div>
				<a href="#bottom" title="ת��ҳ��ײ�" class="unsel" hidefocus="true"><img src="<%=DEF_BBS_HomeUrl%>images/blank.gif" class="head_gobottom" alt="ת��ҳ��ײ�" /></a>
				<a href="<%=DEF_BBS_HomeUrl%>Other/RSS.asp<%
				If GBL_Board_ID > 0 Then Response.Write "?b=" & GBL_Board_ID
				%>" title="RSS���Ĵ˰�����Ϣ" class="unsel" hidefocus="true"><img src="<%=DEF_BBS_HomeUrl%>images/blank.gif" class="head_rss" alt="rss" /></a>
				<a href="<%=DEF_BBS_HomeUrl%>MINI/default.asp<%
				If GBL_Board_ID > 0 Then
					Response.Write "?" & GBL_Board_ID
					If Request.QueryString("ID") <> "" Then
						Response.Write "-" & filterUrlstr(Request.QueryString("ID")) & "-0-0-0-0-0-a-.htm"
					Else
						Response.Write "-0-0-0-0-0-0-b-.htm"
					End If
				End If%>" title="���㷽ʽ��ʾ��̳" class="unsel" hidefocus="true"><img src="<%=DEF_BBS_HomeUrl%>images/blank.gif" class="head_mini" alt="mini" /></a>
				<%
				If GBL_TableHeadString = "" Then Response.Write ""
				%>
		</div>
	</div>
</div>
	<%
	End If%>
</div>
	<%

End Sub

Sub PageExeCuteInfo
%>

				<div class="version">
					Powered by <a href="http://www.leadbbs.com/" target="_blank"><b><%=DEF_Version%></b></a>
					<a href="http://www.leadbbs.com/other/register?"></a>.
				</div>
				<div class="createtime" id="createtime">
				<%
		Response.Write " Page created in " & FormatNumber(cCur(Timer - DEF_PageExeTime1),4,True) & " seconds with " & GBL_DBNum & " queries."
		if OPEN_DEBUG = 1 and DEBUG_User = GBL_CHK_User Then %>
		<div style="text-align:left;background:white;color:black;"><ul><%=sqlstring%></ul></div>
		<%End If%>
				</div>
<%
End Sub

Sub SiteBottom

	%>
	</div>
</div>
			<a name="bottom"></a>
			<div class="bottominfo">
				<div class="area">
				<div class="copyright">
					Copyright <span style="font:11px Tahoma,Arial,sans-serif;">&copy;</span>2003-<%=year(DEF_Now)%>&nbsp;<%=DEF_SiteNameString%>
					- <a href="javascript:;" onclick="LD.Cookie.Clear();">���COOKIE</a><%=DEF_BottomInfo%>
				</div>
				<%PageExeCuteInfo%>
				</div>
			</div>
	<script type="text/javascript">
	<!--
		new LayerMenu('layer_item','layer_iteminfo');
		new LayerMenu('layer_item2','layer_iteminfo2');
		layer_initselect();
		
		var alls = document.getElementsByTagName('form'); 
		for(var i=0; i<alls.length; i++)
		{
			submit_disable(alls[i],1);
		}
		if (typeof initLightbox == 'function')initLightbox();
	-->
	</script>
	<%Global_SiteBottom%>
	<div class="bottom_ad">
	<div class="area">
	<div id="bottom_ad">
	<%Response.Write "<!--"%>
	<!-- #include file=incHtm/Bottom_AD.asp -->
	<%Response.Write "-->"%>
	</div>
	</div>
	</div>
	
<script src=<%=DEF_BBS_HomeUrl%>inc/js/writecapture/writeCapture.js></script>
<script src=<%=DEF_BBS_HomeUrl%>inc/js/writecapture/jquery.writeCapture.js></script>
	<script src="<%=DEF_BBS_HomeUrl%>inc/js/ad.js?123" type="text/javascript">
	</script>
	</body>
	</html><%

End Sub

Sub SiteBottom_Spend

	%>
	<div class="createtime">
	<%
		Response.Write " Page created in " & FormatNumber(cCur(Timer - DEF_PageExeTime1),4,True) & " seconds width " & GBL_DBNum & " queries."
	%>
	</div>
	</body>
	</html>
	<%

End Sub

Sub Global_SiteHead

	If GBL_SiteHeadString <> "" Then Response.Write GBL_SiteHeadString

End Sub

Sub Global_SiteBottom

	If GBL_SiteBottomString <> "" Then Response.Write GBL_SiteBottomString

End Sub

Sub Global_TableHead

	Response.Write "<div class=""stylebox"">"
	If GBL_TableHeadString <> "" Then Response.Write GBL_TableHeadString
	Response.Write "<div class=""nonestylebox"">"

End Sub


Sub Global_TableBottom

	Response.Write "</div>"
	If GBL_TableBottomString <> "" Then Response.Write GBL_TableBottomString
	Response.Write "</div>"

End Sub

Sub Global_ErrMsg(Str)

	If Trim(Str) = "" or Trim(Str) = VbCrLf Then Exit Sub
	%>
<div class="alertbox">
		<b>��ʾ��Ϣ</b>��<br /><br />
		<span class="redfont">
			<%=Str%>
		</span>
		<br />
		[<a href="javascript:history.back()"><b>�����ϴ�ҳ��</b></a>]
</div>
	<%

End Sub

Sub ErrorJump(str)

	CloseDatabase
	Response.Redirect DEF_BBS_HomeUrl & "User/Login.asp?action=err&err=" & urlencode(left(str,500))

End Sub

GetUserNamePassword

Sub GetUserNamePassword

	GBL_CookiePassFlag = 0
	GBL_CHK_Flag = 1
	If dontRequestFormFlag = "" Then
		GBL_CHK_User = Trim(Request.Form("user"))
		GBL_CHK_Pass = Left(Request.Form("pass"),15)
	End If
	'If GBL_CHK_User = "" Then GBL_CHK_User = Trim(Request.QueryString("user"))
	If GBL_CHK_User = "" Then GBL_CHK_User = DecodeCookie(Left(Request.Cookies(DEF_MasterCookies)("User"),255))
	GBL_AppType = htmlencode(Left(Request.Cookies(DEF_MasterCookies & "_apptype"),12))

	GBL_CHK_User = Replace(Left(GBL_CHK_User,20),",","")
	If isArray(GBL_UDT) Then
		If GBL_CHK_User = "" or (GBL_CHK_User <> "" and LCase(GBL_CHK_User) = LCase(GBL_UDT(1))) Then
			GBL_UserID = cCur(GBL_UDT(0))
			GBL_CHK_User = GBL_UDT(1)
			GBL_CHK_UserLimit = GBL_UDT(2)
			GBL_CHK_ShowFlag = ccur(GBL_UDT(3))
			GBL_CHK_MessageFlag = ccur(GBL_UDT(6))
			Dim CkiExp
			If dontRequestFormFlag="" Then
				CkiExp = Request.Form("CkiExp")
			End If
			If CkiExp & "" = "" and GBL_CHK_Pass = "" Then GBL_CHK_Pass = GBL_UDT(9) & ""
	
			Exit Sub
		Else
REM *******Chat Start*******
Chat_SessionFree(GBL_UDT(1))
REM *******Chat End*********
			Set GBL_UDT = Nothing
		End If
	End If
	If GBL_CHK_User = "" Then GBL_CHK_Flag = 0
	GBL_CHK_ShowFlag = 0
	GBL_CHK_UserLimit = 0
	If GBL_CHK_Pass = "" and GBL_CHK_Flag = 1 Then
		GBL_CHK_Pass = Left(DecodeCookie(Left(Request.Cookies(DEF_MasterCookies)("pass"),255)),32)
		If Len(GBL_CHK_Pass) = 32 Then GBL_CookiePassFlag = 1
	End If
	If GBL_CHK_Pass = "" Then GBL_CHK_Flag = 0

End Sub

Sub Free_UDT

	Set GBL_UDT = Nothing
	Set Session(DEF_MasterCookies & "UDT") = Nothing

End Sub

Function CheckPass

	Dim IPADDRESS,SubmitFlag,MD5Pass
	If dontRequestFormFlag="" Then
		SubmitFlag = Request.Form("submitflag")
		If SubmitFlag = "" Then SubmitFlag = request.QueryString("submitflag")
	Else
		SubmitFlag = Request.QueryString("submitflag")
	End If
	If dontRequestFormFlag = "AppLogin" Then SubmitFlag = "AppLogin"
	If GBL_CheckPassDoneFlag = 1 Then
		CheckPass = 0
		Exit Function
	End If
	GBL_CheckPassDoneFlag = 1
	GBL_UserID = 0

	IPADDRESS = GBL_IPAddress
	If GBL_IPAddress = "1.1.1.1" Then
		'GBL_CHK_TempStr = "�Ƿ�IP��ַ���޷���ȫ�ύ�κ����ݡ�" & VbCrLf
		'GBL_CHK_Flag = 0
		'CheckPass = 0
		'Exit Function
	End If
		
	Dim UpdateString

	If GBL_CHK_Flag <> 1 Then
		If GBL_CHK_User <> "" Then Pub_ClearCookie
		If GBL_AppType = "" Then GBL_CHK_User = ""
		GBL_UserID = 0
		CheckPass = 0
		Exit Function
	End If

	Dim Rs
	Dim SQL
	If isArray(GBL_UDT) and dontRequestFormFlag <> "AppLogin" Then
		If cCur(GBL_UDT(11)) > 19800000000000 Then
			SQL = DateDiff("s",RestoreTime(GBL_UDT(11)), DEF_Now)
			If SQL < 0 Then SQL = 0
		Else
			SQL = 0
		End If
		If LCase(GBL_UDT(1)) <> LCase(GBL_CHK_User) Then SQL = 240
	Else
		SQL = 240
	End If
REM *******Chat Start*******
Dim Chat_CreateFlag
Chat_CreateFlag = 0
REM *******Chat End*********
	If SQL >= 240 Then
		If DEF_RepeatLoginTimeOut > 0 and DEF_RepeatLoginTimeOut < DEF_UserOnlineTimeOut Then
			SQL = sql_select("Select T.ID,T.UserName,T.UserLimit,T.ShowFlag,T.Points,T.OnlineTime,T.MessageFlag,T.Login_lastpass,T.Login_falsenum,T.Pass,T.Login_RightIP,T.Prevtime,T.LockIP,T.LastWriteTime,T.LastAnnounceID,T.CharmPoint,T.CachetValue,S.IP,S.LastDoingTime,','+T.Officer+',' from LeadBBS_User as T left join LeadBBS_OnlineUser as S on T.ID=S.UserID Where T.Username = '" & Replace(GBL_CHK_User,"'","''") & "'",1)
		Else
			SQL = sql_select("Select ID,UserName,UserLimit,ShowFlag,Points,OnlineTime,MessageFlag,Login_lastpass,Login_falsenum,Pass,Login_RightIP,Prevtime,LockIP,LastWriteTime,LastAnnounceID,CharmPoint,CachetValue,'',0,','+Officer+',' from LeadBBS_User Where Username = '" & Replace(GBL_CHK_User,"'","''") & "'",1)
		End If
		Set Rs = LDExeCute(SQL,0)
		If ((Rs.Eof) or (Rs.Bof)) Then
			Rs.Close
			Set Rs = Nothing
REM *******Chat Start*******
Chat_SessionFree(GBL_CHK_User)
REM *******Chat End*********
			'GBL_CHK_TempStr = "��������û������ڣ� ��¼ʧ��!" & VbCrLf
			If GBL_CHK_User <> "" Then Pub_ClearCookie
			If GBL_AppType = "" Then GBL_CHK_User = ""
			GBL_CHK_Flag = 0
			GBL_UserID = 0
			Set Session(DEF_MasterCookies & "UDT") = Nothing
			Session(DEF_MasterCookies & "UDT") = ""
			CheckPass = 0
			Exit Function
		Else
REM *******Chat Start*******
			Dim Old_Name
			Old_Name = ""
			If isArray(GBL_UDT) Then Old_Name = GBL_UDT(1)
REM *******Chat End*********
			ReDim GBL_UDT(19)
			GBL_UDT(0) = Rs(0)
			GBL_UDT(1) = Rs(1)
			GBL_UDT(2) = Rs(2)
			GBL_UDT(3) = Rs(3)
			GBL_UDT(4) = Rs(4)
			GBL_UDT(5) = Rs(5)
			GBL_UDT(6) = Rs(6)
			GBL_UDT(7) = Rs(7)

			GBL_UDT(8) = Rs(8)
			GBL_UDT(9) = Rs(9)
			GBL_UDT(10) = Rs(10)
			GBL_UDT(11) = Rs(11)
			GBL_UDT(12) = Rs(12)
			GBL_UDT(13) = Rs(13)
			GBL_UDT(14) = Rs(14)
			GBL_UDT(15) = Rs(15)
			GBL_UDT(16) = Rs(16)
			If DEF_RepeatLoginTimeOut > 0 and DEF_RepeatLoginTimeOut < DEF_UserOnlineTimeOut Then
				GBL_UDT(17) = Rs(17)
				GBL_UDT(18) = Rs(18)
			End If
			GBL_UDT(19) = Rs(19)
			Rs.Close
			Set Rs = Nothing
			Session(DEF_MasterCookies & "UDT") = GBL_UDT
REM *******Chat Start*******
If Old_Name <> "" and Old_Name <> GBL_UDT(1) Then
	Chat_SessionFree(Old_Name) '���µ�¼����ǰ�ͷž��û�
End If
Chat_CreateFlag = 1
REM *******Chat End*********
		End If
	End If

	If DEF_RepeatLoginTimeOut > 0 and DEF_RepeatLoginTimeOut < DEF_UserOnlineTimeOut Then
		If GBL_IPAddress <> GBL_UDT(17) and DateDiff("s",RestoreTime(GBL_UDT(18)),DEF_Now) < DEF_RepeatLoginTimeOut Then
REM *******Chat Start*******
Chat_SessionFree(GBL_CHK_User)
REM *******Chat End*********
			If SubmitFlag <> "" Then GBL_CHK_TempStr = "���û��Ѿ����ߣ�ϵͳ�Ѿ����ó�" & Fix(DEF_RepeatLoginTimeOut/60) & "���Ӻ�������ٴε�¼!" & VbCrLf
			If GBL_CHK_User <> "" Then Pub_ClearCookie
			If GBL_AppType = "" Then GBL_CHK_User = ""
			GBL_CHK_Flag = 0
			Set Session(DEF_MasterCookies & "UDT") = Nothing
			Session(DEF_MasterCookies & "UDT") = ""
			CheckPass = 0
			Exit Function
		End If
	End If

	If GBL_UDT(12) & "" <> "" and GBL_UDT(12) <> GBL_IPAddress Then
REM *******Chat Start*******
Chat_SessionFree(GBL_CHK_User)
REM *******Chat End*********
		If SubmitFlag <> "" Then GBL_CHK_TempStr = "���û��ѱ����Σ�����Ȩʹ�ô��û�!" & VbCrLf
		If GBL_CHK_User <> "" Then Pub_ClearCookie
		If GBL_AppType = "" Then GBL_CHK_User = ""
		GBL_CHK_Flag = 0
		Set Session(DEF_MasterCookies & "UDT") = Nothing
		Session(DEF_MasterCookies & "UDT") = ""
		CheckPass = 0
		Exit Function
	End If
	SQL = 0
	Dim Login_lastpass,Login_falsenum,Pass,Login_RightIP,Prevtime
	GBL_UserID = ccur(GBL_UDT(0))
	GBL_CHK_User = GBL_UDT(1)
	GBL_CHK_UserLimit = ccur(GBL_UDT(2))
	GBL_CHK_ShowFlag = ccur(GBL_UDT(3))
	If GBL_CHK_ShowFlag = 1 Then
		GBL_CHK_ShowFlag = 1
	Else
		GBL_CHK_ShowFlag = 0
	End If
	GBL_CHK_Points = cCur(GBL_UDT(4))
	GBL_CHK_OnlineTime = cCur(GBL_UDT(5))
	GBL_CHK_MessageFlag = ccur(GBL_UDT(6))
	Login_lastpass = Trim(GBL_UDT(7))
	Login_falsenum = cCur(GBL_UDT(8))
	Prevtime = ReStoretime(cCur(GBL_UDT(11)))
	Pass = GBL_UDT(9)
	Login_RightIP = GBL_UDT(10)
	GBL_CHK_LastWriteTime = DateDiff("s",RestoreTime(GBL_UDT(13)),DEF_Now)
	GBL_CHK_LastAnnounceID = cCur(GBL_UDT(14))
	GBL_CHK_CharmPoint = cCur(GBL_UDT(15))
	GBL_CHK_CachetValue = cCur(GBL_UDT(16))

	If GBL_CHK_LastWriteTime < 0 Then GBL_CHK_LastWriteTime = DEF_WriteEventSpace + 1
	If GetTimeValue(PrevTime) > 19800000000000 Then
		PrevTime = DateDiff("s",Prevtime, DEF_Now)
		If PrevTime<0 Then PrevTime=0
	Else
		PrevTime = 0
	End If

	Dim PassCorrect,LastPassCorrect
	PassCorrect = 0
	If GBL_CookiePassFlag = 1 Then
		If MD5(GBL_CHK_User & Pass) = GBL_CHK_Pass Then
			PassCorrect = 1
			MD5Pass = Pass
		Else
			MD5Pass = ""
		End If
	Else
		If Len(GBL_CHK_Pass) > 15 Then
			MD5Pass = GBL_CHK_Pass
		Else
			MD5Pass = MD5(GBL_CHK_Pass)
		End If
		If (Pass = MD5Pass or Mid(MD5Pass,9,16) = Pass) Then PassCorrect = 1
	End If

	'������ȷ���Ѿ����ε�¼һ��������֤
	If PassCorrect=1 and Login_falsenum >= DEF_MaxLoginTimes and IPADDRESS <> Login_RightIP and (Prevtime < DEF_LoginSpaceTime and Prevtime >= 0) Then
		If SubmitFlag <> "" Then GBL_CHK_TempStr = "�˻����¼�����������" & DEF_MaxLoginTimes & "��,����ʱ����,�����ٴε�¼����" & (DEF_LoginSpaceTime-Prevtime) & "��." & VbCrLf
		UpdateString = ",Prevtime=" & GetTimeValue(DEF_Now)
		UpdateSessionValue 11,GetTimeValue(DEF_Now),0
		If GBL_CHK_LastWriteTime < DEF_WriteEventSpace Then
			Call LDExeCute("Update LeadBBS_User Set " & Mid(UpdateString,2) & ",LastWriteTime=" & GetTimeValue(DEF_Now) & " where ID=" & GBL_UserID,1)
			GBL_CHK_LastWriteTime = 240
			UpdateSessionValue 13,GetTimeValue(DEF_Now),0
		End If
REM *******Chat Start*******
Chat_SessionFree(GBL_CHK_User)
REM *******Chat End*********
		CheckPass = 0
		GBL_UserID = 0
		GBL_CHK_Flag = 0
		GBL_CHK_UserLimit = 0
		GBL_CHK_Points = 0
		GBL_CHK_OnlineTime = 0
		GBL_CHK_MessageFlag = 0
		If GBL_CHK_User <> "" Then Pub_ClearCookie
		If GBL_AppType = "" Then GBL_CHK_User = ""
		GBL_CHK_CharmPoint = 0
		GBL_CHK_CachetValue = 0
		Set Session(DEF_MasterCookies & "UDT") = Nothing
		Session(DEF_MasterCookies & "UDT") = "1"
		Exit function
	End If


	LastPassCorrect = 0
	If GBL_CookiePassFlag = 1 Then
		If MD5(GBL_CHK_User & Login_lastpass) = GBL_CHK_Pass Then
			LastPassCorrect = 1
			MD5Pass = Pass
		Else
			MD5Pass = ""
		End If
	Else
		If Len(GBL_CHK_Pass) > 15 Then
			MD5Pass = GBL_CHK_Pass
		Else
			MD5Pass = MD5(GBL_CHK_Pass)
		End If
		If (Login_lastpass = MD5Pass) Then LastPassCorrect = 1
	End If

	If LastPassCorrect = 0 Then  '�ж����һ������ 
		If(Login_falsenum<DEF_MaxLoginTimes) Then
			UpdateString = ",Prevtime=" & GetTimeValue(DEF_Now) & ",Login_IP='" & Replace(IPADDRESS,"'","''") & "',Login_falsenum=Login_falsenum+1,Login_lastpass='" & Replace(MD5Pass,"'","''") & "'"
			UpdateSessionValue 11,GetTimeValue(DEF_Now),0
			UpdateSessionValue 8,1,1
			UpdateSessionValue 7,MD5Pass,0
			If SubmitFlag <> "" Then GBL_CHK_TempStr = "��¼�������" & Login_falsenum+1 & "��,������" & (DEF_MaxLoginTimes-Login_falsenum-1) & "�ε�¼���Գ���." & VbCrLf
		Else
			If (Prevtime < DEF_LoginSpaceTime and Prevtime >= 0) Then
				If IPADDRESS = Login_RightIP Then
				Else
					If SubmitFlag <> "" Then GBL_CHK_TempStr = "�˻����¼�����������" & DEF_MaxLoginTimes & "��,����ʱ����,�����ٴε�¼����" & (DEF_LoginSpaceTime-Prevtime) & "��." & VbCrLf
					UpdateString = ",Prevtime=" & GetTimeValue(DEF_Now)
					UpdateSessionValue 11,GetTimeValue(DEF_Now),0
					If GBL_CHK_LastWriteTime < DEF_WriteEventSpace Then
						Call LDExeCute("Update LeadBBS_User Set " & Mid(UpdateString,2) & ",LastWriteTime=" & GetTimeValue(DEF_Now) & " where ID=" & GBL_UserID,1)
						GBL_CHK_LastWriteTime = 240
						UpdateSessionValue 13,GetTimeValue(DEF_Now),0
					End If
REM *******Chat Start*******
Chat_SessionFree(GBL_CHK_User)
REM *******Chat End*********
					CheckPass = 0
					GBL_UserID = 0
					GBL_CHK_Flag = 0
					GBL_CHK_UserLimit = 0
					GBL_CHK_Points = 0
					GBL_CHK_OnlineTime = 0
					GBL_CHK_MessageFlag = 0
					If GBL_CHK_User <> "" Then Pub_ClearCookie
					If GBL_AppType = "" Then GBL_CHK_User = ""
					GBL_CHK_CharmPoint = 0
					GBL_CHK_CachetValue = 0
					Set Session(DEF_MasterCookies & "UDT") = Nothing
					Session(DEF_MasterCookies & "UDT") = "1"
					Exit function
				End If
			Else
				UpdateString = UpdateString & ",Prevtime=" & GetTimeValue(DEF_Now) & ",Login_IP='" & Replace(IPADDRESS,"'","''") & "',Login_falsenum=1,Login_lastpass='" & Replace(MD5Pass,"'","''") & "'"
				UpdateSessionValue 11,GetTimeValue(DEF_Now),0
				UpdateSessionValue 13,GetTimeValue(DEF_Now),0
				UpdateSessionValue 7,MD5Pass,0
				UpdateSessionValue 8,1,1
			End If
		End If
	Else
		If PrevTime >= 240 Then
			If GBL_CHK_LastWriteTime > DEF_WriteEventSpace Then
				UpdateString = UpdateString & ",Prevtime=" & GetTimeValue(DEF_Now) & ",Login_IP='" & Replace(IPADDRESS,"'","''") & "'"
				UpdateSessionValue 11,GetTimeValue(DEF_Now),0
			End If
		End If
	End If

	If PassCorrect = 0 Then
		If cstr(IPADDRESS&chr(0)) = cstr(Login_RightIP&chr(0)) Then
			UpdateString = UpdateString & ",Login_RightIP='as'"
			UpdateSessionValue 10,"as",0
		End If
		If inStr(UpdateString,"Prevtime") = 0 then
			UpdateString = UpdateString & ",Prevtime=" & GetTimeValue(DEF_Now)
			UpdateSessionValue 11,GetTimeValue(DEF_Now),0
		End If
		Call LDExeCute("Update LeadBBS_User Set " & Mid(UpdateString,2) & " where ID=" & GBL_UserID,1)
		If GBL_CHK_TempStr = "" Then GBL_CHK_TempStr = "�����������, ��¼ʧ��! " & VbCrLf
REM *******Chat Start*******
Chat_SessionFree(GBL_CHK_User)
REM *******Chat End*********
		GBL_CHK_Flag = 0
		CheckPass = 0
		GBL_UserID = 0
		GBL_CHK_UserLimit = 0
		GBL_CHK_Points = 0
		GBL_CHK_OnlineTime = 0
		GBL_CHK_MessageFlag = 0
		If GBL_CHK_User <> "" Then Pub_ClearCookie
		If GBL_AppType = "" Then GBL_CHK_User = ""
		GBL_CHK_CharmPoint = 0
		GBL_CHK_CachetValue = 0
		Set Session(DEF_MasterCookies & "UDT") = Nothing
		Session(DEF_MasterCookies & "UDT") = ""
		Exit Function
	Else
		If Trim(SubmitFlag) <> "" Then
			If GBL_CHK_LastWriteTime > DEF_WriteEventSpace Then
				'������ȷ�򲻼�¼�ϴβ���ʱ��(�����¼,����ע�͵�һ��)
				'UpdateString = ",Prevtime=" & GetTimeValue(DEF_Now) & ",Login_IP='" & Replace(IPADDRESS,"'","''") & "',Login_RightIP='" & Replace(IPADDRESS,"'","''") & "',Login_falsenum=0,Login_oknum=Login_oknum+1"
				UpdateString = ",Login_IP='" & Replace(IPADDRESS,"'","''") & "',Login_RightIP='" & Replace(IPADDRESS,"'","''") & "',Login_falsenum=0,Login_oknum=Login_oknum+1"
				UpdateSessionValue 11,GetTimeValue(DEF_Now),0
				UpdateSessionValue 10,IPADDRESS,0
				UpdateSessionValue 8,0,0
			End If
		End If
		Rem ��¼�ɹ�
		If dontRequestFormFlag="" Then
			IPADDRESS = Request.Form("CkiExp")
			If IPADDRESS = "" Then IPADDRESS = Request.QueryString("CkiExp")
		Else
			If dontRequestFormFlag = "AppLogin" Then
				IPADDRESS = "365"
			Else
				IPADDRESS = ""
			End If
		End If
		If IPADDRESS <> "" Then
			If Len(GBL_CHK_Pass) > 32 Then GBL_CHK_TempStr = "<font color=red class=redfont>��¼ʧ�ܣ��������������!</font>" & VbCrLf
			Select Case IPADDRESS
				Case "-1": IPADDRESS = 0
				Case "365": IPADDRESS = 365
				Case "1": IPADDRESS = 1
				Case "2": IPADDRESS = 2
				Case "7": IPADDRESS = 7
				Case "31": IPADDRESS = 31
				Case "3650": IPADDRESS = 3650
				Case else: IPADDRESS = -99
			End Select
			If IPADDRESS <> -99 Then
				If IPADDRESS > 0 Then Response.Cookies(DEF_MasterCookies).Expires = DateAdd("d",DEF_Now,IPADDRESS)
				Response.Cookies(DEF_MasterCookies)("user") = CodeCookie(GBL_CHK_User)
				Response.Cookies(DEF_MasterCookies)("pass") = CodeCookie(MD5(GBL_CHK_User & Pass))
				'Response.Cookies(DEF_MasterCookies)("pass") = CodeCookie(GBL_CHK_Pass)
				Response.Cookies(DEF_MasterCookies)("expires") = GetTimeValue(DateAdd("d",DEF_Now,IPADDRESS))
				Response.Cookies(DEF_MasterCookies).Domain = DEF_AbsolutHome
			End If
			If GBL_CHK_ShowFlag = 1 and DEF_EnableUserHidden = 1 Then
				IPADDRESS = "�����û�"
				SQL = 0
			Else
				IPADDRESS = GBL_CHK_User
				SQL = cCur(GBL_CHK_UserLimit)
			End If
			If GBL_CHK_LastWriteTime < DEF_WriteEventSpace Then
				Call LDExeCute("Update LeadBBS_onlineUser set UserID=" & GBL_UserID & ",UserName='" & Replace(IPADDRESS,"'","''") & "',HiddenFlag=" & SQL & ",LastDoingTime=" & GetTimeValue(DEF_Now) & ",IP='" & GBL_IPAddress & "' where SessionID=" & cCur(Session.SessionID),1)
			End If
			'Set Session(DEF_MasterCookies & "UDT") = Nothing
			'Session(DEF_MasterCookies & "UDT") = "1"
		End If
	End If
	If UpdateString <> "" Then
		'�����������д��ʱ��,���м������ ,LastWriteTime=" & GetTimeValue(DEF_Now) & " 
		Call LDExeCute("Update LeadBBS_User Set " & Mid(UpdateString,2) & ",LastDoingTime=" & GetTimeValue(DEF_Now) & " where ID=" & GBL_UserID,1)
		GBL_CHK_LastWriteTime = 240
		'UpdateSessionValue 13,GetTimeValue(DEF_Now),0
	End If
	GBL_UserID = cCur(GBL_UserID)
	CheckPass = GBL_UserID
REM *******Chat Start*******
If Chat_CreateFlag >= 1 Then Chat_SessionCreate(GBL_UDT(1))
REM *******Chat End*********

End Function


Sub Pub_ClearCookie
	If GBL_AppType = "" Then Response.Cookies(DEF_MasterCookies)("User") = ""
	Response.Cookies(DEF_MasterCookies)("pass") = ""
	If GBL_AppType = "" Then Response.Cookies(DEF_MasterCookies)("expires") = ""
	If GBL_AppType = "" Then Response.Cookies(DEF_MasterCookies).Expires = Date - 1
	Response.Cookies(DEF_MasterCookies).Domain = DEF_AbsolutHome
End Sub

Dim GBL_BusyTimes
GBL_BusyTimes = 0

Function CheckWriteEventSpace

	Dim BusyTimes
	BusyTimes = Session(DEF_MasterCookies & "_BusyTimes")
	If isNumeric(BusyTimes) = 0 Then
		BusyTimes = 0
		Session(DEF_MasterCookies & "_BusyTimes") = 0
	End If
	BusyTimes = Fix(cCur(BusyTimes))
	If GBL_CheckPassDoneFlag = 0 Then
		If isArray(GBL_UDT) Then
			GBL_CHK_LastWriteTime = DateDiff("s",RestoreTime(GBL_UDT(13)),DEF_Now)
		Else
			GBL_UserID = CheckPass
		End If
	End If
	If GBL_CHK_LastWriteTime < DEF_WriteEventSpace and GBL_UserID <=0 Then
		BusyTimes = BusyTimes + 1
		If BusyTimes > 1 Then
			'GBL_CHK_TempStr = "���Ĳ�����Ƶ�����Ժ����ԣ�"
			CheckWriteEventSpace = 0
		Else
			CheckWriteEventSpace = 1
		End If
	Else
		CheckWriteEventSpace = 1
		BusyTimes = 0
	End If
	Session(DEF_MasterCookies & "_BusyTimes") = BusyTimes

End Function

Sub UpdateLastWriteTime

	'LDExeCute("Update LeadBBS_User Set Prevtime=" & GetTimeValue(DEF_Now) & ",LastWriteTime=" & GetTimeValue(DEF_Now) & " where ID=" & GBL_UserID,1)
	UpdateSessionValue 13,GetTimeValue(DEF_Now),0

End Sub

Sub UpdateSessionValue(N,Value,T)

	If isArray(Session(DEF_MasterCookies & "UDT")) = False Then Exit Sub
	If (DEF_RepeatLoginTimeOut <= 0 or DEF_RepeatLoginTimeOut > DEF_UserOnlineTimeOut) and N >=17 Then Exit Sub
	Dim TA
	TA = Session(DEF_MasterCookies & "UDT")
	If isArray(TA) = False Then Exit Sub
	If T = 0 Then
		TA(N) = Value
	Else
		TA(N) = cCur(TA(N)) + Value
	End If
	Session(DEF_MasterCookies & "UDT") = TA

End Sub

Function CheckSupervisorUserName

	If GBL_CHK_User <> "" and inStr(GBL_CHK_User,",") = 0 and inStr(LCase("," & DEF_SupervisorUserName & ","),"," & LCase(GBL_CHK_User) & ",") > 0 Then
		If Session(DEF_MasterCookies & "Manager") = "manage" Then
			CheckSupervisorUserName = 1
		Else
			CheckSupervisorUserName = 0
		End If
	Else
		CheckSupervisorUserName = 0
	End If

End Function

Function DisplayAnnounceTitle(str,Sty)

	If Sty >= 60 Then
		DisplayAnnounceTitle = "<span class=""grayfont"">���ӵȴ������...</span>"
		Exit Function
	End If
	Dim s
	s = str
	If Sty <> 1 Then s = htmlencode(str)
	If Sty >= 5 and Sty <= 8 Then s = "<strong>" & s & "</strong>"
	Select Case Sty
		Case 1: DisplayAnnounceTitle = "<span class=""word-break-all"">" & s & "</span>"
		Case 2,6: DisplayAnnounceTitle = "<span class=""word-break-all redfont"">" & s & "</span>"
		Case 3,7: DisplayAnnounceTitle = "<span class=""word-break-all greenfont"">" & s & "</span>"
		Case 4,8: DisplayAnnounceTitle = "<span class=""word-break-all bluefont"">" & s & "</span>"
		Case else: DisplayAnnounceTitle = "<span class=""word-break-all"">" & s & "</span>"
	End Select

End Function

Sub SetActiveUserCount

	Dim Rs,Temp
	Set Rs = LDExeCute("Select Count(*) from LeadBBS_OnlineUser",0)
	If Rs.Eof Then
		Application.Lock
		Application(DEF_MasterCookies & "ActiveUsers") = 0
		Application.UnLock
	Else
		Temp = Rs(0)
		If isNull(Temp) Then Temp = 0
		Application.Lock
		Application(DEF_MasterCookies & "ActiveUsers") = cCur(Temp)
		Application.UnLock
	End If
	Rs.Close
	Set Rs = Nothing

End Sub

Sub UpdateBoardApplicationInfo(BoardID,Value,N)

	If isArray(Application(DEF_MasterCookies & "BoardInfo" & BoardID)) = True Then
		Dim TmpArray
		TmpArray = Application(DEF_MasterCookies & "BoardInfo" & BoardID)
		TmpArray(N,0) = Value
		Application.Lock
		Application(DEF_MasterCookies & "BoardInfo" & BoardID) = TmpArray
		Application.UnLock
	End If

End Sub

Sub ReloadPubMessageInfo

	Dim Rs
	Set Rs = LDExeCute(sql_select("Select title,SendTime from LeadBBS_InfoBox where ToUser='' Order by ID DESC",DEF_TopicContentMaxListNum),0)
	If Not Rs.Eof Then
		Application.Lock
		Application(DEF_MasterCookies & "PubMsg") = Rs.GetRows(-1)
		Application.UnLock
	Else
		Application.Lock
		Application(DEF_MasterCookies & "PubMsg") = VbCrLf & VbCrLf
		Application.UnLock
	End If
	Rs.Close
	Set Rs = Nothing

End Sub

Sub ReloadBoardStyleInfo(ID)

	If GBL_ConFlag = 0 Then Exit Sub
	Dim Rs,Temp
	Set Rs = LDExeCute(sql_select("Select T1.StyleID,T1.ScreenWidth,T1.DisplayTopicLength,T1.DefineImage,T1.SiteHeadString,T1.SiteBottomString,T1.TableHeadString,T1.TableBottomString,T1.ShowBottomSure,T1.TempletID,T2.TempletFlag from LeadBBS_Skin as T1 Left Join LeadBBS_Templet as T2 on T1.TempletID=T2.ID Where T1.StyleID=" & ID,1),0)
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		GBL_Board_BoardLimit = 0
		Application.Lock
		Application(DEF_MasterCookies & "Style" & ID) = "yes"
		Application.UnLock
		Exit Sub
	Else
		GBL_SiteHeadString = Rs(4)
		GBL_SiteBottomString = Rs(5)
		GBL_TableHeadString = Rs(6)
		GBL_TableBottomString = Rs(7)
		Temp = Rs.GetRows(1)
		Application.Lock
		Set Application(DEF_MasterCookies & "Style" & ID) = Nothing
		Application(DEF_MasterCookies & "Style" & ID) = Temp
		Application.UnLock
		Temp = Application(DEF_MasterCookies & "Style" & ID)
		Temp(4,0) = GBL_SiteHeadString
		Temp(5,0) = GBL_SiteBottomString
		Temp(6,0) = GBL_TableHeadString
		Temp(7,0) = GBL_TableBottomString
		Application(DEF_MasterCookies & "Style" & ID) = Temp
		Rs.Close
		Set Rs = Nothing
	End If

End Sub

Sub ReloadStatisticData

	dim Rs,SQL
	'allanc(9,0),topicanc(10,0) todayanc(11,0),newuser(12,0)
	SQL = sql_select("select OnlineTime,UserCount,MaxOnline,MaxolTime,PageCount,UploadNum,MaxAnnounce,MaxAncTime,YesterdayAnc,(select sum(AnnounceNum) from LeadBBS_Boards),(select sum(TopicNum) from LeadBBS_Boards),(select sum(TodayAnnounce) from LeadBBS_Boards),(" & sql_select("select UserName from LeadBBS_User Order by ID DESC",1) & ") from LeadBBS_SiteInfo",1)
	Set Rs = LDExeCute(SQL,0)
	If Not Rs.Eof Then
		SQL = Rs.GetRows(1)
		Rs.Close
		Set Rs = Nothing
		Application.Lock
		Application(DEF_MasterCookies & "StatisticData") = SQL
		Application.UnLock
	Else
		Rs.Close
		Set Rs = Nothing
		Call LDExeCute("insert into LeadBBS_SiteInfo(OnlineTime,Version) Values(0,'" & LeftTrue(Replace(DEF_Version,"'","''"),20) & "')",1)

		dim sql2
		sql2 = sql_select("select UserName from LeadBBS_User Order by ID DESC",1)
		select Case DEF_UsedDataBase
			case 2:
				sql2 = "select t.UserName from (" & sql2 & ")as t"
			case else:
		end select
		SQL = sql_select("select OnlineTime,UserCount,MaxOnline,MaxolTime,PageCount,UploadNum,MaxAnnounce,MaxAncTime,YesterdayAnc,(select sum(AnnounceNum) from LeadBBS_Boards),(select sum(TopicNum) from LeadBBS_Boards),(select sum(TodayAnnounce) from LeadBBS_Boards),(" & sql2 & ") from LeadBBS_SiteInfo",1)
		Set Rs = LDExeCute(SQL,0)
		SQL = Rs.GetRows(1)
		Rs.Close
		Set Rs = Nothing
		Application.Lock
		Application(DEF_MasterCookies & "StatisticData") = SQL
		Application.UnLock
	End If

End Sub

Sub ReloadVesion

	dim Rs,SQL
	SQL = sql_select("select Version from LeadBBS_SiteInfo",1)
	Set Rs = LDExeCute(SQL,0)
	If Not Rs.Eof Then
		SQL = Rs(0)
		Rs.Close
		Set Rs = Nothing
	Else
		SQL = ""
		Rs.Close
		Set Rs = Nothing
	End If
	Application.Lock
	If inStr(SQL & "","LeadBBS") Then
		Application(DEF_MasterCookies & "Version") = SQL
	Else
		Application(DEF_MasterCookies & "Version") = "LeadBBS"
	End If
	Application.UnLock

End Sub

Sub UpdateStatisticDataInfo(Value,N,T)

	If isArray(Application(DEF_MasterCookies & "StatisticData")) = False Then ReloadStatisticData
	If isArray(Application(DEF_MasterCookies & "StatisticData")) = True Then
		Dim TmpArray
		TmpArray = Application(DEF_MasterCookies & "StatisticData")
		If T = 0 Then
			TmpArray(N,0) = Value
		Else
			If inStr(TmpArray(N,0),"-") Then TmpArray(N,0) = 0
			If isNumeric(TmpArray(N,0) & "") = 0 Then TmpArray(N,0) = 0
			TmpArray(N,0) = cCur(TmpArray(N,0)) + Value
		End If
		Application.Lock
		Application(DEF_MasterCookies & "StatisticData") = TmpArray
		Application.UnLock
	End If

End Sub

Sub Log_InsertEvent(LogInfo)

	Call LDExeCute("insert into LeadBBS_Log(LogType,LogTime,LogInfo,UserName,IP,BoardID) Values(9," & GetTimeValue(DEF_Now) & ",'" & Replace(Replace(htmlencode(Left(LogInfo,14)),"\","\\"),"'","''") & "','" & Replace(Replace(htmlencode(Left(GBL_CHK_User,14)),"\","\\"),"'","''") & "','" & Replace(GBL_IPAddress,"'","''") & "'," & GBL_Board_ID & ")",1)

End Sub
%>