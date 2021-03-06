<!-- #include file=../inc/BBSsetup.asp -->
<!-- #include file=../inc/User_Setup.ASP -->
<!-- #include file=../inc/Upload_Setup.asp -->
<!-- #include file=../inc/Board_Popfun.asp -->
<!-- #include file=../inc/Constellation.asp -->
<!-- #include file=../inc/Limit_fun.asp -->
<!-- #include file=inc/UserTopic.asp -->
<!-- #include file=../Search/inc/Upload_fun.asp -->
<!-- #include file=inc/Bind_Fun.asp -->
<%
DEF_BBS_HomeUrl = "../"
Dim GBL_ID,GBL_Name,GBL_NoneLimitFlag
Dim Evol,EvolString

Main

Sub Main

	GBL_Name = Request.QueryString("Name")
	GBL_ID = Left(Request.QueryString("ID"),14)
	If GBL_ID="" or isNumeric(GBL_ID)=0 Then GBL_ID = 0
	GBL_ID = cCur(GBL_ID)
	If GBL_ID =  GBL_UserID Then GBL_Name = GBL_CHK_User
	
	Evol = Left(Request.QueryString("Evol"),6)
	
	initDatabase
	
	Select Case Evol
		Case "A":EvolString = "查看用户资料"
		Case "n":EvolString = "查看用户发表的帖子"
		Case "g":EvolString = "查看用户发表的主题"
		Case "e":EvolString = "查看用户发表的精华帖子"
		Case "l":EvolString = "查看用户上传附件"
		Case "more":EvolString = "查看用户的更多信息"
		Case "f": EvolString = "好友信息"
		Case "uf": EvolString = "好友信息"
		Case "bag": EvolString = "收藏夹"
		case "bind": EvolString = "绑定网站"
		case "unbind": 
			Unbind
			exit sub
		Case Else: EvolString = "查看用户资料"
				Evol = "A"
	End Select

	BBS_SiteHead DEF_SiteNameString & " - " & EvolString,0,"<span class=navigate_string_step>" & EvolString & "</span>"
	UpdateOnlineUserAtInfo GBL_board_ID,EvolString
	
	If GBL_ID = 0 and GBL_Name = "" Then
		If GBL_ID = 0 Then GBL_ID = GBL_UserID
		GBL_CHK_TempStr = ""
		If GBL_ID = 0 Then
			GBL_CHK_TempStr = "找不到用户，要查看自己的资料请先登录。" & VbCrLf
		End If
	Else
		GBL_CHK_TempStr = ""
	End If
	If GBL_Name = "" and GBL_ID = GBL_UserID and GBL_UserID > 0 Then GBL_Name = GBL_CHK_User
	
	GBL_NoneLimitFlag = CheckSupervisorUserName  '管理员无限制
	
	If GBL_ID <> GBL_UserID or GBL_CHK_User <> GBL_Name Then
		UserTopicTopInfo("")
	Else
		UserTopicTopInfo("user")
	End If
	If GBL_UserID < 1 Then GBL_CHK_TempStr = "请确认您的身份，游客无权" & EvolString
	if GBL_CHK_TempStr <> "" Then
		Response.Write "<div class='alert redfont'>" & GBL_CHK_TempStr & "</div>"
	Else
		GBL_CHK_TempStr = ""
		Select Case Evol
			Case "n":DisplayUserAnc
			Case "g":DisplayUserTopic
			Case "e":DisplayAncGood
			Case "l":DisplayUpload
			Case "more":If LookMoreInfo = 0 Then Response.Write "<div class='alert redfont'>" & GBL_CHK_TempStr & "</div>"
			Case "f": DisplayFriend
			Case "uf": DisplayFriend
			Case "bag": DisplayFavorite
			Case "bind": DisplayBind
			Case Else: If LookUserInfo = 0 Then Response.Write "<div class='alert redfont'>" & GBL_CHK_TempStr & "</div>"
		End Select
	End If
	UserTopicBottomInfo
	closeDataBase
	SiteBottom

End Sub

Rem 显示用户资料
Function LookUserInfo

	Dim Form_Pass,Form_Mail,Form_Address,Form_SessionID
	Dim Form_Sex,Form_Birthday,Form_ApplyTime,Form_ICQ,Form_OICQ
	Dim Form_Prevtime,Form_Userphoto,Form_UserLevel,Form_Homepage,Form_Underwrite
	Dim Form_Officer,Form_Points,Form_OnlineTime
	Dim Form_FaceUrl,Form_FaceWidth,Form_FaceHeight,Form_NongLiBirth,Form_UserLimit
	Dim Form_AnnounceNum,Form_AnnounceTopic,Form_AnnounceGood,Form_NotSecret,Form_AnnounceNum2
	Dim Form_LastDoingTime,Form_CachetValue,Form_CharmPoint,LastWriteTime
	Dim Rs,SQL
	SQL = "Select ID,UserName,Mail,Address,Sex,ICQ,OICQ,Userphoto,Homepage,underwrite,birthday,NotSecret,ApplyTime,UserLevel,Officer,Points,Onlinetime,Prevtime,NongLiBirth,UserLimit,AnnounceNum,AnnounceTopic,AnnounceGood,LastDoingTime,CachetValue,CharmPoint,LastWriteTime,FaceUrl,FaceWidth,FaceHeight,AnnounceNum2,SessionID from LeadBBS_User where "
	If GBL_Name <> "" Then
		Set Rs = LDExeCute(sql_select(SQL & "UserName='" & Replace(GBL_Name,"'","''") & "'",1),0)
	Else
		Set Rs = LDExeCute(sql_select(SQL & "id=" & GBL_ID,1),0)
	End If
	If Rs.Eof Then
		GBL_CHK_TempStr = "用户不存在，要查看的用户可能已经删除，或者是游客的名字。<br>" & VbCrLf
		LookUserInfo = 0
		GBL_CHK_Flag = 0
		Rs.Close
		Set Rs = Nothing
		Exit Function
	End If

	GBL_ID = cCur(Rs(0))
	GBL_Name = Rs(1)
	Form_Mail = Rs(2)
	Form_Address = Rs(3)
	Form_Sex = Rs(4)
	Form_ICQ = Rs(5)
	Form_OICQ = Rs(6)
	Form_Userphoto = Rs(7)
	Form_Homepage = Rs(8)
	Form_Underwrite = Rs(9)
	Form_birthday = Rs(10)
	Form_NotSecret = ccur(Rs(11))
	If Form_NotSecret = 1 Then
		Form_NotSecret = 1
	Else
		Form_NotSecret = 0
	End If

	REM 特殊数据
	Form_ApplyTime = Rs(12)
	Form_UserLevel = Rs(13)
	Form_Officer = Rs(14)
	Form_Points = Rs(15)
	Form_OnlineTime = Rs(16)
	Form_Prevtime = Rs(17)
	Form_NongLiBirth = Rs(18)
	Form_UserLimit = Rs(19)
	Form_AnnounceNum = cCur(Rs(20))
	Form_AnnounceTopic = cCur(Rs(21))
	Form_AnnounceGood = cCur(Rs(22))
	Form_LastDoingTime = Rs(23)
	Form_CachetValue = cCur(Rs(24))
	Form_CharmPoint = cCur(Rs(25))
	LastWriteTime = cCur(Rs(26))
	Dim Temp

	If DEF_AllDefineFace <> 0 Then
		Form_FaceUrl = Rs(27)
		Form_FaceWidth = Rs(28)
		Form_FaceHeight = Rs(29)
	End If
	Form_AnnounceNum2 = Rs(30)
	Form_SessionID = Rs(31)
	Rs.Close
	Set Rs = Nothing
	LookUserInfo_NavInfo

	'------------special version start--------------
	If ccur(Form_OnlineTime) > 0 Then
	Set Rs = LDExeCute("select count(*) from LeadBBS_User where OnlineTime>" & Form_OnlineTime & " or (OnlineTime=" & Form_OnlineTime & " and ID<" & GBL_ID & ")" ,0)
	If Rs.Eof Then
		Form_SessionID = 0
	Else
		Form_SessionID = Rs(0)
		If isNull(Form_SessionID) Then Form_SessionID = 0
		Form_SessionID = cCur(Form_SessionID)
		If Form_SessionID > 0 Then Form_SessionID = Form_SessionID + 1
	End If
	End If
	'------------special version end--------------
	%>
		<table border=0 cellpadding="0" cellspacing="0" width=100%>
		<tr>
		<td valign=top>
			<table border=0 cellpadding="0" cellspacing="0" class="blanktable splitupright" style="">
			<tr>
				<td width=90>
					论坛排名：
				</td>
				<td>
					<%
					If cCur(Form_SessionID) > 0 Then
						Response.Write "<b><font color=blue class=bluefont>" & Form_SessionID & "</font></b>"
					Else
						Response.Write "无"
					End If%></td>
			</tr><%If Form_mail <> "" and (Form_NotSecret = 1 or GBL_UserID=GBL_ID) Then%>
			<tr>
				<td>
					电子邮件：
				</td>
				<td>
					<div class=word-break-all>
					<a href="mailto:<%=HtmlEncode(Form_mail)%>"><%=HtmlEncode(Form_mail)%></a>
					</div>
					</td>
			</tr><%End If
			If Form_homepage <> "" Then%>
			<tr>
				<td>
					主页地址：
				</td>
				<td>
					<div class=word-break-all>
					<%
					If Left(lcase(Form_homepage),4)<>"http" Then Form_homepage = "http://" & Form_homepage
					Response.Write "<a href=""" & HtmlEncode(Form_homepage) & """ target=_blank>" & HtmlEncode(Form_homepage) & "</a>"
					%>
					</div></td>
			</tr><%
			End If
			If Form_icq <> "" and Form_icq <> "0" Then%>
			<tr>
				<td>
					ICQ 号码：
				</td>
				<td>
					<%=HtmlEncode(Form_icq)%></td>
			</tr><%
			End If
			If Form_oicq <> "" and Form_oicq <> "0" and (Form_NotSecret = 1 or GBL_UserID=GBL_ID) Then%>
			<tr>
				<td>
					OICQ号码：
				</td>
				<td>
					<%=HtmlEncode(Form_oicq)%></td>
			</tr><%
			End If
			If Form_address <> "" and (Form_NotSecret = 1 or GBL_UserID=GBL_ID) Then%>
			<tr>
				<td>
					用户地址：
				</td>
				<td>
					<div class=word-break-all><%=HtmlEncode(Form_address)%>
					</div></td>
			</tr><%
			End If%>
			<tr>
				<td>
					用户性别：
				</td>
				<td>
					<%=Form_sex%></td>
			</tr><%if len(Form_birthday)=14 and (Form_NotSecret = 1 or GBL_UserID=GBL_ID) Then%>
			<tr>
				<td>
					用户生日： 
				</td>
				<td>
					<%If len(Form_birthday)=14 Then%>
					<%=RestoreTime(Left(Form_birthday,8))%></td>
					<%End If%></td>
			</tr><%End If%>
			<%If len(Form_Birthday) = 14 Then%>
			<tr>
				<td>
					星座生肖：
				</td>
				<td>
					<%
					Rs = RestoreTime(Left(Form_Birthday,8))
					If isTrueDate(Rs) Then
						Response.Write Replace(Replace(Constellation(Rs),".gif","b.gif"),"<img width=15 height=15","<img width=80 height=80")
					End If
					
					Rs = RestoreTime(Left(Form_NongLiBirth,8))
					If Len(Rs) = 10 Then%>
					<%=Replace(Replace(DisplayBirthAnimal(Left(Rs,4)),"s.gif",".gif"),"<img width=15 height=15","<img width=80 height=80")%>
			<%
					End If
			%>
				</td>
			</tr>
			<%
			End If
			If Form_Underwrite <> "" Then%>
			<tr>
				<TD colspan=2>
				<script src="<%=DEF_BBS_HomeUrl%>a/inc/leadcode.js?ver=20080728.225" type="text/javascript"></script>
					<div class=a_signature>
					<span id=UnderWrite_info class="word-break-all">
						<%=PrintTrueText(Form_Underwrite)%>
					</span>
				<script type="text/javascript">
				<!--
					var GBL_domain="<%=Temp%>";
					var DEF_DownKey="<%=UrlEncode(DEF_DownKey)%>";
					HU="<%=DEF_BBS_HomeUrl%>";
					leadcode_uw('UnderWrite_info');
				-->
				</script>
					</div>
					</td>
			</tr><%End If%>
			<tr>
				<td colspan=2>
				<hr class=splitline></td>
			</tr>
			<tr>
				<td>
					申请时间：
				</td>
				<td>
					<%=RestoreTime(Form_ApplyTime)%></td>
			</tr>
			<tr>
				<td>
					最后活动：
				</td>
				<td>
					<%
					'If cCur(Form_LastDoingTime) > cCur(Form_Prevtime) Then Form_Prevtime = Form_LastDoingTime
					If cCur(LastWriteTime) > cCur(Form_LastDoingTime) Then Form_LastDoingTime = LastWriteTime
					Response.Write RestoreTime(Form_LastDoingTime)%></td>
			</tr>
			<tr>
				<td>
					<%=DEF_PointsName(3)%>：
				</td>
				<td>
					<%=DEF_UserLevelString(Form_UserLevel)%></td>
			</tr>
			<tr>
				<td>
					<%=DEF_PointsName(0)%>：
				</td>
				<td>
					<%=HtmlEncode(Form_Points)%></td>
			</tr>
			<tr>
				<td>
					发表帖子：
				</td>
				<td>
					<%
					If Form_AnnounceNum = 0 Then
						Response.Write "无任何帖子"
					Else
						Response.Write "现存<b>" & Form_AnnounceNum & "</b>篇"
						Response.Write "，主题<b>" & Form_AnnounceTopic & "</b>篇，回复<b>" & Form_AnnounceNum - Form_AnnounceTopic & "</b>篇"
						Response.Write "<br>精华<b>" & Form_AnnounceGood & "</b>篇"
					End If
					%> 历史累计<b><%=Form_AnnounceNum2%></b>篇</td>
			</tr>
			<tr>
				<td>
					<%=DEF_PointsName(4)%>：
				</td>
				<td>
					<%=clng(cCur(Form_OnlineTime)/60)%></td>
			</tr><%
			If Form_CachetValue <> 0 Then%>
			<tr>
				<td>
					<%=DEF_PointsName(2)%>：
				</td>
				<td>
			<%
				If Form_CachetValue > 0 Then
					Response.Write "<font color=blue class=bluefont>＋" & Form_CachetValue & "</font><br>"
				Else
					Response.Write Form_CachetValue & "<br>"
				End If%></td>
			</tr>
			<%
			End If
			If Form_CharmPoint <> 0 Then%>
			<tr>
				<td>
					<%=DEF_PointsName(1)%>：
				</td>
				<td>
					<b><font color=red><%=Form_CharmPoint%></font></b> <a href=alipay/Payment.asp>立即充值</a></td>
			</tr>
			<%
			End If
			If Form_Officer <> "0" and Form_Officer <> "" Then%>
			<tr>
				<td>
					<%=DEF_PointsName(9)%>：
				</td>
				<td>
					<%=DisplayOfficerString(Form_Officer)%></td>
			</tr><%End If%><%
			If GetBinarybit(Form_UserLimit,8) = 1 or GetBinarybit(Form_UserLimit,10) = 1 or GetBinarybit(Form_UserLimit,14) = 1 or GetBinarybit(Form_UserLimit,2) = 1 Then%>
			<tr>
				<TD valign=top>
					其它信息：
				</td>
				<td>
					<%
			If GetBinarybit(Form_UserLimit,10) = 1 Then
				Response.Write "<font color=555555>职务：</font>" & DEF_PointsName(6) & "<br>"
			ElseIf GetBinarybit(Form_UserLimit,14) = 1 Then
				Response.Write "<font color=555555 class=grayfont>担任</font><b>" & DEF_PointsName(7) & "</b><font color=555555 class=grayfont>一职</font>"
			ElseIf GetBinarybit(Form_UserLimit,8) = 1 Then
				Response.Write "<font color=555555 class=grayfont>担任</font><b>" & DEF_PointsName(8) & "</b><font color=555555 class=grayfont>一职</font>"
			End If

			If GetBinarybit(Form_UserLimit,2) = 1 Then
				Response.Write " <font color=555555 class=grayfont>已经是</font>" & DEF_PointsName(5) & "<br>"
			End If%></td>
			</tr><%End If%><%If GBL_UserID=GBL_ID Then%>
			<tr>
				<td colspan=2>
				<hr class=splitline></td>
			</tr>
			<tr>
				<td height=25 valign=top>
					我的权限<br>和设置：<p>
					<font color=888888 class=grayfont>某些设定<br>
					仅版主以<br>上才有用</font></td>
				<td>
					<table cellpadding="0" cellspacing="0"><%
			Dim TempN
			For TempN = 1 to LimitUserStringDataNum
				If (GetBinarybit(Form_UserLimit,8) = 1 or GetBinarybit(Form_UserLimit,10) = 1) or GBL_NoneLimitFlag = 1 Then
					Response.Write "<tr height=20><td>" & LimitUserStringData(tempN)
					If GetBinarybit(Form_UserLimit,TempN+1) = 1 Then
						Response.Write "</td><td>是</td></tr>"
					Else
						Response.Write "</td><td>否</td></tr>"
					End If
				Else
					If TempN = 4 or TempN = 8 or TempN = 5 or TempN = 10 or TempN = 11 or TempN = 14 or TempN = 15 Then
						'此类资料对于非版主以上用户无意义,所以屏蔽,另,是否专业用户也不作显示处理
					Else
						Response.Write "<tr height=20><td>" & LimitUserStringData(tempN)
						If GetBinarybit(Form_UserLimit,TempN+1) = 1 Then
							Response.Write "</td><td>是</td></tr>"
						Else
							Response.Write "</td><td>否</td></tr>"
						End If
					End If
				End If
			Next%></table></td>
			</tr><%End If%>
			</table>
			<%
			If isNull(Form_FaceUrl) Then Form_FaceUrl = ""
			If DEF_AllDefineFace = 0 or Trim(Form_FaceUrl) = "" Then%>
			<img src=<%=DEF_BBS_HomeUrl%>images/face/<%=string(4-len(cstr(Form_userphoto)),"0")&Form_userphoto%>.gif align=middle>
			<%Else%>
				<img src="<%=htmlencode(Form_FaceUrl)%>" align=middle width=<%=Form_FaceWidth%> height=<%=Form_FaceHeight%>>
			<%End If%>

			<div class=title><%=Server.HtmlEncode(GBL_Name)%></div>
			<%If GBL_CHK_User <> GBL_Name Then%>
			<div class=value2><a href="../a/Processor.asp?action=AddFriend&FriendName=<%=UrlEncode(GBL_Name)%>" onclick="return(pub_msg(this,'anc_msgbody','&SureFlag=1'));">加为好友</a></div>
			<div class=value2><a href="SendMessage.asp?SdM_ToUser=<%=HtmlEncode(GBL_Name)%>" onclick="return(sendprivatemsg(this,'<%=DEF_BBS_HomeUrl%>'));">发送短消息</a></div>
			<%End If%>
			<div class=value2><a href="LookUserInfo.asp?ID=<%=GBL_ID%>&Evol=more">查看更多信息</a></div>
		</td>
		</tr>
		</table>
	<%

End Function

Function PrintTrueText(tempString)

	If tempString<>"" Then
		PrintTrueText=Replace(Replace(Replace(Replace(Replace(htmlEncode(tempString),VbCrLf & " ","<br />" & "&nbsp;"),VbCrLf,"<br />" & VbCrLf),"   "," &nbsp; "),"  "," &nbsp;"),chr(9)," &nbsp; &nbsp; &nbsp;")
		If Left(PrintTrueText,1) = chr(32) Then
			PrintTrueText = "&nbsp;" & Mid(PrintTrueText,2)
		End If
	Else
		PrintTrueText=""
	End If

End Function

Function DisplayOfficerString(Officer)

	Dim Officer_Temp,Temp_N,dotFlag
	dotFlag = 0
	Officer_Temp = split(Officer,",")
	For Temp_N = 0 to Ubound(Officer_Temp,1)
		If isNumeric(Officer_Temp(Temp_N)) Then
			Officer_Temp(Temp_N) = cCur(Officer_Temp(Temp_N))
			If Officer_Temp(Temp_N)>=0 and Officer_Temp(Temp_N)<=DEF_UserOfficerNum Then
				If dotFlag = 0 Then
					dotFlag = 1
					DisplayOfficerString = DisplayOfficerString & DEF_UserOfficerString(Officer_Temp(Temp_N))
				Else
					DisplayOfficerString = DisplayOfficerString & "," & DEF_UserOfficerString(Officer_Temp(Temp_N))
				End If
			End If
		End If
	Next

End Function

Rem 显示用户发表的帖子
Sub DisplayUserAnc

	Dim Rs,SQL,NewNum,RecordCount
	If GBL_ID > 0 Then
		Set Rs = LDExeCute(sql_select("Select ID,UserName,AnnounceNum from LeadBBS_User where ID=" & GBL_ID,1),0)
	Else
		Set Rs = LDExeCute(sql_select("Select ID,UserName,AnnounceNum from LeadBBS_User where UserName='" & Replace(GBL_Name,"'","''") & "'",1),0)
	End If
	If Not Rs.Eof Then
		GBL_Name = Rs(1)
		GBL_ID = cCur(Rs(0))
		RecordCount = cCur(Rs(2))
	Else
		Response.Write "<div class=alert>错误，此用户不存在！</div>"
		Rs.Close
		Set Rs = Nothing
		Exit Sub
	End if
	Rs.close
	Set Rs = Nothing

	GBL_CHK_TempStr=""
	Dim UpDownPageFlag
	UpDownPageFlag = Request("UpDownPageFlag")

	Dim Start,key

	Dim SQLendString

	Start = Left(Trim(Request("Start")),14)
	Start = start
	If isNumeric(Start)=0 or Start="" Then Start=999999999
	Start = cCur(Start)
	If Start = 0 Then Start=999999999

	Dim SQLCountString,whereFlag
	whereFlag = 1
	SQLendString = " where UserID=" & GBL_ID

	SQLCountString = SQLendString
	If UpDownPageFlag = "1" and Start>0 then
		If whereFlag = 1 Then
			SQLendString = SQLendString & " and ID>" & Start
		Else
			SQLendString = SQLendString & " where ID>" & Start
			whereFlag = 1
		End If
	Else
		If whereFlag = 1 Then
			SQLendString = SQLendString & " and ID<" & Start
		Else
			SQLendString = SQLendString & " where ID<" & Start
			whereFlag = 1
		End If
	end If

	If UpDownPageFlag = "1" then
		SQLendString = SQLendString & " Order by  ID ASC"
	Else
		SQLendString = SQLendString & " Order by ID DESC"
	End If
	
	Dim MaxRecordID,MinRecordID
	MaxRecordID = 0

	SQL = "select Max(id) from LeadBBS_Announce " & SQLCountString
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
	SQL = "select Min(id) from LeadBBS_Announce " & SQLCountString
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
	Dim FirstID,LastID

	SQL = sql_select("select T1.ID,T1.Title,T1.Length,T1.ndatetime,T1.Hits,T1.FaceIcon,T1.ChildNum,T1.BoardID,T1.GoodFlag,T1.TitleStyle,T1.ParentID,T2.ForumPass,T2.BoardLimit,T2.OtherLimit,T2.HiddenFlag,T1.RootIDBak from LeadBBS_Announce as T1 left join LeadBBS_Boards as T2 on T1.BoardID=T2.BoardID " & SQLendString,DEF_MaxListNum)
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

		Dim EndwriteQueryString,PageSplictString
		EndwriteQueryString = "?ID=" & GBL_ID & "&Evol=n"
		If key<>"" Then EndwriteQueryString = EndwriteQueryString & "&key=" & urlencode(key)

		PageSplictString = PageSplictString & "<div class=j_page>"
		If FirstID >= MaxRecordID Then
			'PageSplictString = PageSplictString & "首页 " & VbCrLf
			'PageSplictString = PageSplictString & " 上页 " & VbCrLf
		Else
			PageSplictString = PageSplictString & "<a href=LookUserInfo.asp" & EndwriteQueryString & "&Start=0>首页</a> " & VbCrLf
			PageSplictString = PageSplictString & " <a href=LookUserInfo.asp" & EndwriteQueryString & "&Start=" & FirstID & "&UpDownPageFlag=1>上页</a> " & VbCrLf
		End if
	
		If LastID <= MinRecordID Then
			'PageSplictString = PageSplictString & " 下页 " & VbCrLf
			'PageSplictString = PageSplictString & " 尾页 " & VbCrLf
		else
			PageSplictString = PageSplictString & " <a href=LookUserInfo.asp" & EndwriteQueryString & "&Start=" & LastID & ">下页</a> " & VbCrLf
			PageSplictString = PageSplictString & " <a href=LookUserInfo.asp" & EndwriteQueryString & "&Start=1&UpDownPageFlag=1>尾页</a> " & VbCrLf
		end if
		
		PageSplictString = PageSplictString & "<b>共" & RecordCount & "</b>"
		'If (RecordCount mod DEF_MaxListNum)=0 Then
		'	PageSplictString = PageSplictString & " 计<b>" & clng(RecordCount/DEF_MaxListNum) & "</b>页"
		'Else
		'	If RecordCount>=DEF_MaxListNum Then
		'		PageSplictString = PageSplictString & " 计<b>" & clng(RecordCount/DEF_MaxListNum)+1 & "</b>页"
		'	Else
		'		PageSplictString = PageSplictString & " 计<b>1</b>页"
		'	End If
		'End If
		'PageSplictString = PageSplictString & " 每页<b>" & DEF_MaxListNum & "</b>条记录"
		PageSplictString = PageSplictString & "</div>"
	
	End If

	LookUserInfo_NavInfo
	%>
	<table width="100%" border="0" cellspacing="0" cellpadding="0" class=table_in>
	  <tr class=tbinhead>
	    <td><div class=value>帖子</div></td>
	    <td width=80><div class=value title="回复/点击">人气</div></td>
	    <td width=125><div class=value>发表时间</div></td>
	  </tr>
	<%
	If Num = -1 Then
		Response.Write "<tr><td colspan=3 class=tdbox>没有相关的帖子!</td></tr>"
	end if
	
	
	Dim TempN,Temp,Temp1
	
	if Num <> -1 then
		i=1
		LastID = GetData(0,ubound(getdata,2))
		for n= MinN to MaxN Step StepValue
			Response.Write "<tr>"
			'Response.Write "<td class=tdbox><img src=../images/bf/face" & GetData(5,N) & ".gif align=absbottom></td>"
			Response.Write "<td class=tdbox><a href=../a/a.asp?B=" & GetData(7,n) & "&ID=" & GetData(15,N)
			If cCur(GetData(10,n)) > 0 Then
				Response.Write "&RID=" & GetData(0,N) & "#F" & GetData(0,N)
			End If
			Response.Write ">"

			GetData(6,N) = cCur(GetData(6,N))
			Temp1 = Fix((GetData(6,N)+1)/DEF_TopicContentMaxListNum)
			If ((GetData(6,N)+1) mod DEF_TopicContentMaxListNum) > 0 Then Temp1 = Temp1 + 1
			If GetData(6,N)>=DEF_TopicContentMaxListNum Then
				Temp = DEF_BBS_DisplayTopicLength - (Len(Temp1) + 3)
			Else
				Temp = DEF_BBS_DisplayTopicLength
			End If

			If ccur(GetData(8,n)) = 1 Then Temp = Temp - 3
			If GBL_NoneLimitFlag = 0 and GBL_CheckLimitTitle(GetData(11,n),GetData(12,n),GetData(13,n),GetData(14,n)) = 1 Then
				GetData(1,n) = "此帖子标题已设置为隐藏"
				GetData(9,n) = 1
			End If

			If cCur(GetData(10,N)) > 0 and Left(GetData(1,N),3) = "re:" and GetData(1,N) <> "" Then GetData(1,N) = Mid(GetData(1,N),4)
			If GetData(9,n) <> 1 Then
				If strLength(GetData(1,N))>Temp-1 Then GetData(1,N) = LeftTrue(GetData(1,N),Temp-4) & "..."
			Else
				If strLength(GetData(1,N))>Temp-1 Then GetData(1,N) = LeftTrueHTML(GetData(1,N),Temp-4)
			End If
			Response.Write DisplayAnnounceTitle(GetData(1,n),GetData(9,n))
			Response.Write "</a>"

			If GetData(6,N)>=DEF_TopicContentMaxListNum Then
				Response.Write " [<a href=../a/a.asp?B=" & GetData(7,N) & "&ID=" & GetData(0,N) & "&AUpflag=1&ANum=1 title=" & GetData(2,n) & "字节>" & Temp1 & "</b></a>]"
			End If

			If ccur(GetData(8,n)) = 1 Then
				Response.Write "<img src=../images/" & GBL_DefineImage & "jh1.GIF border=0 title=精华帖子 align=absbottom>"
			End If
			Response.Write "</td><td class=tdbox><em>"
			Response.Write GetData(6,N) & "/" & GetData(4,N)
			Response.Write "</em></td><td class=tdbox><EM>"
			Response.Write Left(RestoreTime(GetData(3,n)),16) & "</EM></td>"
			Response.Write "</tr>" & VbCrLf
			i=i+1
		next
	End If
	If PageSplictString<>"" Then Response.Write "<tr><td colspan=3 class=tdbox>" & PageSplictString & "</td></tr>"
	%>
	      </table>
	<%

End Sub

Rem 显示用户发表的主题
Function DisplayUserTopic

	Dim Rs,SQL,NewNum,RecordCount
	
	If GBL_ID > 0 Then
		Set Rs = LDExeCute(sql_select("Select ID,UserName,AnnounceTopic from LeadBBS_User where ID=" & GBL_ID,1),0)
	Else
		Set Rs = LDExeCute(sql_select("Select ID,UserName,AnnounceTopic from LeadBBS_User where UserName='" & Replace(GBL_Name,"'","''") & "'",1),0)
	End If
	If Not Rs.Eof Then
		GBL_Name = Rs(1)
		GBL_ID = cCur(Rs(0))
		RecordCount = cCur(Rs(2))
	Else
		RecordCount = 0
		Response.Write "<div class=alert>错误，此用户不存在！</div>"
		Rs.Close
		Set Rs = Nothing
		Exit Function
	End if
	Rs.close
	Set Rs = Nothing
	
	GBL_CHK_TempStr=""
	Dim UpDownPageFlag
	UpDownPageFlag = Request("UpDownPageFlag")

	Dim Start,key
	
	Dim SQLendString

	Start = Left(Trim(Request.QueryString("Start")),14)
	Start = start
	If isNumeric(Start)=0 or Start="" Then Start=999999999
	Start = cCur(Start)
	If Start = 0 Then Start=999999999

	Dim SQLCountString,whereFlag
	whereFlag = 1
	select case DEF_UsedDataBase
		case 0,2:
			SQLendString = " where UserID=" & GBL_ID & " and parentID=0"
		case Else
			SQLendString = " where UserID=" & GBL_ID & " "
	End select

	SQLCountString = SQLendString
	If UpDownPageFlag = "1" and Start>0 then
		If whereFlag = 1 Then
			SQLendString = SQLendString & " and ID>" & Start
		Else
			SQLendString = SQLendString & " where ID>" & Start
			whereFlag = 1
		End If
	Else
		If whereFlag = 1 Then
			SQLendString = SQLendString & " and ID<" & Start
		Else
			SQLendString = SQLendString & " where ID<" & Start
			whereFlag = 1
		End If
	end If

	If UpDownPageFlag = "1" then
		SQLendString = SQLendString & " Order by  ID ASC"
	Else
		SQLendString = SQLendString & " Order by ID DESC"
	End If
	
	Dim MaxRecordID,MinRecordID
	MaxRecordID = 0

	select case DEF_UsedDataBase
		case 0,2:
			SQL = "select Max(id) from LeadBBS_Announce " & SQLCountString
		case Else
			SQL = "select Max(id) from LeadBBS_Topic " & SQLCountString
	End select
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
	
	select case DEF_UsedDataBase
		case 0,2:
			SQL = "select Min(id) from LeadBBS_Announce " & SQLCountString
		case Else
			SQL = "select Min(id) from LeadBBS_Topic " & SQLCountString
	End select
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

	Dim FirstID,LastID	

	select case DEF_UsedDataBase
		case 0,2:
			SQL = sql_select("select T1.ID,T1.Title,T1.Length,T1.ndatetime,T1.Hits,T1.FaceIcon,T1.ChildNum,T1.BoardID,T1.GoodFlag,T1.TitleStyle,T2.ForumPass,T2.BoardLimit,T2.OtherLimit,T2.HiddenFlag from LeadBBS_Announce as T1 left join LeadBBS_Boards as T2 on T1.BoardID=T2.BoardID " & SQLendString,DEF_MaxListNum)
		case Else
			SQL = sql_select("select T1.ID,T1.Title,T1.Length,T1.ndatetime,T1.Hits,T1.FaceIcon,T1.ChildNum,T1.BoardID,T1.GoodFlag,T1.TitleStyle,T2.ForumPass,T2.BoardLimit,T2.OtherLimit,T2.HiddenFlag from LeadBBS_Topic as T1 left join LeadBBS_Boards as T2 on T1.BoardID=T2.BoardID " & SQLendString,DEF_MaxListNum)
	End select
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
	
		Dim EndwriteQueryString,PageSplictString
		EndwriteQueryString = "?ID=" & GBL_ID & "&Evol=g"
		If key<>"" Then EndwriteQueryString = EndwriteQueryString & "&key=" & urlencode(key)
	
		PageSplictString = PageSplictString & "<div class=j_page>"
		If FirstID >= MaxRecordID Then
			'PageSplictString = PageSplictString & "首页 " & VbCrLf
			'PageSplictString = PageSplictString & " 上页 " & VbCrLf
		else
			PageSplictString = PageSplictString & "<a href=LookUserInfo.asp" & EndwriteQueryString & "&Start=0>首页</a> " & VbCrLf
			PageSplictString = PageSplictString & " <a href=LookUserInfo.asp" & EndwriteQueryString & "&Start=" & FirstID & "&UpDownPageFlag=1>上页</a> " & VbCrLf
		end if
	
		if LastID<MaxRecordID and LastID<>0 then
		else
		end if
	
		If LastID <= MinRecordID Then
			'PageSplictString = PageSplictString & " 下页 " & VbCrLf
			'PageSplictString = PageSplictString & " 尾页 " & VbCrLf
		else
			PageSplictString = PageSplictString & " <a href=LookUserInfo.asp" & EndwriteQueryString & "&Start=" & LastID & ">下页</a> " & VbCrLf
			PageSplictString = PageSplictString & " <a href=LookUserInfo.asp" & EndwriteQueryString & "&Start=1&UpDownPageFlag=1>尾页</a> " & VbCrLf
		end if
		
		PageSplictString = PageSplictString & "<b>共" & RecordCount & "</b>"
		'If (RecordCount mod DEF_MaxListNum)=0 Then
		'	PageSplictString = PageSplictString & " 计<b>" & clng(RecordCount/DEF_MaxListNum) & "</b>页"
		'Else
		'	If RecordCount>=DEF_MaxListNum Then
		'		PageSplictString = PageSplictString & " 计<b>" & clng(RecordCount/DEF_MaxListNum)+1 & "</b>页"
		'	Else
		'		PageSplictString = PageSplictString & " 计<b>1</b>页"
		'	End If
		'End If
		'PageSplictString = PageSplictString & " 每页<b>" & DEF_MaxListNum & "</b>条记录"
		PageSplictString = PageSplictString & "</div>"
	
	End If

	LookUserInfo_NavInfo
	%>
	<table width="100%" border="0" cellspacing="0" cellpadding="0" class=table_in>
	  <tr class=tbinhead>
	    <td><div class=value>主题</div></td>
	    <td width=80><div class=value title="回复/点击">人气</div></td>
	    <td width=125><div class=value>发表时间</div></td>
	  </tr>
	<%
	If Num = -1 Then
		Response.Write "<tr><td colspan=3 class=tdbox>没有任何主题!</td></tr>"
	end if
	
	
	Dim TempN,Temp,Temp1
	
	if Num <> -1 then
		i=1
		LastID = GetData(0,ubound(getdata,2))
		for n= MinN to MaxN Step StepValue
			Response.Write "<tr>"
			'Response.Write "<td class=tdbox><img src=../images/bf/face" & GetData(5,N) & ".gif align=absbottom width=20 height=20></td>"
			Response.Write "<td class=tdbox><a href=../a/a.asp?B=" & GetData(7,n) & "&ID=" & GetData(0,N) & ">"
			
			GetData(6,N) = cCur(GetData(6,N))
			Temp1 = Fix((GetData(6,N)+1)/DEF_TopicContentMaxListNum)
			If ((GetData(6,N)+1) mod DEF_TopicContentMaxListNum) > 0 Then Temp1 = Temp1 + 1
			If GetData(6,N)>=DEF_TopicContentMaxListNum Then
				Temp = DEF_BBS_DisplayTopicLength - (Len(Temp1) + 3)
			Else
				Temp = DEF_BBS_DisplayTopicLength
			End If

			If ccur(GetData(8,n)) = 1 Then Temp = Temp - 3
			If GBL_NoneLimitFlag = 0 and GBL_CheckLimitTitle(GetData(10,n),GetData(11,n),GetData(12,n),GetData(13,n)) = 1 Then
				GetData(1,n) = "此帖子标题已设置为隐藏"
				GetData(9,n) = 1
			End If

			If GetData(9,n) <> 1 Then
				If strLength(GetData(1,N))>Temp-1 Then GetData(1,N) = LeftTrue(GetData(1,N),Temp-4) & "..."
			Else
				If strLength(GetData(1,N))>Temp-1 Then GetData(1,N) = LeftTrueHTML(GetData(1,N),Temp-4)
			End If
			Response.Write DisplayAnnounceTitle(GetData(1,n),GetData(9,n))
			Response.Write "</a>"

			If GetData(6,N)>=DEF_TopicContentMaxListNum Then
				Response.Write " [<a href=../a/a.asp?B=" & GetData(7,N) & "&ID=" & GetData(0,N) & "&AUpflag=1&ANum=1 title=" & GetData(2,n) & "字节>" & Temp1 & "</b></a>]"
			End If
	
			If ccur(GetData(8,n)) = 1 Then
				Response.Write "<img src=../images/" & GBL_DefineImage & "jh1.GIF border=0 title=精华帖子 align=absbottom>"
			End If
			Response.Write "</td><td class=tdbox><em>"
			Response.Write GetData(6,N) & "/" & GetData(4,N)
			Response.Write "</em></td><td width=125 class=tdbox><em>"
			Response.Write Left(RestoreTime(GetData(3,n)),16) & "</em></td>"
			Response.Write "</tr>" & VbCrLf
			i=i+1
		next
	End If
	If PageSplictString<>"" Then Response.Write "<tr><td colspan=3 class=tdbox>" & PageSplictString & "</td></tr>"
	%>
	      </table>
	<br><%

End Function

Rem 查看用户发表的精华帖子
Function DisplayAncGood

	Dim Rs,SQL,NewNum,RecordCount
	
	If GBL_ID > 0 Then
		Set Rs = LDExeCute(sql_select("Select ID,UserName,AnnounceGood from LeadBBS_User where ID=" & GBL_ID,1),0)
	Else
		Set Rs = LDExeCute(sql_select("Select ID,UserName,AnnounceGood from LeadBBS_User where UserName='" & Replace(GBL_Name,"'","''") & "'",1),0)
	End If
	If Not Rs.Eof Then
		GBL_Name = Rs(1)
		GBL_ID = cCur(Rs(0))
		RecordCount = cCur(Rs(2))
	Else
		RecordCount = 0
		Response.Write "<br><br>&nbsp; &nbsp; &nbsp; 错误，此用户不存在！"
		Rs.Close
		Set Rs = Nothing
		Exit Function
	End if
	Rs.close
	Set Rs = Nothing
	
	GBL_CHK_TempStr=""
	Dim UpDownPageFlag
	UpDownPageFlag = Request("UpDownPageFlag")

	Dim Start,key
	Dim SQLendString

	Start = Left(Trim(Request("Start")),14)
	Start = start
	If isNumeric(Start)=0 or Start="" Then Start=999999999
	Start = cCur(Start)
	If Start = 0 Then Start=999999999

	Dim SQLCountString,whereFlag
	whereFlag = 1
	SQLendString = " where GoodFlag=1 and UserID=" & GBL_ID

	SQLCountString = SQLendString
	If UpDownPageFlag = "1" and Start>0 then
		If whereFlag = 1 Then
			SQLendString = SQLendString & " and ID>" & Start
		Else
			SQLendString = SQLendString & " where ID>" & Start
			whereFlag = 1
		End If
	Else
		If whereFlag = 1 Then
			SQLendString = SQLendString & " and ID<" & Start
		Else
			SQLendString = SQLendString & " where ID<" & Start
			whereFlag = 1
		End If
	end If

	If UpDownPageFlag = "1" then
		SQLendString = SQLendString & " Order by  ID ASC"
	Else
		SQLendString = SQLendString & " Order by ID DESC"
	End If
	
	Dim MaxRecordID,MinRecordID
	MaxRecordID = 0

	select case DEF_UsedDataBase
		case 0,2:
			SQL = "select Max(id) from LeadBBS_Announce " & SQLCountString
		case Else
			SQL = "select Max(id) from LeadBBS_Topic " & SQLCountString
	End select
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

	select case DEF_UsedDataBase
		case 0,2:
			SQL = "select Min(id) from LeadBBS_Announce " & SQLCountString
		case Else
			SQL = "select Min(id) from LeadBBS_Topic " & SQLCountString
	End select
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

	Dim FirstID,LastID

	select case DEF_UsedDataBase
		case 0,2:
			SQL = sql_select("select T1.ID,T1.Title,T1.Length,T1.ndatetime,T1.Hits,T1.FaceIcon,T1.ChildNum,T1.BoardID,T1.GoodFlag,T1.ParentID,T1.TitleStyle,T2.ForumPass,T2.BoardLimit,T2.OtherLimit,T2.HiddenFlag,T1.RootIDBak from LeadBBS_Announce as T1 left join LeadBBS_Boards as T2 on T1.BoardID=T2.BoardID " & SQLendString,DEF_MaxListNum)
		case Else
			SQL = sql_select("select T1.ID,T1.Title,T1.Length,T1.ndatetime,T1.Hits,T1.FaceIcon,T1.ChildNum,T1.BoardID,T1.GoodFlag,0,T1.TitleStyle,T2.ForumPass,T2.BoardLimit,T2.OtherLimit,T2.HiddenFlag,T1.ID from LeadBBS_Topic as T1 left join LeadBBS_Boards as T2 on T1.BoardID=T2.BoardID " & SQLendString,DEF_MaxListNum)
	End select
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
		i = 1
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
	
		Dim EndwriteQueryString,PageSplictString
		EndwriteQueryString = "?ID=" & GBL_ID & "&Evol=e"
		If key<>"" Then EndwriteQueryString = EndwriteQueryString & "&key=" & urlencode(key)
	
		PageSplictString = PageSplictString & "<div class=j_page>"
		If FirstID >= MaxRecordID Then
			'PageSplictString = PageSplictString & "首页 " & VbCrLf
			'PageSplictString = PageSplictString & "上页 " & VbCrLf
		Else
			PageSplictString = PageSplictString & "<a href=LookUserInfo.asp" & EndwriteQueryString & "&Start=0>首页</a> " & VbCrLf
			PageSplictString = PageSplictString & "<a href=LookUserInfo.asp" & EndwriteQueryString & "&Start=" & FirstID & "&UpDownPageFlag=1>上页</a> " & VbCrLf
		End if

		If LastID <= MinRecordID Then
			'PageSplictString = PageSplictString & "下页 " & VbCrLf
			'PageSplictString = PageSplictString & "尾页 " & VbCrLf
		Else
			PageSplictString = PageSplictString & "<a href=LookUserInfo.asp" & EndwriteQueryString & "&Start=" & LastID & ">下页</a>" & VbCrLf
			PageSplictString = PageSplictString & "<a href=LookUserInfo.asp" & EndwriteQueryString & "&Start=1&UpDownPageFlag=1>尾页</a>" & VbCrLf
		End If

		PageSplictString = PageSplictString & "<b>共" & RecordCount & "</b>"
		'If (RecordCount mod DEF_MaxListNum)=0 Then
		'	PageSplictString = PageSplictString & " 计<b>" & clng(RecordCount/DEF_MaxListNum) & "</b>页"
		'Else
		'	If RecordCount>=DEF_MaxListNum Then
		'		PageSplictString = PageSplictString & " 计<b>" & clng(RecordCount/DEF_MaxListNum)+1 & "</b>页"
		'	Else
		'		PageSplictString = PageSplictString & " 计<b>1</b>页"
		'	End If
		'End If
		'PageSplictString = PageSplictString & " 每页<b>" & DEF_MaxListNum & "</b>条记录"
		PageSplictString = PageSplictString & "</div>"
	End If

	LookUserInfo_NavInfo
	%>
	
	<table width="100%" border="0" cellspacing="0" cellpadding="0" class=table_in>
	  <tr class=tbinhead>
	    <td><div class=value>帖子</div></td>
	    <td width=80><div class=value title="回复/点击">人气</div></td>
	    <td width=125><div class=value>发表时间</div></td>
	  </tr>
	<%
	If Num = -1 Then
		Response.Write "<tr><td colspan=3 class=tdbox>没有任何帖子!</td></tr>"
	End if

	Dim TempN,Temp,Temp1

	if Num <> -1 then
		i = 1
		LastID = GetData(0,ubound(getdata,2))
		for n= MinN to MaxN Step StepValue
			Response.Write "<tr>"
			'Response.Write "<td class=tdbox><img src=../images/bf/face" & GetData(5,N) & ".gif align=absbottom></td>"
			Response.Write "<td class=tdbox><a href=../a/a.asp?B=" & GetData(7,n) & "&ID=" & GetData(15,N)
			If cCur(GetData(9,n)) > 0 Then
				Response.Write "&RID" & GetData(0,N) & "#F" & GetData(0,N)
			End If
			Response.Write ">"
			GetData(6,N) = cCur(GetData(6,N))
			Temp1 = Fix((GetData(6,N)+1)/DEF_TopicContentMaxListNum)
			If ((GetData(6,N)+1) mod DEF_TopicContentMaxListNum) > 0 Then Temp1 = Temp1 + 1
			If GetData(6,N)>=DEF_TopicContentMaxListNum Then
				Temp = DEF_BBS_DisplayTopicLength - (Len(Temp1) + 3)
			Else
				Temp = DEF_BBS_DisplayTopicLength
			End If
		
			'If ccur(GetData(8,n)) = True Then Temp = Temp - 3
			If GBL_NoneLimitFlag = 0 and GBL_CheckLimitTitle(GetData(11,n),GetData(12,n),GetData(13,n),GetData(14,n)) = 1 Then
				GetData(1,n) = "此帖子标题已设置为隐藏"
				GetData(10,n) = 1
			End If
			
			If GetData(10,n) <> 1 Then
				If strLength(GetData(1,N))>Temp-1 Then GetData(1,N) = LeftTrue(GetData(1,N),Temp-4) & "..."
			Else
				If strLength(GetData(1,N))>Temp-1 Then GetData(1,N) = LeftTrueHTML(GetData(1,N),Temp-4)
			End If
			Response.Write DisplayAnnounceTitle(GetData(1,n),GetData(10,n))
			Response.Write "</a>"

			If GetData(6,N)>=DEF_TopicContentMaxListNum Then
				Response.Write " [<a href=../a/a.asp?B=" & GetData(7,N) & "&ID=" & GetData(0,N) & "&AUpflag=1&ANum=1 title=" & GetData(2,n) & "字节>" & Temp1 & "</b></a>]"
			End If
	
			'If ccur(GetData(8,n)) = 1 Then
			'	Response.Write "<img src=../images/" & GBL_DefineImage & "jh1.GIF border=0 title=精华帖子 align=absbottom width=15 height=16>"
			'End If
			Response.Write "</td><td class=tdbox><em>"
			Response.Write GetData(6,N) & "/" & GetData(4,N)
			Response.Write "</em></td><td class=tdbox><em>"
			Response.Write Left(RestoreTime(GetData(3,n)),16) & "</em></td>"
			Response.Write "</tr>" & VbCrLf
			i=i+1
		next
	End If
	If PageSplictString<>"" Then Response.Write "<tr><td colspan=3 class=tdbox>" & PageSplictString & "</td></tr>"
	%>
	      </table>
	<br><%

End Function

Rem 显示上传附件
Sub DisplayUpload

	Dim Rs,SQL,NewNum,RecordCount
	If GBL_ID > 0 Then
		Set Rs = LDExeCute(sql_select("Select ID,UserName,UploadNum from LeadBBS_User where ID=" & GBL_ID,1),0)
	Else
		Set Rs = LDExeCute(sql_select("Select ID,UserName,UploadNum from LeadBBS_User where UserName='" & Replace(GBL_Name,"'","''") & "'",1),0)
	End If
	If Not Rs.Eof Then
		GBL_Name = Rs(1)
		GBL_ID = cCur(Rs(0))
		RecordCount = cCur(Rs(2))
	Else
		LookUserInfo_NavInfo
		Response.Write "<div class=alert>此用户不存在！</div>"
		Rs.Close
		Set Rs = Nothing
		Exit Sub
	End if
	Rs.close
	Set Rs = Nothing

	If GBL_UserID <> GBL_ID and GBL_NoneLimitFlag = 0 Then
		LookUserInfo_NavInfo
		Response.Write "<div class='alert redfont'>附件已设置为只允许作者本人查看！</div>"
		Exit Sub
	End If

	LookUserInfo_NavInfo
	CALL Upload_List(GBL_ID,RecordCount,"../User/LookUserInfo.asp?ID=" & GBL_ID & "&Evol=l",1)


End Sub

Rem 查看更多信息
Function LookMoreInfo

	Dim Online_OnlineFlag,Online_LastDoingTime,Online_IP,Online_StartTime,LookUserLevel

	Dim Form_ID,Form_IP,Form_Login_oknum,SessionID,Browser,System
	Dim Rs,AtUrl,AtInfo,Login_RightIP,OlUserName

	Dim OlID

	OlID = Left(Request.QueryString("OlID"),14)
	If OlID="" or isNumeric(OlID)=0 Then OlID = 0
	OlID = cCur(OlID)
	If OlID > 0 Then
		Set Rs = LDExeCute(sql_select("Select LastDoingTime,StartTime,IP,AtUrl,AtInfo,UserID,SessionID,Browser,System,ID,UserName from LeadBBS_onlineUser where ID=" & OlID,1),0)
	Else
		Set Rs = LDExeCute(sql_select("Select LastDoingTime,StartTime,IP,AtUrl,AtInfo,UserID,SessionID,Browser,System,ID,UserName from LeadBBS_onlineUser where UserID=" & GBL_ID,1),0)
	End If
	If Rs.Eof Then
		Online_OnlineFlag = 0
	Else
		Online_OnlineFlag = 1
		Online_LastDoingTime = Rs(0)
		Online_StartTime = Rs(1)
		Online_IP = Rs(2)
		AtUrl = Rs(3)
		AtInfo = Rs(4)
		GBL_ID = cCur(Rs(5))
		SessionID = Rs(6)
		Browser = Rs(7)
		System = Rs(8)
		OlID = cCur(Rs(9))
		OlUserName = Rs(10)
	End If
	Rs.Close
	Set Rs = Nothing
	If Browser = "" Then Browser = "未知"
	If System = "" Then System = "未知"

	Dim ShowFlag

	Set Rs = LDExeCute(sql_select("Select UserName,ShowFlag,UserLevel,IP,ID,Login_okNum,Login_RightIP from LeadBBS_User where id=" & GBL_ID,1),0)
	If Online_OnlineFlag = 0 Then
		If Rs.Eof Then
			GBL_CHK_TempStr = "用户不存在<br>" & VbCrLf
			LookUserInfo = 0
			GBL_CHK_Flag = 0
			Rs.Close
			Set Rs = Nothing
			Exit Function
		End If
	End If

	If Not Rs.Eof Then
		GBL_Name = Rs(0)
		ShowFlag = Rs(1)
		REM 特殊数据
		Form_IP = Rs(3)
		Form_ID = Rs(4)
		Form_Login_oknum = Rs(5)
		Login_RightIP = Rs(6)
		If (ccur(ShowFlag) = 1) and DEF_EnableUserHidden = 1 and (GBL_NoneLimitFlag = 0) Then Online_OnlineFlag = 0
	Else
		GBL_Name = ""
		ShowFlag = 1
		Form_IP = "0.0.0.0"
		Form_ID = 0
		Form_Login_oknum = 0
	End If
	Rs.Close
	Set Rs = Nothing

	Set Rs = LDExeCute(sql_select("Select UserLevel from LeadBBS_User where ID=" & GBL_UserID,1),0)
	If Not Rs.Eof Then
		LookUserLevel = cCur(Rs(0))
	Else
		LookUserLevel = 0
	End If
	Rs.Close
	Set Rs = Nothing
	
	If GBL_UserID = GBL_ID or GBL_NoneLimitFlag = 1 Then LookUserLevel=15

	Dim Old_GBL_CHK_User
	Old_GBL_CHK_User = GBL_CHK_User
	GBL_CHK_User = GBL_Name
	If GBL_NoneLimitFlag = 1 Then
		'Form_IP = "218.53.238.111"
		'Online_IPAddress = "218.53.238.111"
		'Online_IP = "218.53.238.111"
	End If
	GBL_CHK_User = Old_GBL_CHK_User
	
	If GBL_CHK_User = GBL_Name Then GBL_NoneLimitFlag = 1 '如果是查看自己的信息无限制
	If (GBL_ID > 0 and (OlID=0 or LookUserLevel >= 15)) or OlUserName = GBL_Name Then
		LookUserInfo_NavInfo
	Else
		GBL_Name = OlUserName
	End If
	%>
	<div class=title>更多关于<%
	If (GBL_ID = 0 or GBL_NoneLimitFlag = 0) and OlID > 0 Then
		Response.Write "第[" & OlID & "]号在线人员"
	Else
		Response.Write "用户[<span class=redfont>" & htmlencode(GBL_Name) & "</span>]"
	End If%>的信息</div>
			<table border=0 cellpadding="0" cellspacing="0" class=blanktable><%
			If GBL_ID > 0 and (OlID=0 or LookUserLevel >= 15) Then%>
			<tr>
				<TD class=tdbox>
					用户编号：
				</td>
				<td>
					<a href=LookUserInfo.asp?ID=<%=GBL_ID%>><%=GBL_ID%></a></td>
			</tr>
			<tr>
				<td>
					用户名：
				</td>
				<td>
					<%=Server.HtmlEncode(GBL_Name)%></td>
			</tr><%end If
			If Online_OnlineFlag=1 Then%>
			<tr>
				<td>
					浏览器及系统：
				</td>
				<td>
					<%=Browser%> / <%=System%></td>
			</tr><%End If
			If GBL_ID > 0 and (Online_OnlineFlag > 1 or GBL_NoneLimitFlag = 1) Then%>
			<tr>
				<td>
					登录次数：
				</td>
				<td>
					<%If LookUserLevel>=4 Then
						Response.Write Form_Login_oknum & "次"
					Else
						Response.Write "需要5级用户才能查看"
					End If%></td>
			</tr><%
			End If
			If LookUserLevel>=0 Then%>
			<tr>
				<td>
					是否在线：
				</td>
				<td>
					<%
					If Online_OnlineFlag=1 Then
						Response.Write "是"
					else
						Response.Write "否或隐身"
					End If%></td>
			</tr><%If Online_OnlineFlag=1 Then%>
			<tr>
				<td>
					当前浏览网页：
				</td>
				<td>
					<a href="<%=AtUrl%>"><%=AtInfo%></a></td>
			</tr><%End If
			End If%>
			<tr>
				<td>
					在线登录时间：
				</td>
				<td>
					<%			
					If LookUserLevel>=9 Then
						If Online_OnlineFlag=1 Then
							Response.Write RestoreTime(Online_StartTime)
						else
							Response.Write "</font>离线或隐身</font>"
						end If
					Else
						Response.Write "需要10级用户才能查看"
					End If%></td>
			</tr><%
					If LookUserLevel>=15 Then%>
			<tr>
				<td>
					最后动作时间：
				</td>
				<td>
					<%
						If Online_OnlineFlag=1 Then
							Response.Write RestoreTime(Online_LastDoingTime)
						else
							Response.Write "离线或隐身"
						end If
					'Else
					'	Response.Write "需要7级用户才能查看"%></td>
			</tr><%
			End If
			'If GBL_NoneLimitFlag = 0 and Online_OnlineFlag=1 Then
			'	Form_IP=Online_IP
			'End If
			
			Dim Online_IPAddress,Form_IPAddress
			Form_IPAddress = GetIPAddressData(Form_IP,LookUserLevel)
			If GBL_NoneLimitFlag = 1 or GBL_UserID = GBL_ID Then
			Else
				Form_IPAddress = ""
			End If
			
			
			If GBL_NoneLimitFlag = 1 Then Online_IPAddress = GetIPAddressData(Online_IP,LookUserLevel)
			If GBL_NoneLimitFlag = 1 Then
			Else
				Online_IPAddress = ""
				Online_IP = ""
			End If
			%>
			<%If GBL_NoneLimitFlag = 1 and Online_OnlineFlag=1 Then
			'<tr>
			'	<td>
			'		Session标识：
			'	</td>
			'	<td>
			'		SessionID</td>
			'</tr>
			%>
			<tr>
				<td>
					在线IP地址：
				</td>
				<td>
					<%=Online_IP%><br>
					<%=Online_IPAddress%></td>
			</tr><%End If
			If GBL_ID = 0 Then
				Form_IP = "游客无注册IP地址"
				Form_IPAddress = "游客无注册地理位置"
			End If
			If GBL_NoneLimitFlag = 1 Then%>
			<tr>
				<td>
					注册IP地址：
				</td>
				<td>
					<%=Form_IP%>
					<br>
					<%=Form_IPAddress%></td>
			</tr><%End If

			If Login_RightIP = "3u7s9_d9299Xls" Then Login_RightIP = "无"
			If GBL_NoneLimitFlag = 1 Then%>
			<tr>
				<td>
					登录IP地址：
				</td>
				<td>
					<%=Login_RightIP%>
					<br>
					<%
				If Login_RightIP <> "" and Login_rightIP <> "无" Then Response.Write GetIPAddressData(Login_RightIP,LookUserLevel)%>
				</td>
			</tr>
			<%End If%>
			</table>
	<%

End Function

Function GetIPAddressData(IP,LookUserLevel)

	Dim sip,num
	Dim str1,str2,str3,str4
	if ip<>"" then
		sip=ip
		If inStr(sip,".") = 0 Then Exit Function
		str1=left(sip,instr(sip,".")-1)
		sip=mid(sip,instr(sip,".")+1)
		If inStr(sip,".") = 0 Then Exit Function
		str2=left(sip,instr(sip,".")-1)
		sip=mid(sip,instr(sip,".")+1)
		If inStr(sip,".") = 0 Then Exit Function
		str3=left(sip,instr(sip,".")-1)
		str4=mid(sip,instr(sip,".")+1)
		if isNumeric(str1)=0 or isNumeric(str2)=0 or isNumeric(str3)=0 or isNumeric(str4)=0 then
		else
			num=cint(str1)*16777216+cint(str2)*65536+cint(str3)*256+cint(str4)-1
		end if
	else
		ip="0.0.0.0"
		num=0
		str1="0.0"
	End If
	If GBL_NoneLimitFlag = 0 or LookUserLevel < 1 Then Exit Function
	Dim Rs
	'Set Rs = LDExeCute(sql_select("select country,city,ip2 from LeadBBS_IPAddress where ip1 <=" & num & " and ip2 >=" & num,1),0)
	Set Rs = LDExeCute(sql_select("select country,city,ip2 from LeadBBS_IPAddress where ip1 <=" & num & " order by ip1 DESC",3),0)

	GetIPAddressData = ""
	Do While Not Rs.Eof
		If cCur(Rs(2)) >= num Then
			GetIPAddressData = Rs("Country") & " " & rs("city")
			Exit Do
		End If
		Rs.MoveNext
	Loop
	Rs.Close
	Set Rs = Nothing
	If GetIPAddressData = "" Then GetIPAddressData="未知"

End Function

Sub DisplayFriend

	Dim Rs,SQL

	Dim UpDownPageFlag
	UpDownPageFlag = Request("UpDownPageFlag")

	Dim Start

	Dim SQLendString

	Start = Left(Trim(Request("Start")),14)
	Start = start
	If isNumeric(Start)=0 or Start="" Then Start=999999999
	Start = cCur(Start)
	If Start = 0 Then Start=999999999


	Dim NotSecret

	If GBL_UserID <> GBL_ID Then
		Set Rs = LDExeCute(sql_select("Select UserName,NotSecret from LeadBBS_User where ID=" & GBL_ID & " Order by ID ASC",1),0)
		If Rs.Eof Then
			GBL_Name = ""
		Else
			GBL_Name = Rs(0)
			NotSecret = Rs(1)
			If ccur(NotSecret) = 1 Then
				NotSecret = "1"
			Else
				NotSecret = "0"
			End If
		End if
		Rs.Close
		Set Rs = Nothing
		If GBL_Name = "" Then
			GBL_ID = 0
		End If
	
		If NotSecret = "0" and GBL_CHK_User <> GBL_Name Then
			LookUserInfo_NavInfo
			Response.Write "<div class=alert>此用户已设置隐私资料保密。</div>"
			Exit Sub
		End If
	End If
	LookUserInfo_NavInfo
	
	Dim SelfFlag
	If GBL_CHK_User <> GBL_Name Then
		SelfFlag = 0
	Else
		SelfFlag = 1
	End If

	Dim FirstID,LastID
	Dim SQLCountString,whereFlag
	Dim MaxRecordID,MinRecordID
	If Request.QueryString("need") <> "23" Then
		whereFlag = 1
		If Evol = "uf" Then
			SQLendString = " where T1.FriendUserID=" & GBL_ID
		Else
			SQLendString = " where T1.UserID=" & GBL_ID
		End If
	
		SQLCountString = SQLendString
		If UpDownPageFlag = "1" and Start>0 then
			If whereFlag = 1 Then
				SQLendString = SQLendString & " and T1.ID>" & Start
			Else
				SQLendString = SQLendString & " where T1.ID>" & Start
				whereFlag = 1
			End If
		Else
			If whereFlag = 1 Then
				SQLendString = SQLendString & " and T1.ID<" & Start
			Else
				SQLendString = SQLendString & " where T1.ID<" & Start
				whereFlag = 1
			End If
		end If
	
		If UpDownPageFlag = "1" then
			SQLendString = SQLendString & " Order by T1.ID ASC"
		Else
			SQLendString = SQLendString & " Order by T1.ID DESC"
		End If
		
		MaxRecordID = 0
	
		SQL = "select Max(T1.id) from LeadBBS_FriendUser as T1 " & SQLCountString
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
		
		SQL = "select Min(T1.id) from LeadBBS_FriendUser as T1 " & SQLCountString
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
	
		If Evol = "uf" Then
			SQL = sql_select("select T2.UserName,T2.Mail,T2.Sex,T2.LastDoingTime,T2.Userphoto,T2.FaceUrl,T2.FaceWidth,T2.FaceHeight,T1.ID from LeadBBS_FriendUser as T1 left join LeadBBS_User as T2 on T1.UserID=T2.ID " & SQLendString,DEF_MaxListNum)
		Else
			SQL = sql_select("select T2.UserName,T2.Mail,T2.Sex,T2.LastDoingTime,T2.Userphoto,T2.FaceUrl,T2.FaceWidth,T2.FaceHeight,T1.ID from LeadBBS_FriendUser as T1 left join LeadBBS_User as T2 on T1.FriendUserID=T2.ID " & SQLendString,DEF_MaxListNum)
		End If
	Else
		SQL = "select T2.UserName,T2.Mail,T2.Sex,T2.LastDoingTime,T2.Userphoto,T2.FaceUrl,T2.FaceWidth,T2.FaceHeight,T1.ID from (LeadBBS_FriendUser as T1 right join LeadBBS_onlineUser as T3 on T1.FriendUserID=T3.UserID) right join LeadBBS_User as T2 on T3.UserID=T2.ID where T1.UserID=" & GBL_UserID
	End If
	Set Rs = LDExeCute(SQL,0)
	Dim Num
	Dim GetData
	If Not rs.Eof Then
		GetData = Rs.GetRows(DEF_MaxListNum)
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
		
		LastID = cCur(GetData(8,MaxN))
		FirstID = cCur(GetData(8,MinN))
	
		Dim EndwriteQueryString,PageSplictString
		EndwriteQueryString = "?Evol=" & Evol & "&ID=" & GBL_ID
		If Request.QueryString("need") = "23" Then EndwriteQueryString = EndwriteQueryString & "&need=23"
	
		PageSplictString = PageSplictString & "<div class=j_page>"
		If FirstID >= MaxRecordID Then
			'PageSplictString = PageSplictString & "首页" & VbCrLf
			'PageSplictString = PageSplictString & " 上页" & VbCrLf
		else
			PageSplictString = PageSplictString & "<a href=LookUserInfo.asp" & EndwriteQueryString & "&Start=0>首页</a> " & VbCrLf
			PageSplictString = PageSplictString & " <a href=LookUserInfo.asp" & EndwriteQueryString & "&Start=" & FirstID & "&UpDownPageFlag=1>上页</a> " & VbCrLf
		end if
	
		if LastID<MaxRecordID and LastID<>0 then
		else
		end if
	
		If LastID <= MinRecordID Then
			'PageSplictString = PageSplictString & " 下页" & VbCrLf
			'PageSplictString = PageSplictString & " 尾页" & VbCrLf
		else
			PageSplictString = PageSplictString & " <a href=LookUserInfo.asp" & EndwriteQueryString & "&Start=" & LastID & ">下页</a> " & VbCrLf
			PageSplictString = PageSplictString & " <a href=LookUserInfo.asp" & EndwriteQueryString & "&Start=1&UpDownPageFlag=1>尾页</a> " & VbCrLf
		end if
		
		'PageSplictString = PageSplictString & "共<b>" & recordCount & "好友</b>"
		'If (recordCount mod DEF_MaxListNum)=0 Then
		'	PageSplictString = PageSplictString & " 计<b>" & clng(recordCount/DEF_MaxListNum) & "</b>页"
		'Else
		'	If recordCount>=DEF_MaxListNum Then
		'		PageSplictString = PageSplictString & " 计<b>" & clng(recordCount/DEF_MaxListNum)+1 & "</b>页"
		'	Else
		'		PageSplictString = PageSplictString & " 计<b>1</b>页"
		'	End If
		'End If
		'PageSplictString = PageSplictString & " 每页<b>" & DEF_MaxListNum & "</b>个"
		PageSplictString = PageSplictString & "</div>"
	
	End If
	
	Dim colNum
	colNum = 2
	%>
	
	<script src="<%=DEF_BBS_HomeUrl%>inc/js/p_list.js?ver=20090601.2" type="text/javascript"></script>
	<script type="text/javascript">
		p_url = "DeleteMessage.asp";
		p_para = "AjaxFlag=1&FriendFlag=1&DeleteSureFlag=dk9@dl9s92lw_SWxl&MessageID=";
		p_command = 'alert(tmp);this.location="LookUserInfo.asp?Evol=f";';
		p_type = 1;
		function killall(str)
		{
			//window.open('DelFriend.asp?kasdie=3&ClearFlag='+str,'','width=450,height=37,scrollbars=auto,status=no');
			if (confirm('删除操作将不可逆,确定继续吗?'))
			p_once("&ClearFlag="+str,1);
		}
	</script>
	<table width="100%" border="0" cellspacing="0" cellpadding="0" class=table_in>
	  <tr class=tbinhead>
	    <td width=<%=DEF_AllFaceMaxWidth + 30%> align=center><div class=value>用户</div></td>
	    <td><div class=value>信息</div></td><%If GBL_CHK_User = GBL_Name and Evol = "f" Then
	    	colNum = 3%>
	    <td width=80><div class=value>删除</div></td><%End If%>
	  </tr>
	<%
	If Num = -1 Then
		response.write "<tr><td colspan=" & colNum & " class=tdbox>没有好友信息。</td></tr>"
	end if

	Dim TempN,Temp,Temp1
	
	Dim Index
	Index = 0
	if Num <> -1 then
		i=1
		LastID = GetData(0,ubound(getdata,2))
		For n = MinN to MaxN Step StepValue
			Response.Write "<tr><td align=center class=tdbox>"
			If DEF_AllDefineFace = 0 or GetData(5,N) & "" = "" Then
				If GetData(4,N)<>"" and isNumeric(GetData(4,N)) Then
					Response.Write "<img src=../images/face/" & string(4-len(cstr(GetData(4,N))),"0")&GetData(4,N) & ".gif align=middle>"
				Else
					Response.Write "<img src=../images/null.gif align=middle>"
				End If
			Else
				Response.Write "<img src=""" & htmlencode(GetData(5,N)) & """ align=middle width=" & GetData(6,N) & " height=" & GetData(7,N) & ">"
			End If
			Response.Write "<a href=LookUserinfo.asp?name=" & htmlencode(GetData(0,n)) & "><div class=user>" & htmlencode(GetData(0,n)) & "</div></a>"

			Response.Write "</td><td class=tdbox>"
			Response.Write "<ul>"
			If SelfFlag = 1 Then Response.Write "<li>邮箱：<a href=""mailto:." & htmlencode(GetData(1,n)) & """>" & htmlencode(GetData(1,n)) & "</a></li>"
			Response.Write "<li>最后活动: " & RestoreTime(GetData(3,n)) & "</li>"
			Response.Write "<li><a href=""SendMessage.asp?SdM_ToUser=" & htmlencode(GetData(0,n)) & """><img src=../images/" & GBL_DefineImage & "message.GIF border=0 title=给" & htmlencode(GetData(0,n)) & "发消息 align=middle>发送短消息</a></li>"
			If GBL_CHK_User = GBL_Name and Evol = "f" Then
				Response.Write "</td><td class=tdbox>"
				%>
				<input class="fmchkbox" type="checkbox" name="ids" id="ids<%=Index%>" value="<%=urlencode(GetData(8,n))%>" /><%
				Response.Write "<a href='javascript:p_once(" & urlencode(GetData(8,n)) & ");'>删除</a>"
				Index = Index + 1
			End If
			Response.Write "</td></tr>" & VbCrLf
			i=i+1
		next
	End If
	Response.Write "<tr><td colspan=" & colNum & " class=tdbox>" & PageSplictString
	%></td></tr>
	
	<tr><td colspan=<%=colNum%> class=tdbox align=right>
	<input class="fmchkbox" type="checkbox" name="selmsg" id="selmsg" value="1" onclick="achoose();" />选择所有记录
	<input type=button value="批量删除" onclick="pchoose();" class="fmbtn btn_4">
	</td></tr>
	</table>
	<br>
	<%If GBL_CHK_User = GBL_Name Then%>
	<div class=value2>	
	<a href='../a/Processor.asp?action=AddFriend&b=0&ID=0&FriendName=' onclick="return(pub_msg(this,'anc_msgbody','&Dir=<%=DEF_BBS_HomeUrl%>'));" target=_blank>添加好友</a>
	<a href='javascript:killall("dkeJje5");'><img src=../images/<%=GBL_DefineImage%>clear.gif width=16 border=0 align=middle>清空好友</a>
	</div>
	<div class=value2>
		<%If Request.QueryString("need") = "23" Then%>
		<a href=LookUserInfo.asp?Evol=f>查看我的好友</a>
		<%Else%>
		<a href=LookUserInfo.asp?Evol=f&need=23>查看我的在线好友</a>
		<%End If%>
	</div><%
	End If%>
	<div class=value2>
	<%If Evol = "uf" Then%>
		<a href=LookUserInfo.asp?Evol=f>查看<b><%=htmlencode(GBL_Name)%></b>的好友</a>
	<%Else%>
		<a href="LookUserInfo.asp?Evol=uf&id=<%=urlencode(GBL_ID)%>">查看添加<b><%=htmlencode(GBL_Name)%></b>为好友的用户</a>
	<%End If%>
	</div>
	<%
	If GBL_UserID>0 and CheckSupervisorUserName = 1 Then
		%>
		<hr class=splitline>
		<a href='javascript:killall("dkeJje6");'><img src=../images/<%=GBL_DefineImage%>clear.gif width=16 border=0 align=middle title=添空所有人的好友列表>清空全部好名名单(限管理员)</a><%
	End If

End Sub

Sub DisplayFavorite

	Dim Rs,SQL

	Dim NotSecret

	If GBL_UserID <> GBL_ID Then
		If GBL_ID = 0 Then
			If GBL_Name = "" Then
				Exit Sub
			Else
				Set Rs = LDExeCute(sql_select("Select UserName,NotSecret,ID from LeadBBS_User where UserName='" & Replace(GBL_Name,"'","''") & "'",1),0)
			End if
		Else
			Set Rs = LDExeCute(sql_select("Select UserName,NotSecret,ID from LeadBBS_User where ID=" & GBL_ID & " Order by ID ASC",1),0)
		End If
		If Rs.Eof Then
			GBL_Name = ""
		Else
			GBL_ID = cCur(Rs(2))
			GBL_Name = Rs(0)
			NotSecret = Rs(1)
			If ccur(NotSecret) = 1 Then
				NotSecret = "1"
			Else
				NotSecret = "0"
			End If
		End if
		Rs.Close
		Set Rs = Nothing
		If GBL_Name = "" Then
			GBL_ID = 0
		End If

		If NotSecret = "0" and GBL_CHK_User <> GBL_Name Then
			LookUserInfo_NavInfo
			Response.Write "<div class=alert>此用户已设置隐私资料保密。</div>"
			Exit Sub
		End If
	End If
	LookUserInfo_NavInfo

	Dim UpDownPageFlag
	UpDownPageFlag = Request("UpDownPageFlag")

	Dim Start,key

	Dim SQLendString
	Dim FirstID,LastID,RecordCount

	Start = Left(Trim(Request("Start")),14)
	Start = start
	If isNumeric(Start)=0 or Start="" Then Start=999999999
	Start = cCur(Start)
	If Start = 0 Then Start=999999999

	Dim SQLCountString,whereFlag
	whereFlag = 1
	SQLendString = " where T1.UserID=" & GBL_ID

	SQLCountString = SQLendString
	If UpDownPageFlag = "1" and Start>0 then
		If whereFlag = 1 Then
			SQLendString = SQLendString & " and T1.ID>" & Start
		Else
			SQLendString = SQLendString & " where T1.ID>" & Start
			whereFlag = 1
		End If
	Else
		If whereFlag = 1 Then
			SQLendString = SQLendString & " and T1.ID<" & Start
		Else
			SQLendString = SQLendString & " where T1.ID<" & Start
			whereFlag = 1
		End If
	End If

	If UpDownPageFlag = "1" then
		SQLendString = SQLendString & " Order by T1.ID ASC"
	Else
		SQLendString = SQLendString & " Order by T1.ID DESC"
	End If
	
	Dim MaxRecordID,MinRecordID
	MaxRecordID = 0

	SQL = "select count(*) from LeadBBS_CollectAnc as T1 " & SQLCountString
	Set Rs = LDExeCute(SQL,0)
	If Rs.Eof then
		RecordCount=0
	Else
		RecordCount = rs(0)
		If RecordCount="" or isNull(RecordCount) or len(RecordCount)<1 Then RecordCount=0
		RecordCount = ccur(RecordCount)
	End If
	Rs.Close
	Set Rs = Nothing

	If RecordCount > 0 Then
		SQL = "select Max(T1.id) from LeadBBS_CollectAnc as T1 " & SQLCountString
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
		
		SQL = "select Min(T1.id) from LeadBBS_CollectAnc as T1 " & SQLCountString
		Set Rs = LDExeCute(SQL,0)
	
		If not Rs.Eof Then
			If Rs(0) <> "" Then
				MinRecordID = cCur(Rs(0))
			Else
				MinRecordID = 0
			End If
		End If
		Rs.Close
		Set Rs = Nothing
	
		SQL = sql_select("select T1.ID,T2.Title,T2.Length,T2.ndatetime,T2.Hits,T2.FaceIcon,T2.ChildNum,T2.BoardID,T2.GoodFlag,T2.Username,T2.ID,T2.TitleStyle from LeadBBS_CollectAnc as T1 Left join LeadBBS_Announce as T2 on T1.AnnounceID=T2.ID " & SQLendString,DEF_MaxListNum)
		Set Rs = LDExeCute(SQL,0)
		Dim Num
		Dim GetData
		If Not rs.Eof Then
			GetData = Rs.GetRows(DEF_MaxListNum)
			Num = Ubound(GetData,2)
		Else
			Num = -1
		End If
		Rs.close
		Set Rs = Nothing
	Else
		Num = -1
		MinRecordID = 0
		MaxRecordID = 0
	End If

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
	
		Dim EndwriteQueryString,PageSplictString
		EndwriteQueryString = "?Evol=bag&ID=" & GBL_ID
		If key<>"" Then EndwriteQueryString = EndwriteQueryString & "&key=" & urlencode(key)
	
		PageSplictString = "<div class=j_page>"
		If FirstID >= MaxRecordID Then
			'PageSplictString = PageSplictString & "首页" & VbCrLf
			'PageSplictString = PageSplictString & " 上页" & VbCrLf
		Else
			PageSplictString = PageSplictString & "<a href=LookUserInfo.asp" & EndwriteQueryString & "&Start=0>首页</a> " & VbCrLf
			PageSplictString = PageSplictString & " <a href=LookUserInfo.asp" & EndwriteQueryString & "&Start=" & FirstID & "&UpDownPageFlag=1>上页</a> " & VbCrLf
		End If
	
		If LastID<MaxRecordID and LastID<>0 then
		Else
		End If
	
		If LastID <= MinRecordID Then
			'PageSplictString = PageSplictString & " 下页" & VbCrLf
			'PageSplictString = PageSplictString & " 尾页" & VbCrLf
		Else
			PageSplictString = PageSplictString & " <a href=LookUserInfo.asp" & EndwriteQueryString & "&Start=" & LastID & ">下页</a> " & VbCrLf
			PageSplictString = PageSplictString & " <a href=LookUserInfo.asp" & EndwriteQueryString & "&Start=1&UpDownPageFlag=1>尾页</a> " & VbCrLf
		End If

		PageSplictString = PageSplictString & "<b>共" & RecordCount & "</b>"
		'If (RecordCount mod DEF_MaxListNum)=0 Then
		'	PageSplictString = PageSplictString & " 计<b>" & clng(RecordCount/DEF_MaxListNum) & "</b>页"
		'Else
		'	If RecordCount>=DEF_MaxListNum Then
		'		PageSplictString = PageSplictString & " 计<b>" & clng(RecordCount/DEF_MaxListNum)+1 & "</b>页"
		'	Else
		'		PageSplictString = PageSplictString & " 计<b>1</b>页"
		'	End If
		'End If
		'PageSplictString = PageSplictString & " 每页<b>" & DEF_MaxListNum & "</b>条收藏帖"
		PageSplictString = PageSplictString & "</div>"
	
	End If
	%>
	
	<script src="<%=DEF_BBS_HomeUrl%>inc/js/p_list.js?ver=20090601.2" type="text/javascript"></script>
	<script type="text/javascript">
		p_url = "DeleteMessage.asp";
		p_para = "AjaxFlag=1&FriendFlag=2&DeleteSureFlag=dk9@dl9s92lw_SWxl&MessageID=";
		p_command = 'alert(tmp);this.location="LookUserInfo.asp?Evol=bag";';
		p_type = 1;
		function killall(str)
		{
			if (confirm('删除操作将不可逆,确定继续吗?'))
			p_once("&ClearFlag="+str,1);
		}
	</script>
	<table width="100%" border="0" cellspacing="0" cellpadding="0" class=table_in>
	  <tr class=tbinhead>
	    <td><div class=value>主题</div></td>
	    <td width=80><div class=value>人气</div></td>
	    <td width=210><div class=value>发表时间/作者</div></td><%
	    Dim ColNum
	    ColNum = 3
	    If GBL_ID = GBL_UserID Then
	    	ColNum = 4%>
	    <td width=80><div class=value>删除</div></td><%
	    End If%>
	  </tr>
	<%
	If Num = -1 Then
		response.write "<tr><td colspan=" & ColNum & " class=tdbox>暂无收藏!</td></tr>"
	End If

	Dim TempN,Temp,Temp1
	
	Dim Index
	Index = 0
	If Num <> -1 then
		i=1
		LastID = GetData(0,ubound(getdata,2))
		For n= MinN to MaxN Step StepValue
			If isNull(GetData(6,N)) Then
				GetData(1,n) = "<span class=grayfont>该收藏帖已经不存在(原编号" & GetData(0,n) & ")，已经被管理员删除。</span>"
				GetData(0,n) = 0
				GetData(2,n) = 0
				GetData(3,n) = "19000101000000"
				GetData(4,n) = 0
				GetData(5,n) = 0
				GetData(6,n) = 0
				GetData(7,n) = 0
				GetData(8,n) = 0
				GetData(9,n) = "游客"
				GetData(10,n) = ""
				GetData(11,n) = 1
			Else
				GetData(0,n) = cCur(GetData(0,n))
			End If
			Response.Write "<tr><td class=tdbox>"
			If GetData(0,n) > 0 Then Response.Write "<a href=../a/a.asp?B=" & GetData(7,n) & "&ID=" & GetData(10,N) & ">"

			GetData(6,N) = cCur(GetData(6,N))
			Temp1 = Fix((GetData(6,N)+1)/DEF_TopicContentMaxListNum)
			If ((GetData(6,N)+1) mod DEF_TopicContentMaxListNum) > 0 Then Temp1 = Temp1 + 1
			If GetData(6,N)>=DEF_TopicContentMaxListNum Then
				Temp = DEF_BBS_DisplayTopicLength - (Len(Temp1) + 3)
			Else
				Temp = DEF_BBS_DisplayTopicLength
			End If
		
			If ccur(GetData(8,n)) = 1 Then Temp = Temp - 3
			
			If GetData(11,n) <> 1 and strLength(GetData(1,N))>Temp-1 Then GetData(1,N) = LeftTrue(GetData(1,N),Temp-4) & "..."
			Response.Write DisplayAnnounceTitle(GetData(1,n),GetData(11,n))
			If GetData(0,n) > 0 Then Response.Write "</a>"

			If GetData(6,N)>=DEF_TopicContentMaxListNum Then
				Response.Write " [<a href=../a/a.asp?B=" & GetData(7,N) & "&ID=" & GetData(10,N) & "&AUpflag=1&ANum=1 title=" & GetData(2,n) & "字节>" & Temp1 & "</b></a>]"
			End If

			If ccur(GetData(8,n)) = 1 Then
				Response.Write "<img src=../images/" & GBL_DefineImage & "jh1.GIF border=0 title=精华帖子 align=absbottom>"
			End If
			Response.Write "</td><td class=tdbox><em>"
			Response.Write GetData(6,N) & "/" & GetData(4,N)
			Response.Write "</em></td><td class=tdbox><em>"
			If GetData(9,n) <> "游客" then
				Response.Write Left(RestoreTime(GetData(3,n)),16) & "</em> <a href=LookUserInfo.asp?name=" & urlencode(GetData(9,n)) & ">" & htmlencode(GetData(9,n)) & "</a></td>"
			Else
				Response.Write Left(RestoreTime(GetData(3,n)),16) & "</em> " & htmlencode(GetData(9,n)) & "</td>"
			End If
			If GBL_ID = GBL_UserID Then
				%>
				<td class=tdbox>
				<input class="fmchkbox" type="checkbox" name="ids" id="ids<%=Index%>" value="<%=GetData(0,n)%>" /><%
				Response.Write "<a href='javascript:p_once(" & GetData(0,n) & ");'>删除</a>"
				Index = Index + 1
				%>
				</td>
				<%
			End If
			Response.Write "</td></tr>" & VbCrLf
			i=i+1
		Next
	End If
	Response.Write "<tr><td colspan=" & ColNum & " class=tdbox>" & PageSplictString
	%>
		</td></tr>
	<%If GBL_ID = GBL_UserID Then%>
	<tr><td colspan=<%=ColNum%> class=tdbox align=right>
	<input class="fmchkbox" type="checkbox" name="selmsg" id="selmsg" value="1" onclick="achoose();" />选择所有记录
	<input type=button value="批量删除" onclick="pchoose();" class="fmbtn btn_4">
	</td></tr>
	<tr><td colspan=<%=ColNum%> class=tdbox align=right>
	<div class=value2>
	<a href='javascript:killall("dkeJje5");'><img src=../images/<%=GBL_DefineImage%>clear.gif width=16 border=0 align=middle>清空我的收藏夹</a>
	<%	If GBL_UserID>0 and CheckSupervisorUserName = 1 Then%><a href='javascript:killall("dkeJje6");'><img src=../images/<%=GBL_DefineImage%>clear.gif width=16 border=0 align=middle>清空所有人的收藏夹(管理员)</a><%End If%>
	</div>
	</td></tr>
	<%End If%>
	</table><%

End Sub

Sub LookUserInfo_NavInfo

	Response.Write "<div class='user_item_nav fire'><ul>"
	Response.Write "<li><div class=name>" & htmlencode(GBL_Name) & "</div></li>"
	If Evol = "A" or Evol = "" Then
		Response.Write "	<li><div class=navactive><span>基本资料</span></div></li>"
	Else
		Response.Write "	<li><a href=LookUserInfo.asp?ID=" & GBL_ID & "&Evol=A>基本资料</a></li>"
	End If
	If Evol = "n" Then
		Response.Write "	<li><div class=navactive>所有帖子</div></li>"
	Else
		Response.Write "	<li><a href=LookUserInfo.asp?ID=" & GBL_ID & "&Evol=n>所有帖子</a></li>"
	End If
	If Evol = "g" Then
		Response.Write "	<li><div class=navactive>主题帖子</div></li>"
	Else
		Response.Write "	<li><a href=LookUserInfo.asp?ID=" & GBL_ID & "&Evol=g>主题帖子</a></li>"
	End If
	If Evol = "e" Then
		Response.Write "	<li><div class=navactive>精华帖子</div></li>"
	Else
		Response.Write "	<li><a href=LookUserInfo.asp?ID=" & GBL_ID & "&Evol=e>精华帖子</a></li>"
	End If
	If Evol = "l" Then
		Response.Write "	<li><div class=navactive>上传附件</div></li>"
	Else
		Response.Write "	<li><a href=LookUserInfo.asp?ID=" & GBL_ID & "&Evol=l>上传附件</a></li>"
	End If
	If Evol = "f" or Evol = "uf" Then
		Response.Write "	<li><div class=navactive>好友信息</div></li>"
	Else
		Response.Write "	<li><a href=LookUserInfo.asp?ID=" & GBL_ID & "&Evol=f>好友信息</a></li>"
	End If
	If Evol = "bag" Then
		Response.Write "	<li><div class=navactive>收藏夹</div></li>"
	Else
		Response.Write "	<li><a href=LookUserInfo.asp?ID=" & GBL_ID & "&Evol=bag>收藏夹</a></li>"
	End If
	Response.Write "</ul></div>"
	

End Sub
%>