<!-- #include file=../inc/BBSsetup.asp -->
<!-- #include file=../inc/User_Setup.ASP -->
<!-- #include file=../inc/Board_Popfun.asp -->
<!-- #include file=../inc/Limit_Fun.asp -->
<!-- #include file=inc/Poll_fun.asp -->
<!-- #include file=../inc/Constellation.asp -->
<!-- #include file=../inc/AD_Fun.asp -->
<!-- #include file=../inc/UBBCode_Setup.asp -->
<%
Const LMT_RefreshEnable = 0 '用户重复浏览帖子是否计算浏览量
Const LMT_ViewTopicOpinion = 1 '重复直接显示主题评价列表
Const LMTDEF_RepostMsg = 0 '回复帖子是否默认短消息通知帖主,0．默认不通知 1.回复全部通知(注意与回帖(a2.asp文件)设置保持一直)
Const LMTDEF_ConvetType = 0 'ubb编码转换方式：0,客户端JS转换,1.leadbbs dll组件服务端转换 2.vbscript服务端转换
Const LMTDEF_ShareID = "<a class=""bshareDiv"" href=""http://www.bshare.cn/share""></a><script type=""text/javascript"" charset=""utf-8"" src=""http://static.bshare.cn/b/buttonLite.js#uuid=0e896bee-a954-4993-a6e5-f56844e9c9fc&style=2""></script>" '可以填写各站或自行编写类型的分享代码(HTML格式，注意手工删除换行符),保持为空则关闭分享代码;

Dim LMTDEF_ShareID_Exist
LMTDEF_ShareID_Exist = ""

DEF_BBS_HomeUrl = "../"
Dim A_NotReplay,Page
A_NotReplay = 0
Dim A_ID
Dim A_RootMinID,A_RootMaxID,A_ChildNum,A_ParentID,A_Data
A_RootMaxID = 0
A_RootMinID = 0
A_ChildNum = 0

Dim A_Title,A_TitleStyle,A_TitleNoHTML,A_RootID,A_RootIDBak,Form_TopicType
A_RootID = 0
A_RootIDBak = 0

Dim A_BoardUrl,A_BoardStr

Dim R_ID,R_NDatetime
R_ID = 0

Dim LMT_EnableRewrite
LMT_EnableRewrite = GetBinarybit(DEF_Sideparameter,16)

Function DisplayAnnounceForm

	GBL_CHK_TempStr = ""
	If A_ID = 0 Then
		CheckBoardAnnounceLimit
	Else
		CheckBoardReAnnounceLimit
	End If
	CheckUserAnnounceLimit
	If GBL_CHK_TempStr <> "" Then Exit Function
	If A_NotReplay = 1 Then Exit Function
	GBL_CHK_Points = cCur(GBL_CHK_Points)
	If GBL_CHK_Points < 1 Then
		If isArray(GBL_UDT) Then GBL_CHK_Points = cCur(GBL_UDT(4))
	End If
%>
<script type="text/javascript">
<!--
	var ValidationPassed = true,submitflag=0;
	function submitonce(theform)
	{	submitflag = 1;
		<%If DEF_EnableAttestNumber > 2 and (DEF_AttestNumberPoints = 0 or GBL_CHK_Points < DEF_AttestNumberPoints) Then%>
		
		if(theform.ForumNumber.value=="")
		{
			alert("请输入验证码!\n");
			ValidationPassed = false;
			theform.ForumNumber.focus();
			submitflag = 0;
			return;
		}
		<%End If%>
		
		if(typeof(edt_init)=="function")
		{
			edt_checkContent();
		}
			
		ValidationPassed = true;
		submit_disable(theform);
	}
	
	var escflag = 0;
	function edt_disabl_esc()
	{
		if(event.keyCode==27)return(false);
	}
	function edt_disablesc_sim(obj)
	{
		if(escflag==1)return;
		escflag = 1;
		obj.value = "loading...";
		obj.disabled=true;
		edt_import();
	}

	function edt_import()
	{
		$import("inc/leadedit.js?ver=20080729.32","js","",function(){
			edt_heigh = 110;
			getAJAX("a2.asp","ol=1","$id('leadeditor').innerHTML=tmp;edt_import_init();",1);
			}
			);
	}
	function edt_import_init()
	{
		edt_heigh = 110;
		edt_init();
		edt_setmode(0);
		edt_setmode(<%
	If DEF_UbbDefaultEdit = "1" Then
		Response.Write "0"
	Else
		Response.Write "1"
	End If%>);
		new LayerMenu('layer_item','layer_iteminfo');
		edt_initdone=1;
		window.onbeforeunload = function(){if(edt_getdoclen()>0&&submitflag==0)return("您的帖子未发表，确定取消吗？");}
	}
-->
</script>
		<%Global_TableHead%>
<div class="contentbox">
<!-- #include file=inc/post_layer.asp -->
		<form action="a2.asp" method="post" id="LeadBBSFm" name="LeadBBSFm" onsubmit="submitonce(this);return ValidationPassed;">
		<table border="0" cellpadding="0" cellspacing="0" width="100%" class="tablebox anc_table">
		<tr class="tbhead">
			<td><div class="value"><%
			If cCur(A_ID)=0 Then
				Response.Write "发表新帖子"
			Else
				Response.Write "回复帖子"
			End If%> 注意: *为必填项</div></td>
		</tr>
		</table>
		<table border="0" cellpadding="0" cellspacing="0" width="100%" class="tablebox anc_table">
		<%If GBL_CHK_User = "" Then%>
		<tr>
			<td width="<%=DEF_BBS_LeftTDWidth%>" class="tdleft">*验证信息</td>
			<td class="tdright">
				用户名 <input maxlength="20" name="User" value="<%=htmlencode(GBL_CHK_User)%>" size="14" class="fminpt input_2" />
				密码 <input maxlength="20" type="password" name="Pass" value="<%'=htmlencode(GBL_CHK_Pass)%>" size="14" class="fminpt input_2" />
					<a href="<%=DEF_BBS_HomeUrl%>User/<%=DEF_RegisterFile%>">注册新用户</a>
				</td>
		</tr><%End If%>
		<tr>
			<td width="<%=DEF_BBS_LeftTDWidth%>" class="tdleft">*帖子名称</td>
			<td class="tdright">
				<input name="submitflag" value="true" type="hidden" />
				<input name="LMT_DefaultEdit" value="1" type="hidden" />
				<input name="BoardID" value="<%=GBL_board_ID%>" type="hidden" />
				<input name="ID" value="<%=A_ID%>" type="hidden" />
				
				<input maxlength="255" name="Form_Title" size="55" value="<%
				If Left(A_Title,3) <> "Re:" and A_RootIDBak > 0 Then 
					Response.Write "Re:" & htmlencode(A_TitleNoHTML)
				Else
					Response.Write htmlencode(A_TitleNoHTML)
				End If%>" tabindex="50" class="fminpt input_4 notfocus" type="text" tabindex="52" /> 长度不得超过255字</td>
		</tr>
		<tr>
			<td valign="top" width="<%=DEF_BBS_LeftTDWidth%>" class="tdleft">
				内容(最大<%=Fix(DEF_MaxTextLength/1024)%>K)<br />
				<label>
				<input class="fmchkbox" type="checkbox" name="Form_HTMLFlag" value="2" checked="checked" />解析UBB代码<br />
				</label>
				<a href="<%=DEF_BBS_HomeUrl%>User/Help/Ubb.asp" target="_blank">内容支持插入UBB标签<br />使用方法请参考帮助</a>
			</td>
			<td class="tdright">
				<div id="leadeditor"><textarea cols="80" name="Form_Content" id="Form_Content" rows="6" tabindex="51" onfocus="edt_disablesc_sim(this);" class="fmtxtra"></textarea></div>
				<script type="text/javascript">
				if($id("Form_Content").disabled==true){
				$id("Form_Content").disabled=false;
				if($id("Form_Content").value=="loading...")$id("Form_Content").value = "";
				}
				</script>
			</td>
		</tr>
		<tr>
			<td class="tdleft">其它选项</td>
			<td class="tdright">
				<label>
				<input class="fmchkbox" type="checkbox" name="Form_NoUserUnderWriteFlag" value="checkbox" checked="checked" />显示签名
				</label>
				<label>
				<input class="fmchkbox" type="checkbox" name="Form_NotReplay" value="checkbox" />锁定帖子
				</label>
				<label>
				<input class=fmchkbox type="checkbox" name="Form_RepostMsg" value="checkbox"<%If LMTDEF_RepostMsg=1 Then Response.Write " checked"%>>回复短消息通知
				</label>
				<span class="grayfont">Alt+S或Ctrl+Enter快速提交</span>
				<%
				If DEF_EnableAttestNumber > 2 and (DEF_AttestNumberPoints = 0 or GBL_CHK_Points < DEF_AttestNumberPoints) Then%>
				<div style="line-height:400%">验证码
				<%
					displayVerifycode%></div><%
				End If%>
			</td>
		</tr>
		<tr>
			<td class="tdleft">&nbsp;</td>
			<td class="tdright">
			<br />
			<input name="submit2" type="submit" value="立即回复" class="fmbtn btn_3" />
			<br /><br />
			</td>
		</tr>		
		</table>
		</form>
</div>
<%
	Global_TableBottom

End Function

Function GetRequestValue

	If cStr(A_ID) = "" Then A_ID = Left(Request.QueryString("ID"),14)
	If isNumeric(A_ID) = 0 Then A_ID = 0
	A_ID = cCur(A_ID)
	R_ID = Left(Request.QueryString("RID"),14)
	If isNumeric(R_ID) = 0 Then R_ID = 0
	R_ID = cCur(R_ID)

	If A_NotReplay = 1 Then Exit Function

End Function

function GetUserName(UserID)

	Dim Rs,SQL
	SQL = sql_select("Select UserName from LeadBBS_User Where ID=" & UserID,1)
	Set Rs = LDExeCute(SQL,0)
	If Rs.Eof Then
		GetUserName = ""
	Else
		SQL = Rs(0)
		If isNull(SQL) Then SQL = 0
		GetUserName = SQL
	End If
	Rs.Close
	Set Rs = Nothing

End Function

Sub UpdateAnnounceApplicationInfo(AncID,IndexN,Value,tp,tid)

	Dim GetDataTop,AllTopNum,N,Str
	If tid = 0 Then
		Str = ""
	Else
		Str = tid
	End if
	AllTopNum = -1
	GetDataTop = Application(DEF_MasterCookies & "TopAnc" & Str)
	If isArray(GetDataTop) = False Then
		'If GetDataTop <> "yes" Then ReloadTopAnnounceInfo(tid)
		Exit Sub
	Else
		AllTopNum = Ubound(GetDataTop,2)
	End If

	For N = 0 to AllTopNum
		If cCur(AncID) = cCur(GetDataTop(0,N)) Then
			If tp = 1 Then
				GetDataTop(IndexN,N) = cCur(GetDataTop(IndexN,N)) + Value
			Else
				GetDataTop(IndexN,N) = Value
			End If
			Application.Lock
			Application(DEF_MasterCookies & "TopAnc" & Str) = GetDataTop
			Application.UnLock
			Exit Sub
		End If
	Next

End Sub

Function DisplayOfficerString(Officer)

	Dim Officer_Temp,Temp_N,dotFlag
	dotFlag = 0
	Officer_Temp = Split(Officer,",")
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

Function GetTopicInfo

	If A_ID = 0 Then Exit Function
	Dim Rs,SQL,Form_NeedValue
	A_ParentID = -1

	Dim ac,rd,Hits
	ac = Request("ac")
	rd = Left(Request("rd"),14)
	If isNumeric(rd) = 0 or inStr(rd,".") then rd = 0
	rd = cCur(rd)
	If rd = 0 or R_ID > 0 Then
		SQL = ""
	Else
		Select Case ac
			Case "pre": 
					select case DEF_UsedDataBase
						case 0,2:
							SQL = sql_select("Select ID,RootID,TopicType,NeedValue,ParentID,ChildNum,Title,hits,RootMaxID,RootMinID,NotReplay,RootIDBak,TitleStyle,VisitIP from LeadBBS_Announce where ParentID=0 and boardid=" & GBL_board_ID & " and RootID>" & rd & " order by RootID ASC",1)
						case Else
						SQL = sql_select("Select ID,RootID,TopicType,NeedValue,0,ChildNum,Title,hits,RootMaxID,RootMinID,NotReplay,ID,TitleStyle,VisitIP from LeadBBS_Topic where boardid=" & GBL_board_ID & " and RootID>" & rd & " order by RootID ASC",1)
					End select
			Case "nxt": 
					Rem For Access
					'SQL = "Select max(rootID) from LeadBBS_Announce where ParentID=0 and boardid=" & GBL_board_ID & " and RootID<" & rd
					'Set Rs = LDExeCute(SQL,0)
					'If Not Rs.Eof Then
					'	SQL = Rs(0)
					'	If isNull(SQL) Then SQL = 0
					'	SQL = cCur(SQL)
					'	If SQL >0 Then
							'SQL = sql_select("Select ID,RootID,TopicType,NeedValue,ParentID,ChildNum,Title,hits,RootMaxID,RootMinID,NotReplay,RootIDBak,TitleStyle,VisitIP from LeadBBS_Announce where ParentID=0 and boardid=" & GBL_board_ID & " and RootID=" & SQL,1)
					'	Else
					'		SQL = ""
					'	End If
					'Else
					'	SQL = ""
					'End If
					'Rs.Close
					'Set Rs = Nothing
					Rem SQL
					select case DEF_UsedDataBase
						case 0,2:
							SQL = sql_select("Select ID,RootID,TopicType,NeedValue,ParentID,ChildNum,Title,hits,RootMaxID,RootMinID,NotReplay,RootIDBak,TitleStyle,VisitIP from LeadBBS_Announce where ParentID=0 and boardid=" & GBL_board_ID & " and RootID<" & rd & " order by RootID DESC",1)
						case Else
							SQL = sql_select("Select ID,RootID,TopicType,NeedValue,0,ChildNum,Title,hits,RootMaxID,RootMinID,NotReplay,ID,TitleStyle,VisitIP from LeadBBS_Topic where boardid=" & GBL_board_ID & " and RootID<" & rd & " order by RootID DESC",1)
					End select
			Case Else: SQL = ""
		End Select
	End If

	If SQL <> "" Then
		Set Rs = LDExeCute(SQL,0)
		If Not Rs.Eof Then
			A_ID = cCur(Rs(0))
			A_RootID = cCur(Rs(1))
			Form_TopicType = Rs(2)
			Form_NeedValue = Rs(3)
			A_ParentID = cCur(Rs(4))
			A_ChildNum = cCur(Rs(5))
			A_Title = Rs(6)
			A_RootMaxID = cCur(Rs(8))
			A_RootMinID = cCur(Rs(9))
			A_NotReplay = Rs(10)
			A_RootIDBak = cCur(Rs(11))
			A_TitleStyle = Rs(12)
			ac = Trim(Rs(13))
			Hits = cCur(Rs(7)) + 1
			Rs.Close
			Set Rs = Nothing
			If ac <> GBL_IPAddress or LMT_RefreshEnable = 1 Then
				CALL LDExeCute("Update LeadBBS_Announce Set hits=" & Hits & ",VisitIP='" & GBL_IPAddress & "' where ID=" & A_ID,1)
				If DEF_UsedDataBase = 1 Then CALL LDExeCute("Update LeadBBS_Topic Set hits=" & Hits & ",VisitIP='" & GBL_IPAddress & "' where ID=" & A_ID,1)
				If inStr(application(DEF_MasterCookies & "TopAncList"),"," & A_ID & ",") Then
					UpdateAnnounceApplicationInfo A_ID,5,Hits,0,0
				Else
					If inStr(application(DEF_MasterCookies & "TopAncList" & GBL_Board_BoardAssort),"," & A_ID & ",") Then UpdateAnnounceApplicationInfo A_ID,5,Hits,0,GBL_Board_BoardAssort
				End If
			End If
		Else
			Rs.Close
			Set Rs = Nothing
		End If
	End If

	Dim BoardID
	If A_ParentID = -1 and R_ID = 0 Then
		BoardID = Request.QueryString("Aq")
		If ((BoardID = "" or BoardID = "1") and Request.QueryString("Ap") = "" and (Request.QueryString("ANum") = "" or Request.QueryString("AUpflag") = "")) and Request.QueryString("Re") <> "1" Then
			SQL = sql_select("Select T1.ID,T1.ParentID,T1.TopicSortID,T1.BoardID,T1.RootID,T1.ChildNum,T1.Layer,T1.Title,T1.Content,T1.FaceIcon,T1.ndatetime,T1.LastTime,T1.Hits,T1.Opinion,T1.UserName,T1.UserID,T1.HTMLFlag,T1.UnderWriteFlag,T1.NotReplay,T1.IPAddress,T2.Mail,T2.OICQ,T2.Userphoto,T2.UserLevel,T2.Homepage,T2.Underwrite,T2.Points,T2.Officer,T2.OnlineTime,T2.Birthday,T2.ApplyTime,T2.Sex,T2.LastDoingTime,T2.AnnounceNum2,T2.FaceUrl,T2.FaceWidth,T2.FaceHeight,T2.UserLimit,T1.TopicType,T1.NeedValue,T1.OtherInfo,T2.NongLiBirth,T2.ShowFlag,T2.NotSecret,T2.UserTitle,T1.TitleStyle,T1.PollNum,T2.CachetValue,T2.CharmPoint,T1.RootMaxID,T1.RootMinID,T1.RootIDBak,T1.VisitIP,T1.lastInfo from LeadBBS_Announce as T1 " & get_index("IX_LeadBBS_Announce_RootIDBak") & "left join LeadBBS_User as T2 on T2.Id=T1.Userid where T1.RootIDBak=" & A_ID & " order by T1.ID ASC",DEF_TopicContentMaxListNum)
			Set Rs = LDExeCute(SQL,0)
			If Not Rs.Eof Then
				A_Data = Rs.GetRows(-1)
				Rs.Close
				Set Rs = Nothing
				A_ParentID = cCur(A_Data(1,0))
			End If
			If A_ParentID = -1 or A_ParentID > 0 Then
				SQL = sql_select("Select T1.ID,T1.ParentID,T1.TopicSortID,T1.BoardID,T1.RootID,T1.ChildNum,T1.Layer,T1.Title,T1.Content,T1.FaceIcon,T1.ndatetime,T1.LastTime,T1.Hits,T1.Opinion,T1.UserName,T1.UserID,T1.HTMLFlag,T1.UnderWriteFlag,T1.NotReplay,T1.IPAddress,T2.Mail,T2.OICQ,T2.Userphoto,T2.UserLevel,T2.Homepage,T2.Underwrite,T2.Points,T2.Officer,T2.OnlineTime,T2.Birthday,T2.ApplyTime,T2.Sex,T2.LastDoingTime,T2.AnnounceNum2,T2.FaceUrl,T2.FaceWidth,T2.FaceHeight,T2.UserLimit,T1.TopicType,T1.NeedValue,T1.OtherInfo,T2.NongLiBirth,T2.ShowFlag,T2.NotSecret,T2.UserTitle,T1.TitleStyle,T1.PollNum,T2.CachetValue,T2.CharmPoint,T1.RootMaxID,T1.RootMinID,T1.RootIDBak,T1.VisitIP,T1.lastInfo from LeadBBS_Announce as T1 left join LeadBBS_User as T2 on T2.Id=T1.Userid where T1.ID=" & A_ID,1)
				Set Rs = LDExeCute(SQL,0)
				If Not Rs.Eof Then
					A_Data = Rs.GetRows(-1)
				Else
					A_ID = 0
				End If
				Rs.Close
				Set Rs = Nothing
			End If
			If isArray(A_Data) Then
				BoardID = cCur(A_Data(3,0))
				If  BoardID <> GBL_Board_ID Then
					GBL_Board_ID = BoardID
					Borad_GetBoardIDValue(GBL_Board_ID)
					CheckPass
					CheckisBoardMaster
				End If
				A_RootID = A_Data(4,0)
				Form_TopicType = A_Data(38,0)
				Form_NeedValue = A_Data(39,0)
				A_ParentID = cCur(A_Data(1,0))
				A_ChildNum = cCur(A_Data(5,0))
				A_RootMaxID = cCur(A_Data(49,0))
				A_RootMinID = cCur(A_Data(50,0))
				A_Title = A_Data(7,0)
				If A_ParentID > 0 and LCase(Left(A_Title,3)) = "re:" Then A_Title = Mid(A_Title,4)
				A_NotReplay = A_Data(18,0)
				A_RootIDBak = cCur(A_Data(51,0))
				A_TitleStyle = A_Data(45,0)
				ac = Trim(A_Data(52,0))
				Hits = cCur(A_Data(12,0)) + 1
				If ac <> GBL_IPAddress or LMT_RefreshEnable = 1 Then
					CALL LDExeCute("Update LeadBBS_Announce Set hits=" & Hits & ",VisitIP='" & GBL_IPAddress & "' where ID=" & A_ID,1)
					If DEF_UsedDataBase = 1 Then CALL LDExeCute("Update LeadBBS_Topic Set hits=" & Hits & ",VisitIP='" & GBL_IPAddress & "' where ID=" & A_ID,1)
					If A_ParentID = 0 and inStr(application(DEF_MasterCookies & "TopAncList"),"," & A_ID & ",") Then
						UpdateAnnounceApplicationInfo A_ID,5,Hits,0,0
					Else
						If A_ParentID = 0 and inStr(application(DEF_MasterCookies & "TopAncList" & GBL_Board_BoardAssort),"," & A_ID & ",") Then UpdateAnnounceApplicationInfo A_ID,5,Hits,0,GBL_Board_BoardAssort
					End If
				End If
			End If
		End If

		If A_ParentID = -1 Then
			SQL = sql_select("Select T1.ID,T1.ParentID,T1.TopicSortID,T1.BoardID,T1.RootID,T1.ChildNum,T1.Layer,T1.Title,T1.Content,T1.FaceIcon,T1.ndatetime,T1.LastTime,T1.Hits,T1.Opinion,T1.UserName,T1.UserID,T1.HTMLFlag,T1.UnderWriteFlag,T1.NotReplay,T1.IPAddress,T2.Mail,T2.OICQ,T2.Userphoto,T2.UserLevel,T2.Homepage,T2.Underwrite,T2.Points,T2.Officer,T2.OnlineTime,T2.Birthday,T2.ApplyTime,T2.Sex,T2.LastDoingTime,T2.AnnounceNum2,T2.FaceUrl,T2.FaceWidth,T2.FaceHeight,T2.UserLimit,T1.TopicType,T1.NeedValue,T1.OtherInfo,T2.NongLiBirth,T2.ShowFlag,T2.NotSecret,T2.UserTitle,T1.TitleStyle,T1.PollNum,T2.CachetValue,T2.CharmPoint,T1.RootMaxID,T1.RootMinID,T1.RootIDBak,T1.VisitIP,T1.lastInfo from LeadBBS_Announce as T1 left join LeadBBS_User as T2 on T2.Id=T1.Userid where T1.ID=" & A_ID,1)
			Set Rs = LDExeCute(SQL,0)
			If Not Rs.Eof Then
				A_Data = Rs.GetRows(-1)
			Else
				A_ID = 0
			End If
			Rs.Close
			Set Rs = Nothing
			If isArray(A_Data) Then
				BoardID = cCur(A_Data(3,0))
				If  BoardID <> GBL_Board_ID Then
					GBL_Board_ID = BoardID
					Borad_GetBoardIDValue(GBL_Board_ID)
					CheckPass
					CheckisBoardMaster
				End If
				A_RootID = A_Data(4,0)
				Form_TopicType = A_Data(38,0)
				Form_NeedValue = A_Data(39,0)
				A_ParentID = cCur(A_Data(1,0))
				A_ChildNum = cCur(A_Data(5,0))
				A_RootMaxID = cCur(A_Data(49,0))
				A_RootMinID = cCur(A_Data(50,0))
				A_Title = A_Data(7,0)
				If A_ParentID > 0 and LCase(Left(A_Title,3)) = "re:" Then A_Title = Mid(A_Title,4)
				A_NotReplay = A_Data(18,0)
				A_RootIDBak = cCur(A_Data(51,0))
				A_TitleStyle = A_Data(45,0)
				ac = Trim(A_Data(52,0))
				Hits = cCur(A_Data(12,0)) + 1
				If ac <> GBL_IPAddress or LMT_RefreshEnable = 1 Then
					CALL LDExeCute("Update LeadBBS_Announce Set hits=" & Hits & ",VisitIP='" & GBL_IPAddress & "' where ID=" & A_ID,1)
					If DEF_UsedDataBase = 1 Then CALL LDExeCute("Update LeadBBS_Topic Set hits=" & Hits & ",VisitIP='" & GBL_IPAddress & "' where ID=" & A_ID,1)
					If A_ParentID = 0 and inStr(application(DEF_MasterCookies & "TopAncList"),"," & A_ID & ",") Then
						UpdateAnnounceApplicationInfo A_ID,5,Hits,0,0
					Else
						If A_ParentID = 0 and inStr(application(DEF_MasterCookies & "TopAncList" & GBL_Board_BoardAssort),"," & A_ID & ",") Then UpdateAnnounceApplicationInfo A_ID,5,Hits,0,GBL_Board_BoardAssort
					End If
				End If
				If A_ParentID = 0 Then Set A_Data = Nothing
			End If
		End If
	End If

	If A_ParentID > 0 or R_ID > 0 Then
		'R_ID = cCur(A_ID)
		
		If R_ID > 0 Then
			SQL = sql_select("Select Title,RootMaxID,RootMinID,Hits,ChildNum,ID,TitleStyle,VisitIP,RootID,NotReplay,BoardID from LeadBBS_Announce where ID=" & A_ID & " and ParentID=0",1)
		Else
			If A_RootIDBak > 0 then
				SQL = sql_select("Select Title,RootMaxID,RootMinID,Hits,ChildNum,ID,TitleStyle,VisitIP,RootID,NotReplay,BoardID from LeadBBS_Announce where ID=" & A_RootIDBak,1)
			Else
				select case DEF_UsedDataBase
					case 0,2:
						SQL = sql_select("Select Title,RootMaxID,RootMinID,Hits,ChildNum,ID,TitleStyle,VisitIP,RootID,NotReplay,BoardID from LeadBBS_Announce where ParentID=0 and boardid=" & GBL_board_ID & " and RootID=" & A_RootID,1)
					case Else
						SQL = sql_select("Select Title,RootMaxID,RootMinID,Hits,ChildNum,ID,TitleStyle,VisitIP,RootID,NotReplay,BoardID from LeadBBS_Topic where boardid=" & GBL_board_ID & " and RootID=" & A_RootID,1)
				End select
			End If
		End If
		Set Rs = LDExeCute(SQL,0)
		If Rs.Eof Then
			Rs.Close
			Set Rs = Nothing
			GBL_CHK_TempStr = "指定的帖子不存在或已被删除。"
		Else
			A_Title = Rs(0)
			A_RootMaxID = cCur(Rs(1))
			A_RootMinID = cCur(Rs(2))
			A_ChildNum = cCur(Rs(4))
			A_RootIDBak = cCur(Rs(5))
			A_TitleStyle = Rs(6)
			ac = Trim(Rs(7))
			A_RootID = cCur(Rs(8))
			A_NotReplay = Rs(9)
			Hits = cCur(Rs(3)) + 1
			Rs.Close
			Set Rs = Nothing
		End If
	End If

	If Form_TopicType > 0 and A_NotReplay = 0 Then
		If GBL_CheckPassDoneFlag = 0 and GBL_CheckPassDoneFlag = 0 Then
			GBL_UserID = cCur(CheckPass)
			CheckisBoardMaster
			GBL_CHK_TempStr = ""
		End If
		Form_NeedValue = cCur(Form_NeedValue)
		Select case Form_TopicType
			Case 2: '仅本版版主
				If GBL_CHK_User <> "" Then
					If GBL_BoardMasterFlag <5 Then
						A_NotReplay = 1
					End If
				Else
					A_NotReplay = 1
				End If
			Case 4: '仅版主
				If GBL_BoardMasterFlag <4 Then
					A_NotReplay = 1
				End If
			Case 6:
				If GBL_CHK_User = "" or GetBinarybit(GBL_CHK_UserLimit,2) <> 1 Then
					A_NotReplay = 1
				End If
			Case 51:
				If GBL_CHK_Points < Form_NeedValue Then
					A_NotReplay = 1
				End If
			Case 53:
				If GBL_CHK_OnlineTime < Form_NeedValue*60 Then
					A_NotReplay = 1
				End If
		End Select
	End If

End Function



Sub B_DisplaySplitPageString(PageSplitString,css)
%>
	<div class="<%=css%> fire">
		<div class="a_post_image">
			<div class="layer_item">
				<a href="../a/a2.asp?b=<%=GBL_board_ID%>&amp;ID=<%=A_ID%>&amp;submitflag=first" class="b_repost_link"><img src="../images/blank.gif" class="b_repost" /></a>
				<div class="layer_iteminfo">
					<ul class="menu_list"><li><a href="../a/a2.asp?B=<%=GBL_board_ID%>">发表新主题</a></li>
					<li><a href="../a/a2.asp?B=<%=GBL_board_ID%>&amp;VoteFlag=yes">发起投票</a></li><%
					If A_NotReplay = 1 Then
					Else%>
					<li><a href="a2.asp?B=<%=GBL_board_ID%>&amp;ID=<%=A_ID%>&amp;submitflag=first">回复此主题</a></li><%
					End If%>
					</ul>
				</div>
			</div>
		</div>
		<%=PageSplitString%>
	</div>
<%
End Sub

Function DisplayTopic

	If Form_TopicType = 39 Then
		A_ChildNum = 0
		A_NotReplay = 1
	End If
	Dim Rs,SQL
	Dim ALL_FirstID,ALL_LastID
	ALL_FirstID = A_RootMaxID
	ALL_LastID = A_RootMinID

	Dim ALL_Count
	ALL_Count = A_ChildNum + 1

	Dim LMT_First,Temp1,Temp2
	Dim SQLEndString,Upflag,WhereFlag
	WhereFlag = 0

	Dim JumpOnly
	JumpOnly = 1
	LMT_First = Left(Request.QueryString("AFirst"),14)
	If isNumeric(LMT_First)=0 Then LMT_First=0
	LMT_First = cCur(LMT_First)
	If LMT_First > 0 Then JumpOnly = 0
	
	Dim LastNum,LastNumBak
	LastNum = 0
	
	LastNumBak = (ALL_Count mod DEF_TopicContentMaxListNum)
	If LastNumBak = 0 Then LastNumBak = DEF_TopicContentMaxListNum
	
	Upflag = Request.QueryString("AUpflag")
	If Upflag<>"1" and Upflag<>"0" Then Upflag="0"
	If Upflag = "1" Then
		LastNum = Request.QueryString("ANum")
		If LastNum <> "" Then
			LastNum = LastNumBak
		End If
		JumpOnly = 0
	End If

	Dim JMPRootID
	JMPRootID = Left(Request.QueryString("Ar"),14)
	If isNumeric(JMPRootID)=0 Then JMPRootID=0
	JMPRootID = Fix(cCur(JMPRootID))
	If JMPRootID>0 Then JumpOnly = 0

	Dim JMPage
	JMPage = Left(Request.QueryString("Aq"),14)
	If isNumeric(JMPage) = 0 Then JMPage = 0
	JMPage = Fix(cCur(JMPage))
	If JMPage > DEF_MaxJumpPageNum+1 Then JMPage = 0
	If JMPage = 0 Then JumpOnly = 0

	Dim MaxPage
	Page = Left(Request.QueryString("Ap"),14)
	If Page<>"" Then JumpOnly = 0
	If isNumeric(Page) = 0 or inStr(Page,".") > 0 Then Page = 0
	Page = cCur(Page)
	MaxPage = Fix(All_Count / DEF_TopicContentMaxListNum)
	If (All_Count mod DEF_TopicContentMaxListNum)<>0 Then MaxPage = MaxPage + 1
	
	Dim old_MaxPage,old_page
	old_MaxPage = MaxPage
	old_page = Page
	
	if JumpOnly = 0 Then MaxPage = MaxPage - 1
	If Page > MaxPage or LastNum > 0 Then
		Page = MaxPage
	End If
	
	
	If JMPage > Maxpage+1 or Maxpage < 0 Then JMPage = 0
	If Upflag="0" and JMPage+Page > MaxPage Then JMPage = 0
	If Upflag="1" and JMPage+Page < 0 Then JMPage = 0
	If JMPRootID > ALL_FirstID+1 and JumpOnly = 0 Then JMPage = 0
	If JMPRootID < ALL_LastID-1 and JumpOnly = 0 Then JMPage = 0

	If Upflag="1" Then
		Page = Page - JMPage
	Else
		Page = Page + JMPage
	End If

	If R_ID > 0 Then
		A_ID = A_RootIDBak
		Dim Index
		Set Rs = LDExeCute("Select Count(*) from LeadBBS_Announce where RootIDBak=" & A_ID & " and ID<" & R_ID,0)
		If Rs.Eof Then
			Index = 0
		Else
			Index = ccur(Rs(0))
		End If
		If isNull(Index) Then Index = 0
		Index = Index + 1
		JMPage = Fix(Index / DEF_TopicContentMaxListNum)
		If (Index mod DEF_TopicContentMaxListNum)<>0 Then JMPage = JMPage + 1
		Page = JMPage-1
		JMPRootID = A_ID-1
		LastNum = 0
		Upflag = "0"
		LMT_First = 0
		Set A_Data = Nothing
		A_ParentID = 0
	ElseIf Page = 0 or isArray(A_Data) Then '开启此项则当页数为0时即忽略一切信息的返回首页
		JMPage = 0
		JMPRootID = 0
		LastNum = 0
		Upflag = "0"
		LMT_First = 0
	End If

	Dim HaveIDFlag
	Dim LastID,FirstID

If isArray(A_Data) Then
	HaveIDFlag = 1
	GetData_2 = A_Data
	Temp2 = Ubound(GetData_2,2)
	FirstID_2 = cCur(GetData_2(0,0))
	LastID_2 = cCur(GetData_2(0,Temp2))
	If FirstID_2<LastID_2 Then
		SQL = FirstID_2
		FirstID_2 = LastID_2
		LastID_2 = SQL
	End If

	LastID = LastID_2
	FirstID = FirstID_2
Else
	If Temp1+1<DEF_TopicContentMaxListNum Then
		SQLEndString = " where T1.RootIDBak=" & A_RootIDBak
		WhereFlag = 1
		If Upflag="0" Then
			If JMPage > 0 Then
				If JMPRootID<>0 Then
					If WhereFlag = 1 Then
						SQLEndString = SQLEndString & " And T1.ID>" & JMPRootID
					Else
						SQLEndString = SQLEndString & " Where T1.ID>" & JMPRootID
						WhereFlag = 1
					End If
				End If
			Else
				If LMT_First<>0 Then
					If WhereFlag = 1 Then
						SQLEndString = SQLEndString & " And T1.ID>" & LMT_First
					Else
						SQLEndString = SQLEndString & " Where T1.ID>" & LMT_First
						WhereFlag = 1
					End If
				End If
			End If
		Else
			If JMPage > 0 Then
				If WhereFlag = 1 Then
					SQLEndString = SQLEndString & " And T1.ID<" & JMPRootID
				Else
					SQLEndString = SQLEndString & " Where T1.ID<" & JMPRootID
					WhereFlag = 1
				End If
			Else
				If LMT_First>=ALL_FirstID Then
					If LMT_First<>0 Then
						If WhereFlag = 1 Then
							SQLEndString = SQLEndString & " And T1.ID<" & LMT_First
						Else
							SQLEndString = SQLEndString & " Where T1.ID<" & LMT_First
							WhereFlag = 1
						End If
					End If
				Else
					If LMT_First<>0 Then
						If WhereFlag = 1 Then
							SQLEndString = SQLEndString & " And T1.ID<" & LMT_First
						Else
							SQLEndString = SQLEndString & " Where T1.ID<" & LMT_First
							WhereFlag = 1
						End If
					End If
				End If
			End If
		End If

		Dim NoPage
		NoPage = 0
		
		If Page < 0 or (Page > MaxPage and MaxPage>=(DEF_MaxJumpPageNum-1)) or (Page > (DEF_MaxJumpPageNum-1) and Page<(MaxPage-DEF_MaxJumpPageNum+1)) Then NoPage = 1
		If (LMT_First > 0 or LastNum>0 or NoPage = 1) and JMPage < 1 Then
			If Upflag="0" Then
				SQLEndString = SQLEndString & " order by T1.ID ASC"
			Else
				SQLEndString = SQLEndString & " order by T1.ID DESC"
			End If
			If LastNum > 0 Then
				SQL = LastNum
			Else
				SQL = DEF_TopicContentMaxListNum
			End If
		Else
			If JMPage > 0 Then
				If Upflag="0" Then
					SQLEndString = SQLEndString & " order by T1.ID ASC"
					Upflag="0"
					SQL = (JMPage-1) * DEF_TopicContentMaxListNum + DEF_TopicContentMaxListNum
				Else
					SQLEndString = SQLEndString & " order by T1.ID DESC"
					Upflag="1"
					SQL = (JMPage-1) * DEF_TopicContentMaxListNum + DEF_TopicContentMaxListNum
				End If
			Else
				If Page < DEF_MaxJumpPageNum Then
					SQLEndString = SQLEndString & " order by T1.ID ASC"
					Upflag="0"
					SQL = Page * DEF_TopicContentMaxListNum + DEF_TopicContentMaxListNum
				Else
					SQLEndString = SQLEndString & " order by T1.ID DESC"
					Upflag="1"
					SQL = (MaxPage-Page) * DEF_TopicContentMaxListNum + LastNumBak
				End If
			End If
		End If

		Dim FirstID_2,LastID_2
		Dim GetData_2

'新代码开始
If DEF_UsedDataBase = 0 or DEF_UsedDatabase = 2 and SQL>1000 Then
		Dim TmpSQL,moveNum
	select case DEF_UsedDataBase
	case 0:
		TmpSQL = sql_select("Select T1.ID from LeadBBS_Announce as T1 " & SQLEndString,sql)
		Set Rs = LDExeCute(TmpSQL,0)
		If (LMT_First = 0 and LastNum = 0 and Page >= 1 and NoPage = 0) or JMPage > 0 Then
			If Not Rs.Eof Then
				If JMPage > 0 Then
					If Upflag="0" Then
						Rs.Move (JMPage-1)* DEF_TopicContentMaxListNum
					Else
						If Page < MaxPage Then
							Rs.Move (JMPage-2) * DEF_TopicContentMaxListNum + DEF_TopicContentMaxListNum
						End If
					End If
				Else
					If Page < DEF_MaxJumpPageNum Then
						Rs.Move Page * DEF_TopicContentMaxListNum
					Else
						If Page < MaxPage Then
							Rs.Move (MaxPage-Page-1) * DEF_TopicContentMaxListNum + LastNumBak
						End If
					End If
				End If
			End If
		End If
	case 2:
		moveNum = 0
		If (LMT_First = 0 and LastNum = 0 and Page >= 1 and NoPage = 0) or JMPage > 0 Then
			If JMPage > 0 Then
				If Upflag="0" Then
					moveNum = (JMPage-1)* DEF_TopicContentMaxListNum
				Else
					If Page < MaxPage Then
						moveNum = (JMPage-2) * DEF_TopicContentMaxListNum + DEF_TopicContentMaxListNum
					End If
				End If
			Else
				If Page < DEF_MaxJumpPageNum Then
					moveNum = Page * DEF_TopicContentMaxListNum
				Else
					If Page < MaxPage Then
						moveNum = (MaxPage-Page-1) * DEF_TopicContentMaxListNum + LastNumBak
					End If
				End If
			End If
		End If
		TmpSQL = sql_select("Select T1.ID from LeadBBS_Announce as T1 " & SQLEndString,moveNum&"," & sql)
		Set Rs = LDExeCute(TmpSQL,0)
	end select
	Dim Cur_RootID
	If Not Rs.Eof Then
		Cur_RootID = Rs(0)
	Else
		Cur_RootID = 0
	End If
	Rs.Close
	Set Rs = Nothing
	Dim SQLEndString_J
	SQLEndString_J = SQLEndString
	If JumpOnly = 1 and inStr(SQLEndString_J," order by T1.ID ASC") Then
		SQLEndString_J = Replace(SQLEndString_J," order by T1.ID ASC"," and T1.ID>=" & Cur_RootID & " order by T1.ID ASC")
	Else
		SQLEndString_J = Replace(SQLEndString_J,">" & JMPRootID & " ",">=" & Cur_RootID & " ")
		SQLEndString_J = Replace(SQLEndString_J,"<" & JMPRootID & " ","<=" & Cur_RootID & " ")
		SQLEndString_J = Replace(SQLEndString_J,">" & LMT_First & " ",">=" & Cur_RootID & " ")
		SQLEndString_J = Replace(SQLEndString_J,"<" & LMT_First & " ","<=" & Cur_RootID & " ")
	End If
	TmpSQL = sql_select("Select T1.ID,T1.ParentID,T1.TopicSortID,T1.BoardID,T1.RootID,T1.ChildNum,T1.Layer,T1.Title,T1.Content,T1.FaceIcon,T1.ndatetime,T1.LastTime,T1.Hits,T1.Opinion,T1.UserName,T1.UserID,T1.HTMLFlag,T1.UnderWriteFlag,T1.NotReplay,T1.IPAddress,T2.Mail,T2.OICQ,T2.Userphoto,T2.UserLevel,T2.Homepage,T2.Underwrite,T2.Points,T2.Officer,T2.OnlineTime,T2.Birthday,T2.ApplyTime,T2.Sex,T2.LastDoingTime,T2.AnnounceNum2,T2.FaceUrl,T2.FaceWidth,T2.FaceHeight,T2.UserLimit,T1.TopicType,T1.NeedValue,T1.OtherInfo,T2.NongLiBirth,T2.ShowFlag,T2.NotSecret,T2.UserTitle,T1.TitleStyle,T1.PollNum,T2.CachetValue,T2.CharmPoint from LeadBBS_Announce as T1 " & get_index("IX_LeadBBS_Announce_RootIDBak") & "left join LeadBBS_User as T2 on T2.Id=T1.Userid " & SQLEndString_J,DEF_MaxListNum)
	Set Rs = LDExeCute(TmpSQL,0)
Else
'新代码结束

		MoveNum = 0
			If (LMT_First = 0 and LastNum = 0 and Page >= 1 and NoPage = 0) or JMPage > 0 Then
				If JMPage > 0 Then
					If Upflag="0" Then
						MoveNum = (JMPage-1)* DEF_TopicContentMaxListNum
					Else
						If Page < MaxPage Then
							MoveNum = (JMPage-2) * DEF_TopicContentMaxListNum + DEF_TopicContentMaxListNum
						End If
					End If
				Else
					If Page < DEF_MaxJumpPageNum Then
						MoveNum = Page * DEF_TopicContentMaxListNum
					Else
						If Page < MaxPage Then
							MoveNum = (MaxPage-Page-1) * DEF_TopicContentMaxListNum + LastNumBak
						End If
					End If
				End If
			end if
		select case def_useddatabase
			case 0,1:
				SQL = sql_select("Select T1.ID,T1.ParentID,T1.TopicSortID,T1.BoardID,T1.RootID,T1.ChildNum,T1.Layer,T1.Title,T1.Content,T1.FaceIcon,T1.ndatetime,T1.LastTime,T1.Hits,T1.Opinion,T1.UserName,T1.UserID,T1.HTMLFlag,T1.UnderWriteFlag,T1.NotReplay,T1.IPAddress,T2.Mail,T2.OICQ,T2.Userphoto,T2.UserLevel,T2.Homepage,T2.Underwrite,T2.Points,T2.Officer,T2.OnlineTime,T2.Birthday,T2.ApplyTime,T2.Sex,T2.LastDoingTime,T2.AnnounceNum2,T2.FaceUrl,T2.FaceWidth,T2.FaceHeight,T2.UserLimit,T1.TopicType,T1.NeedValue,T1.OtherInfo,T2.NongLiBirth,T2.ShowFlag,T2.NotSecret,T2.UserTitle,T1.TitleStyle,T1.PollNum,T2.CachetValue,T2.CharmPoint from LeadBBS_Announce as T1 left join LeadBBS_User as T2 on T2.Id=T1.Userid " & SQLEndString,sql)
			case 2:
				SQL = sql_select("Select T1.ID,T1.ParentID,T1.TopicSortID,T1.BoardID,T1.RootID,T1.ChildNum,T1.Layer,T1.Title,T1.Content,T1.FaceIcon,T1.ndatetime,T1.LastTime,T1.Hits,T1.Opinion,T1.UserName,T1.UserID,T1.HTMLFlag,T1.UnderWriteFlag,T1.NotReplay,T1.IPAddress,T2.Mail,T2.OICQ,T2.Userphoto,T2.UserLevel,T2.Homepage,T2.Underwrite,T2.Points,T2.Officer,T2.OnlineTime,T2.Birthday,T2.ApplyTime,T2.Sex,T2.LastDoingTime,T2.AnnounceNum2,T2.FaceUrl,T2.FaceWidth,T2.FaceHeight,T2.UserLimit,T1.TopicType,T1.NeedValue,T1.OtherInfo,T2.NongLiBirth,T2.ShowFlag,T2.NotSecret,T2.UserTitle,T1.TitleStyle,T1.PollNum,T2.CachetValue,T2.CharmPoint from LeadBBS_Announce as T1 left join LeadBBS_User as T2 on T2.Id=T1.Userid " & SQLEndString,MoveNum & "," & sql)
		end select
		Set Rs = LDExeCute(SQL,0)
		If (LMT_First = 0 and LastNum = 0 and Page >= 1 and NoPage = 0) or JMPage > 0 Then
			If Not Rs.Eof Then
				If DEF_UsedDatabase <> 2 then
					Rs.Move MoveNum
				End If
			End If
		End If
'新代码开始
End If
'新代码结束
		If Not Rs.Eof Then
			HaveIDFlag = 1
			GetData_2 = Rs.GetRows(DEF_TopicContentMaxListNum)
		Else
			HaveIDFlag = 0
		End If
		Rs.Close
		Set Rs = Nothing

		If HaveIDFlag = 1 Then
			Temp2 = Ubound(GetData_2,2)
			FirstID_2 = cCur(GetData_2(0,0))
			LastID_2 = cCur(GetData_2(0,Temp2))
	
			If FirstID_2<LastID_2 Then
				SQL = FirstID_2
				FirstID_2 = LastID_2
				LastID_2 = SQL
			End If
	
			LastID = LastID_2
			FirstID = FirstID_2
		Else
			Temp2 = 0
		End If
	Else
		HaveIDFlag = 0
	End If
End If

	Dim For1,For2,StepValue,DotFlag
	DotFlag = 0
	Dim RewriteFileName
	Dim ReWriteFlag
	If LMT_EnableRewrite = 0 or JumpOnly = 0 or Request.QueryString("E") <> "" Then
		RewriteFileName = "a.asp"
		ReWriteFlag = 0
	Else
		RewriteFileName = "topic"
		ReWriteFlag = 1
	End If
	
	If A_ParentID = 0 Then
		If JumpOnly = 1 Then
			page = page -1
		end if
		If ReWriteFlag = 0 Then
			SQL = "?B=" & GBL_board_ID & "&amp;ID=" & A_ID
		Else
			SQL = "-" & GBL_board_ID & "-" & A_ID
		End If
		Dim PageSplitString
		PageSplitString = "<div class=""j_page"">"
		If LastID <= All_LastID Then
		Else
			If ReWriteFlag = 0 Then
				PageSplitString = PageSplitString & "<a href=""" & RewriteFileName & "" & SQL & A_BoardUrl & """>1</a>"
			Else
				PageSplitString = PageSplitString & "<a href=""" & RewriteFileName & "" & SQL & "-1" & A_BoardUrl & ".html"">1</a>"
			End If
			
			if page <> 1 Then
				If ReWriteFlag = 0 Then
					PageSplitString = PageSplitString & "<a href=""" & RewriteFileName & "" & SQL & "&amp;AFirst=" & LastID & "&amp;AUpflag=1&amp;Ap=" & Page-1 & A_BoardUrl & """>上页"
				Else
					PageSplitString = PageSplitString & "<a href=""" & RewriteFileName & "" & SQL & "-" & Page & A_BoardUrl & ".html"">上页"
				End If
				If (Page - DEF_DisplayJumpPageNum) > 0 Then PageSplitString = PageSplitString & "…"
				PageSplitString = PageSplitString & "</a>"
			End If
		End If
	
		Dim DN,N
		DN = DEF_DisplayJumpPageNum
	
		If MaxPage > 0 Then
			If JumpOnly = 1 Then
				MaxPage = MaxPage - 1
			end if
			
			For1 = Page - DN
			For2 = Page + DN
			If For1 < 0 Then
				For1 = 0
			'ElseIf For1 > 0 Then
			'	PageSplitString = PageSplitString & "…"
			End If
			If For2 >= MaxPage Then For2 = MaxPage
			For N = For1 to For2
				If N = Page Then
					PageSplitString = PageSplitString & "<b>" & N + 1 & "</b>"
				Else
					If N <> MaxPage and N <> 0 Then
						If (N-Page) > 0 Then
							If ReWriteFlag = 0 Then
								PageSplitString = PageSplitString & "<a href=""" & RewriteFileName & "" & SQL & "&amp;Ar=" & FirstID & "&amp;AUpflag=0&amp;Ap=" & page & "&amp;Aq=" & N-Page & A_BoardUrl & """>" & N + 1 & "</a>"
							Else
								PageSplitString = PageSplitString & "<a href=""" & RewriteFileName & "" & SQL & "-" & N + 1 & A_BoardUrl & ".html"">" & N + 1 & "</a>"
							End If
						Else
							If ReWriteFlag = 0 Then
								PageSplitString = PageSplitString & "<a href=""" & RewriteFileName & "" & SQL & "&amp;Ar=" & LastID & "&amp;AUpflag=1&amp;Ap=" & page & "&amp;Aq=" & Page-N & A_BoardUrl & """>" & N + 1 & "</a>"
							Else
								PageSplitString = PageSplitString & "<a href=""" & RewriteFileName & "" & SQL & "-" & N + 1 & A_BoardUrl & ".html"">" & N + 1 & "</a>"
							End If
						End If
					End If
				End If
				DotFlag = 2
			Next
			If For2 < MaxPage Then
				'PageSplitString = PageSplitString & "…"
				DotFlag = 1
			End If
		Else
			PageSplitString = PageSplitString & "<b>1</b>"
		End If
		If FirstID >= All_FirstID Then
			'PageSplitString = PageSplitString & "下页"
			'PageSplitString = PageSplitString & "尾页"
		Else
			If page <> MaxPage-1 Then
				If ReWriteFlag = 0 Then
					PageSplitString = PageSplitString & "<a href=""" & RewriteFileName & "" & SQL & "&amp;AFirst=" & FirstID & "&amp;Ap=" & Page+1 & A_BoardUrl & """>"
				Else
					PageSplitString = PageSplitString & "<a href=""" & RewriteFileName & "" & SQL & "-" & Page+1+1 & A_BoardUrl & ".html"">"
				End If
				If (Page + DN) < MaxPage Then PageSplitString = PageSplitString & "…"
				PageSplitString = PageSplitString & "下页</a>"
			End If
			If ReWriteFlag = 0 Then
				PageSplitString = PageSplitString & "<a href=""" & RewriteFileName & "" & SQL & "&amp;AUpflag=1&amp;ANum=1&amp;Ap=" & MaxPage & A_BoardUrl & """>" & MaxPage + 1 & "</a>"
			Else
				PageSplitString = PageSplitString & "<a href=""" & RewriteFileName & "" & SQL & "-" & MaxPage+1 & A_BoardUrl & ".html"">" & MaxPage + 1 & "</a>"
			End If
		End If
		Rs = Temp2+Temp1
	
		If HaveIDFlag = 1 Then Rs = Rs+1
		If ReWriteFlag = 0 Then
			sql = RewriteFileName & SQL & "&amp;Ar=" & ALL_LastID-1 & "&amp;Ap=-1&AUpflag=0&amp;Aq='+(parseInt(this.value))"
		Else
			sql = RewriteFileName & SQL & "-'+(parseInt(this.value))+'" & A_BoardUrl & ".html'"
		End If
		'PageSplitString = PageSplitString & " 此主题共有" & ALL_Count &"帖 此页" & Rs & "帖 每页" & DEF_TopicContentMaxListNum & "帖"
		If MaxPage > DEF_DisplayJumpPageNum*2 Then PageSplitString = PageSplitString & "<input type=""text"" title=""输入页数,按Enter键跳转。"" size=""2"" onkeydown=""javascript:if(event.keyCode==13){location='" & sql & ";return false;}"">"
		PageSplitString = PageSplitString & "</div>"
	
	End If
	For1 = 0
	For2 = 0
	If (Temp2+Temp1 < 3 or A_ParentID > 0) and GBL_ShowBottomSure = 0 then GBL_SiteBottomString = ""
	If HaveIDFlag = 1 Then
		CALL B_DisplaySplitPageString(PageSplitString,"b_box_none")
		Global_TableHead
		If GBL_BoardMasterFlag >= 5 Then%>
<script src="<%=DEF_BBS_HomeUrl%>inc/js/p_list.js?ver=<%=DEF_Jer%>" type="text/javascript"></script>
		<%End If%>
<script type="text/javascript">
function a_command(cstr,obj,action)
{
	layer_view(cstr,obj,'','','anc_delbody','Processor.asp','',1,'AjaxFlag=1&action=' + action,1);return(false);
}
function a_msg(obj,action)
{
	layer_view('',obj,'','','anc_msgbody','Processor.asp','',1,'AjaxFlag=1&action=' + action);return false;
}
	<%If GBL_BoardMasterFlag >= 5 Then%>
function delbody_view(obj)
{
	layer_create("anc_msgbody");
	$id('anc_msgbody').innerHTML="<div class=ajaxbox><a href=\"javascript:;\" onclick=\"a_command('删除帖子',$id('" + obj.id + "'),'Del&b=<%=GBL_Board_ID%>&ID='+p_getselected());\">批量删除，已选择 <b id=layer_selectnum>" + p_getnum() + "</b> 条记录</a><br><input class=\"fmchkbox\" type=\"checkbox\" name=\"selmsg\" id=\"selmsg\" value=\"1\" onclick=\"achoose();\" />选择全部</div>";
	layer_view('',obj,'','','anc_msgbody','','',0,'',0,0);
}
	<%End If%>
</script>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" class="tablebox anc_headtable">
			<tr class="tbhead">
			<td>
				<div class="value" style="float:left">
				<span class=layerico><a href="javascript:void(0)" onclick="copyClipboard('Text',location.href,'帖子地址已复制到您的剪帖板!','<%=DEF_BBS_HomeUrl%>',this);">
				复制本帖地址</a></span>
				
				</div>
			<%If GBL_CHK_User <> "" and GBL_BoardMasterFlag >= 5 Then%>
				<div class="value" style="float:right">
				<div class="layer_item">
				<span><div class="layer_item_title">主题管理</div></span>
				<div class="layer_iteminfo" id="menu_anc_manage" onclick="this.style.display='none';">
				<ul class="menu_list">
					<li><a href="#Processor.asp?action=Top&b=<%=GBL_board_ID%>&ID=<%=A_RootIDBak%>" onclick="return(a_command('提升帖子',this,'Top&b=<%=GBL_board_ID%>&ID=<%=A_RootIDBak%>'));"><img src="../images/<%=GBL_DefineImage%>ti.GIF" alt="提取主题到最首位置" class="absmiddle" /> 提升主题</a></li>
					<li><a href="Processor.asp?action=Repair&b=<%=GBL_board_ID%>&ID=<%=A_RootIDBak%>" onclick="return(a_command('修复帖子/归入专题',this,'Repair&b=<%=GBL_board_ID%>&ID=<%=A_RootIDBak%>'));"><img src="../images/<%=GBL_DefineImage%>repair.GIF" alt="修复此主题帖子或归入专题" class="absmiddle" /> 修复/归入专题</a></li>
					
				<%If GetBinarybit(GBL_CHK_UserLimit,9) = 0 Then%>
					<li><a href="Processor.asp?action=Move&b=<%=GBL_board_ID%>&ID=<%=A_RootIDBak%>" onclick="return(a_command('转移帖子',this,'Move&b=<%=GBL_board_ID%>&ID=<%=A_RootIDBak%>'));"><img src="../images/<%=GBL_DefineImage%>move.GIF" alt="转移此主题到其它版面" class="absmiddle" /> 转移主题</a></li>
					
					<li><a href="Processor.asp?action=mirror&b=<%=GBL_board_ID%>&ID=<%=A_RootIDBak%>" onclick="return(a_command('镜像帖子',this,'mirror&b=<%=GBL_board_ID%>&ID=<%=A_RootIDBak%>'));"><img src="../images/<%=GBL_DefineImage%>move.GIF" alt="镜像此主题现状到其它版面" class="absmiddle" /> 镜像主题</a></li>
					
				<%End If
				
				If GetBinarybit(GBL_Board_BoardLimit,5) = 0 and GetBinarybit(GBL_CHK_UserLimit,5) = 0 Then
					Response.Write "<li><a href=""Processor.asp?action="
					If GBL_Board_ID <> 444 and DEF_EnableDelAnnounce = 0 Then
						SQL = "Move&b=" & GBL_Board_ID & "&ID=" & A_RootIDBak & "&BoardID2=444"
					Else
						SQL = "Del&b=" & GBL_Board_ID & "&ID=" & A_RootIDBak
					End If
					%>" onclick="return(a_command('删除帖子',this,'<%=SQL%>'));"><img src="../images/<%=GBL_DefineImage%>Del.GIF" alt="删除此帖子" class="absmiddle" /> 删除主题</a></li>
					<%
				End If%>
				<li><a href="Processor.asp?action=TopAnc&b=<%=GBL_board_ID%>&ID=<%=A_RootIDBak%>" onclick="return(a_command('帖子：固顶/取消固顶',this,'TopAnc&b=<%=GBL_board_ID%>&ID=<%=A_RootIDBak%>'));"><img src="../images/<%=GBL_DefineImage%>maketop.GIF" alt="主题固顶或取消固顶" class="absmiddle" /> 版面固顶</a></li>
				<%
				If GBL_BoardMasterFlag >= 6 Then%>
				<li><a href="Processor.asp?action=AllTopAnc&amp;b=<%=GBL_board_ID%>&amp;ID=<%=A_RootIDBak%>&amp;part=1" onclick="return(a_command('帖子：区固顶/取消区固顶',this,'AllTopAnc&b=<%=GBL_board_ID%>&ID=<%=A_RootIDBak%>&part=1'));"><img src="../images/<%=GBL_DefineImage%>makeparttop.GIF" alt="主题区固顶或取消区固顶" class="absmiddle" /> 区固顶</a></li><%
				End If
				If GBL_BoardMasterFlag >= 7 Then%>
				<li><a href="Processor.asp?action=AllTopAnc&amp;b=<%=GBL_board_ID%>&amp;ID=<%=A_RootIDBak%>" onclick="return(a_command('帖子：总固顶/取消总固顶',this,'AllTopAnc&b=<%=GBL_board_ID%>&ID=<%=A_RootIDBak%>'));"><img src="../images/<%=GBL_DefineImage%>makealltop.GIF" alt="主题总固顶或取消总固顶" align="middle" /> 总固顶</a></li><%
				End If
				%>
				</ul>
				</div>
				</div>
				</div>
	<%End If%>
			<div class="value" style="float:right">
			<span class="layerico">
			<%If GBL_CHK_User <> "" Then
					%><a href="Processor.asp?action=Collect&amp;b=<%=GBL_board_ID%>&amp;ID=<%=A_RootIDBak%>" onclick="return(a_msg(this,'Collect&SureFlag=1&b=<%=GBL_Board_ID%>&amp;ID=<%=A_RootIDBak%>'));">加入收藏</a>
				<%End If%>
			</span>
			</div>
		</td></tr></table>
	<%
	Else
		If GBL_ShowBottomSure = 0 Then GBL_SiteBottomString = ""
	End If
	
	If HaveIDFlag = 1 Then
		If Upflag="0" Then
			For1 = 0
			For2 = Temp2
			StepValue = 1
		Else
			For1 = Temp2
			For2 = 0
			StepValue = -1
		End If
		
		If GetData_2(38,For1) > 0 and cCur(GetData_2(1,For1)) = 0 and GetData_2(38,For1) <> 80 Then
			DisplayDataPage For1,For1,StepValue,GetData_2
		Else
			If cCur(GetData_2(1,For1)) = 0 and (GetData_2(16,For1) = 0 or GetData_2(16,For1) = 2) Then GetData_2(8,For1) = PrintTrueText(GetData_2(8,For1))
		End If
		DisplayJsAnnounce For1,For2,StepValue,GetData_2
		Global_TableBottom
		CALL B_DisplaySplitPageString(PageSplitString,"b_box_none2")
	Else
		Response.Write "<div class=""alertbox redfont"">指定的帖子不存在或已被删除。</div>"
		'Response.Redirect DEF_BBS_HomeUrl & "b/b.asp?b=" & GBL_Board_ID
	End If

End Function

Function DisplayDataPage(For1,For2,StepValue,GetData)

	Dim N,Flag,i,Rs
	Flag = 0
	i = 0
	For N = For1 to For2 Step StepValue
		i = i + 1
		If isNull(GetData(23,n)) Then
			GetData(20,n) = ""
			GetData(21,n) = 0
			GetData(22,n) = 0
			GetData(23,n) = 0
			GetData(24,n) = ""
			GetData(25,n) = ""
			GetData(26,n) = 0
			GetData(27,n) = 0
			GetData(28,n) = 0
			GetData(29,n) = 0
			GetData(30,n) = 0
			GetData(31,n) = "密"
			GetData(32,n) = 0
			GetData(33,n) = 0
			GetData(34,n) = ""
			GetData(35,n) = 0
			GetData(36,n) = 0
			GetData(37,n) = 0
			GetData(40,n) = ""
			GetData(42,n) = 1
			GetData(44,n) = ""
			GetData(47,n) = 0
			GetData(48,n) = 0
		End If
		If GetData(16,n) = 0 or GetData(16,n) = 2 Then GetData(8,n) = PrintTrueText(GetData(8,n))

		If GetData(38,n) > 0 and cCur(GetData(1,n)) = 0 and GetData(38,n) <> 80 Then
			Dim OnlineFlag
			GetData(39,n) = cCur(GetData(39,n))
			If GBL_CheckPassDoneFlag = 0 and GBL_CheckPassDoneFlag = 0 Then
				GBL_UserID = cCur(CheckPass)
				CheckisBoardMaster
				GBL_CHK_TempStr = ""
			End If
			GetData(15,n) = cCur(GetData(15,n))
			OnlineFlag = 0
			GBL_CHK_TempStr = ""
			Select case GetData(38,n)
				Case 1,2,61,62:
					If GetData(38,n) = 1 or GetData(38,n) = 61 Then
						Rs = "查看"
					Else
						Rs = "回复"
					End If
					If GBL_CHK_User <> "" Then
						If GBL_BoardMasterFlag < 5 and GetData(15,n) <> GBL_UserID and (GetData(38,n) = 1 or GetData(38,n) = 61) Then
							GBL_CHK_TempStr = "此帖只有本版" & DEF_PointsName(8) & "才能" & Rs
						Else
							GetData(40,N) = GetData(40,N) & " 此帖只有本版" & DEF_PointsName(8) & "才能" & Rs
							OnlineFlag = 1
						End If
					Else
						GBL_CHK_TempStr = "此帖只有本版" & DEF_PointsName(8) & "才能" & Rs
					End If
				Case 3,4,63,64:
					If GetData(38,n) = 3 or GetData(38,n) = 63 Then
						Rs = "查看"
					Else
						Rs = "回复"
					End If
					If GBL_BoardMasterFlag < 4 and GetData(15,n) <> GBL_UserID and (GetData(38,n) = 3 or GetData(38,n) = 63) Then
						GBL_CHK_TempStr = "此帖只有" & DEF_PointsName(8) & "才能" & Rs
					Else
						GetData(40,N) = GetData(40,N) & " 此帖只有" & DEF_PointsName(8) & "才能" & Rs
						OnlineFlag = 1
					End If
				Case 7,67:
					If GBL_UserID > 0 Then
						If GBL_UserID <> GetData(15,n) Then
							Set Rs = LDExeCute(sql_select("Select ID from LeadBBS_Announce where UserID=" & GBL_UserID & " and ParentID=" & GetData(0,N),1),0)
							If Rs.Eof Then
								GBL_CHK_TempStr = "回复此帖才能查看内容"
							Else
								GetData(40,N) = GetData(40,N) & " 回复此帖才能查看内容"
								OnlineFlag = 1
							End If
							Rs.Close
							Set Rs = Nothing
						Else
							GetData(40,N) = GetData(40,N) & " 回复此帖才能查看内容"
							OnlineFlag = 1
						End If
					Else
						GBL_CHK_TempStr = "回复此帖才能查看内容"
					End If
				Case 54,114,49,109:
					Dim Temp
					If GetData(38,n) = 49 or GetData(38,n) = 109 Then
						Temp = DEF_PointsName(1)
					Else
						Temp = DEF_PointsName(0)
					End If
					If GBL_UserID > 0 Then
						If GBL_UserID <> GetData(15,n) and GetData(39,n) > 0 Then
							Set Rs = LDExeCute(sql_select("Select ID from LeadBBS_VoteUser where AnnounceID=" & GetData(0,N) & " and UserName='" & Replace(GBL_CHK_User,"'","''") & "'",1),0)
							If Rs.Eof Then
								If GBL_CHK_Points < GetData(39,n) Then
									GBL_CHK_TempStr = "购买此帖需要花费" & GetData(39,n) & "" & Temp & "，可惜你的" & Temp & "不够，不能购买"
								Else
									GBL_CHK_TempStr = "购买此帖需要花费" & GetData(39,n) & "" & Temp & "，<a href=""javascript:void(0);"" onclick=""if($id('BuyResult').style.display=='none'){$id('BuyResult').style.display='block';getAJAX('a.asp','ol=3&amp;B=" & GBL_board_ID & "&amp;AnnounceID=" & GetData(0,N) & "','BuyResult');}"">点击购买帖子</a>"
								End If
							Else
								OnlineFlag = 1
							End If
							Rs.Close
							Set Rs = Nothing
						Else
							OnlineFlag = 1
						End If
					Else
						GBL_CHK_TempStr = "游客无权浏览限制帖子"
					End If
				Case 5,6,65,66:
					If GetData(38,n) = 5 or GetData(38,n) = 65 Then
						Rs = "查看"
					Else
						Rs = "回复"
					End If
					If (GBL_CHK_User = "" or (GetBinarybit(GBL_CHK_UserLimit,2) <> 1 and GetData(15,n) <> GBL_UserID) or GBL_UserID = 0) and (GetData(38,n) = 5 or GetData(38,n) = 65) Then
						GBL_CHK_TempStr = "此帖只有" & DEF_PointsName(5) & "才能" & Rs
					Else
						GetData(40,N) = GetData(40,N) & " 此帖只有" & DEF_PointsName(5) & "才能" & Rs
						OnlineFlag = 1
					End If
				Case 50,51,110,111:
					If GetData(38,n) = 50 or GetData(38,n) = 110 Then
						Rs = "查看"
					Else
						Rs = "回复"
					End If
					If GBL_CHK_Points < GetData(39,n) and GetData(15,n) <> GBL_UserID and (GetData(38,n) = 50 or GetData(38,n) = 110) Then
						GBL_CHK_TempStr = "此帖需要" & DEF_PointsName(0) & "" & GetData(39,n) & "才能" & Rs
					Else
						GetData(40,N) = GetData(40,N) & " 此帖需要" & DEF_PointsName(0) & "" & GetData(39,n) & "才能" & Rs
						OnlineFlag = 1
					End If
				Case 52,53,112,113:
					If GetData(38,n) = 52 or GetData(38,n) = 112 Then
						Rs = "查看"
					Else
						Rs = "回复"
					End If
					If GBL_CHK_OnlineTime < GetData(39,n)*60 and GetData(15,n) <> GBL_UserID and (GetData(38,n) = 52 or GetData(38,n) = 112) Then
						GBL_CHK_TempStr = "此帖需要" & DEF_PointsName(4) & GetData(39,n) & "才能" & Rs
					Else
						GetData(40,N) = GetData(40,N) & " 此帖需要" & DEF_PointsName(4) & GetData(39,n) & "才能" & Rs
						OnlineFlag = 1
					End If
				Case 55,115:
					If GetData(39,n) > 0 Then
						GetData(39,n) = GetUserName(GetData(39,n))
						If GetData(39,n) <> "" Then
							If GetData(39,n) <> GBL_CHK_User and GBL_UserID <> GetData(15,n) Then
								GBL_CHK_TempStr = "此帖只限用户<span class=""greenfont"">" & GetData(39,n) & "</span>才能查看"
							Else
								GetData(40,N) = GetData(40,N) & " 此帖只限用户<span class=""greenfont"">" & GetData(39,n) & "</span>才能查看"
								OnlineFlag = 1
							End If
						End If
					End If
				Case 39:
					GetData(8,n) = "<b>此帖为镜像帖，<a href=""a.asp?b=" & htmlencode(GetData(16,n)) & "&id=" & GetData(39,n) & """>点此查看原始帖。。。</a></b>"
					GetData(16,n) = 3
			End Select
			If GetData(38,n) > 60 Then
				If OnlineFlag = 1 Then
					'GetData(8,n) = Replace(Replace(GetData(8,n),"[HIDDEN]" & VbCrLf,"<br /><span class=""grayfont"">……………………………………………………隐藏内容…</span><br />",1,1,0),"[/HIDDEN]","<br /><span class=""grayfont"">…………………………………………………………………</span><br />",1,1,0)
					GetData(8,n) = Replace(Replace(GetData(8,n),"[HIDDEN]","<br /><span class=""grayfont"">……………………………………………………隐藏内容…</span><br />",1,1,0),"[/HIDDEN]","<br /><span class=""grayfont"">…………………………………………………………………</span><br />",1,1,0)
				Else
					OnlineFlag = inStr(GetData(8,n),"[HIDDEN]")
					If OnlineFlag > 0 Then
						Rs = inStr(OnlineFlag,GetData(8,n),"[/HIDDEN]")
						If Rs > OnlineFlag + 8 Then
							GetData(8,n) = Left(GetData(8,n),OnlineFlag-1) & "<br />" & GetFobStr("隐藏内容，" & GBL_CHK_TempStr) & Mid(GetData(8,n),Rs + 9)
						End If
					End If
				End If
			ElseIf GetData(38,n) > 0 and GBL_CHK_TempStr <> "" Then
				If OnlineFlag = 0 Then GetData(8,n) = GetFobStr(GBL_CHK_TempStr)
			End If
			GBL_CHK_TempStr = ""
		End If
	Next

End Function

Function GetFobStr(Str)

	GetFobStr = "<div class=""content_hidden""><span class=""grayfont"">……………………………………………………隐藏内容…<br /></span>" & _
				"<span class=""bluefont"">" & Str & "</span>[<a href=""../User/help/help.asp#lmt"">说明</a>]<span class=""grayfont""><br />" & _
				"…………………………………………………………………<br /></span></div><span id=""BuyResult"" style=""display:none"">提交中...</span>"

End Function

Function PrintTrueText(tempString)

	If tempString<>"" Then
		PrintTrueText=Replace(Replace(Replace(Replace(Replace(Replace(htmlEncode(tempString),VbCrLf & " ","<br />" & "&nbsp;"),"[P] ","[P]&nbsp;"),VbCrLf,"<br />" & VbCrLf),"   "," &nbsp; "),"  "," &nbsp;"),chr(9)," &nbsp; &nbsp; &nbsp;")
		If Left(PrintTrueText,1) = chr(32) Then
			PrintTrueText = "&nbsp;" & Mid(PrintTrueText,2)
		End If
	Else
		PrintTrueText=""
	End If

End Function

Function DisplayJsAnnounce(For1,For2,StepValue,GetData)

Dim SupervisorFlag,Temp
SupervisorFlag = CheckSupervisorUserName

Temp = LCase(Request.ServerVariables("server_name"))
If inStr(Temp,".") <> inStrRev(Temp,".") Then Temp = Mid(Temp,inStr(Temp,".") + 1)
%>
<script src="inc/leadcode.js?ver=20080728.225" type="text/javascript"></script>
<script type="text/javascript">
var GBL_domain="|<%=DEF_SafeUrl%>|";
var DEF_DownKey="<%=UrlEncode(DEF_DownKey)%>";
HU="<%=DEF_BBS_HomeUrl%>";

function appandOpinion(obj,id,num)
{
	if($id('opinion' + id))
	{
		$id('opinion' + id).style.display=($id('opinion' + id).style.display=='none')?'block':'none';
		return;
	}
	var tmp = document.createElement("div");
	tmp.setAttribute('id','opinion' + id);
	getAJAX('a.asp','ol=5&B=<%=GBL_board_ID%>&ID=' + id + '&num=' + num,'opinion' + id);
	insertAfter(tmp,obj.parentNode.parentNode);
}
</script>
<%
Dim Flag,i,n,Index
Flag = 0
i = 0
Index = 0


Response.Write "<div class=""contentbox"">"

dim bbsObj,outstr
if LMTDEF_ConvetType = 1 then
	Set bbsObj = CreateObject("leadbbs.bbsCode")
End If
%>
<script>
function swap_ancinfo(j,f)
{	
	if(f==1)
	{
		j.parentNode.parentNode.className="anc_table_div_show";
	}
	else
	{
		j.parentNode.parentNode.className="anc_table_div";
	}
}
</script>
<%

For N = For1 to For2 Step StepValue

	If ccur(GetData(42,n)) = 1 Then
		GetData(42,n) = 1
	Else
		GetData(42,n) = 0
	End If

	If isNull(GetData(23,n)) Then
		GetData(21,n) = 0
		GetData(22,n) = 0
		GetData(23,n) = 0
		GetData(26,n) = 0
		GetData(27,n) = 0
		GetData(28,n) = 0
		GetData(29,n) = 0
		GetData(30,n) = 0
		GetData(32,n) = 0
		GetData(33,n) = 0
		GetData(35,n) = 0
		GetData(36,n) = 0
		GetData(37,n) = 0
		GetData(42,n) = 1
		GetData(47,n) = 0
		GetData(48,n) = 0
	End If
	If (GetData(16,n) = 0 or GetData(16,n) = 2) and cCur(GetData(1,n)) <> 0 Then GetData(8,n) = PrintTrueText(GetData(8,n))
	
	If DEF_UbbLinkNum > 0 Then
		Dim ii
		dim re
		set re = New RegExp
		re.Global = True
		re.IgnoreCase = True
		If GetData(16,n) = 2 Then
			GetData(8,n) = Replace(GetData(8,n),VbCrLf,chr(3))			
			re.Pattern="(\[\/URL\])"
			GetData(8,n)=re.Replace(GetData(8,n),"[/URL]" & VbCrLf)
			For ii = 0 to DEF_UbbLinkNum - 1
				re.Pattern="([^\]])(" & DEF_UbbLinkData(ii) & ")"
				GetData(8,n) = re.Replace(GetData(8,n),"$1[url=" & DEF_UbbLinkUrl(ii) & "][color=blue][b]$2[/b][/color][/url]")
			Next
			GetData(8,n) = Replace(GetData(8,n),VbCrLf,"")
			GetData(8,n) = Replace(GetData(8,n),chr(3),VbCrLf)
		Else
			For ii = 0 to DEF_UbbLinkNum - 1
				re.Pattern="([^\>])(" & DEF_UbbLinkData(ii) & ")"
				GetData(8,n) = re.Replace(GetData(8,n),"$1<a href=" & DEF_UbbLinkUrl(ii) & " target=_blank><span class=bluefont color=blue><b>$2</b></span></a>")
			Next
		End If
		Set re = Nothing
	End If
	'-------------HTMLAnnounce Start--------------%>
	<!-- #include file=../inc/Templet/HTML/Normal_2.asp -->
	<%
	'-------------HTMLAnnounce End--------------
	If GetData(45,n) = 30 Then
		GetData(8,n) = "<span class=""redfont""><b>该用户因此帖永远禁言，此操作不可重复！</b></span>[<a href=""javascript:void(0)"" onclick=""$id('SpeLimit" & GetData(0,n) & "').style.display='block';"">点击查看]</a>。<div id=""SpeLimit" & GetData(0,n) & """ style=""display:none"" class=""a_quote""><table width=""100%"" cellpadding=""0"" cellspacing=""0""><tr><td>" & GetData(8,n) & "</td></tr></table></div>"
	Else
		If GetData(45,n) >=60 Then GetData(8,n) = GetFobStr("此帖有待管理人员审核才能查看")
	End If
	'----广告代码开始------
	If DEF_AD_DataNum > 0 Then
	If page*DEF_TopicContentMaxListNum+i = 2 or cCur(GetData(1,n)) = 0 Then '仅在楼主跟一楼位置显示广告
		Response.Write "<div class=a_topicad>"
		Response.Write AD_GetAdString
		Response.Write "</div>"
	End If
	End If
	'----广告代码结束------
	If GetBinarybit(GetData(37,n),7) = 1 and GetData(45,n) <> 30 Then
		Response.Write "<div class=""a_content"">"
		Response.Write GetFobStr("该用户发言已经被屏蔽")
		Response.Write "</div>"
	Else
		If Lcase(Left(GetData(7,n),3)) <> "re:" or cCur(GetData(1,n)) = 0 Then Response.Write "<div class='a_anctitle word-break-all'><b>" & DisplayAnnounceTitle(GetData(7,n),GetData(45,n)) & "</b></div>" & VbCrLf
		'If GetData(38,n) = 80 Then DisplayVoteForm GetData(0,n),0
		Response.Write "<div class=""a_content"">"
		Response.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" style=""table-layout:fixed; word-break:break-all;word-wrap:break-word;"" class=""a_content_table""><tr><td class=""a_content_td"">"
		If DEF_AnnounceFontSize <> "0" then Response.Write "<div style=""font-size:" & DEF_AnnounceFontSize & """>"
		Response.Write "<div id=""Content" & GetData(0,n) & """ class=""word-break-all"">"
		If GetData(16,n) <> 2 Then
			Response.Write GetData(8,n)
		Else
			if LMTDEF_ConvetType = 1 then
				if inStr(lcase(Request.ServerVariables("HTTP_USER_AGENT")),"msie") then
					Response.Write bbsObj.convertcode(GetData(8,n),DEF_BBS_HomeUrl,DEF_DownKey,"|" & DEF_SafeUrl & "|",outstr,"msie")
				else
					Response.Write bbsObj.convertcode(GetData(8,n),DEF_BBS_HomeUrl,DEF_DownKey,"|" & DEF_SafeUrl & "|",outstr,"other")
				end if
				'Response.Write "<div class=value2><em> Converted by leadbbs.bbsCode(" & outstr & " ms)</em></div>"
				'GetData(40,N) = GetData(40,N) & " Converted by leadbbs.bbsCode(" & outstr & " ms)</em>"
			else
				Response.Write GetData(8,n)
			end if
		End If
		
		
		Response.Write "</div>"

		If DEF_AnnounceFontSize <> "0" then Response.Write "</div>"
		Response.Write "</td></tr></table></div>"
		If GetData(38,n) = 54 or GetData(38,n) = 114 or GetData(38,n) = 49 or GetData(38,n) = 109 Then
			If GetData(38,n) = 49 or GetData(38,n) = 109 Then
				Temp = DEF_PointsName(1)
			Else
				Temp = DEF_PointsName(0)
			End If
			If cCur(GetData(46,n)) = 0 Then
				Response.Write "<div class=""a_contentnote"">[ <em>此帖为" & Temp & "出售帖，目前还没有人购买。</em> ]</div>"
			Else
				Response.Write "<div class=""a_contentnote"">[ <em>此帖为" & Temp & "出售帖，已经有<a href=""#no"" onclick=""if($id('BuyUser').style.display=='none'){$id('BuyUser').style.display='block';getAJAX('a.asp','ol=4&amp;B=" & GBL_board_ID & "&amp;ID=" & GetData(0,N) & "','BuyUser');}""><b>" & GetData(46,n) & "</b></a>人购买了此帖</em> ]</div><div id=""BuyUser"" style=""display:none"">提交中...</div>"
			End If
		End If
		If GetData(40,N) <> "" Then Response.Write "<div class=""a_contentnote"">[ <em>" & GetData(40,N) & "</em> ]</div>"

		If GetData(13,n) <> "" Then
			Temp = Len(GetData(13,n)) - Len(Replace(GetData(13,n),"|",""))
			If Temp = 2 or Temp = 3 Then
				If Temp = 2 Then
					Temp = Split(GetData(13,n),"|")
					If isNumeric(Temp(1)) = 0 Then Temp(1) = 0
					Temp(1) = Fix(cCur(Temp(1)))
					%>
					<div class="a_opinion<%If Temp(1) < 0 Then%>_un<%End If%> fire"><div class="a_opinion2 fire">此帖受到评价<%
					Response.Write "：" & htmlencode(Temp(2)) & " "
					If Temp(1) > 0 Then
						Temp(1) = "<span class=""bluefont"">＋" & Temp(1) & "</span>"
						Response.Write DEF_PointsName(2) & "" & Temp(1) 
					ElseIf Temp(1) < 0 Then
						Temp(1) = "<span class=""redfont"">" & Temp(1) & "</span>"
						Response.Write DEF_PointsName(2) & "" & Temp(1) 
					End If
					If Temp(0) <> "[LeadBBS]" Then
					Response.Write " by <a href=""../User/LookUserInfo.asp?name=" & urlEncode(Temp(0)) & """>" & Temp(0) & "</a>"
					Else
						Response.Write "<span class=""grayfont""> （此评价由系统生成）</span>"
					End If
					%>
					</div></div>
					<%
				Else
					Temp = Split(GetData(13,n),"|")
					If isNumeric(Temp(0)) = 0 Then Temp(0) = 0
					Temp(0) = Fix(cCur(Temp(0)))

					If isNumeric(Temp(1)) = 0 Then Temp(1) = 0
					Temp(1) = Fix(cCur(Temp(1)))
					
					If isNumeric(Temp(2)) = 0 Then Temp(2) = 0
					Temp(2) = Fix(cCur(Temp(2)))
					%>
					<div class="a_opinion<%If Temp(0)+Temp(1)+Temp(2) < 0 Then%>_un<%End If%> fire"><div class="a_opinion2 fire"><a href="javascript:;" id="opinionlink<%=GetData(0,n)%>" onclick="appandOpinion(this,<%=GetData(0,n)%>,<%=Temp(3)%>);">此帖受到<b><%=Temp(3)%></b>次评价</a><%
					
					If Temp(0) <> 0 or Temp(1) <> 0 or Temp(2) <> 0 Then
						Response.Write ", 累计："
						If Temp(0) > 0 Then
							Response.Write DEF_PointsName(0) & "＋" & Temp(0)
						ElseIf Temp(0) < 0 Then
							Response.Write DEF_PointsName(0) & Temp(0)
						End If
						Response.Write " "
						If Temp(1) > 0 Then
							Response.Write DEF_PointsName(2) & "<span class=""bluefont"">＋" & Temp(1) & "</span>"
						ElseIf Temp(1) < 0 Then
							Response.Write DEF_PointsName(2) & "<span class=""redfont"">" & Temp(1) & "</span>"
						End If
						Response.Write " "
						If Temp(2) > 0 Then
							Response.Write DEF_PointsName(1) & "<span class=""greenfont"">＋" & Temp(2) & "</span>"
						ElseIf Temp(2) < 0 Then
							Response.Write DEF_PointsName(1) & "<span class=""redfont"">" & Temp(2) & "</span>"
						End If
					End If
					%>
					</div></div>
					<%If LMT_ViewTopicOpinion = 1 and cCur(GetData(1,n)) = 0 Then
						%><script type="text/javascript">appandOpinion($id('opinionlink<%=GetData(0,n)%>'),<%=GetData(0,n)%>,<%=Temp(3)%>)</script>
					<%
					End If
				End If
			End If
		End If

		If GetData(38,n) = 80 Then
			Response.Write "<br />"
			DisplayVoteForm GetData(0,n),0
			Response.Write "<br />"
		End If
		If DEF_EnableUnderWrite = 0 or trim(GetData(25,n)&" ")="" Then GetData(25,n) = DEF_SiteNameString & "感谢您的参与"
		
		if LMTDEF_ShareID <> "" and cCur(GetData(1,n)) = 0 then 'share code
			LMTDEF_ShareID_Exist = "yes"
		%>
		<div style="float:right;margin:20px 0px 20px 20px;" id="shareHTML"></div>
		<%
		end if
		Response.Write "</td></tr><tr><td valign=""bottom"" class=""tdbottom"">"
		Response.Write "<div class=""a_contentbottom"">"
		If GetData(17,n) = 1 And trim(GetData(25,n)&" ")<>"" Then
			Response.Write "<div class=""a_signature"""
			if LMTDEF_ConvetType = 1 and GetData(16,n) = 2 then
				Response.Write " title=""编码转换时间: " & outstr & " ms"">"
			Else
				Response.Write ">"
			End If
			If DEF_EnableUnderWrite = 0 Then
				Response.Write GetData(25,n)
			Else
				Response.Write "<span id=""UnderWrite" & GetData(0,n) & """ class=""word-break-all"">"
				Response.Write PrintTrueText(GetData(25,n))
				%></span>
				<script type="text/javascript">
				<!--
					leadcode_uw('UnderWrite<%=GetData(0,n)%>');
				-->
				</script>
				<%
			End If
			Response.Write "</div>"
		End If
		Response.Write "</div>"
	End If
		%>
		</td>
	</tr>
	</table>
	</div>
	<%If GetData(16,n) = 2 and LMTDEF_ConvetType <> 1 Then
		%>
	<script type="text/javascript">
	<!--
	leadcode('Content<%=GetData(0,n)%>');
	-->
	</script>
	<%
	End If
Next
if LMTDEF_ConvetType = 1 then
	Set bbsObj = Nothing
	%><script type="text/javascript">
	leadcodebycom();
	</script>
	<%
End If

%>
</div>

<%

End Function

Sub Main

	GBL_CHK_PWdFlag = 0
	If CheckSupervisorUserName = 1 Then
		GBL_CHK_PWdFlag = 1
	End If
	initDatabase
	CheckisBoardMaster
	If dontRequestFormFlag = "" Then
	Select Case Left(Request.Form("ol"),1)
	Case "1"
		GBL_CHK_PWdFlag = 1
		CheckPollTitleID
		CloseDataBase
		Exit Sub
	Case "2","4"
		GBL_CHK_PWdFlag = 1
		PollUserList
		CloseDataBase
		Exit Sub
	Case "3"
		GBL_CHK_PWdFlag = 1
		DisplayBuyAnnounce
		CloseDataBase
		Exit Sub
	Case "5"
		GBL_CHK_PWdFlag = 1
		OpinionUserList
		CloseDataBase
		Exit Sub
	End Select
	End If

	GBL_CHK_TempStr = ""
	GetRequestValue
	If GBL_CHK_TempStr = "" Then GetTopicInfo
	
	GetBoardUrlString

	Dim Temp
	If A_TitleStyle >= 60 Then
		A_TitleNoHTML = "帖子等待审核中..."
		A_Title = "<span class=""grayfont"">帖子等待审核中...</span>"
		A_NotReplay = 1
	Else
		If A_TitleStyle = 1 Then
			A_TitleNoHTML = KillHTMLLabel(A_Title)
		Else
			A_TitleNoHTML = A_Title
		End If
	End If
	Temp = htmlencode(A_TitleNoHTML)
	If strLength(Temp)>DEF_BBS_DisplayTopicLength-6 Then
		Temp = LeftTrue(Temp,DEF_BBS_DisplayTopicLength-9) & "..."
	Else
		Temp = Temp
	End if

	GetBoardNavigateString

	DEF_GBL_Description = A_TitleNoHTML & " " & DEF_SiteNameString
	If GBL_CHK_TempStr = "" Then
		BBS_SiteHead A_TitleNoHTML & " - " & DEF_SiteNameString,0,GetBoardNavigateString & "<span class=""navigate_string_step"">查看帖子</span>"
	Else
		BBS_SiteHead DEF_SiteNameString,0,GetBoardNavigateString & "<span class=""navigate_string_step"">查看帖子</span>"
	End If

	Boards_Body_Head("")
	CheckAccessLimit

	If GetBinarybit(DEF_Sideparameter,18) = 1 Then
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
	If GBL_CHK_TempStr = "" and A_RootIDBak > 0 Then
	%>
	
	<div class="b_box_none fire">
		<div class="a_headinfo">
		<ul><li>主题：<b><a href="<%
		If LMT_EnableRewrite = 0 Then%>a.asp?B=<%=GBL_board_ID%>&amp;ID=<%=A_RootIDBak%><%
		Else
		%>topic-<%=GBL_board_ID%>-<%=A_RootIDBak%>-1.html<%
		End If%>"><%
		Response.Write htmlEncode(LeftTrue(A_TitleNoHTML,35))
		If StrLength(A_TitleNoHTML) > 35 Then Response.Write "..."%></a></b></li><li><%
		If A_ChildNum > 0 Then
			Response.Write "回复：<b>" & A_ChildNum & "</b> 帖"
		Else
			Response.Write "暂无回复"
		End If%></li></ul></div>
		<div class="a_box_list">
			<ul><%
			If A_ParentID = 0 or R_ID > 0 Then%>
			<li><a href="a.asp?B=<%=GBL_board_ID%>&amp;ID=<%=A_ID%>&amp;ac=pre&amp;rd=<%=A_RootID & A_BoardUrl%>">上篇</a></li><li>
			<a href="<%=DEF_BBS_HomeUrl%>b/<%			
			If LMT_EnableRewrite = 0 or Request.QueryString("E") <> "" Then
				Response.Write "b.asp?B=" & GBL_Board_ID & A_BoardUrl
			Else
				Response.Write "forum-" & GBL_Board_ID & A_BoardUrl
				If A_BoardUrl = "" Then Response.Write "-1"
				Response.Write ".html"
			End If%>">返回版面</a></li><li>
			<a href="a.asp?B=<%=GBL_board_ID%>&amp;ID=<%=A_ID%>&amp;ac=nxt&amp;rd=<%=A_RootID & A_BoardUrl%>">下篇</a></li><%
			Else%><li>
			<a href="a.asp?B=<%=GBL_board_ID%>&amp;ID=<%=A_RootIDBak%>&RID=<%=A_ID%>#F<%=A_ID%>"><b>此帖属于回复内容，点击查看完整帖子</b></a></li><%
			End If%></ul>
		</div>
	</div><%
	End If
	If GBL_CHK_TempStr = "" Then
		UpdateOnlineUserAtInfo GBL_board_ID,GBL_Board_BoardName & "→" & Temp
		DisplayTopic
		GetRequestValue
		DisplayAnnounceForm
		GBL_CHK_TempStr = ""
	Else
		If GBL_ShowBottomSure = 0 Then GBL_SiteBottomString = ""
		Global_ErrMsg GBL_CHK_TempStr
	End If
	CloseDatabase
	If GetBinarybit(DEF_Sideparameter,18) = 1 Then
	%></div>
	<%End If
	Boards_Body_Bottom
	If LMTDEF_ShareID <> "" and LMTDEF_ShareID_Exist = "yes" Then
	'share code start
		'=LMTDEF_ShareID%>
		
	<script type="text/javascript">
	var sharehtml = " <%=replace(replace(replace(LMTDEF_ShareID,"\","\\"),"""","\"""),"script","s\x63ript")%>";
	$id('shareHTML').innerHTML = sharehtml;
	js_Reload($id('shareHTML'));
	
	</script>
	<%
	'share code end
	End If
	SiteBottom

End Sub

Sub GetBoardUrlString

	Dim Tmp,Tmp2,E
	E = Request.QueryString("E")
	
If LMT_EnableRewrite = 1 and E = "" Then
	Tmp = Left(Request.QueryString("q"),14)
	If isNumeric(Tmp) = 0 Then Tmp = 0
	Tmp = Fix(cCur(Tmp))
	If Tmp > 0 Then A_BoardUrl = A_BoardUrl & "-" & Tmp
Else	
	Tmp = Left(Request.QueryString("r"),14)
	If isNumeric(Tmp) = 0 Then Tmp = 0
	Tmp = Fix(cCur(Tmp))
	If Tmp > 0 Then A_BoardUrl = "&r=" & Tmp
	
	Tmp = Left(Request.QueryString("p"),14)
	If isNumeric(Tmp) = 0 Then Tmp = 0
	Tmp = Fix(cCur(Tmp))
	If Tmp > 0 Then A_BoardUrl = A_BoardUrl & "&p=" & Tmp
	
	Tmp = Left(Request.QueryString("Upflag"),14)
	If Tmp = "0" or Tmp = "1" Then A_BoardUrl = A_BoardUrl & "&Upflag=" & Tmp
	
	Tmp = Left(Request.QueryString("Num"),14)
	If Tmp <> "" Then Tmp = "1"
	If Tmp = "1" Then A_BoardUrl = A_BoardUrl & "&Num=1"
	
	Tmp = Left(Request.QueryString("q"),14)
	If isNumeric(Tmp) = 0 Then Tmp = 0
	Tmp = Fix(cCur(Tmp))
	If Tmp > 0 Then A_BoardUrl = A_BoardUrl & "&q=" & Tmp
	
	Tmp = Left(Request.QueryString("RootID"),14)
	If isNumeric(Tmp) = 0 Then Tmp = 0
	Tmp = Fix(cCur(Tmp))
	If Tmp > 0 Then A_BoardUrl = A_BoardUrl & "&RootID=" & Tmp
	
	Tmp = E
	If Tmp = "1" Then
		A_BoardStr = "<span class=""navigate_string_step""><a href=""../b/b.asp?B=" & GBL_Board_ID & "&amp;E=1"">专题区</a></span>"
		Tmp = Left(Request.QueryString("EID"),14)
		If isNumeric(Tmp) = 0 Then Tmp = 0
		Tmp = Fix(cCur(Tmp))
		If Tmp < 1 Then Tmp = 0
		Tmp2 = ""
		If Tmp > 0 Then Tmp2 = GetEName(Tmp)
		If Tmp2 = "" Then Tmp = 0
		A_BoardUrl = A_BoardUrl & "&E=1"
		If Tmp > 0 Then
			A_BoardStr = A_BoardStr & "<span class=""navigate_string_step""><a href=""../b/b.asp?B=" & GBL_Board_ID & "&amp;E=1&amp;EID=" & tmp & """>" & Tmp2 & "</a></span>"
			A_BoardUrl = A_BoardUrl & "&amp;EID=" & Tmp
		End If
	ElseIf Tmp <> "" Then
		A_BoardStr = "<span class=""navigate_string_step""><a href=""../b/b.asp?B=" & GBL_Board_ID & A_BoardUrl & """>精华区</a></span>"
		'A_BoardUrl = ""
		A_BoardUrl = A_BoardUrl & "&amp;E=0"
	End If
End If

End Sub

Function GetBoardNavigateString

	If GBL_board_ID = 0 Then Exit Function
	Dim Temp,TempStr,N,Rewrite
	Temp = cCur(Application(DEF_MasterCookies & "BoardInfo" & GBL_board_ID)(26,0))
	Do While Temp > 0
		If isArray(Application(DEF_MasterCookies & "BoardInfo" & Temp)) = False Then
			ReloadBoardInfo(Temp)
			If isArray(Application(DEF_MasterCookies & "BoardInfo" & Temp)) = False Then Exit Do
		End If
		If LMT_EnableRewrite = 0 or Request.QueryString("E") <> "" Then
			Rewrite = "b.asp?B=" & Temp & A_BoardUrl
		Else
			Rewrite = "forum-" & Temp & A_BoardUrl
			If A_BoardUrl = "" Then Rewrite = Rewrite & "-1"
			Rewrite = Rewrite & ".html"
		End If
		TempStr = "<span class=""navigate_string_step""><a href=""" & DEF_BBS_HomeUrl & "b/" & Rewrite & """>" & Application(DEF_MasterCookies & "BoardInfo" & Temp)(0,0) & "</a></span>" & TempStr
		Temp = cCur(Application(DEF_MasterCookies & "BoardInfo" & Temp)(26,0))
		N = N + 1
		If N > 10 Then Exit Do
	Loop
	If LMT_EnableRewrite = 0 or Request.QueryString("E") <> "" Then
		Rewrite = "b.asp?B=" & GBL_Board_ID & A_BoardUrl
	Else
		Rewrite = "forum-" & GBL_Board_ID & A_BoardUrl
		If A_BoardUrl = "" Then Rewrite = Rewrite & "-1"
		Rewrite = Rewrite & ".html"
	End If
	TempStr = TempStr & "<span class=""navigate_string_step""><a href=""" & DEF_BBS_HomeUrl & "b/" & Rewrite & """>" & GBL_Board_BoardName & "</a></span>"
	GetBoardNavigateString = TempStr & A_BoardStr

End Function

Function GetEName(ID)

	Dim TArray,N,Num
	TArray = Application(DEF_MasterCookies & "BoardInfo" & GBL_Board_ID & "_TI")
	If isArray(TArray) = False Then Exit Function
	Num = Ubound(TArray,2)
	For N = 0 To Num
		If ID = cCur(TArray(0,N)) Then
			GetEName = TArray(1,n)
			Exit Function
		End If
	Next
	GetEName = ""

End Function

Main
%>