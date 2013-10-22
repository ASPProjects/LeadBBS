<!-- #include file=../inc/BBSsetup.asp -->
<!-- #include file=../inc/MD5.asp -->
<%


Server.ScriptTimeOut = 999999
Response.Buffer = False
DEF_BBS_HomeUrl = "../"
Dim Con,GBL_CHK_TempStr

UpdateDatabase

Sub OpenDatabase

	on error resume next
	Dim DB
	DB = Request("db")
	Set con = Server.CreateObject("ADODB.Connection")
	select case DEF_UsedDataBase
		case 0,2:	
			If DB = "" Then db = DEF_AccessDatabase
			Con.ConnectionString = db
		case Else
			If DB = "" Then db = Server.MapPath(DEF_BBS_HomeUrl & DEF_AccessDatabase)
			Con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & db
	End select
	'Con.ConnectionString = "driver={Microsoft Access Driver (*.mdb)};dbq=" & db
	Con.Open
	Con.CommandTimeout = 3600
	If err Then
		%>
		数据库连接错误，请确定数据库连接串是否正确！<br><br><font color=red><%=err.description%></font>
		<br><br><a href=Update.asp><b>&lt;&lt;返回升级界面</b></a>
		<%Err.clear
		Response.End
	End If

End Sub

Sub CloseDatabase

	Con.Close
	Set Con = Nothing

End Sub


Function LDExeCute(sql,flag)

	on error resume next
	If flag = 0 or flag = 3 Then
		Set LDExeCute = Con.ExeCute(SQL)
	Else
		Con.ExeCute(SQL)
	End If
	
	If Err Then
		Response.Write "<p>以下SQL语句执行出错：</p><p><font color=gray>" & server.htmlencode(SQL) & "</font></P>"
		Response.Write "<p>错误描述: <font color=red>" & err.description & "</font></p>"
		Err.Clear
	End If

End Function

Function CheckSupervisorUserName

	If Session(DEF_MasterCookies & "Manager") = "manage" Then
		CheckSupervisorUserName = 1
	Else
		CheckSupervisorUserName = 0
	End If

End Function

Sub Closebbs

	Application.Lock
	application(DEF_MasterCookies & "SiteEnableFlagzoieiu") = 0
	application(DEF_MasterCookies & "SiteDisbleWhyszoieiu") = "论坛更新中，请稍候来访."
	Application.UnLock

End Sub

Sub restartbbs

	Application.Lock
	application(DEF_MasterCookies & "SiteEnableFlagzoieiu") = 1
	application(DEF_MasterCookies & "SiteDisbleWhyszoieiu") = ""
	Application.UnLock
	on error resume next
	Application.Contents.RemoveAll

End Sub

Sub UpdateDatabase

	If CheckSupervisorUserName = 0 Then
		Response.Write "Time out."
		Response.End
	End If
	%>
	<html>
	<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
	<title>LeadBBS 6.0/6.1 升级程序</title>
	<style>
	html {height:100%; } 
body{color:black;font: 12px "helvetica neue", "lucida grande", helvetica, arial, sans-serif; background-color:#ffffff;
height:100%}
body{padding: 0px;margin: 0px;}

p { margin: 0px 0px 5px 0px;line-height:1.5em;}
textarea{font-size:9pt;overflow-y:auto;width: 95%; word-break: break-all;word-wrap:break-word;}
input{font-size:9pt;}
input:focus, input:hover { background-color: #f1f1f1; }
select{font-size:9pt;height:20px;color:black;background-color:#f5fafe}
table{text-align: left;}
.fminpt{padding:0px 4px 0px 4px;border-right:#B8D5EA 1px solid;border-top:#B8D5EA 1px solid;font-size:9pt;border-left:#B8D5EA 1px solid;border-bottom:#B8D5EA 1px solid;height:17px;line-height:17px;vertical-align: middle;}
.input_1{width:40px;}
.input_2{width:100px;}
.input_3{width:150px;}
.input_4{width:300px;}
.fmchkbox{font-size:9pt;vertical-align: middle;border:0px;}
.word-break-all{word-break: break-all;word-wrap:break-word;}
.clicktext{cursor: pointer;color:#0055aa;}

.fmtxtra{padding:4px;border-right:#B8D5EA 1px solid;border-top:#B8D5EA 1px solid;font-size:9pt;border-left:#B8D5EA 1px solid;border-bottom:#B8D5EA 1px solid;}
.fmbtn{
   }
td{font-size:9pt;}
li{font-size:9pt;}
ul{font-size:9pt;}
a{color:black;text-decoration:none;}
a:hover{text-decoration:none;color:#0055aa;}
/*a.visit:visited {padding-right:12px; background: url(../../images/style/0/visited.gif) no-repeat 100% 50%;}*/
.unsel{outline: none;-moz-user-select: none;}

.grayfont{color:gray;}
.redfont{color:red;}

.frame_table{border-bottom:#b8d5ea 1px solid;border-top:#b8d5ea 1px solid;border-left:1px solid #b8d5ea;border-right:1px solid #b8d5ea;background-color:#ffffff;table-layout:fixed;}
.frame_table .tdbox{padding:6px;border-top:1px dotted #E7D1B0;}
.frame_tbhead td{background-color:#f5f5f5;padding-bottom:1px;padding-top:1px;}/*padding for Mozilla*/
.frame_tbhead .value{background-color:#f5f5f5;padding-left:6px;padding-right:3px;padding-top:6px;padding-bottom:5px;}
.alert{color:red;font-weight:bold;padding-bottom:12px;padding-top:12px;}
.alertdone{color:green;font-weight:bold;padding-bottom:12px;padding-top:12px;}
.frameline{line-height:26px;margin:2px 0px 2px 18px;padding:0px;}
.note {color:gray;font-size:8pt;}
.frame_body{margin: 15px;}
</style>
	</head>
	<body>
		<div class="frame_body">
		<br><br>
		<p><b style="font-size:14px;">　适用于LeadBBS 6.0/6.1及以上版本的升级(及配置扩展参数)工具</b></p>
		<div class="frameline">1. 导出扩展参数：若论坛的文件被手动替换，此功能将恢复论坛的原先所存储的配置。</div>
		<div class="frameline">2. 配置扩展参数：更多论坛的参数设置选项可以在此找到。</div>
		<div class="frameline">3. 检测是否有版本更新：将您的论坛与官方连接比较，检测是否有新的更新。</div>
		<div class="frameline">4. 立即更新补丁：将您的论坛与官方连接比较，并升级到最新版本。</div>
		<div style="width:90%;margin-bottom:216px;BORDER: #EEE0CB 5px solid; BACKGROUND: #F9F5F0; text-align:left;width:500px;padding:22px;line-height:2.0">

		
	<%
	If Request("sure") = "1" Then
		dim startflag:startflag=request("startflag")
		if startflag <> "1" then%>
		<p><form action=update.asp method=post>
			<br><b><span color=ff0000 class=redfont>注意：继续将进行以下操作<br></span></b>
			<br>
			<ul><li>将当前网站设置与数据库对比，若与数据库中的配置不符将替换当前配置
			<%If LCase(Request("checkversion")) = "updateversion" Then
			%><li><font color=blue><b>您选择了自动更新，若有更新将会强制替换相应的更新文件</b></font></li>
			<%end If%>
			<li>为保证更新安全，检测程序将强制暂停论坛运行，直到操作完成</li>
			</li>
			</ul>
			<input type=hidden name=sure value="<%=server.htmlencode(request("sure"))%>">
			<input type=hidden name=SubmitFlag value="<%=server.htmlencode(request("SubmitFlag"))%>">
			<input type=hidden name=startflag value="1">
			<input type=hidden name=checkversion value="<%=server.htmlencode(request("checkversion"))%>">
			
			<input type=submit value=确定继续 class=fmbtn>
			</form>
		<%
		else
			OpenDatabase
			Closebbs
			If LCase(Request("checkversion")) = "checkversion" Then
				Update62_initBBSdata
				Update_CheckVersion
			Else
				Update62_initBBSdata
				If LCase(Request("checkversion")) = "updateversion" Then
					Update62_CopyFile
				End If
			End If
			restartbbs
			CloseDatabase
		end if
	Else
		%>
		<form action=Update.asp method="post">
		<p>
		注意：此升级程序仅限LeadBBS 6.0/6.1及更高版本的升级更新。
		</p>
		<br />
		<font color=blue>如果您的版本更旧，请先至官方下载6.0版本更新。</font>
		<p>
		<!--
		<input class="fmchkbox" type="checkbox" name="leadbbs40" value="1" />我的数据库还是4.0版本<br />
		-->
		</p>
		<br /><font color=red>警告：升级程序将会强制替换本地文件，<b>务必先作好备份</b>！</font><br />
		<br />
		注意：此程序将完成数据库及本地文件的更新，不需要额外进行手动更新。
		<p>数据库链接串：<input name=db value="<%If DEF_UsedDataBase <> 1 Then
				Response.Write server.htmlencode(DEF_AccessDatabase)
			Else
				Response.Write server.htmlencode(Server.MapPath(DEF_BBS_HomeUrl & DEF_AccessDatabase))
			End If%>" size=40 class='fminpt input_3'>
		</p>
		<input name=sure value=1 type=hidden>
		<p><input type=submit value="开始更新" class="fmbtn btn_4"></p>
		</form>
		<%
	End If
	Update_PageBottom

End Sub

Sub Update_PageBottom

	%>
		</div>
		</div>
		<div style="padding:60px;"></div>
	</body>
	</html>
	<%
End Sub

Dim SubmitFlag
Dim FSFlag
Dim GBL_LeadBBS_Setup_Data '临时读取的SetupRID记录数据数组
Dim SetupRID_1050
'0 临时备份目录名称
'1 a2.asp/Const LMTDEF_MaxReAnnounce值 =
'
Dim GBL_Update_LineStr '获取的临时文件字符串行。
Dim GBL_UpdateVersion '内部版本号
GBL_UpdateVersion = 0
Dim Update_UpdateFileFlag
Update_UpdateFileFlag = 0
Dim GBL_ParaCount
GBL_ParaCount = 44

Sub Update62_UpdateDatabase

	If Update_CheckFields("saveData","LeadBBS_Setup") = False Then
		If DEF_UsedDataBase = 0 Then
			CALL LDExeCute("ALTER TABLE LeadBBS_Setup ADD saveData text NOT NULL CONSTRAINT DF_LeadBBS_Setup_saveData DEFAULT ''",1)
		Else
			CALL LDExeCute("ALTER TABLE LeadBBS_Setup ADD saveData memo default ''",1)
			CALL LDExeCute("ALTER TABLE LeadBBS_Setup ALTER COLUMN saveData memo Default ''",1)			
		End If

		If SubmitFlag = "" Then
		If DEF_UsedDataBase = 0 Then
			CALL LDExeCute("CREATE NONCLUSTERED INDEX IX_LeadBBS_Announce_TopicType ON LeadBBS_Announce(TopicType,NeedValue) ON [PRIMARY]",1)
			CALL LDExeCute("CREATE NONCLUSTERED INDEX IX_LeadBBS_Announce_lastTime ON LeadBBS_Announce(ParentID,BoardID,LastTime) ON [PRIMARY]",1)
			CALL LDExeCute("CREATE NONCLUSTERED INDEX IX_LeadBBS_User_LastDoingTime ON LeadBBS_User(LastDoingTime DESC,ID) ON [PRIMARY]",1)
		Else
			CALL LDExeCute("CREATE INDEX IX_LeadBBS_Announce_TopicType ON LeadBBS_Announce(TopicType,NeedValue)",1)
			CALL LDExeCute("CREATE INDEX IX_LeadBBS_Announce_lastTime ON LeadBBS_Announce(ParentID,BoardID,LastTime)",1)
			CALL LDExeCute("CREATE INDEX IX_LeadBBS_User_LastDoingTime ON LeadBBS_User(LastDoingTime DESC,ID)",1)
		End If
		End If
	End If

End Sub

Dim CurN

Sub Update62_initBBSdata

	SubmitFlag = Request("SubmitFlag")

	Dim RID,ValueStr,ClassNum,saveData
	ReDim SetupRID_1050(5,100)

	CALL Update_ECHO("<div class=alertdone>初始化检测。。。</div>",1)
	'检测FSO
	If Update_CheckFSO = 0 Then
		CALL Update_ECHO("权限不足：空间不支持FSO操作.",1)
		Exit Sub
	Else
		CALL Update_ECHO("权限检测：完成.",0)
	End If
	
	'检测数据库
	'Update62_UpdateDatabase

	'初始化备份文件目录位置
	If Update_CheckSetupRIDExist(1050," and ClassNum=0") = 0 Then
		RID = 1050
		Randomize
		ValueStr = Right(MD5(Rnd*10000000*hour(now)),8)
		SetupRID_1050(0,0) = ValueStr
		ClassNum = 0
		saveData = ""
		CALL Update_InsertSetupRID(RID,ValueStr,ClassNum,saveData," and ClassNum=0")
		CALL Update_ECHO("备份文件目录未建立,初始化完成,临时备份文件存放目录为<u>" & ValueStr & "</u>",0)
	Else
		SetupRID_1050(0,0) = GBL_LeadBBS_Setup_Data(2,0)
		CALL Update_ECHO("获取备份文件目录：<u>" & SetupRID_1050(0,0) & "</u>",0)
	End If
	SetupRID_1050(1,0) = ""
	SetupRID_1050(2,0) = "临时备份文件存放目录"
	Update_CreateFolder(DEF_BBS_HomeUrl & SetupRID_1050(0,0) & "/")
	
	If Update_CheckSetupRIDExist(1002," and ClassNum=0") = 0 Then
		RID = 1002
		ValueStr = "20100101001"
		GBL_UpdateVersion = ValueStr
		ClassNum = 0
		saveData = "内部版本号"
		CALL Update_InsertSetupRID(RID,ValueStr,ClassNum,saveData," and ClassNum=0")
		CALL Update_ECHO("初始化内部版本号为<u>" & ValueStr & "</u>",0)
	Else
		GBL_UpdateVersion = cCur(GBL_LeadBBS_Setup_Data(2,0))
		CALL Update_ECHO("获取内部版本号：<u>" & GBL_UpdateVersion & "</u>",0)
	End If
	
	'获取文件配置
	If SubmitFlag = "" Then CALL Update_ECHO("<div class=alertdone>获取论坛配置信息。。。</div>",1)

	CurN = 1
	CALL Update_GetFileParaValue("帖子相关$$$$:$a/a2.asp","Const LMT_EnableOtherGuestName",CurN,"开放论坛是否允许使用""游客""以外的名字")
	CALL Update_GetFileParaValue("a/a2.asp","Const LMT_BuyAnnounceMaxPoints",CurN,"购买帖消耗的最大积分")
	CALL Update_GetFileParaValue("a/a2.asp","Const LMTDEF_MaxReAnnounce",CurN,"允许的最大回复帖数，仅在允许树形下限制才会有效")
	CALL Update_GetFileParaValue("a/a2.asp","Const LMTDEF_MinAnnounceLength",CurN,"发帖需要最少字数")
	CALL Update_GetFileParaValue("a/a2.asp","Const LMTDEF_NotReplyDate",CurN,"最后回复时间至今高于多少天的帖子则禁止回复,对版主及以上无效")
	CALL Update_GetFileParaValue("a/a2.asp","Const LMTDEF_NeedCachetValue",CurN,"设定多少威望用户可以自己归类专题")
	CALL Update_GetFileParaValue("a/a2.asp","Const LMTDEF_ColorSpend",CurN,"设定帖子颜色消耗多少财富值")
	CALL Update_GetFileParaValue("a/a2.asp","Const LMTDEF_RepostMsg",CurN,"回复帖子是否默认短消息通知帖主,0．默认不通知 1.回复全部通知 2.仅被引用时才通知")
	
	CALL Update_GetFileParaValue("a/a.asp","Const LMT_RefreshEnable",CurN,"用户重复浏览帖子是否计算浏览量")
	CALL Update_GetFileParaValue("a/a.asp","Const LMTDEF_RepostMsg",CurN,"回复帖子是否默认短消息通知帖主,0．默认不通知 1.回复全部通知(注意与回帖(a2.asp文件)设置保持一直)")
	
	CALL Update_GetFileParaValue("a/Editannounce.asp","Const LMTDEF_MinAnnounceLength",CurN,"编辑提交的帖子内容需要最少字数")
	CALL Update_GetFileParaValue("a/Editannounce.asp","Const LMT_BuyAnnounceMaxPoints",CurN,"购买帖消耗的最大积分")
	CALL Update_GetFileParaValue("a/Editannounce.asp","Const LMTDEF_NeedCachetValue",CurN,"设定多少声望用户可以自己归类专题")
	
	CALL Update_GetFileParaValue("附件下载方式$$$$:$a/file.asp","Const LMT_RedirectFile",CurN,"附件显示方式：0,读取下载，隐藏真实地址但性能稍差 1.转址下载 高性能但暴露真实地址")
	
	CALL Update_GetFileParaValue("收藏帖子数$$$$:$a/Processor.asp","Const LMT_MaxCollectAnnounce",CurN,"最多允许收藏帖子数量")
	CALL Update_GetFileParaValue("短消息通知$$$$:$a/Processor.asp","Const LMT_Prc_anonymity",CurN,"管理者是否匿名短消息通知用户： 0 匿名为系统 1 原操作人")
	CALL Update_GetFileParaValue("a/Processor.asp","Const LMT_Prc_MsgFlag",CurN,"管理员是否短消息通知用户： 0 默认选项为不通知,但可选择是否通知 1 默认短消息通知,也可选择是否通知 2.强制短消息通知,不可是否通知")
	
	CALL Update_GetFileParaValue("最大好友数$$$$:$a/inc/AddFriend.asp","Const LMT_MaxFriendNum",CurN,"允许添加的最多好友数目")

	CALL Update_GetFileParaValue("红包帖子$$$$:$a/inc/DelAnnounce.asp","AncIDStr = ",CurN,"红包主题ID列表，逗号分隔，回复此类帖子将奖励随机声望(1-3)，并且此类帖子将禁止删除回复(但可编辑)")
	CALL Update_GetFileParaValue("a/a2.asp","AncIDStr = ",CurN,"红包帖子主题ID列表，逗号分隔，回复此类帖子将奖励随机声望(1-3)，注意与[DelAnnounce.asp]配置保持一致")
	
	CALL Update_GetFileParaValue("默认发帖模式$$$$:$a/inc/Editor_Fun.asp","Const Edt_MiniMode",CurN,"发帖界面：0-传统简约模式 1.多功能模式")
	
	CALL Update_GetFileParaValue("多媒体播放个数及是否自动播放$$$$:$a/inc/leadcode.js","var vnum = ",CurN,"1  forbid play,-2 allow 3 video to play at same time. 0: allow one")
	CALL Update_GetFileParaValue("a/inc/leadcode.js","var autoplay = ",CurN,"0.manual play 1.auto play")
	
	CALL Update_GetFileParaValue("评价帖子设置$$$$:$a/inc/MakeGoodAnnounce.asp","Const DEF_AllowPunish",CurN,"是否允许普通用户惩罚发帖用户：1.允许普遍用户惩罚发帖用户　０。禁止")
	CALL Update_GetFileParaValue("a/inc/MakeGoodAnnounce.asp","Const DEF_AllowOpinionNum",CurN,"允许普通用户评价次数 0,禁止,-1 允许无限 >0 指定次数")
	CALL Update_GetFileParaValue("a/inc/MakeGoodAnnounce.asp","Const DEF_MasterNolimit",CurN,"版主及管理员评价次数是否无限：　１，无限，０，限制同普通用户次")
	CALL Update_GetFileParaValue("a/inc/MakeGoodAnnounce.asp","Const DEF_AllowBoardMasterCachetValue",CurN,"是否允许版主评价声望：1.是 0.否")
	
	CALL Update_GetFileParaValue("投票限制$$$$:$a/inc/Poll_fun.asp","Const LMT_PollNeedPoints",CurN,"用户投票帖子需要达到的积分，可以为负。")
	
	CALL Update_GetFileParaValue("RSS$$$$:$other/RSS.asp","Const RSS_ViewNumer",CurN,"最多允许显示的RSS记录条数")
	
	CALL Update_GetFileParaValue("搜索$$$$:$Search/Search.asp","Const Sch_AllContent",CurN,"是否允许全部搜索,即同时搜索标题和内容，设为99表示采用hubbledotnet引擎ajax调用搜索，设为98采用组件方式调用hubbledotnet搜索")
	CALL Update_GetFileParaValue("Search/Search.asp","Const Sch_AncTitle",CurN,"是否允许帖子标题搜索")
	CALL Update_GetFileParaValue("Search/Search.asp","Const Sch_AncContent",CurN,"是否允许帖子内容搜索")
	CALL Update_GetFileParaValue("Search/Search.asp","Const Sch_LimitTime",CurN,"限制搜索时间(单位秒)")
	
	CALL Update_GetFileParaValue("Search/inc/Search_fun.asp","Const DEF_BBS_MaxListPage",CurN,"搜索结果最多显示页数(过大可能影响性能，默认请设为10)")
	CALL Update_GetFileParaValue("Search/inc/Search_fun.asp","Const DEF_BBS_MaxWords",CurN,"搜索结果的帖子内容略要显示长度(最多显示字节)")
	
	CALL Update_GetFileParaValue("短消息$$$$:$User/LookMessage.asp","Const LMT_LookedMsgExpiresDay",CurN,"短消息阅读后的保存期限(单位天)")
	CALL Update_GetFileParaValue("User/SendMessage.asp","Const DEF_User_MaxReceiveUser",CurN,"定义允许同时发送短消息给多少个用户，默认值为5")
	
	CALL Update_GetFileParaValue("支付宝$$$$:$User/alipay/alipay_Config.asp","partner = ",CurN,"支付宝获取id，您先需要一个支付宝账号，再从相应网址获取id(<a href=https://www.alipay.com/himalayas/practicality_customer.htm?customer_external_id=C4335329546596834111&market_type=from_agent_contract&pro_codes=F7F62F29651356BB target=_blank>点此获取</a>)")
	CALL Update_GetFileParaValue("User/alipay/alipay_Config.asp","key = ",CurN,"支付宝获取的密钥，您先需要一个支付宝账号，再从相应网址获取密钥(<a href=https://www.alipay.com/himalayas/practicality_customer.htm?customer_external_id=C4335329546596834111&market_type=from_agent_contract&pro_codes=F7F62F29651356BB target=_blank>点此获取</a>)")
	
	CALL Update_GetFileParaValue("短消息$$$$:$User/inc/Fun_SendMessage.asp","Const LMT_SendMsgExpiresDate",CurN,"定义新发送短消息保存天数(过期自动删除)")
	CALL Update_GetFileParaValue("User/inc/UserTopic.asp","Const LMT_MaxMessageNumber",CurN,"用户收件箱允许的最多接收记录，超过将无法接收新消息。")
	
	CALL Update_GetFileParaValue("多媒体重复播放设置$$$$:$a/inc/leadcode.js","var playcount = ",CurN,"play loop count:0-100,0=always replay")
	
	CALL Update_GetFileParaValue("注册验证设置$$$$:$User/" & DEF_RegisterFile,"Const LMT_RegVerifyQuestion = ",CurN,"注册验证提示信息，可以是HTML格式，比如使用图片，若不填写表示不开启注册验证信息。")
	CALL Update_GetFileParaValue("User/" & DEF_RegisterFile,"Const LMT_RegVerifyAnswer = ",CurN,"注册验证需要填写的答案。")
	
	CALL Update_GetFileParaValue("QQ互联设置$$$$:$app/qqlogin/oauth.asp","Const apiKey = ",CurN,"APP ID,您需要从腾讯平台申请获取资料：(<a href=http://connect.qq.com/ target=_blank>点此申请</a>)")
	CALL Update_GetFileParaValue("app/qqlogin/oauth.asp","Const secretKey = ",CurN,"APP KEY,您需要从腾讯平台申请获取")
	CALL Update_GetFileParaValue("app/qqlogin/oauth.asp","Const callback = ",CurN,"CALL BACK,回调地址，注意只需要填写域名，不包括http及目录。")
	
	CALL Update_GetFileParaValue("帖内分享代码设置$$$$:$a/a.asp$$$$:$textarea","Const LMTDEF_ShareID = ",CurN,"可以填写各站或自行编写类型的分享代码(HTML格式，注意手工删除换行符),保持为空则关闭分享代码;")
	GBL_ParaCount = CurN - 1

	'保证或更新配置
	If SubmitFlag = "" Then CALL Update_ECHO("<div class=alertdone>检测并保存配置。。。</div>",1)
	Dim N,TmpNewStr,TmpNewStr2
	Dim filename,tmp,title
	
	For N = 1 to GBL_ParaCount
		If inStr(SetupRID_1050(1,N),"$$$$:$") Then
			tmp = Split(SetupRID_1050(1,N),"$$$$:$")
			title = tmp(0)
			filename = tmp(1)
		else
			filename = SetupRID_1050(1,N)
		End If
	
			If Update_CheckSetupRIDExist(1050," and ClassNum=" & N) = 0 Then
				RID = 1050
				ValueStr = SetupRID_1050(0,N)
				ClassNum = N
				saveData = filename & " | " & SetupRID_1050(2,N)
				CALL Update_InsertSetupRID(RID,ValueStr,ClassNum,saveData," and ClassNum=" & N)
				If SubmitFlag = "" Then CALL Update_ECHO("配置项" & N & "(<span class=grayfont>" & SetupRID_1050(2,N) & "</span>)存储完成，值为：<u>" & ValueStr & "</u>",0)
			Else
				If SetupRID_1050(0,N) <> GBL_LeadBBS_Setup_Data(2,0) Then
					If GBL_LeadBBS_Setup_Data(2,0) = "error" and SetupRID_1050(0,N) <> "error" Then
						GBL_LeadBBS_Setup_Data(2,0) = SetupRID_1050(0,N)
						RID = 1050
						ValueStr = SetupRID_1050(0,N)
						ClassNum = N
						saveData = filename & " | " & SetupRID_1050(2,N)
						CALL Update_InsertSetupRID(RID,ValueStr,ClassNum,saveData," and ClassNum=" & N)
					End If
					CALL Update_ECHO("当前配置项<u>" & N & "</u>(<span class=grayfont>" & SetupRID_1050(2,N) & "</span>)与实际不符，已从存储数据读取并更新现用配置。",1)
					
					If Right(SetupRID_1050(4,N),3) = " = " or Right(SetupRID_1050(4,N),2) = "= " or Right(SetupRID_1050(4,N),1) = "=" Then
						TmpNewStr2 = ""
					Else
						TmpNewStr2 = " = "
					End If
					SetupRID_1050(0,N) = GBL_LeadBBS_Setup_Data(2,0)
					If LCase(Right(filename,3)) = ".js" Then
						TmpNewStr = SetupRID_1050(4,N) & TmpNewStr2 & SetupRID_1050(0,N) & ";"
						If SetupRID_1050(2,N) <> "" Then TmpNewStr = TmpNewStr & " //" & SetupRID_1050(2,N)
					Else
						TmpNewStr = SetupRID_1050(4,N) & TmpNewStr2 & SetupRID_1050(0,N) & ""
						If SetupRID_1050(2,N) <> "" Then TmpNewStr = TmpNewStr & " '" & SetupRID_1050(2,N)
					End If
					If Right(SetupRID_1050(3,N),2) = VbCrLf and Right(TmpNewStr,2) <> VbCrLf Then TmpNewStr = TmpNewStr & VbCrLf
					CALL Update_ReplaceFileStr(filename,SetupRID_1050(3,N),TmpNewStr)
				Else
					If SubmitFlag = "" Then CALL Update_ECHO("当前配置项<u>" & N & "</u>确认无误。(<span class=grayfont>" & SetupRID_1050(2,N) & "</span>)。",0)
				End If
			End If
	Next
	
	'检测并保存BBSSetup.asp, Ubbcode_Setup.asp,User_Setup.ASP,Upload_Setup.asp,AD_Data.asp 
	'检测并保存User/inc/Contact_info.asp User_Reg.asp
	CALL Update_ECHO("<div class=alertdone>检测并保存配置文件。。。</div>",1)
	Dim FileSetupData
	FileSetupData = Array("inc/BBSSetup.asp", "inc/Ubbcode_Setup.asp","inc/User_Setup.ASP","inc/Upload_Setup.asp","inc/AD_Data.asp","User/inc/Contact_info.asp","User/inc/User_Reg.asp")
	Dim FileContent
	
	For N = 0 to Ubound(FileSetupData,1)
		If Update_CheckSetupRIDExist(1051," and ClassNum=" & N) = 0 Then
			RID = 1051
			ValueStr = FileSetupData(N)
			ClassNum = N
			saveData = ADODB_LoadFile(DEF_BBS_HomeUrl & FileSetupData(N))
			CALL Update_InsertSetupRID(RID,ValueStr,ClassNum,saveData," and ClassNum=" & N)
			If SubmitFlag = "" Then CALL Update_ECHO("配置文件(<span class=grayfont>" & FileSetupData(N)& "</span>)已保存。",0)
		Else
			FileContent = ADODB_LoadFile(DEF_BBS_HomeUrl & FileSetupData(N))
			If FileContent <> GBL_LeadBBS_Setup_Data(4,0) Then
				ADODB_SaveToFile GBL_LeadBBS_Setup_Data(4,0),DEF_BBS_HomeUrl & FileSetupData(N)
				If SubmitFlag = "" Then CALL Update_ECHO("配置文件<u>" & FileSetupData(N) & "</u>与存储数据不符，当前配置已完成更新。",1)
			Else
				If SubmitFlag = "" Then CALL Update_ECHO("配置文件<u>" & FileSetupData(N) & "</u>完成检测。",0)
			End If
		End If
	Next
	
	'rem Licence Save
	Dim Licence
	Licence = Update_GetLicence
	If Update_CheckSetupRIDExist(1001,"") = 0 Then
		If Licence <> "error" Then
			RID = 1001
			ValueStr = Licence
			ClassNum = 0
			saveData = ""
			CALL Update_InsertSetupRID(RID,ValueStr,ClassNum,saveData,"")
		End If
	Else
		HeadStr = "leadbbs.com/other/register?"
		If Licence = "error" or Licence <> GBL_LeadBBS_Setup_Data(2,0) Then
			Dim HeadStr
			If Licence = "error" Then
				Licence = ""
				HeadStr = "leadbbs.com/other/register?"
			Else
				HeadStr = "leadbbs.com/other/register?"
			End If
			If GBL_LeadBBS_Setup_Data(2,0) <> "" and GBL_LeadBBS_Setup_Data(2,0) <> "error" and DEF_UsedDataBase <> 1 Then
				CALL Update_ReplaceFileStr("inc/Board_Popfun.asp",HeadStr & Licence & """>licence</a>",HeadStr & GBL_LeadBBS_Setup_Data(2,0) & """>licence</a>")
				CALL Update_ReplaceFileStr("inc/Board_Popfun.asp",HeadStr & Licence & """></a>",HeadStr & GBL_LeadBBS_Setup_Data(2,0) & """>licence</a>")
			Else
				RID = 1001
				ValueStr = ""
				ClassNum = 0
				saveData = ""
				CALL Update_InsertSetupRID(RID,ValueStr,ClassNum,saveData,"")
				CALL Update_ReplaceFileStr("inc/Board_Popfun.asp",HeadStr & Licence & """>licence</a>",HeadStr & "" & """></a>")
				CALL Update_ReplaceFileStr("inc/Board_Popfun.asp",HeadStr & Licence & """></a>",HeadStr & "" & """></a>")
			End If
		Else
			If DEF_UsedDataBase <> 0 Then
				RID = 1001
				ValueStr = ""
				ClassNum = 0
				saveData = ""
				CALL Update_InsertSetupRID(RID,ValueStr,ClassNum,saveData,"")
				CALL Update_ReplaceFileStr("inc/Board_Popfun.asp",HeadStr & Licence & """>licence</a>",HeadStr & "" & """></a>")
				CALL Update_ReplaceFileStr("inc/Board_Popfun.asp",HeadStr & Licence & """></a>",HeadStr & "" & """></a>")
			End If
		End If
	End If

	'手动配置参数
	If SubmitFlag <> "" and LCase(Request("checkversion")) <> "checkversion" and LCase(Request("checkversion")) <> "updateversion" Then
		Update_SetupFilePara
		Exit Sub
	End If

End Sub

Function Update_GetLicence

	Dim FileName,Str
	FileName = "inc/Board_Popfun.asp"
	Str = "leadbbs.com/other/register"
	Dim fs,WriteFile,fileContent,ID,Tmp,FSFlag
	If FSFlag = 1 Then
		Set fs = Server.CreateObject(DEF_FSOString)
		Set WriteFile = fs.OpenTextFile(Server.MapPath(DEF_BBS_HomeUrl & FileName),1,True)
		If Not WriteFile.AtEndOfStream Then
			fileContent = WriteFile.ReadAll
		End If
		WriteFile.Close
		Set fs = Nothing
	Else
		fileContent = ADODB_LoadFile(DEF_BBS_HomeUrl & FileName)
	End If
	
	Tmp = InStr(LCase(fileContent),LCase(Str))
	If Tmp < 1 Then
		GBL_Update_LineStr = "@@@@@@nothing$string@@@@@@@@@@@@"
		Update_GetLicence = "error"
		Exit Function
	End If

	GBL_Update_LineStr = Mid(fileContent,Tmp,300)
	
	Dim BottomStr
	BottomStr = VbCrLf
	If inStr(GBL_Update_LineStr,BottomStr) > 2 Then
		GBL_Update_LineStr = Left(GBL_Update_LineStr,inStr(GBL_Update_LineStr,BottomStr)-2)
	Else
		GBL_Update_LineStr = "@@@@@@nothing$string@@@@@@@@@@@@"
		Update_GetLicence = "error"
		Exit Function
	End If
	
	
	Tmp = InStr(GBL_Update_LineStr,"?")
	If Tmp < 1 Then
		GBL_Update_LineStr = "@@@@@@nothing$string@@@@@@@@@@@@"
		Update_GetLicence = "error"
		Exit Function
	End If

	Dim Tmp2
	Tmp2 = Mid(GBL_Update_LineStr,Tmp+1)
	
	Tmp2 = Left(Tmp2,32)
	
	ID = Trim(Tmp2)
	Do while Right(ID,1) = "	"
		ID = Left(ID,Len(ID)-1)
	Loop
	Do while Left(ID,1) = "	"
		ID = Right(ID,Len(ID)-1)
	Loop
	ID = Trim(Tmp2)
	Tmp2 = ID
	Dim allowstr,N
	allowstr = Array("a","b","c","d","e","f","1","2","3","4","5","6","7","8","9","0")
	For N = 0 to Ubound(allowstr,1)
		Tmp2 = Replace(Tmp2,allowstr(N),"")
	Next
	If Tmp2 = "" and Len(ID) = 32 Then
		Update_GetLicence = ID
	Else
		Update_GetLicence = "error"
		GBL_Update_LineStr = "@@@@@@nothing$string@@@@@@@@@@@@"
	End If

End Function

Sub Update_GetFileParaValue(f,fStr,Index,Note)

	Dim fileStr

	Dim filename,tmp,title
		If inStr(f,"$$$$:$") Then
			tmp = Split(f,"$$$$:$")
			title = tmp(0)
			filename = tmp(1)
		else
			filename = f
		End If
	
	fileStr = fStr
	SetupRID_1050(0,Index) = Update_CheckFileInStr(fileName,fileStr)
	If SetupRID_1050(0,Index) = "error" Then
		SetupRID_1050(1,Index) = f
		SetupRID_1050(2,Index) = Note
		SetupRID_1050(3,Index) = GBL_Update_LineStr
		SetupRID_1050(4,Index) = fStr
		CALL Update_ECHO("获取配置" & Index & "(<u>" & fileName & "/" & fileStr & "</u>)失败：请从官方重新下载原文件更新并替换<u>" & fileName & "</u>.",1)
	Else
		SetupRID_1050(1,Index) = f
		SetupRID_1050(2,Index) = Note
		SetupRID_1050(3,Index) = GBL_Update_LineStr
		SetupRID_1050(4,Index) = fStr
		If SubmitFlag = "" Then CALL Update_ECHO("提取配置" & Index & "完成。",0)
	End If
	CurN = CurN + 1

End Sub

Sub Update_ReplaceFileStr(FileName,OldStr,NewStr)

	Dim fs,WriteFile,fileContent
	If FSFlag = 1 Then
		Set fs = Server.CreateObject(DEF_FSOString)
		Set WriteFile = fs.OpenTextFile(Server.MapPath(DEF_BBS_HomeUrl & FileName),1,True)
		If Not WriteFile.AtEndOfStream Then
			fileContent = WriteFile.ReadAll
		End If
		WriteFile.Close
		Set fs = Nothing

		If OldStr = "" Then
			fileContent = fileContent & NewStr
		Else
			fileContent = Replace(fileContent,OldStr,NewStr)
		End If
		Set fs = Server.CreateObject(DEF_FSOString)
		Set WriteFile = fs.CreateTextFile(Server.MapPath(DEF_BBS_HomeUrl & FileName),True)
		WriteFile.Write fileContent
		WriteFile.Close
		Set fs = Nothing
	Else
		fileContent = ADODB_LoadFile(DEF_BBS_HomeUrl & FileName)
		If OldStr = "" Then
			fileContent = fileContent & NewStr
		Else
			fileContent = Replace(fileContent,OldStr,NewStr)
		End If
		ADODB_SaveToFile fileContent,DEF_BBS_HomeUrl & FileName
		Response.Write GBL_CHK_TempStr
	End If

End Sub

Function Update_CheckSetupRIDExist(RID,extend)

	Dim Rs
	Set Rs = LDExeCute(sql_select("Select ID,RID,ValueStr,ClassNum,SaveData from LeadBBS_Setup where RID=" & RID & extend,1),0)
	If Rs.Eof Then
		Update_CheckSetupRIDExist = 0
		Set GBL_LeadBBS_Setup_Data = Nothing
		GBL_LeadBBS_Setup_Data = ""
	Else
		Update_CheckSetupRIDExist = 1
		GBL_LeadBBS_Setup_Data = Rs.GetRows(-1)
		GBL_LeadBBS_Setup_Data(2,0) = Trim(GBL_LeadBBS_Setup_Data(2,0))
	End If
	Rs.Close
	Set Rs = Nothing

End Function

Sub Update_InsertSetupRID(RID,ValueStr,ClassNum,saveData,extend)

	If Update_CheckSetupRIDExist(RID,extend) = 0 Then
		CALL LDExeCute("insert into LeadBBS_Setup(RID,ValueStr,ClassNum,saveData) values(" & Rid & ",'" & Replace(ValueStr,"'","''") & "'," & ClassNum & ",'" & Replace(saveData,"'","''") & "')",1)
	Else
		CALL LDExeCute("Update LeadBBS_Setup Set ValueStr='" & Replace(ValueStr,"'","''") & "',ClassNum=" & ClassNum & ",saveData='" & Replace(saveData,"'","''") & "' where RID=" & RID & extend,1)
	End If

End Sub

Function Update_CreateFolder(folder)

	If FSFlag = 0 Then
		CALL Update_ECHO("空间不支持FSO,目录操作忽略.",1)
		Exit Function
	End If
	Dim FS
	Set FS = Server.CreateObject(DEF_FSOString)
	Dim TDIR
	TDIR = Server.MapPath(Replace(Replace(folder,"/","\"),"\\","\"))
	If Not FS.FolderExists(TDIR) then
		FS.CreateFolder(TDIR)
	End If
	Set FS = Nothing
	
End Function

Function Update_CheckFSO

	On Error Resume Next
	Dim FS
	FSFlag = 1
	If DEF_FSOString <> "" Then
		Err.Clear
		Set FS = Server.CreateObject(DEF_FSOString)
		If Err Then
			FSFlag = 0
			Err.Clear
			Set fs = Nothing
		End If
	Else
		Err.Clear
		Set FS = Server.CreateObject("Scripting.FileSystemObject")
		If Err Then
			FSFlag = 0
			Err.Clear
			Set fs = Nothing
		End If
	End If
	
	If FSFlag = 0 Then
		Update_CheckFSO = 0
	Else
		Update_CheckFSO = 1
	End If

End Function

Function Update_CheckFields(FieldsName,TableName)

	Dim Flag,sql,RS,i
	Flag=False 
	sql=sql_select("select * from "&TableName,1)
	Set RS=LDExeCute(sql,0) 
	for i = 0 to RS.Fields.Count - 1
		If LCase(RS.Fields(i).Name) = LCase(FieldsName) then
			Flag = True
			Exit For
		else
			Flag = False
		end if
	Next
	Update_CheckFields = Flag

End Function

Sub Update_ECHO(str,t)

	If t = 1 Then
		Response.Write "<p style=""color:red"">" & str & "</p>" & VbCrLf
	Else
		Response.Write "<p>" & str & "</p>" & VbCrLf
	End If
	'Response.Flush

End Sub

Function Update_CheckFileInStr(FileName,Str)

	Dim fs,WriteFile,fileContent,ID,Tmp,FSFlag
	If FSFlag = 1 Then
		Set fs = Server.CreateObject(DEF_FSOString)
		Set WriteFile = fs.OpenTextFile(Server.MapPath(DEF_BBS_HomeUrl & FileName),1,True)
		If Not WriteFile.AtEndOfStream Then
			fileContent = WriteFile.ReadAll
		End If
		WriteFile.Close
		Set fs = Nothing
	Else
		fileContent = ADODB_LoadFile(DEF_BBS_HomeUrl & FileName)
	End If
	
	Tmp = InStr(LCase(fileContent),LCase(Str))
	If Tmp < 1 Then
		Update_CheckFileInStr = "error"
		Exit Function
	End If

	GBL_Update_LineStr = Mid(fileContent,Tmp,3000)
	
	Dim BottomStr
	BottomStr = VbCrLf
	If inStr(GBL_Update_LineStr,BottomStr) > 2 Then
		GBL_Update_LineStr = Left(GBL_Update_LineStr,inStr(GBL_Update_LineStr,BottomStr)-1)
	Else
		Update_CheckFileInStr = "error"
		Exit Function
	End If
	
	
	Tmp = InStr(GBL_Update_LineStr,"=")
	If Tmp < 1 Then
		Update_CheckFileInStr = "error"
		Exit Function
	End If
	
	Dim Tmp2
	Tmp2 = Mid(GBL_Update_LineStr,Tmp+1)
	
	Tmp = InStr(Tmp2,"'")
	If Tmp > 1 Then
		Tmp2 = Left(Tmp2,Tmp-1)
	Else
		Tmp = InStr(Tmp2,";")
		If Tmp > 1 Then
			Tmp2 = Left(Tmp2,Tmp-1)
		Else
			Tmp = InStr(Tmp2,"//")
			If Tmp > 2 Then Tmp2 = Left(Tmp2,Tmp-2)
		End If
	End If
	
	ID = Trim(Tmp2)
	Do while Right(ID,1) = "	"
		ID = Left(ID,Len(ID)-1)
	Loop
	Do while Left(ID,1) = "	"
		ID = Right(ID,Len(ID)-1)
	Loop

	Update_CheckFileInStr = ID

End Function




Function ADODB_LoadFile(ByVal File)

	On Error Resume Next
	Dim objStream,FSFlag,fs,WriteFile
	FSFlag = 1
	If DEF_FSOString <> "" Then
		Set fs = Server.CreateObject(DEF_FSOString)
		If Err Then
			FSFlag = 0
			Err.Clear
			Set fs = Nothing
		End If
	Else
		FSFlag = 0
	End If

	If FSFlag = 1 Then
		Set WriteFile = fs.OpenTextFile(Server.MapPath(File),1,True)
		If Err Then
			GBL_CHK_TempStr = "<br>读取文件失败：" & err.description & "<br>其它可能：确定是否对此文件有读取权限."
			err.Clear
			Set Fs = Nothing
			Exit Function
		End If
		If Not WriteFile.AtEndOfStream Then
			ADODB_LoadFile = WriteFile.ReadAll
			If Err Then
				GBL_CHK_TempStr = "读取文件失败：<p>" & err.description & "</p> 其它可能：确定是否对此文件有读取权限."
				err.Clear
				Set Fs = Nothing
				Exit Function
			End If
		End If
		WriteFile.Close
		Set Fs = Nothing
	Else
		Set objStream = Server.CreateObject("ADODB.Stream")
		If Err.Number=-2147221005 Then 
			GBL_CHK_TempStr = "您的主机不支持ADODB.Stream，无法完成操作，请手工进行"
			Err.Clear
			Set objStream = Nothing
			Exit Function
		End If
		With objStream
			.Type = 2
			.Mode = 3
			.Open
			.LoadFromFile Server.MapPath(File)
			.Charset = "gb2312"
			.Position = 2
			ADODB_LoadFile = .ReadText
			.Close
		End With
		Set objStream = Nothing
	End If
	If Err Then
		GBL_CHK_TempStr = "错误信息：<p>" & err.description & "</p>其它可能：确定是否对此文件有读取权限."
		err.Clear
		Set Fs = Nothing
		Exit Function
	End If

End Function

'存储内容到文件
Sub ADODB_SaveToFile(ByVal strBody,ByVal File)

	On Error Resume Next
	Dim objStream,fs,WriteFile
	FSFlag = 1
	If DEF_FSOString <> "" Then
		Set fs = Server.CreateObject(DEF_FSOString)
		If Err Then
			FSFlag = 0
			Err.Clear
			Set fs = Nothing
		End If
	Else
		FSFlag = 0
	End If
	
	If FSFlag = 1 Then
		Set WriteFile = fs.CreateTextFile(Server.MapPath(File),True)
		WriteFile.Write strBody
		WriteFile.Close
		Set Fs = Nothing
	Else
		Set objStream = Server.CreateObject("ADODB.Stream")
		If Err.Number=-2147221005 Then 
			GBL_CHK_TempStr = "您的主机不支持ADODB.Stream，无法完成操作，请手工进行"
			Err.Clear
			Set objStream = Nothing
			Exit Sub
		End If
		With objStream
			.Type = 2
			.Open
			.Charset = "gb2312"
			.Position = objStream.Size
			.WriteText = strBody
			.SaveToFile Server.MapPath(File),2
			.Close
		End With
		Set objStream = Nothing
	End If
	If Err Then
		GBL_CHK_TempStr = "错误信息：<p>" & err.description & "</p>其它可能：确定是否对此文件有写入权限."
		err.Clear
		Set Fs = Nothing
		Exit Sub
	End If

End Sub


'存储内容到文件
Sub ADODB_SaveToFileBinary(ByVal strBody,ByVal File)

	On Error Resume Next
	Dim objStream
	
		Set objStream = Server.CreateObject("ADODB.Stream")
		If Err.Number=-2147221005 Then 
			GBL_CHK_TempStr = "您的主机不支持ADODB.Stream，无法完成操作，请手工进行"
			Err.Clear
			Set objStream = Nothing
			Exit Sub
		End If
		With objStream
			.Type = 1
			.Open
			'.Charset = "gb2312"
			.Position = objStream.Size
			.Write = strBody
			.SaveToFile Server.MapPath(File),2
			.Close
		End With
		Set objStream = Nothing
	If Err Then
		GBL_CHK_TempStr = "错误信息：<p>" & err.description & "</p>其它可能：确定是否对此文件有写入权限."
		err.Clear
		Set Fs = Nothing
		Exit Sub
	End If
	Response.Write "<span class=grayfont>文件长度:" & LenB(strBody) & " Bytes</span>"

End Sub

Sub Update_SetupFilePara

%>
<form name="pollform3sdx" method="post" action="Update.asp">
<input type="hidden" name="SubmitFlag" value=yes>
<input type="hidden" name="sure" value=1>
<br />
<p>
		<b>设置：<span class=grayfont>论坛扩展参数设置</span></b>
		<br>
		<span class=grayfont>(下面为网站参数，请注意修改，错误的设置将会发生严重错误)<br><br>
		请参考注释修改参数，<span class=redfont>一些设置值为字符串的，注意保留单角双引号</font>。</span>
</p>
<%
If Request.Form("SubmitFlag") = "yes" Then
	Update_SetupFilePara_CheckLinkValue
End If%>
<b><span class=redfont><%=GBL_CHK_TempStr%></span></b>
<%
If Request.Form("SubmitFlag") = "yes" Then
	If GBL_CHK_TempStr <> "" Then
		Update_SetupFilePara_Form
	Else
		Update_SetupFilePara_RefreshValue
		Exit Sub
	End If
Else
	Update_SetupFilePara_Form
End If
%>
</form>
<%
End Sub

Sub Update_SetupFilePara_Form

%>

	<table border=0 cellpadding=0 cellspacing=0 width="100%" class=frame_table>
	<%
	Dim N
	Dim title,filename,tmp

	For N = 1 to GBL_ParaCount
		title = ""
		If inStr(SetupRID_1050(1,N),"$$$$:$") Then
			tmp = Split(SetupRID_1050(1,N),"$$$$:$")
			title = tmp(0)
			filename = tmp(1)
		End If	
		If title <> "" Then%>
		<tr>
			<td class=tdbox width=90>&nbsp;</td>
			<td class=tdbox><b><%=title%></b></td>
		</tr><%
		End If
		%>
		<tr>
			<td class=tdbox width=90>配置项<%=N%></td>
			<td class=tdbox><input class=fminpt type="text" name="Form_SetupRID_<%=N%>" maxlength="2048" size="45" value="<%=server.htmlencode(SetupRID_1050(0,N))%>"><br /> <span class=note><%=SetupRID_1050(2,N)%></span></td>
		</tr>
		<tr>
	<%
	Next%>
	<td class=tdbox>&nbsp;</td>
	<td class=tdbox>
		<input type=submit name=提交 value=提交 class=fmbtn>
		<input type=reset name=取消 value=取消 class=fmbtn>
	</td>
	</tr>
	</table>
<%
End Sub

Sub Update_SetupFilePara_CheckLinkValue

	Dim Val,N,Tmp
	For N = 1 to GBL_ParaCount
		Val = Request("Form_SetupRID_" & N)
		Val = Replace(Replace(Val,chr(13),""),chr(10),"")
		If Val = "" Then
			GBL_CHK_TempStr = "配置项" & N & " 必须填写"
			Exit Sub
		End If
		If inStr(Val,"<" & "%") > 0 or inStr(Val,"%" & ">") > 0 Then
			GBL_CHK_TempStr = "配置项" & N & " 填写错误，不能包括一些屏蔽字串。"
			Exit Sub
		End If
		
		If inStr(SetupRID_1050(0,N),"""") > 0 Then
			If Len(Val) > 1024 Then
				GBL_CHK_TempStr = "配置项" & N & " 过长。"
				Exit Sub
			End If
			If Left(Val,1) <> """" or Right(Val,1) <> """" Then
				GBL_CHK_TempStr = "配置项" & N & " 错误，此值为字符串，必须前后使用单引号。"
				Exit Sub
			End If
			tmp = left(Val,len(Val)-1)
			tmp = right(tmp,len(tmp)-1)
			tmp = Replace(tmp,"""""","@______iei2967z")
			tmp = Replace(tmp,"""","""""")
			tmp = Replace(tmp,"@______iei2967z","""""")
			Val = """" & tmp & """"
		Else
			If isNumeric(Val) = 0 or Len(Val) > 12 Then
				GBL_CHK_TempStr = "配置项" & N & " 错误，此值必须为正确的数字。"
				Exit Sub
			End If
			Val = cCur(Val)
		End If
		SetupRID_1050(0,N) = Val
	Next

End Sub

Sub Update_SetupFilePara_RefreshValue

	Dim N,TmpNewStr,TmpNewStr2
	Dim RID,ValueStr,ClassNum,saveData
	Dim filename,tmp,title
	
	For N = 1 to GBL_ParaCount
		If inStr(SetupRID_1050(1,N),"$$$$:$") Then
			tmp = Split(SetupRID_1050(1,N),"$$$$:$")
			title = tmp(0)
			filename = tmp(1)
		else
			filename = SetupRID_1050(1,N)
		End If
		If Right(SetupRID_1050(4,N),3) = " = " or Right(SetupRID_1050(4,N),2) = "= " or Right(SetupRID_1050(4,N),1) = "=" Then
			TmpNewStr2 = ""
		Else
			TmpNewStr2 = " = "
		End If
		If LCase(Right(filename,3)) = ".js" Then
			TmpNewStr = SetupRID_1050(4,N) & TmpNewStr2 & SetupRID_1050(0,N) & ";"
			If SetupRID_1050(2,N) <> "" Then TmpNewStr = TmpNewStr & " //" & SetupRID_1050(2,N)
		Else
			TmpNewStr = SetupRID_1050(4,N) & TmpNewStr2 & SetupRID_1050(0,N) & ""
			If SetupRID_1050(2,N) <> "" Then TmpNewStr = TmpNewStr & " '" & SetupRID_1050(2,N)
		End If
		If Right(SetupRID_1050(3,N),2) = VbCrLf and Right(TmpNewStr,2) <> VbCrLf Then TmpNewStr = TmpNewStr & VbCrLf
		CALL Update_ReplaceFileStr(filename,SetupRID_1050(3,N),TmpNewStr)
		
		RID = 1050
		ValueStr = SetupRID_1050(0,N)
		SetupRID_1050(0,0) = ValueStr
		ClassNum = N
		saveData = filename & " | " & SetupRID_1050(2,N)
		CALL Update_InsertSetupRID(RID,ValueStr,ClassNum,saveData," and ClassNum=" & N)
	Next
	CALL Update_ECHO("<br /><b><font color=green>成功完成手动配置。</b></font>",1)

End Sub


Const NetFlag = 1
Const NetUrl = "http://update.u1.leadbbs.com/"
Const NativeDir = "Download/"
Const SplitString = "---NdetVeL---"
Const CheckEndString = "LeadBBS_^93857855287569"

Sub Update_CheckVersion
	
	Dim Update,CurFile,CurFile_Name,CurFile_Intro
	Dim FileList
	Dim m
	If NetFlag = 0 Then
		Update = ADODB_LoadFile(NativeDir & "update.txt")
	Else
		Update = BytesToBstr(Update_GetInternetFile(NetUrl & "update.txt"))
		
		If Right(Update,Len(CheckEndString)) <> CheckEndString Then
			CALL Update_ECHO("<div class=alert>可能因网络问题无法连接更新服务器，操作中止。</div>",0)
			Exit Sub
		End If
	End If
	Update = Split(Update,VbCrLf)
	
	CALL Update_ECHO("<div class=alertdone>开始检测是否有补丁更新。。。</div>",0)
	Dim UpdateFlag
	UpdateFlag = 0
	For M = 0 to Ubound(Update,1)
		If Trim(Update(M)) <> "" Then
			If inStr(Update(M),SplitString) > 0 Then
				CurFile = Split(Update(M),SplitString)
				CurFile_Name = CurFile(0)
				CurFile_Intro = " (" & CurFile(1) & ")"
			Else
				CurFile_Name = Update(M)
				CurFile_Intro = ""
			End If
			
			If isNumeric(CurFile_Name) = 0 Then CurFile_Name = 0
			CurFile_Name = cCur(CurFile_Name)
			If CurFile_Name > cCur(GBL_UpdateVersion) Then
				CALL Update_ECHO("<div class=alertdone>检测到新补丁<u>" & CurFile_Name & "</u><span class=redfont>" & CurFile_Intro & "</span></div>",0)
				UpdateFlag = 1
			End If
		End If
	Next
	If UpdateFlag = 0 Then
		CALL Update_ECHO("<div class=alertdone>检测结束，您的论坛已是最新版本，无需更新。</div>",0)
	Else
		CALL Update_ECHO("<div class=alertdone>检测完成，请点击左栏的补丁更新开始更新。</div>",0)
	End If

End Sub
	
Sub Update62_CopyFile
	
	Dim Update,CurFile,CurFile_Name,CurFile_Intro
	Dim FileList
	Dim m
	If NetFlag = 0 Then
		Update = ADODB_LoadFile(NativeDir & "update.txt")
	Else
		Update = BytesToBstr(Update_GetInternetFile(NetUrl & "update.txt"))
		
		If Right(Update,Len(CheckEndString)) <> CheckEndString Then
			CALL Update_ECHO("<div class=alert>可能因网络问题无法连接更新服务器，操作中止。</div>",0)
			Exit Sub
		End If
	End If
	Update = Split(Update,VbCrLf)
	
	CALL Update_ECHO("<div class=alertdone>开始检测新补丁并准备补丁更新。。。</div>",0)
	Dim UpdateFlag
	UpdateFlag = 0
	For M = 0 to Ubound(Update,1)
		If Trim(Update(M)) <> "" Then
			If inStr(Update(M),SplitString) > 0 Then
				CurFile = Split(Update(M),SplitString)
				CurFile_Name = CurFile(0)
				CurFile_Intro = " (" & CurFile(1) & ")"
			Else
				CurFile_Name = Update(M)
				CurFile_Intro = ""
			End If
			
			If isNumeric(CurFile_Name) = 0 Then CurFile_Name = 0
			CurFile_Name = cCur(CurFile_Name)
			If CurFile_Name > cCur(GBL_UpdateVersion) Then
				CALL Update_ECHO("<div class=alertdone>更新补丁<u>" & CurFile_Name & "</u><span class=grayfont>" & CurFile_Intro & "</span>。</div>",0)
				If NetFlag = 0 Then
					FileList = ADODB_LoadFile(NativeDir & CurFile_Name & ".txt")
				Else
					FileList = BytesToBstr(Update_GetInternetFile(NetUrl & CurFile_Name & ".txt"))
				End If
				Update_ExeCuteCopyFIle(FileList)
				
				Dim RID,ValueStr,ClassNum,saveData
				GBL_UpdateVersion = CurFile_Name
				RID = 1002
				ValueStr = GBL_UpdateVersion
				ClassNum = 0
				saveData = "内部版本号"
				CALL Update_InsertSetupRID(RID,ValueStr,ClassNum,saveData," and ClassNum=0")
				CALL Update_ECHO("初始化内部版本号为<u>" & ValueStr & "</u>",0)
				If UpdateFlag = 0 Then Update_CloseSite
				UpdateFlag = 1
				
				If Update_UpdateFileFlag = 1 Then
					Update_OpenSite
					CloseDatabase
					CALL Update_ECHO("<div class=alert>更新模块获得更新，更新强制终止，可点击右侧更新功能继续版本更新。</div>",0)
					Update_PageBottom
					Response.End
				End If
			End If
		End If
	Next
	If UpdateFlag = 0 Then
		CALL Update_ECHO("<div class=alertdone>经检测您的论坛已是最新版本，无需更新。</div>",0)
	Else
		Update_OpenSite
		CALL Update_ECHO("<div class=alertdone>补丁已完成更新，下面重新开始检测论坛配置。</div>",0)
		Update62_initBBSdata
	End If

End Sub

Sub Update_CloseSite

	Application.Lock
	application(DEF_MasterCookies & "SiteEnableFlagzoieiu") = 0
	application(DEF_MasterCookies & "SiteDisbleWhyszoieiu") = "<html><body>论坛处于升级状态，请稍候访问。若长时间无法访问，请联系管理员。</body></html>"
	Application.UnLock

End Sub

Sub Update_OpenSite


	Application.Lock
	application(DEF_MasterCookies & "SiteEnableFlagzoieiu") = 1
	application(DEF_MasterCookies & "SiteDisbleWhyszoieiu") = ""
	Application.UnLock

End Sub

Sub Update_ExeCuteCopyFIle(str)

	Dim GData,GDataFlag
	Dim ListIndex,N,count
	ListIndex = Split(str,VbCrLf)
	count = Ubound(ListIndex,1)
	Dim LineCommandArray,LineCommand,LineDir
	
	Dim thisUrl
	thisUrl = Request.ServerVariables("server_name")
	If Request.ServerVariables("SERVER_PORT") <> "80" Then thisUrl = thisUrl & ":" & Request.ServerVariables("SERVER_PORT")
	thisUrl = "http://" & thisUrl & Replace(Request.Servervariables("SCRIPT_NAME"),"Update.asp","")
	
	Dim Extend,LineDir_Bak,NoteInfo
	For N = 0 to count
		ListIndex(N) = Replace(Trim(ListIndex(N)),"	"," ")
		Do while inStr(ListIndex(N),"  ")
			ListIndex(N) = Replace(ListIndex,"  "," ")
		Loop
		If inStr(ListIndex(N)," ") Then
			LineCommandArray = Split(ListIndex(N)," ")
			LineCommand = LineCommandArray(0)
			LineDir = LCase(Replace(LineCommandArray(1),"/","\"))
			LineDir_Bak = LineDir
			If Left(LineDir,7) = "manage\" and LCase(DEF_ManageDir) <> "manage" Then LineDir = DEF_ManageDir & Mid(LineDir,7)
			If Right(LineDir,13) = "\register.asp" Then LineDir = Mid(LineDir,1,Len(LineDir)-13) & "\" & DEF_RegisterFile
			If Ubound(LineCommandArray,1) = 3 Then
				NoteInfo = LineCommandArray(3)
			Else
				NoteInfo = LineCommand & " " & LineDir
			End If
			Select Case LineCommand
				Case "del":
					If Right(LineDir,1) = "\" Then
						DelFolder(DEF_BBS_HomeUrl & LineDir)
					Else
						CALL DelFile(DEF_BBS_HomeUrl & LineDir,0)
					End If
					CALL Update_ECHO(NoteInfo,0)
				Case "copy":
					GData = ""
					GDataFlag = 0
					If Right(LineDir,1) = "\" Then
						Update_CreateFolder(DEF_BBS_HomeUrl & LineDir)
					Else
						If NetFlag = 1 Then
							If inStrRev(LineDir,".") > 0 Then
								Extend = Replace(Mid(LineDir,inStrRev(LineDir,".")),".","")
							Else
								Extend = ""
							End If
							If inStr("*js*css*asp*htm*html*xml*htc*asa*","*" & LCase(Extend) & "*") > 0 Then
								GData = BytesToBstr(Update_GetInternetFile(NetUrl & LineDir_Bak))
								ADODB_SaveToFile GData,DEF_BBS_HomeUrl & LineDir
								GDataFlag = 1
							Else
								GData = Update_GetInternetFile(NetUrl & LineDir_Bak)
								ADODB_SaveToFileBinary GData,DEF_BBS_HomeUrl & LineDir
								GDataFlag = 1
							End If
						Else
							CALL CopyFiles(Server.MapPath(NativeDir & LineDir_Bak),Server.MapPath(DEF_BBS_HomeUrl & LineDir))
						End If
					End If
					CALL Update_ECHO(NoteInfo,0)
					If LCase(LineDir) = LCase(DEF_ManageDir & "\update.asp") Then
						Update_UpdateFileFlag = 1
					End If
					if GDataFlag = 1 then
						select Case LCase(LineDir)
							Case "inc\ubbcode_setup.asp":
								CALL Update_InsertSetupRID(1051,"inc/Ubbcode_Setup.asp",1,GData," and ClassNum=" & 1)
								CALL Update_ECHO("saved " & LineDir,0)
							case "inc\upload_setup.asp":
								CALL Update_InsertSetupRID(1051,"inc/Upload_Setup.ASP",3,GData," and ClassNum=" & 3)
								CALL Update_ECHO("saved " & LineDir,0)
							case "inc\user_setup.asp":
								CALL Update_InsertSetupRID(1051,"inc/User_Setup.ASP",2,GData," and ClassNum=" & 2)
								CALL Update_ECHO("saved " & LineDir,0)
						end Select
					end if
					
				case "exe":
					If NetFlag = 1 Then
						ADODB_SaveToFile BytesToBstr(Update_GetInternetFile(NetUrl & LineDir_Bak)),LineDir
					Else
						CALL CopyFiles(Server.MapPath(NativeDir & LineDir_Bak),Server.MapPath(LineDir))
					End If
					
					Update_GetInternetFile(Replace(thisUrl,"update.asp","") & LineDir)
					CALL DelFile(LineDir,0)
					CALL Update_ECHO(NoteInfo,0)
				case "sql":
					dim sqlfile
					sqlfile = BytesToBstr(Update_GetInternetFile(NetUrl & LineDir_Bak))
					sqlfile = split(sqlfile,"-@-@-split-@-@-")
					if Ubound(sqlfile)>=1 Then
						select case DEF_UsedDataBase
							case 0:
								if trim(sqlfile(0)) <> "" then call ldexecute(sqlfile(0),1)
							case 2:
								if Ubound(sqlfile)>=2 then
									if trim(sqlfile(2)) <> "" then call ldexecute(sqlfile(2),1)
								end if
							case else
								if trim(sqlfile(1)) <> "" then call ldexecute(sqlfile(1),1)
						end select
					end if
					CALL Update_ECHO(NoteInfo,0)
			End Select
		End If
	Next

End Sub

Function DelFile(FilePath_tmp,t)
        On Error Resume Next
        Dim fso,arrFile,i
        Dim FilePath
        FilePath = FilePath_tmp
        If t = 0 Then FilePath = Server.MapPath(FilePath_tmp)
        
        arrFile = Split(FilePath,"|")
        Set Fso = Server.CreateObject("Scripting.FileSystemObject")
        
        for i=0 to UBound(arrFile)
            FilePath = arrFile(i)
            If Fso.FileExists(FilePath) then
                Fso.DeleteFile FilePath
            End If
            If Fso.folderexists(FilePath) then
                Fso.deleteFolder FilePath
            End If
        Next
        Set fso = nothing
        
        If Err then 
            Err.clear()
            DelFile = false
        else
            DelFile = true
        End If
End Function

 Function DelFolder(FolderPath)
 
        On Error Resume Next
        Dim Fso,arrFolder,i
        
        arrFolder = Split(FolderPath,"|")
        Set Fso = Server.CreateObject("Scripting.FileSystemObject")
        
        For i=0 to UBound(arrFolder)
            FolderPath = arrFolder(i)
            If Fso.folderexists(Server.MapPath(FolderPath)) then
                Fso.deleteFolder Server.MapPath(FolderPath)
            End If
        Next
        
        If Err then
            Err.clear()
            DelFolder = false
        else
            DelFolder = true
        End If
    End Function


Function CopyFiles(TempSource,TempEnd)
        On Error Resume Next
        
        Dim CopyFSO,arrSource,arrEnd
        Dim i,srcName,tarName
        
        CopyFiles = false
        Set CopyFSO = Server.CreateObject("Scripting.FileSystemObject")
        
        If TempSource ="" or TempEnd = "" then
            CopyFiles = false
            Exit Function
        End If
        
        arrSource = Split(TempSource,"|")
        arrEnd    = Split(TempEnd,"|")
        If UBound(arrSource) <> UBound(arrEnd) then
            CopyFiles= false
            Exit Function
        End If
        
        for i=0 to UBound(arrSource)
            srcName = arrSource(i)
            tarName = arrEnd(i)
            If CopyFSO.FileExists(tarName) Then
            	CALL DelFile(tarName,1)
            End If
            IF CopyFSO.FileExists(srcName) and not CopyFSO.FileExists(tarName) then
               CopyFSO.CopyFile srcName,tarName
               CopyFiles = true
            End If
        Next
        Set CopyFSO = Nothing
        
        If Err then 
            'Err.clear()
            CopyFiles = false
        End If
End Function

Function BytesToBstr(body) 

	on error resume next
	If LenB(body) < 1 Then
		BytesToBstr = ""
		Exit Function
	End If
	dim objstream
	set objstream = Server.CreateObject("adodb.stream")
	with objstream
	.Type = 1
	.Mode = 3
	.Open
	.Write body 
	.Position = 0
	.Type = 2
	.Charset = "GB2312"
	
	'.Charset = "UTF-8"
	BytesToBstr = .ReadText
	.Close
	end with
	set objstream = nothing
	If Err and BytesToBstr = "" Then
		BytesToBstr = body
		Err.clear
	End If

End Function

Function Update_GetInternetFile(ur)

	Dim url
	Url = ur
	url = Left(url,5000)
	If url = "" Then Exit Function
	Dim xmlHttp
	Set xmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
	xmlHttp.setTimeouts 5000,5000,5000,15000
	xmlHttp.setOption 2, 13056
	xmlHttp.open "GET", url, False, "", "" 
	
	on error resume next
	xmlHttp.send()
	If Err Then
		Response.Write "<p>错误描述: <font color=red>" & err.description & "</font></p>"
		Err.clear
		Update_GetInternetFile = "err"
		Exit Function
	End If

	If xmlHttp.readystate = 4 then 
	'if xmlHttp.status=200 Then
		Update_GetInternetFile = xmlhttp.Responsebody
	'end if 
	Else 
		Update_GetInternetFile = "err"
	End If
	Set xmlHttp = Nothing

End Function

Sub Update_InsertSetupRID(RID,ValueStr,ClassNum,saveData,extend)

	If Update_CheckSetupRIDExist(RID,extend) = 0 Then
		CALL LDExeCute("insert into LeadBBS_Setup(RID,ValueStr,ClassNum,saveData) values(" & Rid & ",'" & Replace(ValueStr,"'","''") & "'," & ClassNum & ",'" & Replace(saveData,"'","''") & "')",1)
	Else
		CALL LDExeCute("Update LeadBBS_Setup Set ValueStr='" & Replace(ValueStr,"'","''") & "',ClassNum=" & ClassNum & ",saveData='" & Replace(saveData,"'","''") & "' where RID=" & RID & extend,1)
	End If

End Sub

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
%>