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
		���ݿ����Ӵ�����ȷ�����ݿ����Ӵ��Ƿ���ȷ��<br><br><font color=red><%=err.description%></font>
		<br><br><a href=Update.asp><b>&lt;&lt;������������</b></a>
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
		Response.Write "<p>����SQL���ִ�г���</p><p><font color=gray>" & server.htmlencode(SQL) & "</font></P>"
		Response.Write "<p>��������: <font color=red>" & err.description & "</font></p>"
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
	application(DEF_MasterCookies & "SiteDisbleWhyszoieiu") = "��̳�����У����Ժ�����."
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
	<title>LeadBBS 6.0/6.1 ��������</title>
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
		<p><b style="font-size:14px;">��������LeadBBS 6.0/6.1�����ϰ汾������(��������չ����)����</b></p>
		<div class="frameline">1. ������չ����������̳���ļ����ֶ��滻���˹��ܽ��ָ���̳��ԭ�����洢�����á�</div>
		<div class="frameline">2. ������չ������������̳�Ĳ�������ѡ������ڴ��ҵ���</div>
		<div class="frameline">3. ����Ƿ��а汾���£���������̳��ٷ����ӱȽϣ�����Ƿ����µĸ��¡�</div>
		<div class="frameline">4. �������²�������������̳��ٷ����ӱȽϣ������������°汾��</div>
		<div style="width:90%;margin-bottom:216px;BORDER: #EEE0CB 5px solid; BACKGROUND: #F9F5F0; text-align:left;width:500px;padding:22px;line-height:2.0">

		
	<%
	If Request("sure") = "1" Then
		dim startflag:startflag=request("startflag")
		if startflag <> "1" then%>
		<p><form action=update.asp method=post>
			<br><b><span color=ff0000 class=redfont>ע�⣺�������������²���<br></span></b>
			<br>
			<ul><li>����ǰ��վ���������ݿ�Աȣ��������ݿ��е����ò������滻��ǰ����
			<%If LCase(Request("checkversion")) = "updateversion" Then
			%><li><font color=blue><b>��ѡ�����Զ����£����и��½���ǿ���滻��Ӧ�ĸ����ļ�</b></font></li>
			<%end If%>
			<li>Ϊ��֤���°�ȫ��������ǿ����ͣ��̳���У�ֱ���������</li>
			</li>
			</ul>
			<input type=hidden name=sure value="<%=server.htmlencode(request("sure"))%>">
			<input type=hidden name=SubmitFlag value="<%=server.htmlencode(request("SubmitFlag"))%>">
			<input type=hidden name=startflag value="1">
			<input type=hidden name=checkversion value="<%=server.htmlencode(request("checkversion"))%>">
			
			<input type=submit value=ȷ������ class=fmbtn>
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
		ע�⣺�������������LeadBBS 6.0/6.1�����߰汾���������¡�
		</p>
		<br />
		<font color=blue>������İ汾���ɣ��������ٷ�����6.0�汾���¡�</font>
		<p>
		<!--
		<input class="fmchkbox" type="checkbox" name="leadbbs40" value="1" />�ҵ����ݿ⻹��4.0�汾<br />
		-->
		</p>
		<br /><font color=red>���棺�������򽫻�ǿ���滻�����ļ���<b>��������ñ���</b>��</font><br />
		<br />
		ע�⣺�˳���������ݿ⼰�����ļ��ĸ��£�����Ҫ��������ֶ����¡�
		<p>���ݿ����Ӵ���<input name=db value="<%If DEF_UsedDataBase <> 1 Then
				Response.Write server.htmlencode(DEF_AccessDatabase)
			Else
				Response.Write server.htmlencode(Server.MapPath(DEF_BBS_HomeUrl & DEF_AccessDatabase))
			End If%>" size=40 class='fminpt input_3'>
		</p>
		<input name=sure value=1 type=hidden>
		<p><input type=submit value="��ʼ����" class="fmbtn btn_4"></p>
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
Dim GBL_LeadBBS_Setup_Data '��ʱ��ȡ��SetupRID��¼��������
Dim SetupRID_1050
'0 ��ʱ����Ŀ¼����
'1 a2.asp/Const LMTDEF_MaxReAnnounceֵ =
'
Dim GBL_Update_LineStr '��ȡ����ʱ�ļ��ַ����С�
Dim GBL_UpdateVersion '�ڲ��汾��
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

	CALL Update_ECHO("<div class=alertdone>��ʼ����⡣����</div>",1)
	'���FSO
	If Update_CheckFSO = 0 Then
		CALL Update_ECHO("Ȩ�޲��㣺�ռ䲻֧��FSO����.",1)
		Exit Sub
	Else
		CALL Update_ECHO("Ȩ�޼�⣺���.",0)
	End If
	
	'������ݿ�
	'Update62_UpdateDatabase

	'��ʼ�������ļ�Ŀ¼λ��
	If Update_CheckSetupRIDExist(1050," and ClassNum=0") = 0 Then
		RID = 1050
		Randomize
		ValueStr = Right(MD5(Rnd*10000000*hour(now)),8)
		SetupRID_1050(0,0) = ValueStr
		ClassNum = 0
		saveData = ""
		CALL Update_InsertSetupRID(RID,ValueStr,ClassNum,saveData," and ClassNum=0")
		CALL Update_ECHO("�����ļ�Ŀ¼δ����,��ʼ�����,��ʱ�����ļ����Ŀ¼Ϊ<u>" & ValueStr & "</u>",0)
	Else
		SetupRID_1050(0,0) = GBL_LeadBBS_Setup_Data(2,0)
		CALL Update_ECHO("��ȡ�����ļ�Ŀ¼��<u>" & SetupRID_1050(0,0) & "</u>",0)
	End If
	SetupRID_1050(1,0) = ""
	SetupRID_1050(2,0) = "��ʱ�����ļ����Ŀ¼"
	Update_CreateFolder(DEF_BBS_HomeUrl & SetupRID_1050(0,0) & "/")
	
	If Update_CheckSetupRIDExist(1002," and ClassNum=0") = 0 Then
		RID = 1002
		ValueStr = "20100101001"
		GBL_UpdateVersion = ValueStr
		ClassNum = 0
		saveData = "�ڲ��汾��"
		CALL Update_InsertSetupRID(RID,ValueStr,ClassNum,saveData," and ClassNum=0")
		CALL Update_ECHO("��ʼ���ڲ��汾��Ϊ<u>" & ValueStr & "</u>",0)
	Else
		GBL_UpdateVersion = cCur(GBL_LeadBBS_Setup_Data(2,0))
		CALL Update_ECHO("��ȡ�ڲ��汾�ţ�<u>" & GBL_UpdateVersion & "</u>",0)
	End If
	
	'��ȡ�ļ�����
	If SubmitFlag = "" Then CALL Update_ECHO("<div class=alertdone>��ȡ��̳������Ϣ������</div>",1)

	CurN = 1
	CALL Update_GetFileParaValue("�������$$$$:$a/a2.asp","Const LMT_EnableOtherGuestName",CurN,"������̳�Ƿ�����ʹ��""�ο�""���������")
	CALL Update_GetFileParaValue("a/a2.asp","Const LMT_BuyAnnounceMaxPoints",CurN,"���������ĵ�������")
	CALL Update_GetFileParaValue("a/a2.asp","Const LMTDEF_MaxReAnnounce",CurN,"��������ظ������������������������ƲŻ���Ч")
	CALL Update_GetFileParaValue("a/a2.asp","Const LMTDEF_MinAnnounceLength",CurN,"������Ҫ��������")
	CALL Update_GetFileParaValue("a/a2.asp","Const LMTDEF_NotReplyDate",CurN,"���ظ�ʱ��������ڶ�������������ֹ�ظ�,�԰�����������Ч")
	CALL Update_GetFileParaValue("a/a2.asp","Const LMTDEF_NeedCachetValue",CurN,"�趨���������û������Լ�����ר��")
	CALL Update_GetFileParaValue("a/a2.asp","Const LMTDEF_ColorSpend",CurN,"�趨������ɫ���Ķ��ٲƸ�ֵ")
	CALL Update_GetFileParaValue("a/a2.asp","Const LMTDEF_RepostMsg",CurN,"�ظ������Ƿ�Ĭ�϶���Ϣ֪ͨ����,0��Ĭ�ϲ�֪ͨ 1.�ظ�ȫ��֪ͨ 2.��������ʱ��֪ͨ")
	
	CALL Update_GetFileParaValue("a/a.asp","Const LMT_RefreshEnable",CurN,"�û��ظ���������Ƿ���������")
	CALL Update_GetFileParaValue("a/a.asp","Const LMTDEF_RepostMsg",CurN,"�ظ������Ƿ�Ĭ�϶���Ϣ֪ͨ����,0��Ĭ�ϲ�֪ͨ 1.�ظ�ȫ��֪ͨ(ע�������(a2.asp�ļ�)���ñ���һֱ)")
	
	CALL Update_GetFileParaValue("a/Editannounce.asp","Const LMTDEF_MinAnnounceLength",CurN,"�༭�ύ������������Ҫ��������")
	CALL Update_GetFileParaValue("a/Editannounce.asp","Const LMT_BuyAnnounceMaxPoints",CurN,"���������ĵ�������")
	CALL Update_GetFileParaValue("a/Editannounce.asp","Const LMTDEF_NeedCachetValue",CurN,"�趨���������û������Լ�����ר��")
	
	CALL Update_GetFileParaValue("�������ط�ʽ$$$$:$a/file.asp","Const LMT_RedirectFile",CurN,"������ʾ��ʽ��0,��ȡ���أ�������ʵ��ַ�������Բ� 1.תַ���� �����ܵ���¶��ʵ��ַ")
	
	CALL Update_GetFileParaValue("�ղ�������$$$$:$a/Processor.asp","Const LMT_MaxCollectAnnounce",CurN,"��������ղ���������")
	CALL Update_GetFileParaValue("����Ϣ֪ͨ$$$$:$a/Processor.asp","Const LMT_Prc_anonymity",CurN,"�������Ƿ���������Ϣ֪ͨ�û��� 0 ����Ϊϵͳ 1 ԭ������")
	CALL Update_GetFileParaValue("a/Processor.asp","Const LMT_Prc_MsgFlag",CurN,"����Ա�Ƿ����Ϣ֪ͨ�û��� 0 Ĭ��ѡ��Ϊ��֪ͨ,����ѡ���Ƿ�֪ͨ 1 Ĭ�϶���Ϣ֪ͨ,Ҳ��ѡ���Ƿ�֪ͨ 2.ǿ�ƶ���Ϣ֪ͨ,�����Ƿ�֪ͨ")
	
	CALL Update_GetFileParaValue("��������$$$$:$a/inc/AddFriend.asp","Const LMT_MaxFriendNum",CurN,"������ӵ���������Ŀ")

	CALL Update_GetFileParaValue("�������$$$$:$a/inc/DelAnnounce.asp","AncIDStr = ",CurN,"�������ID�б����ŷָ����ظ��������ӽ������������(1-3)�����Ҵ������ӽ���ֹɾ���ظ�(���ɱ༭)")
	CALL Update_GetFileParaValue("a/a2.asp","AncIDStr = ",CurN,"�����������ID�б����ŷָ����ظ��������ӽ������������(1-3)��ע����[DelAnnounce.asp]���ñ���һ��")
	
	CALL Update_GetFileParaValue("Ĭ�Ϸ���ģʽ$$$$:$a/inc/Editor_Fun.asp","Const Edt_MiniMode",CurN,"�������棺0-��ͳ��Լģʽ 1.�๦��ģʽ")
	
	CALL Update_GetFileParaValue("��ý�岥�Ÿ������Ƿ��Զ�����$$$$:$a/inc/leadcode.js","var vnum = ",CurN,"1  forbid play,-2 allow 3 video to play at same time. 0: allow one")
	CALL Update_GetFileParaValue("a/inc/leadcode.js","var autoplay = ",CurN,"0.manual play 1.auto play")
	
	CALL Update_GetFileParaValue("������������$$$$:$a/inc/MakeGoodAnnounce.asp","Const DEF_AllowPunish",CurN,"�Ƿ�������ͨ�û��ͷ������û���1.�����ձ��û��ͷ������û���������ֹ")
	CALL Update_GetFileParaValue("a/inc/MakeGoodAnnounce.asp","Const DEF_AllowOpinionNum",CurN,"������ͨ�û����۴��� 0,��ֹ,-1 �������� >0 ָ������")
	CALL Update_GetFileParaValue("a/inc/MakeGoodAnnounce.asp","Const DEF_MasterNolimit",CurN,"����������Ա���۴����Ƿ����ޣ����������ޣ���������ͬ��ͨ�û���")
	CALL Update_GetFileParaValue("a/inc/MakeGoodAnnounce.asp","Const DEF_AllowBoardMasterCachetValue",CurN,"�Ƿ������������������1.�� 0.��")
	
	CALL Update_GetFileParaValue("ͶƱ����$$$$:$a/inc/Poll_fun.asp","Const LMT_PollNeedPoints",CurN,"�û�ͶƱ������Ҫ�ﵽ�Ļ��֣�����Ϊ����")
	
	CALL Update_GetFileParaValue("RSS$$$$:$other/RSS.asp","Const RSS_ViewNumer",CurN,"���������ʾ��RSS��¼����")
	
	CALL Update_GetFileParaValue("����$$$$:$Search/Search.asp","Const Sch_AllContent",CurN,"�Ƿ�����ȫ������,��ͬʱ������������ݣ���Ϊ99��ʾ����hubbledotnet����ajax������������Ϊ98���������ʽ����hubbledotnet����")
	CALL Update_GetFileParaValue("Search/Search.asp","Const Sch_AncTitle",CurN,"�Ƿ��������ӱ�������")
	CALL Update_GetFileParaValue("Search/Search.asp","Const Sch_AncContent",CurN,"�Ƿ�����������������")
	CALL Update_GetFileParaValue("Search/Search.asp","Const Sch_LimitTime",CurN,"��������ʱ��(��λ��)")
	
	CALL Update_GetFileParaValue("Search/inc/Search_fun.asp","Const DEF_BBS_MaxListPage",CurN,"������������ʾҳ��(�������Ӱ�����ܣ�Ĭ������Ϊ10)")
	CALL Update_GetFileParaValue("Search/inc/Search_fun.asp","Const DEF_BBS_MaxWords",CurN,"�������������������Ҫ��ʾ����(�����ʾ�ֽ�)")
	
	CALL Update_GetFileParaValue("����Ϣ$$$$:$User/LookMessage.asp","Const LMT_LookedMsgExpiresDay",CurN,"����Ϣ�Ķ���ı�������(��λ��)")
	CALL Update_GetFileParaValue("User/SendMessage.asp","Const DEF_User_MaxReceiveUser",CurN,"��������ͬʱ���Ͷ���Ϣ�����ٸ��û���Ĭ��ֵΪ5")
	
	CALL Update_GetFileParaValue("֧����$$$$:$User/alipay/alipay_Config.asp","partner = ",CurN,"֧������ȡid��������Ҫһ��֧�����˺ţ��ٴ���Ӧ��ַ��ȡid(<a href=https://www.alipay.com/himalayas/practicality_customer.htm?customer_external_id=C4335329546596834111&market_type=from_agent_contract&pro_codes=F7F62F29651356BB target=_blank>��˻�ȡ</a>)")
	CALL Update_GetFileParaValue("User/alipay/alipay_Config.asp","key = ",CurN,"֧������ȡ����Կ��������Ҫһ��֧�����˺ţ��ٴ���Ӧ��ַ��ȡ��Կ(<a href=https://www.alipay.com/himalayas/practicality_customer.htm?customer_external_id=C4335329546596834111&market_type=from_agent_contract&pro_codes=F7F62F29651356BB target=_blank>��˻�ȡ</a>)")
	
	CALL Update_GetFileParaValue("����Ϣ$$$$:$User/inc/Fun_SendMessage.asp","Const LMT_SendMsgExpiresDate",CurN,"�����·��Ͷ���Ϣ��������(�����Զ�ɾ��)")
	CALL Update_GetFileParaValue("User/inc/UserTopic.asp","Const LMT_MaxMessageNumber",CurN,"�û��ռ�������������ռ�¼���������޷���������Ϣ��")
	
	CALL Update_GetFileParaValue("��ý���ظ���������$$$$:$a/inc/leadcode.js","var playcount = ",CurN,"play loop count:0-100,0=always replay")
	
	CALL Update_GetFileParaValue("ע����֤����$$$$:$User/" & DEF_RegisterFile,"Const LMT_RegVerifyQuestion = ",CurN,"ע����֤��ʾ��Ϣ��������HTML��ʽ������ʹ��ͼƬ��������д��ʾ������ע����֤��Ϣ��")
	CALL Update_GetFileParaValue("User/" & DEF_RegisterFile,"Const LMT_RegVerifyAnswer = ",CurN,"ע����֤��Ҫ��д�Ĵ𰸡�")
	
	CALL Update_GetFileParaValue("QQ��������$$$$:$app/qqlogin/oauth.asp","Const apiKey = ",CurN,"APP ID,����Ҫ����Ѷƽ̨�����ȡ���ϣ�(<a href=http://connect.qq.com/ target=_blank>�������</a>)")
	CALL Update_GetFileParaValue("app/qqlogin/oauth.asp","Const secretKey = ",CurN,"APP KEY,����Ҫ����Ѷƽ̨�����ȡ")
	CALL Update_GetFileParaValue("app/qqlogin/oauth.asp","Const callback = ",CurN,"CALL BACK,�ص���ַ��ע��ֻ��Ҫ��д������������http��Ŀ¼��")
	
	CALL Update_GetFileParaValue("���ڷ����������$$$$:$a/a.asp$$$$:$textarea","Const LMTDEF_ShareID = ",CurN,"������д��վ�����б�д���͵ķ������(HTML��ʽ��ע���ֹ�ɾ�����з�),����Ϊ����رշ������;")
	GBL_ParaCount = CurN - 1

	'��֤���������
	If SubmitFlag = "" Then CALL Update_ECHO("<div class=alertdone>��Ⲣ�������á�����</div>",1)
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
				If SubmitFlag = "" Then CALL Update_ECHO("������" & N & "(<span class=grayfont>" & SetupRID_1050(2,N) & "</span>)�洢��ɣ�ֵΪ��<u>" & ValueStr & "</u>",0)
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
					CALL Update_ECHO("��ǰ������<u>" & N & "</u>(<span class=grayfont>" & SetupRID_1050(2,N) & "</span>)��ʵ�ʲ������ѴӴ洢���ݶ�ȡ�������������á�",1)
					
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
					If SubmitFlag = "" Then CALL Update_ECHO("��ǰ������<u>" & N & "</u>ȷ������(<span class=grayfont>" & SetupRID_1050(2,N) & "</span>)��",0)
				End If
			End If
	Next
	
	'��Ⲣ����BBSSetup.asp, Ubbcode_Setup.asp,User_Setup.ASP,Upload_Setup.asp,AD_Data.asp 
	'��Ⲣ����User/inc/Contact_info.asp User_Reg.asp
	CALL Update_ECHO("<div class=alertdone>��Ⲣ���������ļ�������</div>",1)
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
			If SubmitFlag = "" Then CALL Update_ECHO("�����ļ�(<span class=grayfont>" & FileSetupData(N)& "</span>)�ѱ��档",0)
		Else
			FileContent = ADODB_LoadFile(DEF_BBS_HomeUrl & FileSetupData(N))
			If FileContent <> GBL_LeadBBS_Setup_Data(4,0) Then
				ADODB_SaveToFile GBL_LeadBBS_Setup_Data(4,0),DEF_BBS_HomeUrl & FileSetupData(N)
				If SubmitFlag = "" Then CALL Update_ECHO("�����ļ�<u>" & FileSetupData(N) & "</u>��洢���ݲ�������ǰ��������ɸ��¡�",1)
			Else
				If SubmitFlag = "" Then CALL Update_ECHO("�����ļ�<u>" & FileSetupData(N) & "</u>��ɼ�⡣",0)
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

	'�ֶ����ò���
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
		CALL Update_ECHO("��ȡ����" & Index & "(<u>" & fileName & "/" & fileStr & "</u>)ʧ�ܣ���ӹٷ���������ԭ�ļ����²��滻<u>" & fileName & "</u>.",1)
	Else
		SetupRID_1050(1,Index) = f
		SetupRID_1050(2,Index) = Note
		SetupRID_1050(3,Index) = GBL_Update_LineStr
		SetupRID_1050(4,Index) = fStr
		If SubmitFlag = "" Then CALL Update_ECHO("��ȡ����" & Index & "��ɡ�",0)
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
		CALL Update_ECHO("�ռ䲻֧��FSO,Ŀ¼��������.",1)
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
			GBL_CHK_TempStr = "<br>��ȡ�ļ�ʧ�ܣ�" & err.description & "<br>�������ܣ�ȷ���Ƿ�Դ��ļ��ж�ȡȨ��."
			err.Clear
			Set Fs = Nothing
			Exit Function
		End If
		If Not WriteFile.AtEndOfStream Then
			ADODB_LoadFile = WriteFile.ReadAll
			If Err Then
				GBL_CHK_TempStr = "��ȡ�ļ�ʧ�ܣ�<p>" & err.description & "</p> �������ܣ�ȷ���Ƿ�Դ��ļ��ж�ȡȨ��."
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
			GBL_CHK_TempStr = "����������֧��ADODB.Stream���޷���ɲ��������ֹ�����"
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
		GBL_CHK_TempStr = "������Ϣ��<p>" & err.description & "</p>�������ܣ�ȷ���Ƿ�Դ��ļ��ж�ȡȨ��."
		err.Clear
		Set Fs = Nothing
		Exit Function
	End If

End Function

'�洢���ݵ��ļ�
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
			GBL_CHK_TempStr = "����������֧��ADODB.Stream���޷���ɲ��������ֹ�����"
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
		GBL_CHK_TempStr = "������Ϣ��<p>" & err.description & "</p>�������ܣ�ȷ���Ƿ�Դ��ļ���д��Ȩ��."
		err.Clear
		Set Fs = Nothing
		Exit Sub
	End If

End Sub


'�洢���ݵ��ļ�
Sub ADODB_SaveToFileBinary(ByVal strBody,ByVal File)

	On Error Resume Next
	Dim objStream
	
		Set objStream = Server.CreateObject("ADODB.Stream")
		If Err.Number=-2147221005 Then 
			GBL_CHK_TempStr = "����������֧��ADODB.Stream���޷���ɲ��������ֹ�����"
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
		GBL_CHK_TempStr = "������Ϣ��<p>" & err.description & "</p>�������ܣ�ȷ���Ƿ�Դ��ļ���д��Ȩ��."
		err.Clear
		Set Fs = Nothing
		Exit Sub
	End If
	Response.Write "<span class=grayfont>�ļ�����:" & LenB(strBody) & " Bytes</span>"

End Sub

Sub Update_SetupFilePara

%>
<form name="pollform3sdx" method="post" action="Update.asp">
<input type="hidden" name="SubmitFlag" value=yes>
<input type="hidden" name="sure" value=1>
<br />
<p>
		<b>���ã�<span class=grayfont>��̳��չ��������</span></b>
		<br>
		<span class=grayfont>(����Ϊ��վ��������ע���޸ģ���������ý��ᷢ�����ش���)<br><br>
		��ο�ע���޸Ĳ�����<span class=redfont>һЩ����ֵΪ�ַ����ģ�ע�Ᵽ������˫����</font>��</span>
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
			<td class=tdbox width=90>������<%=N%></td>
			<td class=tdbox><input class=fminpt type="text" name="Form_SetupRID_<%=N%>" maxlength="2048" size="45" value="<%=server.htmlencode(SetupRID_1050(0,N))%>"><br /> <span class=note><%=SetupRID_1050(2,N)%></span></td>
		</tr>
		<tr>
	<%
	Next%>
	<td class=tdbox>&nbsp;</td>
	<td class=tdbox>
		<input type=submit name=�ύ value=�ύ class=fmbtn>
		<input type=reset name=ȡ�� value=ȡ�� class=fmbtn>
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
			GBL_CHK_TempStr = "������" & N & " ������д"
			Exit Sub
		End If
		If inStr(Val,"<" & "%") > 0 or inStr(Val,"%" & ">") > 0 Then
			GBL_CHK_TempStr = "������" & N & " ��д���󣬲��ܰ���һЩ�����ִ���"
			Exit Sub
		End If
		
		If inStr(SetupRID_1050(0,N),"""") > 0 Then
			If Len(Val) > 1024 Then
				GBL_CHK_TempStr = "������" & N & " ������"
				Exit Sub
			End If
			If Left(Val,1) <> """" or Right(Val,1) <> """" Then
				GBL_CHK_TempStr = "������" & N & " ���󣬴�ֵΪ�ַ���������ǰ��ʹ�õ����š�"
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
				GBL_CHK_TempStr = "������" & N & " ���󣬴�ֵ����Ϊ��ȷ�����֡�"
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
	CALL Update_ECHO("<br /><b><font color=green>�ɹ�����ֶ����á�</b></font>",1)

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
			CALL Update_ECHO("<div class=alert>���������������޷����Ӹ��·�������������ֹ��</div>",0)
			Exit Sub
		End If
	End If
	Update = Split(Update,VbCrLf)
	
	CALL Update_ECHO("<div class=alertdone>��ʼ����Ƿ��в������¡�����</div>",0)
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
				CALL Update_ECHO("<div class=alertdone>��⵽�²���<u>" & CurFile_Name & "</u><span class=redfont>" & CurFile_Intro & "</span></div>",0)
				UpdateFlag = 1
			End If
		End If
	Next
	If UpdateFlag = 0 Then
		CALL Update_ECHO("<div class=alertdone>��������������̳�������°汾��������¡�</div>",0)
	Else
		CALL Update_ECHO("<div class=alertdone>�����ɣ����������Ĳ������¿�ʼ���¡�</div>",0)
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
			CALL Update_ECHO("<div class=alert>���������������޷����Ӹ��·�������������ֹ��</div>",0)
			Exit Sub
		End If
	End If
	Update = Split(Update,VbCrLf)
	
	CALL Update_ECHO("<div class=alertdone>��ʼ����²�����׼���������¡�����</div>",0)
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
				CALL Update_ECHO("<div class=alertdone>���²���<u>" & CurFile_Name & "</u><span class=grayfont>" & CurFile_Intro & "</span>��</div>",0)
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
				saveData = "�ڲ��汾��"
				CALL Update_InsertSetupRID(RID,ValueStr,ClassNum,saveData," and ClassNum=0")
				CALL Update_ECHO("��ʼ���ڲ��汾��Ϊ<u>" & ValueStr & "</u>",0)
				If UpdateFlag = 0 Then Update_CloseSite
				UpdateFlag = 1
				
				If Update_UpdateFileFlag = 1 Then
					Update_OpenSite
					CloseDatabase
					CALL Update_ECHO("<div class=alert>����ģ���ø��£�����ǿ����ֹ���ɵ���Ҳ���¹��ܼ����汾���¡�</div>",0)
					Update_PageBottom
					Response.End
				End If
			End If
		End If
	Next
	If UpdateFlag = 0 Then
		CALL Update_ECHO("<div class=alertdone>�����������̳�������°汾��������¡�</div>",0)
	Else
		Update_OpenSite
		CALL Update_ECHO("<div class=alertdone>��������ɸ��£��������¿�ʼ�����̳���á�</div>",0)
		Update62_initBBSdata
	End If

End Sub

Sub Update_CloseSite

	Application.Lock
	application(DEF_MasterCookies & "SiteEnableFlagzoieiu") = 0
	application(DEF_MasterCookies & "SiteDisbleWhyszoieiu") = "<html><body>��̳��������״̬�����Ժ���ʡ�����ʱ���޷����ʣ�����ϵ����Ա��</body></html>"
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
		Response.Write "<p>��������: <font color=red>" & err.description & "</font></p>"
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