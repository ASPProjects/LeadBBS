<!-- #include file=../../inc/BBSSetup.asp -->
<!-- #include file=../../inc/Board_Popfun.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<%
DEF_BBS_HomeUrl = "../../"
Dim GBL_ID
initDatabase
GBL_CHK_TempStr = ""
GBL_ID = checkSupervisorPass

Manage_sitehead DEF_SiteNameString & " - ����Ա",""
frame_TopInfo
DisplayUserNavigate("ȫ���������ܹ���")
If GBL_CHK_Flag=1 Then
	LoginAccuessFul
Else
DisplayLoginForm
End If	
frame_BottomInfo
closeDataBase
Manage_Sitebottom("none")

Sub LoginAccuessFul

	If DEF_UsedDataBase <> 0 Then
		GBL_CHK_TempStr = "<div class=alert>Access���ݿⲻ֧��ȫ����������!</div>"
		Exit Sub
	End If%>

<div class=frametitle>���ݿ�ȫ���������ÿ�������</div>
<div class=frameline>
		<a href=#29 onclick="javascript:window.open('ExeCuteFullTEXTCommands.asp?ExeFlag=1','','width=300,height=20 scrollbars=yes,status=no');"><span class=greenfont>Ϊ���ݿ�����ȫ������</span></a>
		<a href=#29 onclick="javascript:window.open('ExeCuteFullTEXTCommands.asp?ExeFlag=2','','width=300,height=20 scrollbars=yes,status=no');"><span class=redfont>Ϊ���ݿ����ȫ������</span></a> (�Ѿ������мɲ�Ҫ������,������������������)
</div>
<div class=frameline>
		<a href=#29 onclick="javascript:window.open('ExeCuteFullTEXTCommands.asp?ExeFlag=3','','width=300,height=20 scrollbars=yes,status=no');"><span class=greenfont>����ȫ�������������</span></a>
		<a href=#29 onclick="javascript:window.open('ExeCuteFullTEXTCommands.asp?ExeFlag=4','','width=300,height=20 scrollbars=yes,status=no');"><span class=redfont>ֹͣȫ�������������</span></a> (��̳�����ӵ�ʲô����Ҳ�Ѳ���������)
</div>
<div class=frameline>
		<a href=#29 onclick="javascript:window.open('ExeCuteFullTEXTCommands.asp?ExeFlag=5','','width=300,height=20 scrollbars=yes,status=no');"><span class=greenfont>�������º�̨�е�����</span></a>
		<a href=#29 onclick="javascript:window.open('ExeCuteFullTEXTCommands.asp?ExeFlag=6','','width=300,height=20 scrollbars=yes,status=no');"><span class=redfont>ֹͣ���º�̨�е�����</span></a> (��̳�����˷��˰��쵫��������������
</div>
<div class=frameline>
		<a href=#29 onclick="javascript:window.open('ExeCuteFullTEXTCommands.asp?ExeFlag=7','','width=300,height=20 scrollbars=yes,status=no');"><span class=bluefont>����ǰһϵ�и��ٵı仯������ȫ������(��������)</span></a>
</div>
		
<div class=frametitle>������������</div>
<div class=frameline>
		<a href=#29 onclick="javascript:window.open('ExeCuteFullTEXTCommands.asp?ExeFlag=8','','width=300,height=20 scrollbars=yes,status=no');"><span class=bluefont>���MSSQL��ǰʹ�����ݿ���־(ɾ���󲻿ɻָ���־������־��ʱ��ʹ�ô������ע�⾭�����)</span></a><br>
		<a href=#29 onclick="javascript:window.open('ExeCuteFullTEXTCommands.asp?ExeFlag=9','','width=300,height=20 scrollbars=yes,status=no');"><span class=bluefont>����MSSQL��ǰʹ�����ݿ���־�ļ�(��СLog�ļ����ͷ�Ӳ�̿ռ��ϵͳ)</span></a><br>
		<a href=#29 onclick="javascript:window.open('ExeCuteFullTEXTCommands.asp?ExeFlag=10','','width=300,height=20 scrollbars=yes,status=no');"><span class=bluefont>����MSSQL��ǰʹ�����ݿ�������ļ�(<span class=redfont>��С��ʹ�ô��ʹ��ȫ���������ݿ���ܻ����һЩ���ȶ�</span>)</span></a>
</div>
<%
	DisplayOtherInfo

End Sub

Sub DisplayOtherInfo

	Response.Write "<div class=frametitle>���ݿ�����ο�</div>"
	Response.Write "<table border=0 cellpadding=0 cellspacing=0 width=100% class=frame_table>"
	Dim Rs,SQL
	SQL = "Select @@TRANCOUNT,@@VERSION,@@SERVERNAME,@@LANGUAGE,@@CONNECTIONS,@@CPU_BUSY,@@IDLE,@@IO_BUSY,@@LOCK_TIMEOUT,@@MAX_CONNECTIONS,@@TOTAL_READ,@@TOTAL_WRITE,CURRENT_USER,APP_NAME(),HOST_NAME(),DB_NAME(DB_ID()),DATABASEPROPERTY(DB_NAME(DB_ID()), 'IsFulltextEnabled')"
	Set Rs = LDExeCute(SQL,0)
	If Not Rs.Eof Then
		Response.Write "<tr><td class=tdbox width=200>��ǰ���ӵĻ������</td><td class=tdbox>" & Rs(0) & "��</td>"
		Response.Write "<tr><td class=tdbox>��ǰ��װ�����ڡ��汾�ʹ���������</td><td class=tdbox>" & Rs(1) & "</td>"
		Response.Write "<tr><td class=tdbox>���ط���������</td><td class=tdbox>" & Rs(2) & "</td>"
		Response.Write "<tr><td class=tdbox>��ǰʹ�õ�������</td><td class=tdbox>" & Rs(3) & "</td>"
		Response.Write "<tr><td class=tdbox>���ϴ������������ӻ���ͼ���Ӵ���</td><td class=tdbox>" & Rs(4) & "��</td>"
		Response.Write "<tr><td class=tdbox>���ϴ���������CPU�Ĺ���ʱ��</td><td class=tdbox>" & Rs(5) & "���루����ϵͳ��ʱ���ķֱ��ʣ�</td>"
		Response.Write "<tr><td class=tdbox>���ϴ����������õ�ʱ��</td><td class=tdbox>" & Rs(6) & "���루����ϵͳ��ʱ���ķֱ��ʣ�</td>"
		Response.Write "<tr><td class=tdbox>���ϴ�����������ִ���������ʱ��</td><td class=tdbox>" & Rs(7) & "���루����ϵͳ��ʱ���ķֱ��ʣ�</td>"
		
		Response.Write "<tr><td class=tdbox>���ص�ǰ�Ự�ĵ�ǰ����ʱ����</td><td class=tdbox>" & Rs(8) & "����</td>"
		Response.Write "<tr><td class=tdbox>�����ͬʱ�û����ӵ������</td><td class=tdbox>" & Rs(9) & "��(32767��ʾδ����)</td>"
		Response.Write "<tr><td class=tdbox>���ϴ��������ȡ���̵Ĵ���</td><td class=tdbox>" & Rs(10) & "�Σ����Ƕ�ȡ���ٻ��棩</td>"
		Response.Write "<tr><td class=tdbox>���ϴ�������д����̵Ĵ���</td><td class=tdbox>" & Rs(11) & "��</td>"
		Response.Write "<tr><td class=tdbox>��ǰ��¼�û���</td><td class=tdbox>" & Rs(12) & "</td>"
		Response.Write "<tr><td class=tdbox>��ǰ�Ự��Ӧ�ó�������</td><td class=tdbox>" & Rs(13) & "</td>"
		Response.Write "<tr><td class=tdbox>����վ����</td><td class=tdbox>" & Rs(14) & "</td>"
		Response.Write "<tr><td class=tdbox>���ݿ�����</td><td class=tdbox>" & Rs(15) & "</td>"
		Response.Write "<tr><td class=tdbox>���ݿ��Ƿ�ȫ������</td><td class=tdbox>" & Replace(Replace(Rs(16) & "","0","��"),"1","��") & "</td>"

		Response.write "</tr>"
	End If
	Rs.Close
	Set Rs = Nothing
	Response.Write "</table>"
	
	Response.Write "<div class=frametitle>�鿴���ݿ����Ϣ</div><div class=frameline><a href=TableInfo.asp?tb=LeadBBS_Announce>�������鿴��LeadBBS_Announce��Ϣ</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_Assort>�������鿴��LeadBBS_Assort</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_Boards>�������鿴��LeadBBS_Boards</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_CollectAnc>�������鿴��LeadBBS_CollectAnc</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_ForbidIP>�������鿴��LeadBBS_ForbidIP</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_FriendUser>�������鿴��LeadBBS_FriendUser</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_GoodAssort>�������鿴��LeadBBS_GoodAssort</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_InfoBox>�������鿴��LeadBBS_InfoBox</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_IPAddress>�������鿴��LeadBBS_IPAddress</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_Link>�������鿴��LeadBBS_Link</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_onlineUser>�������鿴��LeadBBS_onlineUser</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_Setup>�������鿴��LeadBBS_Setup</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_SiteInfo>�������鿴��LeadBBS_SiteInfo</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_SpecialUser>�������鿴��LeadBBS_SpecialUser</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_TopAnnounce>�������鿴��LeadBBS_TopAnnounce</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_Upload>�������鿴��LeadBBS_Upload</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_User>�������鿴��LeadBBS_User</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_UserFace>�������鿴��LeadBBS_UserFace</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_VoteItem>�������鿴��LeadBBS_VoteItem</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_VoteUser>�������鿴��LeadBBS_VoteUser</a></div>"

End Sub%>