<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Board_popfun.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<%
DEF_BBS_HomeUrl = "../../"

Const LMT_RankNumber = 1000  '���������

Dim GBL_ID
initDatabase
GBL_CHK_TempStr = ""
GBL_ID = checkSupervisorPass

Manage_sitehead DEF_SiteNameString & " - ����Ա",""
frame_TopInfo
DisplayUserNavigate("�޸���̳ͳ����������û�")
If GBL_CHK_Flag=1 Then
	If Request.Form("submitflag") = "yes" then
		If Request.Form("a") = "m" Then
			ReMakeRank
		Else
			ClearOnlineUser
		End If
	Else
		%><div class=frametitle>һ��������߻�Ա</div>
			<div class=frameline>ע�⣺�˹��ܽ�������¹��ܣ�</div>
			<ol class=listli>
				<li>�����ǰ���ߵ�������Ա�������οͣ���</li>
				<li>������������Ա����Ҫ��2�������ҵĻ��������³�Ϊ���߻�Ա</li>
				<li>���ÿ������(�������ذ���)����������Ϊ��</li>
				<li>�������������Ϊ��</li>
			</ol>
			<div class=alert>ȷ����Ϣ�� ���Ҫ��ʼ���������Աô��</div>
			
			<div class=frameline>
			<form action=ClearOnlineUser.asp method=post>
			<input type=hidden name=submitflag value="yes">
			<input type=submit value=�����ʼ���������Ա class=fmbtn>
			</form>
			</div>
			
			<div class=frameline><a href=../SiteManage/RepairSite.asp>������������޸�ÿ�������������������������������������</a>
			</div>
		
			<div class=frametitle>�������������û�����</div>
			
			<div class=frameline>
			�û�����������(����ʱ��)����������ֻ��ǰ<%=LMT_RankNumber%>���û����������ʸ�<br>
			����������Ҫ���ĵ�ʱ��ǳ�֮�󣬽��龡���ٽ��д������
			</div>
			
			<div class=frameline>
			<form action=ClearOnlineUser.asp method=post>
			<input type=hidden name=submitflag value="yes">
			<input type=hidden name=a value="m">
			<input type=submit value=������������û����� class=fmbtn>
			</form>
			</div>
		<%
	End If
Else
DisplayLoginForm
End If
frame_BottomInfo
closeDataBase
Manage_Sitebottom("none")

Sub ClearOnlineUser

	If DEF_UsedDataBase = 1 Then
		CALL LDExeCute("delete from LeadBBS_onlineUser",1)
	Else
		CALL LDExeCute("TRUNCATE TABLE LeadBBS_onlineUser",1)
	End If
	Application.Lock
	Application(DEF_MasterCookies & "ActiveUsers") = 0
	Application.UnLock
	
	Dim SQL,I
	If isArray(Application(DEF_MasterCookies & "BListAll")) = True Then
		SQL = Ubound(Application(DEF_MasterCookies & "BListAll"),2)
		Application.Lock
		For I = 0 To SQL
			Application(DEF_MasterCookies & "BDOL" & Application(DEF_MasterCookies & "BListAll")(0,I)) = 0
		Next
		Application.UnLock
	End If
	Response.write "<div class=frameline><span class=greenfont>�ɹ�������������û���</span>[" & DEF_Now & "]</div>"

End Sub

Sub ReMakeRank

	Server.ScriptTimeOut = 6000
	Dim Rs,N,OnlineTime
	Con.CommandTimeout = 600
	Set Rs = LDExeCute(sql_select("Select ID,OnlineTime,SessionID From LeadBBS_User Order by OnlineTime DESC",LMT_RankNumber),0)
	If Not Rs.Eof Then
		For N = 1 to 1000
			If Not Rs.Eof Then
				If cCur(Rs(2)) <> N Then CALL LDExeCute("Update LeadBBS_User Set SessionID=" & N & " where ID=" & Rs(0),1)
				OnlineTime = Rs(1)
				Rs.MoveNext
			Else
				Exit For
			End If
		Next
		CALL LDExeCute("Update LeadBBS_User Set SessionID=0 where OnlineTime<" & OnlineTime & " and SessionID<>0",1)
	End If
	Rs.Close
	Set Rs = Nothing
	Response.write "<div class=frameline><span class=greenfont>�ɹ����������û�������</span>[" & DEF_Now & "]</div>"

End Sub
%>