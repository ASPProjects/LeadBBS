<!-- #include file=../../inc/BBSSetup.asp -->
<!-- #include file=../../inc/Board_Popfun.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<%
Server.ScriptTimeOut = 600
DEF_BBS_HomeUrl = "../../"
Dim GBL_ID
initDatabase
GBL_CHK_TempStr = ""
GBL_ID = checkSupervisorPass

Dim GBL_EXEString

Manage_sitehead DEF_SiteNameString & " - ����Ա",""
frame_TopInfo
DisplayUserNavigate("����Access���ݿ�")
If GBL_CHK_Flag=1 Then
	LoginAccuessFul
Else
DisplayLoginForm
End If
frame_BottomInfo
closeDataBase
Manage_Sitebottom("none")


function Copyfiles(tempsource,tempend)

    'on error resume next
    Dim fs
    Set fs = Server.CreateObject(DEF_FSOString)

    If fs.FileExists(tempend) then
       Response.Write "Ŀ�걸���ļ�" & tempend & "�Ѵ��ڣ�����ɾ��!"
       Set fs=nothing
       Exit Function
    End If
    
    If fs.FileExists(tempsource) then
    Else
       Response.Write "Ҫ���Ƶ�Դ���ݿ��ļ�"&tempsource&"������!"
       Set fs=nothing
       Exit Function
    End If
    fs.CopyFile tempsource,tempend
    Response.Write "�Ѿ��ɹ������ļ�"&tempsource&"��"&tempend&"!"
    Set fs = Nothing

end function

Function DeleteFiles(path)

	If DEF_FSOString = "" Then Exit Function
	on error resume next
	Dim fs
	Set fs = Server.CreateObject(DEF_FSOString)
	If err <> 0 Then
		Err.Clear
		Response.Write "<br>��������֧��FSO��Ӳ�̱����ļ�δɾ����"
		Exit Function
	End If
	If fs.FileExists(path) Then
		fs.DeleteFile path,True
		DeleteFiles = 1
	Else
		DeleteFiles = 0
	End If
	Set fs = Nothing
     
End Function

Function LoginAccuessFul

	If DEF_FSOString = "" Then
		Response.Write "<p><b>��������֧�ִ˹���!</b></p>"
		Exit Function
	End If
	If DEF_UsedDataBase <> 1 Then
		Response.Write "<p><b>�˹��ܽ���Access���ݿ���Ч!</b></p>"
		Exit Function
	End If
	Dim action
	action = Request.Form("action")
	If action <> "backup" and action <> "delbackup" and action <> "CompactDatabase" Then action = ""
	If action = "" Then
		DisplayStringForm
	Else
		If action = "backup" Then
			Response.Write "<p><br><br>�������ݿ⿪ʼ����վ��ͣһ���û���ǰ̨����......<br>"
			Application.Lock
			application(DEF_MasterCookies & "SiteEnableFlagzoieiu") = 0
			application(DEF_MasterCookies & "SiteDisbleWhyszoieiu") = "��̳��ͣ�У����Ժ򼸷��Ӻ�����..."
			Application.UnLock
			CloseDatabase
			Copyfiles Server.Mappath(DEF_BBS_HomeUrl & DEF_AccessDatabase),Server.Mappath(DEF_BBS_HomeUrl & DEF_AccessDatabase & ".BAK")
			OpenDatabase
			Response.write "<p>�������..."
			Application.Lock
			application(DEF_MasterCookies & "SiteEnableFlagzoieiu") = 1
			application(DEF_MasterCookies & "SiteDisbleWhyszoieiu") = ""
			Application.UnLock
			Response.write "<p>��վ�ָ���������..."
		ElseIf action = "delbackup" Then
			Response.Write "<p><br><br>��ɾ����������Ϊ�ļ���������ڽ�ɾ��...<br>"
			Deletefiles(Server.Mappath(DEF_BBS_HomeUrl & DEF_AccessDatabase & ".temp"))
			If Deletefiles(Server.Mappath(DEF_BBS_HomeUrl & DEF_AccessDatabase & ".BAK")) = 1 Then
				Response.write "<p>�ɹ�ɾ��..."
			Else
				Response.write "<p>�����ļ������ڣ�����Ҫɾ��..."
			End If
		Else
			CompactDatabase
		End If
		Response.Write "<p><br><b>������ɣ�<a href=BackupDatabase.asp>������ﷵ��</a></b>"
	End If

End Function

Function Deletefiles(path)

    on error resume next
    Dim fs
    Set fs=Server.CreateObject(DEF_FSOString)
    If fs.FileExists(path) Then
      fs.DeleteFile path,True
      deletefiles = 1
    Else
      deletefiles = 0
    End If
    Set fs = nothing
         
End Function

Function CompactDatabase

	on error resume next
	CloseDatabase
	Dim fs, Engine
	Set fs = CreateObject(DEF_FSOString)
	If fs.FileExists(Server.Mappath(DEF_BBS_HomeUrl & DEF_AccessDatabase)) Then
		Response.Write "<p><br><br>ѹ�����ݿ⿪ʼ����վ��ͣһ���û���ǰ̨����......<br>"
		Application.Lock
		'Application.Contents.RemoveAll()
		FreeApplicationMemory
		application(DEF_MasterCookies & "SiteEnableFlagzoieiu") = 0
		application(DEF_MasterCookies & "SiteDisbleWhyszoieiu") = "��̳��ͣ�У����Ժ򼸷��Ӻ�����..."
		Application.UnLock
		Set Engine = CreateObject("JRO.JetEngine")
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath(DEF_BBS_HomeUrl & DEF_AccessDatabase), "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath(DEF_BBS_HomeUrl & DEF_AccessDatabase & ".temp")
		If Err Then
			Response.Write "<font color=red class=redfont>���ݿ�ѹ��ʧ�ܣ����ܿռ䲻֧�ִ˲�����</font>"
			err.Clear
			Exit Function
		End If
		fs.CopyFile Server.Mappath(DEF_BBS_HomeUrl & DEF_AccessDatabase & ".temp"),Server.Mappath(DEF_BBS_HomeUrl & DEF_AccessDatabase)
		If Err Then
			Response.Write "<font color=red class=redfont>���ݿ�ѹ���ɹ������޷��滻ԭ���ݿ⣬ѹ���ɹ�������ݿ���Ϊ " &  Server.Mappath(DEF_BBS_HomeUrl & DEF_AccessDatabase) & ".temp </font>"
			err.Clear
			Exit Function
		End If
		fs.DeleteFile(Server.Mappath(DEF_BBS_HomeUrl & DEF_AccessDatabase & ".temp"))
		If Err Then
			Response.Write "<font color=red class=redfont>���ݿ�ѹ���ɹ��������滻ԭ���ݿ�ɹ�����ɾ����ʱ�ļ�ʧ�ܣ����ֶ�ɾ�����ݿ�Ŀ¼�����.temp�ļ���</font>"
			err.Clear
			Exit Function
		End If
		Set fs = Nothing
		Set Engine = nothing
		Response.write "<p>ѹ�����ݿ����..."
		Application.Lock
		application(DEF_MasterCookies & "SiteEnableFlagzoieiu") = 1
		application(DEF_MasterCookies & "SiteDisbleWhyszoieiu") = ""
		Application.UnLock
		Response.write "<p>��վ�ָ���������..."
	Else
		Set fs = Nothing
		Response.Write "<p><br><br>���ݿ����ƻ�·������ȷ. ѹ��ʧ��!" & vbCrLf
	End If
	OpenDatabase

End Function

Function DisplayStringForm

%>
<p>
	���ݿ⽫�Զ����ݳ�BAK�ļ�����������ʱ����ͣ��վ���κη��ʡ�<br>
	����ʱ�������ݿ��С����������������ռ䲻�㣬���ܻ�����ʧ�ܡ�<br>
	������Ŀǰ�����ݿ������ <b><a href="<%=DEF_BBS_HomeUrl & DEF_AccessDatabase%>"><%=DEF_AccessDatabase%></a></b><br>
	������ݣ�ϵͳ���Զ����ݵ��ļ� <b><a href="<%=DEF_BBS_HomeUrl & DEF_AccessDatabase & ".BAK"%>"><%=DEF_AccessDatabase & ".BAK"%></a></b><br>
	<font color=red class=redfont>��������ļ��Ѿ���������ɾ���������ܿ�ʼ���ݡ�<br>
	��ע���ڱ��ݺ��������ݿ⵽���أ�Ȼ��ɾ�����ݵ����ݿ⣬�Է����˶������ء�</font>
	<p>
	<form action=BackupDatabase.asp method=post>
		<input type=hidden name=action value="backup">
		<input type=submit value=��ʼ�������ݿ� class=fmbtn>
	</form>	
	
	<form action=BackupDatabase.asp method=post>
		<input type=hidden name=action value="delbackup">
		<input type=submit value=ɾ���������ݿ� class=fmbtn>
	</form>
	
	<form action=BackupDatabase.asp method=post>
		<input type=hidden name=action value="CompactDatabase">
		<input type=submit value=ѹ�����޸����ݿ� class=fmbtn>
	</form>
	
	
	<b><a href="<%=DEF_BBS_HomeUrl & DEF_AccessDatabase & ".BAK"%>">���ر������ݿ�<%=DEF_AccessDatabase & ".BAK"%></a>
	</b>
	<p>
	<font color=red class=redfont>ע�⣬�������ݿ����ѡ�񡰱��ݵ����ݿ⡱��<br>
	��ǰ��վ�����У����ص�ǰʹ�õ����ݿ���ܽ������𻵵ġ�</font>
	<p>�Уӣ����(����)���ݿⲢ�Ƿ���WEB���棬�ǲ���ֱ�����صģ����¼ftp�����������´���
	<p><font color=Red class=redfont>���棺<b>ѹ�����޸����ݿ�</b>���ܱ�Ȼ�ᵼ����̳��������������ͣ��̳�����У�����ʹ�ñ������ݿ⹦�ܣ������ر��ݺõ����ݿ⣬�ڱ���ʹ��Access���������ѹ���������ϴ��滻���ݿ⡣�ڼ��������ر�֤��̳���ڹر�״̬��ʹ����̳�������ܹرգ���������䲻�����κι���ҳ�棩�����������ѹ�����ݿ�����ϴ��滻���ݿ��ļ��������ٴ�ʹ�ùر���̳���ܡ�
<%

End Function

Function FreeApplicationMemory

	Response.Write "<p><b>�ͷ���̳�����б�</b><table>" & VbCrLf
	Dim Thing
	For Each Thing in Application.Contents
		If Left(Thing,Len(DEF_MasterCookies)) = DEF_MasterCookies Then
			Response.Write "<tr><td><font color=Gray class=grayfont>" & thing & "</font></td><td>&nbsp;"
			If isObject(Application.Contents(Thing)) Then
				Application.Contents(Thing).close
				Set Application.Contents(Thing) = Nothing
				Application.Contents(Thing) = null
				Response.Write "����ɹ��ر�"
			ElseIf isArray(Application.Contents(Thing)) Then
				Set Application.Contents(Thing) = Nothing
				Application.Contents(Thing) = null
				Response.Write "����ɹ��ͷ�"
			Else
				Response.Write htmlencode(Application.Contents(Thing))
				Application.Contents(Thing) = null
			End If
			Response.Write "</td></tr>"
		End If
	Next
	Response.Write "</table>"

End Function%>