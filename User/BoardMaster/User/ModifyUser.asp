<!-- #include file=../../../inc/BBSsetup.asp -->
<!-- #include file=../../../inc/Upload_Setup.asp -->
<!-- #include file=../../../inc/Board_Popfun.asp -->
<!-- #include file=../inc/BoardMaster_Fun.asp -->
<%
DEF_BBS_HomeUrl = "../../../"
initDatabase
GBL_CHK_TempStr = ""

SiteHead(DEF_SiteNameString & " - " & DEF_PointsName(6) & "����")

UserTopicTopInfo
DisplayUserNavigate("ǿ���޸��û�����")%>
<br><br><%
If GBL_CHK_Flag=1 and BDM_isBoardMasterFlag = 1 and BDM_SpecialPopedomFlag = 1 Then
	LoginAccuessFul
Else%>
<table width=96%>
<tr>
	<td>
<%
	If Request("submitflag")="" Then
		Response.Write "<br><b>���ȵ�¼</b>"
	Else
		Response.Write "<br><p align=left><font color=ff0000 class=redfont><b>" & GBL_CHK_TempStr & "</b></font>"
	End If
	DisplayLoginForm
	Response.Write "</p>"%>
	</td>
</tr>
</table><%
End If

UserTopicBottomInfo
closeDataBase
SiteBottom
If GBL_ShowBottomSure = 1 Then Response.Write GBL_SiteBottomString

Dim GBL_ModifyMode,GBL_UserName,GBL_UserName_UserID,GBL_UserName_FaceUrl
Dim GBL_UserName_UnderWrite,GBL_UserName_UserTitle
GBL_ModifyMode = 0

Function LoginAccuessFul

	If Request.Form("submitflag") <> "" Then
		CheckForm
		If GBL_CHK_TempStr = "" Then
			ModifyUser
			Response.Write GBL_CHK_TempStr
			GBL_CHK_TempStr = ""
			DisplayForm
		Else
			DisplayForm
		End If
	Else
		DisplayForm
	End If

End Function

Function ModifyUser

	Response.Write "<p><b>��ʼ����û�<u>" & htmlencode(GBL_UserName) & "</u>���������ϣ�</b></p>" & VbCrLf
	If inStr(GBL_ModifyMode,",1,") Then
		If GBL_UserName_FaceUrl & "" = "" Then
			Response.Write "<p><font color=Red class=redfont>�������ͷ�� ���û�ͷ���Ѿ���Ĭ��ͷ���Թ�������</font></p>"
			DeleteUploadFace(GBL_UserName_UserID)
		Else
			CALL LDExeCute("Update LeadBBS_User Set FaceUrl='' where ID=" & GBL_UserName_UserID,1)
			Response.Write "<p><font color=Green class=greenfont>�������ͷ�� �ɹ������</font></p>"
			DeleteUploadFace(GBL_UserName_UserID)
		End If
	End If

	If inStr(GBL_ModifyMode,",2,") Then
		If GBL_UserName_UnderWrite & "" = "" Then
			Response.Write "<p><font color=Red class=redfont>����û�ǩ���� ���û���ǩ�����ݣ��Թ�������</font></p>"
		Else
			CALL LDExeCute("Update LeadBBS_User Set UnderWrite='',PrintUnderWrite='' where ID=" & GBL_UserName_UserID,1)
			Response.Write "<p><font color=Green class=greenfont>����û�ǩ���� �ɹ������</font></p>"
		End If
	End If

	If inStr(GBL_ModifyMode,",3,") Then
		If GBL_UserName_UserTitle & "" = "" Then
			Response.Write "<p><font color=Red class=redfont>����û�ͷ�Σ� ���û���ͷ�Σ��Թ�������</font></p>"
		Else
			CALL LDExeCute("Update LeadBBS_User Set UserTitle='' where ID=" & GBL_UserName_UserID,1)
			Response.Write "<p><font color=Green class=greenfont>����û�ͷ�Σ� �ɹ������</font></p>"
		End If
	End If

	If CheckSupervisorUserName = 0 Then
		CALL LDExeCute("Update LeadBBS_User Set LastWriteTime=" & GetTimeValue(DEF_Now) & " where ID=" & GBL_UserID,1)
	End If

End Function

Function CheckForm

	If CheckWriteEventSpace = 0 Then
		Response.Write "<b><font color=Red Class=redfont>���Ĳ�����Ƶ�����Ժ������ύ!</font></b> <br>" & VbCrLf
		Exit Function
	End If
	
	GBL_ModifyMode = Replace("," & Left(Request.Form("GBL_ModifyMode"),10) & ","," ","")
	GBL_UserName = Left(Request.Form("GBL_UserName"),20)
	If isNumeric(Replace(GBL_ModifyMode,",","")) = 0 Then
		GBL_CHK_TempStr = "���󣬲���ѡ��ѡ�����"
		Exit Function
	End If

	If GBL_UserName = "" Then
		GBL_CHK_TempStr = "�����������û�����"
		Exit Function
	End If
	
	If CheckUserNameExist(GBL_UserName) = 0 Then
		GBL_CHK_TempStr = "�����û��������ڣ�"
		Exit Function
	End If

End Function

Function DisplayForm

	If GBL_CHK_TempStr <> "" Then Response.Write "<p><font color=ff0000 class=redfont><b>" & GBL_CHK_TempStr & "</b></font></p>"%>

			<%If Request.Form("submitflag") = "LKOkxk2" or Request.Form("submitflag") = "" Then%>
			<p>
		  <b>�����û�����</b>
          <form action=ModifyUser.asp method=post id=fobform name=fobform>
			�� �� ���� <input name=GBL_UserName value="<%=htmlencode(GBL_UserName)%>" class=fminpt><br>
			<input name=submitflag type=hidden value="LKOkxk2">
			ѡ�������<input name=GBL_ModifyMode value=1<%If inStr(GBL_ModifyMode,",1,") Then Response.Write " checked"%> type=checkbox>�������ͷ��
			<input name=GBL_ModifyMode value=2<%If inStr(GBL_ModifyMode,",2,") Then Response.Write " checked"%> type=checkbox>����û�ǩ��
			<input name=GBL_ModifyMode value=3<%If inStr(GBL_ModifyMode,",3,") Then Response.Write " checked"%> type=checkbox>����û�ͷ��
			<br><br>
			<input type=submit value="�ύ" class=fmbtn> <input type=reset value="ȡ��" class=fmbtn></form>
			<br>
			<b>˵����</b><br><br>
			1.����û�����ͷ��󣬴��û�ͷ��ָ�Ϊ��̳���е�ͷ��<br>
			&nbsp; ���Ա��Ϊ1��Ů�Ա��Ϊ2�����Ա�Ϊ0��<br>
			2.����û�ǩ��������ָ�����û�ǩ������ȫ������<br>
			3.����û�ͷ�ν�����ָ�����û�ͷȡ��<br>
			4.ĳЩ�ض��û����ϲ������޸�
			<%End If%>

<%End Function

Rem ���ĳ�û����Ƿ����
Function CheckUserNameExist(UserName)

	If UserName <> "" and inStr(UserName,",") = 0 and inStr(Lcase(DEF_SupervisorUserName),"," & Lcase(UserName) & ",") > 0 Then
		'��������ͬ����ʾ��Ϊ���Է�����Ա���ֱ�й©��ʵ��Ӧ����ʾ����Ա���ܱ�����
		GBL_CHK_TempStr = "�����û���" & htmlencode(UserName) & "�����ڣ�"
		CheckUserNameExist = 0
		Exit Function
	End If

	Dim Rs
	Set Rs = LDExeCute(sql_select("Select ID,UserName,FaceUrl,UnderWrite,UserTitle from LeadBBS_User where UserName='" & Replace(UserName,"'","''") & "'",1),0)
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		CheckUserNameExist = 0
		GBL_CHK_TempStr = "�����û���" & htmlencode(UserName) & "�����ڣ�"
		Exit Function
	Else
		GBL_UserName_UserID = cCur(Rs(0))
		GBL_UserName = Rs(1)
		GBL_UserName_FaceUrl = Rs(2)
		GBL_UserName_UnderWrite = Rs(3)
		GBL_UserName_UserTitle = Rs(4)
	End if
	Rs.Close
	Set Rs = Nothing
		
	CheckUserNameExist = 1

End Function


Function DeleteUploadFace(DelUserID)

	If DEF_FSOString = "" Then
		Response.Write "<p><font color=Red class=redfont>��̳��֧������ɾ���ļ����Թ��ϴ�ͷ��ɾ����</font>"
		Exit Function
	End If
	Dim SQL,Rs
	SQL = "Select ID,PhotoDir,SPhotoDir from LeadBBS_UserFace where UserID=" & DelUserID & " order by ID ASC"
	Set Rs = LDExeCute(SQL,0)
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		Response.Write "<p><b><font color=Red class=redfont>�û����ϴ�ͷ���Թ�ɾ��!</font></b>"
	Else
		If Rs("PhotoDir") <> "" Then DeleteFiles(Server.Mappath(DEF_BBS_HomeUrl & DEF_BBS_UploadPhotoUrl & "face/" & Rs("PhotoDir")))
		If Rs("SPhotoDir") <> "" Then DeleteFiles(Server.Mappath(DEF_BBS_HomeUrl & DEF_BBS_UploadPhotoUrl & "face/" & Rs("SPhotoDir")))
		Rs.Close
		Set Rs = Nothing
		CALL LDExeCute("Delete from LeadBBS_UserFace where UserID=" & DelUserID,1)
		Response.Write "<p><b><font color=green class=greenfont>����û��ϴ�ͷ���ɾ��!</font></b>"
	End If

End Function

Function DeleteFiles(path)

	'on error resume next
	Dim fs
	Set fs=Server.CreateObject(DEF_FSOString)
	If fs.FileExists(path) then
		fs.DeleteFile path,True
		DeleteFiles = 1
	Else
		DeleteFiles = 0
	End If
    Set fs=nothing
         
End Function
%>