<!-- #include file=../inc/BBSsetup.asp -->
<!-- #include file=../inc/Upload_Setup.asp -->
<!-- #include file=../inc/Limit_Fun.asp -->
<!-- #include file=../inc/Board_Popfun.asp -->
<%
DEF_BBS_HomeUrl = "../"
InitDatabase
UpdateOnlineUserAtInfo GBL_board_ID,"ɾ���ռ���Ϣ"

Dim FileID,Upd_SpendFlag
FileID = Left(Request("FileID"),14)
If isNumeric(FileID) = 0 or FileID = "" or InStr(FileID,",") > 0 Then FileID = 0
FileID = cCur(FileID)

If FileID < 0 Then FileID = 0

GBL_CHK_TempStr = ""
If GBL_UserID=0 Then
	GBL_CHK_TempStr = GBL_CHK_TempStr & "��û�е�¼<br>" & VbCrLf
Else
	If DEF_FSOString = "" Then GBL_CHK_TempStr = "��������֧��ɾ���������ܣ�<br>" & VbCrLf
End If
siteHead("   ɾ������")
%>
<script language=javascript>
	window.moveTo(window.screen.width/2-225,window.screen.height/2-18);
</script>
<table align=center border=0 cellpadding=0 cellspacing=0>
<tbody>
<tr> 
	<td height=21 width="650">
	<%
	Dim Rs,SQL,SQLendString,ClearFlag
	If GBL_CHK_TempStr = "" Then
		CheckisBoardMaster
		If DEF_Upd_SpendFlag = 0 and GBL_BoardMasterFlag >=4 Then
			Upd_SpendFlag = 0
		Else
			Upd_SpendFlag = 1
		End If
		
		If CheckSupervisorUserName = 1 or (GetBinarybit(GBL_CHK_UserLimit,11) = 1 and GBL_BoardMasterFlag >=4) Then
			SQL = sql_select("Select ID,PhotoDir,SPhotoDir,UserID from LeadBBS_Upload where id=" & FileID,1)
		Else
			SQL = sql_select("Select ID,PhotoDir,SPhotoDir,UserID from LeadBBS_Upload where UserID=" & GBL_UserID & " and id=" & FileID,1)
		End If
		Set Rs = LDExeCute(SQL,0)
		If Rs.Eof Then
			Rs.Close
			Set Rs = Nothing
			Response.Write "�Ҳ�����¼��<br>" & VbCrLf
		Else
			If Request.Form("DeleteSureFlag")="dk9@dl9s92lw_SWxl" Then
				Dim UploadUserID
				UploadUserID = Rs("UserID")
				%>
				�ɹ�ɾ�����Ϊ<font color=ff0000 class=redfont><%=rs("ID")%></font>���ϴ�����!
				<%
				If Rs("PhotoDir") <> "" Then DeleteFiles(Server.Mappath(Replace(Replace(DEF_BBS_HomeUrl & DEF_BBS_UploadPhotoUrl & Rs("PhotoDir"),"/","\"),"\\","\")))
				If Rs("SPhotoDir") <> "" Then DeleteFiles(Server.Mappath(Replace(Replace(DEF_BBS_HomeUrl & DEF_BBS_UploadPhotoUrl & Rs("SPhotoDir"),"/","\"),"\\","\")))
				Rs.Close
				Set Rs = Nothing
				CALL LDExeCute("Delete from LeadBBS_Upload where id=" & FileID,1)
				If Upd_SpendFlag = 1 and GBL_UserID = cCur(UploadUserID) Then
					If DEF_UploadDeletePoints > 0 Then
						Rs = ",Points=Points-" & DEF_UploadDeletePoints
					Else
						Rs = ",Points=Points+" & (0-DEF_UploadDeletePoints)
					End If
				Else
					Rs = ""
				End If
				CALL LDExeCute("Update LeadBBS_User Set UploadNum=UploadNum-1" & Rs & " where id=" & UploadUserID,1)
				CALL LDExeCute("update LeadBBS_SiteInfo set UploadNum=UploadNum-1",1)
				UpdateStatisticDataInfo -1,5,1
			Else
				%>
				<form name=DellClientForm action=DeleteUpload.asp method=post>
					<input type=hidden name=DeleteSureFlag value="dk9@dl9s92lw_SWxl">
					<input type=hidden name=ClearFlag value="<%=htmlencode(ClearFlag)%>">
					<input type=hidden name=FileID value="<%=htmlencode(FileID)%>">
					<b>ȷ��Ҫɾ�����Ϊ<font color=ff0000 class=redfont><%=htmlencode(rs("ID"))%></font>�ĸ�����</b>
					<%If Upd_SpendFlag = 1 Then%><br>ɾ��������<%
						If DEF_UploadDeletePoints > 0 Then
							Response.Write "<font color=red class=redfont>����" & DEF_UploadDeletePoints & "" & DEF_PointsName(0) & "</font>"
						ElseIf DEF_UploadDeletePoints < 0 Then
							Response.Write "<font color=green class=greenfont title=�����Լ�ɾ��������Ӧ�ı仯>����" & 0-DEF_UploadDeletePoints & "" & DEF_PointsName(0) & "</font>"
						End If
						Response.Write "��ֻ��ɾ���Լ��ϴ��ĸ�������Ч��"
					End If%>
					<p><input type=submit value=ȷ�� class=fmbtn>
					<input type=button value=��ɾ onclick="javascript:window.close();" class=fmbtn>
				</form>
				<%
				Rs.Close
				Set Rs = Nothing
			End If
		End If			
	Else%>
		<table width=96%>
		<tr>
		<td>
		<%Response.Write "<p align=left><font color=ff0000 class=redfont><b>" & GBL_CHK_TempStr & "</b>"
		Response.Write "</p>"%>
		</td>
		</tr>
		</table>
	<%End If%></td>
</tr>

</table>
<html>
<head>
	<title>
		<%=DEF_SiteNameString%>
	</title>
	<meta HTTP-EQUIV="Content-Type" content="text/html; charset=gb2312">
	<link rel="stylesheet" type="text/css" href="<%=DEF_BBS_homeurl%>/inc/style.css">
</head>
<%

closeDataBase

Function DeleteFiles(path)

	If DEF_FSOString = "" Then Exit Function
	on error resume next
	Dim fs
	Set fs = Server.CreateObject(DEF_FSOString)
	If err <> 0 Then
		Err.Clear
		Response.Write "<br>��������֧��FSO��Ӳ���ļ�δɾ����"
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
%>