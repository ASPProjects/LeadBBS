<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/UBBicon_Setup.ASP -->
<!-- #include file=../../inc/Board_popfun.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<%
DEF_BBS_HomeUrl = "../../"
Const MaxLinkNum = 100
initDatabase
GBL_CHK_TempStr = ""
CheckSupervisorPass

Dim Form_DEF_UBBiconNote
Redim Form_DEF_UBBiconNote(DEF_UBBiconNumber)

GetDefaultValue

SiteHead(DEF_SiteNameString & " - ����Ա")
UserTopicTopInfo
DisplayUserNavigate("��̳����ע�Ͳ�������")
If GBL_CHK_Flag=1 Then
	UBBiconSetup
Else%>
	<table width=96%>
	<tr>
	<td>
	<%
	If Request("submitflag")="" Then
		Response.Write "<br><b>���ȵ�¼</b>"
	Else
		Response.Write "<br><p align=left><font color=ff0000 class=RedFont><b>" & GBL_CHK_TempStr & "</b></font>"
	End If
	DisplayLoginForm
	Response.Write "</p>"%>
	</td>
	</tr>
	</table>
<%End If
UserTopicBottomInfo
closeDataBase
SiteBottom
If GBL_ShowBottomSure = 1 Then Response.Write GBL_SiteBottomString

Function UBBiconSetup

%>
<form name="pollform3sdx" method="post" action="UBBiconSetup.asp">
<input type="hidden" name="SubmitFlag" value=yes>
<p>
		���ã�<a href=SiteSetup.asp>��̳���ò���</a> <a href=UploadSetup.asp>�ϴ�����</a>
		<a href=../User/UserSetup.asp>�û�ע�����</a>
		<a href=UbbcodeSetup.asp>UBB�������</a>
		<font color=gray class=GrayFont>UBB����ע��</font>
		</b>
		<br><font color=8888888 class=GrayFont>(����Ϊ��վ��������ע���޸ģ���������ý��ᷢ�����ش���)<br><br>
		��������ú�����վ�����������У��뽫LeadBBS���°��inc/UBBicon_Setup.ASP���ǻ�ȥ</font>
</p>
<%If Request.Form("SubmitFlag") <> "" Then
	GetFormValue
End If%>
<b><font color=ff0000 class=RedFont><%=GBL_CHK_TempStr%></font></b>
<p>
<%
If Request.Form("SubmitFlag") <> "" Then
	If GBL_CHK_TempStr <> "" Then
		DisplayDatabaseLink
	Else
		MakeDataBaseLinkFile
		Exit Function
	End If
Else
	DisplayDatabaseLink
End If
%>
<br>
<input type=submit name=�ύ value=�ύ class=fmbtn>
<input type=reset name=ȡ�� value=ȡ�� class=fmbtn>
</form>
<%

End Function

Function DisplayDatabaseLink

	Dim n,m
	%>
	<table border=0 cellpadding=5 cellspacing=1 width="100%" bgcolor="<%=DEF_BBS_LightColor%>" class=TBBG1>
	<tr class=TBBG9>
		<td valign=top><b>����ע��</b></td>
		<td>
			<table border=0 cellpadding=1 cellspacing=0>
			<tr>
				<td>&nbsp;���</td>
				<td>&nbsp;����</td>
				<td>&nbsp;ע��</td>
			</tr><%
		m = Ubound(DEF_UBBiconNote)
		For n = 0 to DEF_UBBiconNumber - 1
			If n = 0 or (n mod 16) = 0 Then Response.Write "<tr><td colspan=3>&nbsp;</td></tr>"
			If n > m Then
				%>
				<tr>
					<td>&nbsp;&nbsp;<%=Right(" " & n + 1,2)%></td>
					<td>&nbsp;<img src="<%=DEF_BBS_HomeUrl%>images/UBBicon/em<%=Right("0" & n + 1,2)%>.GIF" width=20 height=20 align=absmiddle border=0></td>
					<td>&nbsp;<input class=fminpt type="text" name="Form_DEF_UBBiconNote<%=N%>" maxlength="100" size="50" value=""></td>
				</tr>
				<%
			Else
				%>
				<tr>
					<td>&nbsp;&nbsp;<%=Right(" " & n + 1,2)%></td>
					<td>&nbsp;<img src="<%=DEF_BBS_HomeUrl%>images/UBBicon/em<%=Right("0" & n + 1,2)%>.GIF" width=20 height=20 align=absmiddle border=0></td>
					<td>&nbsp;<input class=fminpt type="text" name="Form_DEF_UBBiconNote<%=N%>" maxlength="100" size="50" value="<%=htmlencode(Form_DEF_UBBiconNote(n))%>"></td>
				</tr>
				<%
			End If
		Next
		%>
			</table></td>
	</tr>
	</table>
	<%

End Function

Function GetDefaultValue

	Dim N
	For N= 0 to Ubound(DEF_UBBiconNote)
		If N > DEF_UBBiconNumber Then Exit Function
		Form_DEF_UBBiconNote(n) = DEF_UBBiconNote(N)
	Next

End Function

Function GetFormValue

	Dim n
	For n = 0 to DEF_UBBiconNumber
		Form_DEF_UBBiconNote(n) = Trim(Request.Form("Form_DEF_UBBiconNote" & N))
	Next
	
	For n = 0 to DEF_UBBiconNumber
		If inStr(Form_DEF_UBBiconNote(n),"""") or inStr(Form_DEF_UBBiconNote(n),"%") Then
			GBL_CHK_TempStr = "��" & N & "��ű����������ܰ��������Ż�ٷֺ�<br>" & VbCrLf
		End If
	Next

End Function

Function MakeDataBaseLinkFile

	Dim n,TempStr
	TempStr = ""
	TempStr = TempStr & chr(60) & "%" & VbCrLf
	TempStr = TempStr & "Dim DEF_UBBiconNote" & VbCrLf	
	TempStr = TempStr & "DEF_UBBiconNote = Array("
	For n = 0 to DEF_UBBiconNumber - 1
		If n = 0 Then
			TempStr = TempStr & """" & Form_DEF_UBBiconNote(n) & """"
		Else
			TempStr = TempStr & ",""" & Form_DEF_UBBiconNote(n) & """"
		End If
	Next
	TempStr = TempStr & ")" & VbCrLf

	TempStr = TempStr & "%" & chr(62) & VbCrLf
	
	ADODB_SaveToFile TempStr,"../../inc/UBBicon_Setup.ASP"
	If GBL_CHK_TempStr = "" Then
		Response.Write "<br><font color=Green class=GreenFont>2.�ɹ�������ã�</font>"
	Else
		%><%=GBL_CHK_TempStr%><br>��������֧������д���ļ����ܣ���ʹ��FTP�ȹ��ܣ���<font color=Red Class=RedFont>../../inc/UBBicon_Setup.ASP</font>�ļ��滻�ɿ�������(ע�ⱸ��)<p>
		<textarea name="fileContent" cols="80" rows="30" class=fmtxtra><%=Server.htmlencode(TempStr)%></textarea><%
		GBL_CHK_TempStr = ""
	End If

End Function%>