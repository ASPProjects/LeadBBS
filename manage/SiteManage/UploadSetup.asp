<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Upload_Setup.asp -->
<!-- #include file=../../inc/Board_popfun.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<%
DEF_BBS_HomeUrl = "../../"
Const MaxLinkNum = 100
initDatabase
GBL_CHK_TempStr = ""
CheckSupervisorPass

Dim Form_DEF_BBS_UploadPhotoUrl,Form_DEF_EnableDotNetUpload,Form_DEF_UploadFileType
Dim Form_DEF_FileMaxBytes,Form_DEF_FaceMaxBytes
Dim Form_DEF_UploadSpendPoints,Form_DEF_UploadDeletePoints,Form_DEF_UploadOneDayMaxNum,Form_DEF_UploadFaceNeedPoints
Dim Form_DEF_UploadVersionString,Form_DEF_Upd_SpendFlag,Form_DEF_UploadOnceNum,Form_DEF_UploadSwidth,Form_DEF_UploadSheight,Form_DEF_DownSpend

GetDefaultValue

Manage_sitehead DEF_SiteNameString & " - ����Ա",""
frame_TopInfo
DisplayUserNavigate("��̳�ϴ���������")
If GBL_CHK_Flag=1 Then
	UploadSetup
Else
DisplayLoginForm
End If
frame_BottomInfo
closeDataBase
Manage_Sitebottom("none")

Function UploadSetup

%>
<form name="pollform3sdx" method="post" action="UploadSetup.asp">
<input type="hidden" name="SubmitFlag" value=yes>
<p>
		���ã�<a href=SiteSetup.asp>��̳���ò���</a> <span class=grayfont>�ϴ�����</span>
		<a href=../User/UserSetup.asp>�û�ע�����</a>
		<a href=UbbcodeSetup.asp>UBB�������</a>
		<br><span class=grayfont>(����Ϊ��վ��������ע���޸ģ���������ý��ᷢ�����ش���)<br><br>
		��������ú�����վ�����������У��뽫LeadBBS���°��inc/Upload_Setup.asp���ǻ�ȥ</span>
</p>
<%If Request.Form("SubmitFlag") <> "" Then
	GetFormValue
End If%>
<b><span class=redfont><%=GBL_CHK_TempStr%></span></b>
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
</p>
<input type=submit name=�ύ value=�ύ class=fmbtn>
<input type=reset name=ȡ�� value=ȡ�� class=fmbtn>
</form>
<%

End Function

Sub DisplayDatabaseLink

		%>
		<table border=0 cellpadding=0 cellspacing=0 width="100%" class=frame_table>
		<tr>
			<td class=tdbox width=100>�ϴ�Ŀ¼</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_BBS_UploadPhotoUrl" maxlength="50" size="30" value="<%=htmlencode(Form_DEF_BBS_UploadPhotoUrl)%>"><span class=grayfont> Ĭ��Ϊimages/upload/���������̳��Ŀ¼����</span></td>
		</tr>
		<tr>
			<td class=tdbox>�ϴ���ʽ</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_EnableDotNetUpload value=0<%If Form_DEF_EnableDotNetUpload = 0 Then%> checked<%End If%>></td><td>FileUp����ϴ�</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_EnableDotNetUpload value=1<%If Form_DEF_EnableDotNetUpload = 1 Then%> checked<%End If%>></td><td>ʹ��.Net�ϴ�</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_EnableDotNetUpload value=2<%If Form_DEF_EnableDotNetUpload = 2 Then%> checked<%End If%>></td><td>ʹ�������</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>�ϴ�����</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_UploadFileType" maxlength="1024" size="50" value="<%=htmlencode(Form_DEF_UploadFileType)%>"><span class=grayfont><br>�����ϴ���������ļ����ͣ���չ������Сд����ð�ŷָ�</span></td>
		</tr>
		<tr>
			<td class=tdbox>�ļ���С</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_FileMaxBytes" maxlength="8" size="10" value="<%=htmlencode(Form_DEF_FileMaxBytes)%>"><span class=grayfont>(�Զ��������ϴ��ļ������ֵ����λ�ֽ�)</span></td>
		</tr>
		<tr>
			<td class=tdbox>ͷ���С</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_FaceMaxBytes" maxlength="8" size="10" value="<%=htmlencode(Form_DEF_FaceMaxBytes)%>"><span class=grayfont>(�Զ��������ϴ�ͷ������ֵ����λ�ֽڣ���Ϊ���ʾ��ֹ)</span></td>
		</tr>
		<tr>
			<td class=tdbox>�ϴ�����<%=DEF_PointsName(0)%></td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_UploadSpendPoints" maxlength="14" size="10" value="<%=htmlencode(Form_DEF_UploadSpendPoints)%>"><span class=grayfont><br>�û�ÿ��һ������������<%=DEF_PointsName(0)%>������Ϊ����ֵ������Ϊ����ֵ����Ϊ����</span></td>
		</tr>
		<tr>
			<td class=tdbox>ɾ������<%=DEF_PointsName(0)%></td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_UploadDeletePoints" maxlength="14" size="10" value="<%=htmlencode(Form_DEF_UploadDeletePoints)%>"><span class=grayfont><br>�û�ÿɾ��һ���Լ��ϴ��ĸ���������<%=DEF_PointsName(0)%>������Ϊ����ֵ������Ϊ����ֵ����Ϊ����(�����Լ�ɾ���Լ��ĸ�������Ч)</span></td>
		</tr>
		<tr>
			<td class=tdbox>�����Ƿ�Ӱ�����</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_Upd_SpendFlag value=0<%If Form_DEF_Upd_SpendFlag = 0 Then%> checked<%End If%>></td><td>��Ӱ��</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_Upd_SpendFlag value=1<%If Form_DEF_Upd_SpendFlag = 1 Then%> checked<%End If%>></td><td>�����趨Ӱ��</td>
          		<td><span class=grayfont><br>�����趨�ϴ���ɾ������,�Ƿ�Ӱ��<%=DEF_PointsName(8)%>������Ȩ����Ա��<%=DEF_PointsName(0)%></span>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>ÿ���ϴ�����</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_UploadOneDayMaxNum" maxlength="14" size="10" value="<%=htmlencode(Form_DEF_UploadOneDayMaxNum)%>"><span class=grayfont><br>�����û�ÿ�������ϴ�����฽���������������ô���100������Ϊ0��ʾ������</span></td>
		</tr>
		<tr>
			<td class=tdbox>�ϴ�ͷ����Ҫ<%=DEF_PointsName(0)%></td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_UploadFaceNeedPoints" maxlength="14" size="10" value="<%=htmlencode(Form_DEF_UploadFaceNeedPoints)%>"><span class=grayfont><br>�����û�ֻ��<%=DEF_PointsName(0)%>ֵ���ڴ�ֵ�����ϴ���̳ͷ��</span></td>
		</tr>
		<tr>
			<td class=tdbox>ˮӡ����</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_UploadVersionString" maxlength="50" size="30" value="<%=htmlencode(Form_DEF_UploadVersionString)%>"><span class=grayfont> ���ϴ���ͼƬ���������Զ�����ˮӡ����</span></td>
		</tr>
		<tr>
			<td class=tdbox>�������������</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_UploadOnceNum" maxlength="14" size="10" value="<%=htmlencode(Form_DEF_UploadOnceNum)%>"><span class=grayfont><br>��һ�α༭�򷢱���,����ͬʱ�ϴ�����฽������</span></td>
		</tr>
		<tr>
			<td class=tdbox>ͼƬ�������Ը߿�</td>
			<td class=tdbox>
			�����ͼ��� <input class=fminpt type="text" name="Form_DEF_UploadSwidth" maxlength="14" size="10" value="<%=htmlencode(Form_DEF_UploadSwidth)%>">
			�����ͼ�߶� <input class=fminpt type="text" name="Form_DEF_UploadSheight" maxlength="14" size="10" value="<%=htmlencode(Form_DEF_UploadSheight)%>">
			<span class=grayfont><br>��Ҫ����: ����ͼ�����. ��λ: ����<br>���ϴ���,���Զ�����ͼƬ��С��������ϴ��趨Ҫ������ͼ</span></td>
		</tr>
		</table>
		<%

End Sub

Sub GetDefaultValue

	Form_DEF_BBS_UploadPhotoUrl = DEF_BBS_UploadPhotoUrl
	Form_DEF_EnableDotNetUpload = DEF_EnableDotNetUpload
	Form_DEF_UploadFileType = DEF_UploadFileType	
	Form_DEF_FileMaxBytes = DEF_FileMaxBytes
	Form_DEF_FaceMaxBytes = DEF_FaceMaxBytes
	Form_DEF_UploadSpendPoints = DEF_UploadSpendPoints
	Form_DEF_UploadDeletePoints = DEF_UploadDeletePoints
	Form_DEF_UploadOneDayMaxNum = DEF_UploadOneDayMaxNum
	Form_DEF_UploadFaceNeedPoints = DEF_UploadFaceNeedPoints
	Form_DEF_UploadVersionString = DEF_UploadVersionString
	Form_DEF_Upd_SpendFlag = DEF_Upd_SpendFlag
	Form_DEF_UploadOnceNum = DEF_UploadOnceNum
	Form_DEF_UploadSwidth = DEF_UploadSwidth
	Form_DEF_UploadSheight = DEF_UploadSheight

End Sub

Sub GetFormValue

	Form_DEF_BBS_UploadPhotoUrl = Trim(Request.Form("Form_DEF_BBS_UploadPhotoUrl"))
	Form_DEF_EnableDotNetUpload = Trim(Request.Form("Form_DEF_EnableDotNetUpload"))
	Form_DEF_UploadFileType = Trim(Request.Form("Form_DEF_UploadFileType"))
	Form_DEF_FileMaxBytes = Trim(Request.Form("Form_DEF_FileMaxBytes"))
	Form_DEF_FaceMaxBytes = Trim(Request.Form("Form_DEF_FaceMaxBytes"))
	Form_DEF_UploadSpendPoints = Trim(Request.Form("Form_DEF_UploadSpendPoints"))
	Form_DEF_UploadDeletePoints = Trim(Request.Form("Form_DEF_UploadDeletePoints"))
	Form_DEF_UploadOneDayMaxNum = Trim(Request.Form("Form_DEF_UploadOneDayMaxNum"))
	Form_DEF_UploadFaceNeedPoints = Trim(Request.Form("Form_DEF_UploadFaceNeedPoints"))
	Form_DEF_UploadVersionString = Trim(Request.Form("Form_DEF_UploadVersionString"))
	Form_DEF_Upd_SpendFlag = Trim(Request.Form("Form_DEF_Upd_SpendFlag"))
	
	Form_DEF_UploadOnceNum = Trim(Request.Form("Form_DEF_UploadOnceNum"))
	Form_DEF_UploadSwidth = Trim(Request.Form("Form_DEF_UploadSwidth"))
	Form_DEF_UploadSheight = Trim(Request.Form("Form_DEF_UploadSheight"))

	If inStr(Form_DEF_BBS_UploadPhotoUrl,"""") or inStr(Form_DEF_BBS_UploadPhotoUrl,"%") Then GBL_CHK_TempStr = "�ϴ�Ŀ¼���ܰ��������Ż�ٷֺ�<br>" & VbCrLf
	If isNumeric(Form_DEF_EnableDotNetUpload) = 0 Then GBL_CHK_TempStr = "�ϴ���ʽ����Ϊ����<br>" & VbCrLf
	If inStr(Form_DEF_UploadFileType,"""") or inStr(Form_DEF_UploadFileType,"%") Then GBL_CHK_TempStr = "�ϴ����Ͳ��ܰ��������Ż�ٷֺ�<br>" & VbCrLf
	Form_DEF_UploadFileType = Replace(Replace(Trim(LCase(Form_DEF_UploadFileType))," ",""),"?","")
	
	Dim FobType,N
	FobType = Array("htw","ida","asp","asa","idq","cer","cdx","htr","idc","shtm","shtml","stm","printer","asax","ascx","ashx","asmx","aspx","axd","vsdisco","rem","soap","config","cs","csproj","vb","vbproj","webinfo","licx","resx","resources","php","cgi")
	For N = 0 to Ubound(FobType)
		If inStr(":" & Form_DEF_UploadFileType & ":",":." & FobType(N) & ":") Then
			GBL_CHK_TempStr = "Ϊ�˰�ȫ���ϴ����Ͳ���ʹ����չ�� " & FobType(N) & " ��<br>" & VbCrLf
			Exit Sub
		End If
	Next
	If isNumeric(Form_DEF_FileMaxBytes) = 0 Then GBL_CHK_TempStr = "�ϴ��ļ���С����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_FaceMaxBytes) = 0 Then GBL_CHK_TempStr = "�ϴ�ͷ���С����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_UploadSpendPoints) = 0 Then GBL_CHK_TempStr = "�ϴ�����" & DEF_PointsName(0) & "ֵ����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_UploadDeletePoints) = 0 Then GBL_CHK_TempStr = "ɾ������" & DEF_PointsName(0) & "ֵ����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_UploadOneDayMaxNum) = 0 Then GBL_CHK_TempStr = "ÿ���ϴ�������������Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_UploadFaceNeedPoints) = 0 Then GBL_CHK_TempStr = "�ϴ�ͷ����Ҫ" & DEF_PointsName(0) & "����Ϊ����<br>" & VbCrLf
	If inStr(Form_DEF_UploadVersionString,"""") or inStr(Form_DEF_UploadVersionString,"%") Then GBL_CHK_TempStr = "�ϴ��ļ�ˮӡ��ʶ���ܰ��������Ż�ٷֺ�<br>" & VbCrLf
	If isNumeric(Form_DEF_Upd_SpendFlag) = 0 Then GBL_CHK_TempStr = "�����Ƿ�Ӱ��" & DEF_PointsName(8) & "" & DEF_PointsName(0) & "����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_UploadOnceNum) = 0 Then GBL_CHK_TempStr = "���������������Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_UploadSwidth) = 0 Then GBL_CHK_TempStr = "ͼƬ�������Կ���Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_UploadSheight) = 0 Then GBL_CHK_TempStr = "ͼƬ�������Ը���Ϊ����<br>" & VbCrLf

End Sub

Sub MakeDataBaseLinkFile

	Dim n,TempStr
	TempStr = ""
	TempStr = TempStr & chr(60) & "%" & VbCrLf
	TempStr = TempStr & "const DEF_BBS_UploadPhotoUrl=" & Chr(34) & Replace(Form_DEF_BBS_UploadPhotoUrl,Chr(34),Chr(34) & Chr(34)) & Chr(34) & VbCrLf
	TempStr = TempStr & "const DEF_EnableDotNetUpload = " & Form_DEF_EnableDotNetUpload & VbCrLf
	TempStr = TempStr & "const DEF_UploadFileType=" & Chr(34) & Replace(Form_DEF_UploadFileType,Chr(34),Chr(34) & Chr(34)) & Chr(34) & VbCrLf
	TempStr = TempStr & "Const DEF_FileMaxBytes = " & Form_DEF_FileMaxBytes & VbCrLf
	TempStr = TempStr & "Const DEF_FaceMaxBytes = " & Form_DEF_FaceMaxBytes & VbCrLf
	TempStr = TempStr & "Const DEF_UploadSpendPoints = " & Form_DEF_UploadSpendPoints & VbCrLf
	TempStr = TempStr & "Const DEF_UploadDeletePoints = " & Form_DEF_UploadDeletePoints & VbCrLf
	TempStr = TempStr & "Const DEF_UploadOneDayMaxNum = " & Form_DEF_UploadOneDayMaxNum & VbCrLf
	TempStr = TempStr & "Const DEF_UploadFaceNeedPoints = " & Form_DEF_UploadFaceNeedPoints & VbCrLf
	TempStr = TempStr & "const DEF_UploadVersionString=" & Chr(34) & Replace(Form_DEF_UploadVersionString,Chr(34),Chr(34) & Chr(34)) & Chr(34) & VbCrLf
	TempStr = TempStr & "const DEF_Upd_SpendFlag = " & Form_DEF_Upd_SpendFlag & VbCrLf
	TempStr = TempStr & "const DEF_UploadOnceNum = " & Form_DEF_UploadOnceNum & VbCrLf
	TempStr = TempStr & "const DEF_UploadSwidth = " & Form_DEF_UploadSwidth & VbCrLf
	TempStr = TempStr & "const DEF_UploadSheight = " & Form_DEF_UploadSheight & VbCrLf

	TempStr = TempStr & "%" & chr(62) & VbCrLf
	
	ADODB_SaveToFile TempStr,"../../inc/Upload_Setup.asp"
	CALL Update_InsertSetupRID(1051,"inc/Upload_Setup.ASP",3,TempStr," and ClassNum=" & 3)
	If GBL_CHK_TempStr = "" Then
		Response.Write "<br><span class=greenfont>2.�ɹ�������ã�</span>"
	Else
		%><%=GBL_CHK_TempStr%><br>��������֧������д���ļ����ܣ���ʹ��FTP�ȹ��ܣ���<span class=redfont>../../inc/Upload_Setup.asp</span>�ļ��滻�ɿ�������(ע�ⱸ��)<p>
		<textarea name="fileContent" cols="80" rows="30" class=fmtxtra><%=Server.htmlencode(TempStr)%></textarea><%
		GBL_CHK_TempStr = ""
	End If

End Sub%>