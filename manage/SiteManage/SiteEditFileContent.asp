<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Board_popfun.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<%
DEF_BBS_HomeUrl = "../../"
Dim GBL_ID
initDatabase
GBL_CHK_TempStr = ""
GBL_ID = checkSupervisorPass

Manage_sitehead DEF_SiteNameString & " - ����Ա",""
frame_TopInfo
DisplayUserNavigate("�༭�ļ�����")
If GBL_CHK_Flag=1 and GBL_CHK_TempStr = "" Then
	SiteEditFileContent
Else
	DisplayLoginForm
End If
frame_BottomInfo
closeDataBase
Manage_Sitebottom("none")

Sub SiteEditFileContent

	Dim FileName,File,TmpStr
	File = Request("File")
	If isNumeric(File) = 0 Then File = 0
	File = Fix(cCur(File))
	If File < -4 or File > DEF_BoardStyleStringNum or File = -2 Then
		Response.Write "<div class=alert>����,Ҫ�༭���ļ�������!</div>" & VbCrLf
		Exit Sub
	End If
	Select Case File
		Case -3: FileName = "../../User/inc/Contact_info.asp"
				 TmpStr = "<b>���߱༭Contact_info.asp�ļ�����ϵ������Ϣ��</b>��HTML�﷨"
		Case -4: FileName = "../../../other/80bbs/test.txt"
				 TmpStr = "<b>���߱༭../../../other/80bbs/test.txt�ļ�</b>��HTML�﷨"
		'Case -2: FileName = "../../inc/BBSSetup.asp"
		'		 TmpStr = "<b>���߱༭BBSSetup.asp��̳�����ļ�</b>����ע�⣬�޸�ǰ��ñ���BBSSetup.asp�ļ�"
		Case -1: FileName = "../../User/inc/User_Reg.asp"
				 TmpStr = "<b>�༭���û�ע����̳Э������</b>(html��ʽ)"
		Case Else: FileName = "../../inc/style" & File & ".css"
				 TmpStr = "<b>�༭�����ʽ����-" & DEF_BoardStyleString(File) & "</b> CSS�ļ���ʽ"
	End Select
	DisplayEditFileContent FileName,TmpStr,File

End Sub

Sub DisplayEditFileContent(FileName,TmpStr,FileParNum)

	'If DEF_FSOString = "" Then
	'	Response.Write "<p><br><font color=red class=redfont>��̳�����óɲ�֧�����߱༭�ļ�����!</font></p>" & VbCrLf
	'	Exit Sub
	'End If
	Dim fileContent

	If Request.Form("SubmitFlag") = "" Then
		FileContent = ADODB_LoadFile(FileName)
		If GBL_CHK_TempStr <> "" Then
			Response.Write "<p>" & GBL_CHK_TempStr
			GBL_CHK_TempStr = ""
			Exit Sub
		End If
	Else
		fileContent = Request.Form("fileContent")
		Dim TempContent
		TempContent = Lcase(fileContent)
		If inStr(TempContent,"<%") or inStr(TempContent,"include") or inStr(TempContent,"server") Then
			Response.Write "<p><br><font color=red class=redfont>�����в��ܺ���<%��include��Server���ַ�!</font></p>" & VbCrLf
			Exit Sub
		End If
		
		ADODB_SaveToFile fileContent,FileName
		
		If FileName = "../../User/inc/Contact_info.asp" Then
			CALL Update_InsertSetupRID(1051,"inc/Contact_info.ASP",5,fileContent," and ClassNum=" & 5)
		End If
		
		If FileName = "../../User/inc/User_Reg.asp" Then
			CALL Update_InsertSetupRID(1051,"inc/User_Reg.ASP",6,fileContent," and ClassNum=" & 6)
		End If
		
		If GBL_CHK_TempStr <> "" Then
			Response.Write "<p>" & GBL_CHK_TempStr
			GBL_CHK_TempStr = ""
			Exit Sub
		Else
			Response.Write "<p><font color=green class=greenfont><b>�ɹ������ļ����ݣ�</b></font></p>" & VbCrLf
		End If
	End If
	%>
	<div class=frameline><%=TmpStr%></div>
	<form action=SiteEditFileContent.asp method=post>
		<input type=hidden value=<%=FileParNum%> name=File>
		<input type=hidden value=yes name=SubmitFlag>
		<div class=frameline>
		<textarea name="fileContent" cols="80" rows="35" class=fmtxtra><%If fileContent <> "" Then Response.Write VbCrLf & server.htmlEncode(fileContent)%></textarea><p>
		</div>

		<div class=frameline>
		<input type="submit" name="save" value="�޸�" class=fmbtn>
		<input type="reset" name="Reset" value="ȡ��" class=fmbtn>
		</div>
	</form>
	<%

End Sub
%>