<%
on error resume next
application.contents.removeall
If err Then
	FreeApplicationMemory
	Err.clear
End If

Function FreeApplicationMemory

	Dim Thing
	For Each Thing in Application.Contents
		If Left(Thing,Len(DEF_MasterCookies)) = DEF_MasterCookies Then
			If isObject(Application.Contents(Thing)) Then
				Application.Contents(Thing).close
				Set Application.Contents(Thing) = Nothing
				Application.Contents(Thing) = null
				'Response.Write "����ɹ��ر�"
			ElseIf isArray(Application.Contents(Thing)) Then
				Set Application.Contents(Thing) = Nothing
				Application.Contents(Thing) = null
				'Response.Write "����ɹ��ͷ�"
			Else
				Response.Write htmlencode(Application.Contents(Thing))
				Application.Contents(Thing) = null
			End If
		End If
	Next
	on error resume next
	Application.Contents.RemoveAll

End Function%>
��̳�ɹ�������� ʹ����ɺ�ע��ʹ��FTPɾ�����ļ�