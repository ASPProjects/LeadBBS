<!-- #include file=../../inc/BBSSetup.asp -->
<!-- #include file=Chat_Fun.asp -->
<!-- #include file=Inc/Chat_Setup.asp -->
<%
Sub Chat_GetWorldChat(User)

	Dim Index,World_Index,Temp

	Index = Session(DEF_MasterCookies & "_Chat_5_Index")
	World_Index = Application(DEF_MasterCookies & "_Chat_S_Index_" & User)
	If Index <> World_Index and Index <> -1 Then
		Response.Write "1 �����µ�˽��" & VbCrLf
	Else
		If isArray(Session(DEF_MasterCookies & "UDT")) = False Then
			Response.Write "9 stop"
		Else
			Temp = Session(DEF_MasterCookies & "UDT")(6)
			If ccur(Temp) = 1 Then
				Response.Write "9 mess"
			Else
				Response.Write "9 null"
			End If
		End If
	End If

End Sub

Sub Chat_GetInfo

	Dim User
	User = Left(Request.Form("User"),20)
	If isArray(Application(DEF_MasterCookies & "_Chat_S_Data_" & User)) = False Then
		Response.Write "9 guest"
		Exit Sub '�ο��޷�����
	Else
		If Application(DEF_MasterCookies & "_Chat_S_ID_" & User) <> cCur(Session.SessionID) Then
			Response.Write "9 stop"
			Exit Sub '�ǵ�ǰ�����޷�����
		End If
	End If
	Dim tmp
	tmp = Session(DEF_MasterCookies & "_Chat_GetTime")
	If Timer < tmp Then
		Session(DEF_MasterCookies & "_Chat_GetTime") = Timer
		tmp = Timer
	End If
	If Timer - tmp < 6 Then
		'��ֹ������������Ϣ
		Response.Write "9 none"
		Exit Sub
	End If
	Session(DEF_MasterCookies & "_Chat_GetTime") = Timer
	Chat_GetWorldChat(User)

End Sub

Chat_GetInfo
%>