<!-- #include file=../../inc/BBSSetup.asp -->
<!-- #include file=../../inc/Board_Popfun.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<%
DEF_BBS_HomeUrl = "../../"
Dim GBL_ID
initDatabase
GBL_CHK_TempStr = ""
GBL_ID = checkSupervisorPass

Manage_sitehead DEF_SiteNameString & " - ����Ա",""
frame_TopInfo
DisplayUserNavigate("�鿴���ݿ��ṹ")
If GBL_CHK_Flag=1 Then
	LoginAccuessFul
Else
DisplayLoginForm
End If
frame_BottomInfo
closeDataBase
Manage_Sitebottom("none")

Sub LoginAccuessFul

	If DEF_UsedDataBase <> 0 Then
		GBL_CHK_TempStr = "<div class=alert>Access���ݿⲻ֧��ȫ����������!</div>"
		Exit Sub
	End If
	Dim TBName
	TBName = Request("TB")

	If TBName = "" or Len(TBName) > 255 Then
		GBL_CHK_TempStr = "<div class=alert>�˱�����!</div>"
		Exit Sub
	End If
	DisplayTableColInfo(TBName)

End Sub

Function DisplayTableColInfo(Name)

	Dim Rs,SQL
	Dim N,Tmp,Tmp2
	Response.Write "<div class=frametitle>���ݿ��" & Name & "�ṹ</div>"
	SQL = "exec SP_SpaceUsed '" & Replace(Name,"'","''") & "'"
	Set Rs = LDExeCute(SQL,0)
	If Not Rs.Eof Then Tmp = Rs.GetRows(-1)
	Rs.Close
	Set Rs = Nothing
	If isArray(Tmp) = True Then
		Response.Write "<div class=frameline>������" & Tmp(0,0) & "<br>"
		Response.Write "���е�������" & Tmp(1,0) & "<br>"
		Response.Write "�����Ŀռ�������" & Tmp(2,0) & "<br>"
		Response.Write "���е�������ʹ�õĿռ�����" & Tmp(3,0) & "<br>"
		Response.Write "���е�������ʹ�õĿռ�����" & Tmp(4,0) & "<br>"
		Response.Write "����δ�õĿռ�����" & Tmp(5,0) & "</div>"
	End If
	Set Tmp = Nothing
	Response.Write "<table border=0 cellpadding=0 cellspacing=0 width=100% class=frame_table>"
	N = 1
	SQL = "Select COL_NAME( OBJECT_ID('" & Replace(Name,"'","''") & "')," & N & "),COL_LENGTH('" & Replace(Name,"'","''") & "',COL_NAME( OBJECT_ID('" & Replace(Name,"'","''") & "') ," & N & "))"
	Set Rs = LDExeCute(SQL,0)
	Tmp = Rs(0)
	Tmp2 = Rs(1)
	Rs.Close
	Set Rs = Nothing
	Do While Not(IsNull(Tmp)) and Len(Tmp) > 0
		Response.Write "<tr><td class=tdbox width=200>" & htmlencode(Tmp) & "</td><td class=tdbox>" & Tmp2 & "</td>"
		N = N + 1
		SQL = "Select COL_NAME( OBJECT_ID('" & Replace(Name,"'","''") & "')," & N & "),COL_LENGTH('" & Replace(Name,"'","''") & "',COL_NAME( OBJECT_ID('" & Replace(Name,"'","''") & "') ," & N & "))"
		Set Rs = LDExeCute(SQL,0)
		Tmp = Rs(0)
		Tmp2 = Rs(1)
		Rs.Close
		Set Rs = Nothing
	Loop
	Response.Write "</table>"
	Response.Write "<div class=frameline><a href=TableInfo.asp?tb=LeadBBS_Announce>�������鿴��LeadBBS_Announce��Ϣ</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_Assort>�������鿴��LeadBBS_Assort</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_Boards>�������鿴��LeadBBS_Boards</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_CollectAnc>�������鿴��LeadBBS_CollectAnc</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_ForbidIP>�������鿴��LeadBBS_ForbidIP</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_FriendUser>�������鿴��LeadBBS_FriendUser</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_GoodAssort>�������鿴��LeadBBS_GoodAssort</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_InfoBox>�������鿴��LeadBBS_InfoBox</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_IPAddress>�������鿴��LeadBBS_IPAddress</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_Link>�������鿴��LeadBBS_Link</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_onlineUser>�������鿴��LeadBBS_onlineUser</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_Setup>�������鿴��LeadBBS_Setup</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_SiteInfo>�������鿴��LeadBBS_SiteInfo</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_SpecialUser>�������鿴��LeadBBS_SpecialUser</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_TopAnnounce>�������鿴��LeadBBS_TopAnnounce</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_Upload>�������鿴��LeadBBS_Upload</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_User>�������鿴��LeadBBS_User</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_UserFace>�������鿴��LeadBBS_UserFace</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_VoteItem>�������鿴��LeadBBS_VoteItem</a>"
	Response.Write "<br><a href=TableInfo.asp?tb=LeadBBS_VoteUser>�������鿴��LeadBBS_VoteUser</a></div>"

End Function%>