<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Board_Popfun.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<!-- #include file=inc/ForumBoard_fun.asp -->
<%
DEF_BBS_HomeUrl = "../../"
Dim GBL_ID,GBL_DELETEID
initDatabase
GBL_CHK_TempStr = ""
GBL_ID = checkSupervisorPass

Manage_sitehead DEF_SiteNameString & " - ����Ա",""
GBL_CHK_TempStr = ""
frame_TopInfo
DisplayUserNavigate("ɾ����̳�հ���")
GBL_DELETEID = Left(Request("GBL_DELETEID"),14)
If isNumeric(GBL_DELETEID)=0 Then GBL_DELETEID=0
GBL_DELETEID = cCur(GBL_DELETEID)
If GBL_CHK_Flag=1 Then
	If Request.Form("DeleteSuer")="E72ksiOkw2" Then
			GBL_BoardID = GBL_DELETEID
			If DeleteForumBoard(GBL_DELETEID)>0 Then
				Response.Write "<p><font color=008800 class=greenfont><b>�Ѿ��ɹ�ɾ��IDΪ" & GBL_DELETEID & "����̳���棡</b></font></p>"
			Else
				Response.Write "<p><font color=ff0000 class=redfont><b>" & GBL_CHK_TempStr & "</b></font></p>"
			End If
		Else
			%><p><form action=ForumBoardDelete.asp method=post>
			ע�⣺ɾ����̳��ɾ��һ����̳���Ӻ������û���Ϣ��<br>
			��������ɾ��ǰ���ȷ�ϴ���̳�Ѿ�Ϊ�ա�<br><br>
			<font color=Red class=redfont>����</font>����ɾ������ʱ��ͬʱɾ���˰����ר����Ϣ<br><br>
			<b><font color=ff0000 class=redfont>ȷ����Ϣ�� ���Ҫɾ������̳������<br><br>
			<input type=hidden name=GBL_DELETEID value="<%=urlencode(GBL_DELETEID)%>">
			<input type=hidden name=DeleteSuer value="E72ksiOkw2">
			
			<input type=button value=���� onclick="javascript:history.go(-1);" class=fmbtn>
			<input type=submit value=ȷ��ɾ�� class=fmbtn>
			</form>
		<%End If
Else
DisplayLoginForm
End If
closeDataBase
frame_BottomInfo
Manage_Sitebottom("none")
%>