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
DisplayUserNavigate("���߱༭�����ļ�")
If GBL_CHK_Flag=1 Then
	%>
	<p><ul>
	<li><a href=SiteEditFileContent.asp?file=-1>�༭���û�ע����̳Э������</a><br>
	<li><a href=SiteEditFileContent.asp?file=-3>���߱༭��ϵ���ǣ��������ǣ�����</a><br>
	</ul>
	<ul>
	<%
	Dim N
	For N = 0 to DEF_BoardStyleStringNum%>
	<li><a href=SiteEditFileContent.asp?file=<%=N%>>�༭�����ʽ����-<%=DEF_BoardStyleString(N)%></a> &nbsp; [<a href=DefineStyleParameter.asp?StyleID=<%=N%>>����������</a>]
	<%Next%>
	</ul>
	ע�⣬�����ķ�������֧���ļ�д�룬������ʹ���������κι��ܣ���Ҫ�ֶ�����Դ������������á�
	<%
Else
DisplayLoginForm
End If
frame_BottomInfo
closeDataBase
Manage_Sitebottom("none")
%>