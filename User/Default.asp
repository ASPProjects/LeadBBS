<!-- #include file=../inc/BBSSetup.asp -->
<!-- #include file=../inc/Board_Popfun.asp -->
<!-- #include file=inc/UserTopic.asp -->
<%
DEF_BBS_HomeUrl = "../"

Main

Sub Main

	initDatabase
	GBL_CHK_TempStr = ""
	
	BBS_SiteHead DEF_SiteNameString & " - �û���ҳ",0,"<span class=navigate_string_step>�û���ҳ</span>"
	UpdateOnlineUserAtInfo GBL_board_ID,"�û���ҳ"

	UserTopicTopInfo("user")

	If GBL_CHK_Flag = 1 Then
		LoginAccuessFul
	Else
		If Request("submitflag")="" Then
			DisplayLoginForm("���ȵ�¼")
		Else
			DisplayLoginForm(GBL_CHK_TempStr)
		End If
	End If
	UserTopicBottomInfo
	closeDataBase
	SiteBottom

End Sub

Sub LoginAccuessFul%>

	<b>�������⣺</b>

	<div class=title>1.����޸��Ѿ�ע������ϣ�</div>

	<div class=value3>�����ߵ�<b>�޸��ҵ�����</b>�Ϳɽ����޸��Լ������ϡ�</div>

	<div class=title>2.�Ƿ����޸��û�����</div>

	<div class=value3>Ĭ�Ϲ��ܲ�֧�֣�ֻ�������������˺š�</div>

	<div class=title>3.ΪʲôҪ<b>�˳���¼</b>��</div>

	<div class=value3>��¼������������ϻ᳤�ڴ�����δ�رյ�������У����ǹص�����Ȼ�����ڵ��Ե�Ӳ���У����˳���¼����������������û���Ϣ��</div>

	<div class=title>4.��¼���Ƿ�����´��Զ���¼��</div>

	<div class=value3>��¼�������<b>�˳���¼</b>������û������������Cookie���Ժ�Ͳ���Ҫ�ٴε�¼����Ȼ���Կ���ʹ��<b>�˳���¼</b>���������û���ݵ�¼��</div>

	<div class=title>5.�鿴�ҵ�����</div>

	<div class=value3>�鿴���ʺŵ���Ϣ��<%=DEF_PointsName(0)%>��ע��ʱ��һ�����ϡ�</div>
    
<%End Sub%>