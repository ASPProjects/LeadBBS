<!-- #include file=../../inc/BBSSetup.asp -->
<!-- #include file=../../inc/Board_Popfun.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<%
DEF_BBS_HomeUrl = "../../"
Dim GBL_ID
initDatabase
GBL_CHK_TempStr = ""
GBL_ID = checkSupervisorPass
closeDataBase

Manage_sitehead DEF_SiteNameString & " - ����Ա",""
frame_TopInfo
If GBL_CHK_Flag=1 Then
	select Case request("action")
		Case "blockdelete":
			DisplayUserNavigate("����ɾ����̳����")
			BlockDelete
		case else
			DisplayUserNavigate("����������̳����")
			BlockUpdate
	end select
Else
DisplayLoginForm
End If
frame_BottomInfo
Manage_Sitebottom("none")

sub BlockUpdate%>
<script language=javascript>
	function blockupdate(url,str)
	{
	   if (confirm(str))
	   {
		document.location.href=url;
	   }
	}
</script>
<table width="97%" border="0" cellspacing="0" cellpadding="0">
<tr>
	<td>
		<div class=alert>���棺</div>
		<div class=frameline>�����޸���������������¶�����Ҫ���С������̳���ݽ϶ִ࣬��ʱ�佫��ǳ��������������й��ܽ�����Ӱ����������ܣ��˷������ϵ�������վ��ִ���ڼ佫�ܵ����ظ��š�
			<br>
			�������¼��޸���̳����(�����ʱ�������������κ�һ����ϸȷ��)
		</div>
		<div class=frametitle>
			1.<a href="javascript:blockupdate('UpdateUnderWritePrintColumn.asp?flag=1','�˲�������ʱ����ܻ�ǳ����ã���������ͣ��̳�ٽ��и��£�\n��ȷ��Ҫ �������������������� ��');"><b>��������������������</b></a>
		</div>
		<div class=frameline>�˹��ܺ��������滻���������еĲ������ݣ���������<br>�滻http://w.leadbbs.com/ Ϊ http://www.leadbbs.com/�����ã�����
		</div>
		
		<div class=frametitle>
			2.<a href="javascript:blockupdate('UpdateUnderWritePrintColumn.asp?flag=UpdateRootMaxMinAnnounceID','�˲�������ʱ����ܻ�ǳ����ã���������ͣ��̳�ٽ��и��£�\n��ȷ��Ҫ �޸���LeadBBS_Announce���������� ��');"><b>�޸���LeadBBS_Announce����������</b></a>
		</div>
		<div class=frameline>�����п��ܻᷢ����ҳ���󣬵��鿴��������ʱ������ҳ��ҳת�ƴ���ʱ���������д˳��������̳��һ�еĴ������ִ�д˳����ʱ�������ʱ�ر���̳���У��Ա�֤���٣������ĸ�����ϡ�Ĭ���ִ��ʱ��99999�룬��ֱ��ȫ�����ݸ�����ϡ�
		</div>

		<div class=frametitle>
			3.<a href="javascript:blockupdate('UpdateUnderWritePrintColumn.asp?flag=UpdateUserAnnounce','�˲�������ʱ����ܻ�ǳ����ã���������ͣ��̳�ٽ��и��£�\n��ȷ��Ҫ ����ͳ�������û��������� ��');"><b>����ͳ�������û���������(���ؼ�<%=DEF_PointsName(0)%>)</b></a>
			<a href="javascript:blockupdate('UpdateUnderWritePrintColumn.asp?flag=UpdateUserAnnounce&ReCount=1','�˲�������ʱ����ܻ�ǳ����ã���������ͣ��̳�ٽ��и��£�\n��ȷ��Ҫ ����ͳ�������û��������� ��');"><b>����ͳ�������û���������(�ؼ���<%=DEF_PointsName(0)%>)</b></a>
		</div>
		<div class=frameline>
		����ͳ�������û��������������������������������������������¼���<%=DEF_PointsName(3)%>��ִ�д˳����ʱ�������ʱ�ر���̳���У��Ա�֤���٣������ĸ�����ϡ�Ĭ���ִ��ʱ��99999�룬��ֱ��ȫ�����ݸ�����ϡ�
		</div>

		<div class=frametitle>
			4.<a href="javascript:blockupdate('UpdateUnderWritePrintColumn.asp?flag=UpdateRootMaxMinAnnounceID&BlockType=3','�˲�������ʱ����ܻ�ǳ����ã���������ͣ��̳�ٽ��и��£�\n��ȷ��Ҫ ���²��������û���ũ������ ��');"><b>���²��������û���ũ������</b></a>
		</div>
		<div class=frameline>
			�û���д�������ǹ������գ��˳���ǿ������һ��ת����һ�㲻�����и��´��ִ�д˳����ʱ�������ʱ�ر���̳���У��Ա�֤���٣������ĸ�����ϡ�Ĭ���ִ��ʱ��99999�룬��ֱ��ȫ�����ݸ�����ϡ�
		</div>
	</td>
</tr>
</table>

<%End sub

sub BlockDelete%>

<script language=javascript>
	function blockupdate(url,str)
	{
	   if (confirm(str))
	   {
			document.location.href=url;
	   }
	}
</script>
<table width="97%" border="0" cellspacing="0" cellpadding="0">
<tr>
	<td>
		<p><b>����ɨ�貢����ɾ����̳����(�����ʱ�������������κ�һ����Ҫ��ȷ��)</b></p>
		<div class=frametitle>
		<a href="javascript:blockupdate('UpdateUnderWritePrintColumn.asp?flag=DeleteBlankUser','�˲�������ʱ����ܻ�ǳ����ã���������ͣ��̳�ٽ��и��£�\n��ȷ��Ҫ ɾ�����κη������ҵ�����ʱ����û� ��');"><b>ɾ�����κη�����һ����ǰע����<%=DEF_PointsName(4)%>����100���û�</b></font></a>
		</div>
		<div class=frameline>
			�˲�����ɾ�����κη�������һ����ǰע�ᣬ������ʱ��С��100���ӵ��û���������ɾ���û���ͬʱ����ͬʱɾ����صĺ������ϣ��ղ����ӣ���̳����Ϣ���ϴ�����������ɾ����Ӧ��ͶƱ���ϡ�ִ�д˳����ʱ�������ʱ�ر���̳���У��Ա�֤���٣������ĸ�����ϡ�Ĭ���ִ��ʱ��99999�룬��ֱ��ȫ�����ݸ�����ϡ�
			<span class=redfont>�����������֧��FSO��������ɾ���ϴ�������</span>
		</div>
		
		<div class=frametitle>
			<a href="DeleteExpiresAnnounceData.asp"><b>����ɾ��ָ����������̳����</b></a>
		</div>
		<div class=frameline>
		��ָ����Ҫɾ������(�����Ⲣɾ����Ӧ�Ļظ�)���ڵİ��棬�����µ��ռ䣮<br>
		ע�⣺����ɾ���������⼰��������Ļظ�����
		</div>
		<div class=frametitle><a href="UpdateUnderWritePrintColumn.asp?flag=DeleteBlankUser&dflag=upload"><b>����ɾ����ָ̳����������ʷ����</b></a>
		</div>
		<div class=frameline>
		ɾ��ָ������֮��������ϴ�����(�������ݿ⼰Ӳ���ļ���ɾ���ļ���ҪFSO֧��)
		</div>
	</td>
</tr>
</table>

<%End sub%>