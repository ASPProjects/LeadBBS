<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Board_popfun.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<%
DEF_BBS_HomeUrl = "../../"
Dim GBL_ID
initDatabase
GBL_CHK_TempStr = ""
GBL_ID = checkSupervisorPass
Const LMT_MaxListLogNum = 300

Manage_sitehead DEF_SiteNameString & " - ����Ա",""

frame_TopInfo
DisplayUserNavigate("��̳��־")
If GBL_CHK_Flag=1 Then
	If Request("clear") = "yes" Then
		ClearForumLog
	Else
		DisplayForumLog
	End If
Else
DisplayLoginForm
End If
frame_BottomInfo
closeDataBase
Manage_Sitebottom("none")

Function DisplayForumLog

	'0-ϵͳ��־
	'1-������־
	'51-����Ա��¼��־
	'52-�ܰ�����¼��־
	'53-��ͨ������¼��־
	'54-��ͨ��Ա��¼��־
	'101-����ɾ��������־
	'102-��������������־
	'103-����ת��������־
	'104-�����޸�������־
	'105-����������Ա��־
	'106-�����̶�������־
	'151-�ܰ�������û���־
	'152-�ܰ������IP��־
	'152-�ܰ���ǿ���޸��û�������־
	'153-�ܰ���ǿ���޸��û�������־
	'154-�ܰ����̶ܹ�������־
	'201-���������̶�������־
	%>
	<script language=javascript>
	var lastID=0,Count=0;
	function s(ID,LogType,LogTime,LogInfo,UserName,IP,BoardID)
	{
		if(ID=="")return;
		Count +=1;lastID=ID;
		if(BoardID==0){BoardID="";}else{BoardID="����:" + BoardID;}
		LogTime = LogTime.substr(0,4) + "-" + LogTime.substr(4,2) + "-" + LogTime.substr(6,2) + " " + LogTime.substr(8,2) + ":" + LogTime.substr(10,2) + ":" + LogTime.substr(12,2)
		switch(parseInt(LogType))
		{
			case 0: LogType="<span class=greenfont>ϵͳ��־</span>";break;
			case 1: LogType="������־";break;
			case 9: LogType="��̳��̬";break;
			case 51: LogType="����Ա��¼";break;
			case 52: LogType="<%=DEF_PointsName(6)%>��¼";break;
			case 53: LogType="������¼";break;
			case 54: LogType="��ͨ��Ա��¼";break;
			case 101: LogType="����ɾ������";break;
			case 102: LogType="������������";break;
			case 103: LogType="����ת������";break;
			case 104: LogType="�����޸�����";break;
			case 105: LogType="����������Ա";break;
			case 106: LogType="�����̶�����";break;
			case 151: LogType="<%=DEF_PointsName(6)%>����û�";break;
			case 152: LogType="<%=DEF_PointsName(6)%>���IP";break;
			case 153: LogType="<%=DEF_PointsName(6)%>ǿ���޸��û�����";break;
			case 154: LogType="<%=DEF_PointsName(6)%>�̶ܹ�����";break;
			case 201: LogType="<%=DEF_PointsName(7)%>���̶�����";break;
		}
		document.write("<tr><td class=tdbox>" + ID + "<br>" + BoardID + "</td><td class=tdbox>" + LogType + "<br><a href=\"<%=DEF_BBS_HomeUrl%>User/LookUserInfo.asp?name=" + escape(UserName) + "\" target=_blank>" + UserName + "</a></td><td class=tdbox>" + IP + "<br>" + LogTime+"</td><td class=tdbox>" + LogInfo + "</tr>");
	}
	</script>
				<div class=frameline>
				<b><span class=grayfont>������־(<%=LMT_MaxListLogNum%>��)</span> <a HREF=ForumLog.asp?clear=yes>���2��ǰ����̳��־(���ٱ��������300����־)</a></b></div>
				<table border=0 cellpadding=0 cellspacing=0 width=100% class=frame_table>
				<tbody>
				<tr class=frame_tbhead>
					<td width=74><div class=value>���|����</div></td>
					<td width=90><div class=value>����/�û�</div></td>
					<td width=176><div class=value>IP��ַ/ʱ��</div></td>
					<td><div class=value>��־��Ϣ</div></td>
				</tr>
				<%
	Dim FirstID
	FirstID = Left(Request("ID"),14)
	If isNumeric(FirstID) = 0 Then FirstID = 0
	FirstID = cCur(Fix(FirstID))

	Dim Rs,SQL
	If FirstID = 0 Then
		SQL = sql_select("select ID,LogType,LogTime,LogInfo,UserName,IP,BoardID from LeadBBS_Log Order by id DESC",LMT_MaxListLogNum)
	Else
		SQL = sql_select("select ID,LogType,LogTime,LogInfo,UserName,IP,BoardID from LeadBBS_Log where ID<" & FirstID & " Order by id DESC",LMT_MaxListLogNum)
	End If

	OpenDatabase
	Set Rs = LDExeCute(SQL,0)
	Dim Num
	Dim GetData
	If Not rs.Eof Then
		Response.Write "<script language=javascript>" & VbCrLf & "s("""
		Response.Write Rs.GetString(,,""",""",""");" & VbCrLf & "s(""","")
		%>","","","");
		</script>
		<%
	Else
		Num = -1
	End If
	Rs.close
	Set Rs = Nothing
	closeDataBase%>
				</table>
				<%If FirstID > 0 Then Response.Write "<a href=ForumLog.asp>������ҳ</a> "%>
		<script language=javascript>
			if(Count>=<%=LMT_MaxListLogNum%>)document.write("<a href=ForumLog.asp?id=" + lastID + ">��һҳ</a>");
		</script>
	<%

End Function

Sub ClearForumLog

	If Request.Form("submitflag") = "yes" then
		Dim SQL,Rs,FilterTime,LogData
		FilterTime = GetTimeValue(DateAdd("d", -2, DEF_Now))
		SQL = sql_select("Select ID,LogTime from LeadBBS_Log order by ID DESC",300)
		Set Rs = LDExeCute(SQL,0)
		If Not Rs.Eof Then
			LogData = Rs.GetRows(-1)
			Rs.Close
			Set Rs = Nothing
			SQL = cCur(LogData(1,Ubound(LogData,2)))
			If SQL < FilterTime Then FilterTime = SQL
		Else
			Rs.Close
			Set Rs = Nothing
		End If
		Response.Write "<p>���������ִ���������...<p>"
		SQL = "Delete from LeadBBS_Log where LogTime<" & FilterTime
		Response.Write "<p>" & SQL
		Con.CommandTimeout = 120
		CALL LDExeCute(SQL,1)
		Response.Write "<p>ִ����ϣ��ɹ��������ǰ����̳��־(���ٱ��������300����־)��"
		Response.Write "<p><a href=ForumLog.asp>������ﷵ�ز鿴��־��</a>"
	Else
			%><p><br>
				ע�⣺�˹��ܽ�������¹��ܣ�<br><br>
				&nbsp; &nbsp; &nbsp; 1.�����̳����֮ǰ����̳��־������󽫲��ָܻ���־��<br>
				<br>
				<b><font color=ff0000 class=redfont>ȷ����Ϣ�� ���Ҫ��ʼ�������ô��</font></b><br><br>
				<form action=ForumLog.asp method=post name=LeadBBSFm id=LeadBBSFm>
				<input name=submitflag value=yes type=hidden>
				<input name=clear value=yes type=hidden>
				<input type=button value="�����ʼ���" onclick="javascript:LeadBBSFm.submit();this.disabled=true;" class=fmbtn>
				</form>
			<%
	End If

End Sub
%>