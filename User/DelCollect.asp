<!-- #include file=../inc/BBSsetup.asp -->
<!-- #include file=../inc/Board_Popfun.asp -->
<%
DEF_BBS_HomeUrl = "../"
InitDatabase
UpdateOnlineUserAtInfo GBL_board_ID,"ɾ���ռ���Ϣ"
Dim GBL_ID,Form_ID

Dim DelID
DelID = Left(Request("DelID"),14)
If isNumeric(DelID) = 0 or DelID = "" or InStr(DelID,",") > 0 Then DelID = 0
DelID = Fix(cCur(DelID))

If DelID < 0 Then DelID = 0

GBL_CHK_TempStr=""
GBL_ID = GBL_UserID
Form_ID = GBL_ID
Form_ID = cCur(Form_ID)
If Form_ID < 0 Then Form_ID = 0
If Form_ID=0 Then
	GBL_CHK_TempStr = GBL_CHK_TempStr & "��û�е�¼<br>" & VbCrLf
End If

siteHead("   ɾ���ղ�����")%>
<script language=javascript>
	window.moveTo(window.screen.width/2-225,window.screen.height/2-18);
</script>
<table align=center border=0 cellpadding=0 cellspacing=0>
<tbody>
<tr> 
	<td height=21 width="650">
	<%
	Dim Rs,SQL,SQLendString,ClearFlag
	If GBL_CHK_TempStr = "" Then
		ClearFlag = Request("ClearFlag")
		If ClearFlag = "dkeJje5" or ClearFlag = "dkeJje6" Then
			If GBL_UserID<1 or CheckSupervisorUserName = 0 or ClearFlag = "dkeJje5" Then
				SQL = "delete from LeadBBS_CollectAnc where UserID=" & GBL_UserID
			Else
				SQL = "delete from LeadBBS_CollectAnc"
			End If

			If Request.Form("DeleteSureFlag")="dk9@dl9s92lw_SWxl" Then
				%>
				�ɹ�ɾ�������ղ�����!
				<%
				CALL LDExeCute(SQL,1)
			Else
				%>
				<form name=DellClientForm action=DelCollect.asp method=post>
					<input type=hidden name=DeleteSureFlag value="dk9@dl9s92lw_SWxl">
					<input type=hidden name=ClearFlag value="<%=htmlencode(ClearFlag)%>">
					<input type=hidden name=DelID value="<%=htmlencode(DelID)%>">
					<b><%If GBL_UserID<1 or CheckSupervisorUserName = 0 or ClearFlag = "dkeJje5" Then%>ȷ��Ҫɾ�����������ղ�������ɾ���󽫲��ָܻ���<%Else
					%><font color=Red class=redfont>���ǹ���Ա��ȷ��Ҫɾ�������˵��ղ�������ɾ�����޷��ָ���</font><%End If%></b>
					<p><input type=submit value=ȷ�� class=fmbtn>
					<input type=button value=��ɾ onclick="javascript:window.close();" class=fmbtn>
				</form>
				<%
			End If
		Else
			If GBL_UserID<1 or CheckSupervisorUserName = 0 Then
				SQL = sql_select("Select * from LeadBBS_CollectAnc where UserID=" & GBL_UserID & " and id=" & DelID,1)
			Else
				SQL = sql_select("Select * from LeadBBS_CollectAnc where id=" & DelID,1)
			End If
			Set Rs = LDExeCute(SQL,0)
			If Rs.Eof Then
				Rs.Close
				Set Rs = Nothing
				Response.Write "�Ҳ�����¼��<br>" & VbCrLf
			Else
				Rs.Close
				Set Rs = Nothing
				If Request.Form("DeleteSureFlag")="dk9@dl9s92lw_SWxl" Then
					%>
					�ɹ�ɾ�����Ϊ<font color=ff0000 class=redfont><%=DelID%></font>���ղ�����!
					<%
					CALL LDExeCute("Delete from LeadBBS_CollectAnc where id=" & DelID,1)
				Else
					%>
					<form name=DellClientForm action=DelCollect.asp method=post>
						<input type=hidden name=DeleteSureFlag value="dk9@dl9s92lw_SWxl">
						<input type=hidden name=ClearFlag value="<%=htmlencode(ClearFlag)%>">
						<input type=hidden name=DelID value="<%=htmlencode(DelID)%>">
						<b>ȷ��Ҫɾ�����Ϊ<font color=ff0000 class=redfont><%=DelID%></font>���ղ�������</b>
						<p><input type=submit value=ȷ�� class=fmbtn>
						<input type=button value=��ɾ onclick="javascript:window.close();" class=fmbtn>
					</form>
					<%
				End If
			End If
		End If
	Else%>
		<table width=96%>
		<tr>
		<td>
		<%Response.Write "<p align=left><font color=ff0000 class=redfont><b>" & GBL_CHK_TempStr & "</b>"
		Response.Write "</p>"%>
		</td>
		</tr>
		</table>
	<%End If%></td>
</tr>

</table>
<%

closeDataBase
%>