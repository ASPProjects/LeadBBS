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

Dim GBL_CTG_DELETEID
GBL_CTG_DELETEID = Left(Request("GBL_CTG_DELETEID"),14)
If isNumeric(GBL_CTG_DELETEID) = 0 Then GBL_CTG_DELETEID = 0
GBL_CTG_DELETEID = cCur(GBL_CTG_DELETEID)
If GBL_CTG_DELETEID < 0 Then GBL_CTG_DELETEID = 0
GBL_CHK_TempStr=""
If GBL_CTG_DELETEID=0 Then
	GBL_CHK_TempStr = GBL_CHK_TempStr & "û��ѡ��Ҫɾ�����û�<br>" & VbCrLf
End If

frame_TopInfo
DisplayUserNavigate("ɾ���û�")
If GBL_CHK_Flag=1 Then
	If GBL_CHK_TempStr="" Then
		If Request.Form("DeleteSuer")="E72ksiOkw2" Then
			If DeleteUser(GBL_CTG_DELETEID)>0 Then
				Response.Write "<br><p><font color=008800 class=greenfont><b>�Ѿ��ɹ�ɾ��IDΪ" & GBL_CTG_DELETEID & "���û���</b></font></p>"
				CALL LDExeCute("Update LeadBBS_SiteInfo Set UserCount=UserCount-1",1)
				UpdateStatisticDataInfo -1,1,1
			else
				Response.Write "<br><p><font color=ff0000 class=redfont><b>" & GBL_CHK_TempStr & "</b></font></p>"
			End If
		Else
			%>
			<p><form action=UserDelete.asp method=post>
			<b><font color=ff0000 class=redfont>ȷ����Ϣ�� ���Ҫɾ�����û���������û���<%=DEF_PointsName(8)%>,<br>
			�뵽��Ӧ����ɾ�����û�����Ȩ��.<br>
			ɾ���û��󣬲���ɾ�����û���������ӣ������ӽ���Ϊ�οͷ���״̬��<br><br>
			<input type=hidden name=GBL_CTG_DELETEID value="<%=urlencode(GBL_CTG_DELETEID)%>">
			<input type=hidden name=DeleteSuer value="E72ksiOkw2">
			
			<input type=button value=��ɾ�� onclick="javascript:history.go(-1);" class=fmbtn>
			<input type=submit value=��Ȼɾ�� class=fmbtn>
			
			</form>
		<%End If
	Else%>
		<table width=96%>
		<tr>
			<td>
				<%Response.Write GBL_CHK_TempStr%>
			</td>
		</tr>
		</table>
	<%End If
Else
DisplayLoginForm
End If
frame_BottomInfo
closeDataBase
Manage_Sitebottom("none")

rem ɾ��ĳ�û�
Function DeleteUser(ID)

	Dim Rs
	Set Rs = LDExeCute(sql_select("Select ID,UserName from LeadBBS_User where ID=" & ID,1),0)
	If Rs.Eof Then
		DeleteUser = 0
		Rs.Close
		Set Rs = Nothing
		GBL_CHK_TempStr = GBL_CHK_TempStr & "�Ҳ������û���<br>" & VbCrLf
	Else
		GBL_CHK_User = Rs("UserName")
		If CheckSupervisorUserName = 1 Then
			GBL_CHK_TempStr = GBL_CHK_TempStr & "��������Ա����ɾ����<br>" & VbCrLf
			Rs.Close
			Set Rs = Nothing
			DeleteUser = 0
			Exit Function
		End If
		Rs.Close
		Set Rs = Nothing
		CALL LDExeCute("delete from LeadBBS_SpecialUser where UserID=" & ID,1)
		CALL LDExeCute("delete from LeadBBS_User where ID=" & ID,1)
		DeleteUser = 1
	End if

End Function%>