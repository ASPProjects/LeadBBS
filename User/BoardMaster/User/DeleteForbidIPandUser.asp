<!-- #include file=../../../inc/BBSsetup.asp -->
<!-- #include file=../../../inc/Board_popfun.asp -->
<!-- #include file=../../../inc/Limit_Fun.asp -->
<!-- #include file=../inc/BoardMaster_Fun.asp -->
<%
DEF_BBS_HomeUrl = "../../../"
initDatabase
GBL_CHK_TempStr = ""
CheckisBoardMasterFlag

SiteHead(DEF_SiteNameString & " - " & DEF_PointsName(6) & "����")


UserTopicTopInfo
DisplayUserNavigate("ɾ���û�")
            
If GBL_CHK_Flag=1 and BDM_isBoardMasterFlag = 1 and BDM_SpecialPopedomFlag = 1 Then
	If GBL_CHK_TempStr="" Then
		If Request.Form("DeleteSure")="E72ksiOkw2" Then
			If DeleteForbidIPandUser = 1 Then
				Response.Write "<p><font color=008800 class=greenfont><b>�Ѿ��ɹ�������е��ڵ������û������εģɣе�ַ��</b></font></p>"
			else
				Response.Write "<p><font color=ff0000 class=redfont><b>" & GBL_CHK_TempStr & "</b></font></p>"
			End If
		Else
			%>
			<p><form action=DeleteForbidIPandUser.asp method=post>
			<b><font color=ff0000 class=redfont>ȷ����Ϣ��������<%=year(DEF_Now)%>��<%=month(DEF_Now)%>��<%=day(DEF_Now)%>���˶������������ǰ�����ڵ����ݣ��������£�<br>
			&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 1.��������ε�IP��ַ<br>
			&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 2.��������η������ݵĻ�Ա<br>
			&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 3.��������ԵĻ�Ա<br>
			&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 4.�������ֹ�޸ĵĻ�Ա<br>
			&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 5.�ָ������˵�<%=DEF_PointsName(5)%>����ͨ��Ա״̬<br>
			&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 6.����ڵ���ʱ����ǰ��Ȼδ�����ע���Ա<br>
			<br>
			<input type=hidden name=DeleteSure value="E72ksiOkw2">
			
			<input type=submit value=��Ȼ��� class=fmbtn>
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
Else%>
	<table width=96%>
	<tr>
	<td>
	<%
	If Request("submitflag")="" Then
		Response.Write "<br><b>���ȵ�¼</b>"
	Else
		Response.Write "<br><p align=left><font color=ff0000 class=redfont><b>" & GBL_CHK_TempStr & "</b></font>"
	End If
	DisplayLoginForm
	Response.Write "</p>"%>
	</td>
	</tr>
	</table>
<%End If
UserTopicBottomInfo
closeDataBase
SiteBottom
If GBL_ShowBottomSure = 1 Then Response.Write GBL_SiteBottomString

Rem ���ĳ�û����Ƿ����
Function DeleteForbidIPandUser

	Server.ScriptTimeOut = 6000
	'If UserName <> "" and inStr(UserName,",") = 0 and inStr(Lcase(DEF_SupervisorUserName),"," & Lcase(UserName) & ",") > 0 Then
	'	GBL_CHK_TempStr = "�����û���" & htmlencode(UserName) & "�����ڣ�"
	'	DeleteForbidIPandUser = 0
	'	Exit Function
	'End If
	
	Response.Write "<br><p>���ڸ����У�����<p>"
	Dim ExpiresTime
	ExpiresTime = GetTimeValue(year(DEF_Now) & "-" & Month(DEF_Now) & "-" & Day(DEF_Now))
	Dim Rs
	Set Rs = LDExeCute("Select T2.ID,T2.UserLimit,T2.UserName,T1.Assort from LeadBBS_SpecialUser as T1 Left join LeadBBS_User As T2 on T1.UserID=T2.ID where T1.ExpiresTime>0 and T1.ExpiresTime<" & ExpiresTime,0)
	If Rs.Eof Then
		DeleteForbidIPandUser = 1
		Response.Write "<br>���κε��ڵ������û�������Ҫ���£���"
	End If
	Dim GBL_UserName_UserID,GBL_UserName_UserLimit,GBL_UserName,GBL_Assort
	Do while Not Rs.Eof
		GBL_UserName_UserLimit = cCur(Rs(1))
		GBL_UserName_UserID = cCur(Rs(0))
		GBL_UserName = Rs(2)
		GBL_Assort = cCur(Rs(3))
		
		',0-��֤��Ա,1-����,2-�ܰ���,3-���λ�Ա,4-���Ի�Ա,5-���޸Ļ�Ա,6-����ʽ��Ա
		Select Case GBL_Assort
			Case 0:
					If GetBinarybit(GBL_UserName_UserLimit,2) = 1 Then
						Response.Write "<br>�û�" & htmlencode(GBL_UserName) & "�Ѿ����" & DEF_PointsName(5) & "״̬��"
						GBL_UserName_UserLimit = SetBinaryBit(GBL_UserName_UserLimit,2,0)
						CALL LDExeCute("Update LeadBBS_User Set UserLimit=" & GBL_UserName_UserLimit & " where ID=" & GBL_UserName_UserID,1)
						CALL LDExeCute("Delete from LeadBBS_SpecialUser Where Assort=" & GBL_Assort & " and UserID=" & GBL_UserName_UserID,1)
					End If
			Case 3:
					If GetBinarybit(GBL_UserName_UserLimit,7) = 1 Then
						Response.Write "<br>�û�" & htmlencode(GBL_UserName) & "�Ѿ�������η������ݼ�ǩ����"
						GBL_UserName_UserLimit = SetBinaryBit(GBL_UserName_UserLimit,7,0)
						CALL LDExeCute("Update LeadBBS_User Set UserLimit=" & GBL_UserName_UserLimit & " where ID=" & GBL_UserName_UserID,1)
						CALL LDExeCute("Delete from LeadBBS_SpecialUser Where Assort=" & GBL_Assort & " and UserID=" & GBL_UserName_UserID,1)
					End If
			Case 4:
					If GetBinarybit(GBL_UserName_UserLimit,3) = 1 Then
						Response.Write "<br>�û�" & htmlencode(GBL_UserName) & "�Ѿ�������Լ����Ͷ���Ϣ��"
						GBL_UserName_UserLimit = SetBinaryBit(GBL_UserName_UserLimit,3,0)
						CALL LDExeCute("Update LeadBBS_User Set UserLimit=" & GBL_UserName_UserLimit & " where ID=" & GBL_UserName_UserID,1)
						CALL LDExeCute("Delete from LeadBBS_SpecialUser Where Assort=" & GBL_Assort & " and UserID=" & GBL_UserName_UserID,1)
					End If
			Case 5:
					If GetBinarybit(GBL_UserName_UserLimit,4) = 1 Then
						Response.Write "<br>�û�" & htmlencode(GBL_UserName) & "�Ѿ������ֹ�޸����Ӽ��������ϣ�"
						GBL_UserName_UserLimit = SetBinaryBit(GBL_UserName_UserLimit,4,0)
						CALL LDExeCute("Update LeadBBS_User Set UserLimit=" & GBL_UserName_UserLimit & " where ID=" & GBL_UserName_UserID,1)
						CALL LDExeCute("Delete from LeadBBS_SpecialUser Where Assort=" & GBL_Assort & " and UserID=" & GBL_UserName_UserID,1)
					End If
			Case 6:
					If GetBinarybit(GBL_UserName_UserLimit,1) = 1 Then
						Response.Write "<br>δ�����û�" & htmlencode(GBL_UserName) & "�Ѿ����ɹ�ɾ����"
						GBL_UserName_UserLimit = SetBinaryBit(GBL_UserName_UserLimit,1,0)
						CALL LDExeCute("delete from LeadBBS_User where ID=" & GBL_UserName_UserID,1)
						CALL LDExeCute("Update LeadBBS_SiteInfo Set UserCount=UserCount-1",1)
						UpdateStatisticDataInfo -1,1,1
						CALL LDExeCute("Delete from LeadBBS_SpecialUser Where Assort=" & GBL_Assort & " and UserID=" & GBL_UserName_UserID,1)
					End If
			Case Else:
		End Select
		Rs.MoveNext
	Loop
	Rs.Close
	Set Rs = Nothing
	Response.Write "<br><font color=Green Class=greenfont>���������û�������ɣ�</font>"
	Set Rs = LDExeCute("Delete From LeadBBS_ForbidIP where ExpiresTime>0 and ExpiresTime<" & ExpiresTime,0)
	Response.Write "<br><font color=Green Class=greenfont>�������ڵı����Σɣе�ַ�Ѿ��ɹ���ɣ�</font>"
	DeleteForbidIPandUser = 1

End Function%>