<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Board_Popfun.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<!-- #include file=inc/ForumCategory_fun.asp -->
<%
DEF_BBS_HomeUrl = "../../"
Dim GBL_ID,GBL_ModifyID
initDatabase
GBL_CHK_TempStr = ""
GBL_ID = checkSupervisorPass

Manage_sitehead DEF_SiteNameString & " - ����Ա",""

frame_TopInfo
If GBL_CHK_Flag=1 Then
	Select Case Left(Request("action"),5)
		Case "del"
			DisplayUserNavigate("ɾ����̳����")
			ForumCategoryDelete
		Case "edit"
			DisplayUserNavigate("�޸���̳����")
			ForumCategoryModify
		Case "join"
			DisplayUserNavigate("�����̳����")
			ForumCategoryJoin
		Case Else
			DisplayUserNavigate("��̳�������")
			ForumCategoryManage
	End Select
Else
DisplayLoginForm
End If
frame_BottomInfo
closeDataBase
Manage_Sitebottom("none")

Sub ForumCategoryManage

	GBL_CHK_TempStr = ""
	Dim Rs
	Set Rs = LDExeCute("Select AssortID,AssortName,AssortMaster from LeadBBS_Assort order by AssortID",0)
	If Rs.Eof Then
		Response.Write "��û���κη��࣬�����Ӱ�!"
		Rs.Close
		Set Rs = Nothing
		Exit Sub
	Else
		GBL_GetData = Rs.GetRows(-1)
		Rs.Close
		Set Rs = Nothing
	End If
	If GBL_CHK_TempStr<>"" then
		Response.Write GBL_CHK_TempStr
	Else
%>
<div class=frameline><a href=ForumCategoryManage.asp?action=join>������ӷ���</a>
</div>
<div class=frameline>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" class=frame_table>
        <tr bgcolor="#eeeeee" class=frame_tbhead>
          <td width="10%"><div class=value>ID</div></td>
          <td width="45%"><div class=value>��������</div></td>
          <td width="45%"><div class=value><%=DEF_PointsName(7)%></div></td>
        </tr>
	<%
	Dim CountN,TempN
	CountN = Ubound(GBL_GetData,2)
	for TempN=0 to CountN
		Response.Write "        <tr>" & VbCrLf
		Response.Write "          <td class=tdbox>"
		Response.Write GBL_GetData(0,TempN) & "</td>" & VbCrLf
		Response.Write "          <td class=tdbox>" & GBL_GetData(1,TempN) & " <a href=ForumCategoryManage.asp?action=edit&GBL_MODIFYID=" & GBL_GetData(0,TempN) & ">�޸�</a> <a href=ForumCategoryManage.asp?action=del&GBL_DELETEID=" & GBL_GetData(0,TempN) & ">ɾ��</a></td>" & VbCrLf
		Response.Write "          <td class=tdbox>"
		DisplayBoardMastList GBL_GetData(2,TempN),30
		Response.Write "</td>" & VbCrLf
		Response.Write "        </tr>" & VbCrLf
	next
	%>
      </table>
</div>
<%
	End If

End Sub


Sub DisplayBoardMastList(MasterList,Num)

	If MasterList = "" Then
		Response.Write "��"
		Exit Sub
	ElseIf MasterList = "?LeadBBS?" Then
		Response.Write "ȫ�����"
		Exit Sub		
	End If
	Dim temp,n,I
	Temp = split(MasterList,",")
	I = Ubound(temp,1)
	For N = 0 to I
		If N >= Num Then
			'Response.Write "..."
			Exit For
		End If
		Response.Write " <a href=" & DEF_BBS_HomeUrl & "user/LookUserInfo.asp?name=" & urlencode(temp(N)) & ">" & htmlencode(temp(n)) & "</a>"
	Next

	If N >= Num and N <= I Then
		Response.Write "<span style=""cursor:hand"" title=""�������: " & temp(N)
		N = N + 1
		For N = N to I
			Response.Write " " & temp(N)
		Next
		Response.Write """>...</span>"
	End If

End Sub

Sub ForumCategoryDelete

	Dim GBL_DELETEID
	GBL_DELETEID = Left(Request("GBL_DELETEID"),14)
	If isNumeric(GBL_DELETEID)=0 Then GBL_DELETEID=0
	GBL_DELETEID = cCur(GBL_DELETEID)

	If Request.Form("sure") = "E72ksiOkw2" Then
		If DeleteForumAssort(GBL_DELETEID)>0 Then
			Response.Write "<p><font color=008800 class=greenfont><b>�Ѿ��ɹ�ɾ��IDΪ" & GBL_DELETEID & "����̳���࣡</b></font></p>"
		Else
			Response.Write "<p><font color=ff0000 class=redfont><b>" & GBL_CHK_TempStr & "</b></font></p>"
		End If
	Else
		%><p><form action=ForumCategoryManage.asp method=post>
		<b><font color=ff0000 class=redfont>ȷ����Ϣ�� ���Ҫɾ������̳������<br><br>
		
		<input type=hidden name=GBL_DELETEID value="<%=urlencode(GBL_DELETEID)%>">
		<input type=hidden name=sure value="E72ksiOkw2">
		<input type=hidden name=action value="del">
			
		<input type=button value=����ɾ�� onclick="javascript:history.go(-1);" class=fmbtn>
		<input type=submit value=��Ȼɾ�� class=fmbtn>
		</form>
	<%End If

End Sub

Function ForumCategoryJoin

	%>
	<div class=frameline><b>��ӷ���</b></div>
	<div class=frameline><%
	GBL_CHK_TempStr = ""
	If Request.Form("submitflag")="LKOkxk2" Then
		GBL_AssortID = Left(Trim(Request.Form("Form_AssortID")),14)
		GBL_AssortName = Trim(Request.Form("Form_AssortName"))
		GBL_AssortMaster = Trim(Request.Form("GBL_AssortMaster"))
		If CheckFormForumCateGoryData=0 Then
			Response.Write "<div class=alert>���ݲ���ͨ����" & GBL_CHK_TempStr & "</div>" & VbCrLf
			DisplayJoinForm
			          		Else
			If InsertForumAssort = 0 Then
				Response.Write "<div class=alert>�������" & GBL_CHK_TempStr & "</div>" & VbCrLf
				DisplayJoinForm
			Else
				Response.Write "<div class=alert><span class=greenfont><b>��ӳɹ�!</b></span></div>" & VbCrLf
			End If
		End If
	Else
		DisplayJoinForm
	End If%>
	</div><%

End Function

Function DisplayJoinForm%>

	<table class=frame_table><form action=ForumCategoryManage.asp method=post name=form1 id=form1>
	<tr><td class=tdbox width=120>Ԥ������ID��:</td><td class=tdbox><input name=Form_AssortID value="<%=htmlencode(GBL_AssortID)%>" class=fminpt></td></tr>
	<tr><td class=tdbox><input name=submitflag type=hidden value="LKOkxk2">
	<input name=action type=hidden value="join">
	Ԥ����������:</td><td class=tdbox><input name=Form_AssortName value="<%=htmlencode(GBL_AssortName)%>" class=fminpt></td></tr>
	<tr><td class=tdbox>�����������:</td><td class=tdbox><input name=GBL_AssortMaster value="<%=htmlencode(GBL_AssortMaster)%>" class=fminpt>(���ŷָ�,ȫ�������д<span style="cursor:hand" onclick="document.form1.GBL_AssortMaster.value='?LeadBBS?';">?LeadBBS?</span>)</td></tr>
	<tr><td class=tdbox>&nbsp;</td><td class=tdbox><input type=submit value="�ύ" class=fmbtn> <input type=reset value="ȡ��" class=fmbtn></form></td></tr>
	</table>

<%End Function

Function ForumCategoryModify

	%>
	<div class=frameline><b>�޸ķ���</b></div>
	<div class=frameline>
			<%
	GBL_ModifyID = Left(Request("GBL_ModifyID"),14)
	If isNumeric(GBL_ModifyID)=0 Then GBL_ModifyID=0
	GBL_ModifyID = cCur(GBL_ModifyID)
	If GetForumAssortData(GBL_MODIFYID) <> 0 Then
		GBL_AssortID = GBL_GetData(0,0)
		GBL_AssortName = GBL_GetData(1,0)
		GBL_AssortMaster = GBL_GetData(2,0)
		GBL_AssortLimit = GBL_GetData(3,0)
		If isNull(GBL_AssortLimit) Then GBL_AssortLimit = 0
		GBL_CHK_TempStr = ""
		If Request.Form("submitflag")="LKOkxk2" Then
			GBL_AssortID = Trim(Request.Form("Form_AssortID"))
			GBL_AssortName = Trim(Request.Form("Form_AssortName"))
			GBL_AssortMaster = Trim(Request.Form("GBL_AssortMaster"))

			Dim Temp1,TempN,Temp2
			GBL_AssortLimit = 0
			Temp2 = 1
			For TempN = 0 to LimitAssortStringDataNum
				Temp1 = Request.Form("Limit" & TempN+1)
				If Temp1 <> "1" Then Temp1 = "0"
				If Temp1 = "1" Then GBL_AssortLimit = GBL_AssortLimit+cCur(Temp2)
				Temp2 = Temp2*2
			Next
			If CheckFormForumCateGoryData=0 Then
				Response.Write "<div class=alert>���ݲ���ͨ����" & GBL_CHK_TempStr & "</div>" & VbCrLf
				DisplayModifyForm
			Else
				If UpdateForumAssort = 0 Then
					Response.Write "<div class=alert>�޸ĳ���" & GBL_CHK_TempStr & "</div>" & VbCrLf
					DisplayModifyForm
				Else
					Response.Write "<div class=alert><span class=greenfont><b>�޸ĳɹ�!</b></span></div>" & VbCrLf
					ReloadBoardListData
				End If
			End If
		Else
			DisplayModifyForm
		End If
	Else
		Response.Write "<div class=alert>����δѡ��Ҫ�޸ĵķ��ࡣ</div>" & VbCrLf
	End If%>
	</div>
	<%

End Function

Function DisplayModifyForm

	%>
	<table class=frame_table><form action=ForumCategoryManage.asp method=post>
	<tr><td class=tdbox width=120>Ԥ������ID��:</td><td class=tdbox><input name=Form_AssortID value="<%=htmlencode(GBL_AssortID)%>" class=fminpt></td></tr>
	<tr><td class=tdbox><input name=submitflag type=hidden value="LKOkxk2">
		<input name=action type=hidden value="edit">
		<input name=GBL_ModifyID type=hidden value="<%=htmlencode(GBL_ModifyID)%>" class=fminpt>
		Ԥ����������:</td><td class=tdbox><input name=Form_AssortName value="<%=htmlencode(GBL_AssortName)%>" class=fminpt></td></tr>
	<tr><td class=tdbox>�����������:</td><td class=tdbox><input name=GBL_AssortMaster value="<%=htmlencode(GBL_AssortMaster)%>" class=fminpt>(���ŷָ�,ȫ�������д<span style="cursor:hand" onclick="document.form1.GBL_AssortMaster.value='?LeadBBS?';">?LeadBBS?</span>)
	</td></tr>
	<tr>
		<td class=tdbox>
			���ඨ�ƣ�</td>
		<td class=tdbox><%
		Dim TempN
		GBL_AssortLimit = cCur(GBL_AssortLimit)
		For TempN = 0 to LimitAssortStringDataNum%>
			<input type="checkbox" class=fmchkbox name="Limit<%=TempN+1%>" value="1"<%If GetBinarybit(GBL_AssortLimit,TempN+1) = 1 Then
			Response.Write " checked>"
		Else
			Response.Write ">"
		End If%><%=LimitAssortStringData(tempN)%><br>
		<%Next%></td>
	</tr>
	<tr><td class=tdbox>&nbsp;</td><td class=tdbox>
		<input type=submit value="�ύ" class=fmbtn> <input type=reset value="ȡ��" class=fmbtn></td></tr></form>
	</table>

<%End Function%>