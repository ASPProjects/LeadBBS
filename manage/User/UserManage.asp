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
DisplayUserNavigate("�û�����")
If GBL_CHK_Flag=1 Then
	UserBrowser
Else
DisplayLoginForm
End If
frame_BottomInfo
closeDataBase
Manage_Sitebottom("none")

Function UserBrowser

	GBL_CHK_TempStr=""
	Dim Rs,SQL
	Dim UpDownPageFlag
	UpDownPageFlag = Request("UpDownPageFlag")

	Dim Start,recordCount,key
	recordCount=0
	
	Dim SQLendString

	Start = Left(Trim(Request("Start")),14)
	Start = start
	If isNumeric(Start)=0 or Start="" Then Start=0
	Start = cCur(Start)
	key = Request.Form("key")
	If key="" Then key = Request("key")

	Dim SQLCountString,whereFlag
	SQLendString=""
	whereFlag = 0

	Rem ����Ĵ���ʹĿǰ�ݲ��ṩ���з���˫�ز�ѯ

	If key<>"" Then
		If whereFlag = 1 Then
			SQLendString = SQLendString & " and LeadBBS_User.UserName='" & Replace(key,"'","''") & "'"
		Else
			SQLendString = SQLendString & " where LeadBBS_User.UserName='" & Replace(key,"'","''") & "'"
			whereFlag = 1
		End If
	End If
	SQLCountString = SQLendString
	Dim MaxRecordID,MinRecordID
	MaxRecordID = 0
	MinRecordID = 0
	Dim FirstID,LastID	
	If key="" Then
		If UpDownPageFlag = "1" and Start>0 then
			If whereFlag = 1 Then
				SQLendString = SQLendString & " and LeadBBS_User.ID<" & Start
			Else
				SQLendString = SQLendString & " where LeadBBS_User.ID<" & Start
				whereFlag = 1
			End If
		Else
			If whereFlag = 1 Then
				SQLendString = SQLendString & " and LeadBBS_User.ID>" & Start
			Else
				SQLendString = SQLendString & " where LeadBBS_User.ID>" & Start
				whereFlag = 1
			End If
		end If
	
		If UpDownPageFlag = "1" then
			'If DEF_IDFocusFlag<> 2 Then SQLendString = SQLendString & " Order by  LeadBBS_User.ID DESC"
			SQLendString = SQLendString & " Order by  LeadBBS_User.ID DESC"
		Else
			'If DEF_IDFocusFlag<> 1 Then SQLendString = SQLendString & " Order by  LeadBBS_User.ID ASC"
			SQLendString = SQLendString & " Order by  LeadBBS_User.ID ASC"
		End If
	
		SQL = "select Max(id) from LeadBBS_User " & SQLCountString
		Set Rs = LDExeCute(SQL,0)
		
		If not Rs.Eof Then
			If Rs(0) <> "" Then
				MaxRecordID = cCur(Rs(0))
			Else
				MaxRecordID = 0
			End If
		End If
		Rs.Close
		Set Rs = Nothing
		SQL = "select Min(id) from LeadBBS_User " & SQLCountString
		Set Rs = LDExeCute(SQL,0)
	
		If not Rs.Eof Then
			If Rs(0) <> "" Then
				MinRecordID = cCur(Rs(0))
			else
				MinRecordID = 0
			end If
		End If
	
		Rs.Close
		Set Rs = Nothing
		SQL = "select UserCount from LeadBBS_SiteInfo"
		Set Rs = LDExeCute(SQL,0)
		If Rs.Eof then
			recordCount=0
		Else
			recordCount=cCur(rs(0))
			If recordCount="" or isNull(recordCount) or len(recordCount)<1 Then recordCount=0
		End If
		Rs.Close
		Set Rs = Nothing
	Else
		
	End If
	
	SQL = sql_select("select LeadBBS_User.ID,LeadBBS_User.UserName,LeadBBS_User.Points,LeadBBS_User.ApplyTime,LeadBBS_User.Prevtime from LeadBBS_User " & SQLendString,DEF_MaxListNum)
	Set Rs = LDExeCute(SQL,0)
	Dim Num
	Dim GetData
	If Not rs.Eof Then
		GetData = Rs.GetRows(-1)
		Num = Ubound(GetData,2)+1
		If Num > Recordcount Then RecordCount = Num
	Else
		Num = 0
	End If
	Rs.close
	Set Rs = Nothing
	
	Dim i,N
	If Num>0 Then
		i=1
	
	Dim MinN,MaxN,StepValue
	SQL = ubound(getdata,2)
	If UpDownPageFlag = "1" then
		MinN = SQL
		MaxN = 0
		StepValue = -1
	Else
		MinN = 0
		MaxN = SQL
		StepValue = 1
	End If
	
	LastID = cCur(GetData(0,MaxN))
	FirstID = cCur(GetData(0,MinN))
	
	
	Dim EndwriteQueryString,PageSplictString
	EndwriteQueryString = "?GBL_CTG_ID=0"
	If key<>"" Then EndwriteQueryString = EndwriteQueryString & "&key=" & urlencode(key)

	PageSplictString = PageSplictString & "<table border=0 cellspacing=0 cellpadding=0><tr><td>&nbsp;"
	if FirstID>MinRecordID and FirstID<>0 then
		PageSplictString = PageSplictString & "<a href=UserManage.asp" & EndwriteQueryString & "&Start=0&SubmitFlag=3829EwoqIaNfoG>��ҳ</a> " & VbCrLf
	else
		PageSplictString = PageSplictString & "<span class=grayfont>��ҳ</span> " & VbCrLf
	end if

	if FirstID > MinRecordID and FirstID<>0 then
		PageSplictString = PageSplictString & " <a href=UserManage.asp" & EndwriteQueryString & "&Start=" & FirstID & "&UpDownPageFlag=1&SubmitFlag=3829EwoqIaNfoG>��ҳ</a> " & VbCrLf
	else
		PageSplictString = PageSplictString & " <span class=grayfont>��ҳ</span> " & VbCrLf
	end if

	if LastID<MaxRecordID and LastID<>0 then
		PageSplictString = PageSplictString & " <a href=UserManage.asp" & EndwriteQueryString & "&Start=" & LastID & "&SubmitFlag=3829EwoqIaNfoG>��ҳ</a> " & VbCrLf
	else
		PageSplictString = PageSplictString & " <span class=grayfont>��ҳ</span> " & VbCrLf
	end if

	if LastID < MaxRecordID and LastID<>0 then
		PageSplictString = PageSplictString & " <a href=UserManage.asp" & EndwriteQueryString & "&Start=" & MaxRecordID+1 & "&UpDownPageFlag=1&SubmitFlag=3829EwoqIaNfoG>βҳ</a> " & VbCrLf
	else
		PageSplictString = PageSplictString & " <span class=grayfont>βҳ</span> " & VbCrLf
	end if
	PageSplictString = PageSplictString & "��<b>" & recordCount & "</b>����Ϣ"
	If (recordCount mod DEF_MaxListNum)=0 Then
		PageSplictString = PageSplictString & " ��<b>" & clng(recordCount/DEF_MaxListNum) & "</b>ҳ"
	Else
		If recordCount>=DEF_MaxListNum Then
			PageSplictString = PageSplictString & " ��<b>" & clng(recordCount/DEF_MaxListNum)+1 & "</b>ҳ"
		Else
			PageSplictString = PageSplictString & " ��<b>1</b>ҳ"
		End If
	End If
	PageSplictString = PageSplictString & " ÿҳ<b>" & DEF_MaxListNum & "</b>��"
	PageSplictString = PageSplictString & "</td><td><form action=UserManage.asp><input size=6 name=key value=" & chr(34) & htmlencode(key) & """ class=fminpt><input type=submit name=submit value=�� class=fmbtn>[����ȫ��]</td></form></tr></table>"
	%>
	<script language=javascript>
	function opw(f,r,id)
	{
		document.location.href = f+'&'+r+'='+id;
	}
	</script>
	<table border=0 cellpadding=0 cellspacing=0 width="100%" class=frame_table>
	<tr class=frame_tbhead>
		<td wdith=66><div class=value>ID</div></td>
		<td width=50%><div class=value>����</div></td>
		<td wdith=66><div class=value><%=DEF_PointsName(0)%></div></td>
		<td wdith=120><div class=value>ע��ʱ��</div></td>
		<td wdith=120><div class=value>����¼</div></td>
	</tr>
<%
		For N = MinN to MaxN Step StepValue
			%>
	<tr bgcolor="<%=DEF_BBS_LightestColor%>" class=TBBG9>
		<td class=tdbox><%=GetData(0,n)%></td>
		<td class=tdbox>
  			<a href=<%=DEF_BBS_HomeUrl%>User/LookUserInfo.asp?id=<%=GetData(0,n)%>><%=htmlencode(GetData(1,n))%></a>
  			<a href=UserModify.asp?Form_ID=<%=GetData(0,n)%> title=�޸��û����ϼ�Ȩ��><span class=greenfont>�޸�</span></a>
  			<a href=UserDelete.asp?GBL_CTG_DELETEID=<%=GetData(0,n)%>><span class=redfont title=��������ɾ���û�����>ɾ</span></a>
  			<a href='javascript:opw("UpdateUserAnnounce2.asp?B=<%=GBL_board_ID%>","ID",<%=GetData(0,n)%>);' title=�޸��û���������<%=DEF_PointsName(3)%>���ݺͶ���Ϣ״̬>�޸�</a>
  			<a href='javascript:opw("DelUserAllAnnounce.asp?B=<%=GBL_board_ID%>","DelUserID",<%=GetData(0,n)%>);' title=ɾ�����û��ĺ������ϣ������ղأ��������ӣ��ϴ����������ϣ�����<%=DEF_PointsName(0)%>>ɾ����</a>
  			<a href='javascript:opw("DelUserAllAnnounce.asp?B=<%=GBL_board_ID%>&dflag=onlyupload","DelUserID",<%=GetData(0,n)%>);' title=����ͬɾ�����ϵ���������ֻɾ���û����ϴ�����>ɾ����</a></td>
		<td class=tdbox><%=GetData(2,n)%></td>
		<td class=tdbox><%=RestoreTime(Left(GetData(3,n),8))%></td>
		<td class=tdbox><%=RestoreTime(Left(GetData(4,n),8))%></td>
	</tr><%
			i = i+1
			If i > DEF_MaxListNum then Exit For
		next
%>
	<tr bgcolor=<%=DEF_BBS_TableHeadColor%> class=TBfour>
		<td class=tdbox height="30" valign="bottom" colspan=5>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td height="20">
					<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" height="3">
					<tr>
						<td></td>
					</tr>
					</table>
					<%=PageSplictString%></td>
			</tr>
			</table>
		</td>
	</tr>
	</table>
	<%
	Else
		Response.Write "<div class=alert>û�з��������ļ�¼��</div>" & VbCrLf
	End If

End Function
%>