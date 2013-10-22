<!-- #include file=../inc/BBSSetup.asp -->
<!-- #include file=../inc/User_Setup.ASP -->
<!-- #include file=../inc/Board_Popfun.asp -->
<!-- #include file=inc/UserTopic.asp -->
<!-- #include file=../inc/IncHtm/Top_BoardTop.asp -->
<%
DEF_BBS_HomeUrl = "../"
GBL_CHK_PWdFlag = 0
initDatabase

Dim Tmp_NavStr,Tmp_Para
Tmp_Para = Left(Request.QueryString,1)
Select Case Tmp_Para
	Case "S": Tmp_NavStr = DEF_PointsName(0) & "���а�"
	Case "u": Tmp_NavStr = DEF_PointsName(4) & "���а�"
	Case "p": Tmp_NavStr = "��ˮ���а�"
	Case "e": Tmp_NavStr = "�����û�"
	Case "r": Tmp_NavStr = "�����û�"
	Case "b": Tmp_NavStr = "���淢������"
	Case Else: Tmp_NavStr = DEF_PointsName(0) & "���а�"
End Select

BBS_SiteHead DEF_SiteNameString & " - " & Tmp_NavStr,0,"<span class=navigate_string_step>" & Tmp_NavStr & "</span>"
UpdateOnlineUserAtInfo 0,Tmp_NavStr

UserTopicTopInfo("forum")
UserTop_NavInfo

Select Case Tmp_Para
	Case "S": DisplayUserPointsTop(DEF_MaxListNum)
	Case "u": DisplayUserOnlineTimeTop(DEF_MaxListNum)
	Case "p": DisplayUserAncTop(DEF_MaxListNum)
	Case "e": DisplayUserNewest(DEF_MaxListNum)
	Case "r": DisplayUserFind
	Case "b": DisplayBoardTop
	Case Else: DisplayUserPointsTop(DEF_MaxListNum)
End Select
UserTopicBottomInfo
closeDataBase
SiteBottom

Sub UserTop_NavInfo

	Dim Evol
	Evol = Tmp_Para
	If Evol = "b" Then Exit Sub
	If inStr("Super",Evol) = 0 Then Evol = ""

	Response.Write "<div class='user_item_nav fire'><ul>"
	If Evol = "S" or Evol = "" Then
		Response.Write "	<li><div class=navactive>" & DEF_PointsName(0) & "���а�</div></li>"
	Else
		Response.Write "	<li><a href=UserTop.asp?S>" & DEF_PointsName(0) & "���а�</a></li>"
	End If
	If Evol = "u" Then
		Response.Write "	<li><div class=navactive>" & DEF_PointsName(4) & "���а�</div></li>"
	Else
		Response.Write "	<li><a href=UserTop.asp?u>" & DEF_PointsName(4) & "���а�</a></li>"
	End If
	If Evol = "p" Then
		Response.Write "	<li><div class=navactive>��ˮ���а�</div></li>"
	Else
		Response.Write "	<li><a href=UserTop.asp?p>��ˮ���а�</a></li>"
	End If
	If Evol = "e" Then
		Response.Write "	<li><div class=navactive>�����û�</div></li>"
	Else
		Response.Write "	<li><a href=UserTop.asp?e>�����û�</a></li>"
	End If
	If Evol = "r" Then
		Response.Write "	<li><div class=navactive>�����û�</div></li>"
	Else
		Response.Write "	<li><a href=UserTop.asp?r>�����û�</a></li>"
	End If
	Response.Write "</ul></div>"
	

End Sub

Sub DisplayUserPointsTop(Number)

	Dim Rs,SQL
	
	Rem Order by Points��Order by LeadBBS_User.Points�ǲ�һ����(��ָAccess��),�ɶ��Access�����ὫpointsΪ0���ŵ���ǰ��
	Rem ����,����һ��Where�ж�,��������Access�ر�������Ч��
	SQL = sql_select("select LeadBBS_User.ID,LeadBBS_User.UserName,LeadBBS_User.Points,LeadBBS_User.ApplyTime,LeadBBS_User.Lastdoingtime,LeadBBS_User.UserLevel from LeadBBS_User Order by LeadBBS_User.Points DESC",Number)

	Set Rs = LDExeCute(SQL,0)
	If Not Rs.Eof Then%>
	<script language=javascript>
	var Number=1,i,t=1;
	var DEF_UserLevelString=new Array(<%
		Response.Write """" & DEF_UserLevelString(0) & """"
		for SQL = 1 to DEF_UserLevelNum
			Response.Write ",""" & DEF_UserLevelString(SQL) & """"
		Next
	%>);
	function s(d0,d1,d2,d4,d5,d6)
	{
		if(d0=="")return;
		document.write("<tr");
		document.write("><td class=tdbox>" + t + "</td><td class=tdbox><a href=LookUserInfo.asp?id=" + d0 + ">" + d1 + "</a></td><td class=tdbox>" + d2 + "</td>");
        document.write("<td class=tdbox>" + d4.substr(0,4) + "-" + d4.substr(4,2) + "-" + d4.substr(6,2) + "</td>");
		document.write("<td class=tdbox>" + d5.substr(0,4) + "-" + d5.substr(4,2) + "-" + d5.substr(6,2) + "</td>");
		document.write("<td class=tdbox>" + DEF_UserLevelString[parseInt(d6)] + "</td></tr>");
		t+=1;
	}
	</script>
	<table width="100%" border="0" cellspacing="0" cellpadding="0" class=table_in>
                    <tr class=tbinhead>
                    	<td width=8%><div class=value>����</div></td>
                    	<td width=30%><div class=value>����</div></td>
                    	<td width=15%><div class=value><%=DEF_PointsName(0)%></div></td>
                    	<td width=15%><div class=value>ע��ʱ��</div></td>
                    	<td width=15%><div class=value>�������</div></td>
                    	<td width=17%><div class=value><%=DEF_PointsName(3)%></div></td>
                    </tr><script language=javascript>
<%	
		Response.Write "" & VbCrLf & "s("""
		Response.Write Rs.GetString(,,""",""",""");" & VbCrLf & "s(""","")
		Rs.Close
		Set Rs = Nothing
		Response.Write ""","""","""","""");"
%></script>
                  </table>
	<%
	Else
		Rs.close
		Set Rs = Nothing
		Response.Write "<div class=alert>�����������.</div>" & VbCrLf
	End If

End Sub

Sub DisplayUserOnlineTimeTop(Number)

	Dim Rs,SQL
	SQL = sql_select("select LeadBBS_User.ID,LeadBBS_User.UserName,LeadBBS_User.OnlineTime,LeadBBS_User.ApplyTime,LeadBBS_User.Lastdoingtime,LeadBBS_User.UserLevel from LeadBBS_User Order by LeadBBS_User.OnlineTime DESC",Number)
	Set Rs = LDExeCute(SQL,0)
	Dim Num
	Dim GetData
	If Not Rs.Eof Then%>
	<script language=javascript>
	var Number=1,i,t=1;
	var DEF_UserLevelString=new Array(<%
		Response.Write """" & DEF_UserLevelString(0) & """"
		for SQL = 1 to DEF_UserLevelNum
			Response.Write ",""" & DEF_UserLevelString(SQL) & """"
		Next
	%>);
	function s(d0,d1,d2,d4,d5,d6)
	{
		if(d0=="")return;
		document.write("<tr");
		document.write("><td class=tdbox>" + t + "</td><td class=tdbox><a href=LookUserInfo.asp?id=" + d0 + ">" + d1 + "</a></td><td class=tdbox>" + parseInt(d2/60) + "</td>");
        document.write("<td class=tdbox>" + d4.substr(0,4) + "-" + d4.substr(4,2) + "-" + d4.substr(6,2) + "</td>");
		document.write("<td class=tdbox>" + d5.substr(0,4) + "-" + d5.substr(4,2) + "-" + d5.substr(6,2) + "</td>");
		document.write("<td class=tdbox>" + DEF_UserLevelString[parseInt(d6)] + "</td></tr>");
		t+=1;
	}
	</script>
                  <table width="100%" border="0" cellspacing="0" cellpadding="0" class=table_in>
                    <tr class=tbinhead>
                    	<td width=8%><div class=value>����</div></td>
                    	<td width=30%><div class=value>����</div></td>
                    	<td width=15%><div class=value><%=DEF_PointsName(4)%></div></td>
                    	<td width=15%><div class=value>ע��ʱ��</div></td>
                    	<td width=15%><div class=value>�������</div></td>
                    	<td width=17%><div class=value><%=DEF_PointsName(3)%></div></td>
                    </tr><script language=javascript>
<%	
		Response.Write "" & VbCrLf & "s("""
		Response.Write Rs.GetString(,,""",""",""");" & VbCrLf & "s(""","")
		Rs.Close
		Set Rs = Nothing
		Response.Write ""","""","""","""");"
%></script>
                  </table>
	<%
	Else
		Rs.close
		Set Rs = Nothing
		Response.Write "<div class=alert>�����������.</div>" & VbCrLf
	End If

End Sub

Sub DisplayUserAncTop(Number)

	Dim Rs,SQL
	
	SQL = sql_select("select LeadBBS_User.ID,LeadBBS_User.UserName,LeadBBS_User.AnnounceNum,LeadBBS_User.ApplyTime,LeadBBS_User.Lastdoingtime,LeadBBS_User.UserLevel from LeadBBS_User Order by LeadBBS_User.AnnounceNum DESC",Number)

	Set Rs = LDExeCute(SQL,0)
	If Not Rs.Eof Then%>
	<script language=javascript>
	var Number=1,i,t=1;
	var DEF_UserLevelString=new Array(<%
		Response.Write """" & DEF_UserLevelString(0) & """"
		for SQL = 1 to DEF_UserLevelNum
			Response.Write ",""" & DEF_UserLevelString(SQL) & """"
		Next
	%>);
	function s(d0,d1,d2,d4,d5,d6)
	{
		if(d0=="")return;
		document.write("<tr");
		document.write("><td class=tdbox>" + t + "</td><td class=tdbox><a href=LookUserInfo.asp?id=" + d0 + ">" + d1 + "</a></td><td class=tdbox>" + d2 + "</td>");
        document.write("<td class=tdbox>" + d4.substr(0,4) + "-" + d4.substr(4,2) + "-" + d4.substr(6,2) + "</td>");
		document.write("<td class=tdbox>" + d5.substr(0,4) + "-" + d5.substr(4,2) + "-" + d5.substr(6,2) + "</td>");
		document.write("<td class=tdbox>" + DEF_UserLevelString[parseInt(d6)] + "</td></tr>");t+=1;
	}
	</script>
                  <table width="100%" border="0" cellspacing="0" cellpadding="0" class=table_in>
                    <tr class=tbinhead>
                    	<td width=8%><div class=value>����</div></td>
                    	<td width=30%><div class=value>����</div></td>
                    	<td width=15%><div class=value>����</div></td>
                    	<td width=15%><div class=value>ע��ʱ��</div></td>
                    	<td width=15%><div class=value>�������</div></td>
                    	<td width=17%><div class=value><%=DEF_PointsName(3)%></div></td>
                    </tr><script language=javascript>
<%	
		Response.Write "" & VbCrLf & "s("""
		Response.Write Rs.GetString(,,""",""",""");" & VbCrLf & "s(""","")
		Rs.Close
		Set Rs = Nothing
		Response.Write ""","""","""","""");"
%></script>
                  </table>
	<%
	Else
		Rs.close
		Set Rs = Nothing
		Response.Write "<div class=alert>�����������.</div>" & VbCrLf
	End If

End Sub

Sub DisplayUserNewest(Number)

	Dim Rs,SQL
	SQL = sql_select("select ID,UserName,Points,OnlineTime,ApplyTime,Prevtime from LeadBBS_User Order by ID DESC",Number)
	Set Rs = LDExeCute(SQL,0)
	Dim Num
	If Not Rs.Eof Then	
	%>
	<script language=javascript>
	function s(d0,d1,d2,d3,d4,d5)
	{
		if(d0=="")return;
		document.write("<tr");
		document.write("><td class=tdbox>" + d0 + "</td><td class=tdbox><a href=LookUserInfo.asp?id=" + d0 + ">" + d1 + "</a></td><td class=tdbox>" + d2 + "</td>");
		document.write("<td class=tdbox>" + parseInt(parseInt(d3)/60) + "</td>");
        document.write("<td class=tdbox>" + d4.substr(0,4) + "-" + d4.substr(4,2) + "-" + d4.substr(6,2) + " " + d4.substr(8,2) + ":" + d4.substr(10,2) + "</td>");
		document.write("<td class=tdbox>" + d5.substr(0,4) + "-" + d5.substr(4,2) + "-" + d5.substr(6,2) + " " + d5.substr(8,2) + ":" + d5.substr(10,2) + "</td></tr>");
	}
	</script>
	<table width="100%" border="0" cellspacing="0" cellpadding="0" class=table_in>
		<tr class=tbinhead>
			<td width=12%><div class=value>���</div></td>
			<td width=24%><div class=value>����</div></td>
			<td width=8%><div class=value><%=DEF_PointsName(0)%></div></td>
			<td width=8%><div class=value><%=DEF_PointsName(4)%></div></td>
			<td width=24%><div class=value>ע��ʱ��</div></td>
			<td width=24%><div class=value>����¼</div></td>
		</tr><script language=javascript>
<%	
		Response.Write "" & VbCrLf & "s("""
		Response.Write Rs.GetString(,,""",""",""");" & VbCrLf & "s(""","")
		Rs.Close
		Set Rs = Nothing
		Response.Write ""","""","""","""");"
%></script>
	</table>
	<%
	Else
		Rs.Close
		Set Rs = Nothing
		Response.Write "<div class=alert>�����������.</div>" & VbCrLf
	End If

End Sub

Sub DisplayUserFind

Dim OkFlag
OkFlag = 1
Dim Form_SearchKey
Form_SearchKey = Request.Form("Form_SearchKey")
If Request("SubmitFlag")<>"kdWosoO9w2AXkHouseASP" Then OkFlag = 0

If Len(Form_SearchKey) < 1 Then OkFlag = 0
If OkFlag = 1 Then Form_SearchKey = Left(Form_SearchKey,20)

If OkFlag = 0 Then
	DisplaySearchForm
Else
	Dim Rs,SQL
	SQL = sql_select("select ID,UserName,Points,ApplyTime,Prevtime,UserLevel from LeadBBS_User where UserName ='" & Replace(Form_SearchKey,"'","''") & "'",1)
	Set Rs = LDExeCute(SQL,0)
	Dim Num
	Dim GetData
	If Not Rs.Eof Then
		GetData = Rs.GetRows(1)
		Num = 0
	Else
		Num = -1
	End If
	Rs.close
	Set Rs = Nothing
	
	Dim i,N
	If Num>=0 Then
	%>
	<table width="100%" border="0" cellspacing="0" cellpadding="0" class=table_in>
	<tr class=tbinhead>
		<td wdith=8%><div class=value>ID</div></td>
		<td wdith=30%><div class=value>����</div></td>
		<td wdith=15%><div class=value><%=DEF_PointsName(0)%></div></td>
		<td wdith=15%><div class=value>ע��ʱ��</div></td>
		<td wdith=15%><div class=value>����¼</div></td>
		<td wdith=17%><div class=value><%=DEF_PointsName(3)%></div></td>
	</tr>
<%
		for n= 0 to Num
			%>
	<tr>
		<td class=tdbox><%=GetData(0,n)%></td>
		<td class=tdbox>
			<a href=LookUserInfo.asp?id=<%=GetData(0,n)%>><%=htmlencode(GetData(1,n))%></a></td>
		<td class=tdbox><%=GetData(2,n)%></td>
		<td class=tdbox><%=RestoreTime(Left(GetData(3,n),8))%></td>
		<td class=tdbox><%=RestoreTime(Left(GetData(4,n),8))%></td>
		<td class=tdbox><%=DEF_UserLevelString(GetData(5,n))%></td>
		</tr><%
		next
%>
	</table>
	<%
		DisplaySearchForm
	Else%>
		<div class=title>�����û��� <span class=redfont><%=htmlencode(Form_SearchKey)%></span></div><%
		Response.Write "<div class=alert>�����������.</div>"
		DisplaySearchForm
	End If
End If

End Sub

Sub DisplayBoardTop

	Dim t
	t = DateDiff("s",Top_BoardTop_UpdateTime,DEF_Now)
	If t >= 0 and t <= DEF_UpdateInterval Then
		Top_BoardTop_View
		Exit Sub
	End If
	
	Dim Rs,SQL
	SQL = sql_select("select BoardID,BoardName,AnnounceNum from LeadBBS_Boards where HiddenFlag=0 and BoardID<>444 order by AnnounceNum DESC",DEF_MaxListNum)
	Set Rs = LDExeCute(SQL,0)
	Dim Num
	Dim GetData
	If Not Rs.Eof Then
		GetData = Rs.GetRows(DEF_MaxListNum)
		Num = Ubound(GetData,2)
	Else
		Num = -1
	End If
	Rs.close
	Set Rs = Nothing
	
	Dim i,N,Str
	If Num>=0 Then
Str = "	<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" class=""table_in"">" & VbCrLf &_
"	<tr class=tbinhead>" & VbCrLf &_
"		<td wdith=8% ><div class=value>&nbsp;</div></td>" & VbCrLf &_
"		<td width=77% ><div class=value>����</div></td>" & VbCrLf &_
"		<td wdith=15% ><div class=value>����</div></td>" & VbCrLf &_
"	</tr>" & VbCrLF
		for n= 0 to Num
Str = Str & "	<tr>" & VbCrLf &_
"		<td class=tdbox>" & N+1 & "</td>" & VbCrLf &_
"		<td class=tdbox>" & VbCrLf &_
"			<a href=" & DEF_BBS_HomeUrl & "b/b.asp?b=" & GetData(0,n) & ">" & KillHTMLLabel(GetData(1,n)) & "</a></td>" & VbCrLf &_
"		<td class=tdbox>" & GetData(2,n) & "</td>" & VbCrLf &_
"		</tr>"
		next
Str = Str & "	</table>" & VbCrLf
	Else
		Str = "<div class=alert>���ް���.</div>"
	End If
	Response.Write Str	
	
	Str = "<" & "%" & VbCrLf &_
	"Dim Top_BoardTop_UpdateTime" & VbCrLf &_
	"Top_BoardTop_UpdateTime = """ & htmlencode(DEF_Now) & """" & VbCrLf &_
	"" & VbCrLf &_
	"Sub Top_BoardTop_View" & VbCrLf &_
	"" & VbCrLf &_
	"%" & ">" & VbCrLf &_
	Str &_
	"<" & "%" & VbCrLf &_
	"" & VbCrLf &_
	"End Sub" & VbCrLf &_
	"%" & ">" & VbCrLf
	CALL ADODB_SaveToFile(Str,DEF_BBS_HomeUrl & "inc/IncHtm/Top_BoardTop.asp")

End Sub

Sub DisplaySearchForm

	Dim Form_SearchKey
	Form_SearchKey = Request.Form("Form_SearchKey")
%>
	<form action="UserTop.asp?r" onSubmit="submit_disable(this);" method="post">
	<div class="title">�����������û���: </div>
	<div class="value2">
	<input name="Form_SearchKey" type="text" value="<%=Htmlencode(Form_SearchKey)%>" class="fminpt input_2">
	<input name="SubmitFlag" value="kdWosoO9w2AXkHouseASP" type="hidden">
	</div>
	<div class="value2">
	<input type="submit" value="����" class="fmbtn btn_2">
	</div>
	</Form>
<%
End Sub%>