<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Board_Popfun.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<!-- #include file=inc/ForumBoard_fun.asp -->
<!-- #include file=../../inc/Limit_Fun.asp -->
<%
DEF_BBS_HomeUrl = "../../"
Dim GBL_ID
initDatabase
GBL_CHK_TempStr = ""
GBL_ID = checkSupervisorPass

Manage_sitehead DEF_SiteNameString & " - ����Ա",""
frame_TopInfo
DisplayUserNavigate("�����̳����")
If GBL_CHK_Flag=1 Then
	LoginAccuessFul
Else
DisplayLoginForm
End If
frame_BottomInfo
closeDataBase
Manage_Sitebottom("none")

Function LoginAccuessFul

%>
<b>��Ӱ���</b>
          <%
          GBL_CHK_TempStr = ""
          If Request.Form("submitflag")="LKOkxk2" Then
          	GBL_BoardID = Left(Trim(Request.Form("GBL_BoardID")),14)
          	GBL_BoardAssort = Left(Trim(Request.Form("GBL_BoardAssort")),14)
          	GBL_BoardName = Trim(Request.Form("GBL_BoardName"))
          	GBL_BoardIntro = Trim(Request.Form("GBL_BoardIntro"))
          	GBL_LastWriter = Trim(Request.Form("GBL_LastWriter"))
          	GBL_LastWriteTime = Left(Trim(Request.Form("GBL_LastWriteTime")),14)
          	GBL_TopicNum = Left(Trim(Request.Form("GBL_TopicNum")),14)
          	GBL_AnnounceNum = Left(Trim(Request.Form("GBL_AnnounceNum")),14)
          	GBL_ForumPass = Trim(Request.Form("GBL_ForumPass"))
          	GBL_HiddenFlag = Left(Trim(Request.Form("GBL_HiddenFlag")),14)
          	GBL_MasterList = Trim(Request.Form("GBL_MasterList"))
		GBL_OtherLimit_Part1 = Left(Request.Form("GBL_OtherLimit_Part1"),14)
		GBL_OtherLimit_Part2 = Left(Request.Form("GBL_OtherLimit_Part2"),14)

          	If CheckFormForumBoardData=0 Then
          		Response.Write "<div class=alert>���ݲ���ͨ����" & GBL_CHK_TempStr & "</div>" & VbCrLf
          		DisplayJoinForm
          	Else
          		If InsertForumBoard = 0 Then
          			Response.Write "<div class=alert>�������" & GBL_CHK_TempStr & "</div>" & VbCrLf
          			DisplayJoinForm
          		Else
          			Response.Write "<div class=alertdone>��ӳɹ�!</div>" & VbCrLf
          		End If
          	End If
          Else
          	DisplayJoinForm
          End If

End Function

Function DisplayForumAssortList

	Dim Rs,GetData,N,TempAssort
	If isNumeric(GBL_BoardAssort)=0 Then
		TempAssort = 0
	Else
		TempAssort = cCur(GBL_BoardAssort)
	End If
	Set Rs = LDExeCute(sql_select("Select * from LeadBBS_Assort order by AssortID",1000),0)
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		Exit Function
	Else
		GetData = Rs.GetRows(-1)
		Rs.Close
		Set Rs = Nothing
	End If
	For N = 0 to Ubound(GetData,2)
		If cCur(GetData(0,N))=TempAssort Then
			Response.Write "<option value=" & GetData(0,N) & " selected>" & GetData(1,n) & "</option>" & VbCrLf
		Else
			Response.Write "<option value=" & GetData(0,N) & ">" & GetData(1,n) & "</option>" & VbCrLf
		End If
	Next

End Function

Function DisplayJoinForm%>
          <form action=ForumBoardJoin.asp method=post name=form1 id=form1>
          <table border=0 cellpadding=0 cellspacing=0 width="100%" class=frame_table>
          <tr>
          	<td class=tdbox width=120>��̳������:</td>
          	<td class=tdbox><input name=GBL_BoardID value="<%=htmlencode(GBL_BoardID)%>" class=fminpt>(���444������վ����ר��,������������)
          	<input name=submitflag type=hidden value="LKOkxk2"></td>
          </tr>
          <tr>
          	<td class=tdbox>��̳��������:</td><td class=tdbox><input name=GBL_BoardName value="<%=htmlencode(GBL_BoardName)%>" class=fminpt>(֧��html)</td>
          </tr>
          <tr>
          	<td class=tdbox>���������:</td><td class=tdbox><select name=GBL_BoardAssort><option value=0>��ѡ�����</option><%DisplayForumAssortList%></select></td>
          </tr>
          <tr>
          	<td class=tdbox>���������:<br>����ʹ��HTML</td><td class=tdbox><textarea name=GBL_BoardIntro rows=3 cols=41 class=fmtxtra><%If GBL_BoardIntro <> "" Then Response.Write VbCrLf & htmlEncode(GBL_BoardIntro)%></textarea></td>
          </tr>
          <tr>
          	<td class=tdbox>��󷢱��û�:</td><td class=tdbox><input name=GBL_LastWriter value="<%=htmlencode(GBL_LastWriter)%>" maxlength=20 class=fminpt></td>
          </tr>
          <tr>
          	<td class=tdbox>��󷢱�ʱ��:</td><td class=tdbox><input name=GBL_LastWriteTime value="<%=htmlencode(GBL_LastWriteTime)%>" maxlength=50 class=fminpt></td>
          </tr>
          <tr>
          	<td class=tdbox>��̳��������:</td><td class=tdbox><input name=GBL_TopicNum value="<%=htmlencode(GBL_TopicNum)%>" maxlength=20 class=fminpt></td>
          </tr>
          <tr>
          	<td class=tdbox>��̳��������:</td><td class=tdbox><input name=GBL_AnnounceNum value="<%=htmlencode(GBL_AnnounceNum)%>" maxlength=20 class=fminpt></td>
          </tr>
          <tr>
          	<td class=tdbox>��̳��ʾ״̬:</td><td class=tdbox>
			<select name="GBL_HiddenFlag">
			<%
		Dim TempN
		If GBL_HiddenFlag = "" or inStr(GBL_HiddenFlag,",") > 0 or isNumeric(GBL_HiddenFlag) = 0 Then GBL_HiddenFlag=0
		GBL_HiddenFlag = Clng(GBL_HiddenFlag)
                  For TempN = 0 to GBL_HiddenFlagNum
                  	%><option value="<%=TempN%>"<%If GBL_HiddenFlag = TempN Then Response.Write " Selected"%>><%=GBL_HiddenFlagData(TempN)%></option>
                  <%Next%>
			</select></td>
          </tr>
          <tr>
          	<td class=tdbox>����Ĭ�Ϸ��:</td><td class=tdbox>
			<select name="GBL_BoardStyle">
			<%
		If GBL_BoardStyle = "" or inStr(GBL_BoardStyle,",") > 0 or isNumeric(GBL_BoardStyle) = 0 Then GBL_BoardStyle=0
		GBL_BoardStyle = Clng(GBL_BoardStyle)
                  For TempN = 0 to DEF_BoardStyleStringNum
                  	%><option value="<%=TempN%>"<%If GBL_BoardStyle = TempN Then Response.Write " Selected"%>><%=DEF_BoardStyleString(TempN)%></option>
                  <%Next%>
			</select></td>
          </tr>
          <tr>
          	<td class=tdbox>��̳��������:</td><td class=tdbox><input name=GBL_ForumPass value="<%=htmlencode(GBL_ForumPass)%>" maxlength=20 class=fminpt>(�û�����˰�����Ҫ������Ӧ������)</td>
          </tr>
          <tr>
          	<td class=tdbox><%=DEF_PointsName(8)%>�б�:</td><td class=tdbox><input name=GBL_MasterList value="<%=htmlencode(GBL_MasterList)%>" maxlength=250 size=28 class=fminpt>(���ŷָ�,ȫ�������д<span style="cursor:hand" onclick="document.form1.GBL_MasterList.value='?LeadBBS?';">?LeadBBS?</span>)</td>
	      </tr>
		<tr>
			<td class=tdbox width=80>��̳Ȩ������:</td>
			<td class=tdbox valign=top><%
		for TempN = 0 to LimitBoardStringDataNum%>	 
				<input type="checkbox" class=fmchkbox name="Limit<%=TempN+1%>" value="1"<%If GetBinarybit(GBL_BoardLimit,TempN+1) = 1 Then
					Response.Write " checked>"
				Else
					Response.Write ">"
				End If%><%=LimitBoardStringData(tempN)%><br>
				<%Next%></td>
		</tr>
	<tr>
		<td class=tdbox width=80>�����������:</td>
		<td class=tdbox valign=top>
			<Select name=GBL_OtherLimit_Part1>
				<option value=0<%If GBL_OtherLimit_Part1 = 0 Then Response.Write " selected"%>>====������====</option>
				<option value=1<%If GBL_OtherLimit_Part1 = 1 Then Response.Write " selected"%>>��Ҫ<%=DEF_PointsName(0)%></option>
				<option value=2<%If GBL_OtherLimit_Part1 = 2 Then Response.Write " selected"%>>��Ҫ<%=DEF_PointsName(4)%>[����ʱ��]</option>
				<option value=3<%If GBL_OtherLimit_Part1 = 3 Then Response.Write " selected"%>>��Ҫ<%=DEF_PointsName(1)%></option>
				<option value=4<%If GBL_OtherLimit_Part1 = 4 Then Response.Write " selected"%>>��Ҫ<%=DEF_PointsName(2)%></option>
			</select>
			<input name=GBL_OtherLimit_Part2 value="<%=htmlencode(GBL_OtherLimit_Part2)%>" maxlength=12 size=12 class=fminpt>
		</td>
	</tr>
          <tr>
          	<td class=tdbox><input type=submit value="�ύ" class=fmbtn> <input type=reset value="ȡ��" class=fmbtn></td>
          </tr>
          </table></form>
<%End Function%>