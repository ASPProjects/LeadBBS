<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/User_Setup.asp -->
<!-- #include file=../../inc/Board_Popfun.asp -->
<!-- #include file=../../inc/Limit_Fun.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<!-- #include file=inc/ForumBoard_fun.asp -->
<%
DEF_BBS_HomeUrl = "../../"
Dim GBL_ID,GBL_ModifyID
initDatabase
GBL_CHK_TempStr = ""
GBL_ID = checkSupervisorPass

Manage_sitehead DEF_SiteNameString & " - ����Ա",""
frame_TopInfo
DisplayUserNavigate("��̳�����޸�")
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
<b>�޸İ���</b>
<%
	GBL_ModifyID = Left(Request("GBL_ModifyID"),14)
	If isNumeric(GBL_ModifyID)=0 Then GBL_ModifyID=0
	GBL_ModifyID = cCur(GBL_ModifyID)
	If GetForumBoardData(GBL_MODIFYID) <> 0 Then
		GBL_BoardID=cCur(GBL_GetData(0,0))
		GBL_BoardID_Old = GBL_BoardID
		GBL_BoardAssort = cCur(GBL_GetData(1,0))
		GBL_BoardAssort_Old = GBL_BoardAssort
		GBL_BoardName=GBL_GetData(2,0)
		GBL_BoardIntro=GBL_GetData(3,0)

		GBL_LastWriter=GBL_GetData(4,0)
		GBL_LastWriteTime=GBL_GetData(5,0)
		GBL_TopicNum=cCur(GBL_GetData(6,0))
		GBL_AnnounceNum=cCur(GBL_GetData(7,0))
		GBL_ForumPass = GBL_GetData(8,0)
		GBL_HiddenFlag = GBL_GetData(9,0)
		GBL_MasterList = GBL_GetData(12,0)
		GBL_BoardLimit = GBL_GetData(13,0)
		GBL_OrderID = cCur(GBL_GetData(18,0))
		GBL_OrderID_Old = GBL_OrderID
		GBL_BoardStyle = GBL_GetData(19,0)
		GBL_MasterList_Old = GBL_MasterList
		GBL_StartTime = Right("0000000" & GBL_GetData(20,0),6)
		GBL_EndTime = Right("0000000" & GBL_GetData(21,0),6)

		GBL_BoardHead = Trim(GBL_GetDAta(22,0))
		GBL_BoardBottom = Trim(GBL_GetDAta(23,0))
		GBL_BoardImgUrl = Trim(GBL_GetDAta(24,0))
		GBL_BoardImgWidth = GBL_GetDAta(25,0)
		GBL_BoardImgHeight = GBL_GetDAta(26,0)

		GBL_ParentBoard = cCur(GBL_GetDAta(27,0))
		GBL_ParentBoard_Old = cCur(GBL_ParentBoard)
		GBL_LowerBoard = GBL_GetData(28,0)
		GBL_OtherLimit = GBL_GetData(35,0)

		GBL_LimitHourStart = Mid(GBL_StartTime,1,2)
		GBL_LimitWeekStart = Mid(GBL_StartTime,3,2)
		GBL_LimitMonthStart = Mid(GBL_StartTime,5,2)

		GBL_LimitHourEnd = Mid(GBL_EndTime,1,2)
		GBL_LimitWeekEnd = Mid(GBL_EndTime,3,2)
		GBL_LimitMonthEnd = Mid(GBL_EndTime,5,2)
		
		REM GBL_OtherLimit_Part1 -- ʮ��������������λ,�������Ʒ�ʽ,�������99�����Ʒ�ʽ
		REM GBL_OtherLimit_Part2 -- ʮ���������������Ĳ���,�������Ʒ�ʽ�ľ���Ҫ����ֵ
		GBL_OtherLimit_Part1 = cCur(Right(GBL_OtherLimit,2))
		If Len(GBL_OtherLimit) > 2 Then
			GBL_OtherLimit_Part2 = cCur(Left(GBL_OtherLimit,Len(GBL_OtherLimit)-2))
		Else
			GBL_OtherLimit_Part2 = 0
		End If

		GBL_CHK_TempStr = ""
		'If GBL_ForumPass <> "" Then
		'	Response.Write "<div class=alert>�˰����Ѿ���ֹ�޸ģ�</div>" & VbCrLf
		'Else
			If Request.Form("submitflag")="LKOkxk2" Then
				GBL_BoardID = Left(Trim(Request.Form("GBL_BoardID")),14)
				GBL_BoardAssort = Left(Trim(Request.Form("GBL_BoardAssort")),14)
				GBL_BoardName = Trim(Request.Form("GBL_BoardName"))
				GBL_BoardIntro = Trim(Request.Form("GBL_BoardIntro"))
				GBL_LastWriter = Trim(Request.Form("GBL_LastWriter"))
				GBL_LastWriteTime = Trim(Request.Form("GBL_LastWriteTime"))
				GBL_TopicNum = Left(Trim(Request.Form("GBL_TopicNum")),14)
				GBL_AnnounceNum = Trim(Request.Form("GBL_AnnounceNum"))
				GBL_ForumPass = Trim(Request.Form("GBL_ForumPass"))
				GBL_HiddenFlag = Trim(Request.Form("GBL_HiddenFlag"))
				GBL_MasterList = Trim(Request.Form("GBL_MasterList"))
				GBL_OrderID = Left(Trim(Request.Form("GBL_OrderID")),14)
				GBL_BoardStyle = Left(Trim(Request.Form("GBL_BoardStyle")),14)
		
				GBL_LimitWeekStart = Left(Trim(Request.Form("GBL_LimitWeekStart")),14)
				GBL_LimitWeekEnd = Left(Trim(Request.Form("GBL_LimitWeekEnd")),14)
				GBL_LimitMonthEnd = Left(Trim(Request.Form("GBL_LimitMonthEnd")),14)
				GBL_LimitMonthStart = Left(Trim(Request.Form("GBL_LimitMonthStart")),14)
				GBL_LimithourStart = Left(Trim(Request.Form("GBL_LimithourStart")),14)
				GBL_LimithourEnd = Left(Trim(Request.Form("GBL_LimithourEnd")),14)
				
				GBL_BoardImgUrl = Trim(Request.Form("GBL_BoardImgUrl"))
				GBL_BoardImgWidth = Left(Trim(Request.Form("GBL_BoardImgWidth")),14)
				GBL_BoardImgHeight = Left(Trim(Request.Form("GBL_BoardImgHeight")),14)
				
				GBL_ParentBoard = Left(Trim(Request.Form("BoardID2")),14)
				If Trim(Request.Form("BoardID3")) <> "" Then GBL_ParentBoard = Left(Trim(Request.Form("BoardID3")),14)
				
				GBL_BoardHead = Left(Request.Form("GBL_BoardHead"),8*1024)
				GBL_BoardBottom = Left(Request.Form("GBL_BoardBottom"),8*1024)
				GBL_OtherLimit_Part1 = Left(Request.Form("GBL_OtherLimit_Part1"),14)
				GBL_OtherLimit_Part2 = Left(Request.Form("GBL_OtherLimit_Part2"),14)
				
				If GBL_OtherLimit_Part1 = "5" Then
					GBL_OtherLimit_Part2 = Left(Request.Form("GBL_UserOfficerString"),14)
					If isNumeric(GBL_OtherLimit_Part2) = 0 Then GBL_OtherLimit_Part2 = 0
					GBL_OtherLimit_Part2 = cCur(Fix(GBL_OtherLimit_Part2))
				End If

				If CheckFormForumBoardData = 0 Then
					Response.Write "<div class=alert>���ݲ���ͨ����" & GBL_CHK_TempStr & "</div>" & VbCrLf
					DisplayJoinForm
				Else
					If UpdateForumBoard = 0 Then
						Response.Write "<div class=alert>�޸ĳ���" & GBL_CHK_TempStr & "</div>" & VbCrLf
						DisplayJoinForm
					Else
						Response.Write "<div class=alert><span class=greenfont><b>�޸ĳɹ�!</b></span></div>" & VbCrLf
					End If
				End If
			Else
				DisplayJoinForm
			End If
		'End If
	Else
		Response.Write "<div class=alert>����δѡ��Ҫ�޸ĵİ��档</div>" & VbCrLf
	End If

End Function

Function DisplayForumAssortList

	Dim Rs,GetData,N,TempAssort
	If isNumeric(GBL_BoardAssort)=0 Then
		TempAssort = 0
	Else
		TempAssort = cCur(GBL_BoardAssort)
	End If
	Set Rs = LDExeCute("Select * from LeadBBS_Assort order by AssortID",0)
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
		GetData(1,n) = KillHTMLLabel(GetData(1,n))
		If StrLength(GetData(1,n)) > 31 Then
			GetData(1,n) = LeftTrue(GetData(1,n),28) & "..."
		End If
		If cCur(GetData(0,N))=TempAssort Then
			Response.Write "<option value=" & GetData(0,N) & " selected>" & GetData(1,n) & "</option>" & VbCrLf
		Else
			Response.Write "<option value=" & GetData(0,N) & ">" & GetData(1,n) & "</option>" & VbCrLf
		End If
	Next

End Function

Function DisplayJoinForm%>

	<form action=ForumBoardModify.asp name=form1 id=form1 method=post>
	<table border=0 cellpadding=0 cellspacing=0 width="100%" class=frame_table>
	<tr>
		<td class=tdbox width=120>��̳������:</td>
		<td class=tdbox><input name=GBL_BoardID value="<%=htmlencode(GBL_BoardID)%>" readonly class=fminpt>
		<input name=submitflag type=hidden value="LKOkxk2"></td>
	</tr>
	<tr>
		<td class=tdbox>��������˳��:</td><td class=tdbox><input name=GBL_OrderID value="<%=htmlencode(GBL_OrderID)%>" class=fminpt> <font color=888888 class=grayfont>ԽСԽǰ��</font></td>
	</tr>
	<tr>
		<td class=tdbox>���������:</td><td class=tdbox><select name=GBL_BoardAssort style="width:130;"><option value=0>��ѡ�����</option><%DisplayForumAssortList%></select>
		</td>
	</tr>
	<tr>
		<td class=tdbox>�ϼ�������:</td><td class=tdbox><!-- #include file=../../inc/incHTM/BoardForMoveList.asp -->
		����д������:<input name=BoardID3 value="<%=htmlencode(Trim(Request.Form("BoardID3")))%>" size=6 maxlength=14 class=fminpt>
		<br>�����˴��һ��Ҫ�ǵ�<font color=red class=redfont>������̳�б��޸�</font>����������ͳ�ƻ��г���
		<script>
			var provincebox = document.form1.BoardID2.options,i;
			for(i = 0; i < provincebox.length; i++)
			{
				if(provincebox.options[i].value=="<%=GBL_ParentBoard%>")
				{provincebox.selectedIndex = i;break;}
			}
		</script>
		</td>
	</tr>
	<tr>
		<td class=tdbox>
		<input name=GBL_ModifyID type=hidden value="<%=htmlencode(GBL_ModifyID)%>" class=fminpt>
		��̳��������:</td><td class=tdbox><input name=GBL_BoardName value="<%=htmlencode(GBL_BoardName)%>" class=fminpt></td>
	</tr>
	<tr>
		<td class=tdbox>���������:<br>����ʹ��HTML</td><td class=tdbox><textarea name=GBL_BoardIntro rows=3 cols=41 class=fmtxtra><%If GBL_BoardIntro <> "" Then Response.Write VbCrLf & Server.htmlEncode(GBL_BoardIntro)%></textarea></td>
	</tr>
	<tr>
		<td class=tdbox>��󷢱��û�:</td><td class=tdbox><input name=GBL_LastWriter value="<%=htmlencode(GBL_LastWriter)%>" class=fminpt></td>
	</tr>
	<tr>
		<td class=tdbox>��󷢱�ʱ��:</td><td class=tdbox><input name=GBL_LastWriteTime value="<%=htmlencode(GBL_LastWriteTime)%>" class=fminpt></td>
	</tr>
	<tr>
		<td class=tdbox>��̳��������:</td><td class=tdbox><input name=GBL_TopicNum value="<%=htmlencode(GBL_TopicNum)%>" class=fminpt></td>
	</tr>
	<tr>
		<td class=tdbox>��̳��������:</td><td class=tdbox><input name=GBL_AnnounceNum value="<%=htmlencode(GBL_AnnounceNum)%>" class=fminpt></td>
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
		<td class=tdbox>��̳��������:</td><td class=tdbox><input name=GBL_ForumPass value="<%=htmlencode(GBL_ForumPass)%>" maxlength=20 class=fminpt></td>
	</tr>
	<tr>
		<td class=tdbox>��̳ͼƬ��ַ:</td><td class=tdbox><input name=GBL_BoardImgUrl value="<%=htmlencode(GBL_BoardImgUrl)%>" maxlength=255 size=40 class=fminpt>(��ʹ�þ���·��)</td>
	</tr>
	<tr>
		<td class=tdbox>��̳ͼƬ��С:</td><td class=tdbox>���<input name=GBL_BoardImgWidth value="<%=htmlencode(GBL_BoardImgWidth)%>" maxlength=3 size=3 class=fminpt> �߶�<input name=GBL_BoardImgHeight value="<%=htmlencode(GBL_BoardImgHeight)%>" maxlength=3 size=3 class=fminpt> ע���С��0-200֮��</td>
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
			<table><tr><td><Select name=GBL_OtherLimit_Part1 onchange="if(value<1){GBL_UserOfficerString.style.display='none';GBL_OtherLimit_Part2.style.display='none';}else{if(value<5){GBL_UserOfficerString.style.display='none';GBL_OtherLimit_Part2.style.display='block';}else{GBL_UserOfficerString.style.display='block';GBL_OtherLimit_Part2.style.display='none';}}">
				<option value=0<%If GBL_OtherLimit_Part1 = 0 Then Response.Write " selected"%>>====������====</option>
				<option value=1<%If GBL_OtherLimit_Part1 = 1 Then Response.Write " selected"%>>��Ҫ<%=DEF_PointsName(0)%></option>
				<option value=2<%If GBL_OtherLimit_Part1 = 2 Then Response.Write " selected"%>>��Ҫ<%=DEF_PointsName(4)%>[����ʱ��]</option>
				<option value=3<%If GBL_OtherLimit_Part1 = 3 Then Response.Write " selected"%>>��Ҫ<%=DEF_PointsName(1)%></option>
				<option value=4<%If GBL_OtherLimit_Part1 = 4 Then Response.Write " selected"%>>��Ҫ<%=DEF_PointsName(2)%></option>
				<option value=5<%If GBL_OtherLimit_Part1 = 5 Then Response.Write " selected"%>>ֻ����<%=DEF_PointsName(9)%></option>
			</select></td><td>
			<select name=GBL_UserOfficerString id=GBL_UserOfficerString<%If GBL_OtherLimit_Part1 <> 5 Then Response.Write " Style=""display:none"""%>>
				<%Dim N	
				for N = 0 to DEF_UserOfficerNum
					Response.Write "<option value=" & N
					If N = GBL_OtherLimit_Part2 and GBL_OtherLimit_Part1 = 5 Then Response.Write " selected"
					Response.Write ">" & DEF_UserOfficerString(N) & "</option>" & VbCrLf
				Next%></select></td><td>
			<input name=GBL_OtherLimit_Part2 id=GBL_OtherLimit_Part2 value="<%If GBL_OtherLimit_Part1 <> 5 Then Response.Write htmlencode(GBL_OtherLimit_Part2)%>" maxlength=12 size=12 class=fminpt<%If GBL_OtherLimit_Part1 = 5 or GBL_OtherLimit_Part1 = 0 Then Response.Write " Style=""display:none"""%>>
			</td></tr></table>
		</td>
	</tr>
	<tr>
		<td class=tdbox>��̳��ʱ�趨:<br>
		0��ʾ����</td><td class=tdbox>Сʱ�رգ�<input name=GBL_LimitHourStart value="<%=htmlencode(GBL_LimitHourStart)%>" maxlength=2 size=2 class=fminpt>ʱ-<input name=GBL_LimitHourEnd value="<%=htmlencode(GBL_LimitHourEnd)%>" maxlength=2 size=2 class=fminpt>ʱ (0-23)
		<br>
		���ڹرգ�����<input name=GBL_LimitWeekStart value="<%=htmlencode(GBL_LimitWeekStart)%>" maxlength=2 size=2 class=fminpt>-����<input name=GBL_LimitWeekEnd value="<%=htmlencode(GBL_LimitWeekEnd)%>" maxlength=2 size=2 class=fminpt> (1-7)
		<br>
		�·ݹرգ���<input name=GBL_LimitMonthStart value="<%=htmlencode(GBL_LimitMonthStart)%>" maxlength=2 size=2 class=fminpt>��-��<input name=GBL_LimitMonthEnd value="<%=htmlencode(GBL_LimitMonthEnd)%>" maxlength=2 size=2 class=fminpt>�� (1-31)</td>
	</tr>	
	<tr>
		<td class=tdbox>������ͷ��:<br>ʹ��HTML�﷨<br>�������潫<br>�������û���<br>�����������<br>��</td><td class=tdbox><textarea name=GBL_BoardHead rows=6 cols=51 class=fmtxtra><%If GBL_BoardHead <> "" Then Response.Write VbCrLf & Server.htmlEncode(GBL_BoardHead)%></textarea></td>
	</tr>
	<tr>
		<td class=tdbox>������β��:<br>ʹ��HTML�﷨<br>���������<br>�����ض�</td><td class=tdbox><textarea name=GBL_BoardBottom rows=6 cols=51 class=fmtxtra><%If GBL_BoardBottom <> "" Then Response.Write VbCrLf & Server.htmlEncode(GBL_BoardBottom)%></textarea></td>
	</tr>
	<tr>
		<td class=tdbox colspan=2><input type=submit value="�ύ" class=fmbtn> <input type=reset value="ȡ��" class=fmbtn></td>
	</tr>
	</table></form>

<%End Function%>