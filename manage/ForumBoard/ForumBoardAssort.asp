<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Board_Popfun.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<%
DEF_BBS_HomeUrl = "../../"
initDatabase
GBL_CHK_TempStr = ""
checkSupervisorPass

Dim Action
Action = Left(Request("Action"),14)
If Action <> "Join" and Action <> "Modify" and Action <> "Delete" Then
	Action = "Manage"
End If

Dim LMT_OrderID,LMT_BoardID,LMT_AssortName,LMT_GoodNum
LMT_OrderID = 0
LMT_BoardID = 0
LMT_GoodNum = 0

Dim LMT_ID,Old_Board

Manage_sitehead DEF_SiteNameString & " - ����Ա",""

frame_TopInfo
DisplayUserNavigate("��̳����ר������")%>
<p><a href=ForumBoardAssort.asp>�������ר��</a>
<a href=ForumBoardAssort.asp?action=Join>��Ӱ���ר��</a>
</p>
<%If GBL_CHK_Flag=1 Then
	Select Case Action:
		Case "Join": Join
		Case "Modify": Join
		Case "Delete": Delete
		Case "Manage": Manage
	End Select
Else
DisplayLoginForm
End If
frame_BottomInfo
closeDataBase
Manage_Sitebottom("none")

Function Join

	If Action = "Modify" Then
		LMT_ID = Left(Trim(Request("ID")),14)
		If isNumeric(LMT_ID) = 0 Then LMT_ID = 0
		LMT_ID = Fix(cCur(LMT_ID))
		If LMT_ID = 0 or CheckParentAssortIDExist(LMT_ID) = 0 Then
			Response.Write "<div class=alert>�༭��ר��������!</div>" & VbCrLf
			Exit Function
		End If
	End If
	%>
	<b><%
	If Action = "Modify" Then
		Response.Write "�༭"
	Else
		Response.Write "���"
	End If%>����ר��</b>
	<%
		GBL_CHK_TempStr = ""
		If Request.Form("submitflag")="LKOkxk2" Then
			If CheckFormData=0 Then
				Response.Write "<div class=alert>������Ϣ��" & GBL_CHK_TempStr & "</div>" & VbCrLf
				DisplayJoinForm
	        		Else
				If UpdateAssort = 0 Then
					Response.Write "<div class=alert>�������" & GBL_CHK_TempStr & "</div>" & VbCrLf
					DisplayJoinForm
				Else
					UpdateCacheData("data_goodassort.asp")
					Response.Write "<div class=alertdone>�ɹ�����!</div>" & VbCrLf
				End If
			End If
		Else
			DisplayJoinForm
		End If

End Function

Function DisplayJoinForm

	If Action = "Modify" Then
		DisplayModifyForm
		Exit Function
	End If%>
	<form action=ForumBoardAssort.asp method=post name=form1 id=form1>
	<table border=0 cellpadding=0 cellspacing=0 width="100%" class=frame_table>
	<tr>
		<td class=tdbox width=120>
			<input name=action type=hidden value="Join">
			<input name=submitflag type=hidden value="LKOkxk2">
			������ţ�</td>
		<td class=tdbox align=left><input name=Form_AssortID size=4 maxlength=4 value="<%=htmlencode(LMT_OrderID)%>" class=fminpt>
			��ʾ�ڰ�������е�ǰ��˳������ԽСԽ��ǰ</td>
	</tr>
	<tr>
		<td class=tdbox width=80>
			��������:</td>
		<td class=tdbox>
			<!-- #include file=../../inc/incHTM/BoardForMoveList.asp -->
		<script>
			var provincebox = document.form1.BoardID2.options,i;
			for(i = 0; i < provincebox.length; i++)
			{
				if(provincebox.options[i].value=="<%=LMT_BoardID%>")
				{provincebox.selectedIndex = i;break;}
			}
		</script>��ѡ���ʾ������̳��ר��
		</td>
	</tr>
	<tr>
		<td class=tdbox width=80>
			ר�����ƣ�</td>
		<td class=tdbox align=left><input name=LMT_AssortName size=40 maxlength=255 value="<%=htmlencode(LMT_AssortName)%>" class=fminpt>
			<br>����ʹ��HTML���255��</td>
	</tr>
	<tr>
		<td class=tdbox colspan=2>
			<input name=LMT_GoodNum type=hidden value="0">
			<input type=submit value="�ύ" class=fmbtn> <input type=reset value="ȡ��" class=fmbtn>
		</td>
	</tr>
	</table></form>

<%End Function

Function DisplayModifyForm

	%>
	<form action=ForumBoardAssort.asp method=post name=form1 id=form1>
	<table border=0 cellpadding=0 cellspacing=0 width="100%" class=frame_table>
	<tr>
		<td class=tdbox width=120>
			<input name=action type=hidden value="Modify">
			<input name=submitflag type=hidden value="LKOkxk2">
			<input name=ID type=hidden value="<%=LMT_ID%>">
			������ţ�</td>
		<td class=tdbox align=left><input name=LMT_OrderID size=4 maxlength=4 value="<%=htmlencode(LMT_OrderID)%>" class=fminpt>
			��ʾ�ڰ�������е�ǰ��˳������ԽСԽ��ǰ</td>
	</tr>	
	<tr>
		<td class=tdbox width=80>
			��������:</td>
		<td class=tdbox>
			<!-- #include file=../../inc/incHTM/BoardForMoveList.asp -->
		<script>
			var provincebox = document.form1.BoardID2.options,i;
			for(i = 0; i < provincebox.length; i++)
			{
				if(provincebox.options[i].value=="<%=LMT_BoardID%>")
				{provincebox.selectedIndex = i;break;}
			}
		</script>��ѡ���ʾ������̳��ר��
		</td>
	</tr>
	<tr>
		<td class=tdbox width=80>
			ר�����ƣ�</td>
		<td class=tdbox align=left><input name=LMT_AssortName size=40 maxlength=255 value="<%=htmlencode(LMT_AssortName)%>" class=fminpt>
			<br>����ʹ��HTML���255��</td>
	</tr>
	<tr>
		<td class=tdbox width=80>
			����ͳ�ƣ�</td>
		<td class=tdbox align=left><input name=LMT_GoodNum type=checkbox value="yes" class=fmchkbox checked>����ͳ�ƴ�ר����ӵ�е���������</td>
	</tr>
	<tr>
		<td class=tdbox colspan=2>
			<input type=submit value="�ύ" class=fmbtn> <input type=reset value="ȡ��" class=fmbtn>
		</td>
	</tr>
	</table></form>

<%End Function

Function CheckFormData

	Dim Temp

	LMT_OrderID = Left(Trim(Request.Form("LMT_OrderID")),14)
	If Action = "Join" or Action = "Modify" Then LMT_BoardID = Left(Trim(Request.Form("BoardID2")),14)
	LMT_AssortName = Trim(Request.Form("LMT_AssortName"))

	If isNumeric(LMT_OrderID) = 0 Then LMT_OrderID = 0
	LMT_OrderID = Fix(cCur(LMT_OrderID))
	If LMT_OrderID < 0 Then LMT_OrderID = 0

	If isNumeric(LMT_BoardID) = 0 or LMT_BoardID = "" Then
		LMT_BoardID = 0
		'GBL_CHK_TempStr = "��ѡ����ȷ���������档<br>" & VbCrLf
		'CheckFormData = 0
		'Exit Function
	End If
	
	LMT_BoardID = Fix(cCur(LMT_BoardID))
	Temp = Application(DEF_MasterCookies & "BoardInfo" & LMT_BoardID)
	If isArray(Temp) = False Then
		ReloadBoardInfo(LMT_BoardID)
		Temp = Application(DEF_MasterCookies & "BoardInfo" & LMT_BoardID)
	End If

	If isArray(Temp) = False Then
		'GBL_CHK_TempStr = "�������治���ڣ���ȷ���Ƿ��Ѿ���ȷѡ��!<br>" & VbCrLf
		'CheckFormData = 0
		LMT_BoardID = 0
	End If

	If Len(LMT_AssortName) > 255 or LMT_AssortName = "" Then
		GBL_CHK_TempStr = "������дר�����ֲ��Ҳ��ܳ���255�֡�<br>" & VbCrLf
		CheckFormData = 0
		Exit Function
	End If

	If inStr(LCase(LMT_AssortName),"'") > 0 or inStr(LCase(LMT_AssortName),"<script") > 0 or inStr(LCase(LMT_AssortName),"<\script") > 0 or inStr(LCase(LMT_AssortName),"</script") > 0 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "Error: ר�����ֲ�������뵥���Ż�js����������<br>" & VbCrLf
		CheckFormData = 0
		Exit Function
	End If		

	CheckFormData = 1

End Function

Function UpdateAssort

	If Action = "Join" Then
		LMT_GoodNum = 0
		CALL LDExeCute("inSert into LeadBBS_GoodAssort(OrderID,BoardID,AssortName,GoodNum) Values(" & _
				LMT_OrderID & "," & LMT_BoardID & ",'" & Replace(LMT_AssortName,"'","''") & "'," & LMT_GoodNum & ")",1)
		ReloadTopicAssort(LMT_BoardID)
	Else
		If Request.Form("LMT_GoodNum") <> "" Then
			Dim Rs
			select case DEF_UsedDataBase
			case 0,2:
				Set Rs = LDExeCute("Select count(*) from LeadBBS_Announce where GoodAssort=" & LMT_ID,0)
			case Else
				Set Rs = LDExeCute("Select count(*) from LeadBBS_Topic where GoodAssort=" & LMT_ID,0)
			End select
			If Rs.Eof Then
				LMT_GoodNum = 0
			Else
				LMT_GoodNum = Rs(0)
				If isNull(LMT_GoodNum) Then LMT_GoodNum = 0
				LMT_GoodNum = cCur(LMT_GoodNum)
			End If
			Rs.Close
			Set Rs = Nothing
		End If
		CALL LDExeCute("Update LeadBBS_GoodAssort Set OrderID=" & LMT_OrderID & _
			",AssortName='" & Replace(LMT_AssortName,"'","''") & "'" & _
			",GoodNum=" & LMT_GoodNum & _
			",BoardID=" & LMT_BoardID & _
			" where ID=" & LMT_ID,1)
		ReloadTopicAssort(LMT_BoardID)
		If Old_Board <> LMT_BoardID Then ReloadTopicAssort(Old_Board)
	End If
	UpdateAssort = 1

End Function

Rem ���ר�����ID�Ƿ����
Function CheckParentAssortIDExist(ID)

	Dim Rs
	If ID = 0 Then
		CheckParentAssortIDExist = 1
		Exit Function
	End If
	Set Rs = LDExeCute(sql_select("Select ID,OrderID,BoardID,AssortName,GoodNum from LeadBBS_GoodAssort where ID=" & ID,1),0)
	If Rs.Eof Then
		CheckParentAssortIDExist = 0
	Else
		LMT_OrderID = cCur(Rs("OrderID"))
		LMT_BoardID = cCur(Rs("BoardID"))
		Old_Board = LMT_BoardID
		LMT_AssortName = Rs("AssortName")
		LMT_GoodNum = cCur(Rs("GoodNum"))
		CheckParentAssortIDExist = 1
	End if
	Rs.Close
	Set Rs = Nothing

End Function

Function Manage

	%>
	<script language=javascript>
	var lastID=0,Count=0,oldBoardID;
	function s(ID,OrderID,BoardID,AssortName,GoodNum,BoardName,OrderID)
	{
		if(ID=="")return;
		if(oldBoardID != BoardID)
		{
			oldBoardID = BoardID;
			if(BoardID==0)BoardName="��ר����";
			document.write("<tr><td class=tdbox colspan=6><b>���棺<a href=<%=DEF_BBS_HomeUrl%>b/b.asp?B=" + BoardID + ">" + BoardName + "</a></td></tr>");
		}
		lastID=ID;
		document.write("<tr class=TBBG9><td class=tdbox>" + ID + "</td>");
		document.write("<td class=tdbox><a href=ForumBoardAssort.asp?action=Modify&ID=" + ID + ">" + AssortName + "</a></td>");
		document.write("<td class=tdbox>" + GoodNum + "</td><td class=tdbox><a href=<%=DEF_BBS_HomeUrl%>b/b.asp?B=" + BoardID + ">" + BoardName + "</a></td>");
		document.write("<td class=tdbox>" + OrderID + "</td>");
		document.write("<td class=tdbox><a href=ForumBoardAssort.asp?action=Delete&ID=" + ID + ">ɾ��</a></td></tr>");
	}
	</script>
	<table border=0 cellpadding=0 cellspacing=0 width="100%" class=frame_table>
			<tr class=frame_tbhead>
				<td width=46><div class=value>���</td>
				<td><div class=value>ר������(�޸�)</div></td>
				<td><div class=value>������</div></td>
				<td><div class=value>��������</div></td>
				<td><div class=value>˳��</div></td>
				<td><div class=value>ɾ��</div></td>
			</tr>
				<%
	Dim Rs,SQL
	SQL = "select T1.ID,T1.OrderID,T1.BoardID,T1.AssortName,T1.GoodNum,T2.BoardName,T1.OrderID from LeadBBS_GoodAssort as T1 left join LeadBBS_Boards as T2 on T1.BoardID=T2.BoardID Order by T1.BoardID,T1.OrderID"

	OpenDatabase
	Set Rs = LDExeCute(SQL,0)
	Dim Num
	Dim GetData
	If Not rs.Eof Then
		Response.Write "<script language=javascript>" & VbCrLf & "s('"
		Response.Write Rs.GetString(,,"','","');" & VbCrLf & "s('","")
		%>','','','');
		</script>
		<%
	Else
		Num = -1
	End If
	Rs.close
	Set Rs = Nothing
	closeDataBase%>
	</table>
	<%

End Function

Function Delete

	Dim ID
	ID = Left(Request("ID"),14)
	If isNumeric(ID) = 0 Then ID = 0
	ID = Fix(cCur(ID))
	If Request.Form("DeleteSuer")="E72ksiOkw2" Then
		If DeleteTopicAssort(ID) > 0 Then
			Response.Write "<p><font color=008800 class=greenfont><b>�Ѿ��ɹ�ɾ�����Ϊ" & ID & "�İ���ר����</b></font></p>"
		Else
			UpdateCacheData("data_goodassort.asp")
			Response.Write "<p><font color=ff0000 class=redfont><b>" & GBL_CHK_TempStr & "</b></font></p>"
		End If
	Else
		%><p><form action=ForumBoardAssort.asp method=post>
		ע�⣺ɾ������ר������ɾ��һ��ר���µ�������Ϣ<br>
		<br><b><font color=ff0000 class=redfont>ȷ����Ϣ�� ���Ҫɾ����ר����<br><br>
		
		<input type=hidden name=Action value="Delete">
		<input type=hidden name=ID value="<%=urlencode(ID)%>">
		<input type=hidden name=DeleteSuer value="E72ksiOkw2">

		<input type=submit value=ȷ��ɾ�� class=fmbtn>
		</form>
	<%End If

End Function

Function DeleteTopicAssort(ID)

	GBL_CHK_TempStr = ""
	Dim Rs,BoardID
	Set Rs = LDExeCute(sql_select("select ID,AssortName,BoardID from LeadBBS_GoodAssort where ID=" & ID,1),0)
	If Rs.Eof Then
		GBL_CHK_TempStr = "���󣬲����ڴ�ר������"
		DeleteTopicAssort = 0
		Rs.Close
		Set Rs = Nothing
		Exit Function
	Else
		BoardID = cCur(Rs(2))
	End If
	Rs.Close
	Set Rs = Nothing
	CALL LDExeCute("Update LeadBBS_Announce Set GoodAssort=0 where GoodAssort=" & ID,1)
	If DEF_UsedDataBase = 1 Then CALL LDExeCute("Update LeadBBS_Topic Set GoodAssort=0 where GoodAssort=" & ID,1)
	CALL LDExeCute("Delete from LeadBBS_GoodAssort where ID=" & ID,1)
	ReloadTopicAssort(BoardID)
	DeleteTopicAssort = 1

End Function

Sub ReloadTopicAssort(BoardID)

	Dim Rs
	Set Rs = LDExeCute("select ID,AssortName,0,0,0 from LeadBBS_GoodAssort where BoardID=" & BoardID & " Order by BoardID,OrderID ASC",0)
	If Not Rs.Eof Then
		Application.Lock
		Application(DEF_MasterCookies & "BoardInfo" & BoardID & "_TI") = Rs.GetRows(-1)
		Application.UnLock
	Else
		Application.Lock
		Set Application(DEF_MasterCookies & "BoardInfo" & BoardID & "_TI") = Nothing
		Application(DEF_MasterCookies & "BoardInfo" & BoardID & "_TI") = "yes"
		Application.UnLock
	End If
	Rs.Close
	Set Rs = Nothing

End Sub


Function UpdateCacheData(savefile)

		Dim Rs,GetData,Num
		Set Rs = LDExeCute("select T1.ID,T1.BoardID,T1.AssortName,T2.BoardName from LeadBBS_GoodAssort as T1 left join LeadBBS_Boards as T2 on T1.BoardID=T2.BoardID Order by T1.BoardID,T1.OrderID",0)
	
		If Not Rs.Eof Then
			GetData = Rs.GetRows(-1)
			Num = Ubound(GetData,2)
		Else
			Num = -1
		End If
		Rs.Close
		Set Rs = Nothing
		
		'on error resume next
		Dim TempStr
		TempStr = ""
	
		Dim N,WriteStr
		TempStr = TempStr & "["
	
		If Num = -1 Then
		Else
			dim oldBD,boardid
			oldbd = -1
			For N = 0 to Num
				boardid = ccur(getdata(1,n))
				if oldbd <> boardid then
					oldbd = boardid
					If N = 0 Then
						TempStr = TempStr & "{" & VbCrLf
					Else
						TempStr = TempStr & ",{" & VbCrLf
					End If
					TempStr = TempStr & "	""id"":0" & "," & VbCrLf
					If boardid=0 then getdata(3,n) = "��ר��"
					TempStr = TempStr & "	""text"":""�������:" & htmlencode(KillHTMLLabel(getdata(3,n))) & """" & VbCrLf & "}"
				end if
				WriteStr = ""
				WriteStr = WriteStr & KillHTMLLabel(GetData(2,N))
				If StrLength(WriteStr) > 21 Then
					WriteStr = LeftTrue(WriteStr,18) & "..."
				End If	
				
				TempStr = TempStr & ",{" & VbCrLf
				TempStr = TempStr & "	""id"":" & GetData(0,N) & "," & VbCrLf
				TempStr = TempStr & "	""text"":""" & GetData(0,N) & "." & htmlencode(WriteStr) & """" & VbCrLf & "}"
				'GBL_LowClassString = ""
				'GBL_LoopN = 0
				'GetLowClassString_Json GetData(0,n)
				'If GBL_LowClassString <> "" Then TempStr = TempStr & GBL_LowClassString				
			Next
		End If
	
		TempStr = DEF_pageHeader & TempStr & "]"
		
		ADODB_SaveToFile TempStr,DEF_BBS_HomeUrl & "inc/IncHtm/" & savefile & ""
		If GBL_CHK_TempStr = "" Then
			Response.Write "<br><span class=cms_ok>2.�ɹ������ļ�../../inc/IncHtm/" & savefile & "��</span>"
		Else
			%><p><%=GBL_CHK_TempStr%><br>��������֧������д���ļ����ܣ���ʹ��FTP�ȹ��ܣ�<br>��<span Class=cms_error>inc/IncHtm/<%=savefile%></span>�ļ��滻���¿�������(ע�ⱸ��)<p>
			<textarea name="fileContent" cols="80" rows="20" class=fmtxtra><%=Server.htmlencode(TempStr)%></textarea><%
			GBL_CHK_TempStr = ""
		End If
	
End Function%>