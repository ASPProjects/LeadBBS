<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Board_Popfun.asp -->
<!-- #include file=../../inc/Ubbcode.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<%
Rem -------------------------------------------------------
Rem ------------�˳��������޸���LeadBBS_Announce-----------
Rem ------------��RootMaxID��RootMinID�ֶ�-----------------
Rem ------------����ʱ�������������ȵ���̨�ر���̳---------
Rem -------------------------------------------------------

DEF_BBS_HomeUrl = "../../"
server.scripttimeout=99999
initDatabase
UpdateLastInfoColumn
CloseDatabase

Function UpdateLastInfoColumn()
	
	If CheckSupervisorUserName = 0 or GBL_UserID = 0 Then Exit Function

	If Request.Form("SureFlag") <> "E72ksiOkw2" Then
		%>
			<p><form action=RepairLastInfo.asp method=post>
			<b><font color=ff0000 class=RedFont>ȷ���޸�LastInfo������?<br>
			<br>
			<input type=hidden name=SureFlag value="E72ksiOkw2">
			
			<input type=submit value=ȷ������ class=fmbtn>
			</form>
		<%
	Else
		Dim StartTime,SpendTime,RemainTime
		Dim TempAT,T
		Dim NowID,EndFlag,SQL,Rs
		NowID = 0
		EndFlag = 0
		
		Dim RootMaxID,RootMinID,ChildNum
		Dim N,GetData
	
		Dim RecordCount,CountIndex
		SQL = "Select count(*) from LeadBBS_Announce where parentID=0"
		Set Rs = Con.ExeCute(SQL)
		If Rs.Eof Then
			RecordCount = 0
		Else
			RecordCount = Rs(0)
			If isNull(RecordCount) Then RecordCount = 0
			RecordCount = ccur(RecordCount)
		End If
		Rs.Close
		Set Rs = Nothing
		If RecordCount < 1 Then RecordCount = 1
		CountIndex = 0
		%>
		<p style="font-size:9pt">���濪ʼ�޸���̳������Ϣ������<%=RecordCount%>�����������
	
		<table width="400" border="0" cellspacing="1" cellpadding="1">
			<tr> 
				<td bgcolor=000000>
		<table width="400" border="0" cellspacing="0" cellpadding="1">
			<tr> 
				<td bgcolor=ffffff height=9><img src=../../images/vote.gif width=0 height=16 id=img1 name=img1 align=absmiddle></td></tr></table>
		</td></tr></table> <span id=txt1 name=txt1 style="font-size:9pt">0</span><span style="font-size:9pt">%</span>
		<span id=tm name=tm style="font-size:9pt">���ڹ�����Ҫʱ��...</span>
		<%
		Response.Flush
		StartTime = Now
		Do while EndFlag = 0
			SQL = sql_select("Select ID,LastInfo from LeadBBS_Announce where ParentID=0 and RootIDBak>" & NowID & " order by ID ASC",100)
			Set Rs = Con.ExeCute(SQL)
			GBL_DBNum = GBL_DBNum + 1
			If Rs.Eof Then
				EndFlag = 1
				Rs.Close
				Set Rs = Nothing
				Exit Do
			Else
				GetData = Rs.GetRows(-1)
				Rs.Close
				Set Rs = Nothing
			End If
			For N = 0 to Ubound(GetData,2)
				If StrLength(GetData(1,n)) > 99 Then
					GetData(1,n) = LeftTrue(GetData(1,n),99)
					Con.Execute("Update LeadBBS_Announce set LastInfo='" & Replace(GetData(1,n),"'","''") & "' where ID=" & GetData(0,n))
				End If
				NowID = GetData(0,n)
	
				CountIndex = CountIndex + 1
				If (CountIndex mod 100) = 0 Then
					Response.Write "<script>img1.width=" & Fix((CountIndex/RecordCount) * 400) & ";" & VbCrLf
					Response.Write "txt1.innerHTML=""" & FormatNumber(CountIndex/RecordCount*100,4,-1) & """;" & VbCrLf
					If CountIndex > 300 Then
						SpendTime = Datediff("s",StartTime,Now)
						RemainTime = SpendTime/CountIndex * (RecordCount-CountIndex)
						Response.Write "tm.innerHTML=""" & "��ǰ����:" & GetTimeString(SpendTime) & " ����ʣ��:" & GetTimeString(RemainTime) & """;" & VbCrLf
					End If
					Response.Write "img1.title=""(" & CountIndex & ")"";</script>" & VbCrLf
					Response.Flush
				End If
			Next
		Loop
		%>
		<script>img1.width=400;
		txt1.innerHTML="100";</script>���
		<%
	End If

End Function

Function GetTimeString(Num)

	Dim Str,Temp,Number
	Number = Num
	Temp = Number/(24*60*60)
	If Fix(Temp) > 0 Then Str = Str & Fix(Temp) & "��"
	Number = Number-Fix(Temp)*24*60*60
	Temp = Number/(60*60)
	If Fix(Temp) > 0 Then Str = Str & Fix(Temp) & "ʱ"
	Number = Number-Fix(Temp)*60*60
	Temp = Number/(60)
	If Fix(Temp) > 0 Then Str = Str & Fix(Temp) & "��"
	Number = Number-Fix(Temp)*60
	Temp = Fix(Number)
	If Fix(Temp) > 0 Then Str = Str & Temp & "��"
	GetTimeString = Str

End Function
%>