<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Board_Popfun.asp -->
<!-- #include file=../../inc/Upload_Setup.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<%
Server.ScriptTimeOut = 6000
DEF_BBS_HomeUrl = "../../"
Dim GBL_ID
initDatabase
GBL_CHK_TempStr = ""
GBL_ID = checkSupervisorPass

Manage_sitehead DEF_SiteNameString & " - ����Ա",""

Dim Form_UploadPhotoUrl_Old,Form_UploadPhotoUrl_Now
Form_UploadPhotoUrl_Old = "images/upload/"
Form_UploadPhotoUrl_Now = DEF_BBS_UploadPhotoUrl

If GBL_CHK_Flag = 1 Then
	GBL_CHK_TempStr = ""
	If Request("Form_UploadPhotoUrl_Old") <> "" Then
		Form_UploadPhotoUrl_Old = Replace(Trim(Request("Form_UploadPhotoUrl_Old")),"\","/")
		If Right(Form_UploadPhotoUrl_Old,1) <> "/" Then GBL_CHK_TempStr = "���󣬾��ϴ�·�����󣬱���ʹ��/��Ϊ·����β"
		Form_UploadPhotoUrl_Now = Replace(Trim(Request("Form_UploadPhotoUrl_Now")),"\","/")
		If Form_UploadPhotoUrl_Now = "" or Right(Form_UploadPhotoUrl_Now,1) <> "/" Then GBL_CHK_TempStr = "�������ϴ�·������"
		If Form_UploadPhotoUrl_Now = Form_UploadPhotoUrl_Old Then GBL_CHK_TempStr = "�¾�·��һ���������滻��"
		If StrLength(Form_UploadPhotoUrl_Now) > 150 or StrLength(Form_UploadPhotoUrl_Old) > 150 Then GBL_CHK_TempStr = "·�����������ܳ���150���ַ���"
	End If
	If Request("submitflag") = "yes" and GBL_CHK_TempStr = "" then
		If Request("Form_UploadPhotoUrl_Old") <> "" Then
			RepairUploadFaceUrl
		Else
			RepairSite
		End If
	Else
		If Request("Form_UploadPhotoUrl_Old") = "" Then
		%><form action=RepairSite.asp method=post>
			<div class=frametitle>1.Ĭ���޸�</div>
			<div class=frameline>ע�⣺�˹��ܽ�������¹��ܣ�</div>
			<div class=frameline>
				1.����ͳ��ÿ������(�������ذ���)����������<br>
				2.����ͳ����������Ա<br>
				3.����ͳ����̳ע���û�����<br>
				4.����ͳ����̳�ϴ���������<br>
				5.<span class=bluefont>�޸���������</span><br>
				6.<span class=bluefont>�޸�����ר����</span><br>
			</div>
			<input type="hidden" name="submitflag" value="yes">
			<div class=alert>ȷ����Ϣ�� ���Ҫ��ʼ�޸���������ô��</div>
			<div class=frameline>
			<input class=fmchkbox type="checkbox" name="repairFlag" value="yes" checked>ѡ�����Զ��޸�ÿ�������������������������鿴
			</div>
			<div class=frameline>
			<input type=submit value="�����ʼ�޸�" class=fmbtn>
			</div>
			
			<div class=frameline>
				<a href=../User/ClearOnlineUser.asp>�����Ҫ������е�������Ա��������������</a>
			</div>
			</form>
		<%End If
			If GBL_CHK_TempStr <> "" Then%>
			<div class=alert><%=GBL_CHK_TempStr%></div><%
			End If%>
			<form action=RepairSite.asp method=post>
			<div class=frametitle>2.�ϴ�·����Ϣ�޸�</div>
			<div class=frameline>
			�˹��ܽ�������¹��ܣ��޸��û����еı����ڱ�����վ��ͼƬ·��<br>
			����İ汾Ϊ3.14a��ɵİ汾�������������ı����ϴ�����·��ʱ��<br>���ܻ���Ҫ�˹��ܽ����޸�<br>
			</div>
			<input type="hidden" name="submitflag" value="yes">
			<div class=frameline>
			���ϴ�·����<input class=fminpt type="text" name="Form_UploadPhotoUrl_Old" maxlength="150" size="30" value="<%=htmlencode(Form_UploadPhotoUrl_Old)%>">
			</div>
			<div class=frameline>
			���ϴ�·����<input class=fminpt type="text" name="Form_UploadPhotoUrl_Now" maxlength="150" size="30" value="<%=htmlencode(Form_UploadPhotoUrl_Now)%>">
			</div>
			<div class=frameline><span class=note>ע��·��ָ�����������̳��Ŀ¼��·����Ĭ�ϴ����images/upload����</span>
			</div>
			<div class=alert>���棺�޸�ʱ����ܽϳ��Ҳ����棬���ȷ������д����Ϣ��ȷ����</div>
			<div class=frameline><input class=fmchkbox type="checkbox" name="repairAnnounce" value="yes" checked>ѡ��������޸����������е�ͼƬ·��</div>
			<div class=frameline><input class=fmchkbox type="checkbox" name="repairUserUnderWrite" value="yes" checked>ѡ��������޸��û�ǩ���е�ͼƬ·��</div>
			
			<div class=frameline>
			<input type=submit value="�����ʼ�޸�" class=fmbtn>
			</div>
			<div class=frameline>
				<a href=../User/ClearOnlineUser.asp>�����Ҫ������е�������Ա��������������</a>
			</div>
			</form>
		<%
	End If
Else
	DisplayLoginForm
End If
closeDataBase
Manage_Sitebottom("none")


Function RepairSite

	Dim repairFlag
	repairFlag = Request("repairFlag")
	If repairFlag <> "yes" Then repairFlag = ""
	Dim Rs
	Dim UploadNum,UserCount
	Response.Write "<br>"
	Set Rs = LDExeCute("select count(*) from LeadBBS_User",0)
	If Rs.Eof Then
		UserCount = 0
	Else
		UserCount = Rs(0)
		If isNull(UserCount) Then UserCount = 0
		UserCount = cCur(UserCount)
	End If
	Rs.Close
	Set Rs = Nothing

	Set Rs = LDExeCute("select count(*) from LeadBBS_Upload",0)
	If Rs.Eof Then
		UploadNum = 0
	Else
		UploadNum = Rs(0)
		If isNull(UploadNum) Then UploadNum = 0
		UploadNum = cCur(UploadNum)
	End If
	Rs.Close
	Set Rs = Nothing

	CALL LDExeCute("Update LeadBBS_SiteInfo Set UserCount=" & UserCount & ",UploadNum=" & UploadNum,1)
	ReloadStatisticData

	Response.Write "<br>ע���û��������ϴ��ļ���������ͳ����ɣ�"
	SetActiveUserCount
	Response.Write "<br>��̳��������������ͳ����ɣ�"

	Dim GetData
	Set Rs = LDExeCute("Select BoardID from LeadBBS_Boards",0)
	If Not Rs.Eof Then
		GetData = Rs.GetRows(-1)
		Rs.Close
		Set Rs = Nothing
	Else
		Rs.Close
		Set Rs = Nothing
		Exit Function
	End If
	
	Dim N,m,i
	m = Ubound(GetData,2)
	For N = 0 to m
		Set Rs = LDExeCute("Select count(*) from LeadBBS_OnlineUser Where AtBoardID=" & GetData(0,n),0)
		If Rs.Eof Then
			i = 0
		Else
			i = Rs(0)
			If isNull(i) Then i = 0
			i = cCur(i)
		End If
		Rs.Close
		Set Rs = Nothing
		ReloadBoardInfo(GetData(0,n))
		ReloadTopicAssort(GetData(0,n))
		Response.Write "<br>�����" & GetData(0,n) & "����������ԭ��" & Application(DEF_MasterCookies & "BDOL" & GetData(0,n)) & "�ˣ�ʵ������" & i & "��"
		If repairFlag = "yes" then
			Application.Lock
			Application(DEF_MasterCookies & "BDOL" & GetData(0,n)) = i
			Application.UnLock
		End If
	Next
	Response.Write "<p>����ͳ�ư�������������ɣ���"
	ReloadPubMessageInfo
	Response.Write "<p>�޸�����������ɣ�"
	If repairFlag <> "yes" then Response.Write "<font color=Red Class=redfont>����û��������ɰ������������ĸ��£�</font>"

End Function

Sub ReloadTopicAssort(BoardID)

	Dim Rs
	Set Rs = LDExeCute("select ID,AssortName,0,0,0 from LeadBBS_GoodAssort where BoardID=" & BoardID & " Order by BoardID,OrderID",0)
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

Sub RepairUploadFaceUrl

	Dim MyHomeUrl,ReplaceUrl1,ReplaceUrl2,Len1,Len2
	MyHomeUrl = LCase(Request.Servervariables("SCRIPT_NAME"))
	If Request.ServerVariables("SERVER_PORT") <> "80" Then MyHomeUrl = ":" & Request.ServerVariables("SERVER_PORT") & MyHomeUrl
	MyHomeUrl = Lcase("http://"&Request.ServerVariables("server_name") & MyHomeUrl)
	MyHomeUrl = Replace(MyHomeUrl,LCase(DEF_ManageDir) & "/sitemanage/repairsite.asp","")	
	
	ReplaceUrl1 = MyHomeUrl & Form_UploadPhotoUrl_Old
	ReplaceUrl2 = Replace("../" & Form_UploadPhotoUrl_Old,"//","/")
	Len1 = Len(ReplaceUrl1)
	Len2 = Len(ReplaceUrl2)
	
	Form_UploadPhotoUrl_Now = Replace("../" & Form_UploadPhotoUrl_Now,"//","/")

	Dim UpdateNumber,NoneUpdateNumber
	UpdateNumber = 0
	NoneUpdateNumber = 0
	
	Dim NowID,EndFlag,Temp
	NowID = 0
	EndFlag = 0
	Dim Rs,SQL,GetData,n
	
	Dim RecordCount,CountIndex
	SQL = "Select count(*) from LeadBBS_User where id>" & NowID
	Set Rs = LDExeCute(SQL,0)
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
	
	Application.Lock
	Application("Io_" & GBL_CHK_User) = "start"
	Application.UnLock
	If Request("executepage") = "" Then
	%>
	<div id="errorstr"></div>
	<p style="font-size:9pt" id="bartitle1">���濪ʼ�޸��û��ϴ�ͷ��·��������<%=RecordCount%>���û�������

		<table width="400" border="0" cellspacing="1" cellpadding="1">
			<tr> 
				<td bgcolor=000000>
		<table width="400" border="0" cellspacing="0" cellpadding="1">
			<tr> 
				<td bgcolor=ffffff height=9><img src=../pic/progressbar.gif width=0 height=16 id=img1 name=img1 align=middle></td></tr></table>
		</td></tr></table> <span id=txt1 name=txt1 style="font-size:9pt">0</span><span style="font-size:9pt">%</span>
		<span id=tm1 name=tm1 style="font-size:9pt">���ڹ�����Ҫʱ��...</span>
		<script src="<%=DEF_BBS_HomeUrl%>inc/js/bar.js?ver=<%=DEF_Jer%>" type="text/javascript"></script>
		<script>
			Upl_url = "../BlockUpdate/Io_Info.asp?id=<%=Urlencode(GBL_CHK_User)%>";
			Upl_IOfun = window.setTimeout(Upl_IO,Upl_GetDelay);
		</script>
		<br>
		<iframe src="RepairSite.asp?executepage=yes&submitflag=yes&Form_UploadPhotoUrl_Old=<%=urlencode(Form_UploadPhotoUrl_Old)%>&Form_UploadPhotoUrl_Now=<%=urlencode(Form_UploadPhotoUrl_Now)%>&repairAnnounce=<%=urlencode(Request("repairAnnounce"))%>&repairUserUnderWrite=<%=urlencode(Request("repairUserUnderWrite"))%>" name="infoframe" id="infoframe" hidefocus="" frameborder="no" scrolling="auto" style="margin-top:100px;width:300px;height:150px;">
		<%
			response.Flush
			Exit sub
	end if
	Dim StartTime,SpendTime,RemainTime
	StartTime = Now
		
	Do while EndFlag = 0
		SQL = sql_select("Select ID,FaceUrl from LeadBBS_User where ID>" & NowID & " order by id ASC",1000)
		Set Rs = LDExeCute(SQL,0)
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
			NowID = GetData(0,n)
			GetData(1,n) = LCase(GetData(1,n))
			If Left(GetData(1,n) & "",Len2) = ReplaceUrl2 Then
				GetData(1,n) = replace(GetData(1,n),ReplaceUrl2,Form_UploadPhotoUrl_Now,1,1,0)
				CALL LDExeCute("Update LeadBBS_User Set FaceUrl='" & Replace(GetData(1,n),"'","''") & "' where ID=" & NowID,1)
			ElseIf Left(GetData(1,n) & "",Len1) = ReplaceUrl1 Then
				GetData(1,n) = replace(GetData(1,n),ReplaceUrl1,Form_UploadPhotoUrl_Now,1,1,0)
				CALL LDExeCute("Update LeadBBS_User Set FaceUrl='" & Replace(GetData(1,n),"'","''") & "' where ID=" & NowID,1)
			Else
				NoneUpdateNumber = NoneUpdateNumber + 1
			End If
			CountIndex = CountIndex + 1
			If (CountIndex mod 100) = 0 Then
				SpendTime = Datediff("s",StartTime,Now)
				RemainTime = SpendTime/CountIndex * (RecordCount-CountIndex)
				Application.Lock
				Application("Io_" & GBL_CHK_User) = Fix((CountIndex/RecordCount) * 400) & "|" & FormatNumber(CountIndex/RecordCount*100,4,-1) & "|" & SpendTime & "|" & RemainTime & "|" & CountIndex
				Application.UnLock
			End If
		Next
		If Response.IsClientConnected Then
		Else
			EndFlag = 1
			Application.Contents.Remove("Io_" & GBL_CHK_User)
		End If
	Loop
	%>���
	������<%=UpdateNumber%>���û���<%=NoneUpdateNumber%>���û��������
	<%
	If Request("repairAnnounce") = "yes" Then RepairAnnounceUploadUrl
	If Request("repairUserUnderWrite") = "yes" Then RepairUserUnderWriteUploadUrl
	Application.Contents.Remove("Io_" & GBL_CHK_User)

End Sub

Sub RepairAnnounceUploadUrl

	Dim MyHomeUrl,ReplaceUrl1,ReplaceUrl2,Len1,Len2
	MyHomeUrl = LCase(Request.Servervariables("SCRIPT_NAME"))
	If Request.ServerVariables("SERVER_PORT") <> "80" Then MyHomeUrl = ":" & Request.ServerVariables("SERVER_PORT") & MyHomeUrl
	MyHomeUrl = Lcase("http://"&Request.ServerVariables("server_name") & MyHomeUrl)
	MyHomeUrl = Replace(MyHomeUrl,LCase(DEF_ManageDir) & "/sitemanage/repairsite.asp","")	
	
	ReplaceUrl1 = MyHomeUrl & Form_UploadPhotoUrl_Old
	ReplaceUrl2 = Replace("../" & Form_UploadPhotoUrl_Old,"//","/")
	Len1 = Len(ReplaceUrl1)
	Len2 = Len(ReplaceUrl2)

	Dim UpdateNumber,NoneUpdateNumber
	UpdateNumber = 0
	NoneUpdateNumber = 0

	Dim NowID,EndFlag,Temp
	NowID = 0
	EndFlag = 0
	Dim Rs,SQL,GetData,n
	
	Dim RecordCount,CountIndex
	SQL = "Select count(*) from LeadBBS_Announce where id>" & NowID
	Set Rs = LDExeCute(SQL,0)
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

	Dim StartTime,SpendTime,RemainTime
	StartTime = Now
	
	response.Flush
	SpendTime = Datediff("s",StartTime,Now)
	RemainTime = RecordCount
	
	dim titlestr
	titlestr = "|���濪ʼ�޸������е�ȫ���ϴ�Ŀ¼�µ�ͼƬ·��������" & RecordCount & "�����Ӵ�����"
	Application.Lock
	Application("Io_" & GBL_CHK_User) = "1|0|���ڹ���ʱ��...|start"	
	Application("Io_" & GBL_CHK_User) = "0|0|" & SpendTime & "|" & RemainTime & "|" & CountIndex & titlestr
	Application.UnLock

	Dim UpdateFlag
	Dim GetData1

	dim re
	set re = New RegExp
	re.Global = True
	re.IgnoreCase = True
	
	Do while EndFlag = 0
		SQL = sql_select("Select ID,Content,HTMLFlag from LeadBBS_Announce where ID>" & NowID & " order by id ASC",100)
		Set Rs = LDExeCute(SQL,0)
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
			NowID = GetData(0,n)
			GetData1 = GetData(1,n)
			If GetData(2,n) <> 0 Then
				UpdateFlag = 0
				SQL = "Update LeadBBS_Announce Set"
				
				re.Pattern="(" & Replace(ReplaceUrl2,".","\.") & ")"
				GetData(1,n)=re.Replace(GetData(1,n),Form_UploadPhotoUrl_Now)
				
				re.Pattern="(" & Replace(ReplaceUrl1,".","\.") & ")"
				GetData(1,n)=re.Replace(GetData(1,n),Form_UploadPhotoUrl_Now)
				
				If GetData(1,n) <> GetData1 Then
					UpdateFlag = 1
					SQL = SQL & " Content='" & Replace(GetData(1,n),"'","''") & "'"
				End If
	
				SQL = SQL & " where ID=" & NowID
				If UpdateFlag = 1 Then
					CALL LDExeCute(SQL,0)
					UpdateNumber = UpdateNumber + 1
				Else
					NoneUpdateNumber = NoneUpdateNumber + 1
				End If
			Else
				NoneUpdateNumber = NoneUpdateNumber + 1
			End If
			CountIndex = CountIndex + 1
			If (CountIndex mod 100) = 0 Then
				SpendTime = Datediff("s",StartTime,Now)
				RemainTime = SpendTime/CountIndex * (RecordCount-CountIndex)
				Application.Lock
				Application("Io_" & GBL_CHK_User) = Fix((CountIndex/RecordCount) * 400) & "|" & FormatNumber(CountIndex/RecordCount*100,4,-1) & "|" & SpendTime & "|" & RemainTime & "|" & CountIndex & titlestr
				Application.UnLock
			End If
		Next
		If Response.IsClientConnected Then
		Else
			EndFlag = 1
			Application.Contents.Remove("Io_" & GBL_CHK_User)
		End If
	Loop
	Set Re = Nothing
	%>���
	������<%=UpdateNumber%>�����ӣ�<%=NoneUpdateNumber%>�������������
	<%
	

End Sub


Sub RepairUserUnderWriteUploadUrl

	Dim MyHomeUrl,ReplaceUrl1,ReplaceUrl2,Len1,Len2
	MyHomeUrl = LCase(Request.Servervariables("SCRIPT_NAME"))
	If Request.ServerVariables("SERVER_PORT") <> "80" Then MyHomeUrl = ":" & Request.ServerVariables("SERVER_PORT") & MyHomeUrl
	MyHomeUrl = Lcase("http://"&Request.ServerVariables("server_name") & MyHomeUrl)
	MyHomeUrl = Replace(MyHomeUrl,LCase(DEF_ManageDir) & "/sitemanage/repairsite.asp","")	
	
	ReplaceUrl1 = MyHomeUrl & Form_UploadPhotoUrl_Old
	ReplaceUrl2 = Replace("../" & Form_UploadPhotoUrl_Old,"//","/")
	Len1 = Len(ReplaceUrl1)
	Len2 = Len(ReplaceUrl2)

	Dim UpdateNumber,NoneUpdateNumber
	UpdateNumber = 0
	NoneUpdateNumber = 0

	Dim NowID,EndFlag,Temp
	NowID = 0
	EndFlag = 0
	Dim Rs,SQL,GetData,n
	
	Dim RecordCount,CountIndex
	SQL = "Select count(*) from LeadBBS_User where id>" & NowID
	Set Rs = LDExeCute(SQL,0)
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


	Dim StartTime,SpendTime,RemainTime
	StartTime = Now
	
	response.Flush
	dim titlestr: titlestr = "|���濪ʼ�޸��û�ǩ���е�ȫ���ϴ�Ŀ¼�µ�ͼƬ·��������" & RecordCount & "���û�ǩ��������"
	
	SpendTime = Datediff("s",StartTime,Now)
	RemainTime = RecordCount
	Application.Lock
	Application("Io_" & GBL_CHK_User) = "1|0|���ڹ���ʱ��...|start"
	Application("Io_" & GBL_CHK_User) = "0|0|" & SpendTime & "|" & RemainTime & "|" & CountIndex & titlestr
	Application.UnLock
	
	Dim UpdateFlag,UpdateFlag2
	Dim GetData1,GetData2
	
	
	dim re
	set re = New RegExp
	re.Global = True
	re.IgnoreCase = True	
	Do while EndFlag = 0
		SQL = sql_select("Select ID,Underwrite,PrintUnderWrite from LeadBBS_User where ID>" & NowID & " order by id ASC",100)
		Set Rs = LDExeCute(SQL,0)
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
			NowID = GetData(0,n)
			GetData1 = GetData(1,n)
			GetData2 = GetData(2,n)
			UpdateFlag = 0
			UpdateFlag2 = 0
			SQL = "Update LeadBBS_User Set"
			
			re.Pattern="(" & Replace(ReplaceUrl2,".","\.") & ")"
			GetData(1,n)=re.Replace(GetData(1,n),Form_UploadPhotoUrl_Now)
			
			re.Pattern="(" & Replace(ReplaceUrl1,".","\.") & ")"
			GetData(1,n)=re.Replace(GetData(1,n),Form_UploadPhotoUrl_Now)
			
			If GetData(1,n) <> GetData1 Then
				UpdateFlag = 1
				SQL = SQL & " UnderWrite='" & Replace(GetData(1,n),"'","''") & "'"
			End If
			
			
			re.Pattern="(" & Replace(ReplaceUrl2,".","\.") & ")"
			GetData(2,n)=re.Replace(GetData(2,n),Form_UploadPhotoUrl_Now)
			
			re.Pattern="(" & Replace(ReplaceUrl1,".","\.") & ")"
			GetData(2,n)=re.Replace(GetData(2,n),Form_UploadPhotoUrl_Now)
			
			If GetData(2,n) <> GetData2 Then UpdateFlag2 = 1
			
			If UpdateFlag2 = 1 Then
				If UpdateFlag = 0 Then
					SQL = SQL & " PrintUnderWrite='" & Replace(GetData(2,n),"'","''") & "'"
					UpdateFlag = 1
				Else
					SQL = SQL & ",PrintUnderWrite='" & Replace(GetData(2,n),"'","''") & "'"
				End If
			End If
			
			SQL = SQL & " where ID=" & NowID
			
			If StrLength(GetData(2,n)) > 1024 or StrLength(GetData(1,n)) > 255 Then
			Else
				If UpdateFlag = 1 Then
					CALL LDExeCute(SQL,1)
					UpdateNumber = UpdateNumber + 1
				Else
					NoneUpdateNumber = NoneUpdateNumber + 1
				End If
			End If
			CountIndex = CountIndex + 1
			If (CountIndex mod 100) = 0 Then
				SpendTime = Datediff("s",StartTime,Now)
				RemainTime = SpendTime/CountIndex * (RecordCount-CountIndex)
				Application.Lock
				Application("Io_" & GBL_CHK_User) = Fix((CountIndex/RecordCount) * 400) & "|" & FormatNumber(CountIndex/RecordCount*100,4,-1) & "|" & SpendTime & "|" & RemainTime & "|" & CountIndex & titlestr
				Application.UnLock
			End If
		Next
		If Response.IsClientConnected Then
		Else
			EndFlag = 1
			Application.Contents.Remove("Io_" & GBL_CHK_User)
		End If
	Loop
	Set Re = Nothing
	%>���
	������<%=UpdateNumber%>��ǩ����<%=NoneUpdateNumber%>���û�ǩ���������
	<%
	
	

End Sub
%>