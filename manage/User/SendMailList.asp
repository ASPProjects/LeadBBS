<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Board_popfun.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<!-- #include file=../../User/inc/Mail_fun.asp -->
<%
server.scripttimeout=99999
DEF_BBS_HomeUrl = "../../"
initDatabase
GBL_CHK_TempStr = ""
checkSupervisorPass

Manage_sitehead DEF_SiteNameString & " - ����Ա",""
frame_TopInfo
DisplayUserNavigate("�ʼ�")

Dim Email,Topic,MailBody,MailList

If GBL_CHK_Flag = 1 and GBL_CHK_TempStr = "" Then
	If Request.QueryString = "MailList" Then
		GetMailList
	Else
		SendMailListForm
		frame_BottomInfo
	End If
Else
	DisplayLoginForm
	frame_BottomInfo
	
End If
closeDataBase

Function SendMailListForm

	Dim Rs,Num
	Set Rs = LDExeCute("Select count(*) from LeadBBS_User",0)
	If Not Rs.Eof Then
		Num = Rs(0)
		If isNull(Num) Then Num = 0
		Num = cCur(Num)
	Else
		Num = 0
	End If
	Rs.Close
	Set Rs = Nothing
	
	Dim submitflag
	MailBody = Request.Form("MailBody")
	Email = Request.Form("Email")
	Topic = Request.Form("Topic")
	MailList = Request.Form("MailList")
	submitflag = Left(Request.Form("submitflag"),10)
	
	If submitflag = "1" Then
		If Email = "" or Len(Email) > 150 or inStr(Email,"@") = 0 Then
			Response.Write "<br><br><b><font color=red>�����ַ��������ܳ���150���֣�</font></b>" & VbCrLf
			submitflag = ""
		End If
		If Topic = "" or Len(Topic) > 250 Then
			Response.Write "<br><br><b><font color=red>�ʼ����������д���Ҳ��ܳ���250���֣�</font></b>" & VbCrLf
			submitflag = ""
		End If
		If MailBody = "" or Len(MailBody) > 65535 Then
			Response.Write "<br><br><b><font color=red>�����ʼ����ݱ�����д���Ҳ��ܳ���65535���֣�</font></b>" & VbCrLf
			submitflag = ""
		End If
	End If

	If submitflag = "1" then
		submitflag = 2
		%><form action=SendMailList.asp id=fm1 name=fm1 method=post><p style="font-size:9pt"><br>
			<b style="font-size:9pt"><font color=ff0000 class=redfont>Ⱥ���ʼ��������£������һ����ť������ʼ���ͣ��ڶ�����ť���·��ر༭</font></b><br><br>
			<p style="font-size:9pt">
			<b>���������б�</b><%
					If MailList = "" Then
						Response.Write "�����û�"
					Else
						Response.Write htmlencode(MailList)
					End If%><br><br>
			<b>���������б�</b><%=htmlencode(MailList)%><br><br>
			<b>����ʹ�����䣺</b><%=htmlencode(Email)%><br><br>
			<b>�����ʼ����⣺</b><%=htmlencode(Topic)%><br><br>
			<b>�����ʼ����ݣ�</b><br><br><br><%=MailBody%><br><br>
			<input name=MailList maxlength=224 size=54 value="<%=htmlencode(MailList)%>" class=fminpt type=hidden><br>
			<input name=Email maxlength=224 size=54 value="<%=htmlencode(Email)%>" class=fminpt type=hidden><br>
			<input name=Topic maxlength=224 size=54 value="<%=htmlencode(Topic)%>" class=fminpt type=hidden><br>
			<input name=MailBody value="<%If MailBody <> "" Then Response.Write VbCrLf & htmlEncode(MailBody)%>" type=hidden>
			<input name=submitflag value="<%=submitflag%>" type=hidden><br><br>
			<input type=button value="������￪ʼ����" class=fmbtn onclick="javascript:document.all.fm1.submitflag.value=2;document.all.fm1.submit();" class=fmbtn>
			<input type=button value="������ﷵ�ر༭" class=fmbtn onclick="javascript:document.all.fm1.submitflag.value=0;document.all.fm1.submit();" class=fmbtn>
			</form>
		<%
		frame_BottomInfo
		
	ElseIf submitflag = "2" then
		If DEF_BBS_EmailMode < 1 and DEF_BBS_EmailMode > 3 Then
			Response.Write "������̳��֧���ʼ����ͣ�"
			frame_BottomInfo
			
		Else
			If MailList = "" Then
				SendMailList
			Else
				SendMailList2
			End If
		End If
	Else
		submitflag = 1
		%><form action=SendMailList.asp id=fm1 name=fm1 style="font-size:9pt" method=post><br>
			<div class=alert>ע�⣺</div>
			<ol class=listli>
			<li>���ڿ�����д�ʼ����ݣ��ʼ����ݱ���ΪHTML���룬һЩͼƬ�������ļ���ʹ�����ӣ�</li>
			<li>���ͳ���Ҫ��һ������ȫִ����ϣ�������ִ�����ϵ�ٷ���Ա</li>
			<li>����<%=Num%>���ͻ������ͣ�����ʱ���п��ܷǳ�������������벻Ҫˢ���ظ�ִ��</li>
			<li>��ʼ����ʱ�н���������ʾ��ǰ�ķ��ͽ���</li>
			<li>�ύ����ʾ��ǰ�ʼ���ʽ����Ҫ�ٴ�ȷ�Ϻ���ܿ�ʼ����</li>
			<li>ע���ʼ��ı�������ʾ�Է���˾���ƣ����ڶ��ź���ʾ����д�ı���</li>
			<li>����д�����˱�ʾ���͸����е��û�</li>
			</ol>
			<div class=frametitle>��д�ʼ���Ϣ</div>
			<table border=0 cellpadding=0 cellspacing=0 width="100%" class=frame_table>
			<tr>
			<td class=tdbox width=120>���������б�</td><td class=tdbox><input name=MailList maxlength=2224 size=44 value="<%=htmlencode(MailList)%>" class=fminpt> <span class=note>���ŷָ� �����򷢸������û�</span></td>
			<tr><td class=tdbox>����ʹ�����䣺</td><td class=tdbox><input name=Email maxlength=224 size=44 value="<%=htmlencode(Email)%>" class=fminpt> <span class=note>�˻ػ���ʹ�õ����� ��һ����Ч</span>
			<tr><td class=tdbox>�����ʼ����⣺</td><td class=tdbox><input name=Topic maxlength=224 size=44 value="<%=htmlencode(Topic)%>" class=fminpt> <span class=note>�ڱ���ǰ�Զ������û������ö��ŷָ�</span>
			<tr><td class=tdbox>�����ʼ����ݣ�<br><span class=note>����ʹ��HTML<br>����</span></td><td class=tdbox>
			<textarea cols=53 name=MailBody rows=16 class=fmtxtra><%If MailBody <> "" Then Response.Write VbCrLf & Server.htmlEncode(MailBody)%></textarea>
			<input name=submitflag value="<%=submitflag%>" type=hidden><br><br>
			<input type=submit value="��һ�� &lt;&lt; �����ʼԤ���ʼ�����" onclick="javascript:document.all.fm1.submit();" class=fmbtn>
			</td></tr></table></form>
			<br />
			<div class=frameline><a href=SendMailList.asp?MailList><b><span class=bluefont>�����ȡ�ʼ��б�</span></b></a></div>
			<div class=frameline>
			��ȡ�ʼ��б����û�����ͬ����Ҫ��ͬ��ʱ�䡣���û�������1000ʱ��������Դ�ܴ�Ϊά���������ȶ����м��������������û������棬<span class=redfont>������ʹ�ô����</span>��
			����û������󣬽�����������ݣ�����ʹ�����Ϊ���ش���ҳ��
			</div>
			
		<%
		frame_BottomInfo
		
	End If

End Function

Function SendMailList

	'Response.Clear
	Dim TempAT,T
	Dim NowID,EndFlag,SQL,Rs
	NowID = 0
	EndFlag = 0

	Dim RootMaxID,RootMinID,ChildNum
	Dim N,GetData

	Dim RecordCount,CountIndex
	SQL = "Select count(*) from LeadBBS_User"
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
	%>
	<p style="font-size:9pt">���濪ʼȺ���ʼ�������<%=RecordCount%>����ַ������

	<table width="400" border="0" cellspacing="1" cellpadding="1">
		<tr> 
			<td bgcolor=000000>
	<table width="400" border="0" cellspacing="0" cellpadding="1">
		<tr> 
			<td bgcolor=ffffff height=9><img src=../../images/vote.gif width=0 height=16 id=img1 name=img1 align=middle></td></tr></table>
	</td></tr></table> <span id=txt1 name=txt1 style="font-size:9pt">0</span><span style="font-size:9pt">%</span>
	<script>window.scroll(0,65535);</script>
	<%
	frame_BottomInfo
	
	Response.Flush
	'on error resume next
	Do while EndFlag = 0
		SQL = sql_select("Select ID,mail,UserName from LeadBBS_User where ID>" & NowID & " order by ID ASC",1000)
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
			CountIndex = CountIndex + 1
			If (CountIndex mod 20) = 0 Then
				Response.Write "<script>img1.width=" & Fix((CountIndex/RecordCount) * 400) & ";" & VbCrLf
				Response.Write "txt1.innerHTML=""" & FormatNumber(CountIndex/RecordCount*100,4,-1) & """;" & VbCrLf
				Response.Write "img1.title=""(" & CountIndex & ")"";</script>" & VbCrLf
				Response.Flush
			End If
			If inStr(GetData(1,n) & "","@") > 0 Then
				If GetData(2,n) <> "" Then GetData(2,n) = GetData(2,n) & ","
				SendMail GetData(1,n),GetData(2,n) & Topic
			End If
		Next
	Loop
	%>
	<script>img1.width=400;
	txt1.innerHTML="100";</script>
	<%

End Function


Function SendMailList2

	'Response.Clear
	Dim TempAT,T
	Dim NowID,EndFlag,SQL,Rs
	NowID = 0
	EndFlag = 0
	
	Dim RootMaxID,RootMinID,ChildNum
	Dim N,GetData

	Dim RecordCount,CountIndex
	
	GetData = Split(MailList,",")
	RecordCount = Ubound(GetData,1) + 1
	If RecordCount < 1 Then RecordCount = 1
	CountIndex = 0
	%>
	<p style="font-size:9pt">���濪ʼȺ���ʼ�������<%=RecordCount%>����ַ������

	<table width="400" border="0" cellspacing="1" cellpadding="1">
		<tr> 
			<td bgcolor=000000>
	<table width="400" border="0" cellspacing="0" cellpadding="1">
		<tr> 
			<td bgcolor=ffffff height=9><img src=../../images/vote.gif width=0 height=16 id=img1 name=img1 align=middle></td></tr></table>
	</td></tr></table> <span id=txt1 name=txt1 style="font-size:9pt">0</span><span style="font-size:9pt">%</span>
	<script>window.scroll(0,65535);</script>
	<%
	frame_BottomInfo
	
	Response.Flush
	For N = 0 to RecordCount - 1
		NowID = N
		CountIndex = CountIndex + 1
		If (CountIndex mod 20) = 0 Then
			Response.Write "<script>img1.width=" & Fix((CountIndex/RecordCount) * 400) & ";" & VbCrLf
			Response.Write "txt1.innerHTML=""" & FormatNumber(CountIndex/RecordCount*100,4,-1) & """;" & VbCrLf
			Response.Write "img1.title=""(" & CountIndex & ")"";</script>" & VbCrLf
			Response.Flush
		End If
		If inStr(GetData(n) & "","@") > 0 Then
			SendMail GetData(n),Topic
		End If
	Next
	%>
	<script>img1.width=400;
	txt1.innerHTML="100";</script>
	<%

End Function

Function SendMail(RvEmail,RvTopic)

	Select Case DEF_BBS_EmailMode
		Case 1: If SendEasyMail(RvEmail,RvTopic,MailBody,MailBody) = 1 Then
					'Response.Write "<br><br>���ϳɹ����͵�����ע�����䣡"
				Else
					'Response.Write "<br><br>��̳δ��ȷ�����ʼ����ͣ����Ϸ���ʧ�ܣ�"
				End If
		Case 2: If SendJmail(RvEmail,RvTopic,MailBody) = 1 Then
					'Response.Write "<br><br>���ϳɹ����͵�����ע�����䣡"
				Else
					'Response.Write "<br><br>��̳δ��ȷ�����ʼ����ͣ����Ϸ���ʧ��2��"
				End If
		Case 3: Response.Write "<br><br>Ⱥ���ʼ���֧��ʹ��CDO�ʼ����ͷ�ʽ���ͣ�"
		Case Else:  Response.Write "<br><br>��̳��֧���ʼ����ͻ�δ������"
	End Select

End Function

Sub GetMailList

	Response.Clear
	Response.ContentType = "text/plain"
	Dim Rs,LoopN
	LoopN = 0
	Set Rs = LDExeCute("select mail from LeadBBS_User where Mail <> ''",0)
	If Not Rs.Eof Then Response.Write Rs.GetString(,,"","" & VbCrLf & "","")
	Rs.Close
	Set Rs = Nothing

End Sub%>