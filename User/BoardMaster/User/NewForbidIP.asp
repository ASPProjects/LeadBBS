<!-- #include file=../../../inc/BBSsetup.asp -->
<!-- #include file=../../../inc/Board_Popfun.asp -->
<!-- #include file=../inc/BoardMaster_Fun.asp -->
<%
DEF_BBS_HomeUrl = "../../../"
initDatabase
GBL_CHK_TempStr = ""

SiteHead(DEF_SiteNameString & " - " & DEF_PointsName(6) & "����")

UserTopicTopInfo
DisplayUserNavigate("���Σɣе�ַ")%>
<br><br><%If GBL_CHK_Flag=1 and BDM_isBoardMasterFlag = 1 and BDM_SpecialPopedomFlag = 1 Then
	LoginAccuessFul
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

Dim GBL_IPStart,GBL_IPEnd,GBL_WhyString,GBL_ExpiresTime,GBL_UserName,GBL_UserName_UserID
Dim GBL_AnnounceID,GBL_MessageID
GBL_ExpiresTime = -1

Function LoginAccuessFul

	If DEF_EnableForbidIP = 0 Then
		Response.Write "<br><p><b><font color=Red class=redfont>ϵͳ�Ѿ���ֹ����IP���ܣ���Ҫ����IP��ַ����ϵ����Ա������</font></b></p>"
		Exit Function
	End If
	GBL_UserName = Trim(Left(Request.Form("GBL_UserName"),14))
	GBL_AnnounceID = Left(Request.Form("GBL_AnnounceID"),14)
	GBL_MessageID = Left(Request.Form("GBL_MessageID"),14)
	
	If GBL_MessageID <> "" Then
	ElseIf GBL_AnnounceID <> "" Then
	ElseIf GBL_UserName <> "" Then
		'CheckUserNameExist(GBL_UserName)
	Else
		'GBL_IPStart = Request.Form("GBL_IPStart")
		'GBL_IPEnd = Request.Form("GBL_IPEnd")
	End If
	GBL_ExpiresTime = Left(Request.Form("GBL_ExpiresTime"),14)
	GBL_WhyString = Left(Request.Form("GBL_WhyString"),100)
	
	If isNumeric(GBL_ExpiresTime) = 0 Then GBL_ExpiresTime = -1
	GBL_ExpiresTime = fix(cCur(GBL_ExpiresTime))
	If GBL_ExpiresTime < 0 or GBL_ExpiresTime > 30 Then GBL_ExpiresTime = -1

	If Request.Form("submitflag") <> "" Then
		CheckNewIP
		If GBL_CHK_TempStr = "" Then
			SaveNewIP
			Response.Write GBL_CHK_TempStr
		Else
			DisplayNewIPForm
		End If
	Else
		DisplayNewIPForm
	End If

End Function

Function SaveNewIP

	Dim SQL,Rs,Number
	GBL_IPEnd = Right("000000000000" & cStr(GBL_IPEnd),12)
	GBL_IPStart = Right("000000000000" & cStr(GBL_IPStart),12)
	Number = (Left(GBL_IPEnd,3) * 256 * 256 * 256 + Mid(GBL_IPEnd,4,3) * 256 * 256 + Mid(GBL_IPEnd,7,3) * 256 + Mid(GBL_IPEnd,10,3))-(Left(GBL_IPStart,3) * 256 * 256 * 256 + Mid(GBL_IPStart,4,3) * 256 * 256 + Mid(GBL_IPStart,7,3) * 256 + Mid(GBL_IPStart,10,3)) + 1
	SQL = sql_select("Select ID from LeadBBS_ForbidIP where IPStart<=" & GBL_IPStart & " and IPEnd>=" & GBL_IPEnd,1)
	Set Rs = LDExeCute(SQL,0)
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		SQL = "Insert Into LeadBBS_ForbidIP(IPStart,IPEnd,IPNumber,ExpiresTime,WhyString) Values(" & GBL_IPStart & "," & GBL_IPEnd & "," & Number & "," & GBL_ExpiresTime & ",'" & Replace(GBL_WhyString,"'","''") & "')"
		CALL LDExeCute(SQL,0)
		GBL_CHK_TempStr = "<font color=008800 class=greenfont>�ɹ����δ�IP��,����" & Number & "��!<br>" & VbCrLf
		'GBL_CHK_TempStr = GBL_CHK_TempStr & "��ʼIP��ַ��" & GBL_IPStart & "<br>" & VbCrLf
		'GBL_CHK_TempStr = GBL_CHK_TempStr & "��ֹIP��ַ��" & GBL_IPEnd & "</font><br>" & VbCrLf
	Else
		Rs.Close
		Set Rs = Nothing
		GBL_CHK_TempStr = "<font color=ff0000 class=redfont>���󣺴�IP��ַ���Ѿ��������б���,�����ظ����!</font><br>" & VbCrLf
	End If
	If CheckSupervisorUserName = 0 Then
		CALL LDExeCute("Update LeadBBS_User Set LastWriteTime=" & GetTimeValue(DEF_Now) & " where ID=" & GBL_UserID,1)
	End If

End Function

Function CheckNewIP

	If CheckWriteEventSpace = 0 Then
		Response.Write "<b><font color=Red Class=redfont>���Ĳ�����Ƶ�����Ժ������ύ!</font></b> <br>" & VbCrLf
		Exit Function
	End If
	If GBL_MessageID <> "" or Request.Form("submitflag") = "LKOkxk4" Then
		If CheckMessageID(GBL_MessageID) = 0 Then
			Exit Function
		End If
	ElseIf GBL_AnnounceID <> "" or Request.Form("submitflag") = "LKOkxk3" Then
		If CheckAnnounceID(GBL_AnnounceID) = 0 Then
			Exit Function
		End If
	ElseIf GBL_UserName <> "" or Request.Form("submitflag") = "LKOkxk2" Then
		If CheckUserNameExist(GBL_UserName) = 0 Then
			Exit Function
		End If
	End If
	Dim Tmp_IPStart,Tmp_IPEnd
	Tmp_IPStart = FormatIPaddress(GBL_IPStart)
	Tmp_IPEnd = FormatIPaddress(GBL_IPEnd)

	If isNumeric(GBL_ExpiresTime) = 0 Then GBL_ExpiresTime = -1
	GBL_ExpiresTime = fix(cCur(GBL_ExpiresTime))
	If GBL_ExpiresTime < 0 or GBL_ExpiresTime > 30 Then GBL_ExpiresTime = -1
	If GBL_ExpiresTime = -1 Then
		GBL_CHK_TempStr = "������������ѡ���������ȷѡ�񣬿����Ǵ��û�IP��ַ�����Ϲ滮��"
		Exit function
	End If

	If Len(Tmp_IPStart) <> 15 Then
		GBL_CHK_TempStr = "������ʼ�ɣе�ַ���󣬿����Ǵ��û�IP��ַ�����Ϲ滮"
		Exit function
	End If

	If Len(Tmp_IPStart) <> 15 Then
		GBL_CHK_TempStr = "������ֹ�ɣе�ַ���󣬿����Ǵ��û�IP��ַ�����Ϲ滮"
		Exit function
	End If
	
	Dim NewGBL_IPStart,NewGBL_IPEnd
	NewGBL_IPStart = Left(Replace(Tmp_IPStart,".",""),14)
	NewGBL_IPEnd = Left(Replace(Tmp_IPEnd,".",""),14)
	If isNumeric(NewGBL_IPStart) = 0 Then
		GBL_CHK_TempStr = "������ʼ�ɣе�ַ���󣬱��������֣������Ǵ��û�IP��ַ�����Ϲ滮"
		Exit function
	End If
	
	If isNumeric(NewGBL_IPEnd) = 0 Then
		GBL_CHK_TempStr = "������ֹ�ɣе�ַ���󣬱��������֣������Ǵ��û�IP��ַ�����Ϲ滮"
		Exit function
	End If
	
	NewGBL_IPStart = cCur(NewGBL_IPStart)
	NewGBL_IPEnd = cCur(NewGBL_IPEnd)
	If NewGBL_IPStart > NewGBL_IPEnd Then
		GBL_CHK_TempStr = "������ֹ�ɣе�ַ���ܱ���ʼ�ɣе�ַС�������Ǵ��û�IP��ַ�����Ϲ滮"
		Exit function
	End If
	
	If NewGBL_IPStart > 255255255255 Then
		GBL_CHK_TempStr = "������ʼ�ɣе�ַ�������IP��ַΪ255.255.255.255�������Ǵ��û�IP��ַ�����Ϲ滮"
		Exit function
	End If
	If NewGBL_IPEnd > 255255255255 Then
		GBL_CHK_TempStr = "������ֹ�ɣе�ַ�������IP��ַΪ255.255.255.255�������Ǵ��û�IP��ַ�����Ϲ滮"
		Exit function
	End If

	GBL_IPStart = NewGBL_IPStart
	GBL_IPEnd = NewGBL_IPEnd
	If GBL_ExpiresTime > 0 Then
		GBL_ExpiresTime = GetTimeValue(DateAdd("d",GBL_ExpiresTime,DEF_Now))
	Else
		GBL_ExpiresTime = 0
	End If

End Function

Function DisplayNewIPForm

	Dim N
	If GBL_CHK_TempStr <> "" Then Response.Write "<p><font color=ff0000 class=redfont><b>" & GBL_CHK_TempStr & "</b></font></p>"%>

			<%If Request.Form("submitflag") = "LKOkxk2" or Request.Form("submitflag") = "" Then%>
			<p>
		  <b>���������û��������Σ�������Ҫ���Σɣе�ַ�������û�����</b>
          <form action=NewForbidIP.asp method=post id=fobform name=fobform>
			���ߵ��û�����<input name=GBL_UserName value="<%=htmlencode(GBL_UserName)%>" class=fminpt><br>
			<input name=submitflag type=hidden value="LKOkxk2">
			����ʱ��ѡ��<select name=GBL_ExpiresTime>
							<%For N = 1 to 30
								If N = GBL_ExpiresTime Then
									Response.Write "<option value=" & N & " selected>����" & Right("0" & N,2) & "��</option>"
								Else
									Response.Write "<option value=" & N & ">����" & Right("0" & N,2) & "��</option>"
								End If
							Next%>
							<option value=0<%If GBL_ExpiresTime = 0 Then Response.Write " Selected"%>>��������</option>
						</select>
						<br>
			����ԭ��ע����<input name=GBL_WhyString MaxLength=100 size=30 value="<%=htmlencode(GBL_WhyString)%>" class=fminpt>
			<select onchange="document.fobform.GBL_WhyString.value=this.value;">
				<option value="">===һЩ����ԭ����ѡ��===</option>
				<option value="��������ɫ������">��������ɫ������</option>
				<option value="����̳���ж��⹥��(�ڿ���Ϊ)">����̳���ж��⹥��(�ڿ���Ϊ)</option>
				<option value="��ͣ�Ķ����ˮ��ע�����û�">��ͣ�Ķ����ˮ��ע�����û�</option>
			</select>
			<br><br>
			<input type=submit value="�ύ" class=fmbtn> <input type=reset value="ȡ��" class=fmbtn></form>
			<br><%End If%>

			<%If Request.Form("submitflag") = "LKOkxk3" or Request.Form("submitflag") = "" Then%>
			<p>
		 	<b>���ݷ������������Σ�����ĳ�û����������ӵı��</b>
          	<form action=NewForbidIP.asp method=post id=fobform name=fobform>
			��̳���ӱ�ţ�<input name=GBL_AnnounceID value="<%=htmlencode(GBL_AnnounceID)%>" class=fminpt><br>
			<input name=submitflag type=hidden value="LKOkxk3">
			����ʱ��ѡ��<select name=GBL_ExpiresTime>
							<%For N = 1 to 30
								If N = GBL_ExpiresTime Then
									Response.Write "<option value=" & N & " selected>����" & Right("0" & N,2) & "��</option>"
								Else
									Response.Write "<option value=" & N & ">����" & Right("0" & N,2) & "��</option>"
								End If
							Next%>
							<option value=0<%If GBL_ExpiresTime = 0 Then Response.Write " Selected"%>>��������</option>
						</select>
						<br>
			����ԭ��ע����<input name=GBL_WhyString MaxLength=100 size=30 value="<%=htmlencode(GBL_WhyString)%>" class=fminpt>
			<select onchange="document.fobform.GBL_WhyString.value=this.value;">
				<option value="">===һЩ����ԭ����ѡ��===</option>
				<option value="��������ɫ������">��������ɫ������</option>
				<option value="����̳���ж��⹥��(�ڿ���Ϊ)">����̳���ж��⹥��(�ڿ���Ϊ)</option>
				<option value="��ͣ�Ķ����ˮ��ע�����û�">��ͣ�Ķ����ˮ��ע�����û�</option>
			</select>
			<br><br>
			<input type=submit value="�ύ" class=fmbtn> <input type=reset value="ȡ��" class=fmbtn></form>
			<br>
			<p>ʹ��˵����<font color=888888 class=grayfont>���ӵı�ţ��ڰ����б��У�����������ǰ���ͼ���Ͽ�����ʾ���������<br>
			�����������ڲ鿴��������ʱ������������������ϣ�������ʾ��������ظ����ı��</font><br><br><%End If%>
			

			<%If Request.Form("submitflag") = "LKOkxk4" or Request.Form("submitflag") = "" Then%>
			<p>
			<b>���ݶ���Ϣ��������Σ�����ĳ�û������Ͷ���Ϣ�ı��</b>
			<form action=NewForbidIP.asp method=post id=fobform name=fobform>
			����Ϣ�ı�ţ�<input name=GBL_MessageID value="<%=htmlencode(GBL_MessageID)%>" class=fminpt><br>
			<input name=submitflag type=hidden value="LKOkxk4">
			����ʱ��ѡ��<select name=GBL_ExpiresTime>
							<%For N = 1 to 30
								If N = GBL_ExpiresTime Then
									Response.Write "<option value=" & N & " selected>����" & Right("0" & N,2) & "��</option>"
								Else
									Response.Write "<option value=" & N & ">����" & Right("0" & N,2) & "��</option>"
								End If
							Next%>
							<option value=0<%If GBL_ExpiresTime = 0 Then Response.Write " Selected"%>>��������</option>
						</select>
						<br>
			����ԭ��˵����<input name=GBL_WhyString MaxLength=100 size=30 value="<%=htmlencode(GBL_WhyString)%>" class=fminpt>
			<select onchange="document.fobform.GBL_WhyString.value=this.value;">
				<option value="">===һЩ����ԭ����ѡ��===</option>
				<option value="��������ɫ������">��������ɫ������</option>
				<option value="����̳���ж��⹥��(�ڿ���Ϊ)">����̳���ж��⹥��(�ڿ���Ϊ)</option>
				<option value="��ͣ�Ķ����ˮ��ע�����û�">��ͣ�Ķ����ˮ��ע�����û�</option>
			</select>
			<br><br>
			<input type=submit value="�ύ" class=fmbtn> <input type=reset value="ȡ��" class=fmbtn></form>
			<br>
			<p>ʹ��˵����<font color=888888 class=grayfont>����Ϣ��ſ����ڲ鿴�ռ����б�����ʾ</font><br><br><%End If%>

<%End Function


Function FormatIPaddress(KIP)

	Dim IP
	IP = KIP
	Rem ��ȥ���׵Ŀյ㣬����ʽ����XXX.XXX.XXX.XXX
	Dim Temp1,Temp2,TempN,Temp
	IP = Trim(IP & "")
	If inStr(IP,".") = 0 or Len(IP) = "" Then
		FormatIPaddress = IP
		Exit Function
	End if
	
	Temp1 = Split(IP,".")
	IP = ""
	Temp2 = Ubound(Temp1,1)
	
	TempN = 0
	do while IP = ""
		If Temp1(TempN) <> "" Then
			if IsNumeric(Temp1(TempN)) Then Temp1(TempN) = cStr(cCur(Temp1(TempN)))
			If Len(Temp1(TempN)) < 3 Then
				IP = string(3-len(Temp1(TempN)),"0") & Temp1(TempN)
			else
				IP = Temp1(TempN)
			End If
			TempN = TempN + 1
			Exit Do
		Else
			TempN = TempN + 1
		End If
		If TempN > Temp2 Then Exit do
	Loop
	
	For Temp = TempN to Temp2
		If Temp1(TempN) <> "" Then
			If isNumeric(Temp1(TempN)) = 0 Then
				FormatIPaddress = ""
				Exit Function
			End If
			Temp1(TempN) = Fix(cCur(Temp1(TempN)))
			If Temp1(TempN) < 0 or Temp1(TempN) > 255 Then
				FormatIPaddress = ""
				Exit Function
			End If
			if IsNumeric(Temp1(TempN)) Then Temp1(TempN) = cStr(cCur(Temp1(TempN)))
			If Len(Temp1(TempN)) < 3 Then
				IP = IP & "." & string(3-len(Temp1(TempN)),"0") & Temp1(TempN)
			else
				IP = IP & "." & Temp1(TempN)
			End If
		End If
		TempN = TempN + 1
	Next
	FormatIPaddress = IP
	Rem ���ص�IP��ַ�պ���15λ���������15���ַ����Ǵ�����Ч��IP��ַ

End Function


Rem ���ĳ�û����Ƿ����
Function CheckUserNameExist(UserName)

	If UserName <> "" and inStr(UserName,",") = 0 and inStr(Lcase(DEF_SupervisorUserName),"," & Lcase(UserName) & ",") > 0 Then
		'��������ͬ����ʾ��Ϊ���Է�����Ա���ֱ�й©��ʵ��Ӧ����ʾ����Ա���ܱ�����
		GBL_CHK_TempStr = "�����û���" & htmlencode(UserName) & "�����ڣ�"
		CheckUserNameExist = 0
		Exit Function
	End If

	Dim Rs
	Set Rs = LDExeCute(sql_select("Select ID,UserName from LeadBBS_User where UserName='" & Replace(UserName,"'","''") & "'",1),0)
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		CheckUserNameExist = 0
		GBL_CHK_TempStr = "�����û���" & htmlencode(UserName) & "�����ڣ�"
		Exit Function
	Else
		GBL_UserName_UserID = cCur(Rs(0))
		GBL_UserName = Rs(1)
	End if
	Rs.Close
	Set Rs = Nothing
	
	Set Rs = LDExeCute(sql_select("Select IP from LeadBBS_OnlineUser where UserID=" & GBL_UserName_UserID,1),0)
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		CheckUserNameExist = 0
		GBL_CHK_TempStr = "�����û���" & htmlencode(UserName) & "Ŀǰ�����ߣ��޷�������Σ���ʹ�������ķ�ʽ�����Σ�"
		Exit Function
	Else
		GBL_IPStart = Rs(0)
		GBL_IPEnd = GBL_IPStart
		Rs.Close
		Set Rs = Nothing
	End if
		
	CheckUserNameExist = 1

End Function

Rem ���ĳ����
Function CheckAnnounceID(AnnounceID)

	If isNumeric(AnnounceID) = False Then
		GBL_CHK_TempStr = "�������Ӳ������ڻ���Ȩ���Σ�"
		CheckAnnounceID = 0
		Exit Function
	End If
	AnnounceID = Fix(cCur(AnnounceID))
	If AnnounceID < 1 Then
		GBL_CHK_TempStr = "���󣬱��" & htmlencode(AnnounceID) & "�����Ӳ������ڻ���Ȩ���Σ�"
		CheckAnnounceID = 0
		Exit Function
	End If
	
	Dim Rs
	Set Rs = LDExeCute(sql_select("Select IPAddress,UserName from LeadBBS_Announce where ID=" & AnnounceID,1),0)
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		CheckAnnounceID = 0
		GBL_CHK_TempStr = "���󣬱��" & htmlencode(AnnounceID) & "�����Ӳ������ڻ���Ȩ���Σ�"
		Exit Function
	Else
		GBL_IPStart = Rs(0)
		GBL_IPEnd = GBL_IPStart
		GBL_UserName = Rs(1)
	End if
	Rs.Close
	Set Rs = Nothing

	If GBL_UserName <> "" and inStr(GBL_UserName,",") = 0 and inStr(Lcase(DEF_SupervisorUserName),"," & Lcase(GBL_UserName) & ",") > 0 Then
		'��������ͬ����ʾ��Ϊ���Է�����Ա���ֱ�й©��ʵ��Ӧ����ʾ����Ա���ܱ�����
		GBL_CHK_TempStr = "���󣬱��" & htmlencode(AnnounceID) & "�����Ӳ������ڻ���Ȩ���Σ�"
		CheckAnnounceID = 0
		Exit Function
	End If
	CheckAnnounceID = 1

End Function


Rem ���ĳ����
Function CheckMessageID(MessageID)

	If isNumeric(MessageID) = False Then
		GBL_CHK_TempStr = "���󣬶���Ϣ�������ڻ���Ȩ���Σ�"
		CheckMessageID = 0
		Exit Function
	End If
	MessageID = Fix(cCur(MessageID))
	If MessageID < 1 Then
		GBL_CHK_TempStr = "���󣬱��" & htmlencode(MessageID) & "�Ķ���Ϣ�������ڻ���Ȩ���Σ�"
		CheckMessageID = 0
		Exit Function
	End If
	
	Dim Rs
	Set Rs = LDExeCute(sql_select("Select IP,FromUser from LeadBBS_InfoBox where ID=" & MessageID,1),0)
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		CheckMessageID = 0
		GBL_CHK_TempStr = "���󣬱��" & htmlencode(MessageID) & "�Ķ���Ϣ�������ڻ���Ȩ���Σ�"
		Exit Function
	Else
		GBL_IPStart = Rs(0)
		GBL_IPEnd = GBL_IPStart
		GBL_UserName = Rs(1)
	End if
	Rs.Close
	Set Rs = Nothing

	If GBL_UserName <> "" and inStr(GBL_UserName,",") = 0 and inStr(Lcase(DEF_SupervisorUserName),"," & Lcase(GBL_UserName) & ",") > 0 Then
		'��������ͬ����ʾ��Ϊ���Է�����Ա���ֱ�й©��ʵ��Ӧ����ʾ����Ա���ܱ�����
		GBL_CHK_TempStr = "���󣬱��" & htmlencode(MessageID) & "�Ķ���Ϣ�������ڻ���Ȩ���Σ�"
		CheckMessageID = 0
		Exit Function
	End If
	CheckMessageID = 1

End Function%>