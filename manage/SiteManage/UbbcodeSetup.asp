<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Ubbcode_Setup.asp -->
<!-- #include file=../../inc/Board_popfun.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<%
DEF_BBS_HomeUrl = "../../"
Const MaxLinkNum = 100
initDatabase
GBL_CHK_TempStr = ""
CheckSupervisorPass

Dim Form_FiltrateBadWordString,Form_DEF_MaxUBBNumber,Form_DEF_UbbUnderwriteImages,Form_DEF_UbbDefaultEdit
Dim Form_DEF_UbbIconG,Form_DEF_UbbLinkData,Form_DEF_SafeUrl

GetDefaultValue

Manage_sitehead DEF_SiteNameString & " - ����Ա",""
frame_TopInfo
DisplayUserNavigate("UBB�����������")
If GBL_CHK_Flag = 1 Then
	UbbcodeSetup
Else
DisplayLoginForm
End If
frame_BottomInfo
closeDataBase
Manage_Sitebottom("none")

Function UbbcodeSetup

%>
<form name="pollform3sdx" method="post" action="UbbcodeSetup.asp">
<input type="hidden" name="SubmitFlag" value=yes>
<p>
		���ã�<a href=SiteSetup.asp>��̳���ò���</a> <a href=UploadSetup.asp>�ϴ�����</a>
		<a href=../User/UserSetup.asp>�û�ע�����</a>
		<span class=grayfont>UBB�������</span>
		<br><span class=grayfont>(����Ϊ��վ��������ע���޸ģ���������ý��ᷢ�����ش���)<br><br>
		��������ú�����վ�����������У��뽫LeadBBS���°��inc/UbbcodeSetup.asp���ǻ�ȥ</span>
</p>
<%If Request.Form("SubmitFlag") <> "" Then
	CheckLinkValue
End If%>
<b><span class=redfont><%=GBL_CHK_TempStr%></span></b>
<p>
<%
If Request.Form("SubmitFlag") <> "" Then
	If GBL_CHK_TempStr <> "" Then
		DisplayDatabaseLink
	Else
		MakeDataBaseLinkFile
		Exit Function
	End If
Else
	DisplayDatabaseLink
End If
%>
</p>
<input type=submit name=�ύ value=�ύ class=fmbtn>
<input type=reset name=ȡ�� value=ȡ�� class=fmbtn>
</form>
<%

End Function

Function CheckLinkValue

	GetFormValue

End Function

Function DisplayDatabaseLink

		%>
		<table border=0 cellpadding=0 cellspacing=0 width="100%" class=frame_table>
		<tr>
			<td class=tdbox width=120>���ֹ���</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_FiltrateBadWordString" maxlength="5024" size="50" value="<%=htmlencode(Form_FiltrateBadWordString)%>"><span class=grayfont> <br>
			ʹ��|���ŷָ������ַ��Զ��滻Ϊ����</span></td>
		</tr>
		<tr>
			<td class=tdbox>��������</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_MaxUBBNumber" maxlength="3" size="10" value="<%=htmlencode(Form_DEF_MaxUBBNumber)%>">
			<br><span class=grayfont>(ÿ����������������ʾ����������Ĳ��ٱ���Ϊ����ͼƬ)</span></td>
		</tr>
		<tr>
			<td class=tdbox>ǩ��ͼƬ</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_UbbUnderwriteImages value=0<%If Form_DEF_UbbUnderwriteImages = 0 Then%> checked<%End If%>></td><td>�ر�</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_UbbUnderwriteImages value=1<%If Form_DEF_UbbUnderwriteImages = 1 Then%> checked<%End If%>></td><td>����</td>
          		<td>&nbsp; (<span class=grayfont>�Ƿ�����ͼƬ��Ϊǩ��</span>)</td>
          		</tr></table></td>
		</tr>
		<tr>
			<td class=tdbox>����ģʽ</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=Form_DEF_UbbDefaultEdit value=0<%If Form_DEF_UbbDefaultEdit = 0 Then%> checked<%End If%>></td><td>��ͨģʽ</td>
          		<td><input class=fmchkbox type=radio name=Form_DEF_UbbDefaultEdit value=1<%If Form_DEF_UbbDefaultEdit = 1 Then%> checked<%End If%>></td><td>�߼�ģʽ</td>
          		</tr></table>
          		&nbsp; (<span class=grayfont>ָ��Ĭ�ϵķ���ģʽ���߼�ģʽ�������߱༭��ҳ</span>)</td>
		</tr>
		<tr>
			<td class=tdbox width=50>�������</td>
			<td class=tdbox>
			<textarea cols=80 name=Form_DEF_UbbIconG style="width: 100%;height:110px; word-break: break-all;" class=fmtxtra><%If Form_DEF_UbbIconG <> "" Then Response.Write VbCrLf & Server.htmlEncode(Form_DEF_UbbIconG)%></textarea>
			<br/>
			ʹ�õ��ǵ��ո�ָ������� ������� 1 25 �������������Ӧ���1-25,������û�������
			���50������ �������<%=DEF_UBBiconNumber%></td>
		</tr>
		<tr>
			<td class=tdbox width=50>��������</td>
			<td class=tdbox>
			<textarea cols=180 name=Form_DEF_UbbLinkData style="width: 100%;height:220px; word-break: break-all;" class=fmtxtra><%If Form_DEF_UbbLinkData <> "" Then Response.Write VbCrLf & Server.htmlEncode(Form_DEF_UbbLinkData)%></textarea>
			<br/>
			ʹ�õ��ǵ��ո�ָ������� LeadBBS http://www.leadbbs.com/ ���������е�����leadbbs�ַ�������ӵ�http://www.leadbbs.com/ ע���ַ���ѡ�� �����滻����������
			<br/>���50����Ŀ</td>
			
		</tr>
		<tr>
			<td class=tdbox width=120>��ȫ��ַ</td>
			<td class=tdbox><input class=fminpt type="text" name="Form_DEF_SafeUrl" maxlength="5024" size="50" value="<%=htmlencode(Form_DEF_SafeUrl)%>"><span class=grayfont> <br>
			ʹ��|���ŷָ���ע����ַֻ�ܰ���1-2��ţ�����|leadbbs.com|youbute.com|sina.com.cn|
			</span></td>
		</tr>
		</table>
		<%

End Function

Function GetDefaultValue

	Form_FiltrateBadWordString = FiltrateBadWordString
	Form_DEF_MaxUBBNumber = DEF_MaxUBBNumber
	Form_DEF_UbbDefaultEdit = DEF_UbbDefaultEdit
	Form_DEF_UbbUnderwriteImages = DEF_UbbUnderwriteImages
	Form_DEF_SafeUrl = DEF_SafeUrl
	
	Dim N,I
	Form_DEF_UbbLinkData = ""
	If isArray(DEF_UbbLinkData) and isArray(DEF_UbbLinkUrl) Then
		N = UBound(DEF_UbbLinkData,1)
		If UBound(DEF_UbbLinkUrl,1) < N Then N = UBound(DEF_UbbLinkUrl,1)
		For I = 0 to N
			Form_DEF_UbbLinkData = Form_DEF_UbbLinkData & DEF_UbbLinkData(I) & " " & DEF_UbbLinkUrl(I) & VbCrLf
		Next
	End If
	
	Form_DEF_UbbIconG = ""
	If isArray(DEF_UbbIconG) and isArray(DEF_UbbIconMax) and isArray(DEF_UbbIconMin) Then
		N = UBound(DEF_UbbIconG,1)
		If UBound(DEF_UbbIconMax,1) < N Then N = UBound(DEF_UbbIconMax,1)
		If UBound(DEF_UbbIconMin,1) < N Then N = UBound(DEF_UbbIconMin,1)
		For I = 0 to N
			Form_DEF_UbbIconG = Form_DEF_UbbIconG & DEF_UbbIconG(I) & " " & DEF_UbbIconMin(I) & " " & DEF_UbbIconMax(I) & VbCrLf
		Next
	End If

End Function

Function GetFormValue

	Form_FiltrateBadWordString = Trim(Request.Form("Form_FiltrateBadWordString"))
	Form_DEF_MaxUBBNumber = Trim(Request.Form("Form_DEF_MaxUBBNumber"))
	Form_DEF_UbbUnderwriteImages = Trim(Request.Form("Form_DEF_UbbUnderwriteImages"))
	Form_DEF_UbbDefaultEdit = Trim(Request.Form("Form_DEF_UbbDefaultEdit"))
	Form_DEF_UbbIconG = Trim(Request.Form("Form_DEF_UbbIconG"))
	Form_DEF_UbbLinkData = Trim(Request.Form("Form_DEF_UbbLinkData"))
	Form_DEF_SafeUrl = Trim(Request.Form("Form_DEF_SafeUrl"))

	If inStr(Form_FiltrateBadWordString,"""") or inStr(Form_FiltrateBadWordString,"%") Then GBL_CHK_TempStr = "���ֹ��˲��ܰ��������Ż�ٷֺ�<br>" & VbCrLf
	If isNumeric(Form_DEF_MaxUBBNumber) = 0 Then GBL_CHK_TempStr = "������������Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_UbbUnderwriteImages) = 0 Then GBL_CHK_TempStr = "�Ƿ�����ǩ��ͼƬ����Ϊ����<br>" & VbCrLf
	If isNumeric(Form_DEF_UbbDefaultEdit) = 0 Then GBL_CHK_TempStr = "����ģʽ����Ϊ����<br>" & VbCrLf
	If inStr(Form_DEF_SafeUrl,"""") or inStr(Form_DEF_SafeUrl,"%") Then GBL_CHK_TempStr = "��ȫ��ַ�б��ܰ��������Ż�ٷֺ�<br>" & VbCrLf
	
	Dim TempA,TempB,a,b,c,Num,N,aa,bb,cc,Counter
	TempA = Split(Form_DEF_UbbIconG,VbCrLf)
	aa = ""
	bb = ""
	cc = ""
	Counter = 0
	If isArray(TempA) Then
		Num = Ubound(TempA,1)
		For N = 0 to Num
			TempB = Split(Trim(TempA(N))," ")
			If Ubound(TempB,1) = 2 Then
				a = TempB(0)
				b = TempB(1)
				c = TempB(2)
				If isNumeric(b) and isNumeric(c) Then
					b = Fix(cCur(b))
					c = Fix(cCur(c))
					If c >= b and c <= DEF_UBBiconNumber and b <= DEF_UBBiconNumber Then
						a = Replace(a,"server","")
						a = Replace(a,"<" & "%","")
						a = Replace(a,"%" & ">","")
						If aa = "" Then
							aa = """" & Replace(a,"""","""""") & """"
						Else
							aa = aa & ",""" & Replace(a,"""","""""") & """"
						End If
						If bb = "" Then
							bb = b
						Else
							bb = bb & "," & b
						End If
						If cc = "" Then
							cc = c
						Else
							cc = cc & "," & c
						End If
						Counter = Counter + 1
					End If
				End If
			End If
		Next
		If aa <> "" Then
			aa = "DEF_UbbIconG = Array(" & aa & ")"
			bb = "DEF_UbbIconMin = Array(" & bb & ")"
			cc = "DEF_UbbIconMax = Array(" & cc & ")"
		End if
		Form_DEF_UbbIconG = "DEF_UbbIconGNum = " & Counter
		Form_DEF_UbbIconG = Form_DEF_UbbIconG & VbCrLf & aa & VbCrLf & bb & VbCrLf & cc
	Else
		Form_DEF_UbbIconG = "DEF_UbbIconGNum = 0"
	End If
	Form_DEF_UbbIconG = "Dim DEF_UbbIconG,DEF_UbbIconMax,DEF_UbbIconMin,DEF_UbbIconGNum" & VbCrLf & Form_DEF_UbbIconG
	
	
	TempA = Split(Form_DEF_UbbLinkData,VbCrLf)
	aa = ""
	bb = ""
	Counter = 0
	If isArray(TempA) Then
		Num = Ubound(TempA,1)
		For N = 0 to Num
			TempB = Split(Trim(TempA(N))," ")
			If Ubound(TempB,1) = 1 Then
				a = TempB(0)
				b = TempB(1)
				If a <> "" and b <> "" Then
						a = Replace(a,"<" & "%","")
						a = Replace(a,"%" & ">","")
						b = Replace(b,"server","")
						b = Replace(b,"<" & "%","")
						b = Replace(b,"%" & ">","")
						If aa = "" Then
							aa = """" & Replace(a,"""","""""") & """"
						Else
							aa = aa & ",""" & Replace(a,"""","""""") & """"
						End If
						If bb = "" Then
							bb = """" & Replace(b,"""","""""") & """"
						Else
							bb = bb & ",""" & Replace(b,"""","""""") & """"
						End If
						Counter = Counter + 1
				End If
			End If
		Next
		If aa <> "" Then
			aa = "DEF_UbbLinkData = Array(" & aa & ")"
			bb = "DEF_UbbLinkUrl = Array(" & bb & ")"
		End if
		Form_DEF_UbbLinkData = "DEF_UbbLinkNum = " & Counter
		Form_DEF_UbbLinkData = Form_DEF_UbbLinkData & VbCrLf & aa & VbCrLf & bb
	Else
		Form_DEF_UbbLinkData = "DEF_UbbLinkNum = 0"
	End If
	Form_DEF_UbbLinkData = "Dim DEF_UbbLinkData,DEF_UbbLinkUrl,DEF_UbbLinkNum" & VbCrLf & Form_DEF_UbbLinkData

End Function

Function MakeDataBaseLinkFile

	Dim n,TempStr
	TempStr = ""
	TempStr = TempStr & chr(60) & "%" & VbCrLf
	TempStr = TempStr & "const FiltrateBadWordString=" & Chr(34) & Replace(Form_FiltrateBadWordString,Chr(34),Chr(34) & Chr(34)) & Chr(34) & VbCrLf
	TempStr = TempStr & "const DEF_MaxUBBNumber = " & Form_DEF_MaxUBBNumber & VbCrLf
	TempStr = TempStr & "const DEF_UbbUnderwriteImages = " & Form_DEF_UbbUnderwriteImages & VbCrLf
	TempStr = TempStr & "const DEF_UbbDefaultEdit = " & Form_DEF_UbbDefaultEdit & VbCrLf
	
	TempStr = TempStr & Form_DEF_UbbIconG & VbCrLf
	TempStr = TempStr & Form_DEF_UbbLinkData & VbCrLf
	
	TempStr = TempStr & "const DEF_SafeUrl=" & Chr(34) & Replace(Form_DEF_SafeUrl,Chr(34),Chr(34) & Chr(34)) & Chr(34) & VbCrLf

	TempStr = TempStr & "%" & chr(62) & VbCrLf

	ADODB_SaveToFile TempStr,"../../inc/Ubbcode_Setup.asp"
	CALL Update_InsertSetupRID(1051,"inc/Ubbcode_Setup.asp",1,TempStr," and ClassNum=" & 1)
	
	If GBL_CHK_TempStr = "" Then
		Response.Write "<br><span class=greenfont>2.�ɹ�������ã�</span>"
	Else
		%><%=GBL_CHK_TempStr%><br>��������֧������д���ļ����ܣ���ʹ��FTP�ȹ��ܣ���<span class=redfont>../../inc/Ubbcode_Setup.asp</span>�ļ��滻�ɿ�������(ע�ⱸ��)<p>
		<textarea name="fileContent" cols="80" rows="30" class=fmtxtra><%=Server.htmlencode(TempStr)%></textarea><%
		GBL_CHK_TempStr = ""
	End If

End Function%>