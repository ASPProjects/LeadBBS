<!-- #include file=../../inc/BBSSetup.asp -->
<!-- #include file=../../inc/Board_Popfun.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<%
DEF_BBS_HomeUrl = "../../"
Dim GBL_ID
initDatabase
GBL_CHK_TempStr = ""
GBL_ID = checkSupervisorPass

Dim GBL_SiteDisbleWhyString,GBL_REQ_Flag

Manage_sitehead DEF_SiteNameString & " - ����Ա",""
frame_TopInfo
DisplayUserNavigate("��վ��ͣ/����")
If GBL_CHK_Flag=1 Then
	LoginAccuessFul
Else
	DisplayLoginForm
End If
frame_BottomInfo
closeDataBase
Manage_Sitebottom("none")

Function LoginAccuessFul

	GBL_REQ_Flag = Request("Flag")
	If GBL_REQ_Flag <> "close" and GBL_REQ_Flag<>"open" Then GBL_REQ_Flag = "open"
	If GBL_REQ_Flag = "open" Then
		If application(DEF_MasterCookies & "SiteEnableFlagzoieiu") = 1 or application(DEF_MasterCookies & "SiteDisbleWhyszoieiu") = "" Then
			Response.Write "<div class=alert>��վ�Ѿ��������ˣ�����Ҫ�ٿ���!</div>" & VbCrLf
		Else
			Application.Lock
			application(DEF_MasterCookies & "SiteEnableFlagzoieiu") = 1
			Application.UnLock
			Response.Write "<div class=alert><span class=greenfont><b>��վ�ɹ�����!</b></span></div>" & VbCrLf
		End If
	Else
		If application(DEF_MasterCookies & "SiteEnableFlagzoieiu") = 0 and application(DEF_MasterCookies & "SiteDisbleWhyszoieiu") <> "" Then
			Response.Write "<div class=alert>��վ�Ѿ��رչ��ˣ�����Ҫ�ٹر�!</div>" & VbCrLf
		Else
			If Request.Form("submitflag")="Dieos9xsl29LO_8" Then
				GBL_SiteDisbleWhyString = Request("GBL_SiteDisbleWhyString")
				If Trim(GBL_SiteDisbleWhyString) <> "" Then
					Application.Lock
					application(DEF_MasterCookies & "SiteEnableFlagzoieiu") = 0
					application(DEF_MasterCookies & "SiteDisbleWhyszoieiu") = GBL_SiteDisbleWhyString
					Application.UnLock
					Response.Write "<div class=alert><span class=greenfont><b>��վ�رճɹ������漴��վ��ͣ���ʵ�ԭ��</b></span></div><hr size=1>" & PrintTrueText(GBL_SiteDisbleWhyString) & "<hr size=1>" & VbCrLf
				Else
					Response.Write "<div class=alert>��ͣ�û�����ԭ����Ϊ��!</div>"
				End If
				DisplayStringForm
			Else
				GBL_SiteDisbleWhyString = DEF_Now & " ����վ��ͣ����"
				DisplayStringForm
			End If
			If DEF_UsedDataBase = 1 and DEF_EnableDatabaseCache = 1 Then
				If isObject(Application(DEF_MasterCookies & "con")) = True Then
					Application.Lock
					Application(DEF_MasterCookies & "con").Close
					Set Application(DEF_MasterCookies & "con") = Nothing
					Application(DEF_MasterCookies & "con") = ""
					Application.UnLock
				End If
				GBL_ConFlag = 0
			End If
		End If
	End If

End Function

Function DisplayStringForm

%>
<form action=SiteOpenClose.asp method="post">
	<div class=frameline>Ҫ��ͣ������д��<span class=redfont>��ͣ�û����ʵ�ԭ��</span>:
	</div>
	<div class=frameline>
	<textarea class=fmtxtra name=GBL_SiteDisbleWhyString rows=8 cols=61 ><%If GBL_SiteDisbleWhyString <> "" Then Response.Write VbCrLf & htmlEncode(GBL_SiteDisbleWhyString)%></textarea>
	</div>
	<input name=submitflag type=hidden value="Dieos9xsl29LO_8">
	<input name=Flag type=hidden value="<%=htmlencode(GBL_REQ_Flag)%>">
	
	<div class=frameline>
	<input type=submit value="��ͣ" class=fmbtn> <input type=reset value="ȡ��" class=fmbtn>
	</div>
</form>

	<div class=frameline>
		��ע�⣬��ͣ��վ���벻Ҫ�رմ�������ڣ������µ�¼����ΪCookie���棨��Ҫѡ����Ч����<span class=redfont>�������Ա���޷����¿�����վ</span>(Ϊ�˱�֤��ͣ��Ĺؼ����ݸ��°�ȫ)��
	</div>
<%

End Function

Function PrintTrueText(tempString)

	If tempString<>"" Then
		PrintTrueText=Replace(Replace(Replace(Replace(Replace(htmlEncode(tempString),VbCrLf & " ","<br>" & "&nbsp;"),VbCrLf,"<br>" & VbCrLf),"   "," &nbsp; "),"  "," &nbsp;"),chr(9)," &nbsp; &nbsp; &nbsp;")

		If Left(PrintTrueText,1) = chr(32) Then
			PrintTrueText = "&nbsp;" & Mid(PrintTrueText,2)
		End If
	Else
		PrintTrueText=""
	End If

End Function%>