<!-- #include file=../inc/BBSsetup.asp -->
<!-- #include file=../inc/User_Setup.asp -->
<!-- #include file=../inc/Board_popfun.asp -->
<!-- #include file=../inc/ubbcode.asp -->
<!-- #include file=../inc/Limit_fun.asp -->
<!-- #include file=inc/User_fun.asp -->
<!-- #include file=inc/UserTopic.asp -->
<!-- #include file=inc/Mail_fun.asp -->
<!-- #include file=../inc/Constellation2.asp -->
<!-- #include file=inc/Fun_SendMessage.asp -->
<%
Const LMT_RegVerifyQuestion = "<img src=../images/temp/tmp.jpg>,ע�����ȫ��ʹ�ô�д��ĸ" 'ע����֤��ʾ��Ϣ��������HTML��ʽ������ʹ��ͼƬ��������д��ʾ������ע����֤��Ϣ��
Const LMT_RegVerifyAnswer = "APPLE" 'ע����֤��Ҫ��д�Ĵ𰸡�
DEF_BBS_HomeUrl = "../"

Form_FaceWidth = DEF_AllFaceMaxWidth
Form_FaceHeight = DEF_AllFaceMaxWidth
GBL_CHK_PWdFlag = 0
CursorLocation = 3
initDatabase

If Request.Form("checkflag") = "1" Then
	Reg_CheckInfo
	CloseDatabase
	Response.End
End If

GBL_CHK_TempStr = ""

Dim AttestNumber,Form_Action
AttestNumber = 0
Dim Form_ID,ShowTestNumber

If Def_UserTestNumber = 2 Then
	ShowTestNumber = 0
ElseIf Def_UserTestNumber = 1 Then
	If DEF_EnableAttestNumber = 1 Then
		ShowTestNumber = 3
	Else
		ShowTestNumber = 4
	End If
Else
	ShowTestNumber = DEF_EnableAttestNumber
End If

Dim reg_action,reg_command
reg_action = Left(Request("action"),30)
reg_command = Left(Request("command"),30)

'�����ر�״̬������󶨻���������
If GetBinarybit(DEF_Sideparameter,10) = 0 Then
	reg_action = ""
	reg_command = ""
End If

If reg_action <> "bind" Then
	BBS_SiteHead DEF_SiteNameString & " - ע�����û�",0,"<span class=navigate_string_step>ע�����û�</span>"
	UpdateOnlineUserAtInfo GBL_board_ID,"ע�����û�"
Else	
	BBS_SiteHead DEF_SiteNameString & " - ����/���ʺ�",0,"<span class=navigate_string_step>����/���ʺ�</span>"
	UpdateOnlineUserAtInfo GBL_board_ID,"����/���ʺ�"
End If
UserTopicTopInfo("")


If reg_action <> "bind" or (reg_action = "bind" and reg_command = "reg") Then
	If Request.form("JoinFlag") <> "" Then
		If LMT_EnableRegNewUsers = 1 Then
			If Request.Form("SubmitFlag")="29d98Sasphouseasp8asphnet" Then
				GBL_CHK_TempStr = ""
				ApplyFlag = 1
				checkFormData
				
				If GBL_CHK_Flag = 0 Then
					Response.WRite "<div class=alert>" & GBL_CHK_TempStr & "</div>" & VbCrLf
					JoinForm
				Else
					If saveFormData = 1 Then
						displayAccessFull
					Else
						Response.Write "<div class=alert>" & GBL_CHK_TempStr & "</div>" & VbCrLf
						JoinForm
					End If
				End If
				If Form_UpFlag = 1 Then Set Form_UpClass = Nothing
			Else
				JoinForm
			End If
		Else
			Response.Write "<div class=alert>��ֹ̳ͣ���û�ע���У�����ϵ����Ա��</div>"
		End If
	Else
		DisplayUserAgreement
	End If
Else
	Reg_Bind
End If
UserTopicBottomInfo
closeDataBase
SiteBottom

Sub Reg_CheckInfo
	
	Dim checkitem,checkvalue
	checkitem = Left(Request("checkitem"),30)
	checkvalue = Left(Request("checkvalue"),30)
	Select Case checkitem
		Case "username":
			If CheckUserNameExist(checkvalue) = 1 Then
				Response.Write "<span class=redfont>�û����ѱ�����ע��</span>"
			Else
				Response.Write "<span class=greenfont>��ϲ�����û�δ��ע��</span>"
			End If
		Case "email":
			If IsValidEmail(checkvalue) = false Then
				Response.Write "<span class=redfont>��Ч�������ַ��</span>"
			Else
				If CheckMailExist(checkvalue) = 1 Then
					Response.Write "<span class=redfont>�������ѱ������û�ʹ��</span>"
				Else
					Response.Write "<span class=greenfont>��֤ͨ��</span>"
				End If
			End If
	End Select

End Sub

Sub DisplayUserAgreement

	%><p><form action=<%=DEF_RegisterFile%> method=post>
	<input name="JoinFlag" type="hidden" value="dkls">
	<input name="action" type="hidden" value="<%=htmlencode(reg_action)%>">
	<input name="command" type="hidden" value="<%=htmlencode(reg_command)%>">
	<input type="hidden" value="<%
	If Request("u") <> "" Then
		Response.Write htmlencode(Request("u"))
	Else
		Response.Write reg_getrefer
	End If
	%>" name=u>
<!-- #include file=inc/User_Reg.asp -->
<input type="submit" value="��ͬ��" class="fmbtn btn_3">
<input type="button" value="��ͬ��" class="fmbtn btn_3" onclick="location.href='../Boards.asp';"></form>
<br /><br />
<div class=splitline></div>
<div class=title>�����ӵ���ʺţ�</div>
<div class=value2><a href="login.asp">����̳�ʺŵ�¼</a><%
If GetBinarybit(DEF_Sideparameter,10) = 1 Then%>
<span class=grayfont>������¼��</span><a href="<%=DEF_BBS_HomeUrl%>app/qqlogin/login.asp"><img src="<%=DEF_BBS_HomeUrl%>images/app/1.gif" border="0" style="position:absolute;" /><span style="padding-left:18px;">QQ��¼</span></a><%
End If%></div>
	<%

End Sub

Function JoinForm%>

	<script type="text/javascript">
	<!--
	var user_DEF_BBS_HomeUrl = "<%=DEF_BBS_HomeUrl%>";
	var user_DEF_faceMaxNum = <%=DEF_faceMaxNum%>;
	var user_DEF_AllDefineFace = <%=DEF_AllDefineFace%>;
	var user_ShowTestNumber = <%=ShowTestNumber%>;
	var user_DEF_RegisterFile = "<%=replace(replace(DEF_RegisterFile,"\","\\"),"""","\""")%>";
	var user_DEF_AllFaceMaxWidth = <%=DEF_AllFaceMaxWidth%>;
	var user_DEF_ShortestUserName = <%=DEF_ShortestUserName%>;
	-->
	</script>
	<script src="inc/register.js" type="text/javascript"></script>

<form action=<%=DEF_RegisterFile%> method=post name=LeadBBSFm id="LeadBBSFm" onSubmit="submitonce(this);return ValidationPassed;">
	<input type=hidden value="<%Response.Write htmlencode(Request("u"))%>" name=u>
	<input name="action" type="hidden" value="<%=htmlencode(reg_action)%>">
	<input name="command" type="hidden" value="<%=htmlencode(reg_command)%>">
	<div class=title><%If reg_action <> "bind" then %>���û�ע��<%
			Else%>��������<%
			End If%></div>
	<br>
	<%If DEF_UserNewRegAttestMode = 1 Then Response.Write "<span class=redfont>ע�⣺��ע����û���Ҫ�������ȡ��֤�뼤�����<br>ϸ��д������Ч�����ַ��</span>"%>

			<table border=0  cellpadding="0" cellspacing="0" class="blanktable">
			<tr>
				<td>
					�û����ƣ� 
				</td>
				<td>
					<input class='fminpt input_3' maxlength=14 name="Form_username" size="14" onchange="reg_checkinfo('username',this.value);" value="<% If Form_username<>"" Then Response.Write Server.HtmlEncode(Form_Username)%>">
					<span id="reg_check_username"></span>
				</td>
			</tr>
			<tr>
				<td>
					������룺 
				</td>
				<td>
					<input class=fminpt name=SubmitFlag type=hidden value="29d98Sasphouseasp8asphnet">
					<input class=fminpt name=JoinFlag type=hidden value="3kkdk">
					<input class='fminpt input_3' maxlength=20 name="Form_password1" size=14 type=password value="<% If Form_password1<>"" Then Response.Write Server.HtmlEncode(Form_password1)%>">
				</td>
			</tr>
			<tr>
				<td>
					��֤���룺 
				</td>
				<td>
					<input class='fminpt input_3' maxlength=20 name="Form_password2" size=14 type=password value="<% If Form_password2<>"" Then Response.Write Server.HtmlEncode(Form_password2)%>">
				</td>
			</tr>
			<tr>
				<td>
					�����ʼ��� 
				</td>
				<td>
					<input class='fminpt input_3' maxlength=60 name=Form_mail size=36 onchange="reg_checkinfo('email',this.value);" value="<% If Form_mail<>"" Then Response.Write Server.HtmlEncode(Form_mail)%>">
					<span id="reg_check_email"></span>
				</td>
			</tr>
			<tr>
				<td>
					������ʾ�� 
				</td>
				<td>
	<script type="text/javascript">
	<!--
	function sel_question(list)
	{
		alert('a');
		//if(list.value!='0'&&list.value!='99')$id('Form_Question').value=list.value;if(this.value=='99')$id('Form_Question').type='text';
	}
	-->
	</script>
					<select name="sel_question" onchange="if(this.value!=''&&this.value!='99')$id('Form_Question').value=this.value;if(this.value=='99'){this.style.display='none';$id('Form_Question').style.display='block';}else{$id('Form_Question').style.display='none';}">
						<option value="" selected>--ѡ������--</option>
						<option value="�ҵļ����ǣ�">�ҵļ����ǣ�</option>
						<option value="����������֣�">����������֣�</option>
						<option value="��ϲ���Ե�ʳƷ��">��ϲ���Ե�ʳƷ��</option>
						<option value="99">�Զ���...</option>
					</select>
					<div class=value2><input class='fminpt input_3' type="text" style="display:none;float:right;" maxlength=20 id=Form_Question name=Form_Question size=36 value="<% If Form_Question<>"" Then Response.Write Server.HtmlEncode(Form_Question)%>">
					<div>
				</td>
			</tr>
			<tr>
				<td>
					��ʾ�𰸣�
				</td>
				<td>
					<input class='fminpt input_3' maxlength=20 name=Form_Answer size=36 value="<% If Form_Answer<>"" Then Response.Write Server.HtmlEncode(Form_Answer)%>">
				</td>
			</tr>
			</table>
			<table border=0  cellpadding="0" cellspacing="0" class="blanktable">
			<tr>
			<td>
				<label><input class="fmchkbox" type="checkbox" name="moreinfo" value="1" onclick="if(this.checked){$id('reg_more_info').style.display='block';}else{$id('reg_more_info').style.display='none';}" />��д��������
				</label>
			</td></tr></table>
			<table border=0  cellpadding="0" cellspacing="0" class="blanktable" id="reg_more_info" style="display:none">
			<tr>
				<td>
					������ҳ��
				</td>
				<td>
					<input class=fminpt maxlength=250 name=Form_homepage size=36 value="<% If Form_homepage<>"" Then Response.Write Server.HtmlEncode(Form_homepage)%>">
				</td>
			</tr>
			<tr>
				<td>
					��ϵ��ַ��
				</td>
				<td>
					<input class=fminpt maxlength=150 name=Form_address size=36 value="<% If Form_address<>"" Then Response.Write Server.HtmlEncode(Form_address)%>">
				</td>
			</tr>
			<tr>
				<td>
					ICQ���룺
				</td>
				<td>
					<input class=fminpt maxlength=12 name=Form_icq size=14 value="<% If Form_icq<>"" Then Response.Write Server.HtmlEncode(Form_icq)%>">
				</td>
				<td rowspan="4" valign=bottom>&nbsp;<%If Form_userphoto<>"" and isNumeric(Form_userphoto) Then%><img name=faceimg id=faceimg src=<%=DEF_BBS_HomeUrl%>images/face/<%=string(4-len(cstr(Form_userphoto)),"0")&Form_userphoto%>.gif align=middle width=62 height=62><%Else%><img name=faceimg id=faceimg src=<%=DEF_BBS_HomeUrl%>images/blank.gif align=middle><%End If%></td>
			</tr>
			<tr>
				<td>
					QQ���룺
				</td>
				<td>
					<input class=fminpt maxlength=12 name=Form_oicq size=14 value="<% If Form_oicq<>"" Then Response.Write Server.HtmlEncode(Form_oicq)%>">
				</td>
			</tr>
			<tr>
				<td>
					�Ա�
				</td>
				<td>
					<label>
						<input class=fmchkbox type=radio name=Form_sex value=�� <%If Form_sex = "��" Then Response.Write " checked"%>>��</label>
					<label>
						<input class=fmchkbox type=radio name=Form_sex value=Ů <%If Form_sex = "Ů" Then Response.Write " checked"%>>Ů</label>
					<label>
						<input class=fmchkbox type=radio name=Form_sex value=�� <%If Form_sex = "��" Then Response.Write " checked"%>>����</label>
				</td>
			</tr>
			<tr>
				<td>
					�û�ͷ��
				</td>
				<td>
					<input class=fminpt onchange="javascript:changeface();" maxlength=4 name=Form_userphoto size=4 value="<% If Form_userphoto<>"" Then Response.Write Server.HtmlEncode(string(4-len(cstr(Form_userphoto)),"0")&Form_userphoto)%>">
					<a href="UserModify.asp?action=face" target=_blank onclick="return(pub_command('ѡ��ͷ��',this,'anc_delbody',''));">ͷ��һ����</a>
				</td>
			</tr><%If DEF_AllDefineFace <> 0 and DEF_AllDefineFace <> 2 Then%>
			<tr>
				<td>
					�Զ�ͷ��
				</td>
				<td>
					<input class=fminpt onchange="javascript:changeface2();" maxlength=250 name=Form_FaceUrl size=36 value="<%=HtmlEncode(Form_FaceUrl)%>">
				</td>
			</tr>
			<tr>
				<td>
					ͷ���С��
				</td>
				<td>
					��: <input class=fminpt onchange="javascript:changeface2();" maxlength=<%=len(DEF_AllFaceMaxWidth)%> name=Form_FaceWidth size=3 value="<%=HtmlEncode(Form_FaceWidth)%>">(20-<%=DEF_AllFaceMaxWidth%>)
					��: <input class=fminpt onchange="javascript:changeface2();" maxlength=<%=len(DEF_AllFaceMaxWidth)%> name=Form_FaceHeight size=3 value="<%=HtmlEncode(Form_FaceHeight)%>">(20-<%=DEF_AllFaceMaxWidth%>)
				</td>
			</tr><%End If%>
			<tr>
				<td>
					����
				</td>
				<td>
					
					<input class=fminpt maxlength=4 name=Form_byear size=4 value="<% If Form_byear<>"" Then
						Response.Write Server.HtmlEncode(Form_byear)
					Else
						Response.Write "19"
					End If%>"> �� 
					<input class=fminpt maxlength=2 name=Form_bmonth size=2 value="<% If Form_bmonth<>"" Then Response.Write Server.HtmlEncode(Form_bmonth)%>">
					�� <input class=fminpt maxlength=2 name=Form_bday size=2 value="<% If Form_bday<>"" Then Response.Write Server.HtmlEncode(Form_bday)%>">
					��</td>
			</tr>
			<tr>
				<td>
					����ǩ����
				</td>
				<td>
					<textarea class=fmtxtra name=Form_Underwrite rows=5 cols=34><%If Form_Underwrite <> "" Then Response.Write VbCrLf & htmlEncode(Form_Underwrite)%></textarea>
				</td>
			</tr>
			</table>
			
			<table border=0  cellpadding="0" cellspacing="0" class="blanktable">
			<%If LMT_RegVerifyQuestion <> "" Then%>
			<tr>
				<td>
					ע����֤��<br />
					<span class="grayfont">����ʾ��д</span>
				</td>
				<td>
						<p>
						<%=LMT_RegVerifyQuestion%>
						</p>
						<input class='fminpt input_2' maxlength=100 name="Form_RegVerifyAnswer" size="14" value="<% If Form_RegVerifyAnswer<>"" Then Response.Write Server.HtmlEncode(Form_RegVerifyAnswer)%>">
				</td>
			</tr>
			<%End If%>
			<%If ShowTestNumber > 2 Then%>
			<tr>
				<td>
					�� ֤ �룺
				</td>
				<td>
						<%displayVerifycode%>
				</td>
			</tr><%End If%>
			<tr>
				<td>&nbsp;</td>
				<td>
					<input name=submit type=submit value="����" class="fmbtn btn_2">
					<input name=b1 type=reset value="��д" class="fmbtn btn_2">
				</td>
			</tr>
			</table>
</form>
<%
End Function

Function saveFormData

	Dim Rs
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	Rs.Open sql_select("Select * from LeadBBS_User",1),con,2,2
	Rs.Addnew
	Rs("UserName") = Form_UserName
	Rs("Mail") = Trim(Form_Mail)
	Rs("Address") = Trim(Form_Address)
	Rs("Sex") = Form_Sex
	If Form_ICQ<>"" Then Rs("ICQ") = Form_ICQ
	If Form_OICQ<>"" Then Rs("OICQ") = Form_OICQ
	Rs("Userphoto") = Form_Userphoto
	Rs("Homepage") = Trim(Form_Homepage)
	Rs("Underwrite") = Form_Underwrite
	Rs("PrintUnderwrite") = Form_PrintUnderwrite
	Rs("Pass") = MD5(Form_Password1)
	If Len(Form_birthday) = 14 Then
		Rs("birthday") = Form_birthday
		Dim Temp
		temp = cCur(Left(Form_birthday,4))
		If temp > 1950 and temp < 2050 Then Rs("NongLiBirth") = GetNongLiTimeValue(ConvertToNongLi(RestoreTime(Form_birthday)))
	End If

	REM ��������
	Rs("ApplyTime") = Form_ApplyTime
	Rs("IP") = Form_IP
	Rs("UserLevel") = Form_UserLevel
	Rs("Officer") = Form_Officer
	Rs("Points") = DEF_User_RegPoints
	Rs("Sessionid") = 0
	Rs("Online") = Form_Online
	Rs("Prevtime") = Form_Prevtime
	Rs("Answer") = MD5(Form_Answer)
	Rs("Question") = Form_Question
	Rs("LastDoingTime") = Form_ApplyTime
	Rs("LastWriteTime") = Form_ApplyTime
	If DEF_UserNewRegAttestMode > 0 Then
		Rs("UserLimit") = 1
	Else
		Rs("UserLimit") = 0
	End If

	If Form_FaceWidth < 20 Then Form_FaceWidth = 20
	If Form_FaceHeight < 20 Then Form_FaceHeight = 20
	If DEF_AllDefineFace <> 0 Then
		Rs("FaceUrl") = Form_FaceUrl & ""
		Rs("FaceWidth") = Form_FaceWidth
		Rs("FaceHeight") = Form_FaceHeight
	Else
		Rs("FaceWidth") = 20
		Rs("FaceHeight") = 20
	End If
	Rs("LastAnnounceID") = 0
	Rs.Update

	Rs.Close
	Set Rs = Nothing
	
	Set Session(DEF_MasterCookies & "UDT") = Nothing
	Session(DEF_MasterCookies & "UDT") = ""
	
	Set Rs = LDExeCute(sql_select("Select ID from LeadBBS_User Where UserName='" & Replace(Form_UserName,"'","''") & "'",1),0)
	If Not Rs.Eof Then
		Form_ID = Rs(0)
	Else
		Form_ID = 0
	End If
	Rs.Close
	Set Rs = Nothing
	saveFormData = 1

	Dim Form_ExpiresTime
	If DEF_UserActivationExpiresDay > 0 and DEF_UserActivationExpiresDay < 3650 Then
		Form_ExpiresTime = GetTimeValue(DateAdd("d",DEF_UserActivationExpiresDay,DEF_Now))
	Else
		Form_ExpiresTime = 0
	End If
	If DEF_UserNewRegAttestMode > 0 Then
		If DEF_UserNewRegAttestMode = 1 Then
			Randomize
			AttestNumber = Right(Fix(Rnd*Timer)+Fix(Rnd*cCur(GetTimeValue(DEF_Now))) + 10000,10)
		End If
		CALL LDExeCute("insert into LeadBBS_SpecialUser(UserID,UserName,BoardID,Assort,ndatetime,ExpiresTime) values(" & Form_ID & ",'" & Replace(Form_UserName,"'","''") & "'," & AttestNumber & ",6," & GetTimeValue(DEF_Now) & "," & Form_ExpiresTime & ")",1)
	End If
	
	BindRegUser

	CALL LDExeCute("Update LeadBBS_SiteInfo Set UserCount=UserCount+1",1)
	UpdateStatisticDataInfo 1,1,1
	UpdateStatisticDataInfo Form_UserName,12,0

	SendNewMessage "[LeadBBS]",Form_UserName,"��ӭ������̳��","������̳�Ѿ�ע��ɹ�����ӭ��Ϊ���ǵ�һԱ��",GBL_IPAddress
	SendRegMail

End Function

Sub SendRegMail

	Dim HomeUrl
	HomeUrl = "http://"&Request.ServerVariables("server_name")
	If Request.ServerVariables("SERVER_PORT") <> "80" Then HomeUrl = HomeUrl & ":" & Request.ServerVariables("SERVER_PORT")
	HomeUrl = Lcase(HomeUrl & Request.Servervariables("SCRIPT_NAME"))
	HomeUrl = Replace(HomeUrl,"user/" & LCase(DEF_RegisterFile),"")

	Dim MailBody,Topic,TextBody
	Topic = "����" & DEF_SiteNameString & "�ĳɹ�ע������"
	MailBody = "<html>"
	TextBody = ""
	MailBody = MailBody & "<title>ע����Ϣ</title>"
	MailBody = MailBody & "<BODY>"
	MailBody = MailBody & "<table BORDER=0 WIDTH=95% ALIGN=CENTER><TBODY><tr>"
	MailBody = MailBody & "<TD valign=MIDDLE ALIGN=TOP><HR WIDTH=100% SIZE=1>"
	TextBody = TextBody & "------------------------------------------" & VbCrLf
	MailBody = MailBody & VbCrLf & "<b>" & htmlencode(Form_UserName)&"������</b>��<br><br>"
	TextBody = TextBody & htmlencode(Form_UserName)&"�����ã�" & VbCrLf & VbCrLf
	MailBody = MailBody & "лл��ע�᱾��̳������������ע����Ϣ��<br><br>"
	TextBody = TextBody & "лл��ע�᱾��̳������������ע����Ϣ��" & VbCrLf & VbCrLf
	MailBody = MailBody & "�û�����"&htmlencode(Form_UserName)&"<br>"
	TextBody = TextBody & "�û�����"&htmlencode(Form_UserName) & VbCrLf
	MailBody = MailBody & "�ܡ��룺" & Form_Password1 & "<br>"
	TextBody = TextBody & "�ܡ��룺" & Form_Password1 & VbCrLf
	If DEF_UserNewRegAttestMode = 1 Then
		MailBody = MailBody & "��֤�룺" & AttestNumber & "<br>"
		TextBody = TextBody & "��֤�룺" & AttestNumber & VbCrLf
		MailBody = MailBody & "<p><b><a href=" & HomeUrl & "User/UserGetPass.asp?act=active&user=" & urlencode(Form_UserName) & ">���������������ע����Ϣ���������������û���</a></b><br>"
		TextBody = TextBody & VbCrLf & VbCrLf & "������������ַ������������ע����Ϣ���������������û���" & VbCrLf & HomeUrl & "User/UserGetPass.asp?act=active&user=" & urlencode(Form_UserName) & VbCrLf & VbCrLf
	Else
		MailBody = MailBody & "<p>��ע����û���ȴ���վ����Ա������֤���ܳ�Ϊ��ʽ�û�����ͨ��֮ǰ�ڹ���ʹ���ϻ���һЩ���ơ�<br>"
		TextBody = TextBody & VbCrLf & VbCrLf & "��ע����û���ȴ���վ����Ա������֤���ܳ�Ϊ��ʽ�û�����ͨ����֤֮ǰ�ڹ���ʹ���ϻ���һЩ���ơ�" & VbCrLf
	End If
	MailBody = MailBody & "<br><br>"
	MailBody = MailBody & "<CENTER><font COLOR=RED><a href=""" & HomeUrl & """>��ӭ����������̳��</a></font>"
	MailBody = MailBody & "</td></tr></table><br><HR WIDTH=95% SIZE=1>"
	MailBody = MailBody & "<p ALIGN=CENTER>" & DEF_SiteNameString & " <a href=http://www.leadbbs.com target=_blank class=NavColor>" & DEF_Version & "</a></P>"
	TextBody = TextBody & VbCrLf & "��̳��ַ��" & HomeUrl & VbCrLf
	TextBody = TextBody & "------------------------------------------" & VbCrLf
	MailBody = MailBody & "</body>"
	MailBody = MailBody & "</html>"
	Select Case DEF_BBS_EmailMode
		Case 1: If SendEasyMail(Form_Mail,Topic,MailBody,TextBody) = 1 Then
					Response.Write "<br><br>���ϳɹ����͵�����ע�����䣡"
				Else
					Response.Write "<br><br>��̳δ��ȷ�����ʼ����ͣ����Ϸ���ʧ�ܣ�"
				End If
		Case 2: If SendJmail(Form_Mail,Topic,MailBody) = 1 Then
					Response.Write "<br><br>���ϳɹ����͵�����ע�����䣡"
				Else
					Response.Write "<br><br>��̳δ��ȷ�����ʼ����ͣ����Ϸ���ʧ�ܣ�"
				End If
		Case 3: If SendCDOMail(Form_Mail,Topic,TextBody) = 1 Then
					Response.Write "<br><br>���ϳɹ����͵�����ע�����䣡"
				Else
					Response.Write "<br><br>��̳δ��ȷ�����ʼ����ͣ����Ϸ���ʧ�ܣ�"
				End If
		Case Else: 
	End Select

End Sub

Function displayAccessFull

	Response.Cookies(DEF_MasterCookies)("user") = CodeCookie(Form_Username)
	Response.Cookies(DEF_MasterCookies)("pass") = CodeCookie(Form_password1)
	Response.Cookies(DEF_MasterCookies).Domain = DEF_AbsolutHome
	CALL LDExeCute("Update LeadBBS_onlineUser set UserID=" & Form_ID & ",UserName='" & Replace(Form_Username,"'","''") & "',HiddenFlag=" & DEF_UserNewRegAttestMode & " where sessionID=" & session.sessionID,1)%>
	<div class=title>���Ѿ��ɹ�<%If reg_action = "bind" Then%>�����ʺ�����<%Else%>ע���Ϊ��̳�û�<%End If%>��3���Ӻ�ҳ�潫�Զ�������Ӧҳ�档</a></div>
	<%If DEF_UserNewRegAttestMode = 1 Then
		Response.Write "<div class='value2 greenfont'>ע����û�ֻ�������̳��Ȩ�ޣ������û�����֤���Ѿ��ɹ����͵�����ע�����䡣</div>" & VbCrLf
	ElseIf DEF_UserNewRegAttestMode = 2 Then
		Response.Write "<div class='value2 greenfont'>ע����û�ֻ�������̳��Ȩ�ޣ���ȴ���վ��Ա����������֤���ܳ�Ϊ��ʽ�û���</div>" & VbCrLf
	End If
	
	Dim u
	u = filterUrlstr(Request("u"))
	If u = "" Then u = DEF_BBS_HomeUrl & "Boards.asp"
	%><script type="text/javascript">
		function a_topage()
		{
			this.location.href = "<%=Replace(Replace(u,"\","\\"),"""","\""")%>"; 
		}
		setTimeout("a_topage()",3000);
		</script>

<%End Function

Sub Reg_Bind

	If reg_command = "bind" Then
		reg_BindExistUser
		Exit Sub
	End If
	
	%>
	<div class="title">��ѡ��: <a href="<%=DEF_RegisterFile%>?action=bind&command=bind&u=<%=Reg_GetRefer%>">��������̳�ʺ�</a> / <a href="<%=DEF_RegisterFile%>?action=bind&command=reg&u=<%=Reg_GetRefer%>">�����ʺ�����</div>
	<%

End Sub

Function Reg_GetRefer

	Dim HomeUrl,u
	HomeUrl = "http://"&Request.ServerVariables("server_name")
	u = filterUrlstr(Request.QueryString("u"))
	If Left(u,1) <> "/" and Left(u,1) <> "\" and Left(u,Len(HomeUrl)) <> HomeUrl Then u = ""
	If u = "" Then
		u = filterUrlstr(Lcase(Request.ServerVariables("HTTP_REFERER")))
		If Request.ServerVariables("SERVER_PORT") <> "80" Then HomeUrl = HomeUrl & ":" & Request.ServerVariables("SERVER_PORT")
		If Left(u,Len(HomeUrl)) <> Lcase(HomeUrl) Then u = ""
		If inStr(u,"/user/" & DEF_RegisterFile) > 0 Then u = ""
	End If
	Reg_GetRefer = htmlencode(u)

End Function

Sub reg_BindExistUser

	If request("SubmitFlag") = "" Then
		DisplayLoginForm("����дҪ�󶨵���̳�û���Ϣ:")
	Else
		If GBL_CHK_Flag = 1 and GBL_UserID > 0 Then
			If reg_CheckAppidForUserID(GBL_AppType,GBL_UserID) = 1 Then
				Response.Write "<div class=""redfont""><b><p>����ʧ��: </p></b>���˺��ѱ���.</div>"
			Else
				If reg_checkAppidBinded = 0 Then
					Response.Write "<div class=""redfont""><b>" & GBL_CHK_TempStr & "</b></div>"
				Else
					Form_ID = GBL_UserID
					BindRegUser
					Response.Write "<div class=""greenfont""><b>�ʺ��ѳɹ���!</b></div>"
				End If
			End If
		Else
			Response.Write "<div class=""redfont""><b><p>����ʧ��: </p></b>�����ʺ���Ϣ����.<br /> " & GBL_CHK_Tempstr & "</div>"
		End If
	%>
	
	<%
	End If

end Sub

Sub BindRegUser

	If reg_action = "bind" and (reg_command = "reg" or reg_command = "bind") Then
		CALL LDExeCute("insert into LeadBBS_AppLogin(UserID,appid,GuestName,appType,ndatetime,IPAddress,Token) values(" & Form_ID & ",'" & Replace(Form_App_appid,"'","''") & "','" & Replace(Form_App_GuestName,"'","''") & "'," & GBL_AppType & "," & GetTimeValue(DEF_Now) & ",'" & Replace(GBL_IPAddress,"'","''") & "','" & Replace(Form_App_Token,"'","''") & "')",1)
	End If

End Sub

Function reg_checkAppidBinded
	
	Dim appInfo
	Form_App_GuestName = LeftTrue(GBL_CHK_User,20)
	appInfo = Request.Cookies(DEF_MasterCookies & "_AppInfo")
	Select Case CStr(GBL_AppType)
		Case "1":					
			If inStr(appInfo,",") Then appInfo = Split(appInfo,",")
			If IsArray(appInfo) Then
				If Ubound(appInfo,1) = 2 Then
					Form_App_Token = LeftTrue(appInfo(1),64)
					Form_App_appid = LeftTrue(appInfo(2),64)
				End If
			End If
			If Len(Form_App_appid) < 16 or Form_App_GuestName = "" Then
				GBL_CHK_TempStr = "����ʧ��:QQ������Ϣ�Ѿ�ʧЧ,�����µ�¼. <br>" & VbCrLf
				reg_checkAppidBinded = 0
				Exit Function
			End If
		Case else
			GBL_CHK_TempStr = "����ʧ��:δ֪�Ļ�����. <br>" & VbCrLf
			reg_checkAppidBinded = 0
			Exit Function
	End Select
	If reg_CheckAppid(GBL_AppType,Form_App_appid) = 1 Then
		GBL_CHK_TempStr = "����ʧ��:�˻����ʺ��ѱ��󶨻�����. <br>" & VbCrLf
		reg_checkAppidBinded = 0
		Exit Function
	End If
	reg_checkAppidBinded = 1

End Function
%>