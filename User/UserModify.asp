<!-- #include file=../inc/BBSsetup.asp -->
<!-- #include file=../inc/User_Setup.asp -->
<!-- #include file=../inc/Board_popfun.asp -->
<!-- #include file=../inc/ubbcode.asp -->
<!-- #include file=../inc/Limit_fun.asp -->
<!-- #include file=inc/User_fun.asp -->
<!-- #include file=inc/UserTopic.asp -->
<!-- #include file=../inc/Constellation2.asp -->
<!-- #include file=../a/inc/upload1_fun.asp -->
<%DEF_BBS_HomeUrl = "../"%>
<!-- #include file=../inc/Upload_Fun.asp -->
<%

Dim Form_RevMessageFlag,Form_SoundFlag,Form_UserLimit,Form_Action,AjaxFlag,upload_step
Dim EnableUpload

Main

Sub Page_Expires

	Response.Expires = 0
	Response.ExpiresAbsolute = DEF_Now - 1
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private" 
	Response.CacheControl = "no-cache"

End Sub

Sub Main

	GBL_HeadResource = "<link rel=""stylesheet"" type=""text/css"" href=""" & DEF_BBS_HomeUrl & "inc/js/imgareaselect/imgareaselect-default.css"" />"
	Form_Action = left(Request.QueryString("action"),20)
	If Form_Action = "face" Then
		User_FaceList
		Exit Sub
	End If
	AjaxFlag = left(Request.QueryString("AjaxFlag"),20)
	If AjaxFlag <> "1" then AjaxFlag = ""
	Page_Expires
	Form_FaceWidth = DEF_AllFaceMaxWidth
	Form_FaceHeight = DEF_AllFaceMaxWidth

	CursorLocation = 3
	initDatabase
	if AjaxFlag <> "1" Then BBS_SiteHead DEF_SiteNameString & " - �޸�����",0,"<span class=navigate_string_step>�޸�����</span>"
	UpdateOnlineUserAtInfo GBL_board_ID,"�޸�����"
	GBL_CHK_TempStr=""
	If GBL_UserID = 0 Then GBL_CHK_TempStr = GBL_CHK_TempStr & "δ��¼���������.<br>" & VbCrLf
	
	Form_RevMessageFlag = 0
	
	if AjaxFlag <> "1" Then UserTopicTopInfo("user")
	
	User_GetStartValue
	
	If GBL_CHK_Flag=1 Then
		If GBL_CHK_TempStr = "" Then
			CheckUserModifyLimit
			CheckisBoardMaster
			User_CheckEnableUpload
			If GBL_CHK_TempStr = "" Then
				GetUserData(GBL_UserID)
				If Form_Submitflag="29d98Sasphouseasp8asphnet" Then
					GBL_CHK_TempStr = ""
					checkFormData
					If GBL_CHK_Flag = 0 Then
						If ajaxflag <> "1" then
							Response.Write "<div class='title redfont'>" & GBL_CHK_TempStr & "</div>" & VbCrLf
						else
							Response.Write "<script>alert(""" & replace(GBL_CHK_TempStr,"<br>","") & """)</script>" & VbCrLf
						end if
						JoinForm
					Else
						If SaveFormData = 1 Then
							displayAccessFull
						Else
							Response.Write "<div class='title redfont'>" & GBL_CHK_TempStr & "</div>" & VbCrLf
							JoinForm
						End If
					End If
					If Form_UpFlag = 1 Then Set Form_UpClass = Nothing
				Else
					JoinForm
				End If
			Else
				Response.WRite "<div class='title redfont'>" & GBL_CHK_TempStr & "</div>" & VbCrLf
			End If
		Else
					DisplayLoginForm(GBL_CHK_TempStr)
		End If
	Else
		If Form_Submitflag = "" Then
			DisplayLoginForm("���ȵ�¼")
		Else
			DisplayLoginForm("<span class=redfont>" & GBL_CHK_TempStr & "</span>")
		End If
	End If
	
	closeDataBase
	if AjaxFlag <> "1" Then UserTopicBottomInfo
	if AjaxFlag <> "1" Then SiteBottom

End Sub

Function User_FaceList

	Dim pagen,First,n
	pagen=10
	first = Left(Request("first"),14)
	If isNumeric(first)=0 or isNull(first) then first=1
	first = Fix(cCur(first))
	If first<1 or first>DEF_faceMaxNum then first=1
	If first<>1 then first=cint(first)
	%>
	<table align=center cellpadding="0" cellspacing="0" class="blanktable">
		<%
		Dim t,t2
		for n=first to first+pagen-1
			If n>DEF_faceMaxNum then exit for
			t = string(4-len(cstr(n)),"0")
			%>
			<tr align="center">
				<td>
					<%=t&n%>
			</td>
				<td>
					<a href=#1 onclick="user_setface('<%=t&n%>')"><img src="<%=DEF_BBS_HomeUrl & "images/face/"&t&n%>.gif"></td>
					<%
			n=n+1
			t = string(4-len(cstr(n)),"0")
			if n>DEF_faceMaxNum then exit for
			%><td><%=t&n%></td><td><a href="javascript:;" onclick="user_setface('<%=t&n%>');return false;"><img src="<%=DEF_BBS_HomeUrl & "images/face/"&t&n%>.gif"></td></tr>
			<%
		next%>
		<tr>
			<td colspan="4" align="center">
				<%
			If first-pagen>0 then
				%><a href="javascritp:;" onclick="getAJAX('UserModify.asp?action=face','first=1',$$('ajaxitembody')[0].id);return false;"><<��ҳ</a>
				<a href="javascritp:;" onclick="getAJAX('UserModify.asp?action=face','first=<%=first-pagen%>',$$('ajaxitembody')[0].id);return false;">��һҳ</a> <%
			Else
				Response.Write "<span class=""grayfont""><<��ҳ ��һҳ</span> " & VbCrLf
			End If

			If first+pagen<DEF_faceMaxNum then
				%><a href="javascritp:;" onclick="getAJAX('UserModify.asp?action=face','first=<%=first+pagen%>',$$('ajaxitembody')[0].id);return false;">��һҳ</a>
				<a href="javascritp:;" onclick="getAJAX('UserModify.asp?action=face','first=<%=DEF_faceMaxNum-pagen+1%>',$$('ajaxitembody')[0].id);return false;">βҳ>></a><%
			Else
				Response.Write "<span class=""grayfont"">��һҳ βҳ</span>" & VbCrLf
			End If%>
			</td>
		</tr>
	</table>

<%End Function

Sub User_CheckEnableUpload

	EnableUpload = 1
	
	If DEF_AllDefineFace = 0 Then EnableUpload = 0
	If DEF_UploadFaceNeedPoints > 0 and DEF_UploadFaceNeedPoints > GBL_CHK_Points Then
		EnableUpload = 0
	Else
		If DEF_FaceMaxBytes > 0 and (GBL_CHK_OnlineTime >= DEF_NeedOnlineTime or DEF_NeedOnlineTime = 0) and (DEF_AllDefineFace = 1 or DEF_AllDefineFace = 2) Then
		Else
			EnableUpload = 0
		End If
	End If
	
	'If DEF_EnableGFL = 0 then EnableUpload = 0

End Sub

Function GetUserData(ID)

	Dim Rs
	Set Rs = LDExeCute(sql_select("Select * from LeadBBS_User where id=" & ID,1),0)
	If Rs.Eof Then
		GetUserData = 0
		GBL_CHK_Flag = 0
		Rs.Close
		Set Rs = Nothing
		Exit Function
	End If
	Form_UserName = Rs("UserName")
	Old_Form_UserName = Form_UserName
	Form_Mail = Rs("Mail")
	Form_Address = Rs("Address")
	Form_Sex = Rs("Sex")
	Form_ICQ = Rs("ICQ")
	Form_OICQ = Rs("OICQ")
	Form_Userphoto = Rs("Userphoto")
	Form_Homepage = Rs("Homepage")
	Form_Underwrite = Rs("Underwrite")
	Form_Pass = Rs("Pass")
	'Form_Password1 = Rs("Pass")
	'Form_Password2 = Form_Password1
	Form_birthday = Rs("birthday")
	If len(Form_birthday)=14 Then
		Form_ApplyTime = RestoreTime(Rs("birthday"))
		Form_bday = day(Form_ApplyTime)
		Form_byear = year(Form_ApplyTime)
		Form_bmonth = month(Form_ApplyTime)
	End If

	REM ��������
	Form_ApplyTime = Rs("ApplyTime")
	Form_IP = Rs("IP")
	Form_UserLevel = Rs("UserLevel")
	Form_Officer = Rs("Officer")
	Form_Points = Rs("Points")
	Form_Sessionid = Rs("Sessionid")
	Form_Online = Rs("Online")
	Form_Prevtime = Rs("Prevtime")
	Form_ShowFlag = ccur(Rs("ShowFlag"))
	If Form_ShowFlag = 1 Then
		Form_ShowFlag = "1"
	Else
		Form_ShowFlag = "0"
	End If
	Form_NotSecret = ccur(Rs("NotSecret"))
	If Form_NotSecret = 1 Then
		Form_NotSecret = "1"
	Else
		Form_NotSecret = "0"
	End If
	
	If DEF_AllDefineFace <> 0 Then
		Form_FaceUrl = Rs("FaceUrl")
		Form_FaceWidth = Rs("FaceWidth")
		Form_FaceHeight = Rs("FaceHeight")
		Form_FaceUrl_Old = Form_FaceUrl
		Form_FaceWidth_Old = Form_FaceWidth
		Form_FaceHeight_Old = Form_FaceHeight
	End If
	Form_UserTitle = Rs("UserTitle")
	Form_UserLimit = Rs("UserLimit")
	Form_RevMessageFlag = GetBinaryBit(Form_UserLimit,13)
	Form_SoundFlag = GetBinaryBit(Form_UserLimit,17)
	Form_Question = Rs("Question")
	Form_Answer = Rs("Answer")
	OLd_Form_Question= Form_Question
	OLd_Form_Answer = Form_Answer
	Rs.Close
	Set Rs = Nothing
	GetUserData = 1
	GBL_CHK_Flag = 1

End Function

Sub Modify_NavInfo

	Response.Write "<div class='user_item_nav fire'><ul>"
	If Form_Action = "base" or Form_Action = "" Then
		Response.Write "	<li><span class=navactive><span>�û�����</span></span></li>"
	Else
		Response.Write "	<li><a href=UserModify.asp>�û�����</a></li>"
	End If
	If Form_Action = "uploadface" Then
		Response.Write "	<li><span class=navactive>�趨ͷ��</span></li>"
	Else
		Response.Write "	<li><a href=UserModify.asp?action=uploadface>�趨ͷ��</a></li>"
	End If
	%>
	</ul></div>
	<%

End Sub

Function JoinForm

	if AjaxFlag = "1" then exit function
%>
	<script type="text/javascript">
	var user_DEF_BBS_HomeUrl = "<%=DEF_BBS_HomeUrl%>";
	var user_DEF_faceMaxNum = <%=DEF_faceMaxNum%>;
	var user_DEF_AllFaceMaxWidth = <%=DEF_AllFaceMaxWidth%>;
	var user_DEF_AllDefineFace = <%=DEF_AllDefineFace%>;
	var user_DEF_RegisterFile = "<%=replace(replace(DEF_RegisterFile,"\","\\"),"""","\""")%>";
	-->
	</script>
	<script src="inc/usermodify.js?ver=<%=DEF_Jer%>" type="text/javascript"></script>

	<%Modify_NavInfo%>
	<div class=clear></div>
			<%If EnableUpload = 0 Then%>
			<form action=UserModify.asp?action=<%=urlencode(Form_Action)%><%
			If Form_Action = "uploadface" then response.Write "&ajaxflag=1"
			%> method=post name=LeadBBSFm id="LeadBBSFm" onSubmit="submitonce(this);return ValidationPassed">
			<%Else%>
			<form action=UserModify.asp?dontRequestFormFlag=1&action=<%=urlencode(Form_Action)%><%
			If Form_Action = "uploadface" then response.write "&ajaxflag=1 target=hidden_frame"%> method=post name=LeadBBSFm id=LeadBBSFm enctype="multipart/form-data" onSubmit="submitonce(this);return ValidationPassed">
			<%End If%>
			<input class='fminpt input_2' name=SubmitFlag type=hidden value="29d98Sasphouseasp8asphnet">
			<table border=0 cellpadding="0" class="blanktable">
		<%If Form_Action <> "uploadface" then%>
			<tr>
				<td>
					*�û����ƣ� 
				</td>
				<td>
					<%If inStr(Old_Form_UserName,"#") Then%>
					<div class=value2>��ʹ�õ�����ʱ�û��� <%Response.Write Server.HtmlEncode(Form_Username)%> ����д�µ��û�����
					</div>
					<div class=value2><input onchange="reg_checkinfo('username',this.value);" class='fminpt input_3' maxlength=20 name=Form_UserName size=36 value="<% If Form_UserName<>"" and inStr(Form_UserName,"#") = 0 Then Response.Write Server.HtmlEncode(Form_UserName)%>">
					<span id="reg_check_username"></span></div>
					<%Else%>
					<%Response.Write Server.HtmlEncode(Form_Username)
					End If%>
				</td>
			</tr>
			<%If Old_Form_Answer <> "" Then%>
			<tr>
				<td>
					*�ɵ����룺 
				</td>
				<td>
					<input class='fminpt input_2' maxLength=20 name="oldpass" size=14 type=password> ������ȷ��д
				</td>
			</tr>
			<%End If%>
			<tr>
				<td>
					�µ����룺 
				</td>
				<td>
					<input class='fminpt input_2' maxLength=20 name="Form_password1" size=14 type=password Value="<% If Form_password1<>"" Then Response.Write Server.HtmlEncode(Form_password1)%>"> 
					<%If inStr(Old_Form_UserName,"#") Then%>
					���������
					<%else%>
					���޸����벻����д
					<%end if%>
				</td>
			</tr>
			<tr>
				<td>
					��֤���룺 
				</td>
				<td>
					<input class='fminpt input_2' maxlength=20 name="Form_password2" size=14 type=password Value="<% If Form_password2<>"" Then Response.Write Server.HtmlEncode(Form_password2)%>">
				</td>
			</tr>
			
			<%If Old_Form_Answer = "" Then%>
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
	</script><div class=value2>
					<select name="sel_question" onchange="if(this.value!=''&&this.value!='99')$id('Form_Question').value=this.value;if(this.value=='99'){this.style.display='none';$id('Form_Question').style.display='block';}else{$id('Form_Question').style.display='none';}">
						<option value="" selected>--ѡ������--</option>
						<option value="�ҵļ����ǣ�">�ҵļ����ǣ�</option>
						<option value="����������֣�">����������֣�</option>
						<option value="��ϲ���Ե�ʳƷ��">��ϲ���Ե�ʳƷ��</option>
						<option value="99">�Զ���...</option>
					</select>
					</div>
					<div class=value2><input class='fminpt input_3' type="text" style="display:none;" maxlength=20 id=Form_Question name=Form_Question size=36 value="<% If Form_Question<>"" Then Response.Write Server.HtmlEncode(Form_Question)%>">
					<div>
				</td>
			</tr>
			<tr>
				<td>
					*��ʾ�𰸣�
				</td>
				<td>
					<input class='fminpt input_3' maxlength=20 name=Form_Answer size=36 value="<% If Form_Answer<>"" Then Response.Write Server.HtmlEncode(Form_Answer)%>">
					����վ��¼��������������
				</td>
			</tr>
			<%End If%>
			<tr>
				<td>
					*�����ʼ��� 
				</td>
				<td>
					<input class='fminpt input_3' onchange="reg_checkinfo('email',this.value);" maxLength=60 name=Form_mail size=36 Value="<% If Form_mail<>"" Then Response.Write Server.HtmlEncode(Form_mail)%>">
					<span id="reg_check_email"></span>
				</td>
			</tr>
			<tr>
				<td>
					��ҳ��ַ��
				</td>
				<td>
					<input class='fminpt input_3' maxlength=250 name=Form_homepage size=36 Value="<% If Form_homepage<>"" Then Response.Write Server.HtmlEncode(Form_homepage)%>">
				</td>
			</tr>
			<tr>
				<td>
					��ͥ��ַ��
				</td>
				<td>
					<input class='fminpt input_3' maxlength=150 name=Form_address size=36 Value="<% If Form_address<>"" Then Response.Write Server.HtmlEncode(Form_address)%>">
				</td>
			</tr>
			<tr>
				<td>
					ICQ��
				</td>
				<td>
					<input class='fminpt input_2' maxlength=10 name=Form_icq size=14 Value="<% If Form_icq<>"" and Form_icq <> "0" Then Response.Write Server.HtmlEncode(Form_icq)%>">
				</td>
			</tr>
			<tr>
				<td>
					QQ��
				</td>
				<td>
					<input class='fminpt input_2' maxlength=10 name=Form_oicq size=14 Value="<% If Form_oicq<>"" and Form_oicq <> "0" Then Response.Write Server.HtmlEncode(Form_oicq)%>">
				</td>
			</tr>
			<tr>
				<td>
					�Ա�
				</td>
				<td>
					<label>
						<input class=fmchkbox type=radio name=Form_sex value=�� <%If Form_sex = "��" Then Response.Write " checked"%>>��</label>
					<label">
						<input class=fmchkbox type=radio name=Form_sex value=Ů <%If Form_sex = "Ů" Then Response.Write " checked"%>>Ů</label>
					<label>
						<input class=fmchkbox type=radio name=Form_sex value=�� <%If Form_sex = "��" Then Response.Write " checked"%>>����</label>
				</td>
			</tr><%
		End If
		
		UploadFace

		If Form_Action <> "uploadface" then
			If DEF_UserEnableUserTitle = 1 and Form_UserLevel >= DEF_UserUserTitleNeedLevel Then%>
			<tr>
				<td>
					�û�ͷ�Σ�
				</td>
				<td>
					<input maxlength=18 name=Form_UserTitle size=36 class='fminpt input_2' Value="<% If Form_UserTitle<>"" Then Response.Write Server.HtmlEncode(Form_UserTitle)%>">
				</td>
			</tr><%End If%>
			<tr>
				<td>
					���գ� 
				</td>
				<TD align="left">
					<input class='fminpt input_1' maxlength=4 name=Form_byear size=4 Value="<% If Form_byear<>"" Then
						Response.Write Server.HtmlEncode(Form_byear)
					Else
						Response.Write "19"
					End If%>"> �� 
					<input class='fminpt input_1' maxlength=2 name=Form_bmonth size=2 Value="<% If Form_bmonth<>"" Then Response.Write Server.HtmlEncode(Form_bmonth)%>">
					�� <input class='fminpt input_1' maxlength=2 name=Form_bday size=2 Value="<% If Form_bday<>"" Then Response.Write Server.HtmlEncode(Form_bday)%>">
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
			<tr>
				<td>
					�Ƿ�����
				</td>
				<td>
					<input class=fmchkbox type=radio name=Form_ShowFlag value=0 <%If Form_ShowFlag = "0" Then Response.Write " checked"%>>��������
					<input class=fmchkbox type=radio name=Form_ShowFlag value=1 <%If Form_ShowFlag = "1" Then Response.Write " checked"%>>����
				</td>
			</tr>
			<tr>
				<td>
					�Ƿ��ܣ�
				</td>
				<td>
					<input class=fmchkbox type=radio name=Form_NotSecret value=0 <%If Form_NotSecret = "0" Then Response.Write " checked"%>>��Ϣ����
					<input class=fmchkbox type=radio name=Form_NotSecret value=1 <%If Form_NotSecret = "1" Then Response.Write " checked"%>>��Ϣ����
				</td>
			</tr>
			<tr>
				<td>
					�� �� Ϣ��
				</td>
				<td>
					<div class=value2>1.������������
					<input class=fmchkbox type=radio name=Form_RevMessageFlag value=0 <%If Form_RevMessageFlag = "0" Then Response.Write " checked"%>>���������˵Ķ���Ϣ
					<input class=fmchkbox type=radio name=Form_RevMessageFlag value=1 <%If Form_RevMessageFlag = "1" Then Response.Write " checked"%>>���޽��պ��Ѷ���Ϣ
					</div>
					<div class=value2>2.������ʾ���� <input class=fmchkbox type=radio name=Form_SoundFlag value=0 <%If Form_SoundFlag = "0" Then Response.Write " checked"%>>����������ʾ����Ϣ
					<input class=fmchkbox type=radio name=Form_SoundFlag value=1 <%If Form_SoundFlag = "1" Then Response.Write " checked"%>>��ֹ������ʾ����Ϣ
					</div>
				</td>
			</tr>
		<%end if%>
			<tr>
				<td>&nbsp;</td>
				<TD height="30">
					<div id="submitdiv" style="<%If DEF_EnableGFL = 0 and Form_Action = "uploadface" Then response.write "display:none;"%>">
					<input name=submit id=submit type=submit value="�޸�" onclick="form_onsubmit(this.form)" class="fmbtn btn_2">
					<input name=b1 type=reset value="��д" class="fmbtn btn_2">
					<div>
				</td>
			</tr>
			</table>
</form>
<div class=title>ʹ��˵����</div>
<ol>
<%If Form_Action <> "uploadface" then%>
<li>�Ƿ���ѡ����Ϊ���ܣ����������Լ��ĵ�ַ�����գ��ѣѣ��������ϣ����Ҳ���ʾ����¼�ľ���ʱ�䣬IP��ַʼ�ձ��ܣ��������������û�һ��</li>
<li>��Ϊ����״̬��ÿ�ε�¼���ߺ���������û�</li>
<li>�ﵽ<%=DEF_UserUserTitleNeedLevel%>����ע���û������Զ���ͷ�Σ����������޸�</li>
<li>����Ϣ����Ϊֻ���պ���ѡ��԰���������Ȩ���û���Ч</li>
<li>ĳЩ���øĶ�����Ҫ�Ժ�����µ�¼������Ч</li>
<%End If%>
<%If Form_Action = "uploadface" then%>
<li>�ﵽ<%=DEF_UploadFaceNeedPoints%><%=DEF_PointsName(0)%><%If DEF_NeedOnlineTime > 0 Then Response.Write "������" & Fix(DEF_NeedOnlineTime/60) & "����"%>����ʹ���Զ����ϴ�ͼƬ��Ϊͷ��</li>
<%
If DEF_EnableAttestNumber > 2 and DEF_AttestNumberPoints > 0 Then
	%><li>���趨<%=DEF_PointsName(0) & DEF_AttestNumberPoints%>���Ϸ�������֤��"<%
End If
%>
<%
	If DEF_AllDefineFace = 0 Then
		Response.Write "<li>��ֹ�Զ���ͷ��"
	ElseIf DEF_AllDefineFace = 2 Then
		Response.Write "<li>����վ��ͼƬ��Ϊͷ�񣬽�ֹ����վ��ͼƬ��Ϊͷ��"
	ElseIf DEF_AllDefineFace = 3 Then
		Response.Write "<li>����վ��ͼƬ��Ϊͷ�񣬽�ֹ�ϴ��Զ���ͷ��"
	End If%>
	<%
End If%>
<li>�༭ͼƬʱ�����ֽ�������</li>
</ol>
<%

End Function

Sub UploadFace

		If Form_Action = "uploadface" then%>
			<tr>
				<td colspan=2>
				<%If EnableUpload = 1 Then%>
				<a href="javascript:;" class="grayfont" onclick="$('#upload_type').value=1;<%if DEF_EnableGFL = 0 then%>$('#submitdiv').hide();<%end if%>$('#face_http3').show();$('#face_http').hide();$('#face_http2').show();$('#face_system').hide();$('#face_upload').show();">�ϴ�ͷ��</a> &nbsp; 
				<%End If%>
				<a href="javascript:;" class="grayfont" onclick="$('#upload_type').value=2;$('#submitdiv').show();$('#face_http3').show();$('#face_http2').hide();$('#face_http2').hide();$('#face_system').show();$('#face_upload').hide();">ʹ��ϵͳͷ��</a> &nbsp; 
				<%If DEF_AllDefineFace <> 0 Then%>
					<a href="javascript:;" class="grayfont" onclick="$('#upload_type').value=3;$('#submitdiv').show();$('#face_http3').hide();$('#face_http2').show();$('#face_http').show();$('#face_system').hide();$('#face_upload').hide();">ʹ������ͼƬ</a>
				<%End If%>
				</span>
				</td>
			</tr>
			<tr id="face_system"<%If EnableUpload = 1 Then response.write " style=""display:none"""%>>
				<td>
					ϵͳͷ��
				</td>
				<td>
					<input name=upload_type id=upload_type type=hidden value="-1">
					<input name=upload_step id=upload_step type=hidden value="">
					<input name=upload_x1 id=upload_x1 type=hidden value="">
					<input name=upload_x2 id=upload_x2 type=hidden value="">
					<input name=upload_y1 id=upload_y1 type=hidden value="">
					<input name=upload_y2 id=upload_y2 type=hidden value="">
					<input name=upload_filename id=upload_filename type=hidden value="">
					<input class='fminpt input_1' onchange="javascript:changeface();" maxlength=4 name=Form_userphoto size=4 Value="<% If Form_userphoto<>"" Then Response.Write Server.HtmlEncode(string(4-len(cstr(Form_userphoto)),"0")&Form_userphoto)%>">
					<a href="UserModify.asp?action=face" target=_blank onclick="return(pub_command('ѡ��ͷ��',this,'anc_delbody',''));">ͷ��һ����</a>
				</td>
			</tr>
		<%
			If DEF_AllDefineFace <> 0 Then%>
			<tr id="face_http" style="display:none;">
				<td>
					�Զ���ͷ���ַ��
				</td>
				<td>
					<input class='fminpt input_4' onchange="javascript:changeface2();" maxlength=250 id="Form_FaceUrl" name=Form_FaceUrl size=36 Value="<%=HtmlEncode(Form_FaceUrl)%>">
				</td>
			</tr>
			<%
			End If
				
			If EnableUpload = 1 Then%>
			<tr height=30 id="face_upload">
				<td> <span id="uptext" name="uptext">�ϴ�ͷ��</span></td>
				<td>
					<span id=FileStr><input type=file id=file size=11 name=userface class="fminpt uninit_upload" onchange="submituploadajax();"></span>
					<span class=grayfont>������ѡ��һ��GIF��JPGͼƬ�ϴ����༭����<%=FormatNumber((DEF_FaceMaxBytes/1024),0)%>K</span>
					
					<script type="text/javascript">
					init_uploadform();
					</script>
  					<script type="text/javascript" src="<%=DEF_BBS_HomeUrl%>inc/js/imgareaselect/imgareaselect.js?ver=<%=DEF_Jer%>"></script>
				</td>
			</tr><%
			End If
			
			If DEF_AllDefineFace <> 0 Then%>
			<tr id="face_http2" style="display:none;">
				<td>
					ͷ���С��
				</td>
				<td>
					��: <input class='fminpt input_1' onchange="javascript:changeface2();" maxlength=<%=len(DEF_AllFaceMaxWidth)%> name=Form_FaceWidth size=3 Value="<%=HtmlEncode(Form_FaceWidth)%>">(20-<%=DEF_AllFaceMaxWidth%>)
					��: <input class='fminpt input_1' onchange="javascript:changeface2();" maxlength=<%=len(DEF_AllFaceMaxWidth)%> name=Form_FaceHeight size=3 Value="<%=HtmlEncode(Form_FaceHeight)%>">(20-<%=DEF_AllFaceMaxWidth*2%>)
				</td>
			</tr>
			<tr id="face_http3">
				<td valign=top>
					ͷ��Ԥ����</td>
				<td valign=top>
					<div>
					<%If DEF_AllDefineFace = 0 or Form_FaceUrl & "" = "" Then
						If Form_userphoto<>"" and isNumeric(Form_userphoto) Then
							%><img name=faceimg id=faceimg src=<%=DEF_BBS_HomeUrl%>images/face/<%=string(4-len(cstr(Form_userphoto)),"0")&Form_userphoto%>.gif align=middle><%
						Else
							%><img name=faceimg id=faceimg src=<%=DEF_BBS_HomeUrl%>images/blank.gif align=middle><%
						End If%>
					<%Else%>
						<img name=faceimg id=faceimg src="<%=htmlencode(Form_FaceUrl)%>" align=middle width=<%=Form_FaceWidth%> height=<%=Form_FaceHeight%>>
					<%End If%>
					</div>
					
					<div style="float:left;margin-left:25px;<%If DEF_EnableGFL = 0 Then response.write "display:none;"%>" id=ajaxphoto>
					<img src=<%=DEF_BBS_HomeUrl%>images/blank.gif id="photo">
					</div>
					
					<div class=clear></div>
					<div id=selectok></div>
					<div id=selectinfo></div>
					</td>
			</td>
			<%
			End If
		end if

End Sub

Function displayAccessFull
	
	If AjaxFlag <> "1" then %>
	<p>
	<%If GBL_CHK_TempStr <> "" Then Response.Write "<font color=red class=redfont><b>������ʾ:" & GBL_CHK_TempStr & "</b></font><p>"%>
	<b><font color=Green class=greenfont>���������ѳɹ��޸ģ�</font></b><br><%
	End if

End Function

Function SaveFormData

	select case Form_Action
		case "uploadface":
		if AjaxFlag <> "1" then
			sql = "update LeadBBS_User set Userphoto=" & Replace(Form_Userphoto,"'","''")
			If DEF_AllDefineFace <> 0 Then
				sql = sql & ",FaceUrl='" & Replace(Form_FaceUrl,"'","''") & "'"
				sql = sql & ",FaceWidth=" & Replace(Form_FaceWidth,"'","''")
				sql = sql & ",FaceHeight=" & Replace(Form_FaceHeight,"'","''")
			Else
				Form_FaceUrl = ""
				Form_FaceWidth = ""
				Form_FaceHeight = ""
			End If
			sql = sql & " Where id=" & GBL_UserID
			CALL LDExeCute(SQL,1)
		end if
			
			If EnableUpload = 1 Then
				if User_ModifyUserFace = 2 then
					sql = "update LeadBBS_User set Userphoto=" & Replace(Form_Userphoto,"'","''")
					If DEF_AllDefineFace <> 0 Then
						sql = sql & ",FaceUrl='" & Replace(Form_FaceUrl,"'","''") & "'"
						sql = sql & ",FaceWidth=" & Replace(Form_FaceWidth,"'","''")
						sql = sql & ",FaceHeight=" & Replace(Form_FaceHeight,"'","''")
					Else
						Form_FaceUrl = ""
						Form_FaceWidth = ""
						Form_FaceHeight = ""
					End If
					sql = sql & " Where id=" & GBL_UserID
					CALL LDExeCute(SQL,1)
				end if
			End If
		Case else
			Form_RevMessageFlag = GetFormData("Form_RevMessageFlag")
			If Form_RevMessageFlag = "1" Then
				Form_RevMessageFlag = 1
			Else
				Form_RevMessageFlag = 0
			End If
		
			Form_SoundFlag = GetFormData("Form_SoundFlag")
			If Form_SoundFlag = "1" Then
				Form_SoundFlag = 1
			Else
				Form_SoundFlag = 0
			End If
		
			Form_UserLimit = SetBinaryBit(Form_UserLimit,13,Form_RevMessageFlag)
			Form_UserLimit = SetBinaryBit(Form_UserLimit,17,Form_SoundFlag)
		
			Dim Rs,temp
			Set Rs = LDExeCute(sql_select("Select * from LeadBBS_User where id=" & GBL_UserID,1),0)
			If Rs.Eof Then
				Rs.Close
				Set Rs = Nothing
				GBL_CHK_TempStr = GBL_CHK_TempStr & "�����������<br>" & VbCrLf
				SaveFormData = 0
				Exit Function
			End If
			Rs.Close
			Set Rs = Nothing
			If Form_ICQ = "" Then
				Form_ICQ = 0
			Else
				If isNumeric(Form_ICQ) = 0 Then Form_ICQ = 0
				Form_ICQ = cCur(Fix(Form_ICQ))
			End If
			If Form_OICQ = "" Then
				Form_OICQ = 0
			Else
				If isNumeric(Form_OICQ) = 0 Then Form_OICQ = 0
				Form_OICQ = cCur(Fix(Form_OICQ))
			End If
			
			Form_Password2 = MD5(Form_Password1)
			If Form_Password1 <> "" Then
				Form_Password2 = ",Pass='" & Replace(Form_Password2,"'","''") & "'"
			Else
				Form_Password2 = ""
			End If
			
			Dim Form_NongLiBirth
			If Len(Form_birthday)=14 Then
				temp = cCur(Left(Form_birthday,4))
				If temp > 1950 and temp < 2050 Then
					Form_NongLiBirth = ",NongLiBirth=" & Replace(GetNongLiTimeValue(ConvertToNongLi(RestoreTime(Form_birthday))),"'","''")
				Else
					'���������������Դ���
					Form_NongLiBirth = ",NongLiBirth=" & Replace(GetTimeValue(DateAdd("m",-1,RestoreTime(Form_birthday))),"'","''")
				End If
				Form_birthday = ",birthday=" & Replace(Form_birthday,"'","''")
			Else
				Form_birthday = ",birthday=0"
				Form_NongLiBirth = ",NongLiBirth=0"
			End If
			If DEF_UserEnableUserTitle = 1 and Form_UserLevel >= DEF_UserUserTitleNeedLevel Then
				Form_UserTitle = ",UserTitle='" & Replace(Form_UserTitle,"'","''") & "'"
			Else
				Form_UserTitle = ""
			End If
			
			If Old_Form_Answer = "" then
				Form_UserTitle = Form_UserTitle & ",username='" & replace(Form_UserName,"'","''") & "'"
				Form_UserTitle = Form_UserTitle & ",question='" & replace(Form_question,"'","''") & "'"
				Form_UserTitle = Form_UserTitle & ",answer='" & md5(Form_answer) & "'"
				GBL_CHK_User = Form_UserName
			end if
			Dim SQL
			SQL = "Update LeadBBS_User Set " & _
			"Mail='" & Replace(Form_Mail,"'","''") & "'" & _
			",Address='" & Replace(Form_Address,"'","''") & "'" & _
			",Sex='" & Replace(Form_Sex,"'","''") & "'" & _
			",ICQ=" & Replace(Form_ICQ,"'","''") & _
			",OICQ=" & Replace(Form_OICQ,"'","''") & _
			",Homepage='" & Replace(Form_Homepage,"'","''") & "'" & _
			",Underwrite='" & Replace(Form_Underwrite,"'","''") & "'" & _
			",PrintUnderwrite='" & Replace(Form_PrintUnderwrite,"'","''") & "'" & _
			Form_Password2 & _
			Form_birthday & _
			Form_NongLiBirth & _
			",ShowFlag=" & Replace(Form_ShowFlag,"'","''") & _
			",NotSecret=" & Replace(Form_NotSecret,"'","''") & _
			",LastWriteTime=" & GetTimeValue(DEF_Now) & _
			",UserLimit=" & Replace(Form_UserLimit,"'","''") & _
			Form_UserTitle & _
			" Where id=" & GBL_UserID
			CALL LDExeCute(SQL,1)
		
			Response.Cookies(DEF_MasterCookies)("user") = CodeCookie(Form_UserName)
			If Form_Password1 = "" Then
			Else
				Response.Cookies(DEF_MasterCookies)("pass") = CodeCookie(Form_Password1)
				Response.Cookies(DEF_MasterCookies).Domain = DEF_AbsolutHome
			End If
		
			Dim TA
			TA = Session(DEF_MasterCookies & "UDT")
			If isArray(TA) Then
				TA(0) = GBL_UserID
				TA(1) = GBL_CHK_User
				TA(2) = Form_UserLimit
				If Form_ShowFlag = "0" Then
					TA(3) = 0
				Else
					TA(3) = 1
				End If
				If Form_Password1 <> "" Then TA(9) = MD5(Form_Password1)
				Session(DEF_MasterCookies & "UDT") = TA
			End If
	end select
	SaveFormData = 1

End Function

Sub Processor_Msg(str)

	%>
		<script>
				parent.upload_resetajax("<%=str%>");
		</script>
	<%

End Sub

function User_ModifyUserFace

	User_ModifyUserFace = 1
	GBL_CHK_TempStr = ""
	Dim file,FileName
	set file = Form_UpClass.file("userface")	
	Dim upload_x1,upload_x2,upload_y1,upload_y2,upload_filename
	upload_x1 = toNum(GetFormData("upload_x1"),0)
	upload_x2 = toNum(GetFormData("upload_x2"),0)
	upload_y1 = toNum(GetFormData("upload_y1"),0)
	upload_y2 = toNum(GetFormData("upload_y2"),0)
	upload_filename = trim(GetFormData("upload_filename"))
	upload_step = trim(GetFormData("upload_step"))
	if upload_step <> "1" then upload_step = ""
	
	dim cropflag
	cropflag = 1
	if(upload_x2-upload_x1<20 or upload_y2-upload_y1<20) then
		cropflag = 0
	End If
	if upload_filename <> "" and upload_step <> "1" then
		FileName = upload_filename
	else
		FileName = Right(Trim(file.fileName),50)
	end if
	If FileName = "" Then
		Set file = Nothing
		Processor_Msg("ok2")
		User_ModifyUserFace = 2
		Exit function
	End If
	
	Dim FileType,Tmp,FileSize
	
	if upload_step = "1" then
		FileType = LCase(file.FileType)
		FileSize = file.FileSize
		Tmp = InStr(FileType,"/")
		If Tmp > 0 Then
			Tmp = Left(FileType,Tmp-1)
		Else
			Tmp = FileType
		End If
		GBL_CHK_TempStr = ""
		If Tmp = "image" and (inStr(FileType,"pjpeg") or inStr(FileType,"jpeg") or inStr(FileType,"gif")) Then
		Else
			GBL_CHK_TempStr = "ͷ���ļ���ʽ����,����ͷ��ʧ��."
		End If
	Else
		FileSize = DEF_FaceMaxBytes - 1
	End If
	
	If FileSize > DEF_FaceMaxBytes Then
		GBL_CHK_TempStr = "ͷ���ļ���С������ָ����С,����ͷ��ʧ��."
	End If
	
	if ajaxflag = "1" and upload_step = "1" then
		If GBL_CHK_TempStr <> "" Then			
			Processor_Msg("error")
			exit function
		end if
		dim tmpFile : tmpFile = FileName
		If inStrRev(tmpFile,".")>0 Then tmpFile = Mid(tmpFile,inStrRev(tmpFile,".")+1)
		If tmpFile = "gif" or tmpFile = "jpg" or tmpFile = "jpeg" or tmpFile = "jpe" or tmpFile = "bmp" or tmpFile = "png" Then
			tmpFile = DEF_BBS_HomeUrl & "temp/uface_" & GBL_UserID & "." & tmpFile
			file.saveas Server.MapPath(tmpFile)
			Set file = Nothing
			If DEF_EnableGFL = 1 then Processor_Msg(tmpFile)
		Else
			Processor_Msg("error")
		end if
		if DEF_EnableGFL = 1 then exit function
	end if
	If GBL_CHK_TempStr <> "" Then
		Processor_Msg("error")
		exit function
	end if
	
	UploadPhotoUrl = UploadPhotoUrl & "face/"
	PhotoDirectory = PhotoDirectory & "face/"

	Dim NewFileName
	NewFileName = FileName
	If Instr(NewFileName,"\") or inStr(NewFileName,"/") Then
		NewFileName = LCase(Mid(NewFileName,InstrRev(Replace(NewFileName & "","/","\"), "\") + 1))
	End If
	If inStr(NewFileName,".") = 0 Then
		GBL_CHK_TempStr = "û����ȷ��ѡ��Ҫ�ϴ����ļ�, ע���ϴ����ļ���ʽ."
		Processor_Msg("error")
		Exit function
	End If

	Pic_Name1 = GetSaveFileName(NewFileName)
	
	UploadPhotoUrl2 = UploadPhotoUrl & Pic_Name2
	UploadPhotoUrl = UploadPhotoUrl & Pic_Name1
	Pic_Name = Pic_Name1
	
	
	Dim Temp,Old_pic_name
	Old_pic_name = pic_name
	If inStrRev(pic_name,".")>0 Then pic_name = Mid(pic_name,inStrRev(pic_name,".")+1)
	
	'file.saveas PhotoDir & Pic_Name
	'FileName = Replace(Replace(FileName,"\","/"),"//","/")
	
	Dim Temp_File
	Temp_File = Server.MapPath(DEF_BBS_HomeUrl & "temp/uface_" & GBL_UserID & "." & pic_name)
	
	if cropflag = 1 and DEF_EnableGFL = 1 then
		Dim MyObj
		Set MyObj = Server.CreateObject("Persits.Jpeg")
		MyObj.open Temp_File
		MyObj.Crop upload_x1,upload_y1,upload_x2,upload_y2
	
		MyObj.save(Temp_File)
		set MyObj = nothing
	end if
	Set file = Nothing
	
	If DEF_EnableGFL = 1 Then
		GBL_Width = GetPicInfo(Temp_File,"width")
		GBL_Height = GetPicInfo(Temp_File,"height")
		If GBL_Width <= DEF_AllFaceMaxWidth and GBL_Height <= DEF_AllFaceMaxWidth Then
			Temp = 2
			call MoveFiles(Temp_File,replace(PhotoDir & pic_name2,"s.","."))
		else
			Temp = SaveSmallPic(Temp_File,PhotoDir & pic_name2,DEF_AllFaceMaxWidth,DEF_AllFaceMaxWidth,-2)
		end if
		If Temp = 4 Then
			If inStrRev(UploadPhotoUrl2,".")>0 Then
				UploadPhotoUrl2 = Left(UploadPhotoUrl2,inStrRev(UploadPhotoUrl2,".")) & "jpg"
			Else
				UploadPhotoUrl2 = UploadPhotoUrl2 & "jpg"
			End if
			If inStrRev(pic_name2,".")>0 Then
				pic_name2 = Left(pic_name2,inStrRev(pic_name2,".")) & "jpg"
			Else
				pic_name2 = pic_name2 & "jpg"
			End if
		ElseIf Temp = 3 Then
			If inStrRev(UploadPhotoUrl2,".")>0 Then
				UploadPhotoUrl2 = Left(UploadPhotoUrl2,inStrRev(UploadPhotoUrl2,".")) & "gif"
			Else
				UploadPhotoUrl2 = UploadPhotoUrl2 & "gif"
			End if
			If inStrRev(pic_name2,".")>0 Then
				pic_name2 = Left(pic_name2,inStrRev(pic_name2,".")) & "gif"
			Else
				pic_name2 = pic_name2 & "gif"
			End if
		Else
			If Temp = 1 Then
				If inStrRev(UploadPhotoUrl2,".")>0 Then
					UploadPhotoUrl2 = Left(UploadPhotoUrl2,inStrRev(UploadPhotoUrl2,".")) & "jpg"
				Else
					UploadPhotoUrl2 = UploadPhotoUrl2 & "jpg"
				End if
				If inStrRev(pic_name2,".")>0 Then
					pic_name2 = Left(pic_name2,inStrRev(pic_name2,".")) & "jpg"
				Else
					pic_name2 = pic_name2 & "jpg"
				End if
			End If
		End If
		If Temp = 1 or Temp = 3 or Temp = 4 Then
			If GBL_Width < 20 or GBL_Height < 20 Then
				GBL_CHK_TempStr = "ͼ���Ȼ�߶�̫С����������,��СҪ��20����"
				DeleteFiles PhotoDir & pic_name2
			Else
				CheckUploadDatabase PhotoDir & pic_name2,""
				CALL LDExeCute("Update LeadBBS_User Set FaceUrl='" & Replace(UploadPhotoUrl2,"'","''") & "',FaceWidth=" & GBL_Width & ",FaceHeight=" & GBL_Height & " where ID=" & GBL_UserID,1)
			End If
			DeleteFiles(PhotoDir & Old_pic_name)
		Else
			If Temp = 2 Then
				If GBL_Width < 20 or GBL_Height < 20 Then
					GBL_CHK_TempStr = "ͼ���Ȼ�߶�̫С����������,��СҪ��20����"
					DeleteFiles PhotoDir & pic_name2
					DeleteFiles(PhotoDir & Old_pic_name)
					DeleteFiles Temp_File
				Else
					CheckUploadDatabase PhotoDir & Old_pic_name,""
					CALL LDExeCute("Update LeadBBS_User Set FaceUrl='" & Replace(UploadPhotoUrl,"'","''") & "',FaceWidth=" & GBL_Width & ",FaceHeight=" & GBL_Height & " where ID=" & GBL_UserID,1)
				End If
			Else
				GBL_CHK_TempStr = "��ͼ���ļ�,�ϴ�����!"
				DeleteFiles(PhotoDir & Old_pic_name)
			End If
		End If
	Else
		If pic_name = "gif" or pic_name = "jpg" or pic_name = "jpeg" or pic_name = "jpe" or pic_name = "bmp" or pic_name = "png" Then
			call MoveFiles(Temp_File,replace(PhotoDir & pic_name2,"s.","."))
			'CheckUploadDatabase PhotoDir & Old_pic_name,""
			If DEF_EnableGFL = 0 then Processor_Msg(UploadPhotoUrl)
			CheckUploadDatabase replace(PhotoDir & pic_name2,"s.","."),""
			response.write "Update LeadBBS_User Set FaceUrl='" & Replace(UploadPhotoUrl,"'","''") & "',FaceWidth=" & GBL_Width & ",FaceHeight=" & GBL_Height & " where ID=" & GBL_UserID
			CALL LDExeCute("Update LeadBBS_User Set FaceUrl='" & Replace(UploadPhotoUrl,"'","''") & "',FaceWidth=" & GBL_Width & ",FaceHeight=" & GBL_Height & " where ID=" & GBL_UserID,1)
		Else
			GBL_CHK_TempStr = "��ͼ���ļ�,�ϴ�����"
			DeleteFiles(PhotoDir & Old_pic_name)
		End If
	End If
	if GBL_CHK_TempStr <> "" Then
		User_ModifyUserFace = 2
		If DEF_EnableGFL = 1 then Processor_Msg("error")
	else
		If DEF_EnableGFL = 1 then Processor_Msg("ok")
	end if
	'DeleteFiles Temp_File

End function


Function GetSaveFileName(name)

	Dim ExtendFileName,TempNum,Temp
	name = Lcase(name)
	name = "1" & Mid(name,inStrRev(name,"."))
	ExtendFileName = Trim(Mid(name,inStrRev(name,".")))
	TempNum = Right("0" & day(DEF_Now),2) & "_" & Right(GetTimeValue(DEF_Now),6)

	If inStr(DEF_UploadFileType,":" & ExtendFileName & ":") < 1 Then ExtendFileName = ".LeadBBS"
	If inStr(":.htw:.ida:.asp:.asa:.idq:.cer:.cdx:.htr:.idc:.shtm:.shtml:.stm:.printer:.asax:.ascx:.ashx:.asmx:.aspx:.axd:.vsdisco:.rem:.soap:.config:.cs:.csproj:.vb:.vbproj:.webinfo:.licx:.resx:.resources:.php:.cgi:",":" & ExtendFileName & ":") Then ExtendFileName = ".LeadBBS"

	GetSaveFileName = TempNum & ExtendFileName
	Pic_Name2 = TempNum & "s" & ExtendFileName
	
	'On Error Resume Next
	Dim FSFlag
	FSFlag = 1
	If DEF_FSOString <> "" Then
		Set fs = Server.CreateObject(DEF_FSOString)
		If Err Then
			FSFlag = 0
			Err.Clear
			Set fs = Nothing
		End If
	Else
		FSFlag = 0
	End If

	If FSFlag = 0 Then
		Err.Clear
		GetSaveFileName = Left(GetTimeValue(DEF_Now),8) & GetSaveFileName
		Pic_Name2 = Left(GetTimeValue(DEF_Now),8) & Pic_Name2
		PhotoDir = Server.MapPath(PhotoDirectory) & "\"
		Set Fs = Nothing
		Dim Rs
		Set Rs = LDExeCute(sql_select("Select ID,PhotoDir from LeadBBS_UserFace Where UserID=" & GBL_UserID,1),0)
		If Not Rs.Eof Then
			Pic_Name2 = Replace(Rs("PhotoDir") & "","\","/")
			Rs.Close
			Set Rs = Nothing
			If Pic_Name2 <> "" Then
				If inStrRev(Pic_Name2,"/") Then Pic_Name2 = Mid(Pic_Name2,inStrRev(Pic_Name2,"/") + 1)
				If inStrRev(Pic_Name2,".") Then Pic_Name2 = Left(Pic_Name2,inStrRev(Pic_Name2,".") - 1)
				Pic_Name2 = Pic_Name2 & ExtendFileName
				GetSaveFileName = Pic_Name2
			End If
		Else
			Rs.Close
			Set Rs = Nothing
		End If
		Exit Function
	End If

	Dim TDir,FS
	TDir = Server.MapPath(PhotoDirectory) & "\"
	If Not FS.FolderExists(TDir) then
		GetSaveFileName = 0
		GBL_Chk_TempStr = "���󣬴��ͼ���Ŀ¼�����ڣ�����ϵ��վ���ģ�"
	End If
	
	TDir = TDir & year(DEF_Now) & "\"
	UploadPhotoUrl = UploadPhotoUrl & year(DEF_Now) & "/"
	If Not FS.FolderExists(TDir) then
		FS.CreateFolder(TDir)
	End If

	TDir = TDir & Right("0" & month(DEF_Now),2) & "\"
	UploadPhotoUrl = UploadPhotoUrl & Right("0" & month(DEF_Now),2) & "/"
	If Not FS.FolderExists(TDir) then
		FS.CreateFolder(TDir)
	End If
	
	'TDir = TDir & Right("0" & day(DEF_Now),2) & "\"
	'UploadPhotoUrl = UploadPhotoUrl & Right("0" & day(DEF_Now),2) & "/"
	'If Not FS.FolderExists(TDir) then
	'	FS.CreateFolder(TDir)
	'End If
	
	PhotoDir = TDir

	If FS.FileExists(TDir & GetSaveFileName) then
		For Temp = 0 To 99
			GetSaveFileName = TempNum & "_" & Temp & ExtendFileName
			Pic_Name2 = TempNum & "_" & Temp & "s" & ExtendFileName
			If FS.FileExists(TDir & GetSaveFileName) then
			Else
				Set FS = Nothing
				Exit For
			End If
		Next
		Set FS = Nothing
	Else
		Set FS = Nothing
	End If

End Function

Dim FileUp,Pic_Name1,Pic_Name,PhotoDir,Pic_Name2,UploadPhotoUrl2,NewFileName

Function DeleteFiles(path)

	If DEF_FSOString = "" Then Exit Function
	on error resume next
	Dim fs
	Set fs = Server.CreateObject(DEF_FSOString)
	If err <> 0 Then
		Err.Clear
		'Response.Write "<br>��������֧��FSO��Ӳ���ļ�δɾ����"
		Exit Function
	End If
	If fs.FileExists(path) Then
		fs.DeleteFile path,True
		DeleteFiles = 1
	Else
		DeleteFiles = 0
	End If
	Set fs = Nothing
         
End Function%>