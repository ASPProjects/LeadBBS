<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/User_Setup.ASP -->
<!-- #include file=../../inc/Board_popfun.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<!-- #include file=../../inc/ubbcode.asp -->
<!-- #include file=inc/User_fun.ASP -->
<!-- #include file=../../inc/Limit_Fun.asp -->
<!-- #include file=../../inc/Constellation2.asp -->

<%
DEF_BBS_HomeUrl = "../../"
Dim GBL_ID
CursorLocation = 3
initDatabase

GBL_CHK_TempStr = ""
GBL_ID = checkSupervisorPass

Manage_sitehead DEF_SiteNameString & " - ����Ա",""
GBL_CHK_TempStr=""
Form_ID = Left(Request("Form_ID"),14)
If isNumeric(Form_ID) = 0 Then Form_ID = 0
Form_ID = cCur(Form_ID)
If Form_ID < 0 Then Form_ID = 0

If Form_ID=0 Then
	GBL_CHK_TempStr = GBL_CHK_TempStr & "û��ѡ��Ҫ�޸ĵ��û�<br>" & VbCrLf
End If
frame_TopInfo
DisplayUserNavigate("�û������޸�")
If GBL_CHK_Flag=1 Then
	If GBL_CHK_TempStr = "" Then
		If Request.Form("SubmitFlag")="29d98Sasphouseasp8asphnet" Then
			GBL_CHK_TempStr = ""
			checkFormDate

			If GBL_CHK_Flag = 0 Then
				Response.WRite "<div class=alert>" & GBL_CHK_TempStr & "</div>" & VbCrLf
				JoinForm
			Else
				If saveFormData = 1 Then
					displayAccessFull
				Else
					Response.WRite "<div class=alert>" & GBL_CHK_TempStr & "</div>" & VbCrLf
					JoinForm
				End If
			End If
		Else
			GetUserData(Form_ID)
			JoinForm
		End If
	Else%>
		<div class=frameline>
			<%=GBL_CHK_TempStr%>
		</div>
	<%End If
Else
DisplayLoginForm
End If
frame_BottomInfo
closeDataBase
Manage_Sitebottom("none")

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
	Form_Mail = Rs("Mail")
	Form_Address = Rs("Address")
	Form_Sex = Rs("Sex")
	Form_ICQ = Rs("ICQ")
	Form_OICQ = Rs("OICQ")
	Form_Userphoto = Rs("Userphoto")
	Form_Homepage = Rs("Homepage")
	Form_Underwrite = Rs("Underwrite")
	'Form_Password1 = Rs("Pass")
	'Form_Password2 = Form_Password1
	Form_birthday = Rs("birthday")
	If len(Form_birthday)=14 Then
		Form_birthday = RestoreTime(Rs("birthday"))
		Form_bday = day(Form_birthday)
		Form_byear = year(Form_birthday)
		Form_bmonth = month(Form_birthday)
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
	Form_ID = Rs("ID")
	Form_Login_ip = Rs("Login_ip")
	Form_Login_oknum = Rs("Login_oknum")
	Form_Login_falsenum = Rs("Login_falsenum")
	Form_Login_lastpass = Rs("Login_lastpass")
	Form_Login_RightIP = Rs("Login_RightIP")
	Form_Question = Rs("Question")
	'Form_Answer = Rs("Answer")
	Form_LockIP = Rs("LockIP")
	
	If DEF_AllDefineFace <> 0 Then
		Form_FaceUrl = Rs("FaceUrl")
		Form_FaceWidth = Rs("FaceWidth")
		Form_FaceHeight = Rs("FaceHeight")
	End If
	Form_UserLimit = Rs("UserLimit")
	Form_UserTitle = Rs("UserTitle")
	Form_CachetValue = Rs("CachetValue")
	Form_CharmPoint = Rs("CharmPoint")
	Rs.Close
	Set Rs = Nothing
	GetUserData = 1
	GBL_CHK_Flag = 1

End Function

Function JoinForm%>
<head>
	<style type=text/css>
		.input
		{
			FONT-FAMILY: ����;
			border-left:0px;
			border-right:0px;
			border-top:0px;
			border-bottom:1px groove #0055ff;
			width:240px;
			font-size:9pt
		}
		.inputs
		{
			FONT-FAMILY: ����;
			border-left:0px;
			border-right:0px;
			border-top:0px;
			border-bottom:1px groove #0055ff;
			width:40px;
			font-size:9pt
		}
		.inputss
		{
			FONT-FAMILY: ����;
			border-left:0px;
			border-right:0px;
			border-top:0px;
			border-bottom:1px groove #0055ff;
			width:20px;
			font-size:9pt
		}
	</style>
	<script LANGUAGE="JavaScript" TYPE="text/javascript">
		function setface() 
		{
			window.open('facelist.asp','','width=250,height=450 scrollbars=auto,status=no');
		}
	</script>
	<script language=JavaScript>
	<!--
	ValidationPassed = true;
	function isnum(str)
	{
		rset="";
		for(i=0;i<str.length;i++)
		{
			if(str.charAt(i)>="0" && str.charAt(i)<="9")
			{
			}
			else
			{
				return 0;
			}
		}
		return 1;
	}

	function changeface()
	{
		var temp;
		temp=document.form1.Form_userphoto.value;
		if (temp!="" && isnum(temp)==1 && temp.length==4)
		{
			if (temp > 0 && temp <= <%=DEF_faceMaxNum%>)
			{
				document.faceimg.src='<%=DEF_BBS_HomeUrl%>images/face/'+temp+'.gif';
			}
			else
			{
				alert("����!��ͼ����Ų�����!");
				document.faceimg.src='<%=DEF_BBS_HomeUrl%>images/null.gif';
				document.form1.Form_userphoto.value='';
				ValidationPassed = false;
			}
		}
		else
		{
			alert("����!��ͼ����Ų�����!\nͼ����ű�����4λ��<%if len(Cstr(DEF_faceMaxNum))>4 then Response.Write "������"%>,���� 0001 ,���Ϊ<%=DEF_faceMaxNum%>");
			document.faceimg.src='<%=DEF_BBS_HomeUrl%>images/null.gif';
			document.form1.Form_userphoto.value='';
			ValidationPassed = false;
		}
	}

	<%If DEF_AllDefineFace <> 0 Then%>
	function changeface2()
	{
		var temp,obj;
		obj=document.form1;
		if(obj.Form_FaceWidth.value!="")
		{
			if (! isnum(obj.Form_FaceWidth.value))
			{
				alert("�Զ���ͷ���ȱ��������֣�\n");
				obj.Form_FaceWidth.focus();
				return;
			}
			else
			{
				if(obj.Form_FaceWidth.value<20 || obj.Form_FaceWidth.value><%=DEF_AllFaceMaxWidth%>)
				{
					alert("�Զ���ͷ���ȱ�����20-<%=DEF_AllFaceMaxWidth%>֮�䣡\n");
					obj.Form_FaceWidth.focus();
					return;
				}
			}
		}

		if(obj.Form_FaceHeight.value!="")
		{
			if (! isnum(obj.Form_FaceHeight.value))
			{
				alert("�Զ���ͷ��߶ȱ��������֣�\n");
				obj.Form_FaceHeight.focus();
				return;
			}
			else
			{
				if(obj.Form_FaceHeight.value<20 || obj.Form_FaceHeight.value><%=DEF_AllFaceMaxWidth*2%>)
				{
					alert("�Զ���ͷ��߶ȱ�����20-<%=DEF_AllFaceMaxWidth*2%>֮�䣡\n");
					obj.Form_FaceHeight.focus();
					return;
				}
			}
		}

		temp=document.form1.Form_FaceUrl.value;
		if (temp!="")
		{
			document.faceimg.src=temp;
			document.faceimg.width=obj.Form_FaceWidth.value;
			document.faceimg.height=obj.Form_FaceHeight.value;
		}
	}
	<%End If%>
	function form_onsubmit(obj)
	{
		if(obj.Form_username.value=="")
		{
			alert("����������û���!\n");
			ValidationPassed = false;
			obj.Form_username.focus();
			return;
		}
		
		if(obj.Form_username.value.length<1)
		{
			alert("�û�������������Ҫ1���ַ�!\n");
			ValidationPassed = false;
			obj.Form_username.focus();
			return;
		}

		//if(obj.Form_password1.value=="")
		//{
		//	alert("�������µ�����!\n");
		//	ValidationPassed = false;
		//	obj.Form_password1.focus();
		//	return;
		//}

		//if(obj.Form_password2.value=="")
		//{
		//	alert("�����������֤���룡\n");
		//	ValidationPassed = false;
		//	obj.Form_password2.focus();
		//	return;
		//}

		if(obj.Form_password1.value!=obj.Form_password2.value)
		{
			alert("��������������벻��ͬ��\n");
			ValidationPassed = false;
			obj.Form_password1.focus();
			return;
		}

		if(obj.Form_icq.value!="")
		{
			if (! isnum(obj.Form_icq.value))
			{
				alert("ι,��������ICQ���������˶���,�����ICQ������ô�������֣�\n");
				ValidationPassed = false;
				obj.Form_icq.focus();
				return;
			}
		}

		if(obj.Form_oicq.value!="")
		{
			if (! isnum(obj.Form_oicq.value))
			{
				alert("ι,��������OICQ���������˶���,�����OICQ������ô��������?\n");
				ValidationPassed = false;
				obj.Form_oicq.focus();
				return;
			}
		}

		if(obj.Form_byear.value!="")
		{
			if (! isnum(obj.Form_byear.value))
			{
				alert("ι,����������ĳ�����,����������ô�������֣�\n");
				ValidationPassed = false;
				obj.Form_byear.focus();
				return;
			}
		}

		if(obj.Form_bmonth.value!="")
		{
			if (! isnum(obj.Form_bmonth.value))
			{
				alert("ι,����������ĳ�����,������·���ô�������֣�\n");
				ValidationPassed = false;
				obj.Form_bmonth.focus();
				return;
			}
		}

		if(obj.Form_bday.value!="")
		{
			if (! isnum(obj.Form_bday.value))
			{
				alert("ι,����������ĳ�����,����ĳ�������ô�������֣�\n");
				ValidationPassed = false;
				obj.Form_bday.focus();
				return;
			}
		}

		if(obj.Form_userphoto.value!="")
		{
			if (! isnum(obj.Form_bday.value))
			{
				alert("�û�ͼ��,ֻ����001-318֮������֣�\n");
				ValidationPassed = false;
				obj.Form_bday.focus();
				return;
			}
		}
		
		if(obj.Form_Underwrite.value.length>255)
		{
			alert("�û�ǩ������ҪС��255���ַ�!\n");
			ValidationPassed = false;
			obj.Form_Underwrite.focus();
			return;
		}
		//��������
		

		if(obj.Form_ApplyTime.value=="")
		{
			alert("����ʱ��ɲ�������ѽ!\n");
			ValidationPassed = false;
			obj.Form_ApplyTime.focus();
			return;
		}		

		if(obj.Form_ApplyTime.value!="")
		{
			if (! isnum(obj.Form_ApplyTime.value))
			{
				alert("ι,������������ʱ��,����ʱ����Ҫ�����ֵ��ޣ�\n");
				ValidationPassed = false;
				obj.Form_ApplyTime.focus();
				return;
			}
			if (obj.Form_ApplyTime.value.length!=14)
			{
				alert("����ʱ�������14λ���ޣ�\n");
				ValidationPassed = false;
				obj.Form_ApplyTime.focus();
				return;
			}
		}
		
		if(obj.Form_Online.value=="")
		{
			alert("����״̬�ɲ�������ѽ!\n");
			ValidationPassed = false;
			obj.Form_Online.focus();
			return;
		}
		if (! isnum(obj.Form_Online.value))
		{
			alert("����״̬��������������\n");
			ValidationPassed = false;
			obj.Form_Online.focus();
			return;
		}
		if(obj.Form_Prevtime.value=="")
		{
			alert("����¼ʱ��ɲ�������ѽ!\n");
			ValidationPassed = false;
			obj.Form_Prevtime.focus();
			return;
		}		

		if(obj.Form_Prevtime.value!="")
		{
			if (! isnum(obj.Form_Prevtime.value))
			{
				alert("ι,������������¼ʱ��,����¼ʱ����Ҫ�����ֵ��ޣ�\n");
				ValidationPassed = false;
				obj.Form_Prevtime.focus();
				return;
			}
			if (obj.Form_Prevtime.value.length!=14)
			{
				alert("����¼ʱ�������14λ���ޣ�\n");
				ValidationPassed = false;
				obj.Form_Prevtime.focus();
				return;
			}
		}
		
		if(obj.Form_UserLevel.value=="")
		{
			alert("�û�<%=DEF_PointsName(3)%>�ɲ�������ѽ!\n");
			ValidationPassed = false;
			obj.Form_UserLevel.focus();
			return;
		}
		if (! isnum(obj.Form_UserLevel.value))
		{
			alert("�û�<%=DEF_PointsName(3)%>��������������\n");
			ValidationPassed = false;
			obj.Form_UserLevel.focus();
			return;
		}
		if (obj.Form_UserLevel.value><%=DEF_UserLevelNum%>||obj.Form_UserLevel.value<0)
		{
			alert("�û�<%=DEF_PointsName(3)%>ֵ�����Ǵ����0����С��<%=DEF_UserLevelNum%>��\n");
			ValidationPassed = false;
			obj.Form_UserLevel.focus();
			return;
		}

		if(obj.Form_Points.value=="")
		{
			alert("�û�<%=DEF_PointsName(0)%>�ɲ�������ѽ!\n");
			ValidationPassed = false;
			obj.Form_Points.focus();
			return;
		}
		if (! isnum(obj.Form_Points.value))
		{
			alert("�û�<%=DEF_PointsName(0)%>��������������\n");
			ValidationPassed = false;
			obj.Form_Points.focus();
			return;
		}
		
		if(obj.Form_Officer.value=="")
		{
			alert("<%=DEF_PointsName(9)%>�ɲ�������ѽ!\n");
			ValidationPassed = false;
			obj.Form_Officer.focus();
			return;
		}

		if (! isnum(obj.Form_Login_oknum.value))
		{
			alert("�ɹ��Ǵα�������������\n");
			ValidationPassed = false;
			obj.Form_Login_oknum.focus();
			return;
		}
		if (! isnum(obj.Form_Login_falsenum.value))
		{
			alert("�ְܵǴα�������������\n");
			ValidationPassed = false;
			obj.Form_Login_falsenum.focus();
			return;
		}
		
		<%If DEF_AllDefineFace <> 0 Then%>
		if(obj.Form_FaceWidth.value!="")
		{
			if (! isnum(obj.Form_FaceWidth.value))
			{
				alert("�Զ���ͷ���ȱ��������֣�\n");
				ValidationPassed = false;
				obj.Form_FaceWidth.focus();
				return;
			}
			else
			{
				if(obj.Form_FaceWidth.value<20 || obj.Form_FaceWidth.value><%=DEF_AllFaceMaxWidth%>)
				{
					alert("�Զ���ͷ���ȱ�����20-<%=DEF_AllFaceMaxWidth%>֮�䣡\n");
					ValidationPassed = false;
					obj.Form_FaceWidth.focus();
					return;
				}
			}
		}

		if(obj.Form_FaceHeight.value!="")
		{
			if (! isnum(obj.Form_FaceHeight.value))
			{
				alert("�Զ���ͷ��߶ȱ��������֣�\n");
				ValidationPassed = false;
				obj.Form_FaceHeight.focus();
				return;
			}
			else
			{
				if(obj.Form_FaceHeight.value<20 || obj.Form_FaceHeight.value><%=DEF_AllFaceMaxWidth*2%>)
				{
					alert("�Զ���ͷ��߶ȱ�����20-<%=DEF_AllFaceMaxWidth%>֮�䣡\n");
					ValidationPassed = false;
					obj.Form_FaceHeight.focus();
					return;
				}
			}
		}<%End if%>
		ValidationPassed = true;
		return true;
	}
	-->
	</script>
</head>

<form action=UserModify.asp method=post name=form1 onSubmit="return ValidationPassed">
	<div class=frametitle>�û������޸�</div>

	<table border=0 cellpadding=0 cellspacing=0 width="100%" class=frame_table>
			<tr>
				<td class=tdbox width=120>
					<p>*�û����ƣ� 
				</td>
				<td class=tdbox>
					<p>
					<input maxLength=20 name="Form_username" size=36 class=fminpt Value="<% If Form_username<>"" Then Response.Write Server.HtmlEncode(Form_Username)%>">
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>&nbsp;�µ����룺 
				</td>
				<td class=tdbox>
					<input name=SubmitFlag type=hidden value="29d98Sasphouseasp8asphnet">
					<input name=Form_ID type=hidden value="<%=htmlencode(Form_ID)%>">
					<input maxLength=20 name="Form_password1" size=36 class=fminpt type=password Value="<% If Form_password1<>"" Then Response.Write Server.HtmlEncode(Form_password1)%>">
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>&nbsp;��֤���룺 
				</td>
				<td class=tdbox>
					<input maxlength=20 name="Form_password2" size=36 class=fminpt type=password Value="<% If Form_password2<>"" Then Response.Write Server.HtmlEncode(Form_password2)%>">
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>*������ʾ�� 
				</td>
				<td class=tdbox>
					<input maxLength=20 name=Form_Question class=fminpt size=36 Value="<% If Form_Question<>"" Then Response.Write Server.HtmlEncode(Form_Question)%>">
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>&nbsp;��ʾ�𰸣�
				</td>
				<td class=tdbox>
					<input maxlength=20 name=Form_Answer class=fminpt size=36 Value="<% If Form_Answer<>"" Then Response.Write Server.HtmlEncode(Form_Answer)%>"> ���Ŀ��Բ���д
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>�ɣ�������
				</td>
				<td class=tdbox>
					<input maxlength=15 name=Form_LockIP class=fminpt size=36 Value="<% If Form_LockIP<>"" Then Response.Write Server.HtmlEncode(Form_LockIP)%>"> ����������һ��IP��ַ
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>�����ʼ��� 
				</td>
				<td class=tdbox>
					<input maxLength=60 name=Form_mail size=36 class=fminpt Value="<% If Form_mail<>"" Then Response.Write Server.HtmlEncode(Form_mail)%>">
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>��ҳ��ַ��
				</td>
				<td class=tdbox>
					<input maxlength=250 name=Form_homepage size=36 class=fminpt Value="<% If Form_homepage<>"" Then Response.Write Server.HtmlEncode(Form_homepage)%>">
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>ICQ ���룺
				</td>
				<td class=tdbox>
					<input maxlength=10 name=Form_icq size=36 class=fminpt Value="<% If Form_icq<>"" Then Response.Write Server.HtmlEncode(Form_icq)%>">
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>OICQ���룺
				</td>
				<td class=tdbox>
					<input maxlength=10 name=Form_oicq size=36 class=fminpt Value="<% If Form_oicq<>"" Then Response.Write Server.HtmlEncode(Form_oicq)%>">
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>��ĵ�ַ��
				</td>
				<td class=tdbox>
					<input maxlength=150 name=Form_address size=36 class=fminpt Value="<% If Form_address<>"" Then Response.Write Server.HtmlEncode(Form_address)%>">
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>�û�ͷ�Σ�
				</td>
				<td class=tdbox>
					<input maxlength=14 name=Form_UserTitle size=36 class=fminpt Value="<% If Form_UserTitle<>"" Then Response.Write Server.HtmlEncode(Form_UserTitle)%>">
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>����Ա�
				</td>
				<td class=tdbox>
					<table border=0 cellpadding=0 cellspacing=0>
						<tr>
							<td><input class=fmchkbox type=radio name=Form_sex value=�� <%If Form_sex = "��" Then Response.Write " checked"%>></td><td>��</td>
							<td><input class=fmchkbox type=radio name=Form_sex value=Ů <%If Form_sex = "Ů" Then Response.Write " checked"%>></td><td>Ů</td>
							<td><input class=fmchkbox type=radio name=Form_sex value=�� <%If Form_sex = "��" Then Response.Write " checked"%>></td><td>����</td>
		 				</tr>
		  			</table>
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>�û�ͷ��
				</td>
				<td class=tdbox>
					<input onchange="javascript:changeface();" maxlength=4 name=Form_userphoto size=6 class=fminpt Value="<% If Form_userphoto<>"" Then Response.Write Server.HtmlEncode(string(4-len(cstr(Form_userphoto)),"0")&Form_userphoto)%>">
					<span style='cursor:hand' title='�鿴ͷ�����' onclick="setface();">�鿴ͷ�����</span>
					<%If DEF_AllDefineFace = 0 or Form_FaceUrl & "" = "" Then%>
						<%If Form_userphoto<>"" and isNumeric(Form_userphoto) Then%><img name=faceimg id=faceimg src=<%=DEF_BBS_HomeUrl%>images/face/<%=string(4-len(cstr(Form_userphoto)),"0")&Form_userphoto%>.gif align=middle width=62 height=62><%Else%><img name=faceimg id=faceimg src=<%=DEF_BBS_HomeUrl%>images/null.gif align=middle><%End If%>
					<%Else%>
						<img name=faceimg id=faceimg src="<%
						If Left(Lcase(Form_FaceUrl),5) <> "http:" Then Response.Write "../"
						Response.Write htmlencode(Form_FaceUrl)%>" align=middle width=<%=Form_FaceWidth%> height=<%=Form_FaceHeight%>>
					<%End If%>
				</td>
			</tr><%If DEF_AllDefineFace <> 0 Then%>
			<tr>
				<td class=tdbox>
					<p>�Զ�ͷ��
				</td>
				<td class=tdbox>
					<input onchange="javascript:changeface2();" maxlength=250 name=Form_FaceUrl size=26 class=fminpt Value="<%=HtmlEncode(Form_FaceUrl)%>">
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>ͷ���С��
				</td>
				<td class=tdbox>
					�Զ�ͷ���: <input onchange="javascript:changeface2();" maxlength=<%=len(DEF_AllFaceMaxWidth)%> name=Form_FaceWidth size=3 class=fminpt Value="<%=HtmlEncode(Form_FaceWidth)%>">(20-<%=DEF_AllFaceMaxWidth%>)
					��: <input onchange="javascript:changeface2();" maxlength=<%=len(DEF_AllFaceMaxWidth*2)%> name=Form_FaceHeight size=3 class=fminpt Value="<%=HtmlEncode(Form_FaceHeight)%>">(20-<%=DEF_AllFaceMaxWidth%>)
				</td>
			</tr><%End If%>
			<tr>
				<td class=tdbox>
					<p>������գ� 
				</td>
				<td class=tdbox align="left">
					<p>
					<input maxlength=4 name=Form_byear size=4 class=fminpt Value="<% If Form_byear<>"" Then
						Response.Write Server.HtmlEncode(Form_byear)
					Else
						Response.Write "19"
					End If%>"> �� 
					<input maxlength=2 name=Form_bmonth size=2 class=fminpt Value="<% If Form_bmonth<>"" Then Response.Write Server.HtmlEncode(Form_bmonth)%>">
					�� <input maxlength=2 name=Form_bday size=2 class=fminpt Value="<% If Form_bday<>"" Then Response.Write Server.HtmlEncode(Form_bday)%>">
					��</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>ǩ��-UBB��
				</td>
				<td class=tdbox>
					<textarea name=Form_Underwrite rows=5 cols=36 class=fmtxtra><%If Form_Underwrite <> "" Then Response.Write VbCrLf & htmlEncode(Form_Underwrite)%></textarea>
				</td>
			</tr>
			<tr>
				<td class=tdbox colspan=2 bgcolor=F7F7F7 height=25 class=TBfour>
					��չ��Ϣ</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>����ʱ�䣺
				</td>
				<td class=tdbox>
					<input maxlength=14 name=Form_ApplyTime size=14 class=fminpt Value="<% If Form_ApplyTime<>"" Then Response.Write Server.HtmlEncode(Form_ApplyTime)%>">
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>SessionID 
				</td>
				<td class=tdbox>
					<input maxlength=14 name=Form_Sessionid size=14 class=fminpt Value="<% If Form_Sessionid<>"" Then Response.Write Server.HtmlEncode(Form_Sessionid)%>">
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>����״̬��
				</td>
				<td class=tdbox>
					<input maxlength=8 name=Form_Online size=8 class=fminpt Value="<% If Form_Online<>"" Then Response.Write Server.HtmlEncode(Form_Online)%>">
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>����¼��
				</td>
				<td class=tdbox>
					<input maxlength=14 name=Form_Prevtime size=14 class=fminpt Value="<% If Form_Prevtime<>"" Then Response.Write Server.HtmlEncode(Form_Prevtime)%>">
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>�û�<%=DEF_PointsName(3)%>��
				</td>
				<td class=tdbox>
					<input maxlength=8 name=Form_UserLevel size=8 class=fminpt Value="<% If Form_UserLevel<>"" Then Response.Write Server.HtmlEncode(Form_UserLevel)%>">
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>�û�IPַ��
				</td>
				<td class=tdbox>
					<input maxlength=50 name=Form_IP size=36 class=fminpt Value="<% If Form_UserLevel<>"" Then Response.Write HtmlEncode(Form_IP)%>">
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>�û�<%=DEF_PointsName(0)%>��
				</td>
				<td class=tdbox>
					<input maxlength=14 name=Form_Points size=14 class=fminpt Value="<% If Form_Points<>"" Then Response.Write HtmlEncode(Form_Points)%>">
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>�û�<%=DEF_PointsName(2)%>��
				</td>
				<td class=tdbox>
					<input maxlength=44 name=Form_CachetValue size=14 class=fminpt Value="<%If Form_CachetValue<>"" Then Response.Write HtmlEncode(Form_CachetValue)%>">
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>�û�<%=DEF_PointsName(1)%>��
				</td>
				<td class=tdbox>
					<input maxlength=14 name=Form_CharmPoint size=14 class=fminpt Value="<%If Form_CharmPoint<>"" Then Response.Write HtmlEncode(Form_CharmPoint)%>">
				</td>
			</tr>			
			<tr>
				<td class=tdbox>
					<p><%=DEF_PointsName(9)%>��</span>
				</td>
				<td class=tdbox>
					<input maxlength=255 name=Form_Officer size=36 class=fminpt Value="<% If Form_Officer<>"" Then Response.Write HtmlEncode(Form_Officer)%>">
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>���IPַ��
				</td>
				<td class=tdbox>
					<input maxlength=50 name=Form_Login_ip size=36 class=fminpt Value="<% If Form_Login_ip<>"" Then Response.Write HtmlEncode(Form_Login_ip)%>">
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p title=�ɹ���¼����̳�Ĵ���>�ɹ��ǴΣ�
				</td>
				<td class=tdbox>
					<input maxlength=17 name=Form_Login_oknum size=36 class=fminpt Value="<% If Form_Login_oknum<>"" Then Response.Write HtmlEncode(Form_Login_oknum)%>">
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p title=������¼ʧ�ܵĴ���>�ְܵǴΣ�
				</td>
				<td class=tdbox>
					<input maxlength=17 name=Form_Login_falsenum size=36 class=fminpt Value="<% If Form_Login_falsenum<>"" Then Response.Write HtmlEncode(Form_Login_falsenum)%>">
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p title=���һ�ε�¼ʹ�õ����룬������ȷ���>ĩ�����룺
				</td>
				<td class=tdbox>
					<input maxlength=20 type=password name=Form_Login_lastpass size=36 class=fminpt Value="<% If Form_Login_lastpass<>"" Then Response.Write HtmlEncode(Form_Login_lastpass)%>">
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p title=���һ�γɹ���¼����̳�����õģɣе�ַ>����ĩIP��
				</td>
				<td class=tdbox>
					<input maxlength=50 name=Form_Login_RightIP size=36 class=fminpt Value="<% If Form_Login_RightIP<>"" Then Response.Write HtmlEncode(Form_Login_RightIP)%>">
				</td>
			</tr>
			<tr>
				<td class=tdbox bgcolor=F7F7F7 height=25 class=TBBG1>
					�û�Ȩ�ޣ�</td>
				<td class=tdbox><%
				Form_UserLimit = cCur(Form_UserLimit)
				Dim TempN
	for TempN = 0 to LimitUserStringDataNum%>
			<input type="checkbox" class=fmchkbox name="Limit<%=TempN+1%>" value="1"<%If GetBinarybit(Form_UserLimit,TempN+1) = 1 Then
				Response.Write " checked>"
			Else
				Response.Write ">"
			End If%><%=LimitUserStringData(tempN)%><br>
			<%Next%></td>
			</tr>
	<tr>
		<td class=tdbox>&nbsp;</td>
		<td class=tdbox>
			<input name=submit type=submit value=" �� �� " onclick="form_onsubmit(this.form)" class=fmbtn>
			<input name=b1 type=reset value=" �� д " class=fmbtn>
		</td>
	</tr>
	</table>
			<div class=frametitle>ע�⣺</div>
			<div class=frameline>
			<ol class=listli>
			<li>����Ȩ�ޣ�ֻ��ӵ��<%=DEF_PointsName(6)%>Ȩ�޵Ļ�Ա�������ã�����ȡ�������ܹ̣�</li>
			����ĳ����ԱIP��ַ�����λ�Ա���Եȵȣ�</li>
			<li>��ֹɾ������ ��ֹ�������� ��ֹת������ ɾ���ϴ�����������<%=DEF_PointsName(8)%>��<%=DEF_PointsName(6)%>��Ч</li>
			<li><%=DEF_PointsName(8)%>һ����ò�Ҫ�����ģ��ɰ����޸�ʱ�Զ��жϲ���</li>
			<li>��ֹ�޸ĸ������Ϻ��������ݣ��κλ�Ա����Ч�������޸����˵����Ӽ���������</li>
			<li>�û�Ȩ���е�������HTMLȨ��,ֻ���<%=DEF_PointsName(5)%>��<%=DEF_PointsName(8)%>������Ȩ�޵��û���Ч</li>
			</ol>
			</div>
			<div class=frameline>
			<%DisplayOfficerList
			DisplayLevelList%>
			</div>
</form>
<%
End Function

Function displayAccessFull%>
	<p>�޸ĳɹ����������£�<br>
	<br>
	</p>
	<table border=0 cellpadding="0" cellspacing="0" class=frame_table>
			<tr>
				<td class=tdbox width=120>
					<p>*�û����ƣ� 
				</td>
				<td class=tdbox>
					<%=Server.HtmlEncode(Form_Username)%>
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>*�µ����룺 
				</td>
				<td class=tdbox>
					********
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>�����ʼ��� 
				</td>
				<td class=tdbox>
					<%=HtmlEncode(Form_mail)%>
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>��ҳ��ַ��
				</td>
				<td class=tdbox>
					<%=HtmlEncode(Form_homepage)%>
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>ICQ ���룺
				</td>
				<td class=tdbox>
					<%=HtmlEncode(Form_icq)%>
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>OICQ���룺
				</td>
				<td class=tdbox>
					<%=HtmlEncode(Form_oicq)%>
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>��ĵ�ַ��
				</td>
				<td class=tdbox>
					<%=HtmlEncode(Form_address)%>
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>�û�ͷ�Σ�
				</td>
				<td class=tdbox>
					<%=HtmlEncode(Form_UserTitle)%>
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>����Ա�
				</td>
				<td class=tdbox>
					<%=Form_sex%>
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>�û�ͷ��
				</td>
				<td class=tdbox>
					<%If DEF_AllDefineFace = 0 or Form_FaceUrl = "" Then%>
					<img src=<%=DEF_BBS_HomeUrl%>images/face/<%=string(4-len(cstr(Form_userphoto)),"0")&Form_userphoto%>.gif align=middle width=62 height=62>
					<%Else%>
						<img src="<%If Left(Lcase(Form_FaceUrl),5) <> "http:" Then Response.Write "../"
						Response.Write htmlencode(Form_FaceUrl)%>" align=middle width=<%=Form_FaceWidth%> height=<%=Form_FaceHeight%>>
					<%End If%>
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>������գ�
				</td>
				<td class=tdbox align="left">
					<%If len(Form_birthday)=14 Then Response.Write RestoreTime(Form_birthday)%></td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>ǩ��-UBB��
				</td>
				<td class=tdbox>
					<table style="table-layout:fixed; word-break:break-all" width=332 border="0" cellspacing="0" cellpadding="0"><tr><td class=tdbox>
					<%=Form_PrintUnderwrite%></td></tr></table></td>
				</td>
			</tr>
			<tr>
				<td class=tdbox colspan=2 bgcolor=F7F7F7 height=25 class=TBfour>
					��չ��Ϣ</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>����ʱ�䣺
				</td>
				<td class=tdbox>
					<%If len(Form_ApplyTime)=14 Then Response.Write RestoreTime(Form_ApplyTime)%>
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>SessionID 
				</td>
				<td class=tdbox>
					<%=HtmlEncode(Form_Sessionid)%>
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>����״̬��
				</td>
				<td class=tdbox>
					<%=HtmlEncode(Form_Online)%>
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>����¼��
				</td>
				<td class=tdbox>
					<%If len(Form_birthday)=14 Then Response.Write RestoreTime(Form_Prevtime)%>
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>�û�<%=DEF_PointsName(3)%>��
				</td>
				<td class=tdbox>
					<%=Form_UserLevel%>: <%=DEF_UserLevelString(Form_UserLevel)%>
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>�û�IPַ��
				</td>
				<td class=tdbox>
					<%=HtmlEncode(Form_IP)%>
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>�û�<%=DEF_PointsName(0)%>��
				</td>
				<td class=tdbox>
					<%=HtmlEncode(Form_Points)%>
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>�û�<%=DEF_PointsName(2)%>��
				</td>
				<td class=tdbox>
					<%
					Response.Write HtmlEncode(Form_CachetValue)%>
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>�û�<%=DEF_PointsName(1)%>��
				</td>
				<td class=tdbox>
					<%
					Response.Write HtmlEncode(Form_CharmPoint)%>
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p><%=DEF_PointsName(9)%>��
				</td>
				<td class=tdbox>
					<%=Form_Officer%>: <%=DisplayOfficerString(Form_Officer)%>
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>���IPַ��
				</td>
				<td class=tdbox>
					<%=Form_Login_ip%>
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>�ɹ��ǴΣ�
				</td>
				<td class=tdbox>
					<%=Form_Login_oknum%>
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>�ְܵǴΣ�
				</td>
				<td class=tdbox>
					<%=Form_Login_falsenum%>
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>ĩ�����룺
				</td>
				<td class=tdbox>
					<%=Form_Login_lastpass%>
				</td>
			</tr>
			<tr>
				<td class=tdbox>
					<p>����δIP��
				</td>
				<td class=tdbox>
					<%=Form_Login_RightIP%>
				</td>
			</tr>
			</table>
<%End Function

Function saveFormData

	Dim Rs
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	Rs.Open sql_select("Select * from LeadBBS_User where id=" & Form_ID,1),con,1,3
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		GBL_CHK_TempStr = GBL_CHK_TempStr & "�����������<br>" & VbCrLf
		saveFormData = 0
		Exit Function
	End If
	Rs("UserName") = Form_UserName
	Rs("Mail") = Form_Mail
	Rs("Address") = Form_Address
	Rs("Sex") = Form_Sex
	If Form_ICQ = "" Then
		Rs("ICQ") = Null
	Else
		Rs("ICQ") = Form_ICQ
	End If
	If Form_OICQ = "" Then
		Rs("OICQ") = Null
	Else
		Rs("OICQ") = Form_OICQ
	End If
	Rs("Userphoto") = Form_Userphoto
	Rs("Homepage") = Form_Homepage
	Rs("Underwrite") = Form_Underwrite
	Rs("PrintUnderwrite") = Form_PrintUnderwrite
	If Form_Password1 <> "" Then Rs("Pass") = MD5(Form_Password1)
	If Len(Form_birthday)=14 Then
		Rs("birthday") = Form_birthday
		Dim Temp
		temp = cCur(Left(Form_birthday,4))
		If temp > 1950 and temp < 2050 Then
			Rs("NongLiBirth") = GetNongLiTimeValue(ConvertToNongLi(RestoreTime(Form_birthday)))
		Else
			'���������������Դ���
			Rs("NongLiBirth") = GetTimeValue(DateAdd("m",-1,RestoreTime(Form_birthday)))
		End If
	Else
		Rs("birthday") = Null
	End If

	REM ��������
	Rs("ApplyTime") = Form_ApplyTime
	Rs("IP") = Form_IP
	Rs("UserLevel") = Form_UserLevel
	Rs("Officer") = Form_Officer
	Rs("Points") = Form_Points
	Rs("Sessionid") = Form_Sessionid
	Rs("Online") = Form_Online
	Rs("Prevtime") = Form_Prevtime
	If Form_Answer <> "" Then Rs("Answer") = MD5(Form_Answer)
	Rs("Question") = Form_Question
	
	Rs("Login_ip") = Form_Login_ip
	Rs("Login_oknum") = Form_Login_oknum
	Rs("Login_falsenum") = Form_Login_falsenum
	Rs("Login_lastpass") = Form_Login_lastpass
	Rs("Login_RightIP") = Form_Login_RightIP

	If DEF_AllDefineFace <> 0 Then
		Rs("FaceUrl") = Form_FaceUrl
		Rs("FaceWidth") = Form_FaceWidth
		Rs("FaceHeight") = Form_FaceHeight
	End If

	Rs("UserLimit") = Form_UserLimit
	Rs("UserTitle") = Form_UserTitle
	Rs("CachetValue") = Form_CachetValue
	Rs("CharmPoint") = Form_CharmPoint
	Rs("LockIP") = Form_LockIP

	Rs.Update
	Rs.Close
	Set Rs = Nothing
	saveFormData = 1

	UpdateSpecialUserTable Form_UserLimit,Form_ID,Form_UserName

End Function

Sub UpdateSpecialUserTable2(UserLimit,UserID,UserName,N,assort)

	Dim Rs
	Dim Flag
	
	Rem ��֤��Ա
	Flag = GetBinarybit(UserLimit,N)
	If Flag = 0 Then
		CALL LDExeCute("Delete from LeadBBS_SpecialUser where Assort=" & assort & " and UserID=" & UserID,1)
	Else
		Set Rs = LDExeCute(sql_select("Select ID from LeadBBS_SpecialUser Where Assort=" & assort & " and UserID=" & UserID,1),0)
		If Rs.Eof Then
			Rs.Close
			Set Rs = Nothing
			CALL LDExeCute("insert into LeadBBS_SpecialUser(UserID,UserName,BoardID,Assort,ndatetime) values(" & UserID & ",'" & Replace(UserName,"'","''") & "',0," & assort & "," & GetTimeValue(DEF_Now) & ")",1)
		Else
			Rs.Close
			Set Rs = Nothing
		End If
	End If

End Sub

Function UpdateSpecialUserTable(UserLimit,UserID,UserName)

	'����,0-��֤��Ա,1-����,2-�ܰ���,3-���λ�Ա,4-���Ի�Ա,5-���޸Ļ�Ա,6-δ����,7-������,8-רҵ�ͻ�
	
	Rem ��֤��Ա
	CALL UpdateSpecialUserTable2(UserLimit,UserID,UserName,1,6)

	Rem ����ʽ��Ա
	CALL UpdateSpecialUserTable2(UserLimit,UserID,UserName,2,0)
	
	Rem ����
	CALL UpdateSpecialUserTable2(UserLimit,UserID,UserName,8,1)

	Rem �ܰ���
	CALL UpdateSpecialUserTable2(UserLimit,UserID,UserName,10,2)
	
	Rem ���λ�Ա
	CALL UpdateSpecialUserTable2(UserLimit,UserID,UserName,7,3)
	
	Rem ���Ի�Ա
	CALL UpdateSpecialUserTable2(UserLimit,UserID,UserName,3,4)
	
	Rem ��ֹ�޸Ļ�Ա
	CALL UpdateSpecialUserTable2(UserLimit,UserID,UserName,4,5)
	
	Rem רҵ�ͻ�
	CALL UpdateSpecialUserTable2(UserLimit,UserID,UserName,15,8)

End Function%>