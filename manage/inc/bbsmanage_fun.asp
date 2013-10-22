<%
Response.Expires = 0 
Response.ExpiresAbsolute = DEF_Now - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private" 
Response.CacheControl = "no-cache"

Sub displayVerifycode

	Dim Url
	Url = Left(Request("dir"),100)
	If Request("dir") = "" Then
		Url = DEF_BBS_HomeUrl
	End If
%>
		<input name="ForumNumber" id="ForumNumber" maxlength="4" value="<%=htmlencode(Session(DEF_MasterCookies & "RndNum_par") & "")%>" onfocus="verify_load(0,'<%=url%>');" class="fminpt input_1" />
		<img src="<%=Url%>images/blank.gif" id="verifycode" align="middle" onclick="verify_load(1,'<%=url%>');" /> 
		<a href="javascript:;" id=verify_click onclick="this.style.display='none';verify_load(1,'<%=url%>');return false;">�����ʾ��֤��</a>
		<noscript>     
		<div class="verifycode"><img src="<%=Url%>User/number.asp?r=1" id="verifycode" class="verifycode" align="middle" onclick="verify_load(1,'<%=url%>');" /></div>
		</noscript>
<%End Sub

Function DisplayLoginForm
%>
	<table class=login_table>
	<tr>
	<td>
	<div class=title>����Ա��½</div>
	<%
	If Request("submitflag") = "" Then
	Else
		Response.Write "<p class=""alert"">" & GBL_CHK_TempStr & "</p>"
	End If
	%>
	<script language="javascript">
	function submitonce(theform)
	{
		if (document.all||document.getElementById)
		{
			for (i=0;i<theform.length;i++)
			{
				var tempobj=theform.elements[i];
				if(tempobj.type.toLowerCase()=="submit"||tempobj.type.toLowerCase()=="reset")
				tempobj.disabled=true;
			}
		}
	}
	</script>
<form action=<%=DEF_BBS_HomeUrl & DEF_ManageDir%>/Default.asp method="post" onSubmit="submitonce(this);" target="_top">
	<div class="value2">�û������� <input name=user type=text maxlength=20 size=22 value="<%=htmlencode(GBL_CHK_user)%>" class="fminpt input_2"></div>
	<div class="value2">���롡���� <input name=pass type=password maxlength=20 size=22 value="<%'=htmlencode(GBL_CHK_pass)%>" class="fminpt input_2"></div>
	<div class="value2">����𰸣� <input name=MPass type=password maxlength=20 size=22 value="<%'=htmlencode(GBL_CHK_pass)%>" class="fminpt input_2"> <span class="note">��д�һ������õĴ�</span></div>
	<div class="value2">��֤�롡�� <%displayVerifycode%><br>
	<div class="value2">��Ч�ڡ��� <select name="CkiExp" id="CkiExp">
				<option value="-99">��ȫ</option>
				<option value="-1">��Ч</option>
				<option value="365">һ��</option>
				<option value=1>һ��</option>
				<option value=2>����</option>
				<option value=7>һ��</option>
				<option value=31>һ��</option>
			</select> <span class="note">Cookie����ʱ��</span>
	</div>
	<div class="splitline"></div>
	<div class="value2">
	���������� <input name=submitflag type=hidden value="ddddls-+++"> <input type=submit value="��¼" class="fmbtn">
	<input type=reset value="ȡ��" class="fmbtn">
	</div>
</form>
	</td>
	</tr>
	</table>
<%

End Function

Function DisplayUserNavigate(str)

	Dim NewUrl
	NewUrl = DEF_BBS_HomeUrl
	If Left(NewUrl,3) = "../" or Left(NewUrl,3) = "..\" Then NewUrl = Mid(NewUrl,4)
	If DEF_SiteHomeUrl = "" Then DEF_SiteHomeUrl = DEF_BBS_HomeUrl & "Boards.asp"
	Response.Write "<div class=frame_navtitle>"
	Response.Write "<a href=""" & NewUrl & "Default.asp"" target=_top>��̳����ϵͳ</a> &gt;&gt; " & Str
	Response.Write "</div>"

End Function

Sub frame_TopInfo

%>
	<div class="frame_body">

<%End Sub

Sub frame_BottomInfo

%>
	</div>
	<div style="height:100px;"></div>
<%

End Sub

Function CheckSupervisorPass

	If Request.Form("submitflag") = "ddddls-+++" Then
		If Mid(Cstr(Request.ServerVariables("HTTP_REFERER")),8,len(Cstr(Request.ServerVariables("SERVER_NAME")))) <> Cstr(Request.ServerVariables("SERVER_NAME")) Then
			GBL_CHK_Flag = 0
			CheckSupervisorPass = 0
			GBL_CHK_TempStr = "�˲���ֻ�й���Ա���ܲ�����<br>" & VbCrLf
			Exit Function
		End If
		Dim NumCheck
		NumCheck = CheckRndNumber
		If NumCheck = 0 Then
			GBL_CHK_Flag = 0
			CheckSupervisorPass = 0
			GBL_CHK_TempStr = "��֤�����<br>" & VbCrLf
			Exit Function
		End If
	End If

	If GBL_CHK_Flag = 1 and (GBL_CHK_User <> "" and inStr(GBL_CHK_User,",") = 0 and inStr(LCase(DEF_SupervisorUserName),"," & LCase(GBL_CHK_User) & ",") > 0) and GBL_UserID > 0 Then
	Else
		GBL_CHK_Flag = 0
		checkSupervisorPass = 0
		GBL_CHK_TempStr = "�˹���ֻ�й���Ա���ܲ���[1]��<br>" & VbCrLf
		Exit Function
	End If

	If Session(DEF_MasterCookies & "Manager") <> "manage" Then
		If CheckUserAnswer(GBL_CHK_User,Left(Request.Form("MPass"),22)) = 0 Then
			GBL_CHK_Flag = 0
			checkSupervisorPass = 0
			GBL_CHK_TempStr = "�˹���ֻ�й���Ա���ܲ���[2]��<br>" & VbCrLf
			Exit Function
		End If
		Session(DEF_MasterCookies & "Manager") = "manage"
	End If

	GBL_CHK_Flag = 1
	checkSupervisorPass = 1	

End Function

Function CheckRndNumber

	If DEF_EnableAttestNumber = 0 Then
		CheckRndNumber = 1
		Exit Function
	End If

	Dim RndNumber
	RndNumber = Left(Session(DEF_MasterCookies & "RndNum") & "",4)
	If RndNumber = "" Then
		Randomize
		RndNumber = Fix(Rnd*9999)+1
		Session(DEF_MasterCookies & "RndNum") = RndNumber
	End If

	Dim ForumNumber
	ForumNumber = Left(Request.Form("ForumNumber"),4)
	If LCase(RndNumber) = LCase(ForumNumber) Then
		CheckRndNumber = 1
	Else
		CheckRndNumber = 0
	End If
	Randomize

End Function

Function CheckUserAnswer(UserName,Answer)

	Dim Rs
	Set Rs = LDExeCute(sql_select("Select Answer from LeadBBS_User where UserName='" & Replace(UserName,"'","''") & "'",1),0)
	If Rs.Eof Then
		CheckUserAnswer = 0
	Else
		If MD5(Answer) = Rs(0) Then
			CheckUserAnswer = 1
		Else
			CheckUserAnswer = 0
		End If
	End if
	Rs.Close
	Set Rs = Nothing

End Function

Sub Manage_sitehead(headString,classStr)

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
	"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml"  xml:lang="zh-CN" lang="zh-CN">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
	<meta name="description" content="<%=htmlencode(DEF_GBL_Description)%>" />
	<title>
		<%=Replace(headString,"<","&lt;")%> - Powered by <%=DEF_Version%>
	</title>
	<link rel="stylesheet" id="css" type="text/css" href="<%=DEF_BBS_homeUrl & DEF_ManageDir%>/inc/manage.css" title="managecssfile" />
	<script type="text/javascript">
	<!--
	var DEF_MasterCookies = "<%=htmlencode(DEF_MasterCookies)%>";
	var GBL_Style = "<%=GBL_Board_BoardStyle%>";
	-->
	</script>
	<script src="<%=DEF_BBS_HomeUrl%>inc/js/jquery.js" type="text/javascript"></script>
	<script src="<%=DEF_BBS_HomeUrl%>inc/js/common.js" type="text/javascript"></script>
</head>
<body id="body"<%If classStr <> "" Then%> class="<%=classStr%>" scroll="no"<%
Else%><%
End If%>>
<iframe name="hidden_frame" id="hidden_frame" style="display:none"></iframe>
<a name="top"></a>

<%
End Sub

Sub Manage_Sitebottom(flag)

	If flag = "" Then
	%>
	<div class="createtime">
	LeadBBS Copyright <span style="font:11px Tahoma,Arial,sans-serif;">&copy;</span>2003-<%=year(DEF_Now)%>.
	<%
		Response.Write " Page created in " & FormatNumber(cCur(Timer - DEF_PageExeTime1),4,True) & " seconds width " & GBL_DBNum & " queries."
	%>
	</div>
	<%End If%>
	<script type="text/javascript">
	<!--
		if (typeof submit_disable == 'function')
		{
		new LayerMenu('layer_item','layer_iteminfo');
		new LayerMenu('layer_item2','layer_iteminfo2');
		layer_initselect();
		
		var alls = document.getElementsByTagName('form'); 
		for(var i=0; i<alls.length; i++)
		{
			submit_disable(alls[i],1);
		}
		if (typeof initLightbox == 'function')initLightbox();
		}
	-->
	</script>
	</body>
	</html>
	<%

End Sub

Function Update_CheckSetupRIDExist(RID,extend)

	Dim Rs
	Set Rs = LDExeCute(sql_select("Select ID,RID,ValueStr,ClassNum,SaveData from LeadBBS_Setup where RID=" & RID & extend,1),0)
	If Rs.Eof Then
		Update_CheckSetupRIDExist = 0
	Else
		Update_CheckSetupRIDExist = 1
	End If
	Rs.Close
	Set Rs = Nothing

End Function

Sub Update_InsertSetupRID(RID,ValueStr,ClassNum,saveData,extend)

	If Update_CheckSetupRIDExist(RID,extend) = 0 Then
		CALL LDExeCute("insert into LeadBBS_Setup(RID,ValueStr,ClassNum,saveData) values(" & Rid & ",'" & Replace(ValueStr,"'","''") & "'," & ClassNum & ",'" & Replace(saveData,"'","''") & "')",1)
	Else
		CALL LDExeCute("Update LeadBBS_Setup Set ValueStr='" & Replace(ValueStr,"'","''") & "',ClassNum=" & ClassNum & ",saveData='" & Replace(saveData,"'","''") & "' where RID=" & RID & extend,1)
	End If

End Sub
%>