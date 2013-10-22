<%		
function checkInstalled

	dim filestr
	filestr = ADODB_LoadFile(DEF_BBS_HomeUrl & "inc/BBSSetup.asp")
	if inStr(filestr,"Response.Redirect ""install/default.asp""") < 1 then
		printline("<span style='color:white'><b>��̳��װ����������Ҫ���°�װ���ϴ��滻inc/BBSSetup.asp�ļ�.</b></span>")
		checkInstalled = true
	else
		checkInstalled = false
	end if

end function

sub install_head
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
	"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"  xml:lang="zh-CN" lang="zh-CN">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
	<meta name="description" content="LeadBBS ��װ" />
	<title>
		LeadBBS 7.0 ��װ
	</title>
	<link rel="stylesheet" id="css" type="text/css" href="scripts/install.css" title="cssfile" />
	<script src="../inc/js/jquery.js" type="text/javascript"></script>
	
	<script type="text/JavaScript">
	$(window).resize(function(){
	 $('.area').css({
	  position:'absolute',
	  left: (($(window).width() - $('.area').outerWidth())/2)<0?0:(($(window).width() - $('.area').outerWidth())/2),
	   top: (($(window).height() - $('.area').outerHeight())/2 + $(document).scrollTop())<0?0:(($(window).height() - $('.area').outerHeight())/2 + $(document).scrollTop())
	 });
	});
	//��ʼ������
	</script>
	
</head>

<body id="body">

<%end sub

sub install_bottom
%>

	<script type="text/JavaScript">
	$(window).resize();
	</script>
	</body>
	</html>
<%
end sub

sub install_step
%>
	<div class="install_title">LeadBBS��װ��</div>
	<div class="install_step">
	<a href="javascript:;" id="step1"<%If Step >=0 then response.write " class=""on"""%>>һ ��ʼ</a>
	<a href="javascript:;" id="step2"<%If Step >=2 then response.write " class=""on"""%>>�� ��װ���</a>
	<a href="javascript:;" id="step3"<%If Step >=3 then response.write " class=""on"""%>>�� �������ݿ�</a>
	<a href="javascript:;" id="step4"<%If Step >=4 then response.write " class=""on"""%>>�� ���ù���</a>
	<a href="javascript:;" id="step5"<%If Step >=5 then response.write " class=""on"""%>>�� ��ɰ�װ</a>
	</div>
<%
end sub

sub install_contenthead
%>


<div class="area">
<div id="wrap">
  <div id="subwrap">
   <div id="content"><div class="contents">

<%
end sub

sub install_contentbottom
%>

 	</div>
  </div>
 </div>
</div>
<%

end sub

Function toNum(s,default)

	if isNumeric(s) = 0 Then
		toNum = default
	else
		toNum = ccur(s)
	end if

End Function

Function CheckObjInstalled(strClassString,w)

	On Error Resume Next
	Dim Temp
	Err = 0
	Dim TmpObj
	Set TmpObj = Server.CreateObject(strClassString)
	Temp = Err
	If Temp = 0 Then
		CheckObjInstalled = True
		if w = 1 then Response.Write strClassString & "��<font color=green class=greenfont>��</font>"
	ElseIf Temp = -2147221005 Then
		Response.Write strClassString & "��<font color=red class=redfont>���δ��װ</font>"
		if w = 1 then CheckObjInstalled = False
	ElseIf Temp = -2147221477 Then
		if w = 1 then Response.Write strClassString & "��<font color=green class=greenfont>��֧�ִ����</font>"
		CheckObjInstalled = True
	ElseIf Temp = 1 Then
		if w = 1 then Response.Write strClassString & "��<font color=red>��δ֪�Ĵ����������δ��ȷ��װ</font>"
		CheckObjInstalled = False
	End If
	Err.Clear
	Set TmpObj = Nothing
	Err = 0

End Function

Function ADODB_LoadFile(ByVal File)

	On Error Resume Next
	Dim objStream,FSFlag,fs,WriteFile
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

	If FSFlag = 1 Then
		Set WriteFile = fs.OpenTextFile(Server.MapPath(File),1,True)
		If Err Then
			GBL_CHK_TempStr = "<br>��ȡ�ļ�ʧ�ܣ�" & err.description & "<br>�������ܣ�ȷ���Ƿ�Դ��ļ��ж�ȡȨ��."
			err.Clear
			Set Fs = Nothing
			Exit Function
		End If
		If Not WriteFile.AtEndOfStream Then
			ADODB_LoadFile = WriteFile.ReadAll
			If Err Then
				GBL_CHK_TempStr = "��ȡ�ļ�ʧ�ܣ�<p>" & err.description & "</p> �������ܣ�ȷ���Ƿ�Դ��ļ��ж�ȡȨ��."
				err.Clear
				Set Fs = Nothing
				Exit Function
			End If
		End If
		WriteFile.Close
		Set Fs = Nothing
	Else
		Set objStream = Server.CreateObject("ADODB.Stream")
		If Err.Number=-2147221005 Then 
			GBL_CHK_TempStr = "����������֧��ADODB.Stream���޷���ɲ��������ֹ�����"
			Err.Clear
			Set objStream = Nothing
			Exit Function
		End If
		With objStream
			.Type = 2
			.Mode = 3
			.Open
			.LoadFromFile Server.MapPath(File)
			.Charset = "gb2312"
			.Position = 2
			ADODB_LoadFile = .ReadText
			.Close
		End With
		Set objStream = Nothing
	End If
	If Err Then
		GBL_CHK_TempStr = "������Ϣ��<p>" & err.description & "</p>�������ܣ�ȷ���Ƿ�Դ��ļ��ж�ȡȨ��."
		err.Clear
		Set Fs = Nothing
		Exit Function
	End If

End Function

'�洢���ݵ��ļ�
Sub ADODB_SaveToFile(ByVal strBody,ByVal File)

	On Error Resume Next
	Dim objStream,FSFlag,fs,WriteFile
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
	If FSFlag = 1 Then
		Set WriteFile = fs.CreateTextFile(Server.MapPath(File),True)
		WriteFile.Write strBody
		WriteFile.Close
		Set Fs = Nothing
	Else
		Set objStream = Server.CreateObject("ADODB.Stream")
		If Err.Number=-2147221005 Then 
			GBL_CHK_TempStr = "����������֧��ADODB.Stream���޷���ɲ��������ֹ�����"
			Err.Clear
			Set objStream = Nothing
			Exit Sub
		End If
		With objStream
			.Type = 2
			.Open
			.Charset = "gb2312"
			.Position = objStream.Size
			.WriteText = strBody
			.SaveToFile Server.MapPath(File),2
			.Close
		End With
		Set objStream = Nothing
	End If
	If Err Then
		GBL_CHK_TempStr = "������Ϣ��<p>" & err.description & "</p>�������ܣ�ȷ���Ƿ�Դ��ļ���д��Ȩ��."
		err.Clear
		Set Fs = Nothing
		Exit Sub
	End If

End Sub

Function htmlEncode(str)

	If str & "" <> "" Then
		htmlEncode=Replace(Replace(Replace(str,">","&gt;"),"<","&lt;"),"""","&quot;")
	Else
		htmlEncode=str
	End If

End Function


sub install_step1form

%>
<ol>
<li>�˳��������Ŀռ䰲װLeadBBS 7.0��ȷ���ռ�ӵ��ASP����ִ��Ȩ�ޣ�</li>
<li>��װ���������ɹ�ִ��һ�Σ���Ҫ�ٴΰ�װ�������ϴ�ԭʼ�ļ���
<li>�˳����Ȩ�޹�LeadBBS�ٷ���̳���У����˿����ʹ�ô˳���<br />��ҵ��;����Ҫ�����������֧��һ���ķ���.</li>
</ol>
<a href="default.asp?step=2" class="install_submit">ͬ������Ҫ�󲢼�����װ</a>
<%

end sub

sub install_step2form

	Dim flag,f,filestr,readflag
%>
<div class="contenttitle">���¼��ȫ��ͨ�����ܼ�����װ</div>
<ol>
<li><%
f = CheckObjInstalled("Scripting.FileSystemObject",1)
if f = false then Check_com = false%></li>
<li><%
f = CheckObjInstalled("adodb.connection",1)
if f = false then Check_com = false%></li>
<li><%
f = CheckObjInstalled("Scripting.Dictionary",1)
if f = false then Check_com = false%></li>
<li><%
f = CheckObjInstalled("adodb.connection",1)
if f = false then Check_com = false%></li>
<li>��ȡȨ�޼�⣺<%

filestr = ADODB_LoadFile(DEF_BBS_HomeUrl & "inc/BBSSetup.asp")
If GBL_CHK_TempStr <> "" Then
	Check_com = false
	readflag = false
	Response.Write "<span class=""redfont"">��</span> <br /><span class=""grayfont"">(" & htmlEncode(GBL_CHK_TempStr) & ")</span>"
Else
	If inStr(filestr,"DEF_UsedDataBase") > 0 then
		readflag = true
		Response.Write "<font color=green class=greenfont>��</font>"
	else
		readflag = false
		Check_com = false
		Response.Write "<span class=""redfont"">��</span> <span class=""grayfont"">(δ����ȷ��ȡ�ļ�)</span>"
	end if
End If
%>
</li>
<li>дȨ�޼�⣺<%
If readflag = false then
	Response.Write "<span class=""redfont"">��</span> <span class=""grayfont"">(���ȱ���߱��������Ķ�Ȩ��)</span>"
else
	call ADODB_SaveToFile(filestr,DEF_BBS_HomeUrl & "inc/BBSSetup.asp")
	If GBL_CHK_TempStr <> "" Then
		Check_com = false
		Response.Write "<span class=""redfont"">��</span> <br /><span class=""grayfont"">(" & htmlEncode(GBL_CHK_TempStr) & ")</span>"
	Else
		Response.Write "<font color=green class=greenfont>��</font>"
	End If
end if
%>
</li>
</ol>
<div class="contenttitle">���¼����̳��չ��Ҫ������ʵ��������ܻ����Ĭ������</div>
<ol>
<li><%call CheckObjInstalled("Persits.Jpeg",1)%></li>
<li><%call CheckObjInstalled("leadbbs.bbsCode",1)%></li>
<li><%call CheckObjInstalled("JMail.SMTPMail",1)%></li>
</ol>
<%if Check_com = true then%>
<a href="default.asp?step=3" class="install_submit">������װ</a>
<%else%>
<a href="javascript:;" onclick="alert('ֻ��ͨ����һ���ּ����ܼ�����װ�����Ƚ����Ӧ�ռ���������.');" class="install_submit_disbale">ȫ��ͨ����һ���ּ����ܼ���</a>
<%
end if

end sub

function filterStr(str)

	str = replace(str,"""","")
	str = replace(str,"'","")
	str = replace(str,"<","")
	filterStr = str

end function

sub OpenDatabase(constr)

	on error resume next
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.ConnectionString = constr
	Con.Open
	If Err Then
		GBL_CHK_TempStr = "<p>��������: <font color=red>" & err.description & "</font></p>"
	End If

End Sub

Sub CloseDatabase

	on error resume next
	Con.Close
	Set con = Nothing
	
End Sub

function mapFile(file)

	mapFile = Server.MapPath(file)

End function

sub install_step3form

	dim server,port,uid,pwd,databasename,mysqlversion
	dim mysql_server,mysql_port,mysql_uid,mysql_pwd,mysql_databasename
	dim submitflag
	dtype = 2
	dtype = left(request.form("dtype"),1)
	server = left(request.form("server"),32)
	port = toNum(left(request.form("port"),6),0)
	uid = left(request.form("uid"),32)
	pwd = left(request.form("pwd"),32)
	databasename = left(request.form("databasename"),32)
	mysqlversion = left(request.form("mysqlversion"),1)
	
	
	mysql_server = left(request.form("mysql_server"),32)
	mysql_port = toNum(left(request.form("mysql_port"),6),0)
	mysql_uid = left(request.form("mysql_uid"),32)
	mysql_pwd = left(request.form("mysql_pwd"),32)
	mysql_databasename = left(request.form("mysql_databasename"),32)
	
	submitflag = left(request.form("submitflag"),4)

	if dtype = "0" then
		dtype = 0
	elseif dtype = "1" then
		dtype = 1
	else
		dtype = 2
	end if

	if mysqlversion = "1" then
		mysqlversion = 1
	else
		mysqlversion = 0
	end if
	
	if submitflag = "true" then
		select case dtype
			case 0:
				setupstr = "Provider=SQLOLEDB;Data Source=" & filterStr(server)
				if port > 0 then
					setupstr = setupstr & "," & port & ";"
				else
					setupstr = setupstr & ";"
				end if
				setupstr = setupstr & "uid=" & filterStr(uid) & ";pwd=" & filterStr(pwd) & ";database=" & filterStr(databasename) & ";"
				constr = setupstr
				GBL_CHK_TempStr = "<font class=redfont>�ݲ�֧��mssql�汾��װ.</font>"
			case 1:
				setupstr = "data/global.asa"
				constr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & mapFile(DEF_BBS_HomeUrl & setupstr)
			case 2:
				if mysqlversion = 0 then
					setupstr = "Driver={Mysql ODBC 5.1 Driver};SERVER=" & filterStr(mysql_server) & ";"
					if mysql_port > 0 then
						setupstr = setupstr & "PORT=" & mysql_port & ";"
					end if
					setupstr = setupstr & "DATABASE=" & filterStr(mysql_databasename) & ";UID=" & filterStr(mysql_uid) & ";PWD=" & filterStr(mysql_pwd) & ";charset=gbk;"
				else
				end if
				constr = setupstr
		end select

		If dtype <> 0 Then
			GBL_CHK_TempStr = ""
			OpenDatabase(constr)
			CloseDatabase
			If GBL_CHK_TempStr = "" Then
				install_step4form
				exit sub
			end if
		End If
	end if
	%>
	<script type="text/JavaScript">
	$id = function(id){
	return document.getElementById(id);
	}
	function changedatabase(f)
	{
		if(f==1)
		$id("access").style.display="block";
		else
		$id("access").style.display="none";
		
		if(f==0)
		$id("mssql").style.display="block";
		else
		$id("mssql").style.display="none";
		
		if(f==2)
		$id("mysql").style.display="block";
		else
		$id("mysql").style.display="none";
	}

	var ValidationPassed = true;
	function submitonce(theform)
	{
		if(ValidationPassed == false)return;
		submit_disable(theform);
		$id('LeadBBSFm').submit();
	}
	function submit_disable(theform,tp)
	{
		ValidationPassed = false;
	}
	</script>
	<div class="contenttitle" id="databasecontenttitle">�������ݿ�</div>
	<%If GBL_CHK_TempStr <> "" Then
		%>
		<div class="line">
			�������ݿ�ʧ�ܣ��޷�������װ���������ԭ��
		</div>
		<%
		Response.Write GBL_CHK_TempStr
		End If%>

	<form action="default.asp?step=3" method="post" id="LeadBBSFm" name="LeadBBSFm" onsubmit="submitonce(this);return ValidationPassed;">
	<input name="submitflag" value="true" type="hidden" />

<div id="selectdatabase">
	<div class="line">
	<span class="name">���ݿ����ͣ�</span>
	<input class=fmchkbox type=radio name=dtype value=1 <%If dtype = 1 Then Response.Write " checked"%> onclick="changedatabase(1)" />Microsoft Access(mdb)
	<span style="display:none;">
	<input class=fmchkbox type=radio name=dtype value=0 <%If dtype = 0 Then Response.Write " checked"%> onclick="changedatabase(0)" />Microsoft SQL Server
	</span>
	<input class=fmchkbox type=radio name=dtype value=2 <%If dtype = 2 Then Response.Write " checked"%> onclick="changedatabase(2)" />MySQL
	</div>
</div>
	
	<div class="line" id="access"<%
	if dtype <> 1 then response.Write " style=""display:none;"""
	%>>
	<span class="name">�ļ�·����</span>
	<input maxlength=255 type="text" id=Form_Title name=accessfile size="49" value="data/global.asa" disabled="true"> <span class="info">���ɸ���</span>
	</div>

	<div id="mssql"<%
	if dtype <> 0 then response.Write " style=""display:none;"""
	%>>
	<div class="line">
	<span class="name">
	SERVER��
	</span><input maxlength=30 type="text" name=server size="25" value="<%=htmlEncode(server)%>"> <span class="info">���ݿ������IP��ַ����WEBͬ������дlocalhost</span>
	</div>
	
	<div class="line">
	<span class="name">
	�˿ڣ�
	</span><input maxlength=30 type="text" name=port size="25" value="<%if port > 0 then response.write htmlEncode(port)%>"> <span class="info">�˿ںţ���ΪĬ�ϲ�����д</span>
	</div>
	
	<div class="line">
	<span class="name">
	�û�����
	</span><input maxlength=30 type="text" name=uid size="25" value="<%=htmlEncode(uid)%>"> <span class="info">�������ݿ���û���(UID)</span>
	</div>
	
	<div class="line">
	<span class="name">
	���룺
	</span><input maxlength=30 type="password" name=pwd size="25" value="<%=htmlEncode(pwd)%>"> <span class="info">����(PWD)</span>
	</div>
	
	<div class="line">
	<span class="name">
	���ݿ����ƣ�
	</span><input maxlength=30 type="text" name=databasename size="25" value="<%=htmlEncode(databasename)%>"> <span class="info">��Ҫָ�����ݿ�����</span>
	</div>
	</div>
	
	

	<div id="mysql"<%
	if dtype <> 2 then response.Write " style=""display:none;"""
	%>>
	
	<div class="line">
	<span class="name">
	ODBC�汾��
	</span>
	<input class=fmchkbox type=radio name=mysqlversion value=0 <%If mysqlversion = 0 Then Response.Write " checked"%> />Mysql ODBC 5.1 Driver
	<input class=fmchkbox type=radio name=mysqlversion value=1 <%If mysqlversion = 1 Then Response.Write " checked"%> />Mysql ODBC 3.51 Driver
	<span class="info">������ѯ�ռ��ṩ��</span>
	</div>
	
	<div class="line">
	<span class="name">
	SERVER��
	</span><input maxlength=30 type="text" name=mysql_server size="25" value="<%=htmlEncode(mysql_server)%>"> <span class="info">���ݿ������IP��ַ��ͬWEB�ı�������дlocalhost</span>
	</div>
	
	<div class="line">
	<span class="name">
	�˿ڣ�
	</span><input maxlength=30 type="text" name=mysql_port size="25" value="<%if mysql_port > 0 then response.write htmlEncode(mysql_port)%>"> <span class="info">�˿ںţ���ΪĬ�ϲ�����д</span>
	</div>
	
	<div class="line">
	<span class="name">
	�û�����
	</span><input maxlength=30 type="text" name=mysql_uid size="25" value="<%=htmlEncode(mysql_uid)%>"> <span class="info">�������ݿ���û���(UID)</span>
	</div>
	
	<div class="line">
	<span class="name">
	���룺
	</span><input maxlength=30 type="password" name=mysql_pwd size="25" value="<%=htmlEncode(mysql_pwd)%>"> <span class="info">����(PWD)</span>
	</div>
	
	<div class="line">
	<span class="name">
	���ݿ����ƣ�
	</span><input maxlength=30 type="text" name=mysql_databasename size="25" value="<%=htmlEncode(mysql_databasename)%>"> <span class="info">��Ҫָ�����ݿ�����</span>
	</div>
	
	</div>
	<a href="javascript:;" onclick="submitonce(this);" class="install_submit">�ύ����</a>
	</form>
	<%

End sub


Sub install_step4form

	adminuser = left(request.form("adminuser"),20)
	adminpassword = left(request.form("adminpassword"),20)
	adminpassword2 = left(request.form("adminpassword2"),20)
	
	If constr = "" then constr = left(request.form("constr"),255)
	If dtype = "" then dtype = left(request.form("dtype"),20)
	If setupstr = "" then setupstr = left(request.form("setupstr"),255)
	
	if request.form("submitflag") = "true2" then
		if adminuser = "" then
			GBL_CHK_TempStr = "����д����Ա�˺�."
		else
			CheckUsername(adminuser)
		end if
		if adminpassword <> adminpassword2 then
			GBL_CHK_TempStr = GBL_CHK_TempStr & " ������������δ��ͬ."
		Else
			if adminpassword = "" or adminpassword2 = "" then
				GBL_CHK_TempStr = GBL_CHK_TempStr & " ����д����."
			End If
		End If
		if GBL_CHK_TempStr = "" then
			install_step5form
			exit sub
		end if
	end if
%>
	<script type="text/JavaScript">
	$id = function(id){
	return document.getElementById(id);
	}
	var ValidationPassed = true;
	function submitonce(theform)
	{
		if(ValidationPassed == false)return;
		submit_disable(theform);
		alert("�ύ�󣬿�����Ҫ������ʱ�����������ã����д������滻BBSSetup.aspΪԭʼ�ļ������½��а�װ.");
		$id('LeadBBSFm').submit();
	}
	function submit_disable(theform,tp)
	{
		ValidationPassed = false;
	}
	$id('step4').className = "on"
	</script>
	<div class="contenttitle" id="databasecontenttitle">���ù���Ա</div>
	<%If GBL_CHK_TempStr <> "" Then
		%>
		<div class="line">
			�����Ϣ��дδ��ͨ����֤��
		</div>
		<%
		Response.Write "<span class=redfont>" & GBL_CHK_TempStr & "</span>"
		End If%>

	<form action="default.asp?step=4" method="post" id="LeadBBSFm" name="LeadBBSFm" onsubmit="submitonce(this);return ValidationPassed;">
	<input name="submitflag" value="true2" type="hidden" />
	<input name="constr" value="<%=htmlencode(constr)%>" type="hidden" />
	<input name="dtype" value="<%=htmlencode(dtype)%>" type="hidden" />
	<input name="setupstr" value="<%=htmlencode(setupstr)%>" type="hidden" />
	
	<div class="line">
	<span class="name">
	����Ա�˺ţ�
	</span><input maxlength=30 type="text" name=adminuser size="25" value="<%=htmlEncode(adminuser)%>"> <span class="info">���û�������Ϊ��̳��������Ա��ע��</span>
	</div>
	
	<div class="line">
	<span class="name">
	���룺
	</span><input maxlength=20 type="password" name=adminpassword size="25" value="<%=htmlEncode(adminpassword)%>"> <span class="info">��������Ա����</span>
	</div>
	
	<div class="line">
	<span class="name">
	�ٴ��������룺
	</span><input maxlength=20 type="password" name=adminpassword2 size="25" value="<%=htmlEncode(adminpassword2)%>"> <span class="info">�����������������ͬ</span>
	</div>
	
	<a href="javascript:;" onclick="submitonce(this);" class="install_submit">�ύ����</a>
	</form>

<%

End Sub

function exesql(sql,flag)

	if sql = "" then exit function
	on error resume next
	If flag = 0 or flag = 3 Then
		Set exesql = Con.ExeCute(SQL)
	Else
		Con.ExeCute(SQL)
	End If
	If Err Then
		'printline("<p>����SQL���ִ�г�������������ֹ������ϵ�ٷ������</p><p><font color=gray>" & server.htmlencode(SQL) & "</font></P>")
		printline("<p>��������: <font color=red>" & err.description & "</font></p>")
		Err.Clear
	End If

end function

sub install_step5form

%>

	<script type="text/JavaScript">
	$(window).resize();
	</script>
	<script type="text/JavaScript">
	$id = function(id){
	return document.getElementById(id);
	}
	$id('step5').className = "on"
	</script>
<%
	if dtype = "1" then
		dtype = 1
	else
		dtype = 2
	end if
	
	setupstr = filterStr(setupstr)
	constr = filterStr(constr)
	
	OpenDatabase(constr)
	
	If GBL_CHK_TempStr <> "" Then
		Response.Write "<span class=err>��װʧ�ܣ��뷵�����°�װ.</span>"&GBL_CHK_TempStr
		CloseDatabase
		exit sub
	end if
	
	printline("<b>�����ĵȺ���ɰ�װ....</b>")
	dim filestr,arr,n
	select case dtype
		case 2:
			filestr = ADODB_LoadFile("database/mysql.sql")
			arr = split(filestr,";")
			printline("��ʼ��ʼ�����ݿ�...")
			%>
			<div class=errstr>
			<%
			for n = 0 to ubound(arr)
				call exesql(arr(n),1)
			next
			%>
			</div>
			<%
			printline("����ɳ�ʼ�����ݿ�.")
	end select
	
	dim sql,rs
	sql = "select id from leadbbs_user where username='" & filterStr(adminuser) & "'"
	set rs = exesql(sql,0)
	if rs.eof then
		sql = "insert into leadbbs_user(username,pass,answer) values('" & filterStr(adminuser) & "','" & md5(adminpassword) & "','" & md5(adminpassword) & "')"
		call exesql(sql,1)
	else
		printline("����Ա�˺�" & htmlencode(adminuser) & "�Ѵ��ڣ��Թ����.")
	end if
	rs.close
	set rs = nothing
	
	printline("�ѳ�ʼ������Ա.")
	
	CloseDatabase
	filestr = ADODB_LoadFile(DEF_BBS_HomeUrl & "inc/BBSSetup.asp")
	filestr = replace(filestr,"Const DEF_AccessDatabase = """"","Const DEF_AccessDatabase = """ & filterStr(setupstr) & """")
	filestr = replace(filestr,"const DEF_UsedDataBase = 1","const DEF_UsedDataBase = " & dtype)
	filestr = replace(filestr,"const DEF_UsedDataBase = 0","const DEF_UsedDataBase = " & dtype)
	filestr = replace(filestr,"const DEF_UsedDataBase = 2","const DEF_UsedDataBase = " & dtype)
	printline("������ݿ�����.")
	
	if CheckObjInstalled("Persits.Jpeg",0) = True then
		filestr = replace(filestr,"const DEF_EnableGFL = 0","const DEF_EnableGFL = 1")
		printline("�ѿ���aspJpeg�������֧��.")
	else
		filestr = replace(filestr,"const DEF_EnableGFL = 1","const DEF_EnableGFL = 0")
		printline("�ѽ���aspJpeg�������֧��.")
	End If

	filestr = replace(filestr,"const DEF_SupervisorUserName = "",Admin,""","const DEF_SupervisorUserName = ""," & filterStr(adminuser) & ",""")
	

	filestr = replace(filestr,"Response.Redirect ""install/default.asp""","")
	call ADODB_SaveToFile(filestr,DEF_BBS_HomeUrl & "inc/BBSSetup.asp")
	printline("<b>��ɰ�װ.</b>")
	printline("<b>��������ʹ�ô����Ĺ���Ա�����̨�½����༰������Ϣ��</b>")
	%>
	<a href="<%=DEF_BBS_HomeUrl%>manage/" class="install_submit">��������̨����</a>
	<%
	

end sub

sub printline(str)

	%>
	<div class="line"><%=str%></div>
	<%
	response.flush

end sub

Function CheckUsername(Form_UserName)


					Dim TempChar,TempURL,Loop_N
					TempURL = Len(Form_UserName)
					For Loop_N = 1 to TempURL
						TempChar = ASC(Mid(Form_UserName,Loop_N,1))
						If TempChar < 0 Then TempChar = TempChar + 65535
							If TempChar = 32 Then
								If TempURL > Len(Replace(Form_UserName," ","")) + 2 Then '������������ո��Ҳ�ͬʱ��һ��
									CheckUsername = 0
									GBL_CHK_TempStr = "�û������ֻ����ʹ�������ո�!<br>"
									Exit Function
								End If
							Else
								If TempChar < 45 or (TempChar>45 and TempChar<48) or (TempChar>57 and TempChar<65) or (TempChar>90 and TempChar < 95) or TempChar = 96 or (TempChar > 122 and TempChar < 33088) Then
									GBL_CHK_TempStr = "�û������зǷ��ַ�(��ʹ������,��ĸ,�»���)!<br>"
									CheckUsername = 0
									Exit Function
								End If
							End If
						
						
						If TempChar > 65184 Then
							GBL_CHK_TempStr = "�Ƿ����û���,���зǷ��ַ�,��ȷ��!<br>"
							CheckUsername = 0
							Exit Function
						End If
					Next
					CheckUsername = 1

End function
%>