<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Board_Popfun.asp -->
<%
DEF_BBS_HomeUrl = "../../"


Main

Sub Main

	Select Case left(Request("file"),5)
		Case ""
			Plug_MusicBar_Default
		Case "music"
			Plug_MusicBar_Music
		Case "list"
			Plug_MusicBar_List
		Case "edit"
			Plug_MusicBar_Edit
	End Select

End Sub

Plug_MusicBar_Default

Sub Plug_MusicBar_Default

	Dim GBL_FrameUrl
	GBL_FrameUrl = "http://" & LCase(Request.ServerVariables("server_name")) & Replace(LCase(Request.Servervariables("SCRIPT_NAME")),"plug-ins/musicbar/default.asp","")
	If Left(LCase(Request.ServerVariables("HTTP_REFERER")),Len(GBL_FrameUrl)) = GBL_FrameUrl Then
		GBL_FrameUrl = Request.ServerVariables("HTTP_REFERER")
	Else
		GBL_FrameUrl = "../../Boards.asp"
	End If
%>
<html><head><title><%=DEF_SiteNameString%></title></head>
<frameset rows="*,22" cols="*" framespacing="0" frameborder="NO" border="0">
<frame src="<%=GBL_FrameUrl%>" name="mainFrame" scrolling="yes">
<frame src="Default.asp?file=music" name="bottomFrame" scrolling="yes" noresize></frameset><noframes>
<body></body>
</noframes></html>

<%End Sub

Sub Plug_MusicBar_Music

	SiteHead("     ")%>
<style TYPE="text/css">
<!--
a:link,a:active,a:visited
{
	 color: silver;
	 text-decoration: none ;
}
a:hover 
{
	 color: white;
	 text-decoration: none ;
}
td {FONT-SIZE: 9pt;  FONT-FAMILY: "Verdana"; color:#333333}
.bg{background-image:url('images/dot.gif'); background-repeat:repeat-x; background-position:bottom; height:18px }
.bg2{background-image:url('images/bar.gif'); background-repeat:repeat-x; background-position:bottom; height:18px }
.time
{
	font-family: "arial",;
	 font-size: 9pt;
	 color:#333333;
}
-->
</style>
<script languange="Javascript" src="js/bud.js"></script>
<script languange="Javascript" src="js/time.js"></script>
<script languange ="Javascript" src="js/imgchg.js"></script>
<script languange="Javascript">
var blnAutoStart = true;
var blnRndPlay = true; 
var blnStatusBar = false; 
var blnShowVolCtrl = true;
var blnShowPlist = true;
var blnUseSmi = false;
var blnLooptrk = true;
var blnShowMmInfo = false;
</script>
<script language="JavaScript" src="js/ctrchg.js"></script>
<script languange ="Jscript" FOR=Exobud EVENT=openStateChange(sf)> evtOSChg(sf); </script>
<script languange ="Jscript" FOR=Exobud EVENT=playStateChange(ns)> evtPSChg(ns); </script>
<script languange ="Jscript" FOR=Exobud EVENT=error()> evtWmpError(); </script>
<script languange ="Jscript" FOR=Exobud EVENT=Buffering(bf)> evtWmpBuff(bf); </script>
<body onLoad="initExobud();show5();" ondragstart="return false" onselectstart="return false" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0 class=body>
<OBJECT ID=Exobud CLASSID="CLSID:6BF52A52-394A-11d3-B153-00C04F79FAA6"
	type="application/x-oleobject"	width="0" height="0" 
	style="position:relative;left:0px; top:0px;width:0px;height:0px;">
	<param name="autoStart" value="true">
	<param name="balance" value="0">
	<param name="currentPosition" value="0">
	<param name="currentMarker" value="0">
	<param name="enableContextMenu" value="false">
	<param name="enableErrorDialogs" value="false">
	<param name="enabled" value="true">
	<param name="fullScreen" value="false">
	<param name="invokeURLs" value="false">
	<param name="mute" value="false">
	<param name="playCount" value="1">
	<param name="rate" value="1">
	<param name="uiMode" value="none">
	<param name="volume" value="100">
</OBJECT>
<table width=<%=DEF_BBS_ScreenWidth%> height=20  border=0 align="center" cellpadding=0 cellspacing=0>
<tr><td>
<table border=0 cellpadding=0 cellspacing=0><tr>
	<td width=21 height=20   align=center valign=middle NOWRAP><a href=# onclick="top.location=parent.mainFrame.document.location;"><img name="scope" src="images/m1.gif" border=0 alt="�˳��������� -- CnSide MP" border=0></a></td>
	<td class=bg width=7 height=20  NOWRAP>&nbsp;</td>
	<td class=bg valign=middle  width=203 height=20 NOWRAP> <marquee behavior="scroll"  scrolldelay=70 scrollamount=2 width=215 height=12>
		<span id="disp1" width=215 class="title" align=left>Music Player</span> 
		<span id=liveclock width=150></span> </marquee> </td>
	<td width=7 height=20 NOWRAP>&nbsp;</td>
	<td width=7 height=20 NOWRAP>&nbsp;</td>
	<td class=bg width=102 height=20 align=center valign=middle background="img/bg2c.gif" NOWRAP onclick="chgTimeFmt();this.blur();"> 
		<span id="disp2" width=105 class="time" align="center" title ="ʱ����ʾ(Elaps/Laps)" style="cursor:hand;">00:00 
		| 00:00</span> </td>
	<td width=7 height=20 NOWRAP>&nbsp;</td>
	<td width=24 height=20 valign=middle NOWRAP><img name="vmute" src="images/volume.gif" border=0 width=14 height=12 onclick="wmpMute();this.blur();" onMouseOver="imgtog('vmute',2);" onMouseOut="imgtog('vmute',3);" style="cursor:hand;" title="����(Mute)"></td>
	<td width=18 height=20 align="center" valign=bottom NOWRAP class=bg2><img name="vdn" src="images/left.gif" border=0 width=6 height=6 onclick="wmpVolDn();this.blur();" onMouseOver="imgtog('vdn',2);" onMouseOut="imgtog('vdn',3)" style="cursor:hand;" title="����"></td>
	<td width=18 height=20 align="center" valign=bottom NOWRAP class=bg2><img name="vup" src="images/right.gif" border=0 width=6 height=6 onclick="wmpVolUp();this.blur();" onMouseOver="imgtog('vup',2);" onMouseOut="imgtog('vup',3)" style="cursor:hand;" title="����"></td>
	<td width=4 height=20  NOWRAP>&nbsp;</td>
	<td height=20 width=10 valign=middle NOWRAP><img src=../../null.gif height=2 width=2><br><img name="pmode" src="images/r.gif" border=0 width=10 height=10 onclick="chgPMode();this.blur();" onMouseOver="imgtog('pmode',2);" onMouseOut="imgtog('pmode',3)" style="cursor:hand;" title="ģʽ"></td>
	<td width=14 height=20 align="center" valign=middle NOWRAP><img src=../../null.gif height=2 width=2><br><img name="rept" src="images/sel.gif" border=0 width=10 height=10 onclick="chkRept();this.blur();" onMouseOver="imgtog('rept',2);" onMouseOut="imgtog('rept',3)" style="cursor:hand;" title="ѭ��"></td>
	<td width=37 height=20  NOWRAP>&nbsp;</td>
	<td height=20 width=19 valign=middle NOWRAP><img src=../../null.gif height=3 width=2><br><img src="images/rew.gif" alt="ǰһ��" name="prevt" width=13 height=13 border=0 style="cursor:hand;" title="��һ��" onclick="playPrev();this.blur();" onMouseOver="imgtog('prevt',2);" onMouseOut="imgtog('prevt',3)"></td>
	<td height=20 width=18 valign=middle NOWRAP><img src=../../null.gif height=3 width=2><br><img src="images/stop2.gif" alt="��ͣ" name="pauzt" width=13 height=13 border=0 style="cursor:hand;" title="��ͣ/����" onclick="wmpPP();this.blur();" onMouseOver="imgtog('pauzt',2);" onMouseOut="imgtog('pauzt',3)"></td>
	<td height=20 width=18 valign=middle NOWRAP><img src=../../null.gif height=3 width=2><br><img src="images/stop.gif" alt="ֹͣ" name="stopt" width=13 height=13 border=0 style="cursor:hand;" title="ֹͣ" onclick="wmpStop();this.blur();" onMouseOver="imgtog('stopt',2);" onMouseOut="imgtog('stopt',3)"></td>
	<td height=20 width=19 valign=middle NOWRAP><img src=../../null.gif height=3 width=2><br><img src="images/play.gif" alt="����" name="playt" width=13 height=13 border=0 style="cursor:hand;" title="����" onclick="startExobud();this.blur();" onMouseOver="imgtog('playt',2);" onMouseOut="imgtog('playt',3)"></td>
	<td height=20 width=17 valign=middle NOWRAP><img src=../../null.gif height=3 width=2><br><img src="images/ffw.gif" alt="��һ��" name="nextt" width=13 height=13 border=0 style="cursor:hand;" title="��һ��" onclick="playNext();this.blur();" onMouseOver="imgtog('nextt',2);" onMouseOut="imgtog('nextt',3)"></td>
	<td width="30" height=20 valign=middle NOWRAP><img src=../../null.gif height=3 width=2><br><img src="images/list.gif" alt="�����б�" name="plist" width=13 height=13 border=0 style="cursor:hand;" title="�����б�" onclick="openPlist();this.blur();" onMouseOver="imgtog('plist',2);" onMouseOut="imgtog('plist',3)"></td>
	<td width="30" height=20 valign=middle NOWRAP><img src=../../null.gif height=3 width=2><br><a 
		href=Default.asp?file=edit target=mainFrame><img src="images/edit.gif" alt="�༭�����б�" width=13 height=13 border=0 style="cursor:hand;"></a></td>
</tr></table></td></tr>
<tr>
	<td colspan=26 height=0> <div id="capText" style="width:100%;height:60;font-size:11px;color:white;background-color:#555555;display:none;"> 
		<p>LeadBBS.COM MP Player</div></td>
</tr>
</table>
</body>
</html>

<%End Sub

Sub Plug_MusicBar_List

	SiteHead("     ")%>
<script languange ="Javascript" src="js/list.js"></script>
</head>
<body onLoad="dspList();this.focus();" onDragStart="return false" onSelectStart="return false" text="silver" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0 style="border: 0px solid black; margin: 0pt;">
<br>
<table border=0 cellpadding=0 cellspacing=0 width=90% align=center>
<tr>
	<td class=font height="25" align=center valign=middle>�� �� ѡ ��</td>
</tr>
<tr><td width=90% height=*% background="white" valign=top align=left ><div id="mmList"></div></td></tr>
<tr>
	<td height=20 valign=top> 
		<div id="pageList"></div>
		<div align="center"></div>
	</td>
</tr>
</table>
<table border=0 cellpadding=0 cellspacing=0 width=90% align=center>
<tr>
	<td height=25>
		<span id="pageInfo"></span>
		<br><br><a href="#" onclick="chkSel();" onFocus="this.blur()" title="ȫѡ">ȫѡ</a>
		<a href="#" onclick="chkDesel()" onFocus="this.blur()" title="���">���</a>
		<a href="#" onclick="playSel();" onFocus="this.blur()" title="����">����</a> 
		<a href="#" onClick = "window.close();" onFocus='this.blur()' title="�ر�">�ر�</a>
	</td>
</tr>
</table>
<%
	SiteBottom_Spend

End Sub

Sub Plug_MusicBar_Edit

	Dim Master
	Dim EditMod,MyStr
	EditMod = Request("EditMod")
	MyStr = "<table width=600><tr><td><b>���߱༭�����б��ļ�bglist.js</b>--<font color=red>ע������</font></td></tr>"
	If EditMod = "0" Then
		MyStr = MyStr&"<tr><td><li>��ʽ mkList('�����ļ�·��','���ֽ�Ŀ����','��Ļ��ַ','�Ƿ񲥷�(fΪ������,��������)');<li>���� mkList('hi/LeadBBS_H1.mp3','Welcome to LeadBBS.','','');</td></tr>"
	Else
		EditMod = "1"
		MyStr = MyStr&"<tr><td><li>��������������ַ���գ������б���ɾ��;<li>�������ֱ�����ڱ�ŵ���С���ֵ֮��,�ظ����߳��������б���ɾ��;</td></tr>"
	End If
	MyStr = MyStr&"<tr><td><input type=button onclick=""location='default.asp?file=edit&editmod=0'"" value=""�ı�ģʽ"" class=fmbtn>&nbsp;&nbsp;<input type=button onclick=""location='default.asp?file=edit&editmod=1'"" value=""�б�ģʽ"" class=fmbtn></td></tr></table>"
	
	InitDatabase
	BBS_SiteHead DEF_SiteNameString &" - ���ֲ���",GBL_board_ID," >> ��� >> ���ֲ���"
	If CheckSupervisorUserName = 1 Then
		Master = True
	Else
		Master = False
	End If
	CloseDatabase
	
	Global_TableHead
	%>
	<table cellpadding=3 cellspacing=1 align=center width="<%=DEF_BBS_ScreenWidth%>" class=TBone>
	<tr class=TBHead height=24>
		<td align=center><b><font class=HeadFont>�ٷ���LeadBBS���ֲ��Ų�� ��������</font><b></td>
	</tr>
	<%
	If GBL_CHK_User = "" or Master = False Then
		Response.Write "<tr class=TBBG1><td><table cellpadding=3 cellspacing=4><tr><td>���������ԭ������ǣ�<br><br><li>�㲻�ǹ���Ա����Ȩ���룡</li><li>������ǹ���Ա�����Թ���Ա���<a href='" & DEF_BBS_HomeUrl & "User/Login.asp?Relogin=Yes&u=" & urlencode(Request.Servervariables("SCRIPT_NAME") & "?" & Request.QueryString) & "'><font class=NavColor>�ص�¼</font></a>��</li></td></tr></table></td></tr></table>"
	Else
		%>
		<tr class=TBBG1><td align=center>
		<br>
		<%
		DisplayEditFileContent "bglist.js",MyStr,"edit",EditMod
		%>
		</td></tr>
		<%
	End If
	%>
	</table>
	<%
	Global_TableBottom
	SiteBottom
	If GBL_ShowBottomSure = 0 Then GBL_SiteBottomString = ""
	Response.Write GBL_SiteBottomString

End Sub

'FileName ���·��
'TemStr �༭ע��
'FilePar ���ش��ݲ���
Sub DisplayEditFileContent(FileName,TmpStr,FilePar,eMod)

	'If DEF_FSOString = "" Then
	'	Response.Write "<p><br><font color=red class=redfont>��̳�����óɲ�֧�����߱༭�ļ�����!</font></p>" & VbCrLf
	'	Exit Sub
	'End If
	Dim fileContent

	If Request.Form("SubmitFlag") = "" Then
		FileContent = ADODB_LoadFile(FileName)
		If GBL_CHK_TempStr <> "" Then
			Response.Write "<p>" & GBL_CHK_TempStr
			GBL_CHK_TempStr = ""
			Exit Sub
		End If
	Else
		If eMod="0" Then
			fileContent = Request.Form("fileContent")
		Else
			Dim i,mI,mName,mLink,mF,mO
			mI=Request.Form("mName").count
			For i=1 to mI
				fileContent=fileContent&"//"&i&"//"&vbcrlf
			Next
			For i=1 to mI
				If request.Form("mName")(i)<>"" And request.Form("mLink")(i)<>"" Then
					fileContent=Replace(fileContent,"//"&request.Form("mO")(i)&"//","mkList('"&request.Form("mLink")(i)&"','"&request.Form("mName")(i)&"','','"&cht(request.Form("mF"&i))&"');")
				Else
					fileContent=Replace(fileContent,","&request.Form("mO")(i)&","&vbcrlf,"")
				End If
			Next
		End If
		Dim TempContent
		TempContent = Lcase(fileContent)
		If inStr(TempContent,"<%") or inStr(TempContent,"include") or inStr(TempContent,"server") Then
			Response.Write "<p><br><font color=red class=redfont>�����в��ܺ���<%��include��Server���ַ�!</font></p>" & VbCrLf
			Exit Sub
		End If

		ADODB_SaveToFile fileContent,FileName
		If GBL_CHK_TempStr <> "" Then
			Response.Write "<p>" & GBL_CHK_TempStr
			GBL_CHK_TempStr = ""
			Exit Sub
		Else
			Response.Write "<p><font color=green class=greenfont><b>�ɹ������ļ����ݣ�</b></font></p>" & VbCrLf
		End If
	End If
	%>
	<form action=Default.asp method=post>
		<%=TmpStr%><p>
		<input type=hidden value=<%=FilePar%> name=file>
		<input type=hidden value=yes name=SubmitFlag>
		<input type=hidden value=<%=eMod%> name="EditMod">
		<%If eMod = "0" Then%>
		<textarea name="fileContent" cols="96" rows="20" class=fmtxtra><%If fileContent <> "" Then Response.Write VbCrLf & server.htmlEncode(fileContent)%></textarea></p>
		<%Else%>
		<style>
		#list td{ text-align:center;border:1px dashed #000000; height:24px}
		.txt{ border:1px solid #000000; text-align:center; width:220px}
		</style>
		<table width=600 id="list"><tr><td width=30>���</td><td width=240>��������</td><td width=240>������ַ</td><td width=60>�Ƿ񲥷�</td><td width=30>����</td></tr><%
		Dim objRegExp,Matches,j
		Set objRegExp=New RegExp
		objRegExp.IgnoreCase =True
		objRegExp.Global=True
		objRegExp.Pattern="mkList\(\'(.*?)\',\'(.*?)\',\'\',\'(.*?)\'\);"
		Set Matches = objRegExp.Execute(fileContent)
		For j = 0 to Matches.Count - 1
			Response.Write "<tr><td>"&j+1&"</td><td><input name=mName type=text class=txt value="""&Matches(j).SubMatches(1)&"""></td><td><input name=mLink type=text class=txt value="""&Matches(j).SubMatches(0)&"""></td><td><input name=mF"&j+1&" type=checkbox value=""t"" "&Chf(Matches(j).SubMatches(2))&"></td><td><input name=mO type=text class=txt style=""width:20px"" value="""&j+1&"""></td></tr>"&vbcrlf
		Next				
		%></table><div align="left" style="margin-left:70px"><input type="button" onclick="AddRow()" value="����һ��" class=fmbtn></div>
		<Script Language="Javascript">
		function AddRow()
		{
		var i =list.rows.length;
		var newTr = list.insertRow();
		var newTd0 = newTr.insertCell();
		var newTd1 = newTr.insertCell();
		var newTd2 = newTr.insertCell();
		var newTd3 = newTr.insertCell();
		var newTd4 = newTr.insertCell();
		newTd0.innerHTML = i; 
		newTd1.innerHTML= "<input name=mName type=text class=txt value=\"\">";
		newTd2.innerHTML= "<input name=mLink type=text class=txt value=\"\">";
		newTd3.innerHTML= "<input name=mF"+i+" type=checkbox value=\"t\" checked>";
		newTd4.innerHTML= "<input name=mO type=text class=txt style=\"width:20px\" value=\""+i+"\">";
		}
		</script>
		<%End If%>
		<input type="submit" name="save" value="�ύ�༭" class=fmbtn>
		<input type="reset" name="Reset" value="ȡ��" class=fmbtn>
		</form>
	<%

End Sub

Function cht(str)

	If str="t" then
		cht=""
	Else
		cht="f"
	End If

End Function

Function chf(str)

	If str="" then
		chf="checked"
	Else
		chf=""
	End If

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
				GBL_CHK_TempStr = "<br>��ȡ�ļ�ʧ�ܣ�" & err.description & "<br>�������ܣ�ȷ���Ƿ�Դ��ļ��ж�ȡȨ��."
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
			GBL_CHK_TempStr = "<div align='center'>����������֧��ADODB.Stream���޷���ɲ��������ֹ�����</div>"
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
		GBL_CHK_TempStr = "<br>������Ϣ��" & err.description & "<br>�������ܣ�ȷ���Ƿ�Դ��ļ��ж�ȡȨ��."
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
			GBL_CHK_TempStr = "<div align='center'>����������֧��ADODB.Stream���޷���ɲ��������ֹ�����</div>"
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
		GBL_CHK_TempStr = "<br>������Ϣ��" & err.description & "<br>�������ܣ�ȷ���Ƿ�Դ��ļ���д��Ȩ��."
		err.Clear
		Set Fs = Nothing
		Exit Sub
	End If

End Sub%>