<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Board_popfun.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<!-- #include file=inc/skin_fun.asp -->

<%
DEF_BBS_HomeUrl = "../../"
Const MaxLinkNum = 100

Dim StyleID,ScreenWidth,DisplayTopicLength,SiteHeadString,SiteBottomString,DefineImage,TableHeadString,TableBottomString,ShowBottomSure
Dim SmallTableHead,SmallTableBottom,TempletID

Sub Main

	initDatabase
	GBL_CHK_TempStr = ""
	checkSupervisorPass
	
	Manage_sitehead DEF_SiteNameString & " - ����Ա",""
	frame_TopInfo
	DisplayUserNavigate("�༭������")
	If GBL_CHK_Flag=1 Then
		Dim ExtentSkinManager,Action
		Action = Request.QueryString("Action")
		If Action = "" Then				
			If dontRequestFormFlag = "" Then
				Action = Request.Form("Action")
			Else
				Action = Form_UpClass.form("Action")
			End If
		End if
		
		If Action <> "" Then
			Set ExtentSkinManager = New ExtentSkin_Manager
			ExtentSkinManager.ExtentSkin
			Set ExtentSkinManager = Nothing
		Else
			GetDefaultValue
			If GBL_CHK_TempStr <> "" Then
				Response.Write "<br><br>" & GBL_CHK_TempStr
			Else
				Siteskin
			End If
		End If
	Else
		DisplayLoginForm
	End If
	frame_BottomInfo
	closeDataBase
	Manage_Sitebottom("none")

End Sub

Main

Function Siteskin

	%>
	<form name="pollform3sdx" method="post" action="DefineStyleParameter.asp">
	<input type="hidden" name="SubmitFlag" value=yes>
	<p>
		<b>
			��̳����������Զ��� - <%=DEF_BoardStyleString(StyleID)%></b>
			<br><br><span class=grayfont>
			���������ʹ��html�﷨�����������ڱ��ص��Ժú��پ����趨<br></span>
	</p>
	<%If Request.Form("SubmitFlag") <> "" Then
		GetFormValue
	End If
	If GBL_CHK_TempStr <> "" Then%>
	<div class=alert><%=GBL_CHK_TempStr%></div>
	<%
	End If
	If Request("SubmitFlag") <> "" Then
		If GBL_CHK_TempStr <> "" Then
			DisplayDatabaseLink
		Else
			SaveStyleDefine
			Response.Write "<div class=alertdone>�ɹ�������ã�</div>"
			Exit Function
		End If
	Else
		DisplayDatabaseLink
	End If
	%>
	<br>
	<input type=submit name=�ύ value=�ύ class=fmbtn>
	<input type=reset name=ȡ�� value=ȡ�� class=fmbtn>
	</form>
	<%

End Function

Function DisplayDatabaseLink

	%>
		<table border=0 cellpadding=0 cellspacing=0 width=100% class=frame_table>
		<tr>
			<td class=tdbox width=120>�����</td>
			<td class=tdbox><b><%=StyleID%></b> (<%=DEF_BoardStyleString(StyleID)%>)<input name=StyleID value=<%=StyleID%> type=hidden></td>
		</tr>
		<!--
		<tr>
			<td class=tdbox>��̳���</td>
			<td class=tdbox><input class=fminpt type="text" name="ScreenWidth" maxlength="50" size="30" value="<%=htmlencode(ScreenWidth)%>"><font color=gray>(֧��ʹ�ðٷֱȺ;��Կ��)</font></td>
		</tr>
		-->
		<tr>
			<td class=tdbox>���ⳤ��</td>
			<td class=tdbox><input class=fminpt type="text" name="DisplayTopicLength" maxlength="3" size="10" value="<%=htmlencode(DisplayTopicLength)%>"><font color=gray>(��λ���ֽڣ�����������ʾ��󳤶ȣ�750��=54��770��=56,�Ϊ255�ֽ�)</font></td>
		</tr>
		<!--
		<tr>
			<td class=tdbox>�Զ�ͼƬ</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=DefineImage value=0<%If DefineImage = 0 Then%> checked<%End If%>></td><td>��</td>
          		<td><input class=fmchkbox type=radio name=DefineImage value=1<%If DefineImage = 1 Then%> checked<%End If%>></td><td>��</td><td><font color=gray>&nbsp; (�Ƿ��Դ��·��ͼƬ�������images/skin/<%=StyleID%>/����ָ����ʹ��Ĭ��ͼƬ)</font></td></tr></table></td>
		</tr>
		-->
		<tr>
			 <td class=tdbox>��վ�ײ�<br>�����Զ�<br>ʹ��HTML</td>
			<td class=tdbox><textarea name=SiteHeadString rows=5 cols=60 class=fmtxtra><%If SiteHeadString <> "" Then Response.Write VbCrLf & Server.htmlEncode(SiteHeadString)%></textarea></td>
		</tr>
		<tr>
			<td class=tdbox>��վβ��<br>�����Զ�<br>ʹ��HTML</td>
			<td class=tdbox><textarea name=SiteBottomString rows=5 cols=60 class=fmtxtra><%If SiteBottomString <> "" Then Response.Write VbCrLf & Server.htmlEncode(SiteBottomString)%></textarea></td>
		</tr>
		<tr>
			<td class=tdbox>��ֱ��<br>ͷ������<br>֧��HTML</td>
			<td class=tdbox><textarea name=TableHeadString rows=5 cols=60 class=fmtxtra><%If TableHeadString <> "" Then Response.Write VbCrLf & Server.htmlEncode(TableHeadString)%></textarea></td>
		</tr>
		<tr>
			<td class=tdbox>��ֱ��<br>β������<br>֧��HTML</td>
			<td class=tdbox><textarea name=TableBottomString rows=5 cols=60 class=fmtxtra><%If TableBottomString <> "" Then Response.Write VbCrLf & Server.htmlEncode(TableBottomString)%></textarea></td>
		</tr>
		<tr>
			<td class=tdbox>С�ֱ��<br>ͷ������<br>֧��HTML</td>
			<td class=tdbox><textarea name=SmallTableHead rows=5 cols=60 class=fmtxtra><%If SmallTableHead <> "" Then Response.Write VbCrLf & Server.htmlEncode(SmallTableHead)%></textarea></td>
		</tr>
		<tr>
			<td class=tdbox>С�ֱ��<br>β������<br>֧��HTML</td>
			<td class=tdbox><textarea name=SmallTableBottom rows=5 cols=60 class=fmtxtra><%If SmallTableBottom <> "" Then Response.Write VbCrLf & Server.htmlEncode(SmallTableBottom)%></textarea></td>
		</tr>
		<tr>
			<td class=tdbox>β������<br>������ʾ</td>
			<td class=tdbox><table border=0 cellpadding=0 cellspacing=0><tr>
				<td><input class=fmchkbox type=radio name=ShowBottomSure value=0<%If ShowBottomSure = 0 Then%> checked<%End If%>></td><td>��ѡ����ʾ</td>
          		<td><input class=fmchkbox type=radio name=ShowBottomSure value=1<%If ShowBottomSure = 1 Then%> checked<%End If%>></td><td>�϶���ʾ</td><td><font color=gray>&nbsp; (Ϊ�����ۣ�ĳЩҳ ��ײ��Զ������ݲ�����ʾ����ѡ����ѡ�񣢣�����ѡ�����)</font></td></tr></table></td>
		</tr>
		<!--
		<tr>
			<td class=tdbox>ʹ��ģ��</td>
			<td class=tdbox><%DisplayTempletList(TempletID)%> <font color=Gray>����ʹ��JSģ�壬��ѡ���һ��</font></td>
		</tr>
		-->
		</table>
	<%

End Function

Function GetDefaultValue

	StyleID = Left(Request("StyleID"),14)
	If isNumeric(StyleID) = 0 Then StyleID = 0
	StyleID = Fix(cCur(StyleID))
	If StyleID < 0 or StyleID > DEF_BoardStyleStringNum Then StyleID = 0
	
	Dim Rs
	Set Rs = LDExeCute(sql_select("Select * from LeadBBS_Skin Where StyleID=" & StyleID,1),0)
	If Not Rs.Eof Then
		ScreenWidth = Rs("ScreenWidth")
		DisplayTopicLength = Rs("DisplayTopicLength")
		SiteHeadString = Rs("SiteHeadString")
		SiteBottomString = Rs("SiteBottomString")
		DefineImage = ccur(Rs("DefineImage"))
		TableHeadString = Rs("TableHeadString")
		TableBottomString = Rs("TableBottomString")
		SmallTableHead = Rs("SmallTableHead")
		SmallTableBottom = Rs("SmallTableBottom")
		ShowBottomSure = Rs("ShowBottomSure")
		If DefineImage = 1 Then
			DefineImage = 1
		Else
			DefineImage = 0
		End If
		If ShowBottomSure >= 1 Then
			ShowBottomSure = 1
		Else
			ShowBottomSure = 0
		End If
		TempletID = cCur(Rs("TempletID"))
		If isNull(TempletID) Then TempletID = 0
	Else
		ScreenWidth = "770"
		DisplayTopicLength = 56
		SiteHeadString = ""
		SiteBottomString = ""
		DefineImage = 0
		TableHeadString = ""
		TableBottomString = ""
		ShowBottomSure = 0
		TempletID = 0
	End If
	Rs.Close
	Set Rs = Nothing

End Function

Function GetFormValue

	ScreenWidth = Request("ScreenWidth")
	DisplayTopicLength = Left(Request("DisplayTopicLength"),4)
	If isNumeric(DisplayTopicLength) = 0 Then DisplayTopicLength = 56
	DisplayTopicLength = Fix(cCur(DisplayTopicLength))
	DefineImage = Request.Form("DefineImage")
	If DefineImage = "1" Then
		DefineImage = 1
	Else
		DefineImage = 0
	End If
	If DisplayTopicLength < 10 or DisplayTopicLength > 255 Then DisplayTopicLength = 56
	SiteHeadString = Left(Request.Form("SiteHeadString"),DEF_MaxTextLength)
	SiteBottomString = Left(Request.Form("SiteBottomString"),DEF_MaxTextLength)
	TableHeadString = Left(Request.Form("TableHeadString"),DEF_MaxTextLength)
	TableBottomString = Left(Request.Form("TableBottomString"),DEF_MaxTextLength)
	SmallTableHead = Left(Request.Form("SmallTableHead"),DEF_MaxTextLength)
	SmallTableBottom = Left(Request.Form("SmallTableBottom"),DEF_MaxTextLength)
	ShowBottomSure = Left(Request.Form("ShowBottomSure"),4)
	TempletID = Left(Request.Form("TempletID"),4)
	
	If ShowBottomSure = "1" Then
		ShowBottomSure = 1
	Else
		ShowBottomSure = 0
	End If

	If isNumeric(TempletID) = 0 Then TempletID = 0

End Function

Function SaveStyleDefine

	Dim Rs
	Set Rs = LDExeCute(sql_select("Select * from LeadBBS_Skin Where StyleID=" & StyleID,1),0)
	If Not Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		rem '"ScreenWidth='" & Replace(ScreenWidth,"'","''") & "'" & _
		rem ",TempletID=" & TempletID &_
		rem ",DefineImage=" & DefineImage & _
		CALL LDExeCute("Update LeadBBS_Skin Set "&_
		"DisplayTopicLength=" & DisplayTopicLength & _
		",SiteHeadString='" & Replace(SiteHeadString,"'","''") & "'" & _
		",SiteBottomString='" & Replace(SiteBottomString,"'","''") & "'" & _
		",TableHeadString='" & Replace(TableHeadString,"'","''") & "'" & _
		",TableBottomString='" & Replace(TableBottomString,"'","''") & "'" & _
		",SmallTableHead='" & Replace(SmallTableHead,"'","''") & "'" & _
		",SmallTableBottom='" & Replace(SmallTableBottom,"'","''") & "'" & _
		",ShowBottomSure=" & ShowBottomSure & _
		" Where StyleID=" & StyleID,1)
	Else
		Rs.Close
		Set Rs = Nothing
		CALL LDExeCute("insert into LeadBBS_Skin(" & _
		"StyleID,ScreenWidth,DisplayTopicLength,SiteHeadString,SiteBottomString,DefineImage," & _
		"TableHeadString,TableBottomString,SmallTableHead,SmallTableBottom,ShowBottomSure,TempletID) values(" & _
		StyleID & _
		",'" & Replace(ScreenWidth,"'","''") & "'" & _
		"," & DisplayTopicLength & _
		",'" & Replace(SiteHeadString,"'","''") & "'" & _
		",'" & Replace(SiteBottomString,"'","''") & "'" & _
		"," & DefineImage & _
		",'" & Replace(TableHeadString,"'","''") & "'" & _
		",'" & Replace(TableBottomString,"'","''") & "'" & _
		",'" & Replace(SmallTableHead,"'","''") & "'" & _
		",'" & Replace(SmallTableBottom,"'","''") & "'" & _
		"," & ShowBottomSure & _
		"," & TempletID & _
		")",1)
	End If
	ReloadBoardStyleInfo(StyleID)

End Function

Sub DisplayTempletList(TempletID)

	%>
	<script language=javascript>
	var TempletID = <%=TempletID%>;
	function s(ID,TempletName)
	{
		if(ID=="")return;
		if(TempletID == parseInt(ID)){document.write("<option value=" + ID + " selected>" + TempletName);}
		else{document.write("<option value=" + ID + ">" + TempletName);}
	}
	</script>
	<%
	Dim Rs,SQL
	SQL = "select ID,TempletName from LeadBBS_Templet"

	Set Rs = LDExeCute(SQL,0)
	Dim Num
	If Not rs.Eof Then
		Response.Write "<select name=TempletID><script language=javascript>s(""9999"",""HTML���(��JSģ��)"");" & VbCrLf & "s("""
		Response.Write Rs.GetString(,,""",""",""");" & VbCrLf & "s(""","")
		%>","","","");
		</script>
		</select>
		<%
	Else
		Num = -1
		Response.Write "�޿���ģ��"
	End If
	Rs.close
	Set Rs = Nothing

End Sub
%>