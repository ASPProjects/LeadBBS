<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Board_Popfun.asp -->
<!--#include file="inc/StarSetup.asp"-->
<%
DEF_BBS_HomeUrl = "../../"

Main

Sub Main

	BBS_SiteHead DEF_SiteNameString & " - ��������",0,"<span class=""navigate_string_step"">��������</span>"
	Dim Master
	InitDatabase
	If CheckSupervisorUserName = 1 Then
		Master = True
	Else
		Master = False
	End If

	Boards_Body_Head("")
	%>
	<div class="alertbox fire">
	<table cellpadding="0" cellspacing="0" class="table_in">
	<tr class="tbinhead">
		<td><div class="value"><b>���ٷ���LeadBBS��ҳ���ǲ�� �������ǹ������ģ�</b></div></td>
	</tr>
	<tr>
	<td class="tdbox">
	<%
	If GBL_CHK_User="" or Not(Master) Then
		%>
		<div class="alert">���������ԭ������ǣ�</div><br /><br />
		�㲻�ǹ���Ա����Ȩ���룡<
		br />������ǹ���Ա�����Թ���Ա���<a href="<%=DEF_BBS_HomeUrl%>User/Login.asp?Relogin=Yes&u=<%=urlencode(Request.Servervariables("SCRIPT_NAME") & "?" & Request.QueryString)%>"><b>�ص�¼</b></a>��
		</div>
		<%
	Else
		Call Main_Star()
	End If
	CloseDatabase%>
	</td><tr>
	</table>
	<%
	Boards_Body_Bottom
	
	If GBL_ShowBottomSure = 0 Then GBL_SiteBottomString = ""
	SiteBottom

End Sub

Sub Main_Star()

	%>
		<br />
		<ol>
			<li>
			ע����� �����棬���������Ե�ǰ����������ʾ��ʽ���趨�������Ҫ�޸ģ�������Ӧ�ĵ�ѡ��ť��ÿһ�е���ʾ��ʽ�����Զ��壬�Զ���������ǿ��
			</li>
			<li>�����㣬�Ӷ���ҳ��ʾ���ٶ�Ӱ�쿴����Χ�ڣ�0.00�롪��2.0��䡣ÿһ����ʾ��ʽ���ٶȵ�Ӱ�춼��һ��������Ա��ʹ��ʱ���Լ�ʵ��һ�£�����
			</li>
			<li>
			<a href="http://www.leadbbs.com/a/a.asp?B=10&ID=858130" target="_blank">
				<span class="redfont">�ٷ���LeadBBS��ҳ���ǲ�����°����¸�����������</span></a>
			</li>
		</ol>
				<form action="admin_HomePageStar.asp" method="post">
					<%
	Dim Temp_1,Temp_2,Temp_3,Temp_4,Temp_5,Temp_6
	If Request.Form("submit") <> " �ύ " Then
		Temp_1 = GBL_PLUG_HPS_LineFirstType
		Temp_2 = GBL_PLUG_HPS_LineSecondType
		Temp_3 = GBL_PLUG_HPS_ShowType
		Temp_4 = GBL_PLUG_HPS_RefreshSpace
		Temp_5 = GBL_PLUG_HPS_TopMax
		Temp_6 = GBL_PLUG_HPS_Collapse
		%>
		<ol>
		<li><b>��һ��������ʾ��ʽ��</b>
			<br />
			<input type="radio" name="GBL_PLUG_HPS_LineFirstType" value="1" <%if Temp_1="1" Then Response.Write("checked=""checked""") End If%> class="fmchkbox">ÿ�շ�����
			<input type="radio" name="GBL_PLUG_HPS_LineFirstType" value="2" <%if Temp_1="2" Then Response.Write("checked=""checked""") End If%> class="fmchkbox">ÿ�ܷ�����
			<input type="radio" name="GBL_PLUG_HPS_LineFirstType" value="3" <%if Temp_1="3" Then Response.Write("checked=""checked""") End If%> class="fmchkbox">ÿ�·�����
			<input type="radio" name="GBL_PLUG_HPS_LineFirstType" value="4" <%if Temp_1="4" Then Response.Write("checked=""checked""") End If%> class="fmchkbox">ÿ�귢���� 
			<input type="radio" name="GBL_PLUG_HPS_LineFirstType" value="5" <%if Temp_1="5" Then Response.Write("checked=""checked""") End If%> class="fmchkbox">��������
		</li>
		<li><b>�ڶ���������ʾ��ʽ��</b>
			<br />
			<input type="radio" name="GBL_PLUG_HPS_LineSecondType" value="1" <%if Temp_2="1" Then Response.Write("checked=""checked""") End If%> class="fmchkbox">ÿ�շ�����
			<input type="radio" name="GBL_PLUG_HPS_LineSecondType" value="2" <%if Temp_2="2" Then Response.Write("checked=""checked""") End If%> class="fmchkbox">ÿ�ܷ�����
			<input type="radio" name="GBL_PLUG_HPS_LineSecondType" value="3" <%if Temp_2="3" Then Response.Write("checked=""checked""") End If%> class="fmchkbox">ÿ�·�����
			<input type="radio" name="GBL_PLUG_HPS_LineSecondType" value="4" <%if Temp_2="4" Then Response.Write("checked=""checked""") End If%> class="fmchkbox">ÿ�귢����
			<input type="radio" name="GBL_PLUG_HPS_LineSecondType" value="5" <%if Temp_2="5" Then Response.Write("checked=""checked""") End If%> class="fmchkbox">��������
		</li>
		<li><b>���ǲ����ʾ��ʽ��</b>
			<br />
			<input type="radio" name="GBL_PLUG_HPS_ShowType" value="0" <%if Temp_3="0" Then Response.Write("checked=""checked""") End If%> class="fmchkbox">��ֹ��ʾ
			<input type="radio" name="GBL_PLUG_HPS_ShowType" value="1" <%if Temp_3="1" Then Response.Write("checked=""checked""") End If%> class="fmchkbox">������ʾ����
			<input type="radio" name="GBL_PLUG_HPS_ShowType" value="2" <%if Temp_3="2" Then Response.Write("checked=""checked""") End If%> class="fmchkbox">ֻ��ʾ��һ��
			<input type="radio" name="GBL_PLUG_HPS_ShowType" value="3" <%if Temp_3="3" Then Response.Write("checked=""checked""") End If%> class="fmchkbox">ֻ��ʾ�ڶ���
		</li>
		<li>4. <b>���ǲ��ˢ�¼����</b>
			<br />
			<input type="text" size="3" maxlength="3" name="GBL_PLUG_HPS_RefreshSpace" value="<%=Temp_4%>" class="fminpt input_1">���ӣ�Ϊ�������վ�����ټ����������������棬Ҫ������Ϊ5���ӣ����50����
		</li>
		<li><b>��ʾ���Ǽ�¼������</b>
			<br />
			<input type="text" size="2" maxlength="2" name="GBL_PLUG_HPS_TopMax" value="<%=Temp_5%>" class="fminpt input_1"> ����Ϊ3�������ֻ������Ϊ50��
		</li>
		<li><b>Ĭ���Ƿ����</b>
			<br />
			<input type="radio" name="GBL_PLUG_HPS_Collapse" value="0" <%if Temp_6="0" Then Response.Write("checked=""checked""") End If%> class="fmchkbox">Ĭ����ʾ
			<input type="radio" name="GBL_PLUG_HPS_Collapse" value="1" <%if Temp_6="1" Then Response.Write("checked=""checked""") End If%> class="fmchkbox">Ĭ�Ͼ���
		</li>
		</ol>
		<ol>
		<div class=value2>
			<input type="submit" name="Submit" value=" �ύ " class="fmbtn btn_2">
			<input type="reset" name="Submit2" value=" ���� " class="fmbtn btn_2">
		</div>
		</ol>
		<%
	Else
		Temp_1 = Left(Trim(Request.Form("GBL_PLUG_HPS_LineFirstType")),14)
		Temp_2 = Left(Trim(Request.Form("GBL_PLUG_HPS_LineSecondType")),14)
		Temp_3 = Left(Trim(Request.Form("GBL_PLUG_HPS_ShowType")),14)
		Temp_4 = Left(Trim(Request.Form("GBL_PLUG_HPS_RefreshSpace")),14)
		Temp_5 = Left(Trim(Request.Form("GBL_PLUG_HPS_TopMax")),14)
		Temp_6 = Left(Trim(Request.Form("GBL_PLUG_HPS_Collapse")),14)
		If isNumeric(Temp_1) = 0 then Temp_1 = 0
		If isNumeric(Temp_2) = 0 then Temp_2 = 0
		If isNumeric(Temp_3) = 0 then Temp_3 = 0
		If isNumeric(Temp_4) = 0 then Temp_4 = 5
		If isNumeric(Temp_6) = 0 then Temp_6 = 0
		Temp_4 = Fix(cCur(Temp_4))
		If Temp_4 < 5 Then Temp_4 = 5
		If Temp_4 > 50 Then Temp_4 = 50
		
		If isNumeric(Temp_5) = 0 then Temp_5 = 5
		Temp_5 = Fix(cCur(Temp_5))
		If Temp_5 < 3 Then Temp_5 = 3
		If Temp_5 > 50 Then Temp_5 = 50
		
		Dim WriteString
		WriteString = ""
		WriteString = WriteString & "<%" & VbCrLf
		WriteString = WriteString & "const GBL_PLUG_HPS_LineFirstType = " & Temp_1 & VbCrLf
		WriteString = WriteString & "const GBL_PLUG_HPS_LineSecondType = " & Temp_2 & VbCrLf
		WriteString = WriteString & "const GBL_PLUG_HPS_ShowType = " & Temp_3 & VbCrLf
		WriteString = WriteString & "const GBL_PLUG_HPS_RefreshSpace = " & Temp_4 & VbCrLf
		WriteString = WriteString & "const GBL_PLUG_HPS_TopMax = " & Temp_5 & VbCrLf
		WriteString = WriteString & "const GBL_PLUG_HPS_Collapse = " & Temp_6 & VbCrLf & VbCrLf
		WriteString = WriteString & "'####################################################################" & VbCrLf
		WriteString = WriteString & "'##" & VbCrLf
		WriteString = WriteString & "'##��������������ʾ��ʽ����!����ķ�������֧��FSO�����ֶ��޸���ʾ��ʽ!" & VbCrLf
		WriteString = WriteString & "'##����1Ϊÿ�շ�������2Ϊÿ�ܷ�������3Ϊÿ�·�������4Ϊÿ�귢������5Ϊ����������6Ϊ��������ǣ�7Ϊ���Ů����" & VbCrLf
				WriteString = WriteString & "'##�����˳�������ɹ�����Ա������£�ʱ��2004-03-20 16:50 LeadBBS�������� for 3.14" & VbCrLf
		WriteString = WriteString & "'##�������������ֹ����������λ��!" & VbCrLf
		WriteString = WriteString & "'##����ʹ��ǰ���ȿ���װ˵��!" & VbCrLf
		WriteString = WriteString & "'##������ҳ��http://gafc.9126.com/" & VbCrLf
		WriteString = WriteString & "'##�����ٷ���ҳ��http://www.LeadBBS.com/" & VbCrLf
		WriteString = WriteString & "'##������л��ʹ�ñ����!" & VbCrLf
		WriteString = WriteString & "'##" & VbCrLf
		WriteString = WriteString & "'####################################################################" & VbCrLf

		WriteString = WriteString & "%" & chr(62) & VbCrLf
		ADODB_SaveToFile WriteString,"Inc/StarSetup.asp"
		
		If GBL_CHK_TempStr = "" Then
			Response.Write "<br /><span class=greenfont>2.�ɹ�������ã�</span>"
			Response.Write("��ϲ��������������ʾ��ʽ�Ѿ��趨��ϣ�����"&"<br /><br /><br />")
        	Response.Write("<input type=""button"" value=""��������"" onclick=""window.location.href='admin_HomePageStar.asp'"" class=""fmbtn btn_3"">")
		Else
			%><%=GBL_CHK_TempStr%><br />��������֧������д���ļ����ܣ���ʹ��FTP�ȹ��ܣ���<span class="redfont">Inc/StarSetup.asp</span>�ļ��滻�ɿ�������(ע�ⱸ��)
			<p>
			<textarea name="fileContent" cols="80" rows="30" class="fmtxtra"><%=Server.htmlencode(WriteString)%></textarea>
			</p><%
			GBL_CHK_TempStr = ""
		End If
	End If
	%>
					</form>
<%
	Set Application(DEF_MasterCookies & "_PLUG_HPS_DAY") = Nothing
	Application(DEF_MasterCookies & "_PLUG_HPS_DAY") = ""
	Set Application(DEF_MasterCookies & "_PLUG_HPS_OTHER") = Nothing
	Application(DEF_MasterCookies & "_PLUG_HPS_OTHER") = ""
	Application(DEF_MasterCookies & "_PLUG_HPS_M") = ""

End Sub
%>