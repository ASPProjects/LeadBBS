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
DisplayUserNavigate("��վ��������")
If GBL_CHK_Flag=1 Then
	LoginAccuessFul
Else
	DisplayLoginForm
End If
frame_BottomInfo
closeDataBase
Manage_Sitebottom("none")

Function LoginAccuessFul

	If Request.Form("submitflag")="Dieos9xsl29LO_8" Then
		Application.Lock
		If Request("Flag2") <> "" Then
			Application.Contents.RemoveAll()
			Response.Write "<div class=frameline>�ɹ������ͷ���̳������</div>"
		Else
			FreeApplicationMemory
			Response.Write "<div class=frameline>�ɹ������̳�������ã�</div>"
		End If
		If Request("Flag") <> "" then
			application(DEF_MasterCookies & "SiteEnableFlagzoieiu") = 0
			application(DEF_MasterCookies & "SiteDisbleWhyszoieiu") = "��̳�Ѿ��ر�"
			Response.Write "<div class=frameline>��̳�����ɹ��������Ѿ������رգ��������������װ��̳���벻Ҫ�ٷ����κι���Աҳ�森</div>"
		Else
			Response.Write "<div class=frameline>��̳�����ɹ���</div>"
		End If
		Application.UnLock
	Else
		DisplayStringForm
	End If

End Function

Function DisplayStringForm

%>
<div class=frameline>
��ȷ���Ƿ���Ҫ����������̳����������̳һ��״̬������<br>
���൱��Web������������Ľ����<br>
�����������������㣬���ͷ�һЩ��̳ռ�õ��ڴ森
</div>
<form action=SiteReset.asp method="post">
	<div class=alert>ȷ����Ϣ�� �������������̳ô��</div>
	<div class=frameline>
	<input class=fmchkbox type="checkbox" name=Flag value="yes">ѡ�������������Զ��ر���վ����,�������Ҫ�����ķ���װ�����̳,���ڴ˲�����,��Ҫ�ٽ��������Ĺ���Ա����,�Ա��������ڴ泹���ͷ�.
	</div>
	<div class=frameline>
	<input class=fmchkbox type="checkbox" name=Flag2 value="yes">ѡ���򳹵��ͷ��ڴ�ռ��(���ַ�������֧��)
	</div>
	<div class=frameline>
	<input name=submitflag type=hidden value="Dieos9xsl29LO_8">
	<input type=submit value="������̳" class="fmbtn"> <input type=reset value="ȡ��" class="fmbtn">
	</div>
</form>
<div class=frametitle><b>������������̳���ݽ����ͷţ�</b></div>
<div class=frameline>
<table>
	<%
	Dim Thing
	For Each Thing in Application.Contents
		If Left(Thing,Len(DEF_MasterCookies)) = DEF_MasterCookies Then
			Response.Write "<tr><td><font color=Gray class=grayfont>" & thing & "</font></td><td>&nbsp;"
			If isObject(Application.Contents(Thing)) Then
				Response.Write "����"
			ElseIf isArray(Application.Contents(Thing)) Then
				Response.Write "����"
			Else
				Response.Write Application.Contents(Thing)
			End If
			Response.Write "</td></tr>"
		End If
	Next
	Response.Write "</table></div>"

End Function

Function FreeApplicationMemory

	Response.Write "<div class=frametitle>�ͷ���̳�����б�</div><div class=frameline><table>" & VbCrLf
	Dim Thing
	For Each Thing in Application.Contents
		If Left(Thing,Len(DEF_MasterCookies)) = DEF_MasterCookies Then
			Response.Write "<tr><td><font color=Gray class=grayfont>" & thing & "</font></td><td>&nbsp;"
			If isObject(Application.Contents(Thing)) Then
				Application.Contents(Thing).close
				Set Application.Contents(Thing) = Nothing
				Application.Contents(Thing) = null
				Application.Contents.Remove(Thing)
				Response.Write "����ɹ��ر�"
			ElseIf isArray(Application.Contents(Thing)) Then
				Set Application.Contents(Thing) = Nothing
				Application.Contents(Thing) = null
				Application.Contents.Remove(Thing)
				Response.Write "����ɹ��ͷ�"
			Else
				Response.Write htmlencode(Application.Contents(Thing))
				Application.Contents(Thing) = null
				Application.Contents.Remove(Thing)
			End If
			Response.Write "</td></tr>"
		End If
	Next
	Response.Write "</table></div>"

End Function%>