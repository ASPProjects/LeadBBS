<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Board_popfun.asp -->
<!-- #include file=../../inc/Limit_Fun.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<!-- #include file=../../User/inc/Fun_SendMessage.asp -->
<%
Server.ScriptTimeOut = 99999
DEF_BBS_HomeUrl = "../../"
initDatabase
GBL_CHK_TempStr = ""
checkSupervisorPass

Manage_sitehead DEF_SiteNameString & " - ����Ա",""


Dim Sdm_FromUser,Sdm_ToUser,Sdm_Title,Sdm_Content,Smd_ToUserID,SdM_ToUserClass
Sdm_FromUser = GBL_CHK_User

SdM_ToUserClass = 0
frame_TopInfo
DisplayUserNavigate("��̳����ϢȺ��")%>
<p>
<%
If GBL_CHK_Flag=1 Then
	If GBL_CHK_TempStr="" Then
		If Request.Form("submitFlag")<>"" Then
			CheckSubmitFormData
			If GBL_CHK_TempStr = "" Then
				WriteNewMessageToDatabase
			Else
				Response.Write "<div class=alert>" & GBL_CHK_TempStr & "</div>" & VbCrLf
				NewMessageForm
			End If
		Else
			NewMessageForm
		End If
	Else
		Response.Write GBL_CHK_TempStr
	End If
Else
DisplayLoginForm
End If
frame_BottomInfo
closeDataBase
Manage_Sitebottom("none")

Function NewMessageForm

	Dim TempN
%>
	<div class=frametitle>Ⱥ�����µĶ���Ϣ</b>���ӡ�<span class=redfont><b>*</b></span>����Ϊ�����
	</div>
	<br>
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
	function ctlkey()
	{
		if(event.ctrlKey && event.keyCode==13){submitonce(document.frmReg);document.frmReg.submit();}
		if(event.altKey && (event.keyCode==83 || event.keyCode==115)){submitonce(document.frmReg);document.frmReg.submit();}
	}
	var ie = (document.all)? true:false
	if (ie)
	{
		window.document.onkeydown=ctlkey;
	}
	function smilie(smilietext)
	{
		smilietext=smilietext;
		if (document.frmReg.SdM_Content.createTextRange && document.frmReg.SdM_Content.caretPos)
		{
			var caretPos = document.frmReg.SdM_Content.caretPos;
			caretPos.text = caretPos.text.charAt(caretPos.text.length - 1) == ' ' ? smilietext + ' ' : smilietext;document.frmReg.SdM_Content.focus();
		}
		else
		{
			document.frmReg.SdM_Content.value+=smilietext;document.frmReg.SdM_Content.focus();
		}
	}
	function storeCaret (textEl)
	{
		if (textEl.createTextRange) 
		textEl.caretPos = document.selection.createRange().duplicate(); 
	}
	</script>
	<form method="post" action="SendGroupMessage.asp" id="frmReg" name="frmReg" onSubmit="submitonce(this);">
	<table border=0 cellpadding=0 cellspacing=0 width="100%" class=frame_table>
	
	<tr> 
		<td class=tdbox width=120>
			������</td>
		<td class=tdbox>
			<%=Sdm_FromUser%>
			<input name=submitFlag value="<%=Second(time)&minute(time)%>" type=hidden>
		</td>
	</tr>
	<tr>
		<td class=tdbox>
			<%If CheckSupervisorUserName = 0 Then%><font color="#CC0000" class=redfont><b>*</b></font><%End If%>������</td>
		<td class=tdbox>
			<select name=SdM_ToUserClass>
				<option value=0<%If SdM_ToUserClass = 0 Then Response.Write " selected"%>>ȫ�����</option>
				<option value=1<%If SdM_ToUserClass = 1 Then Response.Write " selected"%>>ȫ��<%=DEF_PointsName(6)%></option>
				<option value=2<%If SdM_ToUserClass = 2 Then Response.Write " selected"%>>ȫ��<%=DEF_PointsName(5)%></option>
				<option value=4<%If SdM_ToUserClass = 4 Then Response.Write " selected"%>>ȫ��<%=DEF_PointsName(10)%></option>
				<option value=5<%If SdM_ToUserClass = 5 Then Response.Write " selected"%>>��������������õ������û�(����ʹ��)</option>
				<option value=3<%If SdM_ToUserClass = 3 Then Response.Write " selected"%>>ȫ���û���ռ����Դ�󣬽���ʹ��</option>
		</td>
	</tr>
	<tr> 
		<td class=tdbox align=left width=20%><font color="#CC0000" class=redfont><b>*</b></font>��Ϣ����</td>
		<td class=tdbox height=24 valign=top width=80%>
			<input class=fminpt name=SdM_Title value="<%=htmlencode(SdM_Title)%>" size=60 maxlength=50>
		</td>
	</tr>
	<tr> 
		<td class=tdbox align=left width=20%>����Ϣ����<br><br><a href="javascript:smilie('[IMG][/IMG]');">֧��[IMG]��<br>ǩ����ͼƬ</td>
		<td class=tdbox height=24 valign=top width=80%>
			<textarea cols=58 name="SdM_Content" rows=10" onselect="storeCaret(this);" onclick="storeCaret(this);" onkeyup="storeCaret(this);" class=fmtxtra><%If SdM_Content<>"" Then Response.Write VbCrLf & Htmlencode(SdM_Content)%></textarea>
		</td>
	</tr>
	</table>
	<br>
	<table width=95% border="0" cellspacing="2" cellpadding="2">
	<tr>
		<td class=tdbox width=20%>&nbsp;&nbsp;</td>
		<td class=tdbox>
			<input type="submit" name="Submit" value="�ύ" class=fmbtn> &nbsp;
			<input type="reset" name="reset" value="���" class=fmbtn>
		</td>
	</tr>
	</table>
	</form>
	<p>ע�����ж���Ϣ����ϵͳ��ݷ���</p>
<%
End Function

Function CheckSubmitFormData

	SdM_ToUserClass = Trim(Request.Form("SdM_ToUserClass"))
	SdM_Title = Trim(Request.Form("SdM_Title"))
	SdM_ConTent = Request.Form("SdM_ConTent")

	If SdM_ToUserClass <> "0" and SdM_ToUserClass <> "1" and SdM_ToUserClass <> "2" and SdM_ToUserClass <> "3" and SdM_ToUserClass <> "4" and SdM_ToUserClass <> "5" Then SdM_ToUserClass = -1
	SdM_ToUserClass = cCur(SdM_ToUserClass)
	If SdM_ToUserClass = -1 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "���󣬽��ն���ѡ�����!<br>" & VbCrLf
		Exit Function
	End If

	If len(SdM_Title) > 50 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "������Ϣ�����벻Ҫ����50����. <br>" & VbCrLf
		Exit Function
	End if

	If SdM_Title = "" Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "������Ϣ���������д. <br>" & VbCrLf
		Exit Function
	End if

	If Len(SdM_Content) > DEF_MaxTextLength then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "�������ݲ��ܳ���" & DEF_MaxTextLength & "����!<br>" & VbCrLf
		Exit Function
	End If

End Function


Function WriteNewMessageToDatabase

	Dim Rs,Flag,N,GetData
	Select Case SdM_toUserClass
	Case 0:
		Set Rs = LDExeCute("select UserID,UserName from LeadBBS_SpecialUser Where Assort=1 order by UserID",0)
		If Not Rs.Eof Then
			GetData = Rs.GetRows(-1)
			Rs.Close
			Set Rs = Nothing
			Rs = Ubound(GetData,2)
			Flag = 0
			For N = 0 to Rs
				If Flag <> cCur(GetData(0,N)) Then
					SendMsg(GetData(1,N))
					Flag = cCur(GetData(0,N))
				End If
			Next
		Else
			Rs.Close
			Set Rs = Nothing
			Response.Write "<p>û���κΰ�����Ⱥ��������<br>" & VbCrLf
		End If
	Case 1:
		Set Rs = LDExeCute("select UserName from LeadBBS_SpecialUser Where Assort=2",0)
		If Rs.Eof Then
			Rs.Close
			Set Rs = Nothing
		Else
			GetData = Rs.GetRows(-1)
			Rs.Close
			Set Rs = Nothing
			Rs = Ubound(GetData,2)
			For N = 0 to Rs
				SendMsg(GetData(0,N))
			Next
		End If
	Case 2:
		Set Rs = LDExeCute("select UserName from LeadBBS_SpecialUser Where Assort=0",0)
		If Rs.Eof Then
			Rs.Close
			Set Rs = Nothing
		Else
			GetData = Rs.GetRows(-1)
			Rs.Close
			Set Rs = Nothing
			Rs = Ubound(GetData,2)
			For N = 0 to Rs
				SendMsg(GetData(0,N))
			Next
		End If
	Case 3:
		Set Rs = LDExeCute("select UserName from LeadBBS_User",0)
		Response.Clear
		N = 0
		Response.Write "<p style=font-size:9pt>"
		Do While Not Rs.Eof
			N = N + 1
			SendMsg(Rs(0))
			If (N mod 100) = 0 Then Response.flush
			Rs.MoveNext
		Loop
		Rs.Close
		Set Rs = Nothing
	Case 4:
		Set Rs = LDExeCute("select UserName from LeadBBS_SpecialUser Where Assort=8",0)
		If Rs.Eof Then
			Rs.Close
			Set Rs = Nothing
		Else
			GetData = Rs.GetRows(-1)
			Rs.Close
			Set Rs = Nothing
			Rs = Ubound(GetData,2)
			For N = 0 to Rs
				SendMsg(GetData(0,N))
			Next
		End If
	Case 5:
		Set Rs = LDExeCute("select UserName from LeadBBS_User where LastDoingTime>=" & GetTimeValue(Dateadd("m",-2,DEF_Now)),0)
		Response.Clear
		N = 0
		Response.Write "<p style=font-size:9pt>"
		Do While Not Rs.Eof
			N = N + 1
			SendMsg(Rs(0))
			If (N mod 100) = 0 Then Response.flush
			Rs.MoveNext
		Loop
		Rs.Close
		Set Rs = Nothing
	End Select

End Function

Sub SendMsg(SdM_toUser)

	If CheckUserNameExist(SdM_ToUser) = 0 Then
		Response.Write "<br><span class=redfont>�û� " & htmlencode(SdM_toUser) & " �����ڣ����Ͷ���Ϣʧ�ܣ�</font>"
		Exit Sub
	End If

	SdM_fromUser = "[LeadBBS]"
	SendNewMessage SdM_fromUser,SdM_ToUser,SdM_Title,SdM_Content,GBL_IPAddress

	Response.Write GBL_CHK_TempStr
	If Smd_ToUserID = GBL_UserID Then UpdateSessionValue 6,1,0

End Sub

Rem ���ĳ�û���ID�Ƿ����
Function CheckUserNameExist(UserName)

	Dim Rs
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	Set Rs = LDExeCute(sql_select("Select ID,UserName from LeadBBS_User where UserName='" & Replace(UserName,"'","''") & "'",1),0)
	If Rs.Eof Then
		CheckUserNameExist = 0
		Smd_ToUserID = 0
	Else
		CheckUserNameExist = 1
		Smd_ToUserID = cCur(Rs(0))
		Sdm_ToUser = Rs(1)
	End if
	Rs.Close
	Set Rs = Nothing

End Function
%>