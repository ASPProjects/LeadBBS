<!-- #include file=../inc/BBSsetup.asp -->
<!-- #include file=../inc/Board_popfun.asp -->
<!-- #include file=../inc/Limit_fun.asp -->
<!-- #include file=inc/UserTopic.asp -->
<!-- #include file=inc/Fun_SendMessage.asp -->
<%
Const DEF_User_MaxReceiveUser = 5 '��������ͬʱ���Ͷ���Ϣ�����ٸ��û���Ĭ��ֵΪ5
DEF_BBS_HomeUrl = "../"

Dim Sdm_FromUser,Sdm_ToUser,Sdm_Title,Sdm_Content,Sdm_ToUserID
Dim ReplyMessageID,ModifyMessageID

Dim AjaxFlag

Main

Sub Main

	If Request("AjaxFlag") = "1" Then
		AjaxFlag = 1
	Else
		AjaxFlag = 0
	End If
	If Request.QueryString("go") = "liling" Then
		Main_SelectFriend
	Else
		Main_SendMessage
	End If

End Sub

Sub Main_SendMessage

	initDatabase
	CheckisBoardMaster
	GBL_CHK_TempStr = ""

	Sdm_FromUser = GBL_CHK_User

	If GBL_UserID < 0 Then GBL_UserID = 0
	If GBL_UserID = 0 or Sdm_FromUser = "" Then GBL_CHK_TempStr = GBL_CHK_TempStr & "��δ��¼" & VbCrLf

	ReplyMessageID = Left(Request("ReplyMessageID"),14)
	If isNumeric(ReplyMessageID) = 0 or ReplyMessageID = "" or InStr(ReplyMessageID,",") Then ReplyMessageID = 0
	ReplyMessageID = Fix(cCur(ReplyMessageID))
	If ReplyMessageID < 0 Then ReplyMessageID = 0

	ModifyMessageID = Left(Request("ModifyMessageID"),14)
	If isNumeric(ModifyMessageID) = 0 or ModifyMessageID = "" or InStr(ModifyMessageID,",") Then ModifyMessageID = 0
	ModifyMessageID = Fix(cCur(ModifyMessageID))
	If ModifyMessageID < 0 Then ModifyMessageID = 0

	SdM_ToUser = Request("SdM_ToUser")
	
	Dim submitFlag
	If Request.Form("submitFlag")<>"" Then
		submitFlag = 1
	Else
		submitFlag = 0
	End If

	If AjaxFlag = 0 Then
		BBS_SiteHead DEF_SiteNameString & " - д����Ϣ",0,"<span class=navigate_string_step>д����Ϣ</span>"
		UpdateOnlineUserAtInfo GBL_board_ID,"���Ͷ���Ϣ"
	
		UserTopicTopInfo("user")
	ElseIf submitFlag = 0 Then%>
	<div class="ajaxbox">
	<div id="sndErrString"></div>
	<%
	End If

	If GBL_CHK_Flag = 1 Then
		CheckUserAnnounceLimit
		If GBL_CHK_OnlineTime < DEF_NeedOnlineTime and DEF_NeedOnlineTime > 0 and CheckSupervisorUserName = 0 Then
			GBL_CHK_TempStr = "<div class=alert>��̳��������ʱ��(" & DEF_PointsName(4) & ")" & Fix(DEF_NeedOnlineTime/60) & "�������û�����ʹ�ô˹��ܡ�</div>" & VbCrLf
		End If
		If ModifyMessageID > 0 Then GetMessageValue(ModifyMessageID)
		If GBL_CHK_TempStr = "" Then
			If submitFlag <> 0 Then
				CheckSubmitFormData
				If GBL_CHK_TempStr = "" Then
					WriteNewMessageToDatabase
				Else
					Message_Done GBL_CHK_TempStr,"err"
					If AjaxFlag = 0 Then NewMessageForm
				End If
			Else
				If ReplyMessageID > 0 and ModifyMessageID = 0 Then GetMessageValue(ReplyMessageID)
				NewMessageForm
			End If
		Else
			Message_Done GBL_CHK_TempStr,"err"
		End If
	Else
		If AjaxFlag = 0 Then
			If Request("submitflag")="" Then
				Response.Write DisplayLoginForm("���ȵ�¼")
			Else
				DisplayLoginForm(GBL_CHK_TempStr)
			End If
		Else
			Message_Done "�˺���֤����.","err"
		End If
	End If
	closeDataBase
	If AjaxFlag = 0 Then
		UserTopicBottomInfo
		SiteBottom
	ElseIf submitFlag = 0 Then%>
	</div>
	<%
	End If

End Sub

Sub Message_Done(Str,tp)

	If AjaxFlag = 1 and Request.Form("submitFlag")<>"" Then
		If tp = "err" Then%>
		<script>parent.document.getElementById("sndErrString").innerHTML = "<div class=alert><%=Replace(Replace(Replace(Str,"\","\\"),"""","\"""),VbCrLf,"")%></div>";
		parent.submit_disable(parent.document.getElementById("LeadBBSFm"),1)
		</script>
		<%Else%>
	<script>parent.layer_outmsg("anc_delbody","<div class=\"ajaxbox\"><%=Replace(Replace(Replace(Str,"\","\\"),"""","\"""),VbCrLf,"")%></div>","");</script>
	<%
		End If
	Else
		If tp = "err" Then
			Response.Write "<div class=alert>" & Str & "</div>"
		Else
			Response.Write "<div class=alertdone>" & Str & "</div>"
		End If
	End If

End Sub

Function NewMessageForm

	Dim TempN,Pub,SuperFlag
	
	SuperFlag = CheckSupervisorUserName
	If Request("pub") <> "" Then
		Pub = 1
	Else
		Pub = 0
	End If

	Response.Write "<div class=title>"
	If ModifyMessageID > 0 Then
		Response.Write "�༭"
	Else
		Response.Write "��д�µ�"
	End If
	If SdM_ToUser = "" and Pub = 1 Then
		Response.Write "����"
	Else
		Response.Write "����Ϣ"
	End If
	
	Dim Url
	Url = filterUrlstr(Left(Request("dir"),100))
	If Request("dir") = "" Then
		Url = DEF_BBS_HomeUrl
	End If
	%></div>
	<form method="post" action="<%=Url%>User/SendMessage.asp" id="LeadBBSFm" name="LeadBBSFm"<%
	If AjaxFlag = 1 Then
		Response.Write " target=""hidden_frame"""%> onSubmit="sutmitflag=1;submit_disable(this);return true;"<%
	Else%> onSubmit="return(submitonce(this));"<%
	End If
	%>>
	<input type=hidden name=AjaxFlag value="<%=AjaxFlag%>">
	<table border=0 cellpadding=0 cellspacing=0 width=95% class=table_in>
	
	<tr> 
		<td width=150 class=tdbox>
			������</td>
		<td valign=top class=tdbox>
			<%=Sdm_FromUser%>
			<input name=submitFlag value="<%=Second(time)&minute(time)%>" type=hidden>
			<%If SuperFlag = 1 Then%><input name=pub value="<%If Pub = 1 Then Response.Write "1"%>" type=hidden><%End If%>
			<input name=ModifyMessageID value="<%=ModifyMessageID%>" type=hidden>
		</td>
	</tr>
	<%If DEF_EnableAttestNumber > 2 and (DEF_AttestNumberPoints = 0 or GBL_CHK_Points < DEF_AttestNumberPoints) Then%>
	<tr> 
		<td class=tdbox>��֤��</td>
		<TD class=tdbox>
			<%displayVerifycode%>
		</td>
	</tr>
				<%End If
	If SuperFlag = 0 or Pub = 0 Then%>
	<tr>
		<td class=tdbox>
			<%If SuperFlag = 0 Then%>*<%End If%>������</td>
		<td valign=top class=tdbox><%If ModifyMessageID > 0 Then
				Response.Write htmlencode(SdM_ToUser)
				If SdM_ToUser = "" Then Response.Write "��Ϊ���棬�޽�����"
			Else%>
			<input class='fminpt input_4' name=SdM_ToUser id=SdM_ToUser value="<%=htmlencode(SdM_ToUser)%>" size=41 maxlength=200> <%
				If AjaxFlag = 0 Then DisplayFriendList
			End If%>
		</td>
	</tr>
	<%End If%>
	<tr>
		<td class=tdbox>*����</td>
		<td class=tdbox>
			<input class='fminpt input_4' name=SdM_Title id=SdM_Title value="<%=htmlencode(SdM_Title)%>" size=60 maxlength=100>
		</td>
	</tr>
	<tr>
		<td valign=top class=tdbox>����<%
		If AjaxFlag = 0 Then
		%><div class=value2>֧��[IMG]��ǩ����ͼƬ</a></div><%
		End If%></td>
		<td valign=top class=tdbox>
			<textarea cols=58 name="SdM_Content" id=SdM_Content rows=10"<%
			If AjaxFlag = 0 Then%> onselect="storeCaret(this);" onclick="storeCaret(this);" onkeyup="storeCaret(this);" onkeydown="if(ctlkey(event)==false)return(false);"<%
			End If%> class=fmtxtra><%If SdM_Content<>"" Then Response.Write VbCrLf & Htmlencode(SdM_Content)%></textarea>
		</td>
	</tr>
<%
If AjaxFlag = 0 Then
	If DEF_UBBiconNumber > 0 Then%>
		<tr>
			<td width="<%=DEF_BBS_LeftTDWidth%>" valign=top class=tdbox>
			����<a href=<%=DEF_BBS_HomeUrl%>User/Help/Ubb.asp?icon target=_blank>����</a>
			</td>
			<td class=tdbox>
				<table border="0" cellspacing="0" cellpadding="0"><tr><td>
				<div id=editor_icon>
				</div>
				</td></tr></table>
			</td>
		</tr><%
	End If
End If%>
	<tr>
		<td class=tdbox>&nbsp;</td>
		<td class=tdbox>
			<input type="submit" name="Submit" value="�ύ" class="fmbtn btn_2"> &nbsp;
			<input type="reset" name="reset" value="���" class="fmbtn btn_2">
		</td>
	</tr>
	</table>
	</form>
<%If AjaxFlag = 0 Then%>
	<script type="text/javascript">
	var ValidationPassed = true,submitflag=0;
	function submitonce(theform)
	{
		submitflag = 1;
		if(theform.ForumNumber)
		{
			if(theform.ForumNumber.value=="")
			{
				alert("��������֤��!\n");
				ValidationPassed = false;
				theform.ForumNumber.focus();
				submitflag = 0;
				return false;
			}
			else
			{ValidationPassed = true;}
		}

		submit_disable(theform);
		return true;
	}
	function ctlkey(event)
	{
		if(event.ctrlKey && event.keyCode==13){submitonce($id('LeadBBSFm'));if(ValidationPassed)$id('LeadBBSFm').submit();return(false);}
		if(event.altKey && event.keyCode==83){submitonce($id('LeadBBSFm'));if(ValidationPassed)$id('LeadBBSFm').submit();return(false);}
		return(true);
	}
	function edt_icon(s1)
	{
		var str1="[EM"+s1+"]",str2="";
		var str=str1 + str2;
		var obj=$id('SdM_Content');
		obj.focus();

		if(!isUndef(obj.selectionStart)) 
		{
			str = str1 + obj.value.substr(obj.selectionStart,obj.selectionEnd-obj.selectionStart) + str2;
			obj.value = obj.value.substr(0, obj.selectionStart) + str + obj.value.substr(obj.selectionEnd);
		}
		else if ((document.selection)&&(document.selection.type== "Text"))
		{
			var range = document.selection.createRange();
			var ch_text = range.text;
			range.text = str1 + ch_text + str2;
		} 
		else
		{
			if (obj.createTextRange && obj.caretPos)
			{
				var caretPos = obj.caretPos;
				caretPos.text = str1 + caretPos.text + str2;
				obj.focus();
			}
			else{obj.value+=str;obj.focus();}
		}
	}
	function storeCaret (textEl)
	{
		if (textEl.createTextRange) 
		textEl.caretPos = document.selection.createRange().duplicate(); 
	}

var edt_escfg = 0;

function edt_disable_esc()
{
	if(event.keyCode==27)return(false);
}
function edt_disablesc()
{
	if(edt_escfg==1)return;
	edt_escfg = 1;

	if (document.all)
            $id('body').attachEvent("onkeydown",edt_disable_esc);    
        else    
            $id('body').addEventListener("onkeydown",edt_disable_esc,false);
}
function Msg_Focus()
{
	if(!Browser.is_ie){$id("SdM_Title").focus();return;}
	var obj=$id("SdM_Title");
	var r = obj.createTextRange();
	r.collapse(false);
	r.select();
	obj.focus();
}
edt_disablesc();
setTimeout("Msg_Focus()",500);
$import("../a/edit/icon.asp?f=msg","js");
window.onbeforeunload = function(){if($id("SdM_Content").value.length>0&&submitflag==0)return("���Ķ���Ϣδ����ȷ��ȡ����");}
	</script>
<%
	If CheckSupervisorUserName = 1 Then%><div class=value2>- ���ǹ���Ա������д�����˱�ʾ��������</div><%
	End If
	If DEF_User_MaxReceiveUser >= 2 Then
		%><div class=value2>- ���������д<%=DEF_User_MaxReceiveUser%>�������ˣ��ö��ŷָ�</div>
		<div class=value2>- �·��͵Ķ���Ϣ��ౣ��<%=LMT_SendMsgExpiresDate%>�죬��ʱϵͳ���Զ����.</div><%
	End IF
End If

End Function

Function CheckSdM_ToUserString
	
	Do While InStr(SdM_ToUser,",,")
		SdM_ToUser = Replace(SdM_ToUser,",,",",")
	Loop
	If Left(SdM_ToUser,1) = "," Then SdM_ToUser = Mid(SdM_ToUser,2)
	If Right(SdM_ToUser,1) = "," Then SdM_ToUser = Left(SdM_ToUser,Len(SdM_ToUser) - 1)
	TmpArr = Split(SdM_ToUser,",")
	
	If Len(SdM_ToUser) - Len(Replace(SdM_ToUser,",","")) > DEF_User_MaxReceiveUser - 1 Then
		GBL_CHK_TempStr = "���󣬽��������ֻ����д" & DEF_User_MaxReceiveUser & "�ˣ�"
		CheckSdM_ToUserString = 0
		Exit Function
	Else
		Dim N,M,TmpArr
		TmpArr = Split(SdM_ToUser,",")
		For N = 0 to Ubound(TmpArr,1)
			For M = N + 1 to Ubound(TmpArr,1)
				If LCase(TmpArr(N)) = LCase(TmpArr(M)) Then
					GBL_CHK_TempStr = "���󣬽����˲������ظ���д(�ظ��û�����" & TmpArr(N) & ")"
					CheckSdM_ToUserString = 0
					Exit Function
				End If
			Next
		Next
	End If
	CheckSdM_ToUserString = 1

End Function

Function CheckSubmitFormData

	If ModifyMessageID = 0 Then SdM_ToUser = Trim(Request.Form("SdM_ToUser"))
	SdM_Title = Trim(Request.Form("SdM_Title"))
	SdM_ConTent = Request.Form("SdM_ConTent")
	
	If DEF_EnableAttestNumber > 2 and (DEF_AttestNumberPoints = 0 or GBL_CHK_Points < DEF_AttestNumberPoints) Then
		If CheckRndNumber = 0 Then
			GBL_CHK_TempStr = "��֤����д����!"
			GBL_CHK_Flag = 0
			Exit Function
		End If
	End If

	If SdM_ToUser <> "" Then
		If ModifyMessageID = 0 Then
			If CheckSdM_ToUserString = 0 Then Exit Function
			If CheckUserNameExist = 0 Then
				Exit Function
			ElseIf CheckMessageOver(SdM_ToUser) = 1 Then
				Exit Function
			End If
		End If
	Else
		If CheckSupervisorUserName = 0 Then
			GBL_CHK_TempStr = "���󣬱�����д������!" & VbCrLf
			Exit Function
		Else
			If SdM_Title = "" Then
				GBL_CHK_TempStr = "���ǹ���Ա, ���ڷ������ǹ�����Ϣ, ����д��Ϣ����!" & VbCrLf
				Exit Function
			End If
			If inStr(Lcase(SdM_Title),"<script") or inStr(Lcase(SdM_Title),"</script") Then
				GBL_CHK_TempStr = "���ǹ���Ա, �����Ĺ�����ⲻ��ʹ��JS����!" & VbCrLf
				Exit Function
			End If
		End If
	End If

	If StrLength(SdM_Title) > 100 Then
		GBL_CHK_TempStr = "������Ϣ�����벻Ҫ����100���ַ�. " & VbCrLf
		Exit Function
	End if

	If SdM_Title = "" Then
		GBL_CHK_TempStr = "������Ϣ���������д. " & VbCrLf
		Exit Function
	End if

	If CheckSupervisorUserName = 1 Then
		If Len(SdM_Content) > DEF_MaxTextLength * 2 then
			GBL_CHK_TempStr = "�������ݲ��ܳ���" & DEF_MaxTextLength * 2 & "����!" & VbCrLf
			Exit Function
		End If
	Else
		If Len(SdM_Content) > DEF_MaxTextLength / 2 then
			GBL_CHK_TempStr = "�������ݲ��ܳ���" & DEF_MaxTextLength / 2 & "����!" & VbCrLf
			Exit Function
		End If
	End If

	If ModifyMessageID = 0 Then 
		Dim Temp
		Temp = CheckIsRestSpaceTime(SdM_Title,Left(SdM_Content & "",100))
		Select Case Temp
		Case 1: If CheckSupervisorUserName = 0 Then
				GBL_CHK_TempStr = "����������̫�����Ϣ������Ϣ" & DEF_RestSpaceTime & "���Ӻ��ٷ�����Ϣ!"
				GBL_CHK_Flag = 0
				Exit Function
			End If
		Case 2: GBL_CHK_TempStr = "�벻Ҫ���ظ�����Ϣ!"
			GBL_CHK_Flag = 0
			Exit Function
		Case 3: Exit Function
		End Select
	Else
		If CheckWriteEventSpace = 0 Then
			GBL_CHK_TempStr = "�����޸����ϵĹ������ύ��̫Ƶ�����Ժ������ύ! " & VbCrLf
			GBL_CHK_Flag = 0
			Exit Function
		End If
	End If

End Function
Sub ModifyMessage(ToUser,Title,Content,ModifyID)

	Dim SQL
	Content = Replace(Content & "","\" & VbCrLf,"\\" & VbCrLf & VbCrLf)
	SQL = "Update LeadBBS_InfoBox set Title='" & Replace(Title,"'","''") & "',Content='" & Replace(Content,"'","''") & "' where ID=" & ModifyID
	CALL LDExeCute(SQL,1)

	If ToUser = "" Then ReloadPubMessageInfo

	GBL_CHK_TempStr = "<p align=left>&nbsp; &nbsp; <font color=008800 class=greenfont>�ɹ��༭����Ϣ"

	If CheckSupervisorUserName = 0 Then
		CALL LDExeCute("Update LeadBBS_User Set LastWriteTime=" & GetTimeValue(DEF_Now) & " where ID=" & GBL_UserID,1)
		UpdateSessionValue 13,GetTimeValue(DEF_Now),0
	End If

End Sub

Function WriteNewMessageToDatabase

	If ModifyMessageID = 0 Then
		SendNewMessage SdM_fromUser,SdM_ToUser,SdM_Title,SdM_Content,GBL_IPAddress
	Else
		ModifyMessage SdM_ToUser,SdM_Title,SdM_Content,ModifyMessageID
	End If
	If GBL_CHK_TempStr <> "" Then Message_Done "<div class=value2>" & GBL_CHK_TempStr & "</div>",""

End Function

Function CheckMessageOver(username)

	Dim Rs,tmp

	Dim N,TmpArr
	TmpArr = Split(SdM_ToUser,",")

	For N = 0 to Ubound(TmpArr,1)
		Set Rs = LDExeCute("Select count(*) from LeadBBS_InfoBox where ToUser='" & Replace(TmpArr(N),"'","''") & "'",0)
		If Rs.Eof Then
			CheckMessageOver = 0
		Else
			tmp = Rs(0)
			If isNull(tmp) Then tmp = 0
			tmp = cCur(tmp)
			If tmp > LMT_MaxMessageNumber Then
				GBL_CHK_TempStr = "���󣬽����ˡ�<b>" & htmlencode(TmpArr(N)) & "</b>���ռ��伺��������ʧ��!<br>" & VbCrLf
				CheckMessageOver = 1
				Exit Function
			End If
		End if
		Rs.Close
		Set Rs = Nothing
	Next
	CheckMessageOver = 0

End Function

Rem ���ĳ�û���ID�Ƿ����
Function CheckUserNameExist

	Dim UserLimit
	Dim Rs
	Dim N,TmpArr
	TmpArr = Split(SDM_ToUser,",")
	
	For N = 0 to Ubound(TmpArr,1)
		Set Rs = LDExeCute(sql_select("Select ID,UserName,UserLimit from LeadBBS_User where UserName='" & Replace(TmpArr(N),"'","''") & "'",1),0)
		If Rs.Eof Then
			CheckUserNameExist = 0
			Sdm_ToUserID = 0
			GBL_CHK_TempStr = "�����Ҳ��������ˡ�<b>" & TmpArr(n) & "</b>������ȷ���Ƿ���ڴ���!<br>" & VbCrLf
			Rs.Close
			Set Rs = Nothing
			Exit Function
		Else
			Sdm_ToUserID = Rs(0)
			TmpArr(N) = Rs(1)
			UserLimit = Rs(2)
		End if
		Rs.Close
		Set Rs = Nothing
		If GetBinaryBit(UserLimit,13) = 1 and GBL_BoardMasterFlag < 4 Then '����������Ȩ���߲��ܴ�����
			Set Rs = LDExeCute(sql_select("Select ID from LeadBBS_FriendUser where FriendUserID=" & GBL_UserID & " and UserID=" & Sdm_ToUserID,1),0)
			If Rs.Eof Then
				GBL_CHK_TempStr = htmlencode(TmpArr(N)) & " �Ѿ����óɽ�������պ��ѵĶ���Ϣ��������š�<br>" & VbCrLf
				CheckUserNameExist = 0
			End If
			Rs.Close
			Set Rs = Nothing
		End If
	Next
	SdM_ToUser = ""
	SdM_ToUser = TmpArr(0)
	For N = 1 to Ubound(TmpArr,1)
		SdM_ToUser = SdM_ToUser + "," + TmpArr(N)
	Next
	CheckUserNameExist = 1

End Function

Function GetMessageValue(MessageID)

	Dim Rs,SQL
	SQL = sql_select("Select FromUser,SendTime,Title,Content,ToUser,ReadFlag from LeadBBS_InfoBox where ID=" & MessageID,1)
	Set Rs = LDExeCute(SQL,0)
	If Rs.Eof Then
		GBL_CHK_TempStr = "�����޷��鿴����Ϣ��������ϣ�"
	Else
		If SdM_FromUser <> Rs(0) and SdM_FromUser <> Rs(4) and Rs(4) <> "" Then
			GBL_CHK_TempStr = "�����޷��鿴����Ϣ��������ϣ�"
			Rs.Close
			Set Rs = Nothing
			Exit Function
		End If
		If Rs(4) = "" and CheckSupervisorUserName = 0 Then
			GBL_CHK_TempStr = "�����޷��鿴����Ϣ��������ϣ�"
			Rs.Close
			Set Rs = Nothing
			Exit Function
		End If
		SdM_Title = Rs(2)
		If ModifyMessageID > 0 Then
			If GBL_CHK_User <> Rs(0) and (CheckSupervisorUserName = 0 and Rs(4) = "") Then
				GBL_CHK_TempStr = "�����޷��鿴����Ϣ��������ϣ�"
				Rs.Close
				Set Rs = Nothing
				Exit Function
			End If
			If (ccur(Rs(5)) = 1) Then
				GBL_CHK_TempStr = "��������Ϣ���ģ��޷��ٱ༭��"
				Rs.Close
				Set Rs = Nothing
				Exit Function
			End If
			SdM_Content = Rs(3)
			SdM_ToUser = Rs(4)
			SdM_FromUser = Rs(0)
		Else
			'SdM_Content = Replace(Rs(3) & "",VbCrLf,VbCrLf & ">")
			'If SdM_Content<>"" Then SdM_Content = ">" & SdM_Content
			'SdM_Content = VbCrLf & VbCrLf & VbCrLf & "-----------------------------------------" & VbCrLf & ">[�ظ�" & htmlencode(Rs(0)) & " " & Mid(RestoreTime(Rs(1)),6,11) & "���͵Ķ���Ϣ]" & VbCrLf & ">���⣺" & SdM_Title & VbCrLf & VbCrLf & SdM_Content & VbCrLf
			SdM_Content = Rs(3)
			SQL = inStr(SdM_Content,"[/quote]")
			If inStr(SdM_Content,"[quote]") > 0 and SQL > 0 Then SdM_Content =  Mid(SdM_Content,SQL + 8)
			If Replace(Trim(SdM_Content),VbCrLf,"") <> "" Then SdM_Content = "[b]ԭ����[/b][hr]" & VbCrLf & SdM_Content & VbCrLf
			SdM_Content = "[quote][u]" & GBL_CHK_User & "[/u] �ظ� [u]" & htmlencode(Rs(0)) & "[/u] �� " & Left(RestoreTime(Rs(1)),16) & " ���͵Ķ���Ϣ" & VbCrLf & VbCrLf & "[b]ԭ���⣺[/b][url=LookMessage.asp?MessageID=" & ReplyMessageID & "]" & SdM_Title & "[/url]" & VbCrLf & VbCrLf & SdM_Content & "[/quote]" & VbCrLf
			If Left(SdM_Title,3) <> "Re:" Then SdM_Title = Left("Re:" & SdM_Title,250)
		End If
	End If
	Rs.Close
	Set Rs = Nothing

End Function

Function CheckIsRestSpaceTime(Form_Title,Form_Content)

	If CheckWriteEventSpace = 0 Then
		CheckIsRestSpaceTime = 3
		Exit Function
	End If
	Dim Rs,ndatetime,Temp_ID,Temp_Title,Temp_Content
	Set Rs = LDExeCute(sql_select("Select SendTime,ID,title,Content from LeadBBS_InfoBox where FromUser='" & Replace(Sdm_FromUser,"'","''") & "' Order by ID DESC",1),0)
	If Rs.Eof Then
		CheckIsRestSpaceTime = 0
		Rs.Close
		Set Rs = Nothing
		Exit Function
	Else
		ndatetime = Rs(0)
		Temp_ID = Rs(1)
		Temp_Title = Rs(2)
		Temp_Content = Left(Rs(3) & "",100)
		Rs.Close
		Set Rs = Nothing
	End if
	
	If DateDiff("s", RestoreTime(ndatetime), DEF_Now) < 0 Then
		CheckIsRestSpaceTime = 0
		Exit Function
	End If
	If DateDiff("s", RestoreTime(ndatetime), DEF_Now) < DEF_RestSpaceTime Then
		CheckIsRestSpaceTime = 1
		CALL LDExeCute("Update LeadBBS_InfoBox set SendTime=" & GetTimeValue(DEF_Now) & " where id=" & Temp_ID,1)
	Else
		If Temp_Title = Form_Title and Temp_Content = Form_Content Then
			CheckIsRestSpaceTime = 2
		Else
			CheckIsRestSpaceTime = 0
		End If
	End If

End Function

Sub DisplayFriendList

	%>
	<script type="text/javascript">
	function sendmsg_selfriend(val,obj)
	{
		var num = <%=DEF_User_MaxReceiveUser%>;
		if(val=='')
		{
			layer_view('ѡ�������',$id('user_selfriend'),'','','anc_msgbody','SendMessage.asp?go=liling','',0,'AjaxFlag=1',0,0);
		}
		else
		{
			if(num <2 )
			$id('SdM_ToUser').value=val;
			else
			{
				if($id('SdM_ToUser').value=='')
				{
					$id('SdM_ToUser').value=val;
					layer_view('�ɹ���ӡ�',obj,'','','user_selfriend_alert','','',0,'',0,-55);
				}
				else
				{
					if((","+$id('SdM_ToUser').value+",").indexOf(","+val+",")!=-1)
					{
						layer_view('�����ظ������û��ѱ���ӡ�',obj,'','','user_selfriend_alert','','',0,'',0,-55);
						return(false);
					}
					if(($id('SdM_ToUser').value.length-$replace($id('SdM_ToUser').value,",","").length) < num-1)
					{
					$id('SdM_ToUser').value += ',' + val;
					layer_view('�ɹ���ӡ�',obj,'','','user_selfriend_alert','','',0,'',0,-55);
					}
					else
					layer_view('���ʧ�ܣ����������������д' + num + '�ˡ�',obj,'','','user_selfriend_alert','','',0,'',0,-55);
				}
			}
		}
		return(false);
	}
	
	</script>
	<a href="javascript:;" onclick="sendmsg_selfriend('',this);" class="layerico" id="user_selfriend">ѡ�����</a>
	<%

End Sub


Rem -------�µ������ڲ��ִ���-----

Sub Main_SelectFriend

	initDatabase
	Response.Write "<div class=ajaxbox>"
	If GBL_UserID = 0 Then
		GBL_CHK_TempStr = ""
		If GBL_UserID = 0 Then GBL_CHK_TempStr = "�Ҳ����û���Ҫ�鿴�������ȵ�¼��<br>" & VbCrLf
	Else
		GBL_CHK_TempStr = ""
	End If

	if GBL_CHK_TempStr <> "" Then
		Response.Write "<div class=alert>" & GBL_CHK_TempStr & "</div><hr class=splitline>"
	Else
		GBL_CHK_TempStr = ""
		DisplayCenter
	End If
	closeDataBase
	Response.Write "</div>"

End Sub

Sub DisplayCenter

	Dim Rs,SQL
	Set rs = Server.CreateObject("ADODB.Recordset")

	SQL = sql_select("Select ID,UserName from LeadBBS_User where ID=" & GBL_UserID,1)
	Set Rs = LDExeCute(SQL,0)
	If Rs.Eof Then
		Response.Write "<div class=alert>�����û���֤����</div>"
		Rs.Close
		Set Rs = Nothing
		Exit Sub
	End if
	Rs.close
	Set Rs = Nothing
	
	SQL = "select T2.UserName,T2.UserName from LeadBBS_FriendUser as T1 left join LeadBBS_User as T2 on T1.FriendUserID=T2.ID where T1.UserID=" & GBL_UserID & " Order by T1.ID ASC"
	Set Rs = LDExeCute(SQL,0)
	If Not rs.Eof Then
		Response.Write "<div class=u_friendlist><div class=title>��������ѡ�����</div><ul><li><a href=#1 class=layer_alertclick onclick="&chr(34)&"sendmsg_selfriend('"
		Response.Write Rs.GetString(,,"',this)" & chr(34) & ">","</a></li>" & VbCrLf & "<li><a href=#1 class=layer_alertclick onclick="&chr(34)&"sendmsg_selfriend('","")
		Response.Write "','');""></a></li></ul></div>"
	End If
	Rs.close
	Set Rs = Nothing

End Sub%>