<%Const LMT_MaxFriendNum = 200 '������ӵ���������Ŀ
Function CheckAddFriendSure

	If GetBinarybit(GBL_CHK_UserLimit,1) = 1 Then
		Processor_ErrMsg "����Ȩ�޲��㣬����ʽ�û��޴˹��ܣ�" & VbCrLf
		CheckAddFriendSure = 0
		Exit Function
	End If
	CheckAddFriendSure = 1

End Function


Function DisplayAddFriend

	Dim FriendName,FriendID
	FriendName = Left(Request("FriendName"),20)
	If Request.Form("SureFlag")="1" Then
		Dim Rs,SQL
		SQL = "Select count(*) from LeadBBS_FriendUser where UserID=" & GBL_UserID
		Set Rs = LDExeCute(SQL,0)
		If Rs.Eof Then
			SQL = 0
		Else
			SQL = Rs(0)
			If IsNull(SQL) Then SQL = 0
			SQL = cCur(SQL)
		End If
		Rs.Close
		Set Rs = Nothing

		If SQL > LMT_MaxFriendNum Then
			Processor_ErrMsg "���󣬺������Ѿ�����" & LMT_MaxFriendNum & "�ˣ���������ӣ�" & VbCrLf
			Set Rs = Nothing
			Exit Function
		End if

		SQL = sql_select("Select ID,UserName from LeadBBS_User where UserName='" & Replace(FriendName,"'","''") & "'",1)
		Set Rs = LDExeCute(SQL,0)
		If Rs.Eof Then
			Processor_ErrMsg "����ȷ��д�ĺ������ƣ�" & VbCrLf
			Rs.Close
			Set Rs = Nothing
			Exit Function
		End If
		FriendID = cCur(Rs(0))
		FriendName = Rs(1)
		Rs.Close
		Set Rs = Nothing
		
		SQL = sql_select("Select ID from LeadBBS_FriendUser where FriendUserID=" & FriendID & " and UserID=" & GBL_UserID,1)
		Set Rs = LDExeCute(SQL,0)
		If Not Rs.Eof Then
			//Processor_ErrMsg "<b>" & htmlencode(FriendName) & "</b> �Ѿ������ĺ��ѣ��޷��ظ���ӣ�" & VbCrLf
			Processor_ErrMsg "<div id=collect_msg><b>" & htmlencode(FriendName) & "</b> �Ѿ������ĺ��ѣ��޷��ظ���ӣ�<br /><a href=""javascript:p_url = '" & DEF_BBS_HomeUrl & "User/DeleteMessage.asp';" & VbCrLf & "p_para='AjaxFlag=1&FriendFlag=1&DeleteSureFlag=dk9@dl9s92lw_SWxl&MessageID=';" & VbCrLf & "p_command = '$id(\'collect_msg\').innerHTML=tmp';" & VbCrLf & "p_type = 1;" & VbCrLf & "p_once(" & Rs(0) & ");"">�������ɾ���˺��ѡ�</a></div>" & VbCrLf
			Rs.Close
			Set Rs = Nothing
			Exit Function
		End If
		Rs.Close
		Set Rs = Nothing

		CALL LDExeCute("insert into LeadBBS_FriendUser(FriendUserID,UserID) Values(" & FriendID & "," & GBL_UserID & ")",1)
		Set Rs = Nothing
		If CheckSupervisorUserName = 0 Then
			CALL LDExeCute("Update LeadBBS_User Set LastWriteTime=" & GetTimeValue(DEF_Now) & " where ID=" & GBL_UserID,1)
			UpdateSessionValue 13,GetTimeValue(DEF_Now),0
		End If
		SendNewMessage Prc_User,FriendName,"��̳���ţ���Ӻ���֪ͨ","[url=../User/LookUserInfo.asp?name=" & urlencode(GBL_CHK_User) & "]" & GBL_CHK_User & "[/url]�����Ϊ����." & VbCrLf,GBL_IPAddress
		Processor_Done "�ɹ����" & htmlencode(FriendName) & "�������б�"
	Else
		Processor_Head
		
		Dim Url
		Url = htmlencode(Left(Request("dir"),100))
		If Request("dir") = "" Then
			Url = DEF_BBS_HomeUrl
		End If
		%>
		<form name=DellClientForm action="<%=Url%>a/Processor.asp?action=AddFriend&b=<%=Request("B")%>" onSubmit="submit_disable(this);" method=post<%
	If AjaxFlag = 1 Then
		Response.Write " target=""hidden_frame"""
	End If
	%>>
			<input type=hidden name=SureFlag value="1">
			<input type=hidden name=JsFlag value="1">
			<input type=hidden name=Url value="<%=Url%>">
			<input type=hidden name=AjaxFlag value="<%=AjaxFlag%>">
			<input type=hidden name=ID value="<%=Request("ID")%>">
			<input type=hidden name=BoardID value="<%=Request("B")%>">
			<div class=value2>
			�������֣�
			<input type=input name=FriendName value="<%=FriendName%>" class='fminpt input_2'></div>			
			<div class=value2><br /><input type=submit value=��Ϊ���� class="fmbtn btn_3"></div>
		</form>
		<%Processor_Bottom
	End If

End Function%>