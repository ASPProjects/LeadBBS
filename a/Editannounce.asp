<!-- #include file=../inc/BBSsetup.asp -->
<!-- #include file=../inc/Ubbcode.asp -->
<!-- #include file=../inc/Board_Popfun.asp -->
<!-- #include file=../inc/Limit_Fun.asp -->
<!-- #include file=inc/MakeAnnounceTop.asp -->
<%DEF_BBS_HomeUrl = "../"%>
<!-- #include file=inc/Editor_Fun.asp -->
<!-- #include file=../inc/Upload_Fun.asp -->
<!-- #include file=inc/upload1_fun.asp -->
<%
Const LMTDEF_MinAnnounceLength = 2 '�༭�ύ������������Ҫ��������
Const LMT_BuyAnnounceMaxPoints = 9 '���������ĵ�������
Const LMTDEF_NeedCachetValue = 1 '�趨���������û������Լ�����ר��

Dim Form_EditAnnounceID,Form_BoardID,Form_RootID,Form_ParentID
Dim Form_Title,Form_Content,Form_FaceIcon,Form_ndatetime
Dim Form_Length,Form_UserName,Form_UserID,Form_HTMLFlag,Form_UnderWriteFlag
Dim Form_NoUserUnderWriteFlag,Form_NotReplay
Dim Form_GoodAssort,Form_GoodAssort_Old
Dim Form_TopicType,Form_NeedValue,Form_TitleStyle,Form_Topictype_Old,Form_TitleStyle_Old,Form_Title_Old
Form_TitleStyle = 0

Dim Form_VoteType,Form_VoteItem,Form_Vote_ExpireDay,VoteFlag,Form_VoteType_Old
VoteFlag = 0

Dim LMT_TopicName,LMT_TopicNameNoHTML,LMT_TopicTitleStyle,LMT_RootIDBak,LMT_TopicNameNoHTML_Temp
Dim Upd_SpendFlag,Upd_ErrInfo,Form_UpClass,Form_UpFlag,Form_Submitflag,Form_ForumNumber

Form_HTMLFlag = 2
Dim LMT_RootMaxID,LMT_RootMinID
LMT_RootMaxID = 0
LMT_RootMinID = 0

const PageSplitNum = 10

Dim LMT_DefaultEdit,LMT_EnableUpload
LMT_DefaultEdit = DEF_UbbDefaultEdit

Dim LMT_MaxTextLength,SupervisorFlag,VoteGetData,VoteNumber
SupervisorFlag = CheckSupervisorUserName
If SupervisorFlag = 0 Then
	LMT_MaxTextLength = DEF_MaxTextLength
Else
	LMT_MaxTextLength = DEF_MaxTextLength * 4
End If

EditFlag = 1

Function DisplayAnnounceForm

	Dim Temp
	Temp = GBL_CHK_TempStr
%>

<script type="text/javascript">
<!--
var submitflag=0;
function submitonce(theform)
{
	submitflag = 1;
	var lg;<%If DEF_EnableAttestNumber > 2 and (DEF_AttestNumberPoints = 0 or GBL_CHK_Points < DEF_AttestNumberPoints) Then%>
		
	if(theform.ForumNumber.value=="")
	{
		alert("��������֤��!\n");
		ValidationPassed = false;
		theform.ForumNumber.focus();
		submitflag = 0;
		return;
	}
	else
	{ValidationPassed = true;
	}<%End If%>
	
	edt_checkContent();
	lg = edt_getdoclen();
	if(lg < <%=LMTDEF_MinAnnounceLength%>)
	{
		alert("��������ݳ��Ȳ�����Ҫ�� \n\n����Ҫ��<%=LMTDEF_MinAnnounceLength%>���֣�Ŀǰ����" + lg + "����\n");
		ValidationPassed = false;
		submitflag = 0;
		return;
	}
	if(lg > <%=LMT_MaxTextLength%>)
	{
		alert("��������ݳ�����<%=LMT_MaxTextLength%>���֣�Ŀǰ����" + lg + "����\n");
		ValidationPassed = false;
		submitflag = 0;
		return;
	}
	else
	{
		ValidationPassed = true;
	}
	submit_disable(theform);
}

function storeCaret (textEl)
{
	if (textEl.createTextRange) 
	textEl.caretPos = document.selection.createRange().duplicate(); 
}

function ctlkey(event)
{
	if(event.ctrlKey && event.keyCode==13){submitonce(document.LeadBBSFm);if(ValidationPassed)document.LeadBBSFm.submit();return(false);}
	if(event.altKey && event.keyCode==83){submitonce(document.LeadBBSFm);if(ValidationPassed)document.LeadBBSFm.submit();return(false);}
}
//-->
</script><%
DisplayPreview
Global_TableHead
LMT_EnableUpload = 1
If GBL_UserID < 1 Then LMT_EnableUpload = 0
Select Case DEF_EnableUpload
	Case 0: LMT_EnableUpload = 0
	case 2: If CheckSupervisorUserName = 0 Then LMT_EnableUpload = 0
	Case 3: If GBL_BoardMasterFlag < 4 Then LMT_EnableUpload = 0
	Case 4: If GetBinarybit(GBL_CHK_UserLimit,2) = 0 Then LMT_EnableUpload = 0
	Case 5: If GBL_BoardMasterFlag < 4 and GetBinarybit(GBL_CHK_UserLimit,2) = 0 Then LMT_EnableUpload = 0
End Select

If DEF_Upd_SpendFlag = 0 and GBL_BoardMasterFlag >=4 Then
	Upd_SpendFlag = 0
Else
	Upd_SpendFlag = 1
End If

If Upd_SpendFlag = 1 and DEF_UploadSpendPoints > 0 and DEF_UploadSpendPoints > GBL_CHK_Points Then LMT_EnableUpload = 0
If LMT_EnableUpload = 1 and (GBL_CHK_OnlineTime >= DEF_NeedOnlineTime or DEF_NeedOnlineTime = 0) Then
	LMT_EnableUpload = 1
Else
	LMT_EnableUpload = 0
End If
%>
<div class=contentbox>
		<table border=0 cellpadding=0 cellspacing=0 width="100%" class=tablebox>
		<tr>
			<td class=tbhead><div class=value><%
				Response.Write "�༭���ӣ�"  & LMT_TopicNameNoHTML_Temp%></div></td>
		</tr>
		</table>
		<!-- #include file=inc/post_layer.asp -->
		<%If LMT_EnableUpload = 0 Then %>
		<form action=EditAnnounce.asp method=post id=LeadBBSFm name=LeadBBSFm onSubmit="submitonce(this);return ValidationPassed;">
		<%Else%>
		<form action="EditAnnounce.asp?dontRequestFormFlag=1" id=LeadBBSFm name=LeadBBSFm method="post" enctype="multipart/form-data" onsubmit="submitonce(this);if(ValidationPassed)Upl_submit();return ValidationPassed;">
		<%End If%>
		<table border=0 cellpadding=0 cellspacing=0 width="100%" class=tablebox>
		<tr>
			<td width="<%=DEF_BBS_LeftTDWidth%>" class=tdleft>*ԭ������</td>
			<td class=tdright>
				<input name=test value="testvalue" type=hidden>
				<input maxlength=20 name=Form_User value="<%=htmlencode(Form_UserName)%>" size="20" readonly class='fminpt input_2'>
				<input name=submitflag value="slzOowl_kdO8m610" type=hidden>
				<input name=BoardID value="<%=Form_BoardID%>" type=hidden>
				<input name=ID value="<%=Form_EditAnnounceID%>" type=hidden>
			</td>
                </tr>
		</tr>
		<tr>
			<td width="<%=DEF_BBS_LeftTDWidth%>" class=tdleft>*���ӱ���</td>
			<td class=tdright>
				<%
				If VoteFlag = 0 Then%>
					<input maxlength=255 name=Form_Title size="49" value="<%
					If (Form_Submitflag = "first" or Form_Submitflag = "") and form_ParentID > 0 and Form_Title="" Then
						If Left(LMT_TopicName,3) <> "Re:" Then
							Response.Write "Re:" & htmlencode(LMT_TopicNameNoHTML)
						Else
							Response.Write htmlencode(LMT_TopicNameNoHTML)
						End If
					Else
						Response.Write htmlencode(Form_title)
					End If%>" class='fminpt input_4'><%
				Else%>
              		<input maxlength=255 name=Form_Title size="35" value="<%=htmlencode(Form_Title)%>" class='fminpt input_4'>
					<%If isNumeric(Form_Vote_ExpireDay) = 0 then Form_Vote_ExpireDay = 0
					Form_Vote_ExpireDay = cCur(Form_Vote_ExpireDay)%>
					<select name="Form_Vote_ExpireDay">
						<option value=0>����ʱ��</option>
						<option value=0<%If Form_Vote_ExpireDay = 0 Then Response.Write " selected"%>>��������</option>
						<option value=1<%If Form_Vote_ExpireDay = 1 Then Response.Write " selected"%>>һ��</option>
						<option value=2<%If Form_Vote_ExpireDay = 2 Then Response.Write " selected"%>>����</option>
						<option value=3<%If Form_Vote_ExpireDay = 3 Then Response.Write " selected"%>>����</option>
						<option value=7<%If Form_Vote_ExpireDay = 7 Then Response.Write " selected"%>>һ��</option>
						<option value=10<%If Form_Vote_ExpireDay = 10 Then Response.Write " selected"%>>ʮ��</option>
						<option value=15<%If Form_Vote_ExpireDay = 15 Then Response.Write " selected"%>>�����</option>
						<option value=20<%If Form_Vote_ExpireDay = 20 Then Response.Write " selected"%>>��ʮ��</option>
						<option value=30<%If Form_Vote_ExpireDay = 30 Then Response.Write " selected"%>>һ����</option>
						<option value=45<%If Form_Vote_ExpireDay = 45 Then Response.Write " selected"%>>һ���°�</option>
						<option value=60<%If Form_Vote_ExpireDay = 60 Then Response.Write " selected"%>>������</option>
						<option value=90<%If Form_Vote_ExpireDay = 90 Then Response.Write " selected"%>>������</option>
						<option value=120<%If Form_Vote_ExpireDay = 120 Then Response.Write " selected"%>>�ĸ���</option>
						<option value=180<%If Form_Vote_ExpireDay = 180 Then Response.Write " selected"%>>������</option>
						<option value=240<%If Form_Vote_ExpireDay = 240 Then Response.Write " selected"%>>�˸���</option>
						<option value=365<%If Form_Vote_ExpireDay = 365 Then Response.Write " selected"%>>һ��</option>
					</select>
				<%End If
				If GBL_BoardMasterFlag >= 5 Then%>
				<select name="Form_TitleStyle">
					<option value=0<%If Form_TitleStyle = 0 Then Response.Write " selected"%>>��ʽ</option><%If GBL_BoardMasterFlag >= 9 or Form_TitleStyle_Old = 1 Then%>
					<option value=1<%If Form_TitleStyle = 1 Then Response.Write " selected"%>>HTML</option><%End If%>
					<option value=2<%If Form_TitleStyle = 2 Then Response.Write " selected"%>>��ɫ</option>
					<option value=3<%If Form_TitleStyle = 3 Then Response.Write " selected"%>>��ɫ</option>
					<option value=4<%If Form_TitleStyle = 4 Then Response.Write " selected"%>>��ɫ</option>
					<option value=5<%If Form_TitleStyle = 5 Then Response.Write " selected"%>>����</option>
					<option value=6<%If Form_TitleStyle = 6 Then Response.Write " selected"%>>�غ�</option>
					<option value=7<%If Form_TitleStyle = 7 Then Response.Write " selected"%>>����</option>
					<option value=8<%If Form_TitleStyle = 8 Then Response.Write " selected"%>>����</option>
				</select>
				<%End If

				If cCur(Form_ParentID)=0 and (GBL_CHK_CachetValue >= LMTDEF_NeedCachetValue or GBL_BoardMasterFlag >= 4) Then
					Dim TArray,Num,N,TArray2
					TArray = Application(DEF_MasterCookies & "BoardInfo" & GBL_Board_ID & "_TI")
					TArray2 = Application(DEF_MasterCookies & "BoardInfo" & 0 & "_TI")
					If isArray(TArray) = False Then
						If TArray & "" <> "yes" Then ReloadTopicAssort(GBL_Board_ID)
						TArray = Application(DEF_MasterCookies & "BoardInfo" & GBL_Board_ID & "_TI")
					End If
					If isArray(TArray2) = False Then
						If TArray2 & "" <> "yes" Then ReloadTopicAssort(0)
						TArray2 = Application(DEF_MasterCookies & "BoardInfo" & 0 & "_TI")
					End If
					If isArray(TArray) = True or isArray(TArray2) = True Then%>
					<select name="Form_GoodAssort"  style="width:74">
					<%
						If isArray(TArray) = True Then
							Response.Write "<Option value=0 class=TBBG1>ѡ��ר��" & VbCrLf
							Num = Ubound(TArray,2)
							For N = 0 To Num
								If cCur(TArray(0,N)) = Form_GoodAssort Then
									Response.Write "<Option value=" & TArray(0,N) & " selected>" & TArray(1,N) & VbCrLf
								Else
									Response.Write "<Option value=" & TArray(0,N) & ">" & TArray(1,N) & VbCrLf
								End If
							Next
						End If
						If isArray(TArray2) = True Then
							Response.Write "<Option value=0 class=TBBG1>=��ר��=" & VbCrLf
							Num = Ubound(TArray2,2)
							For N = 0 To Num
								If cCur(TArray2(0,N)) = Form_GoodAssort Then
									Response.Write "<Option value=" & TArray2(0,N) & " selected>" & TArray2(1,N) & VbCrLf
								Else
									Response.Write "<Option value=" & TArray2(0,N) & ">" & TArray2(1,N) & VbCrLf
								End If
							Next
						End If
						Response.Write "</select>" & VbCrLf
					End If
				End If%> �255��</td>
		</tr><%If VoteFlag = 1 Then%>
		<tr>
			<td width="<%=DEF_BBS_LeftTDWidth%>" class=tdleft>*ͶƱѡ��
			<p>ÿ��һ��ͶƱѡ����ܼ���ѡ�������ѡ����ൽ<br><%=DEF_VOTE_MaxNum%>�������޸�ѡ��
			<p><%If isNumeric(Form_VoteType) = 0 then Form_VoteType = 0
				Form_VoteType = cCur(Form_VoteType)%><table border=0 cellpadding=0 cellspacing=0><tr><td><input class=fmchkbox type=radio name=Form_VoteType value=0 <%If Form_VoteType = 0 Then Response.Write " checked"%>></td><td>��ѡƱ</td>
          		<td><input class=fmchkbox type=radio name=Form_VoteType value=1 <%If Form_VoteType = 1 Then Response.Write " checked"%>></td><td>��ѡƱ</td></tr></table>
			</td>
			<td class=tdright>
				<textarea cols=80 name=Form_VoteItem rows=8 style="width: 95%; word-break: break-all;" onkeydown="if(ctlkey(event)==false)return(false);" class=fmtxtra><%If Form_VoteItem <> "" Then Response.Write VbCrLf & htmlEncode(Form_VoteItem)%></textarea>
				</td>
		</tr><%End If%>
		<tr>
			<td width="<%=DEF_BBS_LeftTDWidth%>" class=tdleft>��������</td>
			<td class=tdright>
				<input name=Form_FaceIcon class=fmchkbox type=radio value=0>��
				<input name=Form_FaceIcon class=fmchkbox type=radio value=1<%If Form_FaceIcon=1 Then Response.WRite " CHECKED"%>><img src="../images/<%=GBL_DefineImage%>bf/FACE1.GIF" class=absmiddle>
				<input name=Form_FaceIcon class=fmchkbox type=radio value=2<%If Form_FaceIcon=2 Then Response.WRite " CHECKED"%>><img src="../images/<%=GBL_DefineImage%>bf/FACE2.GIF" class=absmiddle>
				<input name=Form_FaceIcon class=fmchkbox type=radio value=3<%If Form_FaceIcon=3 Then Response.WRite " CHECKED"%>><img src="../images/<%=GBL_DefineImage%>bf/FACE3.GIF" class=absmiddle>
				<input name=Form_FaceIcon class=fmchkbox type=radio value=5<%If Form_FaceIcon=5 Then Response.WRite " CHECKED"%>><img src="../images/<%=GBL_DefineImage%>bf/FACE5.GIF" class=absmiddle>
				<input name=Form_FaceIcon class=fmchkbox type=radio value=6<%If Form_FaceIcon=6 Then Response.WRite " CHECKED"%>><img src="../images/<%=GBL_DefineImage%>bf/FACE6.GIF" class=absmiddle>
				<input name=Form_FaceIcon class=fmchkbox type=radio value=7<%If Form_FaceIcon=7 Then Response.WRite " CHECKED"%>><img src="../images/<%=GBL_DefineImage%>bf/FACE7.GIF" class=absmiddle>
				<input name=Form_FaceIcon class=fmchkbox type=radio value=15<%If Form_FaceIcon=15 Then Response.WRite " CHECKED"%>><img src="../images/<%=GBL_DefineImage%>bf/FACE15.GIF" class=absmiddle>
				<input name=Form_FaceIcon class=fmchkbox type=radio value=16<%If Form_FaceIcon=16 Then Response.WRite " CHECKED"%>><img src="../images/<%=GBL_DefineImage%>bf/FACE16.GIF" class=absmiddle>
			</td>
		</tr><%
		DisplayLeadBBSEditor1
		If Form_ParentID=0 and Form_TopicType <> 80 Then%>
		<tr>
			<td width="<%=DEF_BBS_LeftTDWidth%>" class=tdleft>���ܱ�������</td>
			<td class=tdright>
				<table border="0" cellspacing="0" cellpadding="0" class="blanktable"><tr><td>
					<select name=Form_TopicType onchange="if(this.value<49)$id('NextContactDateDiv').style.display='none';if(this.value>=49)$id('NextContactDateDiv').style.display='block';">
						<%If Form_Topictype_Old <> 54 Then%><option value="0">��ѡ����������...
						<%If DEF_EnableSpecialTopic = 1 and GetBinarybit(GBL_Board_BoardLimit,14) = 1 Then%><option value="7"<%If Form_TopicType = 7 Then Response.Write " selected"%>>�ظ��������ܲ鿴<%End If%>
						<option value="50"<%If Form_TopicType = 50 Then Response.Write " selected"%>>�鿴������Ҫ�ﵽ<%=DEF_PointsName(0)%>
						<option value="51"<%If Form_TopicType = 51 Then Response.Write " selected"%>>�ظ�������Ҫ�ﵽ<%=DEF_PointsName(0)%>
						<option value="52"<%If Form_TopicType = 52 Then Response.Write " selected"%>>�鿴������Ҫ�ﵽ<%=DEF_PointsName(4)%>
						<option value="53"<%If Form_TopicType = 53 Then Response.Write " selected"%>>�ظ�������Ҫ�ﵽ<%=DEF_PointsName(4)%>
						<option value="55"<%If Form_TopicType = 55 Then Response.Write " selected"%>>ֻ��ָ���û��ܲ鿴������<%End If%>
						<%If (Form_TopicType_Old = 54 or (DEF_EnableSpecialTopic = 1 and GetBinarybit(GBL_Board_BoardLimit,14) = 1)) and Form_TopicType_Old <> 49 Then%>
						<option value="54"<%If Form_TopicType = 54 Then Response.Write " selected"%>>���۱���������<%=DEF_PointsName(0)%><%End If%>
						<%If (Form_TopicType_Old = 49 or (DEF_EnableSpecialTopic = 1 and GetBinarybit(GBL_Board_BoardLimit,14) = 1)) and Form_TopicType_Old <> 54 Then%>
						<option value="49"<%If Form_TopicType = 49 Then Response.Write " selected"%>>���۱���������<%=DEF_PointsName(1)%><%End If%>
						<%If Form_Topictype_Old <> 54 Then%><option value="1"<%If Form_TopicType = 1 Then Response.Write " selected"%>>������<%=DEF_PointsName(8)%>���ܲ鿴
						<option value="2"<%If Form_TopicType = 2 Then Response.Write " selected"%>>������<%=DEF_PointsName(8)%>���ܻظ�
						<option value="3"<%If Form_TopicType = 3 Then Response.Write " selected"%>>��<%=DEF_PointsName(8)%>���ܲ鿴
						<option value="4"<%If Form_TopicType = 4 Then Response.Write " selected"%>>��<%=DEF_PointsName(8)%>���ܻظ�
						<option value="5"<%If Form_TopicType = 5 Then Response.Write " selected"%>>��<%=DEF_PointsName(5)%>���ܲ鿴
						<option value="6"<%If Form_TopicType = 6 Then Response.Write " selected"%>>��<%=DEF_PointsName(5)%>���ܻظ�<%End If%>
					</select>
				</td><td>
				<span name=NextContactDateDiv id=NextContactDateDiv<%If Form_TopicType<49 Then Response.Write " style='display:none'"%>>
					<input name=Form_NeedValue value="<%If cStr(Form_NeedValue) <> "0" Then Response.Write htmlencode(Form_NeedValue)%>" size=10 maxlength=10 class='fminpt input_1'></span>
				</td><td>
					<a href=#icon onclick="if(LeadBBSFm.Form_TopicType.value=='0'){alert('�ڲ������ر�ǩǰ��ѡ����������.');}else{addcontent(0,'HIDDEN','/HIDDEN');}">�������ر�ǩ</a>
				</td></tr></table>
		</tr><%End If%>
		<tr class=tdleft>
			<td class=tdleft>����ѡ��</td>
			<td class=tdright>
				<label>
				<input type="checkbox" class=fmchkbox name="Form_UnderWriteFlag" value="checkbox"<%If Form_UnderWriteFlag=1 Then Response.Write " checked"%>>��ʾǩ��</label>
				<label>
				<input type="checkbox" class=fmchkbox name="Form_NotReplay" value="checkbox"<%If Form_NotReplay = 1 Then Response.Write " checked"
						If GBL_BoardMasterFlag < 5 and Form_NotReplay = 1 Then Response.Write " DISABLED"%>><%
						If Form_ParentID=0 Then
							Response.Write "��������"
						Else
							Response.Write "��������"
						End If%></label>
					- <a href="<%=DEF_BBS_HomeUrl%>User/Help/Ubb.asp?colo" target=_blank>��ɫ��</a>
					
				<%If DEF_EnableAttestNumber > 2 and (DEF_AttestNumberPoints = 0 or GBL_CHK_Points < DEF_AttestNumberPoints) Then%>
				<div style="line-height:400%">��֤��
				<%
					displayVerifycode%></div><%
				End If%></div>
			</td>
		</tr>
		<tr>
			<td class=tdleft>&nbsp;</td>
			<td class=tdright>
				<br />
				<input name=submit2 type=submit value="��ɱ༭" class="fmbtn btn_3">
				<input id=Preview_btn type=button value="Ԥ���༭" onclick="edt_preview();" class="fmbtn btn_3">
				<br /><br />
			</td>
		</tr>
		</table>
		</form>
</div>
<%
	Global_TableBottom

End Function

Function GetFormData(name)

	If Form_UpFlag = 0 Then
		GetFormData = Request.Form(name)
	Else
		GetFormData = Form_UpClass.form(name)
	End If

End Function

Sub Get_PublicValue

	If Request.QueryString("dontRequestFormFlag") = "" Then
		Form_UpFlag = 0
	Else
		Form_UpFlag = 1
		Server.ScriptTimeOut=3000
		set Form_UpClass=new upload_Class
		Form_UpClass.ProgressID = Request.QueryString("Upload_ID")
		Form_UpClass.GetUpFile
	End If
	Form_Submitflag = Request.QueryString("submitflag")
	If Form_Submitflag = "" Then Form_Submitflag = GetFormData("submitflag")

	Form_EditAnnounceID = Request.QueryString("ID")
	If Form_EditAnnounceID = "" Then Form_EditAnnounceID = Left(GetFormData("ID"),14)
	If isNumeric(Form_EditAnnounceID) = 0 Then Form_EditAnnounceID = 0
	Form_EditAnnounceID = cCur(Form_EditAnnounceID)

	If GBL_Board_ID = 0 Then
		GBL_Board_ID = GetFormData("BoardID")
		If GBL_Board_ID = "" Then GBL_Board_ID = GetFormData("b")
		GBL_Board_ID = Left(GBL_Board_ID,14)
		If isNumeric(GBL_Board_ID)=0 Then GBL_Board_ID=0
		GBL_Board_ID = Fix(cCur(GBL_Board_ID))
		If GBL_Board_ID > 2147479999 Then GBL_Board_ID = 0
		If GBL_Board_ID > 0 Then Borad_GetBoardIDValue(GBL_Board_ID)
	End If

	Form_BoardID = GBL_board_ID

End Sub

Function GetRequestValue

	If Form_Submitflag = "slzOowl_kdO8m610" Then
		Form_Title = Trim(GetFormData("Form_Title"))
		Form_HTMLFlag = GetFormData("Form_HTMLFlag")
		If Form_HTMLFlag="2" Then
			Form_HTMLFlag=2
		ElseIf Form_HTMLFlag = "1" and ((GetBinarybit(GBL_CHK_UserLimit,16) = 1 and GBL_BoardMasterFlag >= 2) or SupervisorFlag = 1) and GBL_UserID > 0 Then
			Form_HTMLFlag = 1
		Else
			Form_HTMLFlag = 0
		End If
	Else
		Form_HTMLFlag = 2
	End If

	Form_Content = GetFormData("Form_Content")
	
	Form_FaceIcon = Left(GetFormData("Form_FaceIcon"),14)
	If isNumeric(Form_FaceIcon) = 0 Then Form_FaceIcon = 0	
	Form_FaceIcon = Fix(cCur(Form_FaceIcon))
	If Form_FaceIcon < 0 or Form_FaceIcon > 16 Then Form_FaceIcon = 0
	
	Form_NoUserUnderWriteFlag = GetFormData("Form_NoUserUnderWriteFlag")
	If Form_NoUserUnderWriteFlag="checkbox" Then
		Form_NoUserUnderWriteFlag = 1
	Else
		Form_NoUserUnderWriteFlag = 0
	End If

	If GBL_BoardMasterFlag >= 5 or Form_NotReplay = 0 Then
		Form_NotReplay = GetFormData("Form_NotReplay")
		If Form_NotReplay <> "" Then 
			Form_NotReplay = 1
		Else
			Form_NotReplay = 0
		End If
	End If

	If IsNull(Form_TopicType) Then Form_TopicType= 0
	If IsNull(Form_NeedValue) Then Form_NeedValue = 0

	If Form_ParentID = 0 and Form_Topictype_Old <> 80 Then
		Form_TopicType = Left(GetFormData("Form_TopicType"),14)
		If isNumeric(Form_TopicType) = 0 Then Form_TopicType = 0
		Form_TopicType = cCur(Form_TopicType)
		If Not ((Form_TopicType >=0 and Form_TopicType <=7) or (Form_TopicType>=49 and Form_TopicType<=55)) Then Form_TopicType = 0
		If Form_Topictype_Old = 54 or Form_Topictype_Old = 49 or Form_Topictype_Old = 114 or Form_Topictype_Old = 109 Then Form_Topictype = Form_Topictype_Old
		If Form_TopicType = 55 Then
			Form_NeedValue = Left(GetFormData("Form_NeedValue"),20)
		Else
			If Form_TopicType >=49 and Form_TopicType <=54 Then
				Form_NeedValue = Left(GetFormData("Form_NeedValue"),14)
				If isNumeric(Form_NeedValue) = 0 Then Form_NeedValue = 0
				Form_NeedValue = cCur(Form_NeedValue)
				If Form_NeedValue<0 or Form_NeedValue > 999999 Then Form_NeedValue = 0
			Else
				Form_NeedValue = 0
			End If
		End If
		If Form_TopicType = 7 Then
			If DEF_EnableSpecialTopic = 0 or GetBinarybit(GBL_Board_BoardLimit,14) = 0 Then
				Form_TopicType = 0
				Form_NeedValue = 0
			End If
		End If
	Else
		Form_NeedValue = 0
		If Form_Topictype_Old <> 80 Then Form_TopicType = 0
	End If

	Form_GoodAssort = Left(GetFormData("Form_GoodAssort"),14)
	
	Form_UnderWriteFlag = GetFormData("Form_UnderWriteFlag")
	If Form_UnderWriteFlag="checkbox" Then
		Form_UnderWriteFlag = 1
	Else
		Form_UnderWriteFlag = 0
	End If
	
	If VoteFlag = 1 Then
		Form_VoteItem = Trim(GetFormData("Form_VoteItem"))
		Form_Vote_ExpireDay = Left(Trim(GetFormData("Form_Vote_ExpireDay")),14)
		Form_VoteType = Left(Trim(GetFormData("Form_VoteType")),14)
	End If
	Form_TitleStyle = Left(GetFormData("Form_TitleStyle"),14)

	Form_ForumNumber = Left(GetFormData("ForumNumber"),4)

End Function

Function Borad_CheckAnnounceIDExist(ID)

	Dim Rs
	Set Rs = LDExeCute(sql_select("Select id from LeadBBS_Announce where id=" & ID,1),0)
	If Rs.Eof Then
		Borad_CheckAnnounceIDExist = 0
	Else
		Borad_CheckAnnounceIDExist = 1
	End If
	Rs.Close
	Set Rs = Nothing

End Function

Function DisplayOfficerString(Officer)

	Dim Officer_Temp,Temp_N,dotFlag
	dotFlag = 0
	Officer_Temp = split(Officer,",")
	For Temp_N = 0 to Ubound(Officer_Temp,1)
		If isNumeric(Officer_Temp(Temp_N)) Then
			Officer_Temp(Temp_N) = cCur(Officer_Temp(Temp_N))
			If Officer_Temp(Temp_N)>=0 and Officer_Temp(Temp_N)<=DEF_UserOfficerNum Then
				If dotFlag = 0 Then
					dotFlag = 1
					DisplayOfficerString = DisplayOfficerString & DEF_UserOfficerString(Officer_Temp(Temp_N))
				Else
					DisplayOfficerString = DisplayOfficerString & "," & DEF_UserOfficerString(Officer_Temp(Temp_N))
				End If
			End If
		End If
	Next

End Function

function GetUserID(UserName)

	Dim Rs,SQL
	SQL = sql_select("Select ID from LeadBBS_User Where UserName='" & Replace(username,"'","''") & "'",1)
	Set Rs = LDExeCute(SQL,0)
	If Rs.Eof Then
		GetUserID = 0
	Else
		SQL = Rs(0)
		If isNull(SQL) Then SQL = 0
		GetUserID = cCur(SQL)
	End If
	Rs.Close
	Set Rs = Nothing

End Function

function GetUserName(UserID)

	Dim Rs,SQL
	SQL = sql_select("Select UserName from LeadBBS_User Where ID=" & UserID,1)
	Set Rs = LDExeCute(SQL,0)
	If Rs.Eof Then
		GetUserName = ""
	Else
		SQL = Rs(0)
		If isNull(SQL) Then SQL = 0
		GetUserName = SQL
	End If
	Rs.Close
	Set Rs = Nothing

End Function

Function ResumeCode(Tstr)

	Dim str
	str = Tstr
	Str = Replace(str," &nbsp; &nbsp; &nbsp;",chr(9))
	Str = Replace(str,"<br>" & "&nbsp;",VbCrLf & " ")
	Str = Replace(str,"<br>" & "&nbsp;",VbCrLf & " ")
	Str = Replace(str,"<br>" & VbCrLf,VbCrLf)
	Str = Replace(str,"<br>" & VbCrLf,VbCrLf)
	Str = Replace(str,"<br>",VbCrLf)
	Str = Replace(str,"<br>",VbCrLf)
	Str = Replace(str,"&nbsp;"," ")
	str = Replace(str,"&gt;",">")
	Str = Replace(str,"&lt;","<")
	Str = Replace(str,"&quot;","""")
	ResumeCode = Str

End Function

Function Borad_CheckBoardIDExist(ID)

	If isArray(Application(DEF_MasterCookies & "BoardInfo" & ID)) = False Then
		ReloadBoardInfo(ID)
		If isArray(Application(DEF_MasterCookies & "BoardInfo" & ID)) = False Then
			Borad_CheckBoardIDExist = 0
		Else
			Borad_CheckBoardIDExist = 1
		End If
	Else
		Borad_CheckBoardIDExist = 1
	End If

End Function

Dim LMT_GoodAssortIndex
LMT_GoodAssortIndex = -1
Function CheckGoodAssortID(ID)

	Dim TArray,Num,N

	TArray = Application(DEF_MasterCookies & "BoardInfo" & GBL_Board_ID & "_TI")
	If isArray(TArray) = True Then
		Num = Ubound(TArray,2)
		For N = 0 To Num
			If cCur(TArray(0,N)) = ID Then
				CheckGoodAssortID = 1
				LMT_GoodAssortIndex = N
				Exit Function
			End If
		Next
	End If
	TArray = Application(DEF_MasterCookies & "BoardInfo" & 0 & "_TI")
	If isArray(TArray) = True Then
		Num = Ubound(TArray,2)
		For N = 0 To Num
			If cCur(TArray(0,N)) = ID Then
				CheckGoodAssortID = 1
				LMT_GoodAssortIndex = N
				Exit Function
			End If
		Next
	End If
	CheckGoodAssortID = 0

End Function


Sub ChangeGoodAssort(ID,ID2)

	If ID = ID2 Then Exit Sub
	Dim TArray,N,Num,NN,ExitN
	TArray = Application(DEF_MasterCookies & "BoardInfo" & GBL_Board_ID & "_TI")
	If isArray(TArray) = False Then
		'ChangeGoodAssort = 0
		Exit Sub
	End If
	Num = Ubound(TArray,2)
	NN = 0
	ExitN = 2
	If ID = 0 or ID2 = 0 Then ExitN = 1
	For N = 0 To Num
		If ID = cCur(TArray(0,N)) Then
			If cCur(TArray(2,N)) = -1 Then
				TArray(3,N) = 0
				TArray(4,N) = 0
			Else
				TArray(2,N) = cCur(TArray(2,N)) - 1
				TArray(3,N) = 0
				TArray(4,N) = 0
			End If
			NN = NN + 1
			If NN >= ExitN Then Exit For
		End If
		If ID2 = cCur(TArray(0,N)) Then
			If cCur(TArray(2,N)) <> 0 Then
				If cCur(TArray(2,N)) = -1 Then
					TArray(2,N) = 1
					TArray(3,N) = 0
					TArray(4,N) = 0
				Else
					TArray(2,N) = cCur(TArray(2,N)) + 1
					TArray(3,N) = 0
					TArray(4,N) = 0
				End If
			End If
			TArray(2,N) = 0
			NN = NN + 1
			If NN >= ExitN Then Exit For
		End If
	Next
	If NN > 0 Then
		Application.Lock
		Application(DEF_MasterCookies & "BoardInfo" & GBL_Board_ID & "_TI") = TArray
		Application.UnLock
	End If

	If NN < 2 Then
		TArray = Application(DEF_MasterCookies & "BoardInfo" & 0 & "_TI")
		If isArray(TArray) = False Then
			'ChangeGoodAssort = 0
			Exit Sub
		End If
		Num = Ubound(TArray,2)
		NN = 0
		ExitN = 2
		If ID = 0 or ID2 = 0 Then ExitN = 1
		For N = 0 To Num
			If ID = cCur(TArray(0,N)) Then
				If cCur(TArray(2,N)) = -1 Then
					TArray(3,N) = 0
					TArray(4,N) = 0
				Else
					TArray(2,N) = cCur(TArray(2,N)) - 1
					TArray(3,N) = 0
					TArray(4,N) = 0
				End If
				NN = NN + 1
				If NN >= ExitN Then Exit For
			End If
			If ID2 = cCur(TArray(0,N)) Then
				If cCur(TArray(2,N)) <> 0 Then
					If cCur(TArray(2,N)) = -1 Then
						TArray(2,N) = 1
						TArray(3,N) = 0
						TArray(4,N) = 0
					Else
						TArray(2,N) = cCur(TArray(2,N)) + 1
						TArray(3,N) = 0
						TArray(4,N) = 0
					End If
				End If
				TArray(2,N) = 0
				NN = NN + 1
				If NN >= ExitN Then Exit For
			End If
		Next
		If NN > 0 Then
			Application.Lock
			Application(DEF_MasterCookies & "BoardInfo" & 0 & "_TI") = TArray
			Application.UnLock
		End If
	End If

End Sub

Function CheckAnnouceValue

	GBL_CHK_TempStr = ""
	If CheckWriteEventSpace = 0 Then
		GBL_CHK_TempStr = GBL_CHK_TempStr & "�����޸����ϵĹ������ύ��̫Ƶ�����Ժ������ύ! <br>" & VbCrLf
		GBL_CHK_Flag = 0
		Exit Function
	End If

	If DEF_EnableAttestNumber > 2 and (DEF_AttestNumberPoints = 0 or GBL_CHK_Points < DEF_AttestNumberPoints) Then
		If CheckRndNumber = 0 Then
			GBL_CHK_TempStr = "<b><font color=ff0000>��֤����д����!</font></b><br>"
			GBL_CHK_Flag = 0
			Exit Function
		End If
	End If

	If GBL_BoardMasterFlag < 5 or isNumeric(Form_TitleStyle) = 0 then
		Form_TitleStyle = Form_TitleStyle_Old
		If Form_TitleStyle >= 60 Then Form_TitleStyle = Form_TitleStyle - 60
		If Form_TitleStyle <> Form_TitleStyle_Old and (cCur(Form_TitleStyle) + 60) <> cCur(Form_TitleStyle_Old) Then
			Form_TitleStyle = 0
		End If
	Else
		Form_TitleStyle = fix(cCur(Form_TitleStyle))
		If Form_TitleStyle < 0 or Form_TitleStyle > 8 Then Form_TitleStyle = 0
	End If
	If Form_TitleStyle = 1 and GBL_BoardMasterFlag <9 and ((Form_TitleStyle_Old <> 1 and Form_TitleStyle_Old <> 61) or Form_Title <> Form_Title_Old) then Form_TitleStyle = 0
	If Form_TitleStyle_Old >= 60 Then Form_TitleStyle = Form_TitleStyle + 60

	If Form_TitleStyle < 60 and GBL_BoardMasterFlag < 4 and GetBinarybit(GBL_Board_BoardLimit,13) = 1 Then
		Form_TitleStyle = Form_TitleStyle + 60
	End If

	Dim SQL_Temp
	If Form_TitleStyle_Old < 60 and GBL_BoardMasterFlag < 9 and (GetBinarybit(GBL_Board_BoardLimit,13) = 1 or GetBinarybit(GBL_Board_BoardLimit,22) = 1) Then
		SQL_Temp = "Insert into LeadBBS_Assessor(BoardID,Title,UserName,NdateTime,AnnounceID,Content,HTMLFlag,TypeFlag) Values(" & _
				GBL_Board_ID & _
				",'" & Replace(Form_Title,"'","''") & "'" & _
				",'" & Replace(Form_UserName,"'","''") & "'" & _
				"," & GetTimeValue(DEF_Now) & ""
		SQL_Temp = SQL_Temp & "," & Form_EditAnnounceID
		SQL_Temp = SQL_Temp & ",'" & Replace(Replace(Form_Content & "","\" & VbCrLf,"\\" & VbCrLf & VbCrLf),"'","''") & "'"
		SQL_Temp = SQL_Temp & "," & Form_htmlflag
		If GetBinarybit(GBL_Board_BoardLimit,13) = 1 Then
			SQL_Temp = SQL_Temp & ",0"
		Else
			SQL_Temp = SQL_Temp & ",1"
		End If
		SQL_Temp = SQL_Temp & ")"
		CALL LDExeCute(SQL_Temp,1)
	End If

	If cCur(Form_ParentID) <> 0 or (GBL_CHK_CachetValue < LMTDEF_NeedCachetValue and GBL_BoardMasterFlag <= 4) or isNumeric(Form_GoodAssort) = 0 Then
		Form_GoodAssort = Form_GoodAssort_Old
	Else
		Form_GoodAssort = fix(cCur(Form_GoodAssort))
		If Form_GoodAssort <> 0 Then
			If CheckGoodAssortID(Form_GoodAssort) = 0 Then
				GBL_CHK_TempStr = "��������ר��ѡ�����.<br>" & VbCrLf
				CheckAnnouceValue = 0
				Exit Function
			End If
		End If
	End If

	If Form_ParentID = 0 and GetBinarybit(GBL_Board_BoardLimit,23) = 1 and Form_GoodAssort < 1 Then
			GBL_CHK_TempStr = "�˰������ѡ������ר��.<br>" & VbCrLf
			CheckAnnouceValue = 0
			Exit Function
	End If

	If Len(Form_Content)>LMT_MaxTextLength Then
		If (GBL_UserID>0 and SupervisorFlag = 1) Then
			If Len(Form_Content)>LMT_MaxTextLength*4 Then
				GBL_CHK_TempStr = "�����������ݲ��ܳ���" & LMT_MaxTextLength*4 & "�ֽ�.<br>" & VbCrLf
				CheckAnnouceValue = 0
				Exit Function
			End If
		Else
			GBL_CHK_TempStr = "�����������ݲ��ܳ���" & LMT_MaxTextLength & "�ֽ�.<br>" & VbCrLf
			CheckAnnouceValue = 0
			Exit Function
		End If
	End If

	If GBL_UserID<1 Then
		GBL_CHK_TempStr = "������û�������<br>" & VbCrLf
		CheckAnnouceValue = 0
		Exit Function
	End If
	If isNumeric(Form_BoardID)=0 Then
		GBL_CHK_TempStr = "��������һ��������Ҫ�ط���<br>" & VbCrLf
		CheckAnnouceValue = 0
		Exit Function
	End If
	Form_BoardID = cCur(Form_BoardID)
	If Borad_CheckBoardIDExist(Form_BoardID) = 0 Then
		GBL_CHK_TempStr = "��������һ��������Ҫ�ط�.<br>" & VbCrLf
		CheckAnnouceValue = 0
		Exit Function
	End If
	
	If isNumeric(Form_EditAnnounceID)=0 Then Form_EditAnnounceID=0
	Form_EditAnnounceID = cCur(Form_EditAnnounceID)
	If Borad_CheckAnnounceIDExist(Form_EditAnnounceID) = 0 Then
		GBL_CHK_TempStr = "��������Ҫ�༭�����Ӳ����ڣ������Ǹ�ɾ��������ԭ��.<br>" & VbCrLf
		CheckAnnouceValue = 0
		Exit Function
	End If
	
	If Trim(Replace(Replace(Replace(Replace(Form_Title & "","&nbsp;",""),chr(13),""),chr(10),""),chr(0),"")) = "" Then
		GBL_CHK_TempStr = "�������Ʊ�����д.<br>" & VbCrLf
		CheckAnnouceValue = 0
		Exit Function
	End If

	If len(Form_Title)>255 Then
		GBL_CHK_TempStr = "��������̫�����������255�ֽ�.<br>" & VbCrLf
		CheckAnnouceValue = 0
		Exit Function
	End If
	
 	If Trim(Replace(Replace(Form_Content,"&nbsp;",""),VbCrLf,"")) = "" and (Form_ParentID = 0 or LCase(Left(Form_Title,3)) = "re:") Then
		GBL_CHK_TempStr = "������д����������Ϣ.<br>" & VbCrLf
		CheckAnnouceValue = 0
		Exit Function
	ElseIf LMTDEF_MinAnnounceLength > 0 Then
		If (Len(Form_Title) < LMTDEF_MinAnnounceLength or inStr(htmlencode(Form_Title),LMT_TopicNameNoHTML) or LCase(Left(Form_Title,3)) = "re:") Then
			If Form_htmlflag = 2 Then
				If Len(Trim(ResumeCode(Replace(Replace(Replace(Replace(Form_Content,VbCrLf,""),chr(13),""),chr(10),""),chr(0),"")))) < LMTDEF_MinAnnounceLength Then
					GBL_CHK_TempStr = "��������������Ϣ���̡�<br>" & VbCrLf
					CheckAnnouceValue = 0
					Exit Function
				End If
			Else
				If Len(Trim(ResumeCode(Replace(Replace(Replace(Replace(Form_Content,VbCrLf,""),chr(13),""),chr(10),""),chr(0),"")))) < LMTDEF_MinAnnounceLength Then
					GBL_CHK_TempStr = "��������������Ϣ���̡�<br>" & VbCrLf
					CheckAnnouceValue = 0
					Exit Function
				End If
			End If
		End If
	End If

	If Form_TopicType = 54 and Form_NeedValue > LMT_BuyAnnounceMaxPoints Then
		GBL_CHK_TempStr = "���󣬳��������ֻ�ܱ��" & LMT_BuyAnnounceMaxPoints & DEF_PointsName(0) & "��<br>" & VbCrLf
		CheckAnnouceValue = 0
		Exit Function
	End If
	If Form_TopicType = 49 and Form_NeedValue > LMT_BuyAnnounceMaxPoints Then
		GBL_CHK_TempStr = "���󣬳��������ֻ�ܱ��" & LMT_BuyAnnounceMaxPoints & DEF_PointsName(1) & "��<br>" & VbCrLf
		CheckAnnouceValue = 0
		Exit Function
	End If

	Dim TempURL,Loop_N,Temp
	If VoteFlag = 1 Then
		If Replace(Form_VoteItem,VbCrLf,"") = "" Then
			GBL_CHK_TempStr = "����ͶƱѡ�������д.<br>" & VbCrLf
			CheckAnnouceValue = 0
			Exit Function
		End If
		
		Dim Form_VoteItem_Old
		Form_VoteItem_Old = Form_VoteItem
		Temp = 0
		TempURL = Split(Form_VoteItem,VbCrLf)
		Form_VoteItem = ""
		For Loop_N = 0 to Ubound(TempURL,1)
			TempURL(Loop_N) = Trim(TempURL(Loop_N))
			If TempURL(Loop_N) <> "" Then
				If Len(TempURL(Loop_N)) > 48 Then
					GBL_CHK_TempStr = "����ͶƱѡ������̫�������ܳ���24��.<br>" & VbCrLf
					Form_VoteItem = Form_VoteItem_Old
					CheckAnnouceValue = 0
					Exit Function
				ElseIf StrLength(TempURL(Loop_N)) > 48 Then
					GBL_CHK_TempStr = "����ͶƱѡ������̫�������ܳ���24��.<br>" & VbCrLf
					Form_VoteItem = Form_VoteItem_Old
					CheckAnnouceValue = 0
					Exit Function
				End If
				If Temp > 0 Then
					Form_VoteItem = Form_VoteItem & VbCrLf & TempURL(Loop_N)
				Else
					Form_VoteItem = Form_VoteItem & TempURL(Loop_N)
				End If
				Temp = Temp + 1
				If Temp > DEF_VOTE_MaxNum Then
					GBL_CHK_TempStr = "����ͶƱѡ��ܳ���" & DEF_VOTE_MaxNum & "��.<br>" & VbCrLf
					Form_VoteItem = Form_VoteItem_Old
					CheckAnnouceValue = 0
					Exit Function
				End If
			End If
		Next

		If Temp < 2 or Temp < VoteNumber + 1 Then
			GBL_CHK_TempStr = "�޸ĵ�ͶƱѡ��ܼ��٣�ԭ����ͶƱ����" & VoteNumber + 1 & "��ѡ��.<br>" & VbCrLf
			Form_VoteItem = Form_VoteItem_Old
			CheckAnnouceValue = 0
			Exit Function
		End If

		If Left(Form_VoteItem,1) = VbCrLf Then Form_VoteItem = Mid(Form_VoteItem,2)

		If isNumeric(Form_Vote_ExpireDay) = 0 then Form_Vote_ExpireDay = 0
		Form_Vote_ExpireDay = Fix(cCur(Form_Vote_ExpireDay))
		If Form_Vote_ExpireDay < 0 or Form_Vote_ExpireDay > 365 Then
			GBL_CHK_TempStr = "����ͶƱ����ʱ��ѡ�����.<br>" & VbCrLf
			Form_VoteItem = Form_VoteItem_Old
			CheckAnnouceValue = 0
			Exit Function
		End If

		If isNumeric(Form_VoteType) = 0 then Form_VoteType = 0
		Form_VoteType = Fix(cCur(Form_VoteType))
		If Form_VoteType <> 0 and Form_VoteType <> 1 Then
			GBL_CHK_TempStr = "����ͶƱ����ֻ���ǵ�ѡƱ���ѡƱ.<br>" & VbCrLf
			Form_VoteItem = Form_VoteItem_Old
			CheckAnnouceValue = 0
			Exit Function
		End If
		
		If Form_VoteType = 0 and Form_VoteType_Old = 1 Then
			GBL_CHK_TempStr = "����ͶƱ���Ͳ����ɶ�ѡƱ��Ϊ��ѡƱ.<br>" & VbCrLf
			Form_VoteItem = Form_VoteItem_Old
			CheckAnnouceValue = 0
			Exit Function
		End If
	End If

	Form_Title = UBB_FiltrateBadWords(Form_Title)

	'If Left(Form_Title,3) = "Re:" and Form_Title <> "Re:" & LMT_TopicNameNoHTML and Form_ParentID <> 0 Then Form_Title = Mid(Form_Title,4)
	'Form_Title = Replace(Replace(Form_Title,chr(13),""),chr(10),"")

	Form_Length = Len(Form_Content)
	If GBL_Board_ForumPass <> "" or GBL_Board_OtherLimit > 0 or GetBinarybit(GBL_Board_BoardLimit,2) = 1 or GetBinarybit(GBL_Board_BoardLimit,7) = 1 Then
	Else
		If Left(Form_Title,3) = "re:" Then
		
			If Form_HTMLFlag = 2 Then
				GBL_CHK_TempStr = Trim(Left(clearUbbcode(Form_Content),20))
			Else
				GBL_CHK_TempStr = Trim(Left(Form_Content,20))
			End If
			If Form_Length > 20 Then GBL_CHK_TempStr = GBL_CHK_TempStr & "..."
			If Replace(Replace(GBL_CHK_TempStr,chr(13),""),chr(10),"") <> "" Then Form_Title = "re:" & GBL_CHK_TempStr
			GBL_CHK_TempStr = ""
		ElseIf Left(Form_Title,3) = "Re:" Then
			GBL_CHK_TempStr = Trim(Left(Form_Content,20))
			If Form_Length > 20 Then GBL_CHK_TempStr = GBL_CHK_TempStr & "..."
			If Replace(Replace(GBL_CHK_TempStr,chr(13),""),chr(10),"") <> "" Then Form_Title = "re:" & GBL_CHK_TempStr
			GBL_CHK_TempStr = ""
		End If
	End If

	If Form_TopicType = 55 Then
		Form_NeedValue = GetUserID(Form_NeedValue)
		If Form_NeedValue = 0 Then
			GBL_CHK_TempStr = "���������˴���ֻ����ĳ�û��鿴�������û��������ڡ�<br>" & VbCrLf
			Form_NeedValue = Left(Form_NeedValue,20)
			CheckAnnouceValue = 0
			Exit Function
		End If
	End If

	Form_Title = Replace(Replace(Form_Title,chr(13),""),chr(10),"")
	CheckAnnouceValue = 1

End Function

Function SaveEditAnnounceValue

	Dim MaxAnnounceID,MaxRootID
	Form_NoUserUnderWriteFlag = cCur(Form_NoUserUnderWriteFlag)
	Form_UnderWriteFlag = cCur(Form_UnderWriteFlag)
	Dim SQL,Rs
	Dim PollString,PollNum
	Dim New_Form_RootID
	Dim TempURL,Loop_N
	
	Form_Content = UBB_FiltrateBadWords(Form_Content) '���ֹ���
	
	If Form_UpFlag = 1 Then
		Dim Upd_FileInfo,UploadSave
		Set UploadSave = New Upload_Save
		UploadSave.Upload_File
		Upd_FileInfo = UploadSave.Upd_FileInfo
		Upd_ErrInfo = UploadSave.Upd_ErrInfo
	End If

	PollNum = 0
	If Form_htmlflag = 2 and Form_TopicType <> 80 and Form_ParentID = 0 and Form_TopicType <> 54 and Form_TopicType <> 49 Then
		If Upd_FileInfo <> 0 Then
			PollNum = Upd_FileInfo
		End If
		PollString = ",PollNum=" & PollNum
	ElseIf Form_TopicType <> 80 and Form_TopicType <> 54 and Form_TopicType <> 49 Then
		PollString = ",PollNum=0"
	End If

	If Form_TopicType <> 80 and Form_TopicType < 60 and Form_ParentID = 0 and Form_TopicType > 0 Then
		Loop_N = inStr(Form_Content,"[HIDDEN]")
		If Loop_N > 0 Then
			TempURL = inStr(Loop_N,Form_Content,"[/HIDDEN]")
			If TempURL > Loop_N + 9 Then
				Form_TopicType = Form_TopicType + 60
			End If
		End If
	End If

	select case DEF_UsedDataBase
		case 0,2:
			Set Rs = LDExeCute("Select Max(RootID) from LeadBBS_Announce where ParentID = 0 and BoardID=" & GBL_Board_ID & " and RootID<" & DEF_BBS_TOPMinID,0)
		case Else
			Set Rs = LDExeCute("Select Max(RootID) from LeadBBS_Topic where BoardID=" & GBL_Board_ID & " and RootID<" & DEF_BBS_TOPMinID,0)
	End select
	If Rs.Eof Then
		MaxRootID = 1
	Else
		MaxRootID = Rs(0)
		If isNull(MaxRootID) or MaxRootID="" Then MaxRootID=1
		MaxRootID = cCur(MaxRootID)
	End If
	Rs.Close
	Set Rs = Nothing

	New_Form_RootID = Form_RootID

	Rem �ö������ӱ༭һ��, �������
	If Form_ParentID=0 and GBL_BoardMasterFlag >= 5 and Form_RootID > DEF_BBS_TOPMinID Then
		'If Form_RootID < cCur(Application(DEF_MasterCookies & "BoardInfo" & GBL_Board_ID)(11,0)) Then
			select case DEF_UsedDataBase
				case 0,2:
					Set Rs = LDExeCute("Select Max(RootID) from LeadBBS_Announce Where ParentID = 0 and BoardID=" & GBL_Board_ID,0)
				case Else
					Set Rs = LDExeCute("Select Max(RootID) from LeadBBS_Topic Where BoardID=" & GBL_Board_ID,0)
			End select
			If Rs.Eof Then
				New_Form_RootID = DEF_BBS_TOPMinID+1
			Else
				New_Form_RootID = Rs(0)
				If isNull(New_Form_RootID) or New_Form_RootID="" Then New_Form_RootID = DEF_BBS_TOPMinID+1
				New_Form_RootID = cCur(New_Form_RootID)+1
				If New_Form_RootID<DEF_BBS_TOPMinID Then New_Form_RootID=DEF_BBS_TOPMinID
			End If
			Rs.Close
			Set Rs = Nothing
			
			select case DEF_UsedDataBase
				case 0,2:
					SQL = " Update LeadBBS_Announce Set RootID=" & New_Form_RootID & " where ParentID=0 and RootIDBak=" & LMT_RootIDBak
					CALL LDExeCute(SQL,1)
				case Else
					SQL = " Update LeadBBS_Announce Set RootID=" & New_Form_RootID & " where ID=" & LMT_RootIDBak
					CALL LDExeCute(SQL,1)
					SQL = " Update LeadBBS_Topic Set RootID=" & New_Form_RootID & " where ID=" & LMT_RootIDBak
					CALL LDExeCute(SQL,1)
			End select
		'End If
	End If
	If New_Form_RootID<> Form_RootID Then
		UpdateBoardValue(Form_BoardID)
	End If

	SQL = "Update LeadBBS_Announce set Title='" & Replace(Form_Title,"'","''") & "',Content='" & Replace(Form_Content,"'","''") & "',FaceIcon=" & Form_FaceIcon & ",htmlflag=" & Form_htmlflag & ",NotReplay=" & Form_NotReplay & ",UnderWriteFlag=" & Form_UnderWriteFlag &_
	",TopicType=" & Form_TopicType & ",NeedValue=" & Form_NeedValue & ",TitleStyle=" & Form_TitleStyle & ",Length=" & Form_Length & ",GoodAssort=" & Form_GoodAssort & PollString
	If SupervisorFlag = 0 and (Form_UserID <> GBL_UserID or DateDiff("s",RestoreTime(Form_ndatetime),DEF_Now) > DEF_EditAnnounceDelay) Then SQL = SQL & ",OtherInfo='���������" & Replace(LeftTrue(GBL_CHK_User,63),"'","''") & "��" & DEF_Now & "�༭��" & "'"
	If ((Form_TopicType = 54 or Form_TopicType = 49 or Form_TopicType = 114 or Form_TopicType = 109) and (Form_TopicType_Old <> 54 and Form_TopicType_Old <> 49 and Form_TopicType_Old <> 114 and Form_TopicType_Old <> 109)) and Form_ParentID = 0 Then SQL = SQL & ",PollNum=0"
	SQL = SQL & " where ID=" & Form_EditAnnounceID
	If inStr(application(DEF_MasterCookies & "TopAncList"),"," & Form_EditAnnounceID & ",") Then
		UpdateAnnounceApplicationInfo LMT_RootIDBak,2,Form_Title,0,0
		UpdateAnnounceApplicationInfo LMT_RootIDBak,3,Form_FaceIcon,0,0
		UpdateAnnounceApplicationInfo LMT_RootIDBak,14,Form_TopicType,0,0
		UpdateAnnounceApplicationInfo LMT_RootIDBak,16,Form_TitleStyle,0,0
		UpdateAnnounceApplicationInfo LMT_RootIDBak,6,Form_Length,0,0
		If PollString <> "" Then UpdateAnnounceApplicationInfo LMT_RootIDBak,15,PollNum,0,0
	Else
		If inStr(application(DEF_MasterCookies & "TopAncList" & GBL_Board_BoardAssort),"," & Form_EditAnnounceID & ",") Then
			UpdateAnnounceApplicationInfo LMT_RootIDBak,2,Form_Title,0,GBL_Board_BoardAssort
			UpdateAnnounceApplicationInfo LMT_RootIDBak,3,Form_FaceIcon,0,GBL_Board_BoardAssort
			UpdateAnnounceApplicationInfo LMT_RootIDBak,14,Form_TopicType,0,GBL_Board_BoardAssort
			UpdateAnnounceApplicationInfo LMT_RootIDBak,16,Form_TitleStyle,0,GBL_Board_BoardAssort
			UpdateAnnounceApplicationInfo LMT_RootIDBak,6,Form_Length,0,GBL_Board_BoardAssort
			If PollString <> "" Then If PollString <> "" Then UpdateAnnounceApplicationInfo LMT_RootIDBak,15,PollNum,0,GBL_Board_BoardAssort
		End If
	End If

	ChangeGoodAssort Form_GoodAssort_Old,Form_GoodAssort

	'on error resume next
	CALL LDExeCute(SQL,1)

	If Form_ParentID = 0 Then
		If DEF_UsedDataBase = 1 Then
			SQL = "Update LeadBBS_Topic set Title='" & Replace(Form_Title,"'","''") & "',FaceIcon=" & Form_FaceIcon & ",NotReplay=" & Form_NotReplay &_
			",TopicType=" & Form_TopicType & ",NeedValue=" & Form_NeedValue & ",TitleStyle=" & Form_TitleStyle & ",Length=" & Form_Length & ",GoodAssort=" & Form_GoodAssort & PollString
			SQL = SQL & " where ID=" & Form_EditAnnounceID
			CALL LDExeCute(SQL,1)
		End If
		If Form_TitleStyle = 1 Then
			LMT_TopicNameNoHTML = KillHTMLLabel(Form_Title)
		Else
			LMT_TopicNameNoHTML = Form_Title
		End If
		UpdateBoardLastAnnounce
	ElseIf Form_ParentID > 0 and LMT_RootMaxID = Form_EditAnnounceID Then
		SQL = " Update LeadBBS_Announce Set RootID=" & New_Form_RootID & " where ParentID=0 and RootIDBak=" & LMT_RootIDBak
		If left(Form_Title,3) = "re:" Then Form_Title = Mid(Form_Title,4)
		CALL LDExeCute("Update LeadBBS_Announce Set LastInfo='" & Replace(LeftTrue(Form_Title,50),"'","''") & "' where ID=" & LMT_RootIDBak,1)
	End If
	if err Then
		SaveEditAnnounceValue = 0
		GBL_CHK_TempStr = "���󣬷�����̫æ�������ĵ���̫���������ύ��!<br>" & VbCrLf
	Else
		SaveEditAnnounceValue = 1
	End If
	
	Rem ���汣��ͶƱѡ��
	If VoteFlag = 1 Then
		Form_Vote_ExpireDay = cCur(Form_Vote_ExpireDay)
		If Form_Vote_ExpireDay <> 0 Then Form_Vote_ExpireDay = GetTimeValue(DateAdd("d",Form_Vote_ExpireDay,DEF_Now))
		TempURL = Split(Form_VoteItem,VbCrLf)
		For Loop_N = 0 to VoteNumber
			CALL LDExeCute("Update LeadBBS_VoteItem Set VoteType=" & Form_VoteType & ",VoteName='" & Replace(UBB_FiltrateBadWords(TempURL(Loop_N)),"'","''") & "',ExpiresTime=" & Form_Vote_ExpireDay & " where ID=" & VoteGetData(3,Loop_N),1)
		Next
		For Loop_N = VoteNumber + 1 to Ubound(TempURL,1)
			CALL LDExeCute("insert into LeadBBS_VoteItem(AnnounceID,VoteType,VoteName,ExpiresTime) values(" & Form_EditAnnounceID & "," & Form_VoteType & ",'" & Replace(UBB_FiltrateBadWords(TempURL(Loop_N)),"'","''") & "'," & Form_Vote_ExpireDay & ")",1)
		Next
	End If

End Function

sub UpdateBoardLastAnnounce

	Dim Rs,SQL
	Dim LastAnnounceID
	LastAnnounceID = cCur(Application(DEF_MasterCookies & "BoardInfo" & GBL_Board_ID)(19,0))

	If LastAnnounceID = Form_EditAnnounceID or LastAnnounceID = LMT_RootIDBak Then
		CALL LDExeCute("Update LeadBBS_Boards Set LastTopicName='" & Replace(LMT_TopicNameNoHTML,"'","''") & "' where BoardID=" & Form_BoardID,1)
		UpdateBoardApplicationInfo Form_BoardID,LMT_TopicNameNoHTML,20
	End If

End sub

Sub UpdateAnnounceApplicationInfo(AncID,IndexN,Value,tp,tid)

	Dim GetDataTop,AllTopNum,N,Str
	If tid = 0 Then
		Str = ""
	Else
		Str = tid
	End if
	AllTopNum = -1
	GetDataTop = Application(DEF_MasterCookies & "TopAnc" & Str)
	If isArray(GetDataTop) = False Then
		'If GetDataTop <> "yes" Then ReloadTopAnnounceInfo(tid)
		Exit Sub
	Else
		AllTopNum = Ubound(GetDataTop,2)
	End If

	For N = 0 to AllTopNum
		If cCur(AncID) = cCur(GetDataTop(0,N)) Then
			If tp = 1 Then
				GetDataTop(IndexN,N) = cCur(GetDataTop(IndexN,N)) + Value
			Else
				GetDataTop(IndexN,N) = Value
			End If
			Application.Lock
			Application(DEF_MasterCookies & "TopAnc" & Str) = GetDataTop
			Application.UnLock
			Exit Sub
		End If
	Next

End Sub

Function DisplayAnnounceAccessfull

Global_TableHead%>
<div class=contentbox>
		<table border=0 cellpadding=0 cellspacing=0 width="100%" class=tablebox>
		<tr class=tbhead>
			<td><div class=value>�Ѿ��ɹ��༭���桰<%=GBL_Board_BoardName%>���е����ӣ�������ѡ�����²�����</div></td>
		</tr>
		<tr>
			<td class=tdright>
			<br>
			<%If Upd_ErrInfo <> "" Then Response.Write "<font color=Red class=redfont>" & Upd_ErrInfo & "</font><br>"%>
			ҳ�潫��1����Զ����������༭�����ӣ����Լ���ѡ�����²�����<br>
				<script type="text/javascript">
				function a_topage()
				{
					this.location.href = "a.asp?B=<%=GBL_board_ID%>&ID=<%
					Dim Url
					Url = LMT_RootIDBak
					If LMT_RootIDBak <> Form_EditAnnounceID Then Url = Url & "&RID=" & Form_EditAnnounceID & "#F" & Form_EditAnnounceID
					Response.Write Url%>"; 
				}
				setTimeout("a_topage()",1000);
				</script>
				<ul>
					<li><a href=<%=DEF_BBS_HomeUrl%>Boards.asp>������ҳ</a></li>
					<li>����<a href=<%=DEF_BBS_HomeUrl%>b/b.asp?B=<%=GBL_board_ID%>><%=GBL_Board_BoardName%></a>��̳</li>
					<li>��<a href=a.asp?B=<%=GBL_board_ID%>&ID=<%=Url%>>�ձ༭������</a></li>
					<li>��<a href=a.asp?B=<%=GBL_board_ID%>&ID=<%=LMT_RootIDBak%>>�ձ༭������</a></li>
				</ul>
			</td>
		</tr>		
		</table>
</div>
<%
	Global_TableBottom

End Function

Function GetTopicInfo

	If GBL_UserID < 1 Then
		GetTopicInfo = 0
		GBL_CHK_TempStr = "����, �οͲ��ܱ༭����!<br>" & VbCrLf
		Exit Function
	End If

	Dim Rs,SQL

	SQL = "Select RootID,Title,BoardID,RootIDBak,TitleStyle,ParentID,HTMLFlag,Opinion,Content,FaceIcon,ndatetime,Length,UserName,UserID,UnderWriteFlag,NotReplay,TopicType,NeedValue,GoodAssort from LeadBBS_Announce where ID=" & Form_EditAnnounceID
	Set Rs = LDExeCute(SQL,0)
	If Rs.Eof Then
		LMT_RootIDBak = 0
		Rs.Close
		Set Rs = Nothing
		GBL_CHK_TempStr = "����,�����ⲻ����!<br>" & VbCrLf
		Exit Function
	Else
		Form_BoardID = cCur(Rs("BoardID"))
		Form_UserID = cCur(Rs("UserID"))
		If Form_BoardID <> GBL_Board_ID Then
			LMT_RootIDBak = 0
			GBL_CHK_TempStr = "����,�����ⲻ����!<br>" & VbCrLf
			Rs.close
			Set Rs = Nothing
			Exit function
		End If

		If GBL_BoardMasterFlag < 5 and Form_UserID <> GBL_UserID Then
			GetTopicInfo = 0
			GBL_CHK_TempStr = "����, ��û��Ȩ�ޱ༭����!<br>" & VbCrLf
			Rs.Close
			Set Rs = Nothing
			Exit Function
		End If

		Form_ParentID = cCur(Rs("ParentID"))
		Form_RootID = cCur(Rs("RootID"))
		Form_Title = Rs("Title")
		Form_Title_Old = Form_Title
		Form_HTMLFlag = cCur(Rs("HTMLFlag"))

		Form_FaceIcon = cCur(Rs("FaceIcon"))
		Form_ndatetime = cCur(Rs("ndatetime"))
		Form_UserName = Rs("UserName")
		Form_UnderWriteFlag = Rs("UnderWriteFlag")
		Form_NotReplay = Rs("NotReplay")
		Form_Topictype = Rs("TopicType")
		If Form_Topictype <> 80 and Form_Topictype > 60 Then Form_Topictype = Form_Topictype - 60
		Form_Topictype_Old = Form_Topictype
		Form_NeedValue = cCur(Rs("NeedValue"))
		Form_TitleStyle = Rs("TitleStyle")
		Form_TitleStyle_Old = Form_TitleStyle
		If Form_TitleStyle >= 60 Then Form_TitleStyle = Form_TitleStyle - 60
		Form_GoodAssort = cCur(Rs("GoodAssort"))
		Form_GoodAssort_Old = Form_GoodAssort
		If isNull(Form_Topictype) Then Form_Topictype = 0
		If isNull(Form_NeedValue) Then Form_NeedValue = 0
		Form_Content = Rs("Content")
		
		LMT_TopicName = Form_Title
		LMT_RootIDBak = cCur(RS("RootIDBak"))
		LMT_TopicTitleStyle = Rs("TitleStyle")

		If Form_Topictype = 39 Then
			GBL_CHK_TempStr = "�������޷��༭��<br>" & VbCrLf
			GetTopicInfo = 0
			Exit Function
		End If
		If Form_Topictype = 55 and cCur(Form_NeedValue) > 0 Then
			Form_NeedValue = GetUserName(Form_NeedValue)
			If Form_NeedValue <> GBL_CHK_User and GBL_UserID <> Form_UserID and SupervisorFlag = 0 Then
				GBL_CHK_TempStr = "��˽���Ӳ��������˱༭��<br>" & VbCrLf
				Rs.Close
				Set Rs = Nothing
				GetTopicInfo = 0
				Exit Function
			End If
		End If
		If DEF_EditAnnounceExpires > 0 and GBL_BoardMasterFlag < 4 and DateDiff("s",RestoreTime(Form_ndatetime),DEF_Now) > DEF_EditAnnounceExpires Then
			GBL_CHK_TempStr = "�������Ѿ�����������༭��������ޡ�<br>" & VbCrLf
			Rs.Close
			Set Rs = Nothing
			GetTopicInfo = 0
			Exit Function
		End If
		Rs.close
		Set Rs = Nothing
	End If

	If Form_ParentID <> 0 Then
		select case DEF_UsedDataBase
			case 0,2:
				SQL = sql_select("Select Title,BoardID,RootIDBak,TitleStyle,RootMaxID,RootMinID from LeadBBS_Announce where ParentID=0 and RootIDBak=" & LMT_RootIDBak,1)
			case Else
				SQL = sql_select("Select Title,BoardID,ID,TitleStyle,RootMaxID,RootMinID from LeadBBS_Topic where ID=" & LMT_RootIDBak,1)
		End select
		Set Rs = LDExeCute(SQL,0)
		If Rs.Eof Then
			LMT_RootIDBak = 0
			Rs.Close
			Set Rs = Nothing
			GBL_CHK_TempStr = "����,�����ⲻ����!<br>" & VbCrLf
			Exit Function
		End If
		If Form_BoardID <> GBL_Board_ID Then
			LMT_RootIDBak = 0
			GBL_CHK_TempStr = "����,�����ⲻ����!<br>" & VbCrLf
			Rs.close
			Set Rs = Nothing
			Exit function
		Else
			LMT_TopicName = Rs(0)
			LMT_RootIDBak = cCur(RS(2))
			LMT_TopicTitleStyle = Rs(3)
			If isNull(LMT_RootIDBak) then LMT_RootIDBak = 0
			LMT_RootMaxID = cCur(Rs(4))
			LMT_RootMinID = cCur(Rs(5))
			Rs.Close
			Set Rs = Nothing
		End If
	End If

	If GBL_BoardMasterFlag >= 7 and Form_Topictype = 80 and Form_ParentID = 0 Then VoteFlag = 1
	If VoteFlag = 1 Then
		SQL = sql_select("Select VoteName,VoteType,ExpiresTime,ID from LeadBBS_VoteItem where AnnounceID=" & Form_EditAnnounceID,DEF_VOTE_MaxNum)
		Set Rs = LDExeCute(SQL,0)
		If Not Rs.Eof Then VoteGetData = Rs.GetRows(DEF_VOTE_MaxNum)
		Rs.Close
		Set Rs = Nothing
		VoteNumber = 0
		Form_VoteType = 0
		Form_VoteType_Old = 0
		Form_Vote_ExpireDay = 0
		If isArray(VoteGetData) Then
			VoteNumber = Ubound(VoteGetData,2)
			Form_VoteType = ccur(VoteGetData(1,0))
			If Form_VoteType = 1 Then
				Form_VoteType = 1
			Else
				Form_VoteType = 0
			End If
			Form_VoteType_Old = Form_VoteType
			If cCur(VoteGetData(2,0)) = 0 Then
				Form_Vote_ExpireDay = 0
			Else
				Form_Vote_ExpireDay = DateDiff("d",DEF_Now,RestoreTime(VoteGetData(2,0)))
			End If
			Form_VoteItem = VoteGetData(0,0)
			For SQL = 1 to VoteNumber
				Form_VoteItem = Form_VoteItem & VbCrLf & VoteGetData(0,SQL)
			Next
		End If
	End If
	If Form_TitleStyle_Old >= 30 and Form_TitleStyle_Old <= 60 Then GBL_CHK_TempStr = "�������Ѿ���ֹ�༭��<br>" & VbCrLf

End Function

Sub Main

	Get_PublicValue
	initDatabase
	CheckisBoardMaster
	GetTopicInfo
	
	If GetBinarybit(GBL_Board_BoardLimit,16) = 1 Then
		If LMT_DefaultEdit = 1 Then
			LMT_DefaultEdit = 0
		Else
			LMT_DefaultEdit = 1
		End If
	End If

	If LMT_TopicTitleStyle = 1 Then
		LMT_TopicNameNoHTML = KillHTMLLabel(LMT_TopicName)
	Else
		LMT_TopicNameNoHTML = LMT_TopicName
	End If
	Temp = htmlencode(LMT_TopicNameNoHTML)
	Dim Temp
	If Temp = "" Then
		Temp = "�༭����"
	Else
		If strLength(Temp)>DEF_BBS_DisplayTopicLength-5 Then
			LMT_TopicNameNoHTML_Temp = htmlencode(LeftTrue(Temp,DEF_BBS_DisplayTopicLength-8)) & "..."
			Temp = "�༭��" & LMT_TopicNameNoHTML_Temp
		Else
			LMT_TopicNameNoHTML_Temp = htmlencode(Temp)
			Temp = "�༭��" & LMT_TopicNameNoHTML_Temp
		End if
	End If
	BBS_SiteHead DEF_SiteNameString & " - " & KillHTMLLabel(GBL_Board_BoardName) & " - �༭����",GBL_board_ID,"<span class=navigate_string_step>�༭����</span>"
	UpdateOnlineUserAtInfo GBL_board_ID,GBL_Board_BoardName & "��" & Temp
	
	Boards_Body_Head("")
	CheckAccessLimit
	CheckAccessLimit_TimeLimit
	CheckBoardModifyLimit
	CheckUserModifyLimit

	If Form_Submitflag = "" and Form_EditAnnounceID > 0 Then
		%>
		<div class="a_editanc_nav fire">
			<ul>
			<li><a href=../a/a2.asp?B=<%=GBL_board_ID%>><b>��������</b></a></li>
			<li><a href=../b/b.asp?B=<%=GBL_board_ID%>>������</a></li>
			<li><a href=../b/eb.asp?B=<%=GBL_board_ID%>>������</a></li>
			<li><a href="a.asp?B=<%=GBL_board_ID%>&ID=<%=Form_EditAnnounceID%>">��������</a></li>
			</ul>
		</div><%
	End If
	If GBL_CHK_TempStr <> "" Then
		Global_ErrMsg GBL_CHK_TempStr
	Else
		GetAncUploaInfo
		If Form_Submitflag <> "slzOowl_kdO8m610" Then
			DisplayAnnounceForm
		Else
			GetRequestValue
			If CheckAnnouceValue = 1 Then
				If SaveEditAnnounceValue = 1 Then
					If SupervisorFlag = 0 Then
						CALL LDExeCute("Update LeadBBS_User Set LastWriteTime=" & GetTimeValue(DEF_Now) & " where ID=" & GBL_UserID,1)
						UpdateSessionValue 13,GetTimeValue(DEF_Now),0
					End If
					DisplayAnnounceAccessfull
				Else
					Global_ErrMsg GBL_CHK_TempStr
					DisplayAnnounceForm
				End If
			Else
				Global_ErrMsg GBL_CHK_TempStr
				DisplayAnnounceForm
			End If
		End If
	End If
	If Form_UpFlag = 1 Then Set Form_UpClass = Nothing
	closeDataBase
	Boards_Body_Bottom
	SiteBottom

End Sub

Main
%>