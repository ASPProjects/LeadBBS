<!-- #include file=../inc/BBSsetup.asp -->
<!-- #include file=../inc/User_Setup.ASP -->
<!-- #include file=../inc/Ubbcode.asp -->
<!-- #include file=../inc/Board_Popfun.asp -->
<!-- #include file=../inc/Constellation.asp -->
<!-- #include file=inc/MakeAnnounceTop.asp -->
<!-- #include file=../inc/Limit_Fun.asp -->
<%DEF_BBS_HomeUrl = "../"%>
<!-- #include file=inc/Editor_Fun.asp -->
<!-- #include file=../inc/Upload_Fun.asp -->
<!-- #include file=inc/upload1_fun.asp -->
<!-- #include file=../User/inc/Fun_SendMessage.asp -->
<!-- #include file=../User/inc/Bind_Fun.asp -->
<%
Const LMT_EnableOtherGuestName = 0 '������̳�Ƿ�����ʹ��"�ο�"���������
Const LMT_BuyAnnounceMaxPoints = 9 '���������ĵ�������

Dim LMT_EnableUpload
Const LMTDEF_MaxReAnnounce = 1500 '��������ظ������������������������ƲŻ���Ч
Const LMTDEF_MinAnnounceLength = 2 '������Ҫ��������
Const LMTDEF_NotReplyDate = 120 '���ظ�ʱ��������ڶ�������������ֹ�ظ�,�԰�����������Ч
Const LMTDEF_NeedCachetValue = 1 '�趨���������û������Լ�����ר��
Const LMTDEF_ColorSpend = 1 '�趨������ɫ���Ķ��ٲƸ�ֵ
Const LMTDEF_RepostMsg = 2 '�ظ������Ƿ�Ĭ�϶���Ϣ֪ͨ����,0��Ĭ�ϲ�֪ͨ 1.�ظ�ȫ��֪ͨ 2.��������ʱ��֪ͨ

Dim A_NotReplay
Dim Re_ID,Form_TopicSortID,Form_BoardID,Form_RootID
Dim Form_RootMaxID
Dim Form_Layer,Form_Title,Form_Content,Form_FaceIcon,Form_ndatetime,Form_LastTime
Dim Form_Length,Form_UserName,Form_UserID,Form_HTMLFlag,Form_UnderWriteFlag
Dim Form_UserPass,Form_NoUserUnderWriteFlag,Form_NotReplay
Dim Form_AnnounceIsTopFlag,Form_FaceUrl,Form_FaceWidth,Form_FaceHeight
Dim Form_TopicType,Form_NeedValue,Form_GoodAssort
Dim RootTopicType
Dim Form_VoteFlag,Form_VoteItem,Form_Vote_ExpireDay,Form_VoteType,Form_TitleStyle,Form_PollNum
Dim Form_Opinion,Form_Color,Form_ForumNumber
Dim Form_UpClass,Form_UpFlag,Form_Submitflag,Form_Hits
Form_Vote_ExpireDay = 0
Form_VoteType = 0
Form_TitleStyle = 0
Form_NoUserUnderWriteFlag = 1
Form_GoodAssort = 0

Dim LMT_RootID,LMT_TopicName,LMT_ChildNum,LMT_RootIDBak,LMT_TopicTitleStyle,LMT_TopicNameNoHTML,LMT_LastInfo,LMT_TopicNameNoHTML_Temp,Topic_UserName,LMT_ReName
Dim Upd_ErrInfo
Dim Form_EditAnnounceID
Form_EditAnnounceID = 0

LMT_RootIDBak = 0

Form_HTMLFlag = 2

const PageSplitNum = 10

Rem �·������ӵģɣĺ�
Dim NewAnnounceID
NewAnnounceID = 0

Function DisplayAnnounceForm

	Dim Temp
	Temp = GBL_CHK_TempStr
	GBL_CHK_TempStr = ""
	If Re_ID = 0 Then
		CheckBoardAnnounceLimit
	Else
		CheckBoardReAnnounceLimit
	End If
	CheckUserAnnounceLimit
	If GBL_CHK_TempStr <> "" Then Exit Function
	If A_NotReplay = 1 Then Exit Function
%>
<script language=javascript>
<!--
var ValidationPassed = true,submitflag=0;

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
	ValidationPassed = true;
	submit_disable(theform);
}

//-->
</script>
<%If Form_Submitflag = "" and Re_ID > 0 Then%><img src=../images/null.gif height=4 width=2><br><%End If

DisplayPreview

Global_TableHead%>
<div class=contentbox>
<%
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
		<table border=0 cellpadding=0 cellspacing=0 width="100%" class=tablebox>
		<tr class=tbhead>
			<td><div class=value><%
			If cCur(Re_ID)=0 Then
				If Form_VoteFlag <> "" and Re_ID = 0 Then
					Response.Write "������ͶƱ ע��: *Ϊ������"
				Else
					Response.Write "���������� *Ϊ������"
				End If
			Else
				Response.Write LMT_TopicNameNoHTML_Temp
			End If
			If GBL_BoardMasterFlag < 4 and GetBinarybit(GBL_Board_BoardLimit,13) = 1 Then
				Response.Write "<b>�����ڴ˰淢���������Ҫ�ȴ�����Ա��˲���������ʾ��</b>"
			End If%></font></b>
			</td>
		</tr>
		</table>
		<!-- #include file=inc/post_layer.asp -->
		<%If LMT_EnableUpload = 0 Then %>
		<form action=a2.asp method=post id=LeadBBSFm name=LeadBBSFm onsubmit="submitonce(this);return ValidationPassed;">
		<%Else%>
		<form action="a2.asp?dontRequestFormFlag=1" ID=LeadBBSFm name=LeadBBSFm method="post" enctype="multipart/form-data" onsubmit="submitonce(this);if(ValidationPassed)Upl_submit();return ValidationPassed;">
		<%End If%>
		<table border=0 cellpadding=0 cellspacing=0 width="100%" class=tablebox><%
		If GBL_UserID = 0 Then%>
		<tr>
			<td width="<%=DEF_BBS_LeftTDWidth%>" class=tdleft>*�û���</td>
			<td class=tdright>
				<input maxlength=20 name=User type="text" value="<%=htmlencode(Form_UserName)%>" size="20" class='fminpt input_2'> <%
				If GBL_CHK_User = "" Then
					%>û��ע���<a href=../User/<%=DEF_RegisterFile%>>�������ע��</a>���û�<%
				End If
				If GetBinarybit(GBL_Board_BoardLimit,9) = 1 Then%> ���Բ���д<%End If%></td>
                </tr>
        	<tr>
			<td width="<%=DEF_BBS_LeftTDWidth%>" class=tdleft>����</td>
			<td class=tdright>
				<input maxlength=20 type=password name=Pass value="<%'=htmlencode(Form_UserPass)%>" size="20" class='fminpt input_2'><%
				If GetBinarybit(GBL_Board_BoardLimit,9) = 1 Then%> ���Բ���д<%End If%>
			</td>
		</tr><%
		End If%>
		<tr>
			<td width="<%=DEF_BBS_LeftTDWidth%>" class=tdleft>
				<input name=test value="testvalue" type=hidden>
				<%If Re_ID = 0 Then
					%>
					<input id="Form_Color" name="Form_Color" value="<%=Form_Color%>" type="hidden" />
					<div style="float:left">���� <span class="bluefont" title="���вƸ�:<%=GBL_CHK_CharmPoint%>">��<%=LMTDEF_ColorSpend & DEF_PointsName(1)%>��ɫ</span></div>
					<%If GBL_CHK_CharmPoint >= LMTDEF_ColorSpend Then%>
					<div class="layer_item" style="float:left">
						<div id="Form_Color_view" class="layer_item_title color_pannel" style="BACKGROUND-COLOR:<%=Form_Color%>;">ѡ��</div>
					<div class="layer_iteminfo" onclick="this.style.display='none';">
					<ul class="color_list">
					<script type="text/javascript">
					function titlecolor_set(co)
					{
						if(co=="")
						{
							$id("Form_Color").value="";
							$id("Form_Color_view").style.color = "";
							$id("Form_Color_view").style.backgroundColor = "";
							return;
						}
						$id("Form_Color").value="#" + co;
						var r = parseInt(co.substr(0,2),16);
						var g = parseInt(co.substr(2,2),16);
						var b = parseInt(co.substr(4,2),16);
						var tmp = "RGB(" + (256+~r) + "," + (256+~g) + "," + (256+~b) + ")";
						$id("Form_Color_view").style.color = tmp;
						$id("Form_Color_view").style.backgroundColor = "#" + co;
					}
					var Color_n,Color_l,Color_str = "f0f8ff faebd7 00ffff 7fffd4 f0ffff f5f5dc ffe4c4 ffebcd 0000ff 8a2be2 a52a2a deb887 5f9ea0 7fff00 d2691e ff7f50 000000 1e90ff 696969 6495ed fff8dc dc143c 00ffff 00008b 008b8b b8860b a9a9a9 006400 bdb76b 8b008b 556b2f ff8c00 9932cc 8b0000 e9967a 8fbc8f 483d8b 2f4f4f 00ced1 9400d3 ff1493 00bfff b22222 fffaf0 228b22 ff00ff dcdcdc f8f8ff ffd700 daa520 808080 008000 adff2f f0fff0 ff69b4 cd5c5c 4b0082 fffff0 f0e68c e6e6fa fff0f5 7cfc00 fffacd add8e6 f08080 e0ffff fafad2 90ee90 d3d3d3 ffb6c1 ffa07a 20b2aa 87cefa 778899 b0c4de ffffe0 00ff00 32cd32 faf0e6 ff00ff 800000 66cdaa 0000cd ba55d3 9370db 3cb371 7b68ee 00fa9a 48d1cc c71585 191970 f5fffa ffe4e1 ffe4b5 ffdead 000080 fdf5e6 808000 6b8e23 ffa500 ff4500 da70d6 eee8aa 98fb98 afeeee db7093 ffefd5 ffdab9 cd853f ffc0cb dda0dd b0e0e6 800080 ff0000 bc8f8f 4169e1 8b4513 fa8072 f4a460 2e8b57 fff5ee a0522d c0c0c0 87ceeb 6a5acd 708090 fffafa 00ff7f 4682b4 d2b48c 008080 d8bfd8 ff6347 40e0d0 ee82ee f5deb3 ffffff f5f5f5 ffff00 9acd32";
					Color_str=Color_str.split(" ");
					Color_l=Color_str.length;
					for(Color_n=0;Color_n<Color_l;Color_n++)
					if("#"+Color_str[Color_n]=="<%=Form_Color%>")
					document.write("<li style='COLOR: #" + Color_str[Color_n] + "; BACKGROUND-COLOR: #" + Color_str[Color_n] + "' onclick='titlecolor_set(\"" + Color_str[Color_n] + "\")'></li>");
					else
					document.write("<li style='COLOR: #" + Color_str[Color_n] + "; BACKGROUND-COLOR: #" + Color_str[Color_n] + "' onclick='titlecolor_set(\"" + Color_str[Color_n] + "\")'></li>");
					</script>
					<li onclick="titlecolor_set('');" style="width:90%;margin:6px 0px 0px 0px;"><span style="WHITE-SPACE: nowrap;font-size:9pt;">ȡ��ѡ��</span></li>
					</ul>
					</div>
					</div>
					<%
					Else%>
					<div id="Form_Color_view" class="layer_item_title color_pannel" style="BACKGROUND-COLOR:<%=Form_Color%>;"><span class="grayfont">ѡ��</span></div>
					<%
					End If%><%
				Else
					%>���ӱ���<%
				End If%>
				</td>
			<td class=tdright>
				<input name=submitflag value="true" type=hidden>
				<input name=VoteFlag value="<%=htmlencode(Form_VoteFlag)%>" type=hidden>
				<input name=BoardID value="<%=Form_BoardID%>" type=hidden>
				<input name=ID value="<%=Re_ID%>" type=hidden><%If Form_VoteFlag = "" Then%>
				<input maxlength=255 type="text" id=Form_Title name=Form_Title size="49" value="<%
				If (Form_Submitflag = "first" or Form_Submitflag = "") and Re_ID > 0 and Form_Title="" Then
					If Left(LMT_TopicName,3) <> "Re:" Then
						Response.Write "Re:" & htmlencode(LMT_TopicNameNoHTML)
					Else
						Response.Write htmlencode(LMT_TopicNameNoHTML)
					End If
				Else
					Response.Write htmlencode(Form_title)
				End If%>" class='fminpt input_4'><%Else%>
				<input maxlength=255 id=Form_Title name=Form_Title size="35" value="<%=htmlencode(Form_Title)%>" class='fminpt input_4'>
				<%If isNumeric(Form_Vote_ExpireDay) = 0 then Form_Vote_ExpireDay = 0
				Form_Vote_ExpireDay = cCur(Form_Vote_ExpireDay)%>
				<select name="Form_Vote_ExpireDay">
					<option value=0>����ʱ��
					<option value=0<%If Form_Vote_ExpireDay = 0 Then Response.Write " selected"%>>��������
					<option value=1<%If Form_Vote_ExpireDay = 1 Then Response.Write " selected"%>>һ��
					<option value=2<%If Form_Vote_ExpireDay = 2 Then Response.Write " selected"%>>����
					<option value=3<%If Form_Vote_ExpireDay = 3 Then Response.Write " selected"%>>����
					<option value=7<%If Form_Vote_ExpireDay = 7 Then Response.Write " selected"%>>һ��
					<option value=10<%If Form_Vote_ExpireDay = 10 Then Response.Write " selected"%>>ʮ��
					<option value=15<%If Form_Vote_ExpireDay = 15 Then Response.Write " selected"%>>�����
					<option value=20<%If Form_Vote_ExpireDay = 20 Then Response.Write " selected"%>>��ʮ��
					<option value=30<%If Form_Vote_ExpireDay = 30 Then Response.Write " selected"%>>һ����
					<option value=45<%If Form_Vote_ExpireDay = 45 Then Response.Write " selected"%>>һ���°�
					<option value=60<%If Form_Vote_ExpireDay = 60 Then Response.Write " selected"%>>������
					<option value=90<%If Form_Vote_ExpireDay = 90 Then Response.Write " selected"%>>������
					<option value=120<%If Form_Vote_ExpireDay = 120 Then Response.Write " selected"%>>�ĸ���
					<option value=180<%If Form_Vote_ExpireDay = 180 Then Response.Write " selected"%>>������
					<option value=240<%If Form_Vote_ExpireDay = 240 Then Response.Write " selected"%>>�˸���
					<option value=365<%If Form_Vote_ExpireDay = 365 Then Response.Write " selected"%>>һ��
				</select>
				<%End If
				If GBL_BoardMasterFlag >= 5 Then
					If isNumeric(Form_TitleStyle) = 0 then Form_TitleStyle = 0
					Form_TitleStyle = cCur(Form_TitleStyle)
				%>
				<select name="Form_TitleStyle">
					<option value=0<%If Form_TitleStyle = 0 Then Response.Write " selected"%>>��ʽ</option><%If GBL_BoardMasterFlag >= 9 Then%>
					<option value=1<%If Form_TitleStyle = 1 Then Response.Write " selected"%>>HTML</option><%End If%>
					<option value=2<%If Form_TitleStyle = 2 Then Response.Write " selected"%>>��ɫ</option>
					<option value=3<%If Form_TitleStyle = 3 Then Response.Write " selected"%>>��ɫ</option>
					<option value=4<%If Form_TitleStyle = 4 Then Response.Write " selected"%>>��ɫ</option>
					<option value=5<%If Form_TitleStyle = 5 Then Response.Write " selected"%>>����</option>
					<option value=6<%If Form_TitleStyle = 6 Then Response.Write " selected"%>>�غ�</option>
					<option value=7<%If Form_TitleStyle = 7 Then Response.Write " selected"%>>����</option>
					<option value=8<%If Form_TitleStyle = 8 Then Response.Write " selected"%>>����</option>
				</select>
				<%
				End If

				If cCur(Re_ID)=0 and (GBL_CHK_CachetValue >= LMTDEF_NeedCachetValue or GBL_BoardMasterFlag >= 4) Then
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
					<select name="Form_GoodAssort" style="width:74">
					<%
						If isArray(TArray) = True Then
							Response.Write "<Option value=0>ѡ��ר��" & VbCrLf
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
							Response.Write "<Option value=0>=��ר��=" & VbCrLf
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
				End If%> �255��
				</td>
		</tr><%If Form_VoteFlag <> "" Then%>
		<tr>
			<td width="<%=DEF_BBS_LeftTDWidth%>" class=tdleft>*ͶƱѡ��
			<p>ÿ��һ��ͶƱѡ����<%=DEF_VOTE_MaxNum%>��ѡ�ѡ���24�֣��������ϣ����й���
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
				<input name=Form_FaceIcon type=radio value=0>��
				<input name=Form_FaceIcon type=radio value=1<%If Form_FaceIcon=1 Then Response.WRite " CHECKED"%>><img src="../images/<%=GBL_DefineImage%>bf/FACE1.GIF" class=absmiddle>
				<input name=Form_FaceIcon type=radio value=2<%If Form_FaceIcon=2 Then Response.WRite " CHECKED"%>><img src="../images/<%=GBL_DefineImage%>bf/FACE2.GIF" class=absmiddle>
				<input name=Form_FaceIcon type=radio value=3<%If Form_FaceIcon=3 Then Response.WRite " CHECKED"%>><img src="../images/<%=GBL_DefineImage%>bf/FACE3.GIF" class=absmiddle>
				<input name=Form_FaceIcon type=radio value=5<%If Form_FaceIcon=5 Then Response.WRite " CHECKED"%>><img src="../images/<%=GBL_DefineImage%>bf/FACE5.GIF" class=absmiddle>
				<input name=Form_FaceIcon type=radio value=6<%If Form_FaceIcon=6 Then Response.WRite " CHECKED"%>><img src="../images/<%=GBL_DefineImage%>bf/FACE6.GIF" class=absmiddle>
				<input name=Form_FaceIcon type=radio value=7<%If Form_FaceIcon=7 Then Response.WRite " CHECKED"%>><img src="../images/<%=GBL_DefineImage%>bf/FACE7.GIF" class=absmiddle>
				<input name=Form_FaceIcon type=radio value=15<%If Form_FaceIcon=15 Then Response.WRite " CHECKED"%>><img src="../images/<%=GBL_DefineImage%>bf/FACE15.GIF" class=absmiddle>
				<input name=Form_FaceIcon type=radio value=16<%If Form_FaceIcon=16 Then Response.WRite " CHECKED"%>><img src="../images/<%=GBL_DefineImage%>bf/FACE16.GIF" class=absmiddle>
			</td>
		</tr><%
		call DisplayLeadBBSEditor1(Form_HTMLFlag,Form_Content,0,1)
		If Re_ID=0 and Form_VoteFlag = "" Then%>
		<tr>
			<td width="<%=DEF_BBS_LeftTDWidth%>" class=tdleft>���ܱ�������</td>
			<td class=tdright>
				<table border="0" cellspacing="0" cellpadding="0" class="blanktable"><tr><td><select name=Form_TopicType onchange="if(this.value<49)$id('NextContactDateDiv').style.display='none';if(this.value>=49)$id('NextContactDateDiv').style.display='block';">
					<option value="0">��ѡ����������...
					<%If DEF_EnableSpecialTopic = 1 and GetBinarybit(GBL_Board_BoardLimit,14) = 1 Then%><option value="7"<%If Form_TopicType = 7 Then Response.Write " selected"%>>�ظ��������ܲ鿴<%End If%>
					<option value="50"<%If Form_TopicType = 50 Then Response.Write " selected"%>>�鿴������Ҫ�ﵽ<%=DEF_PointsName(0)%>
					<option value="51"<%If Form_TopicType = 51 Then Response.Write " selected"%>>�ظ�������Ҫ�ﵽ<%=DEF_PointsName(0)%>
					<option value="52"<%If Form_TopicType = 52 Then Response.Write " selected"%>>�鿴������Ҫ�ﵽ<%=DEF_PointsName(4)%>
					<option value="53"<%If Form_TopicType = 53 Then Response.Write " selected"%>>�ظ�������Ҫ�ﵽ<%=DEF_PointsName(4)%>
					<option value="55"<%If Form_TopicType = 55 Then Response.Write " selected"%>>ֻ��ָ���û��ܲ鿴������
					<%If DEF_EnableSpecialTopic = 1 and GetBinarybit(GBL_Board_BoardLimit,14) = 1 Then%>
					<option value="54"<%If Form_TopicType = 54 Then Response.Write " selected"%>>���۱���������<%=DEF_PointsName(0)%>
					<option value="49"<%If Form_TopicType = 49 Then Response.Write " selected"%>>���۱���������<%=DEF_PointsName(1)%><%End If%>
					<option value="1"<%If Form_TopicType = 1 Then Response.Write " selected"%>>������<%=DEF_PointsName(8)%>���ܲ鿴
					<option value="2"<%If Form_TopicType = 2 Then Response.Write " selected"%>>������<%=DEF_PointsName(8)%>���ܻظ�
					<option value="3"<%If Form_TopicType = 3 Then Response.Write " selected"%>>��<%=DEF_PointsName(8)%>���ܲ鿴
					<option value="4"<%If Form_TopicType = 4 Then Response.Write " selected"%>>��<%=DEF_PointsName(8)%>���ܻظ�
					<option value="5"<%If Form_TopicType = 5 Then Response.Write " selected"%>>��<%=DEF_PointsName(5)%>���ܲ鿴
					<option value="6"<%If Form_TopicType = 6 Then Response.Write " selected"%>>��<%=DEF_PointsName(5)%>���ܻظ�
					</select></td>
				<td>
				<span name=NextContactDateDiv id=NextContactDateDiv<%If Form_TopicType<49 Then Response.Write " style='display:none'"%>>
					&nbsp; <input name=Form_NeedValue value="<%If cStr(Form_NeedValue) <> "0" Then Response.Write htmlencode(Form_NeedValue)%>" size=10 maxlength=10 class='fminpt input_1'></span>
				</td>
				<td>
					&nbsp; <a href=#icon onclick="if(LeadBBSFm.Form_TopicType.value=='0'){alert('�ڲ������ر�ǩǰ��ѡ����������.');}else{addcontent(0,'HIDDEN','/HIDDEN');}">�������ر�ǩ</a></td></tr></table>
		</tr><%End If%>
		<tr>
			<td class=tdleft valign=top>����ѡ��</td>
			<td class=tdright>
				<label>
				<input class=fmchkbox type="checkbox" name="Form_NoUserUnderWriteFlag" value="checkbox"<%If Form_NoUserUnderWriteFlag=1 Then Response.Write " checked"%>>��ʾǩ��</label>
				<%If Re_ID=0 and GBL_UserID>0 and GBL_BoardMasterFlag >= 5 and GetBinarybit(GBL_CHK_UserLimit,4) = 0 Then%>
				<label>
				<input class=fmchkbox type="checkbox" name="Form_AnnounceIsTop" value="checkbox"<%If Form_AnnounceIsTopFlag=1 Then Response.Write " checked"%>>�����ö�</label><%
				End If%>
				<label>
				<input class=fmchkbox type="checkbox" name="Form_NotReplay" value="checkbox"<%If Form_NotReplay=1 Then Response.Write " checked"%>><%If Re_ID=0 Then
					Response.Write "��������</label>"
				Else
					%>��������</label>
					<label>
					<input class=fmchkbox type="checkbox" name="Form_RepostMsg" value="checkbox"<%
					If GBL_CHK_User = LMT_ReName Then
						Response.Write " disabled=""disabled"""
					ElseIf LMTDEF_RepostMsg=1 or (LMTDEF_RepostMsg=2 and Request.QueryString("repost") = "1") Then
						Response.Write " checked=""checked"""
					End If
					%>>�ظ�����Ϣ֪ͨ</label><%
				End If%>
				- <a href="../User/Help/Ubb.asp?colo" target=_blank>��ɫ��</a>
				<%
				
					Dim ConnetBind
					set ConnetBind = New Connet_Bind
					ConnetBind.BindAnnounceList
					set ConnetBind = nothing
				%>
				<%If DEF_EnableAttestNumber > 2 and (DEF_AttestNumberPoints = 0 or GBL_CHK_Points < DEF_AttestNumberPoints) Then%>
				<div style="line-height:400%">��֤��
				<%
					displayVerifycode%></div><%
				End If%></div>
			</td>
		</tr>
		<tr>
			<td class=tdleft>
			<td class=tdright>
				<br />
				<input name=submit2 type=submit value="��������" class="fmbtn btn_3">
				<input id=Preview_btn type=button value="Ԥ������" onclick="edt_preview();" class="fmbtn btn_3">
				<br/ ><br />
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

Sub GetBoardID

	If dontRequestFormFlag = "" Then
		Form_UpFlag = 0
	Else
		Form_UpFlag = 1
		Server.ScriptTimeOut=3000
		set Form_UpClass=new upload_Class
		Form_UpClass.ProgressID = Request.QueryString("Upload_ID")
		Form_UpClass.GetUpFile
	End If
	If GBL_Board_ID = 0 Then
		GBL_Board_ID = GetFormData("BoardID")
		If GBL_Board_ID = "" Then GBL_Board_ID = GetFormData("b")
		If GBL_Board_ID = "" Then GBL_Board_ID = Request.QueryString("b")
		GBL_Board_ID = Left(GBL_Board_ID,14)
		If isNumeric(GBL_Board_ID)=0 Then GBL_Board_ID=0
		GBL_Board_ID = Fix(cCur(GBL_Board_ID))
		If GBL_Board_ID > 2147479999 Then GBL_Board_ID = 0
	End If

End Sub

Function GetRequestValue
	
	Form_Submitflag = Request.QueryString("submitflag")
	If Form_Submitflag = "" Then Form_Submitflag = GetFormData("submitflag")

	If cStr(Re_ID) = "" Then Re_ID = Left(Request.QueryString("ID"),14)
	If cStr(Re_ID) = "" Then Re_ID = Left(GetFormData("ID"),14)
	If isNumeric(Re_ID) = 0 Then Re_ID = 0
	Re_ID = Fix(cCur(Re_ID))

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

	If A_NotReplay = 1 Then Exit Function

	If Form_Submitflag = "true" Then
		Form_Title = Trim(GetFormData("Form_Title"))
		Form_HTMLFlag = GetFormData("Form_HTMLFlag")
		If Form_HTMLFlag="2" Then
			Form_HTMLFlag=2
		ElseIf Form_HTMLFlag = "1" and ((GetBinarybit(GBL_CHK_UserLimit,16) = 1 and GBL_BoardMasterFlag >= 2) or CheckSupervisorUserName = 1) and GBL_UserID > 0 Then
			Form_HTMLFlag = 1
		Else
			Form_HTMLFlag = 0
		End If
		'If Re_ID<>0 and Form_Title = "" Then Form_Title="�ظ�:"
	Else
		Form_HTMLFlag = 2
	End If

	If Form_Submitflag <> "first" Then
		Form_Content = GetFormData("Form_Content")
	End If

	Form_FaceIcon = Left(GetFormData("Form_FaceIcon"),14)
	If isNumeric(Form_FaceIcon) = 0 Then Form_FaceIcon = 0
	Form_FaceIcon = Fix(cCur(Form_FaceIcon))
	If Form_FaceIcon < 0 or Form_FaceIcon > 16 Then Form_FaceIcon = 0

	Form_UserID = GBL_UserID

	Form_UserName = GBL_CHK_User
	'Form_UserPass = GBL_CHK_pass
	
	If Form_Submitflag <> "" and Form_Submitflag <> "first" Then
		Form_NoUserUnderWriteFlag = GetFormData("Form_NoUserUnderWriteFlag")
		If Form_NoUserUnderWriteFlag="checkbox" Then
			Form_NoUserUnderWriteFlag = 1
		Else
			Form_NoUserUnderWriteFlag = 0
		End If
	
		Form_AnnounceIsTopFlag = GetFormData("Form_AnnounceIsTop")
		If Form_AnnounceIsTopFlag="checkbox" Then
			Form_AnnounceIsTopFlag = 1
		Else
			Form_AnnounceIsTopFlag = 0
		End If
		
		Form_NotReplay = GetFormData("Form_NotReplay")
		If Form_NotReplay <> "" Then 
			Form_NotReplay = 1
		Else
			Form_NotReplay = 0
		End If
	End If

	Form_VoteFlag = Request.QueryString("VoteFlag")
	If Form_VoteFlag = "" Then Form_VoteFlag = Left(GetFormData("VoteFlag"),14)

	If IsNull(Form_TopicType) Then Form_TopicType = 0
	If IsNull(Form_NeedValue) Then Form_NeedValue = 0
	If Re_ID=0 and Form_VoteFlag = "" and Form_TopicType <> 80 Then
		Form_TopicType = Left(GetFormData("Form_TopicType"),14)
		If isNumeric(Form_TopicType) = 0 Then Form_TopicType = 0
		Form_TopicType = cCur(Form_TopicType)
		If Not ((Form_TopicType >=0 and Form_TopicType <=7) or (Form_TopicType>=49 and Form_TopicType<=55)) Then Form_TopicType = 0
		If Form_TopicType = 55 Then
			Form_NeedValue = Left(GetFormData("Form_NeedValue"),20)
		Else
			If Form_TopicType >=49 and Form_TopicType <=54 Then
				Form_NeedValue = Left(GetFormData("Form_NeedValue"),14)
				If isNumeric(Form_NeedValue) = 0 Then Form_NeedValue = 0
				Form_NeedValue = cCur(Form_NeedValue)
				If Form_NeedValue<0 or Form_NeedValue > 999999 Then Form_NeedValue = 0
				If Form_NeedValue = 0 Then Form_TopicType = 0
			Else
				Form_NeedValue = 0
			End If
		End If
		If Form_TopicType = 54 or Form_TopicType = 49 or Form_TopicType = 7 Then
			If DEF_EnableSpecialTopic = 0 or GetBinarybit(GBL_Board_BoardLimit,14) = 0 Then
				Form_TopicType = 0
				Form_NeedValue = 0
			End If
		End If
	Else
		If Form_VoteFlag <> "" and Re_ID=0 Then
			Form_TopicType = 80
		Else
			Form_TopicType = 0
			Form_NeedValue = 0
		End If
	End If
	
	If Form_VoteFlag <> "" and Re_ID=0 Then
		Form_VoteItem = Trim(GetFormData("Form_VoteItem"))
		Form_Vote_ExpireDay = Left(Trim(GetFormData("Form_Vote_ExpireDay")),14)
		Form_VoteType = Left(Trim(GetFormData("Form_VoteType")),14)
	End If
	Form_TitleStyle = Left(GetFormData("Form_TitleStyle"),14)
	Form_GoodAssort = Left(GetFormData("Form_GoodAssort"),14)
	Form_Color = LeftTrue(GetFormData("Form_Color"),7)
	If Form_Color = "--" then Form_Color = ""

	Form_ForumNumber = Left(GetFormData("ForumNumber"),4)

End Function

Function Borad_CheckAnnounceIDExist(ID)

	Borad_CheckAnnounceIDExist = 1
	exit function
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

Function Borad_CheckBoardIDExist(ID)

	Dim Temp
	Temp = Application(DEF_MasterCookies & "BoardInfo" & ID)
	If isArray(Temp) = False Then
		ReloadBoardInfo(ID)
		Temp = Application(DEF_MasterCookies & "BoardInfo" & ID)
	End If
	If isArray(Temp) = False Then
		Borad_CheckBoardIDExist = 0
	Else
		Borad_CheckBoardIDExist = 1
	End If

End Function


Rem ���ĳ�û����Ƿ����
Function CheckUserNameExist(UserName)

	Dim Rs
	Set Rs=LDExeCute(sql_select("Select ID from LeadBBS_User where UserName='" & Replace(UserName,"'","''") & "'",1),0)
	If Rs.Eof Then
		CheckUserNameExist = 0
	Else
		CheckUserNameExist = 1
	End if
	Rs.Close
	Set Rs = Nothing

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

Sub ChangeGoodAssort(ID)

	If ID = 0 Then Exit Sub
	Dim TArray,Temp
	Temp = GBL_Board_ID
	TArray = Application(DEF_MasterCookies & "BoardInfo" & GBL_Board_ID & "_TI")
	If isArray(TArray) = False Then
		'ChangeGoodAssort = 0
		Exit Sub
	End If
	If ID <> cCur(TArray(0,LMT_GoodAssortIndex)) Then		
		Temp = 0
		TArray = Application(DEF_MasterCookies & "BoardInfo" & 0 & "_TI")
		If isArray(TArray) = False Then
			'ChangeGoodAssort = 0
			Exit Sub
		End If
	End If
	If cCur(TArray(2,LMT_GoodAssortIndex)) = -1 Then
		TArray(2,LMT_GoodAssortIndex) = 1
		TArray(3,LMT_GoodAssortIndex) = 0
		TArray(4,LMT_GoodAssortIndex) = 0
	Else
		TArray(2,LMT_GoodAssortIndex) = cCur(TArray(2,LMT_GoodAssortIndex)) + 1
		TArray(3,LMT_GoodAssortIndex) = 0
		TArray(4,LMT_GoodAssortIndex) = 0
	End If
	Application.Lock
	Application(DEF_MasterCookies & "BoardInfo" & Temp & "_TI") = TArray
	Application.UnLock

End Sub

Function CheckAnnouceValue

	If A_NotReplay = 1 Then Exit Function

	GBL_CHK_TempStr = ""

	If DEF_EnableAttestNumber > 2 and (DEF_AttestNumberPoints = 0 or GBL_CHK_Points < DEF_AttestNumberPoints) Then
		If CheckRndNumber = 0 Then
			GBL_CHK_TempStr = "<b><font color=ff0000>��֤����д����!</font></b><br>"
			GBL_CHK_Flag = 0
			Exit Function
		End If
	End If
	If GBL_BoardMasterFlag < 5 or isNumeric(Form_TitleStyle) = 0 then
		Form_TitleStyle = 0
	Else
		Form_TitleStyle = fix(cCur(Form_TitleStyle))
		If Form_TitleStyle < 0 or Form_TitleStyle > 8 Then Form_TitleStyle = 0
		If Form_TitleStyle = 1 and GBL_BoardMasterFlag <9 then Form_TitleStyle = 0
	End If

	If cCur(Re_ID) <> 0 or (GBL_CHK_CachetValue < LMTDEF_NeedCachetValue and GBL_BoardMasterFlag <= 4) or isNumeric(Form_GoodAssort) = 0 Then
		Form_GoodAssort = 0
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

	If Re_ID = 0 and GetBinarybit(GBL_Board_BoardLimit,23) = 1 and Form_GoodAssort < 1 Then
			GBL_CHK_TempStr = "�˰������ѡ������ר��.<br>" & VbCrLf
			CheckAnnouceValue = 0
			Exit Function
	End If
		

	If Len(Form_Content)>LMT_MaxTextLength Then
		GBL_CHK_TempStr = "�����������ݲ��ܳ���" & DEF_MaxTextLength & "�ֽ�.<br>" & VbCrLf
		CheckAnnouceValue = 0
		Exit Function
	End If

	If GetBinarybit(GBL_Board_BoardLimit,9) = 0 Then
		If GBL_UserID<1 Then
			GBL_CHK_TempStr = "������û������󣬻���û��Ѿ�����������ʱ���Ρ�<br>" & VbCrLf
			CheckAnnouceValue = 0
			Exit Function
		End If
		If GBL_CHK_OnlineTime < DEF_NeedOnlineTime and DEF_NeedOnlineTime > 0 and CheckSupervisorUserName = 0 Then
			GBL_CHK_TempStr = "��̳��������ʱ��(" & DEF_PointsName(4) & ")" & Fix(DEF_NeedOnlineTime/60) & "���������û����ܷ��ԡ�<br>" & VbCrLf
			CheckAnnouceValue = 0
			Exit Function
		End If
	Else
		If GBL_UserID <= 0 and Form_UserName <> "" Then
			If CheckUserNameExist(Form_UserName) = 1 Then
				GBL_CHK_TempStr = "ע�⣺�û���" & htmlencode(Form_UserName) & "�Ѿ�����ʹ�ã��벻Ҫʹ�ô��û���������<br>" & VbCrLf
				CheckAnnouceValue = 0
				Exit Function
			End If
		End If
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

	If isNumeric(Re_ID)=0 Then Re_ID=0
	Re_ID = cCur(Re_ID)
	If Re_ID<>0 Then
		If Borad_CheckAnnounceIDExist(Re_ID) = 0 Then
			GBL_CHK_TempStr = "��������Ҫ�ظ������Ӳ����ڣ������Ǹ�ɾ��������ԭ��.<br>" & VbCrLf
			CheckAnnouceValue = 0
			Exit Function
		End If
	End If
	
	If Trim(Replace(Replace(Replace(Replace(Form_Title & "","&nbsp;",""),chr(13),""),chr(10),""),chr(0),"")) = "" Then
		GBL_CHK_TempStr = "�������Ʊ�����д.<br>" & VbCrLf
		CheckAnnouceValue = 0
		Exit Function
	End If

	If strLength(Form_Title)>255 Then
		GBL_CHK_TempStr = "��������̫�����������255�ֽ�.<br>" & VbCrLf
		CheckAnnouceValue = 0
		Exit Function
	End If

 	If Trim(Replace(Replace(Form_Content,"&nbsp;",""),VbCrLf,"")) = "" and (Re_ID = 0 or (htmlencode(Form_Title) = LMT_TopicNameNoHTML or Lcase(left(Form_Title,3)) = "re:")) Then
		GBL_CHK_TempStr = "������д����������Ϣ.<br>" & VbCrLf
		CheckAnnouceValue = 0
		Exit Function
	ElseIf LMTDEF_MinAnnounceLength > 0 Then
		If (Len(Form_Title) < LMTDEF_MinAnnounceLength or inStr(htmlencode(Form_Title),LMT_TopicNameNoHTML) or Lcase(left(Form_Title,3)) = "re:") Then
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
	If Form_VoteFlag <> "" and Re_ID=0 Then
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
				If StrLength(TempURL(Loop_N)) > 48 Then
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
		
		If Temp < 2 Then
			GBL_CHK_TempStr = "��Ȼ��ͶƱ���벻Ҫ��������ѡ��.<br>" & VbCrLf
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
	End If

	Loop_N = CheckIsRestSpaceTime(Form_Title)
	Select Case Loop_N
	Case 1: If CheckSupervisorUserName = 0 Then
			GBL_CHK_TempStr = GBL_CHK_TempStr & "����������̫������ӣ�����Ϣ" & DEF_RestSpaceTime & "���Ӻ��ٷ���!<br>"
			GBL_CHK_Flag = 0
			Exit Function
		End If
	Case 2:	GBL_CHK_TempStr = GBL_CHK_TempStr & "�벻Ҫ���ظ�������!<br>"
		GBL_CHK_Flag = 0
		Exit Function
	Case 3: Exit Function
	End Select

	If Re_ID=0 and GBL_UserID>0 and (GBL_BoardMasterFlag >= 5) and Form_AnnounceIsTopFlag=1 and GetBinarybit(GBL_CHK_UserLimit,4) = 0 Then
		If CheckMakeTopAnnounceOver = 1 then
			GBL_CHK_TempStr = GBL_CHK_TempStr & "����,�ö�������̫�࣬�����ٷ���ֱ���ö�������.<br>"
			GBL_CHK_Flag = 0
			Exit Function
		End If
	End If

	Form_Title = UBB_FiltrateBadWords(Form_Title)

	If Re_ID = 0 Then
		LMT_LastInfo = ""
	Else
		If Form_Title = "Re:" & LMT_TopicNameNoHTML Then
			If Form_HTMLFlag = 2 Then
				LMT_LastInfo = clearUbbcode(Form_Content)
			Else
				LMT_LastInfo = Form_Content
			End If
		Else
			If Form_TitleStyle = 1 Then
				LMT_LastInfo = KillHTMLLabel(Form_Title)
			Else
				LMT_LastInfo = Form_Title
			End if
		End If
		If StrLength(LMT_LastInfo) > 50 Then LMT_LastInfo = LeftTrue(LMT_LastInfo,47) & "..."
		LMT_LastInfo = UBB_FiltrateBadWords(LMT_LastInfo)
	End If

	Form_Length = Len(Form_Content)
	If Left(Form_Title,3) = "Re:" and Form_Title <> "Re:" & LMT_TopicNameNoHTML and Re_ID <> 0 Then Form_Title = Mid(Form_Title,4)
	'If GBL_Board_ForumPass <> "" or GBL_Board_OtherLimit > 0 or GetBinarybit(GBL_Board_BoardLimit,2) = 1 or GetBinarybit(GBL_Board_BoardLimit,7) = 1 Then
	'���ư��棬�ظ�����ͬ����ı���
	'Else
		If Left(Form_Title,3) = "Re:" Then
			If Form_HTMLFlag = 2 Then
				GBL_CHK_TempStr = Trim(Left(clearUbbcode(Form_Content),20))
			Else
				GBL_CHK_TempStr = Trim(Left(Form_Content,20))
			End If
			If Form_Length > 20 Then GBL_CHK_TempStr = GBL_CHK_TempStr & "..."
			If Replace(Replace(GBL_CHK_TempStr,chr(13),""),chr(10),"") <> "" Then Form_Title = "re:" & GBL_CHK_TempStr
			GBL_CHK_TempStr = ""
		End If
	'End If
	Form_Title = Replace(Replace(Form_Title,chr(13),""),chr(10),"")
	
	If Form_TitleStyle = 0 and Re_ID = 0 and Len(Form_Color) = 7 and GBL_CHK_CharmPoint >= LMTDEF_ColorSpend Then
		Temp = "<font color=""" & htmlencode(Form_Color) & """>" & htmlencode(Form_Title) & "</font>"
		If strLength(Temp)>255 Then
			GBL_CHK_TempStr = "������������̫��.<br>" & VbCrLf
			CheckAnnouceValue = 0
			Exit Function
		End If
		Form_Title = Temp
		Form_TitleStyle = 1
	Else
		Form_Color = ""
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

	CheckAnnouceValue = 1

End Function

Function CheckMakeTopAnnounceOver

	Dim Rs,SQL
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	
	select case DEF_UsedDataBase
		case 0,2:
			SQL = "Select count(*) from LeadBBS_Announce Where ParentID = 0 and BoardID=" & GBL_board_ID & " and RootID>=" & DEF_BBS_TOPMinID
		case Else
			SQL = "Select count(*) from LeadBBS_Topic Where BoardID=" & GBL_board_ID & " and RootID>=" & DEF_BBS_TOPMinID
	End select
	Set Rs = LDExeCute(SQL,0)
	If Not Rs.Eof Then
		SQL = Rs(0)
		If isNull(SQL) Then SQL = 0
		SQL = cCur(SQL)
	Else
		SQL = 0
	End If
	Rs.Close
	Set Rs = Nothing

	If SQL > DEF_BBS_MaxTopAnnounce Then
		CheckMakeTopAnnounceOver = 1
	Else
		CheckMakeTopAnnounceOver = 0
	End If

End Function

Function CheckIsRestSpaceTime(Form_Title)

	If CheckWriteEventSpace = 0 and GBL_CHK_User <> "" Then
		CheckIsRestSpaceTime = 3
		Exit Function
	End If
	Dim Rs,ndatetime,Temp_ID,Temp_Title,Temp_Content
	If GBL_CHK_LastAnnounceID > 0 and GetBinarybit(GBL_Board_BoardLimit,9) = 0 or GBL_UserID > 0 Then
		Set Rs = LDExeCute(sql_select("Select ndatetime,ID,title,Content,htmlflag from LeadBBS_Announce where ID=" & GBL_CHK_LastAnnounceID,1),0)
		If Rs.Eof Then
			CheckIsRestSpaceTime = 0
			Rs.Close
			Set Rs = Nothing
			Temp_ID = 0
		Else
			ndatetime = Rs(0)
			Temp_ID = cCur(Rs(1))
			Temp_Title = Rs(2)
			Temp_Content = Rs(3)
			Rs.Close
			Set Rs = Nothing
		End if
	Else
		Set Rs = LDExeCute(sql_select("Select ndatetime,ID,title,Content,htmlflag from LeadBBS_Announce where IPAddress='" & Replace(GBL_IPAddress,"'","''") & "' Order by ID DESC",1),0)
		If Rs.Eof Then
			CheckIsRestSpaceTime = 0
			Rs.Close
			Set Rs = Nothing
			Exit Function
		Else
			ndatetime = Rs(0)
			Temp_ID = cCur(Rs(1))
			Temp_Title = Rs(2)
			Temp_Content = Rs(3)
			Rs.Close
			Set Rs = Nothing
		End if
	End If

	If Temp_ID = 0 Then
		Exit Function
	End If

	If DateDiff("s", RestoreTime(ndatetime), DEF_Now) < 0 Then
		CheckIsRestSpaceTime = 0
		Exit Function
	End If

	If DateDiff("s", RestoreTime(ndatetime), DEF_Now) < DEF_RestSpaceTime and CheckSupervisorUserName = 0 Then
		CheckIsRestSpaceTime = 1
		'CALL LDExeCute("Update LeadBBS_Announce set ndatetime=" & GetTimeValue(DEF_Now) & " where id=" & Temp_ID,1)
	Else
		If (Temp_Title = Form_Title or "Re:" & Temp_Title = Form_Title) and Temp_Content = Form_Content Then
			CheckIsRestSpaceTime = 2
		Else
			CheckIsRestSpaceTime = 0
		End If
	End If

End Function

Function SaveAnnounceValue


	if cstr(GBL_ipaddress)="115.221.54.100" then
		Response.Write "<p>Form_UserName:" & Form_UserName
		Response.Write "<p>Form_UserID:" & Form_UserID
	End If

	If Form_UserName = "" Then
		Form_UserName = "�ο�"
		GBL_UserID = 0
		Form_UserID = 0
	ElseIf Form_UserID < 1 and LMT_EnableOtherGuestName = 0 Then
		Form_UserName = "�ο�"
	End If
	if cstr(GBL_ipaddress)="115.221.54.100" then
		Response.Write "<p>Form_UserName:" & Form_UserName
		Response.Write "<p>Form_UserID:" & Form_UserID
	End If
	Form_UserID = cCur(Form_UserID)
	If inStr(Lcase(Form_UserName),"<script") or inStr(Lcase(Form_UserName),"<\script") or inStr(Lcase(Form_UserName),"</script") or inStr(Lcase(Form_UserName),"\") or inStr(Lcase(Form_UserName),">") or inStr(Lcase(Form_UserName),"<") or inStr(Lcase(Form_UserName),"""") or inStr(Lcase(Form_UserName),",") or inStr(Lcase(Form_UserName),chr(10)) or inStr(Lcase(Form_UserName),chr(13)) Then Form_UserName = "�ο�"
	Form_NoUserUnderWriteFlag = cCur(Form_NoUserUnderWriteFlag)
	Form_UnderWriteFlag = cCur(Form_UnderWriteFlag)
	Dim SQL,Rs,MaxRootID
	Dim NewMaxAnnounceID
	Dim TempURL,Loop_N,Temp
	Rem Ϊ�˼��ݣ���ü���ѡȡ���RootIDֵ�����RootID����MaxID����RootIDȡ��MaxID(RootID<DEF_BBS_TOPMinID)
	
	If Re_ID = 0 or DEF_EnableMakeTopAnc <> GetBinarybit(GBL_Board_BoardLimit,17) Then
		'MaxRootID = cCur(Application(DEF_MasterCookies & "BoardInfo" & GBL_board_ID)(11,0))
		'If MaxRootID >= DEF_BBS_TOPMinID Then
			select case DEF_UsedDataBase
				case 0,2:
					Set Rs = LDExeCute(sql_select("Select RootID from LeadBBS_Announce where ParentID=0 and BoardID=" & GBL_Board_ID & " Order by RootID DESC",DEF_BBS_MaxTopAnnounce + 2),0)
				case Else
					Set Rs = LDExeCute(sql_select("Select RootID from LeadBBS_Topic where BoardID=" & GBL_Board_ID & " Order by RootID DESC",DEF_BBS_MaxTopAnnounce + 2),0)
			End select
			MaxRootID = 0
			Do while Not Rs.Eof
				If cCur(Rs(0)) < DEF_BBS_TOPMinID Then
					MaxRootID = cCur(Rs(0))
					Exit Do
				End If
				Rs.MoveNext
			Loop
			Rs.Close
			Set Rs = Nothing
			
			If MaxRootID >= DEF_BBS_TOPMinID and MaxRootID > 0 Then
				select case DEF_UsedDataBase
					case 0,2:
						Set Rs = LDExeCute("Select Max(RootID) from LeadBBS_Announce where ParentID=0 and BoardID=" & GBL_Board_ID & " and RootID<" & DEF_BBS_TOPMinID,0)
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
			End If
		'End If
	End If

	If Re_ID<>0 Then
		If DEF_EnableTreeView = 1 Then
			Set Rs = LDExeCute(sql_select("Select RootIDBak,Layer,TopicSortID from LeadBBS_Announce where id=" & Re_ID,1),0)
			If Rs.Eof Then
				SaveAnnounceValue = 0
				Rs.Close
				Set Rs = Nothing
				GBL_CHK_TempStr = "�������Ҫ�ظ��������Ҳ���.<br>" & VbCrLf
				Exit Function
			Else
				Form_RootID = cCur(Rs("RootIDBak"))
				Form_Layer = Rs("Layer") + 1
				Form_TopicSortID = cCur(Rs("TopicSortID"))+1
				Rs.Close
				Set Rs = Nothing
			End If
		Else
			Form_RootID = 0
			Form_Layer = 2
			Form_TopicSortID = Form_RootMaxID + 1
		End If

		If Form_Layer>2 Then
			'CALL LDExeCute(" Update LeadBBS_Announce Set ChildNum=ChildNum+1,LastTime=" & GetTimeValue(DEF_Now) & ",LastUser='" & Replace(Form_UserName,"'","''") & "' where ID=" & LMT_RootIDBak,1)
			CALL LDExeCute(" Update LeadBBS_Announce Set ChildNum=ChildNum+1 where id=" & Re_ID,1)
		Else
			'If DEF_EnableTreeView = 1 Then
			'	CALL LDExeCute(" Update LeadBBS_Announce Set ChildNum=ChildNum+1,LastTime=" & GetTimeValue(DEF_Now) & ",LastUser='" & Replace(Form_UserName,"'","''") & "' where id=" & LMT_RootIDBak,1)
			'Else
			'	CALL LDExeCute(" Update LeadBBS_Announce Set ChildNum=ChildNum+1,LastTime=" & GetTimeValue(DEF_Now) & ",LastUser='" & Replace(Form_UserName,"'","''") & "' where id=" & Re_ID,1)
			'End If
		End If
		If DEF_EnableTreeView = 1 Then
			CALL LDExeCute("Update LeadBBS_Announce Set TopicSortID=TopicSortID+1 where BoardID=" & Form_boardID & " and RootIDBak=" & LMT_RootIDBak & " and TopicSortID>=" & Form_TopicSortID,1)
		End If
	Else
		If Re_ID=0 and GBL_UserID>0 and (GBL_BoardMasterFlag >= 5) and Form_AnnounceIsTopFlag=1 and GetBinarybit(GBL_CHK_UserLimit,4) = 0 Then
			select case DEF_UsedDataBase
				case 0,2:
					Set Rs = LDExeCute("Select Max(RootID) from LeadBBS_Announce Where ParentID=0 and BoardID=" & GBL_Board_ID & " and RootID>=" & DEF_BBS_TOPMinID,0)
				case Else
					Set Rs = LDExeCute("Select Max(RootID) from LeadBBS_Topic Where BoardID=" & GBL_Board_ID & " and RootID>=" & DEF_BBS_TOPMinID,0)
			End select
			If Rs.Eof Then
				Form_RootID = DEF_BBS_TOPMinID+1
			Else
				Form_RootID = Rs(0)
				If isNull(Form_RootID) or Form_RootID="" Then Form_RootID = DEF_BBS_TOPMinID+1
				Form_RootID = cCur(Form_RootID)+1
				If Form_RootID<DEF_BBS_TOPMinID Then Form_RootID=DEF_BBS_TOPMinID
			End If
			Rs.Close
			Set Rs = Nothing
		Else
			Form_RootID = MaxRootID+1
		End If
		Form_Layer = 1
		Form_TopicSortID = 1
	End If
	Form_ndatetime = GetTimeValue(DEF_Now)
	Form_LastTime = Form_ndatetime
	
	Form_Content = UBB_FiltrateBadWords(Form_Content) '���ֹ���

	If Form_UpFlag = 1 Then
		Dim Upd_FileInfo,UploadSave
		Set UploadSave = New Upload_Save
		UploadSave.Upload_File
		Upd_FileInfo = UploadSave.Upd_FileInfo
		Upd_ErrInfo = UploadSave.Upd_ErrInfo
	End If
	
	Form_PollNum = 0
	If Form_htmlflag = 2 and Form_TopicType <> 80 and Re_ID = 0 and Form_TopicType <> 54 and Form_TopicType <> 49 Then
		If Upd_FileInfo <> 0 Then
			Form_PollNum = Upd_FileInfo
		End If
	End If

	If Form_TopicType <> 80 and Form_TopicType < 60 and Re_ID = 0 and Form_TopicType > 0 Then
		Loop_N = inStr(Form_Content,"[HIDDEN]")
		If Loop_N > 0 Then
			Temp = inStr(Loop_N,Form_Content,"[/HIDDEN]")
			If Temp > Loop_N + 9 Then
				Form_TopicType = Form_TopicType + 60
			End If
		End If
	End If

	If Form_NeedValue = "" then Form_NeedValue = 0
	Form_NeedValue = cCur(Form_NeedValue)

	If Re_ID <> 0 Then
		Form_TopicType = 0
		Form_NeedValue = 0
	End If

	If GBL_BoardMasterFlag < 4 and GetBinarybit(GBL_Board_BoardLimit,13) = 1 Then
		Form_TitleStyle = Form_TitleStyle + 60
	End If
	SQL = " insert into LeadBBS_Announce(ParentID,TopicSortID,BoardID,RootID," & _
		    "Layer,Title,Content,FaceIcon,ndatetime,LastTime,Length," &_
		    "UserName,UserID,UnderWriteFlag,htmlflag,NotReplay,IPAddress,TopicType,NeedValue,TitleStyle,RootIDBak,VisitIP,GoodAssort,PollNum)" &_
	" values(" & Re_ID & "," & Form_TopicSortID & "," & Form_BoardID & "," & Form_RootID & "," &_
	Form_Layer & ",'" & Replace(Form_Title,"'","''") & "','" & Replace(Replace(Form_Content & "","\" & VbCrLf,"\\" & VbCrLf & VbCrLf),"'","''") & "'," &_
	Form_FaceIcon & "," & Form_ndatetime & "," & Form_LastTime & "," & Form_Length & ",'" &_
	Replace(Form_UserName,"'","''") & "'," & Form_UserID & "," & Form_NoUserUnderWriteFlag & "," & Form_htmlflag & "," & Form_NotReplay & ",'" & Replace(GBL_IPAddress,"'","''") & "'" & _
	"," & Form_TopicType & "," & Form_NeedValue & "," & Form_TitleStyle & "," & LMT_RootIDBak & ",'0.0.0.0'," & Form_GoodAssort & "," & Form_PollNum & ")"

	Dim SQL_Temp
	SQL_Temp = "Insert into LeadBBS_Assessor(BoardID,Title,UserName,NdateTime,AnnounceID,Content,HTMLFlag,TypeFlag) Values(" & _
			GBL_Board_ID & _
			",'" & Replace(Form_Title,"'","''") & "'" & _
			",'" & Replace(Form_UserName,"'","''") & "'" & _
			"," & GetTimeValue(DEF_Now) & ""
	ChangeGoodAssort(Form_GoodAssort)

	CALL LDExeCute(SQL,1)
	
	If Form_Color <> "" and Form_TitleStyle = 1 Then
		Form_Color = ",CharmPoint=CharmPoint-" & LMTDEF_ColorSpend
	Else
		Form_Color = ""
	End if
	If GBL_BoardMasterFlag < 4 and GetBinarybit(GBL_Board_BoardLimit,13) = 1 Then
		Form_Title = "���������..."
		LMT_TopicNameNoHTML = "���������..."
		LMT_LastInfo = ""
	End If

	'����Ҫָ��User
	'�˷�����Ȼ�ǳ��죬����������ʱ����������������
	If DEF_UsedDataBase = 0 or DEF_UsedDatabase = 2 Then
		Set Rs = LDExeCute("select @@IDENTITY as id",0)
		NewAnnounceID = Rs(0)
		Rs.Close
		Set Rs = Nothing
		If isNull(NewAnnounceID) Then NewAnnounceID = 0
		NewAnnounceID = cCur(NewAnnounceID)
		If Re_ID = 0 then LMT_RootIDBak = NewAnnounceID

		If NewAnnounceID = 0 Then
			'SQL = sql_select("Select ID,RootID from LeadBBS_Announce where UserID=" & Form_UserID & " and ParentID=" & Re_ID & " order by id DESC",1)
			SQL = sql_select("Select ID,RootID from LeadBBS_Announce where UserID=" & Form_UserID & " order by id DESC",1)
			'SQL = "Select max(ID) from LeadBBS_Announce where UserID=" & Form_UserID
	
			Set Rs = LDExeCute(SQL,0)
			If Rs.Eof Then
				GBL_CHK_TempStr = "������󣬸շ����������Ѿ�ɾ���������������<br>" & VbCrLf
				Rs.Close
				Set Rs = Nothing
				Exit Function
			End If
			NewAnnounceID = Rs(0)
			'Form_RootID = cCur(Rs(1))
			If isNull(NewAnnounceID) Then NewAnnounceID = 0
			NewAnnounceID = cCur(NewAnnounceID)
			If Re_ID = 0 then LMT_RootIDBak = NewAnnounceID
			Rs.Close
			Set Rs = Nothing
		End If
	Else
		SQL = "Select max(ID) from LeadBBS_Announce where UserID=" & Form_UserID
		Set Rs=LDExeCute(SQL,0)
		GBL_DBNum = GBL_DBNum + 1
		If Rs.Eof Then
			GBL_CHK_TempStr = "������󣬸շ����������Ѿ�ɾ���������������<br>" & VbCrLf
			Rs.Close
			Set Rs = Nothing
			Exit Function
		End If
		NewAnnounceID = Rs(0)
		If isNull(NewAnnounceID) Then NewAnnounceID = 0
		NewAnnounceID = cCur(NewAnnounceID)
		If Re_ID = 0 then LMT_RootIDBak = NewAnnounceID
		Rs.Close
		Set Rs = Nothing
		
		If Re_ID = 0 Then
			SQL = " insert into LeadBBS_Topic(ID,BoardID,RootID," & _
				    "Title,FaceIcon,ndatetime,LastTime,Length," &_
				    "UserName,UserID,NotReplay,TopicType,NeedValue,TitleStyle,VisitIP,GoodAssort,PollNum)" &_
			" values(" & NewAnnounceID & "," & Form_BoardID & "," & Form_RootID & "," &_
			"'" & Replace(Form_Title,"'","''") & "'," &_
			Form_FaceIcon & "," & Form_ndatetime & "," & Form_LastTime & "," & Form_Length & ",'" &_
			Replace(Form_UserName,"'","''") & "'," & Form_UserID & "," & Form_NotReplay & "" & _
			"," & Form_TopicType & "," & Form_NeedValue & "," & Form_TitleStyle & ",'0.0.0.0'," & Form_GoodAssort & "," & Form_PollNum & ")"
			CALL LDExeCute(SQL,1)
		End If
	End If
	
	If Form_UpFlag = 1 Then
		UploadSave.UpdateUpload(NewAnnounceID)
		Set UploadSave = Nothing
	End If

	NewMaxAnnounceID = NewAnnounceID
	If GBL_BoardMasterFlag < 9 and (GetBinarybit(GBL_Board_BoardLimit,13) = 1 or GetBinarybit(GBL_Board_BoardLimit,22) = 1) Then
		SQL_Temp = SQL_Temp & "," & NewAnnounceID
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

	Rem ����MaxRootID
	If Re_ID > 0 Then
		If (DEF_EnableMakeTopAnc <> GetBinarybit(GBL_Board_BoardLimit,17)) and cCur(LMT_RootID)<DEF_BBS_TOPMinID Then
			CALL LDExeCute("Update LeadBBS_Announce Set RootID=" & MaxRootID+1 & ",ChildNum=ChildNum+1,LastTime=" & GetTimeValue(DEF_Now) & ",LastUser='" & Replace(Form_UserName,"'","''") & "',RootMaxID=" & NewMaxAnnounceID & ",LastInfo='" & Replace(Left(LMT_LastInfo,50),"'","''") & "' where ID=" & LMT_RootIDBak,1)
			If DEF_UsedDataBase = 1 Then CALL LDExeCute("Update LeadBBS_Topic Set RootID=" & MaxRootID+1 & ",ChildNum=ChildNum+1,LastTime=" & GetTimeValue(DEF_Now) & ",LastUser='" & Replace(Form_UserName,"'","''") & "',RootMaxID=" & NewMaxAnnounceID & ",LastInfo='" & Replace(Left(LMT_LastInfo,50),"'","''") & "' where ID=" & LMT_RootIDBak,1)
			
			If cCur(Application(DEF_MasterCookies & "BoardInfo" & GBL_board_ID)(12,0)) >= cCur(LMT_RootID) Then
				UpdateBoardValue(GBL_board_ID)
			Else
				CALL LDExeCute(" Update LeadBBS_Boards Set AllMaxRootID=" & MaxRootID+1 & ",LastWriter='" & Replace(Form_UserName,"'","''")  &"',LastWriteTime=" & Form_LastTime & ",LastAnnounceID=" & LMT_RootIDBak & ",LastTopicName='" & Replace(LMT_TopicNameNoHTML,"'","''") & "' where BoardID=" & GBL_board_ID,1)
				UpdateBoardApplicationInfo GBL_Board_ID,MaxRootID+1,11
			End If
			'UpdateBoardValue(GBL_board_ID)
		Else
			CALL LDExeCute(" Update LeadBBS_Boards Set LastWriter='" & Replace(Form_UserName,"'","''")  &"',LastWriteTime=" & Form_LastTime & ",LastAnnounceID=" & LMT_RootIDBak & ",LastTopicName='" & Replace(LMT_TopicNameNoHTML,"'","''") & "' where BoardID=" & GBL_board_ID,1)
			CALL LDExeCute("Update LeadBBS_Announce Set ChildNum=ChildNum+1,LastTime=" & GetTimeValue(DEF_Now) & ",LastUser='" & Replace(Form_UserName,"'","''") & "',RootMaxID=" & NewMaxAnnounceID & ",LastInfo='" & Replace(Left(LMT_LastInfo,50),"'","''") & "' where ID=" & LMT_RootIDBak,1)
			If DEF_UsedDataBase = 1 Then CALL LDExeCute("Update LeadBBS_Topic Set ChildNum=ChildNum+1,LastTime=" & GetTimeValue(DEF_Now) & ",LastUser='" & Replace(Form_UserName,"'","''") & "',RootMaxID=" & NewMaxAnnounceID & ",LastInfo='" & Replace(Left(LMT_LastInfo,50),"'","''") & "' where ID=" & LMT_RootIDBak,1)
		End If

		UpdateBoardApplicationInfo GBL_board_ID,Form_UserName,3
		UpdateBoardApplicationInfo GBL_board_ID,Form_LastTime,4
		UpdateBoardApplicationInfo GBL_board_ID,LMT_RootIDBak,19
		UpdateBoardApplicationInfo GBL_board_ID,LMT_TopicNameNoHTML,20
		UpdateBoardAnnounceNum Application(DEF_MasterCookies & "BoardInfo" & GBL_board_ID)(28,0),0,1,1,0,Form_UserName,Form_LastTime,LMT_RootIDBak,LMT_TopicNameNoHTML
		UpdateStatisticDataInfo 1,9,1
		UpdateStatisticDataInfo 1,11,1
		CALL LDExeCute("Update LeadBBS_User set Points=Points+" & DEF_BBS_AnnouncePoints & ",AnnounceNum=AnnounceNum+1,AnnounceNum2=AnnounceNum2+1,LastAnnounceID=" & NewAnnounceID & Form_Color & " Where ID = " & Form_UserID,1)
		UpdateSessionValue 4,DEF_BBS_AnnouncePoints,1
		UpdateSessionValue 14,NewAnnounceID,0
		If Form_Color <> "" Then UpdateSessionValue 15,0-LMTDEF_ColorSpend,1
		If inStr(application(DEF_MasterCookies & "TopAncList"),"," & LMT_RootIDBak & ",") Then
			UpdateAnnounceApplicationInfo LMT_RootIDBak,10,Form_UserName,0,0
			UpdateAnnounceApplicationInfo LMT_RootIDBak,4,Form_LastTime,0,0
			UpdateAnnounceApplicationInfo LMT_RootIDBak,1,LMT_ChildNum + 1,0,0
			UpdateAnnounceApplicationInfo LMT_RootIDBak,17,LMT_LastInfo,0,0
		Else
			If inStr(application(DEF_MasterCookies & "TopAncList" & GBL_Board_BoardAssort),"," & LMT_RootIDBak & ",") Then
				UpdateAnnounceApplicationInfo LMT_RootIDBak,10,Form_UserName,0,GBL_Board_BoardAssort
				UpdateAnnounceApplicationInfo LMT_RootIDBak,4,Form_LastTime,0,GBL_Board_BoardAssort
				UpdateAnnounceApplicationInfo LMT_RootIDBak,1,LMT_ChildNum + 1,0,GBL_Board_BoardAssort
				UpdateAnnounceApplicationInfo LMT_RootIDBak,17,LMT_LastInfo,0,GBL_Board_BoardAssort
			End If
		End If
		
		'ͬʱ�᾵����(���3��)
		SQL = sql_select("Select ID from LeadBBS_Announce where TopicType=39 and NeedValue=" & Re_ID,3)
		Set Rs = LDExeCute(SQL,0)
		If Not Rs.Eof Then
			SQL = Rs.GetRows(3)
			Rs.Close
			Set Rs = Nothing
			For Temp = 0 to Ubound(SQL,2)
				CALL MakeAnnounceTop(SQL(0,Temp),",LastUser='" & Replace(Form_UserName,"'","''") & "',LastTime=" & GetTimeValue(DEF_Now) & ",ChildNum=ChildNum+1,Hits=" & Form_Hits)
			Next
		Else
			Rs.Close
			Set Rs = Nothing
		End If
	Else
		CALL LDExeCute("Update LeadBBS_Announce Set RootMaxID=ID,RootMinID=ID,RootIDBak=ID where RootIDBak=0",1)
		If DEF_UsedDataBase = 1 Then CALL LDExeCute("Update LeadBBS_Topic Set RootMaxID=ID,RootMinID=ID where ID=" & NewAnnounceID,1)
		CALL LDExeCute("Update LeadBBS_User set Points=Points+" & DEF_BBS_AnnouncePoints * 2 & ",AnnounceNum=AnnounceNum+1,AnnounceNum2=AnnounceNum2+1,AnnounceTopic=AnnounceTopic+1,LastAnnounceID=" & NewAnnounceID & Form_Color & " Where ID = " & Form_UserID,1)
		UpdateSessionValue 4,DEF_BBS_AnnouncePoints * 2,1
		UpdateSessionValue 14,NewAnnounceID,0
		If Form_Color <> "" Then UpdateSessionValue 15,0-LMTDEF_ColorSpend,1

		LMT_TopicName = Form_Title
		If Form_TitleStyle <> 1 Then
			LMT_TopicNameNoHTML = Form_Title
		Else
			LMT_TopicNameNoHTML = KillHTMLLabel(Form_Title)
		End If
		
		If Form_RootID > cCur(Application(DEF_MasterCookies & "BoardInfo" & GBL_board_ID)(11,0)) Then
			If cCur(Application(DEF_MasterCookies & "BoardInfo" & GBL_board_ID)(12,0)) = 0 Then
				UpdateBoardValue(GBL_board_ID)
			Else
				CALL LDExeCute("Update LeadBBS_Boards Set AllMaxRootID=" & Form_RootID & ",LastWriter='" & Replace(Form_UserName,"'","''")  &"',LastWriteTime=" & Form_LastTime & ",LastAnnounceID=" & LMT_RootIDBak & ",LastTopicName='" & Replace(LMT_TopicNameNoHTML,"'","''") & "' where BoardID=" & GBL_board_ID,1)
				UpdateBoardApplicationInfo GBL_Board_ID,Form_RootID,11
			End If
		Else
			CALL LDExeCute("Update LeadBBS_Boards Set LastWriter='" & Replace(Form_UserName,"'","''")  &"',LastWriteTime=" & Form_LastTime & ",LastAnnounceID=" & LMT_RootIDBak & ",LastTopicName='" & Replace(LMT_TopicNameNoHTML,"'","''") & "' where BoardID=" & GBL_board_ID,1)
		End If
		UpdateBoardApplicationInfo GBL_board_ID,Form_RootID,11
		UpdateBoardApplicationInfo GBL_board_ID,Form_UserName,3
		UpdateBoardApplicationInfo GBL_board_ID,Form_LastTime,4
		UpdateBoardApplicationInfo GBL_board_ID,LMT_RootIDBak,19
		UpdateBoardApplicationInfo GBL_board_ID,LMT_TopicNameNoHTML,20
		UpdateStatisticDataInfo 1,9,1
		UpdateStatisticDataInfo 1,10,1
		UpdateStatisticDataInfo 1,11,1
		UpdateBoardAnnounceNum Application(DEF_MasterCookies & "BoardInfo" & GBL_board_ID)(28,0),1,1,1,0,Form_UserName,Form_LastTime,LMT_RootIDBak,LMT_TopicNameNoHTML
	End If
	'on error resume next
	If err Then
		SaveAnnounceValue = 0
		GBL_CHK_TempStr = "���󣬷�����̫æ�������ĵ���̫���������ύ��!<br>" & VbCrLf
		err.clear
	Else
		SaveAnnounceValue = 1
	End If

	Rem ���汣��ͶƱѡ��
	If Form_VoteFlag <> "" and Re_ID=0 Then
		Form_Vote_ExpireDay = cCur(Form_Vote_ExpireDay)
		If Form_Vote_ExpireDay <> 0 Then Form_Vote_ExpireDay = GetTimeValue(DateAdd("d",Form_Vote_ExpireDay,DEF_Now))
		Temp = 0
		TempURL = Split(Form_VoteItem,VbCrLf)
		Form_VoteItem = ""
		For Loop_N = 0 to Ubound(TempURL,1)
			If TempURL(Loop_N) <> "" Then
				CALL LDExeCute("insert into LeadBBS_VoteItem(AnnounceID,VoteType,VoteName,ExpiresTime) values(" & NewAnnounceID & "," & Form_VoteType & ",'" & Replace(UBB_FiltrateBadWords(TempURL(Loop_N)),"'","''") & "'," & Form_Vote_ExpireDay & ")",1)
				Temp = Temp + 1
				If Temp > DEF_VOTE_MaxNum Then
					GBL_CHK_TempStr = "����ͶƱѡ��ܳ���" & DEF_VOTE_MaxNum & "��.<br>" & VbCrLf
					SaveAnnounceValue = 0
					Exit Function
				End If
			End If
		Next
	End If
	
	Rem �������Ϣ֪ͨ����
	If LMT_ReName = "�ο�" or LMT_ReName = "[LeadBBS]" Then LMT_ReName = ""
	If Re_ID > 0 and LMT_ReName <> "" Then
		If GetFormData("Form_RepostMsg") = "checkbox" Then
			SendNewMessage GBL_CHK_User,LMT_ReName,"��̳���ţ����ӻظ�֪ͨ","[color=blue]��������������ܵ��ظ�[/color][hr]" &_
			"[b]���ڰ��棺[/b][url=../b/b.asp?b=" & GBL_Board_ID & "]" & htmlencode(KillHTMLLabel(GBL_Board_BoardName)) & "[/url]" & VbCrLf & _
			"[b]�ظ����ߣ�[/b]" & GBL_CHK_User & VbCrLf & _
			"[b]�ظ����ӣ�[/b][url=../a/a.asp?b=" & GBL_Board_ID & "&id=" & NewAnnounceID & "]" & htmlencode(Form_Title) & "[/url]",GBL_IPAddress
		End If
	End If
	
	Rem ���淢��ͬ��
	dim weiBoFlag
	If GetFormData("bindpost_1_1") = "1" Then
		weiBoFlag = 0
	Else
		weiBoFlag = 1
	End If
	If GetFormData("bindpost_1") = "1" or weiBoFlag = 0 Then
		if weiBoFlag = 0 and GetFormData("bindpost_1") = "0" then weiBoFlag = 2
		Dim ConnetBind
		set ConnetBind = New Connet_Bind
		call ConnetBind.PostShare(1,LeftTrue(Form_Title,72),getInstallDir("a/a2.asp") & "a/a.asp?b=" & GBL_Board_ID & "&id=" & NewAnnounceID,"",LeftTrue(clearUbbcode(GetReContent(Form_Content)),80),"",weiBoFlag)
		set ConnetBind = nothing
	End If

End Function

function getInstallDir(filedir)

	dim HomeUrl
	HomeUrl = Request.ServerVariables("server_name")
	If Request.ServerVariables("SERVER_PORT") <> "80" Then HomeUrl = HomeUrl & ":" & Request.ServerVariables("SERVER_PORT")
	HomeUrl = HomeUrl & replace(lcase(Request.Servervariables("SCRIPT_NAME")),lcase(filedir),"")
	getInstallDir = "http://" & HomeUrl

End function

Function DisplayAnnounceAccessfull

Global_TableHead%>
<div class=contentbox>
		<table border=0 cellpadding=0 cellspacing=0 width="100%" class=tablebox>
		<tr class=tbhead>
			<td><div class=value><%
			If Re_ID=0 Then
				Response.Write "�շ���������"
			Else
				Response.Write "�շ��Ļظ�����"
			End If%> �Ѿ��ɹ��������桰<%=GBL_Board_BoardName%>���У�������ѡ�����²�����</div></td>
		</tr>
		</table>
		<table border=0 cellpadding=0 cellspacing=0 width="100%" class=tablebox>
		<tr>
			<td class=tdright>
			<br>
			<%
			Dim UpdateFlag
			UpdateFlag = UpdateUserLevel(Form_UserID)
			If GBL_BoardMasterFlag < 4 and GetBinarybit(GBL_Board_BoardLimit,13) = 1 Then
				Response.Write "�����ĵȴ���̳����Ա����������ӡ�<br>"
			End If

			
			If Upd_ErrInfo <> "" Then Response.Write "<font color=Red class=redfont>" & Upd_ErrInfo & "</font><br>"%>
			��ҳ�潫��5����Զ�����������������ӣ����Լ���ѡ�����²�����<br>
				<script language=javascript>
				function a_topage()
				{
					this.location.href = "a.asp?B=<%=GBL_board_ID%>&ID=<%=LMT_RootIDBak%>&AUpflag=1&ANum=1"; 
				}
				setTimeout("a_topage()",5000);
				</script>
				<ul>
					<li><a href=../Boards.asp>������ҳ</a><br>
					<li>����<a href=../b/b.asp?B=<%=GBL_board_ID%>><%=GBL_Board_BoardName%></a>��̳<br>
					<%
					If (LMT_ChildNum + 2) > DEF_TopicContentMaxListNum Then%>
					<li>����<a href=a.asp?B=<%=GBL_board_ID%>&ID=<%=LMT_RootIDBak%>>�շ��������</a><br>
					<%End If%>
					<li>��<a href=a.asp?B=<%=GBL_board_ID%>&ID=<%=LMT_RootIDBak%>&AUpflag=1&ANum=1>�շ��������</a>
				</ul>
			</td>
		</tr>
		
		</table>
</div>
<%
REM *******Chat Start*******
If GBL_CheckLimitTitle(GBL_Board_ForumPass,GBL_Board_BoardLimit,GBL_Board_OtherLimit,GBL_Board_HiddenFlag) = 1 Then
	Form_Title = "<font color=gray calss=grayfont>�����ӱ���������Ϊ����</font>"
	Form_TitleStyle = 1
Else
	If Left(Form_Title,3) = "re:" and Form_Title <> "re:" Then Form_Title = Mid(Form_Title,4)
End If
CALL Chat_Appand_pop(3,"<span onclick=c_sc(this.innerHTML) style=cursor: pointer class=c_name>" & GBL_CHK_User & "</span>�������ӣ�<a href=../../a/a.asp?B=" & GBL_board_ID & "&ID=" & LMT_RootIDBak & "&AUpflag=1&ANum=1 target=_blank>" & Replace(DisplayAnnounceTitle(Form_Title,Form_TitleStyle),"""","\""") & "</a>��")
REM *******Chat End*********

	
	If UpdateFlag = 0 Then
		Response.Clear
		CloseDatabase
		Response.Redirect "a.asp?B=" & GBL_board_ID & "&ID=" & LMT_RootIDBak & "&AUpflag=1&ANum=1"
	End If
	Global_TableBottom

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

Function GetReContent(str)

	Dim in1,in2,Str2
	in1 = inStr(str,"[QUOTE]")
	if in1 > 0 Then
		in2 = inStr(str,"[/QUOTE]")
	Else
		in2 = 0
	End If
	If in1 = 0 or in2 = 0 or in1 >= in2 Then
		GetReContent = str
		Exit Function
	End If
	Str2 = Left(str,in1-1) & Mid(str,in2+8)
	If Left(Str2,2) = VbCrLf Then Str2 = Mid(Str2,3)
	If Replace(Trim(Str2),VbCrLf,"") = "" Then Str2 = Replace(Replace(Str,"[QUOTE]",""),"[/QUOTE]","")
	GetReContent = Str2

End Function

Function GetTopicInfo

	If Re_ID = 0 Then Exit Function
	Dim ThisParentID,TmpContent,LMT_LastTime
	ThisParentID = 1
	Dim Rs,SQL,Form_TopicType,Form_NeedValue,TParentID

	TParentID = -1
	
	Dim RootIDBak

	Dim ac,rd
	ac = Request.QueryString("ac")
	rd = Left(Request.QueryString("rd"),14)
	If isNumeric(rd) = 0 or inStr(rd,".") then rd = 0
	rd = cCur(rd)
	If rd = 0 Then
		SQL = ""
	Else
		Select Case ac
			Case "pre": SQL = sql_select("Select t1.ID,t1.RootID,t1.TopicType,t1.NeedValue,t1.ParentID,t1.ChildNum,t1.Title,t1.hits,t1.NotReplay,t1.Content,t1.UserName,t1.RootIDBak,t1.TitleStyle,t1.Opinion,t1.HtmlFlag,T2.UserLimit,T1.VisitIP,T1.RootMaxID,T1.RootMinID,T1.LastTime from LeadBBS_Announce as T1 left join LeadBBS_User as T2 on T1.UserID=T2.ID where t1.ParentID=0 and t1.boardid=" & GBL_board_ID & " and t1.RootID>" & rd & " order by t1.RootID ASC",1)
			Case "nxt": SQL = sql_select("Select t1.ID,t1.RootID,t1.TopicType,t1.NeedValue,t1.ParentID,t1.ChildNum,t1.Title,t1.hits,t1.NotReplay,t1.Content,t1.UserName,t1.RootIDBak,t1.TitleStyle,t1.Opinion,t1.HtmlFlag,T2.UserLimit,T1.VisitIP,T1.RootMaxID,T1.RootMinID,T1.LastTime from LeadBBS_Announce as T1 left join LeadBBS_User as T2 on T1.UserID=T2.ID where t1.ParentID=0 and t1.boardid=" & GBL_board_ID & " and t1.RootID<" & rd & " order by t1.RootID DESC",1)
			Case Else: SQL = ""
		End Select
	End If

	If SQL <> "" Then
		Set Rs = LDExeCute(SQL,0)
		If Not Rs.Eof Then
			Re_ID = cCur(Rs(0))
			'LMT_RootIDBak = Re_ID
			LMT_RootIDBak = cCur(Rs(11))
			LMT_RootID = cCur(Rs(1))
			Form_TopicType = Rs(2)
			Form_NeedValue = Rs(3)
			TParentID = cCur(Rs(4))
			ThisParentID = TParentID
			LMT_ChildNum = cCur(Rs(5))
			LMT_TopicName = Rs(6)
			'Rs(7) = cCur(Rs(7)) + 1
			A_NotReplay = Rs(8)
			Topic_UserName = Rs(10) & ""
			LMT_ReName = Topic_UserName
			RootIDBak = Rs(11)
			LMT_TopicTitleStyle = Rs(12)
			ac = Trim(Rs(16))
			If GetBinarybit(Rs(15),7) = 1 or Form_Submitflag <> "first" or Request.QueryString("repost") <> "1" or LMT_TopicTitleStyle >= 60 or (Form_TopicType > 0 and Form_TopicType <> 80) Then
				'TmpContent = "���û������Ѿ�������Ա���Σ�����������Ч��"
			Else
				Select Case Rs(14)
					Case 0: TmpContent = GetReContent(ResumeCode(Rs(9)))
					Case 1: TmpContent = KillHTMLLabel(GetReContent(Left(Rs(9),500)))
					Case 2: 
							'If LMT_DefaultEdit = 0 Then
							'	TmpContent = clearUbbcode(GetReContent(Rs(9)))
							'Else
								TmpContent = clearUbbcode(GetReContent(Rs(9)))
							'End If
					Case 3: 
							'If LMT_DefaultEdit = 0 Then
							'	TmpContent = clearUbbcode(GetReContent(ResumeCode(Rs(9))))
							'Else
								TmpContent = clearUbbcode(GetReContent(ResumeCode(Rs(9))))
							'End If
							Form_HTMLFlag = 2
				End Select
				Do While Right(TmpContent,2) = VbCrLf
					TmpContent = Mid(TmpContent,1,len(TmpContent)-2)
				Loop
				If Len(TmpContent)>100 Then
					Form_Content = "[QUOTE][b]����������[@" & LMT_ReName & "][url=a.asp?b=" & GBL_Board_ID & "&id=" & Re_ID & "]���������[/url]��[/b]" & VbCrLf & Left(TmpContent,100) & "...[/QUOTE]" & VbCrLf
				Else
					Form_Content = "[QUOTE][b]����������[@" & LMT_ReName & "][url=a.asp?b=" & GBL_Board_ID & "&id=" & Re_ID & "]���������[/url]��[/b]" & VbCrLf & TmpContent & "[/QUOTE]" & VbCrLf
				End If
				'If LMT_DefaultEdit = 0 Then Form_Content = UBB_Code(Form_Content)
			End If
			Form_RootMaxID = cCur(Rs(17))
			LMT_LastTime = cCur(Rs(19))
			Form_Hits = cCur(Rs(7)) + 1
			Rs.Close
			Set Rs = Nothing
		Else
			Rs.Close
			Set Rs = Nothing
		End If
	End If

	If TParentID = -1 Then
		SQL = sql_select("Select T1.RootID,t1.TopicType,t1.NeedValue,t1.ParentID,t1.ChildNum,t1.Hits,t1.title,t1.NotReplay,t1.Content,t1.UserName,t1.RootIDBak,t1.BoardID,t1.TitleStyle,t1.Opinion,t1.HtmlFlag,T2.UserLimit,T1.VisitIP,T1.RootMaxID,T1.RootMinID,T1.LastTime from LeadBBS_Announce as T1 left join LeadBBS_User as T2 on T1.UserID=T2.ID where T1.ID=" & Re_ID,1)
		Set Rs = LDExeCute(SQL,0)
		If Rs.Eof Then
			Rs.Close
			Set Rs = Nothing
			GBL_CHK_TempStr = "����,�����ⲻ����!<br>" & VbCrLf
			Exit Function
		Else
			If cCur(Rs(11)) <> GBL_board_ID Then
				Rs.Close
				Set Rs = Nothing
				GBL_CHK_TempStr = "����,�����ⲻ����!<br>" & VbCrLf
				Exit Function
			Else
				LMT_RootID = Rs(0)
				'LMT_RootIDBak = Re_ID
				Form_TopicType = Rs(1)
				Form_NeedValue = Rs(2)
				TParentID = cCur(Rs(3))
				ThisParentID = TParentID
				LMT_ChildNum = cCur(Rs(4))
				'Rs(5) = cCur(Rs(5)) + 1
				LMT_TopicName = Rs(6)
				A_NotReplay = Rs(7)
				RootIDBak = Rs(10)
				LMT_RootIDBak = cCur(RootIDBak)
				Topic_UserName = Rs(9)
				LMT_ReName = Topic_UserName
				LMT_TopicTitleStyle = Rs(12)

				ac = Trim(Rs(16))
				If GetBinarybit(Rs(15),7) = 1 or Form_Submitflag <> "first" or Request.QueryString("repost") <> "1" or LMT_TopicTitleStyle >= 60 or (Form_TopicType > 0 and Form_TopicType <> 80) Then
					'TmpContent = "���û������Ѿ�������Ա���Σ�����������Ч��"
				Else
					Select Case Rs(14)
						Case 0: TmpContent = GetReContent(ResumeCode(Rs(8)))
						Case 1: TmpContent = KillHTMLLabel(GetReContent(Left(Rs(8),500)))
						Case 2: 
								'If LMT_DefaultEdit = 0 Then
								'	TmpContent = clearUbbcode(GetReContent(Rs(8)))
								'Else
									TmpContent = clearUbbcode(GetReContent(Rs(8)))
								'End If
						Case 3: 
								'If LMT_DefaultEdit = 0 Then
								'	TmpContent = clearUbbcode(GetReContent(ResumeCode(Rs(8))))
								'Else
									TmpContent = clearUbbcode(GetReContent(ResumeCode(Rs(8))))
								'End If
								Form_HTMLFlag = 2
					End Select
					Do While Right(TmpContent,2) = VbCrLf
						TmpContent = Mid(TmpContent,1,len(TmpContent)-2)
					Loop
					If Form_Submitflag = "first" and Request.QueryString("repost") = "1" Then
						If Len(TmpContent)>100 Then
							Form_Content = "[QUOTE][b]����������[@" & LMT_ReName & "][url=a.asp?b=" & GBL_Board_ID & "&id=" & Re_ID & "]���������[/url]��[/b]" & VbCrLf & VbCrLf & Left(TmpContent,100) & "...[/QUOTE]" & VbCrLf
						Else
							Form_Content = "[QUOTE][b]����������[@" & LMT_ReName & "][url=a.asp?b=" & GBL_Board_ID & "&id=" & Re_ID & "]���������[/url]��[/b]" & VbCrLf & VbCrLf & TmpContent & "[/QUOTE]" & VbCrLf
						End If
						'If LMT_DefaultEdit = 0 Then Form_Content = UBB_Code(Form_Content)
					End If
				End If
				Form_RootMaxID = cCur(Rs(17))
				LMT_LastTime = cCur(Rs(19))
				Form_Hits = cCur(Rs(5)) + 1
				Rs.Close
				Set Rs = Nothing
			End If
		End If
	End If

	If TParentID > 0 Then
		If cCur(RootIDBak) > 0 Then
			select case DEF_UsedDataBase
				case 0,2:
					SQL = sql_select("Select Title,Hits,ChildNum,ID,TitleStyle,NotReplay,RootIDBak,RootID,RootMaxID,RootMinID,UserName,TopicType,NeedValue,LastTime from LeadBBS_Announce where ParentID=0 and RootIDBak=" & RootIDBak & " order by ID ASC",1)
				case Else
					SQL = sql_select("Select Title,Hits,ChildNum,ID,TitleStyle,NotReplay,ID,RootID,RootMaxID,RootMinID,UserName,TopicType,NeedValue,LastTime from LeadBBS_Topic where ID=" & RootIDBak & " order by ID ASC",1)
			End select
		Else
			select case DEF_UsedDataBase
				case 0,2:
					SQL = sql_select("Select Title,Hits,ChildNum,ID,TitleStyle,NotReplay,RootIDBak,RootID,RootMaxID,RootMinID,UserName,TopicType,NeedValue,LastTime from LeadBBS_Announce where ParentID=0 and boardid=" & GBL_board_ID & " and RootID=" & LMT_RootID,1)
				case Else
					SQL = sql_select("Select Title,Hits,ChildNum,ID,TitleStyle,NotReplay,id,RootID,RootMaxID,RootMinID,UserName,TopicType,NeedValue,LastTime from LeadBBS_Topic where boardid=" & GBL_board_ID & " and RootID=" & LMT_RootID,1)
			End select
		End If
		Set Rs = LDExeCute(SQL,0)
		If Rs.Eof Then
			Rs.Close
			Set Rs = Nothing
			GBL_CHK_TempStr = "����,�������Ѿ�ɾ��!<br>" & VbCrLf
			Exit Function
		Else
			LMT_TopicName = Rs(0)
			LMT_ChildNum = cCur(Rs(2))
			LMT_RootIDBak = cCur(Rs(6))
			LMT_TopicTitleStyle = Rs(4)
			If A_NotReplay = 0 Then A_NotReplay = Rs(5)
			LMT_RootID = cCur(Rs(7))
			Form_RootMaxID = cCur(Rs(8))
			Topic_UserName = Rs(10)
			Form_TopicType = Rs(11)
			Form_NeedValue = Rs(12)
			LMT_LastTime = cCur(Rs(13))
			Rs.Close
			Set Rs = Nothing
		End If
	End If

	If Form_TopicType = 39 Then
		GBL_CHK_TempStr = "�������޷��ظ���<br>" & VbCrLf
		A_NotReplay = 1
		Exit Function
	End If
	If Form_TopicType > 0 and A_NotReplay = 0 and Form_TopicType <> 80 Then
		Form_NeedValue = cCur(Form_NeedValue)
		Select case Form_TopicType
			Case 2:
				If GBL_CHK_User <> "" Then
					If GBL_BoardMasterFlag < 5 Then
						A_NotReplay = 1
						If Form_Submitflag <> "" Then GBL_CHK_TempStr = "����,ֻ�б���" & DEF_PointsName(8) & "���ܻظ�����!<br>" & VbCrLf
					End If
				Else
					A_NotReplay = 1
				End If
			Case 4:
				If GBL_BoardMasterFlag < 4 Then
					A_NotReplay = 1
					If Form_Submitflag <> "" Then GBL_CHK_TempStr = "����,ֻ��" & DEF_PointsName(8) & "���ܻظ�����!<br>" & VbCrLf
				End If
			Case 6:
				If GBL_CHK_User = "" or GetBinarybit(GBL_CHK_UserLimit,2) <> 1 Then
					A_NotReplay = 1
					If Form_Submitflag <> "" Then GBL_CHK_TempStr = "����,ֻ��" & DEF_PointsName(5) & "���ܻظ�����!<br>" & VbCrLf
				End If
			Case 51:
				If GBL_CHK_Points < Form_NeedValue Then
					A_NotReplay = 1
					If Form_Submitflag <> "" Then GBL_CHK_TempStr = "����,��Ҫ" & DEF_PointsName(0) & "" & Form_NeedValue & "���ϲ��ܻظ�����!<br>" & VbCrLf
				End If
			Case 53:
				If GBL_CHK_OnlineTime < Form_NeedValue*60 Then
					A_NotReplay = 1
					If Form_Submitflag <> "" Then GBL_CHK_TempStr = "����,��Ҫ" & DEF_PointsName(4) & Form_NeedValue & "���ϲ��ܻظ�����!<br>" & VbCrLf
				End If
			Case 55:
					If Form_NeedValue > 0 and Form_Submitflag <> "" Then
						If Form_NeedValue <> GBL_UserID and GBL_CHK_User <> Topic_UserName Then
							A_NotReplay = 1
							GBL_CHK_TempStr = "����ֻ�������˺ͽ����˻ظ�"
						End If
					End If
		End Select
	End If

	Dim RootTopicNeedValue
	
		
	If Len(LMT_LastTime) = 14 and GBL_BoardMasterFlag < 4 Then
		If DateDiff("d",ReStoreTime(LMT_LastTime),Now) > LMTDEF_NotReplyDate Then
			If Form_Submitflag <> "" Then GBL_CHK_TempStr = "���������ظ�ʱ�䳬��" & LMTDEF_NotReplyDate & "�죬���������ظ���"
			A_NotReplay = 1 '
		End If
	End If
	If Form_TopicType > 0 And Form_TopicType <> 80 and Form_Submitflag = "first" and ThisParentID = 0 and Request.QueryString("repost") = "1" Then
		RootTopicNeedValue = cCur(Form_NeedValue)
		RootTopicType = Form_TopicType
		Select case RootTopicType
			Case 1:
				If GBL_CHK_User <> "" Then
					If GBL_BoardMasterFlag < 5 Then
						GBL_CHK_TempStr = "����ֻ�б���" & DEF_PointsName(8) & "�������ûظ�"
					End If
				Else
					GBL_CHK_TempStr = "����ֻ�б���" & DEF_PointsName(8) & "�������ûظ�"
				End If
			Case 3:
				If GBL_BoardMasterFlag < 4 Then
					GBL_CHK_TempStr = "����ֻ��" & DEF_PointsName(8) & "�������ûظ�"
				End If
			Case 5:
				If GBL_CHK_User = "" or GetBinarybit(GBL_CHK_UserLimit,2) <> 1 Then
					GBL_CHK_TempStr = "����ֻ��" & DEF_PointsName(5) & "�������ûظ�"
				End If
			Case 7,54,49:
				If Form_Submitflag = "first" and Request.QueryString("repost") = "1" Then
					A_NotReplay = 1
					GBL_CHK_TempStr = "���󣬴����������ûظ���<br>" & VbCrLf
				End If
			Case 50:
				If GBL_CHK_Points < RootTopicNeedValue Then
					GBL_CHK_TempStr = "������Ҫ" & DEF_PointsName(0) & "" & RootTopicNeedValue & "�������ûظ�"
				End If
			Case 52:
				If GBL_CHK_OnlineTime < RootTopicNeedValue*60 Then
					GBL_CHK_TempStr = "������Ҫ" & DEF_PointsName(4) & RootTopicNeedValue & "�������ûظ�"
				End If
			Case 115:
					If Form_NeedValue > 0 Then
						If Form_NeedValue <> GBL_UserID and GBL_CHK_User <> Topic_UserName Then
							A_NotReplay = 1
							GBL_CHK_TempStr = "����ֻ�������˺ͽ����˻ظ�"
						End If
					End If
		End Select
	End If

End Function

Function UpdateUserLevel(UserID)

	Dim Temp_N,UserLevel,Points,OnlineTime,Save_UserLevel
	Dim Rs
	Set Rs = LDExeCute(sql_select("Select UserLevel,AnnounceNum2,OnlineTime from LeadBBS_User where id=" & UserID,1),0)
	If Rs.Eof Then
		Rs.Close
		Set Rs = Nothing
		Exit Function
	End If
	UserLevel = Rs(0)
	Save_UserLevel = UserLevel
	Points = cCur(Rs(1))
	OnlineTime = cCur(Rs(2))
	Rs.Close
	Set Rs = Nothing

	For Temp_N = 0 To DEF_UserLevelNum
		'�Զ����¼���ȼ�
		'If Points >= DEF_UserLevelPoints(Temp_N) Then UserLevel = Temp_N
		If Points >= DEF_UserLevelPoints(Temp_N) and Temp_N >= UserLevel Then UserLevel = Temp_N
	Next

	Randomize

	REM **** �ɲƸ������趨��ʼ *****
	Rs = Fix(Rnd*1314)+1
	If Rs = 1314 Then
		Rs = 1
		Response.Write "<br>&nbsp;<font color=""blue"" class=""bluefont"">���������࣬��������" & DEF_PointsName(1) & "֮�񣬴������µ�" & DEF_PointsName(1) & "��</font><br>"
REM *******Chat Start*******
	CALL Chat_Appand_pop(3,"<b><span onclick=""c_sc(this.innerHTML)"" style=""cursor: pointer"" class=""c_name"">" & GBL_CHK_User & "</span>�������࣬��������" & DEF_PointsName(1) & "֮�񣬲������µ�" & DEF_PointsName(1) & "��</b>")
REM *******Chat End*********
	Else
		Rs = 0
	End If
	REM **** �ɲƸ����ʽ��� *****
	
	REM **** �ض��������趨��ʼ *****
	
	'������Ҫ����������ID���,ֻ��������
	Dim AncIDStr
	AncIDStr = "" '�����������ID�б����ŷָ����ظ��������ӽ������������(1-3)��ע����[DelAnnounce.asp]���ñ���һ��

	Dim Tn
	If Re_ID > 0 and GBL_CHK_OnlineTime > 3600 and inStr("," & AncIDStr & ",","," & LMT_RootIDBak & ",") Then
		Set Rs = LDExeCute(sql_select("Select ID,Opinion from LeadBBS_Announce where UserID=" & GBL_UserID & " and ParentID=" & LMT_RootIDBak,10),0)
		Dim FirFlag,Opinion
		FirFlag = 0
		Opinion = ""

		Do While Opinion = "" and Not Rs.Eof
			Opinion = Opinion & Rs(1)
			Rs.MoveNext
		Loop
		If Opinion = "" Then
			FirFlag = 1
		Else
			FirFlag = 0
		End If
		Rs.Close
		Set Rs = Nothing
		Rs = 0

		If FirFlag = 1 Then
			Dim Tmp
			Tmp = Fix(Rnd*1314)+1
			If Tmp = 1314 Then
				Rs = 3
			ElseIf Tmp > 1200 Then
				Rs = 2
			Else
				Rs = 1
			End If
			CALL LDExeCute("Update LeadBBS_Announce Set Opinion='" & "[LeadBBS]|0|����ָ��" & Tmp & "�����" & DEF_PointsName(2) & "" & Rs & "' where ID=" & NewAnnounceID,1)
			Response.Write "<br>&nbsp;<font color=red class=redfont>��ϲ����ظ������������µ�" & DEF_PointsName(2) & "��</font><br>"
		End If
	End If
	REM **** �ض��������趨���� *****

	If Save_UserLevel <> UserLevel or Rs >= 1 Then
		CALL LDExeCute("Update LeadBBS_User set UserLevel=" & UserLevel & ",CachetValue=CachetValue+" & Rs & " where id=" & UserID,1)
		UpdateSessionValue 15,Rs,1
	End If
	If Rs >= 1 Then
		UpdateUserLevel = 1
	Else
		UpdateUserLevel = 0
	End If

End Function

Function UpdateBoardAnnounceNum(BoardList,TopicNum,AnnounceNum,TodayAnnounce,GoodNum,LastWriter,LastWriteTime,LastAnnounceID,LastTopicName)

	Dim SafeFlag
	SafeFlag = 0
	'������̳ ��֤���� רҵ�û����� ���ް�������
	If GBL_Board_ForumPass <> "" or GetBinarybit(GBL_Board_BoardLimit,2) = 1 or GetBinarybit(GBL_Board_BoardLimit,15) = 1 or GetBinarybit(GBL_Board_BoardLimit,7) = 1 Then SafeFlag = 1
	Dim SQL,N,Num
	If BoardList = "" or (TopicNum = 0 and AnnounceNum = 0 and TodayAnnounce = 0 and GoodNum = 0) Then Exit Function
	SQL = "Update LeadBBS_Boards Set AnnounceNum=AnnounceNum+" & AnnounceNum & ",AnnounceNum_All=AnnounceNum_All+" & AnnounceNum
	If TopicNum <> 0 Then SQL = SQL & ",TopicNum=TopicNum+" & TopicNum
	If TopicNum <> 0 Then SQL = SQL & ",TopicNum_All=TopicNum_All+" & TopicNum
	If TodayAnnounce <> 0 Then  SQL = SQL & ",TodayAnnounce=TodayAnnounce+" & TodayAnnounce
	If TodayAnnounce <> 0 Then  SQL = SQL & ",TodayAnnounce_All=TodayAnnounce_All+" & TodayAnnounce
	If GoodNum <> 0 Then SQL = SQL & ",GoodNum=GoodNum+" & GoodNum
	If GoodNum <> 0 Then SQL = SQL & ",GoodNum_All=GoodNum_All+" & GoodNum
	If inStr(BoardList,",") = False Then
		If LastWriter <> "" Then SQL = SQL & ",LastWriter='" & Replace(LastWriter,"'","''") & "'"
		If cCur(LastWriteTime) > 0 Then SQL = SQL & ",LastWriteTime=" & LastWriteTime
		If LastTopicName <> "" Then SQL = SQL & ",LastTopicName='" & Replace(LastTopicName,"'","''") & "'"
		If cCur(LastAnnounceID) > 0 Then SQL = SQL & ",LastAnnounceID=" & LastAnnounceID
		SQL = SQL & " where BoardID=" & GBL_Board_ID
		CALL LDExeCute(SQL,1)
	Else
		If SafeFlag = 0 Then '���ܰ��治�����ϼ��������·���
			If LastWriter <> "" Then SQL = SQL & ",LastWriter='" & Replace(LastWriter,"'","''") & "'"
			If cCur(LastWriteTime) > 0 Then SQL = SQL & ",LastWriteTime=" & LastWriteTime
			If LastTopicName <> "" Then SQL = SQL & ",LastTopicName='" & Replace(LastTopicName,"'","''") & "'"
			If cCur(LastAnnounceID) > 0 Then SQL = SQL & ",LastAnnounceID=" & LastAnnounceID
		End If
		SQL = SQL & " where BoardID in(" & BoardList & "," & GBL_Board_ID & ")"
		CALL LDExeCute(SQL,1)
	End If
	BoardList = Split(BoardList,",")
	Num = Ubound(BoardList,1)
	Dim Temp
	For N = 0 To Num
		Temp = Application(DEF_MasterCookies & "BoardInfo" & BoardList(N))
		If isArray(Temp) = False Then
			ReloadBoardInfo(BoardList(N))
			Temp = Application(DEF_MasterCookies & "BoardInfo" & BoardList(N))
		End If
		If isArray(Temp) = True Then
			'��������ͳ����Ϣ���Ӱ��汣���������
			If TopicNum <> 0 Then
				UpdateBoardApplicationInfo BoardList(N),cCur(Temp(5,0))+TopicNum,5
				UpdateBoardApplicationInfo BoardList(N),cCur(Temp(29,0))+TopicNum,29
			End If
			If AnnounceNum <> 0 Then
				UpdateBoardApplicationInfo BoardList(N),cCur(Temp(6,0))+AnnounceNum,6
				UpdateBoardApplicationInfo BoardList(N),cCur(Temp(30,0))+AnnounceNum,30
			End If
			If TodayAnnounce <> 0 Then
				UpdateBoardApplicationInfo BoardList(N),cCur(Temp(18,0))+TodayAnnounce,18
				UpdateBoardApplicationInfo BoardList(N),cCur(Temp(31,0))+TodayAnnounce,31
			End If
			If GoodNum <> 0 Then
				UpdateBoardApplicationInfo BoardList(N),cCur(Temp(13,0))+GoodNum,13
				UpdateBoardApplicationInfo BoardList(N),cCur(Temp(32,0))+GoodNum,32
			End If
			If (SafeFlag = 1 and GBL_Board_ID = cCur(BoardList(N))) or SafeFlag = 0 Then
				If SafeFlag = 0 Then	
					If LastWriter <> "" Then UpdateBoardApplicationInfo BoardList(N),LastWriter,3
					If cCur(LastWriteTime) > 0 Then UpdateBoardApplicationInfo BoardList(N),Form_LastTime,4
					If LastTopicName <> "" Then UpdateBoardApplicationInfo BoardList(N),LastTopicName,20
				End If
				If SafeFlag = 1 or Num = 0 or GBL_Board_ID = cCur(BoardList(N)) Then
					If cCur(LastAnnounceID) > 0 Then UpdateBoardApplicationInfo BoardList(N),LastAnnounceID,19
				Else
					'If cCur(LastAnnounceID) > 0 Then UpdateBoardApplicationInfo BoardList(N),0,19
					If cCur(LastAnnounceID) > 0 Then UpdateBoardApplicationInfo BoardList(N),LastAnnounceID,19
				End If
			End If
		End If
	Next
	'28,T1.ParentBoardStr,29.TopicNum_All,30.AnnounceNum_All,31.TodayAnnounce_All,32.GoodNum_All

End Function

Sub Main

	If dontRequestFormFlag = "" Then
		Select Case Left(Request.Form("ol"),1)
			Case "1":
				CALL Editor_View(Edt_MiniMode,"")
				Exit Sub
		End Select
	End If

	Free_UDT
	GetBoardID
	initDatabase
	GetRequestValue
	CheckisBoardMaster
	GBL_CHK_TempStr = ""
	If Re_ID > 0 Then Form_VoteFlag = ""
	GetTopicInfo

	If GetBinarybit(GBL_Board_BoardLimit,16) = 1 Then
		If LMT_DefaultEdit = 1 Then
			LMT_DefaultEdit = 0
		Else
			LMT_DefaultEdit = 1
		End If
	End If

	If LMT_TopicTitleStyle >= 60 and GBL_BoardMasterFlag < 4 Then
		LMT_TopicNameNoHTML = "���ӵȴ������..."
		LMT_TopicName = "<font color=gray class=grayfont>���ӵȴ������...</font>"
		'LMT_TopicTitleStyle = 1
		A_NotReplay = 1
	Else
		If LMT_TopicTitleStyle = 1 Then
			LMT_TopicNameNoHTML = KillHTMLLabel(LMT_TopicName)
		Else
			LMT_TopicNameNoHTML = LMT_TopicName
		End If
	End If
	LMT_TopicNameNoHTML_Temp = LMT_TopicNameNoHTML
	
	If strLength(LMT_TopicNameNoHTML_Temp)>DEF_BBS_DisplayTopicLength-6 Then
		LMT_TopicNameNoHTML_Temp = htmlencode(LeftTrue(LMT_TopicNameNoHTML_Temp,DEF_BBS_DisplayTopicLength-9)) & "..."
	Else
		LMT_TopicNameNoHTML_Temp = htmlencode(LMT_TopicNameNoHTML_Temp)
	End if
	If Re_ID > 0 Then
		CheckAccessLimit_TimeLimit
		LMT_TopicNameNoHTML_Temp = "�ظ���" & LMT_TopicNameNoHTML_Temp
		BBS_SiteHead DEF_SiteNameString & " - " & KillHTMLLabel(GBL_Board_BoardName) & " - �ظ�����",GBL_board_ID,"<span class=navigate_string_step>�ظ�����</span>"
	Else
		CheckAccessLimit_TimeLimit
		If Form_VoteFlag <> "" and Re_ID = 0 Then
			LMT_TopicNameNoHTML_Temp = "������ͶƱ"
		Else
			LMT_TopicNameNoHTML_Temp = "����������"
		End If
		BBS_SiteHead DEF_SiteNameString & " - " & KillHTMLLabel(GBL_Board_BoardName) & " - " & LMT_TopicNameNoHTML_Temp,GBL_board_ID,"<span class=navigate_string_step>" & LMT_TopicNameNoHTML_Temp & "</span>"
	End If

	Boards_Body_Head("")
	CheckAccessLimit
	If Form_Submitflag <> "" or Re_ID = 0 Then
		If Re_ID = 0 Then
			CheckBoardAnnounceLimit
		Else
			CheckBoardReAnnounceLimit
		End If
		CheckUserAnnounceLimit
	End If
	If GBL_CHK_TempStr = "" Then
		If Form_Submitflag = "" Then
			'GetRequestValue
			Form_NoUserUnderWriteFlag=1
			If Re_ID<>0 Then
				If Left(LMT_TopicName,3) <> "Re:" Then
					Form_Title = "Re:" & LMT_TopicNameNoHTML
				Else
					Form_Title = LMT_TopicNameNoHTML4
				End If
				'Form_Title = ""
			End If
			DisplayAnnounceForm
			GBL_CHK_TempStr = ""
		Else
			If A_NotReplay = 0 and LMT_ChildNum > LMTDEF_MaxReAnnounce and DEF_EnableTreeView = 1 Then
				Global_ErrMsg "�ظ��Ѿ��ﵽ�����Ŀ,�����ٻظ�����,�������⡣"
				GBL_CHK_TempStr = " "
			ElseIf A_NotReplay = 1 and Form_Submitflag = "first" Then
				Global_ErrMsg "������������״̬��������ظ���"
				GBL_CHK_TempStr = " "
			Else
				If CheckAnnouceValue = 1 Then
					If SaveAnnounceValue = 1 Then
						DisplayAnnounceAccessfull
					Else
						If Form_Submitflag <> "first" Then
							Global_ErrMsg GBL_CHK_TempStr
							GBL_CHK_TempStr = " "
						End If
						DisplayAnnounceForm
					End If
				Else
					If Form_Submitflag <> "first" Then
						Global_ErrMsg GBL_CHK_TempStr
					End If
					DisplayAnnounceForm
				End If
			End If
		End If
		UpdateOnlineUserAtInfo GBL_board_ID,GBL_Board_BoardName & "��" & LMT_TopicNameNoHTML_Temp
	Else
		Global_ErrMsg GBL_CHK_TempStr
	End If
	If Form_UpFlag = 1 Then Set Form_UpClass = Nothing
	CloseDatabase
	Boards_Body_Bottom
	SiteBottom

End Sub

Main
%>