<!-- #include file=Editor.asp -->
<%
Const Edt_MiniMode = 1 '�������棺0-��ͳ��Լģʽ 1.�๦��ģʽ
Dim UploadListData,UploadListNum,EditFlag,UploadTable
UploadListNum = 0
EditFlag = 0
UploadTable = "LeadBBS_Upload"
Dim UploadOneDayMaxNum,Upd_SpendFlag
UploadOneDayMaxNum = DEF_UploadOneDayMaxNum
Upd_SpendFlag = 0
Dim upload_NoteLength
upload_NoteLength = 30

Dim LMT_MaxTextLength
If CheckSupervisorUserName = 0 Then
	LMT_MaxTextLength = DEF_MaxTextLength
Else
	LMT_MaxTextLength = DEF_MaxTextLength * 4
End If

Dim LMT_DefaultEdit
LMT_DefaultEdit = DEF_UbbDefaultEdit

Sub ReloadTopicAssort(BoardID)

	Dim Rs
	Set Rs = LDExeCute("select ID,AssortName,0,0,0 from LeadBBS_GoodAssort where BoardID=" & BoardID & " Order by BoardID,OrderID",0)
	If Not Rs.Eof Then
		Application.Lock
		Application(DEF_MasterCookies & "BoardInfo" & BoardID & "_TI") = Rs.GetRows(-1)
		Application.UnLock
	Else
		Application.Lock
		Set Application(DEF_MasterCookies & "BoardInfo" & BoardID & "_TI") = Nothing
		Application(DEF_MasterCookies & "BoardInfo" & BoardID & "_TI") = "yes"
		Application.UnLock
	End If
	Rs.Close
	Set Rs = Nothing

End Sub

Function DisplayLeadBBSEditor1(Form_HTMLFlag,Form_Content,refer,hidemoreinfo)

	If Edt_MiniMode = 0 and refer = 0 Then%>
	<tr>
		<td width="<%=DEF_BBS_LeftTDWidth%>" class=tdleft>����UBB��ǩ</td>
		<td class=tdright>
			<img src="../images/ubb/bold.GIF" style="cursor: pointer" onclick="addcontent(0,'B','/B');" title=���� width=20 height=20 align=middle border=0>
			<img src="../images/ubb/italicize.GIF" style="cursor: pointer" onclick="addcontent(0,'I','/I');" title=б�� width=20 height=20 align=middle border=0>
			<img src="../images/ubb/underline.GIF" style="cursor: pointer" onclick="addcontent(0,'U','/U');" title=�»��� width=20 height=20 align=middle border=0>
			<img src="../images/ubb/center.GIF" style="cursor: pointer" onclick="addcontent(0,'ALIGN','/ALIGN','CENTER');" title=���� width=20 height=20 align=middle border=0>
			<img src="../images/ubb/url1.GIF" style="cursor: pointer" onclick="addcontent(0,'URL','/URL');" title=���� width=20 height=20 align=middle border=0>
			<img src="../images/ubb/email1.GIF" style="cursor: pointer" onclick="addcontent(0,'EMAIL','/EMAIL');" title=�ʼ� width=20 height=20 align=middle border=0>
			<%If DEF_EnableImagesUBB = 1 Then%><img src="../images/ubb/image.GIF" style="cursor: pointer" onclick="addcontent(0,'IMG','/IMG');" title=ͼƬ width=20 height=20 align=middle border=0><%end If%>
			<img src="../images/ubb/swf.GIF" style="cursor: pointer" onclick="addcontent(0,'FLASH','/FLASH');" title=Flash width=20 height=20 align=middle border=0>
			<img src="../images/ubb/code.GIF" style="cursor: pointer" onclick="addcontent(0,'CODE','/CODE');" title=���� width=20 height=20 align=middle border=0>
			<img src="../images/ubb/quote1.GIF" style="cursor: pointer" onclick="addcontent(0,'QUOTE','/QUOTE');" title=���� width=20 height=20 align=middle border=0>
			<img src="../images/ubb/fly.GIF" style="cursor: pointer" onclick="addcontent(0,'FLY','/FLY');" title=���� width=20 height=20 align=middle border=0>
			<img src="../images/ubb/light.GIF" style="cursor: pointer" onclick="addcontent(0,'LIGHT','/LIGHT');" title=��˸���� width=20 height=20 align=middle border=0>
			<img src="../images/ubb/glow.GIF" style="cursor: pointer" onclick="addcontent(0,'GLOW=255,RED,2','/GLOW');" title=���� width=20 height=20 align=middle border=0>
			<img src="../images/ubb/shadow.GIF" style="cursor: pointer" onclick="addcontent(0,'SHADOW=255,RED,2','/SHADOW');" title=��Ӱ width=20 height=20 align=middle border=0>
			<img src="../images/ubb/size3.GIF" style="cursor: pointer" onclick="addcontent(0,'SIZE','/SIZE','3');" title=3���� width=20 height=20 align=middle border=0>
			<img src="../images/ubb/blue.gif" style="cursor: pointer" onclick="addcontent(0,'COLOR','/COLOR','blue');" title=��ɫ�� width=20 height=20 align=middle border=0>
			<img src="../images/ubb/red.GIF" style="cursor: pointer" onclick="addcontent(0,'COLOR','/COLOR','red');" title=��ɫ�� width=20 height=20 align=middle border=0>
			<%If DEF_EnableFlashUBB = 1 then%><img src="../images/ubb/media.gif" style="cursor: pointer" onclick="addcontent(0,'MP=320,309','/MP');" title=����Media�ļ� width=20 height=20 align=middle border=0>
			<img src="../images/ubb/real.gif" style="cursor: pointer" onclick="addcontent(0,'RM=320,260','/RM');" title=����RealPlay�ļ� width=20 height=20 align=middle border=0><%End If%>
		</td>
	</tr><%End If
	if refer = 0 Then%>
	<tr>
		<td valign="top" width="<%=DEF_BBS_LeftTDWidth%>" class=tdleft>
	<%End If
		If hidemoreinfo = 1 then%>
					����(���<%=Fix(LMT_MaxTextLength/1024)%>K)
						<br><br>
						<%If ((GetBinarybit(GBL_CHK_UserLimit,16) = 1 and GBL_BoardMasterFlag >= 2) or CheckSupervisorUserName = 1) or refer > 0 Then%>
						
						���뷽ʽ<%if refer = 0 Then%><br><%end if%>
							<label>
							<input class=fmchkbox type="radio" name="Form_HTMLFlag" value="0"<%If Form_HTMLFlag=0 Then Response.Write " checked"%>>�ı�</label>
							<label>
							<input class=fmchkbox type="radio" name="Form_HTMLFlag" value="1"<%If Form_HTMLFlag=1 Then Response.Write " checked"%>>HTML</label>
							<label>
							<input class=fmchkbox type="radio"" name="Form_HTMLFlag" value="2"<%If Form_HTMLFlag=2 Then Response.Write " checked"%>>UBB</label>
							<%Else%>
							<label>����UBB����
							<input class=fmchkbox type="checkbox" name="Form_HTMLFlag" value="2"<%If Form_HTMLFlag=2 Then Response.Write " checked"%>>
							</label>
							<%End If%>
					<br>
					<a href="../User/Help/Ubb.asp" target=_blank>����֧�ֲ���գ£±�ǩ��ʹ�÷�����ο�����</a>
					<%if refer = 0 Then
						response.write "<br>"
					else
						response.write " - "
					end if%>
					<a href=#icon onclick="alert('���������Ϊ'+edt_getdoclen()+'���֣������<%=DEF_MaxTextLength%>��');">�鿴��������</a>
					<%if refer = 0 Then
						response.write "<br>"
					else
						response.write " - "
					end if%><span class=layerico><a href=#icon onclick="edt_mode?copyClipboard('text',edt_txtobj.value,'�ɹ�����','<%=DEF_BBS_HomeUrl%>',this):copyClipboard('text',edt_doc.body.innerHTML,'�ɹ�����','<%=DEF_BBS_HomeUrl%>',this);">��������</a></span>
					<%if refer = 0 Then
						response.write "<br>"
					else
						response.write "<br><br>"
					end if
			End If%>
	<%if refer = 0 Then%>
		</td>
		<td valign=top class=tdright>
	<%End If%>

<script src="<%=DEF_BBS_HomeUrl%>a/inc/leadedit.js?ver=20080729.22"></script>
<script type="text/javascript">edt_heigh = 220;</script>
<%
CALL Editor_View(Edt_MiniMode,Form_Content)%>
<script type="text/javascript">edt_init();
edt_setmode(0);edt_setmode(<%
	If LMT_DefaultEdit = 1 Then
		Response.Write "0"
	Else
		Response.Write "1"
	End If%>);edt_initdone=1;
if(typeof submitflag != 'undefined')window.onbeforeunload = function(){if(edt_getdoclen()>0&&submitflag==0)return("��������δ����ȷ��ȡ����");}</script>
<%if refer = 0 Then%>
			</td>
		</tr><%
	End If
	If DEF_UBBiconNumber > 0 Then%>
<script type="text/javascript">

</script>
<%
	End If
	If LMT_EnableUpload = 1 Then DisplayUpload(refer)

End Function

Sub DisplayPreview

Dim Temp
Temp = LCase(Request.ServerVariables("server_name"))
If inStr(Temp,".") <> inStrRev(Temp,".") Then Temp = Mid(Temp,inStr(Temp,".") + 1)
%>
<script src="<%=DEF_BBS_HomeUrl%>a/inc/leadcode.js?ver=20080728.225"></script>
<script language=javascript>var GBL_domain="<%=Temp%>";HU="<%=DEF_BBS_HomeUrl%>";var DEF_DownKey="<%=UrlEncode(DEF_DownKey)%>";</script>
<span ID=Preview Style='display:none'><%
Global_TableHead%>
<div class=contentbox>
	<table border=0 cellpadding=0 cellspacing=0 width="100%" class=tablebox>
		<tr class=tbhead>
			<td><div class=value>
			<b>����Ԥ��</b>
			[<a href=#icon onclick="edt_preview(1);">�ر���ʾ</a>]
			</div>
		</td>
	</tr>
	</table>
	<table border=0 cellpadding=0 cellspacing=0 width="100%" class=tablebox>
	<tr class=tdleft>
		<td width=<%=DEF_BBS_LeftTDWidth%> valign=top class=tdleft>
			<b><span id=Preview_UserName> </span></b>
		</td>
		<td valign=top class=tdright>
			<table width=100% style="table-layout:fixed; word-break:break-all"><td>
				<div class=word-break-all>
				<b><span id=Preview_Title> </span></b>
				<span id=Preview_Content class=a_content style=font-size:<%=DEF_AnnounceFontSize%>>
				</span>
				</div>
			</td>
			</tr>
			</table>
		</td>
	</tr>
	</table>
</div>
<%
	Global_TableBottom
%>
<img src=../images/null.gif height=4 width=2><br></span>
<%

End Sub

Sub DisplayUpload(refer)

if refer = 0 Then
%>
<tr>
<td width="<%=DEF_BBS_LeftTDWidth%>" valign=top class=tdleft>
<span id="uptext" name="uptext">�ϴ�����</span>
			<%
			If DEF_Upd_SpendFlag = 1 or GBL_BoardMasterFlag < 4 Then
				If DEF_UploadSpendPoints > 0 Then
					Response.Write " <font color=blue class=bluefont>����" & DEF_UploadSpendPoints & "" & DEF_PointsName(0) & "</font>"
				ElseIf DEF_UploadSpendPoints < 0 Then
					Response.Write " <font color=green class=greenfont>���" & 0-DEF_UploadSpendPoints & "" & DEF_PointsName(0) & "</font>"
				End If%>
				<br>ɾ������<%
				If DEF_UploadDeletePoints > 0 Then
					Response.Write " <font color=blue class=bluefont>����" & DEF_UploadDeletePoints & "" & DEF_PointsName(0) & "</font>"
				ElseIf DEF_UploadDeletePoints < 0 Then
					Response.Write " <font color=green class=greenfont title=�����Լ�ɾ��������Ӧ�ı仯>���" & 0-DEF_UploadDeletePoints & "" & DEF_PointsName(0) & "</font>"
				End If
			End If
%>
</td>
<td class=tdright valign=top>
<%
Else
	Response.Write "<br>"
End If

If EditFlag = 1 Then
	DisplayUploadEdit
End If
If DEF_UploadOnceNum-UploadListNum > 0 Then
%>
<div><b>�ϴ��¸���</b></div>
<div id=upload_node style="display:none;margin-top:5px;">
�ļ� <span><span><input name="file_number" type="file" onchange="upl_onchange(this.name)" id=file_number size="20" class="fminpt"></span></span> ע�� <input name="text_number" type="text" maxlength=<%=upload_NoteLength%> size="20" class='fminpt input_2 note'>
</div>
<div id=new_upload>
<div id=upload0 style="margin-top:5px;">
�ļ� <span><span><input name="file0" type="file" id=file0 size="20" class="fminpt uninit_upload" onchange=upl_onchange(0)></span></span>
ע�� <input name="text0" type="text" maxlength=<%=upload_NoteLength%> size="20" class="fminpt input_2 note">
<span id=upload_del0 style=display:none><a href=#none onclick="upl_remove(this.parentNode.parentNode);">ɾ��</a>
<a href=#none onclick="addcontent(1,'[upload=0]');">����</a></span>
</div>
</div><%End If
%><br>
<div id=upload_doc> </div>
<p>ע��������С����Ϊ <%=int(DEF_FileMaxBytes/1024)%>K<%
If UploadOneDayMaxNum > 0 Then Response.Write " ÿ������ϴ�" & UploadOneDayMaxNum & "��"
%>
<script type="text/javascript">
init_uploadform();
var DEF_UploadFileType="<%=DEF_UploadFileType%>";
var DEF_UploadOnceNum = <%=DEF_UploadOnceNum-UploadListNum%>,DEF_UploadOneDayMaxNum=<%=UploadOneDayMaxNum%>;
if(DEF_UploadOneDayMaxNum<DEF_UploadOnceNum && DEF_UploadOneDayMaxNum > 0)DEF_UploadOnceNum = DEF_UploadOneDayMaxNum;

var Upl_IOfun,Upl_Level=0,Upl_GetDelay=2000,Upl_Start = false,Upl_selnum = 0;

function findtag(obj, tag)
{
	if(!isUndef(obj.getElementsByTagName))
	{
		return obj.getElementsByTagName(tag);
	}
	else if(obj.all && obj.all.tags)
	{
		return obj.all.tags(tag);
	}
	else
	{
		return null;
	}
}

function upl_newid()
{
	for(var n=0;n<=DEF_UploadOnceNum-1;n++)
	{
		if(!$id('file' + n))
		{
			return(n);
		}
	}
	return(-1);
}

function upl_add(id,oldid)
{
	if($id("upload" + id))return;
	var Node = $id('upload_node').cloneNode(true);
	Node.id="upload" + id;
	Node.style.display="";
	
	if($id('upload_del' + oldid))$id('upload_del' + oldid).style.display="";
	if(id==-1)return;
	
	tags = findtag(Node, 'input');
	for(i=0;i<tags.length;i++)
	{
		if(tags[i].name == 'file_number') 
		{
		tags[i].className = "fminpt uninit_upload";
		tags[i].name = 'file' + id;
		tags[i].id = 'file' + id;
		tags[i].onchange = 'upl_onchange(' + id + ')';
		tags[i].unselectable = 'on';
		}
		if(tags[i].name == 'text_number') 
		{
		tags[i].name = 'text' + id;
		tags[i].id = 'text' + id;
		}
	}
	Node.innerHTML = Node.innerHTML.replace("_number", id).replace("_number", id);
	Node.innerHTML+="<span id=upload_del" + id + " style=display:none><a href=#none onclick=\"upl_remove(this.parentNode.parentNode);\">ɾ��</a> <a href=#none onclick=\"addcontent(1,'[upload=" + id + "]');\">����</a></span>";
	$id('new_upload').appendChild(Node);
	init_uploadform();
	Upl_selnum++;
}

function upl_remove(id)
{
	$id('new_upload').removeChild(id);
	Upl_selnum--;
}
                
function upl_onchange(id)
{
	if(isNaN(id))id=id.substr(4,1);
	var file = $id('file' + id).value;
	var ext = file.lastIndexOf('.') == -1 ? '' : file.substr(file.lastIndexOf('.')+1, file.length).toLowerCase();
	if(DEF_UploadFileType.indexOf(':.'+ext+':') == -1 || ext == '')
	{
		alert('��֧���ϴ������͸���!');
		upl_remove($id('upload' + id));
		upl_add(upl_newid(),-1);
		return;
	}
	
	if (document.all||document.getElementById)
	{
		var theform = $id("LeadBBSFm");
		for (i=0;i<theform.length;i++)
		{
			var tempobj=theform.elements[i];
			if(tempobj.type.toLowerCase()=="file" && tempobj!=$id('file' + id))
			{
				if(tempobj.value == $id('file' + id).value)
				{
					alert('���ļ������ϴ��б���.');
					upl_remove($id('upload' + id));
					upl_add(upl_newid(),-1);
					return;
				}
			}
		}
	}
	upl_add(upl_newid(),id);
}

function Upl_IO(ur,lb,id)
{
	Upl_Level += 1;
	getAJAX("Inc/Upload_Info.asp?id=<%=Urlencode(GBL_CHK_User)%>&tt=" + Math.random(),"","Upl_IO_processor(tmp);",1);
	window.clearTimeout(Upl_IOfun);
	if(Upl_Level<2)Upl_IOfun = window.setTimeout(Upl_IO,Upl_GetDelay);
	Upl_Level -= 1;
}

function Upl_IO_processor(str)
{
	if(str=="busy"){window.clearTimeout(Upl_IOfun);return;}
	var tp="upload_doc";
	var tmp;
	if(str!=" ")
	{
		Upl_Start = true;
		tmp = str.split(" ");
		if(tmp.length>=3)
		{
			str = "�ϴ��ٷֱ�: " + parseInt(tmp[0]/tmp[1]*100) + "% " + " ���ϴ� " + parseInt(tmp[0]/1024) + "K ����ʱ�� " + parseInt(tmp[2]) + " ��"
		}
		$id(tp).innerHTML = str + " ";
	}
	else
	{
		if(Upl_Start)
		{
			$id(tp).innerHTML = "�����ϴ���ɣ����Ժ�...";
			Upl_Level=9999;window.clearTimeout(Upl_IOfun);
		}
		else
		{
			$id(tp).innerHTML = "�����ϴ����������Ժ�...";
		}
	}
}

function Upl_getHttp()
{
	var oT = false;
	try
	{
		oT=new XMLHttpRequest;
	}
	catch(e)
	{
		try
		{
			oT=new ActiveXObject("MSXML2.XMLHTTP");
		}
		catch(e2)
		{
			try
			{
				oT=new ActiveXObject("Microsoft.XMLHTTP");
			}
			catch(e3)
			{
				oT=false;
			}
		}
	}
	return(oT);
}

function Upl_submit(){
	if(Upl_selnum<1)return;
	Upl_Start = false;Upl_Level=0;
	Upl_IOfun = window.setTimeout(Upl_IO,100);
}
</script>
<%if refer = 0 Then%>
</td>
</tr>
<%
end if

End Sub

Sub DisplayUploadEdit

	Dim N
	If UploadListNum < 1 Then Exit Sub
	Response.Write "<div><b>�༭����</b></div>"
	For N = 0 to UploadListNum - 1
		Response.Write "<div style=""margin-top:5px;"">"
		Response.WRite "����<input class=fmchkbox type=checkbox name=filedel" & UploadListData(0,N) & " value=1 checked>"
		Response.Write "<a href=#no onclick=""$id('fileedit_" & UploadListData(0,N) & "').style.display='';this.style.display='none';"">�޸�</a> <span id=fileedit_" & UploadListData(0,N) & " style=display:none><span><input name=fileedit" & UploadListData(0,N) & " type=file id=fileedit" & UploadListData(0,N) & " size=15 class=""fminpt uninit_upload""></span></span>"
		Response.Write " ע�� <input name=textedit" & UploadListData(0,N) & " type=text value=""" & htmlencode(UploadListData(8,N)) & """ maxlength=" & upload_NoteLength & " size=20 class='fminpt input_2 note'> " & htmlencode(UploadListData(6,N)) & "</div>"
	Next
	Response.Write "<br>"

End Sub

Sub GetAncUploaInfo

	Dim Rs
	Set Rs = LDExeCute("Select ID,UserID,PhotoDir,SPhotoDir,ndatetime,FileType,FileName,FileSize,Info,AnnounceID,BoardID from " & UploadTable & " where AnnounceID=" & Form_EditAnnounceID,0)
	If Not Rs.Eof Then
		UploadListData = Rs.GetRows(-1)
		UploadListNum = Ubound(UploadListData,2) + 1
	Else
		UploadListNum = 0
	End If
	Rs.Close
	Set Rs = Nothing

End Sub

Class Upload_Save

Private Upd_SaveName_Small,Upd_SaveName,Upd_Extend,PhotoDir,UploadPhotoUrl_Small,UploadPhotoUrl,SQLArr,SQLArrNum,EnableUpload,TodayNum,UploadList,Save_Type
Private UploadProcessNum,Upload_ViewType
Public Upd_ErrInfo,Upd_FileInfo

Private Sub Class_Initialize

	Upload_ViewType = 1 '�ϴ���ʾ��ʽ: 1.����file.asp 0.ֱ����ʾͼƬ��ַ
	Upd_ErrInfo = ""
	Upd_FileInfo = 0
	EnableUpload = 1
	Upd_SpendFlag = 0
	TodayNum = 0
	UploadList = "0"
	
	Dim N
	N = 1
	Select Case DEF_EnableUpload
		Case 0: N = 0
		case 2: If CheckSupervisorUserName = 0 Then N = 0
		Case 3: If GBL_BoardMasterFlag < 4 Then N = 0
		Case 4: If GetBinarybit(GBL_CHK_UserLimit,2) = 0 Then N = 0
		Case 5: If GBL_BoardMasterFlag < 4 and GetBinarybit(GBL_CHK_UserLimit,2) = 0 Then N = 0
	End Select
	
	If DEF_Upd_SpendFlag = 0 and GBL_BoardMasterFlag >=4 Then
		Upd_SpendFlag = 0
	Else
		Upd_SpendFlag = 1
	End If
	If N = 1 and (GBL_CHK_OnlineTime >= DEF_NeedOnlineTime or DEF_NeedOnlineTime = 0) Then
	Else
		EnableUpload = 0
		Exit Sub
	End If

	If Upd_SpendFlag = 1 and DEF_UploadSpendPoints > 0 and DEF_UploadSpendPoints > GBL_CHK_Points Then
		Upd_ErrInfo = Upd_ErrInfo & DEF_PointsName(0) & "����(��Ҫ" & DEF_PointsName(0) & DEF_UploadSpendPoints & ")!');"
		EnableUpload = 0
		Exit Sub
	End If

	If UploadOneDayMaxNum > 0 and CheckSupervisorUserName = 0 Then
		Dim Rs,Num
		Num = 0
		Set Rs = LDExeCute(sql_select("Select NdateTime from " & UploadTable & " where UserID=" & GBL_UserID & " order by id DESC",UploadOneDayMaxNum),0)
		If Not Rs.Eof Then
			Do While Not Rs.Eof
				Num = RestoreTime(Rs(0))
				If Year(Num) = Year(DEF_Now) and Month(Num) = Month(DEF_Now) and Day(Num) = Day(DEF_Now) Then
					TodayNum = TodayNum + 1
				Else
					Exit do
				End If
				Rs.MoveNext
			Loop
			If TodayNum > UploadOneDayMaxNum Then
				Upd_ErrInfo = Upd_ErrInfo & "�޷����ϴ�������ÿ�����ɴ�" & UploadOneDayMaxNum & "������!"
				EnableUpload = 0
			End If
		End If
		Rs.Close
		Set Rs = Nothing
	End If

End Sub

Private Function GetSaveFileName(fname)

	UploadPhotoUrl_Small = ""
	UploadPhotoUrl = ""

	Dim TempNum,Temp,name
	name = Lcase(fname)
	name = "1" & Mid(name,inStrRev(name,"."))
	Upd_Extend = Trim(Mid(name,inStrRev(name,".")))

	If inStr(DEF_UploadFileType,":" & Upd_Extend & ":") < 1 Then Upd_Extend = ".LeadBBS"
	If inStr(":.htw:.ida:.asp:.asa:.idq:.cer:.cdx:.htr:.idc:.shtm:.shtml:.stm:.printer:.asax:.ascx:.ashx:.asmx:.aspx:.axd:.vsdisco:.rem:.soap:.config:.cs:.csproj:.vb:.vbproj:.webinfo:.licx:.resx:.resources:.php:.cgi:",":" & Upd_Extend & ":") Then Upd_Extend = ".LeadBBS"

	TempNum = Right("0" & day(DEF_Now),2) & "_" & Right(GetTimeValue(DEF_Now),6)

	GetSaveFileName = TempNum & Upd_Extend
	Upd_SaveName_Small = TempNum & "s" & Upd_Extend
	
	'On Error Resume Next
	Dim FSFlag
	FSFlag = 1
	If DEF_FSOString <> "" Then
		Err.Clear
		Set fs = Server.CreateObject(DEF_FSOString)
		If Err Then
			FSFlag = 0
			Err.Clear
			Set fs = Nothing
		End If
	Else
		FSFlag = 0
	End If

	If FSFlag = 0 Then
		Err.Clear
		GetSaveFileName = Left(GetTimeValue(DEF_Now),8) & GetSaveFileName
		Upd_SaveName_Small = Left(GetTimeValue(DEF_Now),8) & Upd_SaveName_Small
		PhotoDir = Server.MapPath(PhotoDirectory) & "\"
		Set Fs = Nothing
		Exit Function
	End If

	Dim TDir,FS
	TDir = Server.MapPath(PhotoDirectory) & "\"
	If Not FS.FolderExists(TDir) then
		GetSaveFileName = 0
		Upd_ErrInfo = Upd_ErrInfo & "<br>�������Ŀ¼��������ϵվ��!"
	End If
	
	TDir = TDir & year(DEF_Now) & "\"
	UploadPhotoUrl = UploadPhotoUrl & year(DEF_Now) & "/"
	If Not FS.FolderExists(TDir) then
		FS.CreateFolder(TDir)
	End If

	TDir = TDir & Right("0" & month(DEF_Now),2) & "\"
	UploadPhotoUrl = UploadPhotoUrl & Right("0" & month(DEF_Now),2) & "/"
	If Not FS.FolderExists(TDir) then
		FS.CreateFolder(TDir)
	End If
	
	'TDir = TDir & Right("0" & day(DEF_Now),2) & "\"
	'UploadPhotoUrl = UploadPhotoUrl & Right("0" & day(DEF_Now),2) & "/"
	'If Not FS.FolderExists(TDir) then
	'	FS.CreateFolder(TDir)
	'End If
	
	PhotoDir = TDir

	Dim Exist
	Exist = 0
	If FS.FileExists(TDir & GetSaveFileName) Then
		Exist = 1
	Else
		If FS.FileExists(TDir & Replace(GetSaveFileName,Upd_Extend, "s.jpg")) Then
			Exist = 1
		End if
	End If
	If Exist = 1 then
		For Temp = 0 To 99
			GetSaveFileName = TempNum & "_" & Temp & Upd_Extend
			Upd_SaveName_Small = TempNum & "_" & Temp & "s" & Upd_Extend

			Exist = 0
			If FS.FileExists(TDir & GetSaveFileName) Then
				Exist = 1
			Else
				If FS.FileExists(TDir & Replace(GetSaveFileName,Upd_Extend, "_s.jpg")) Then Exist = 1
			End If
			
			If Exist = 1 then
			Else
				Set FS = Nothing
				Exit For
			End If
		Next
		Set FS = Nothing
	Else
		Set FS = Nothing
	End If

End Function

Private Sub initUploadArr

	'7 �Ƿ�Ϊ�༭(0 ���ϴ� 1�༭) 8.file��ID�� 9.�༭���Ƿ�ɾ��
	Dim N,start
	If EditFlag = 0 or UploadListNum = 0 Then
		ReDim SQLArr(9,DEF_UploadOnceNum)
		UploadProcessNum = DEF_UploadOnceNum - 1
		For N = 0 to UploadProcessNum
			SQLArr(7,N) = 0
			SQLArr(8,N) = N
			SQLArr(9,N) = 0
		Next
	Else
		start = DEF_UploadOnceNum - UploadListNum
		If start < 0 Then start = 0
		ReDim SQLArr(9,start + DEF_UploadOnceNum)

		For N = 0 to start - 1
			SQLArr(7,N) = 0
			SQLArr(8,N) = N
			SQLArr(9,N) = 0
		Next
		UploadProcessNum = start + UploadListNum - 1
		For N = start to UploadProcessNum
			SQLArr(7,N) = 1
			SQLArr(8,N) = UploadListData(0,n-start)
			If Form_UpClass.form("filedel" & UploadListData(0,n-start)) <> "1" Then
				SQLArr(9,N) = 1
			Else
				SQLArr(9,N) = 0
			End If
		Next
	End If

End Sub

Public Sub Upload_File
	
	If EnableUpload = 0 Then Exit Sub
	Dim FileType,File,FileName,N,FileSize,Tmp,Info
	Dim TmpPoints
	TmpPoints = GBL_CHK_Points
	initUploadArr

	dim infoLen
	If lcase(UploadTable) = "leadbbs_upload" then
		infoLen = 30
	else
		upload_NoteLength = 255
		infoLen = 255
	end if

	For N = 0 to UploadProcessNum
		If SQLArr(7,N) = 1 Then
			Set file = Form_UpClass.file("fileedit" & SQLArr(8,N))
		Else
			Set file = Form_UpClass.file("file" & SQLArr(8,N))
		End If
		FileName = Trim(file.fileName)
		If inStr(FileName,".") = 0 and file.FileSize > 0 Then FileName = "leadbbs.rar"
		If StrLength(FileName) > 50 Then FileName = LeftTrue(FileName,25) & "___" & Right(FileName,11)

		SQLArr(0,N) = 0
		SQLArr(1,N) = ""
		SQLArr(2,N) = ""
		SQLArr(3,N) = 2
		SQLArr(4,N) = ""
		SQLArr(5,N) = 0
		SQLArr(6,N) = ""

		'�༭ɾ������ļ��������±���
		If SQLArr(7,N) = 1 Then
			Info = LeftTrue(Form_UpClass.form("textedit" & SQLArr(8,N)),infoLen)
		Else
			Info = LeftTrue(Form_UpClass.form("text" & SQLArr(8,N)),infoLen)
		End If
		If FileName <> "" Then
			FileType = LCase(file.FileType)
			Tmp = InStr(FileType,"/")
			If Tmp > 0 Then
				Tmp = Left(FileType,Tmp-1)
			Else
				Tmp = FileType
			End If
			'0-image 1.flash 2.other 3.text 4.audio 5.viedo
			Save_Type = 2
			Upd_FileInfo = 7
			Select Case Tmp
			Case "image":
				If inStr(FileType,"pjpeg") or inStr(FileType,"jpeg") or inStr(FileType,"gif") or inStr(FileType,"png") Then Save_Type = 0
				Upd_FileInfo = 4
				If FileName = "leadbbs.rar" Then FileName = "leadbbs.gif"
			Case "text": Save_Type = 3
				Upd_FileInfo = 8
				If FileName = "leadbbs.rar" Then FileName = "leadbbs.txt"
			Case "audio": if FileType = "audio/x-ms-wma" or FileType = "audio/mp3" or FileType = "audio/wav" or FileType = "audio/mid" or FileType = "audio/mpeg" Then Save_Type = 4
				Upd_FileInfo = 1
				If FileName = "leadbbs.rar" Then FileName = "leadbbs.wma"
			Case "video": if FileType = "video/x-msvideo" or FileType = "video/avi" or FileType = "video/x-ms-asf" or FileType = "video/x-ms-wmv" Then Save_Type = 5
				Upd_FileInfo = 1
				If FileName = "leadbbs.rar" Then FileName = "leadbbs.wma"
			Case "application": if FileType = "application/x-shockwave-flash" Then
						Save_Type = 1
						Upd_FileInfo = 3
						If FileName = "leadbbs.rar" Then FileName = "leadbbs.swf"
					Else
						Upd_FileInfo = 7
					End If
			Case else:
				Upd_FileInfo = 7
			End Select
			FileSize = file.FileSize
			
			Upd_SaveName = GetSaveFileName(FileName)
			UploadPhotoUrl_Small = UploadPhotoUrl & Upd_SaveName_Small
			UploadPhotoUrl = UploadPhotoUrl & Upd_SaveName

			If Save_Type = 0 and FileSize > 2097152 Then 'ͼƬ���ֻ����2M
				Upd_ErrInfo = Upd_ErrInfo & "<br>����(ͼƬ) " & HtmlEncode(Upd_SaveName) & " ������С���ϴ�ʧ��!"
			ElseIf FileSize > DEF_FileMaxBytes Then
				Upd_ErrInfo = Upd_ErrInfo & "<br>���� " & HtmlEncode(Upd_SaveName) & " ������С���ϴ�ʧ��!"
			ElseIf FileSize < 1 Then
				Upd_ErrInfo = Upd_ErrInfo & "<br>���� " & HtmlEncode(Upd_SaveName) & " Ϊ�գ��ϴ�ʧ��!"
			ElseIf inStr(DEF_UploadFileType,":" & Upd_Extend & ":") < 1 Then
				Upd_ErrInfo = Upd_ErrInfo & "<br>���� " & HtmlEncode(Upd_SaveName) & " ���ʹ����ϴ�ʧ��!"
			Else
				If Upd_SpendFlag = 1 and DEF_UploadSpendPoints > 0 and DEF_UploadSpendPoints > TmpPoints Then
					Upd_ErrInfo = Upd_ErrInfo & "<br>���� " & HtmlEncode(Upd_SaveName) & " �ϴ�ʧ��(" & DEF_PointsName(0) & "����)!"
				ElseIf UploadOneDayMaxNum > 0 and TodayNum >= UploadOneDayMaxNum Then
					Upd_ErrInfo = Upd_ErrInfo & "<br>���� " & HtmlEncode(Upd_SaveName) & " �ϴ�ʧ��(�������ϴ���������)!"
				Else
					TmpPoints = TmpPoints - DEF_UploadSpendPoints
					TodayNum = TodayNum + 1

					file.saveas PhotoDir & Upd_SaveName
					If Save_Type = 0 Then
						ProcessFile
					Else
						UploadPhotoUrl_Small = ""
					End If
					SQLArr(0,N) = GBL_UserID
					SQLArr(1,N) = UploadPhotoUrl
					SQLArr(2,N) = UploadPhotoUrl_Small
					SQLArr(3,N) = Save_Type
					SQLArr(4,N) = FileName
					SQLArr(5,N) = FileSize
					SQLArr(6,N) = Info
				End If
			End If
		ElseIf SQLArr(7,N) = 1 and SQLArr(9,N) = 0 and Info <> "" Then
			SQLArr(6,N) = Info
		End If
		Set file = Nothing
	Next
	Saved
	
	If Upd_FileInfo = 0 Then
		For N = 0 to Ubound(SQLArr,2)
			If SQLArr(0,N) > 0 Then
				Select Case SQLArr(3,N)
					Case 0:	Upd_FileInfo = 4 'image
					Case 1:	Upd_FileInfo = 3 'flash
					Case 3:	Upd_FileInfo = 8 'text
					Case 4:	Upd_FileInfo = 1 'audio
					Case 5: Upd_FileInfo = 1 'video
					Case Else: Upd_FileInfo = 7 'other
				End Select
				Exit For
			End If
		Next
	End If
	If Upd_FileInfo = 0 and UploadListNum > 0 Then
		For N = 0 to UploadListNum - 1
			If cCur(UploadListData(0,N)) > 0 Then
				Select Case UploadListData(5,N)
					Case 0:	Upd_FileInfo = 4 'image
					Case 1:	Upd_FileInfo = 3 'flash
					Case 3:	Upd_FileInfo = 8 'text
					Case 4:	Upd_FileInfo = 1 'audio
					Case 5: Upd_FileInfo = 1 'video
					Case Else: Upd_FileInfo = 7 'other
				End Select
				Exit For
			End If
		Next
	End If

End Sub

Public Sub UpdateUpload(id)

	CALL LDExeCute("Update " & UploadTable & " Set AnnounceID=" & id & " where id in(" & UploadList & ")",1)

End Sub

Private Sub DeleteUpload(id)

	CALL LDExeCute("Delete from " & UploadTable & " where id in(" & id & ")",1)

End Sub

Private Sub Saved

	Dim N,Num,SQL,UserID,Rs,AncID
	Num  = 0
	
	'ID,UserID,PhotoDir,SPhotoDir,ndatetime,FileType,FileName,FileSize,Info,AnnounceID,BoardID
	Dim DelContentStr
	Dim EditN
	EditN = 0

	Dim TmpHome
	TmpHome = Left(DEF_BBS_UploadPhotoUrl,1)
	If TmpHome = "/" or TmpHome = "\" Then
		TmpHome = ""
	Else
		TmpHome = DEF_BBS_HomeUrl
	End If

	If EditFlag = 1 Then
		AncID = Form_EditAnnounceID
	Else
		If Form_EditAnnounceID > 0 Then
			AncID = Form_EditAnnounceID
		Else
			AncID = 0
		End If
	End If

	Dim DelNum,re
	DelNum = 0
	For N = 0 to UploadProcessNum
		If SQLArr(7,N) = 1 Then '�༭����
			If SQLArr(9,N) = 1 Then 'ɾ���������ı༭����
				CALL ChangeUploadNum(UploadListData(1,EditN),-1,0)
				DelContentStr = "[upload=" & UploadListData(0,EditN) & "," & UploadListData(5,EditN) & "]" & UploadListData(6,EditN) & "[/upload]"
				DeleteUpload(UploadListData(0,EditN))
				If UploadListData(2,EditN) <> "" Then DeleteFiles(Server.MapPath(Replace(PhotoDirectory & UploadListData(2,EditN),"/","\")))
				If UploadListData(3,EditN) <> "" Then DeleteFiles(Server.MapPath(Replace(PhotoDirectory & UploadListData(3,EditN),"/","\")))
				Form_Content = Replace(Form_Content,VbCrLf & DelContentStr,"")
				Form_Content = Replace(Form_Content,DelContentStr,"")
				If Upload_ViewType <> 1 Then
					set re = New RegExp
					re.Global = True
					re.IgnoreCase = True		
					re.Pattern="\[IMG*([0-9=]*),(absmiddle|left|right|top|middle|bottom|absbottom|baseline|texttop)\](/|../|http://|https://|ftp://)(" & Replace(DEF_BBS_UploadPhotoUrl & UploadListData(2,EditN),"\","\\") & ")\[\/IMG]"
					Form_Content = re.Replace(Form_Content,"")
					Set re = Nothing
				End If
				UploadListData(0,EditN) = 0
			Else '�޸ĸ�����Ϣ
				If SQLArr(4,N) <> "" Then
					SQL = "Update " & UploadTable & " Set " &_
						"PhotoDir='" & Replace(SQLArr(1,N),"'","''") & "'" &_
						",SPhotoDir='" & Replace(SQLArr(2,N),"'","''") & "'" &_
						",ndatetime=" & GetTimeValue(DEF_Now) &_
						",FileType=" & SQLArr(3,N) &_
						",FileName='" & Replace(SQLArr(4,N),"'","''") & "'" &_
						",FileSize=" & SQLArr(5,N) &_
						",Info='" & Replace(SQLArr(6,N),"'","''") & "'" &_
					 	" where ID=" & UploadListData(0,EditN)
					CALL LDExeCute(SQL,1)
					If UploadListData(2,EditN) <> "" Then DeleteFiles(Server.MapPath(Replace(PhotoDirectory & UploadListData(2,EditN),"/","\")))
					If UploadListData(3,EditN) <> "" Then DeleteFiles(Server.MapPath(Replace(PhotoDirectory & UploadListData(3,EditN),"/","\")))
					DelContentStr = "[upload=" & UploadListData(0,EditN) & "," & UploadListData(5,EditN) & "]" & UploadListData(6,EditN) & "[/upload]"
					Form_Content = Replace(Form_Content,DelContentStr,"[upload=" & UploadListData(0,EditN) & "," & SQLArr(3,N) & "]" & SQLArr(4,N) & "[/upload]")

					If Upload_ViewType <> 1 and (SQLArr(3,N) = 0 or  UploadListData(5,EditN) = 0) Then
						Dim Tmp
						If SQLArr(3,N) = 0 Then
							Tmp = "[upload=" & UploadListData(0,EditN) & "," & SQLArr(3,N) & "]" & SQLArr(4,N) & "[/upload]"
							DelContentStr = "[img]" & TmpHome & DEF_BBS_UploadPhotoUrl & SQLArr(1,N) & "[/img]"
							Form_Content = Replace(Form_Content,VbCrLf & Tmp,DelContentStr)
							Form_Content = Replace(Form_Content,Tmp,DelContentStr)

							set re = New RegExp
							re.Global = True
							re.IgnoreCase = True
							re.Pattern="\[IMG*([0-9=]*),(absmiddle|left|right|top|middle|bottom|absbottom|baseline|texttop)\](/|../|http://|https://|ftp://)(" & Replace(DEF_BBS_UploadPhotoUrl & UploadListData(2,EditN),"\","\\") & ")\[\/IMG]"
							Form_Content = re.Replace(Form_Content,DelContentStr)
							Set re = Nothing
						Else
							set re = New RegExp
							re.Global = True
							re.IgnoreCase = True
							re.Pattern="\[IMG*([0-9=]*),(absmiddle|left|right|top|middle|bottom|absbottom|baseline|texttop)\](/|../|http://|https://|ftp://)(" & Replace(DEF_BBS_UploadPhotoUrl & UploadListData(2,EditN),"\","\\") & ")\[\/IMG]"
							DelContentStr = "[upload=" & UploadListData(0,EditN) & "," & SQLArr(3,N) & "]" & SQLArr(4,N) & "[/upload]"

							Form_Content = re.Replace(Form_Content,DelContentStr)
							Set re = Nothing
						End If
					End If
					UploadListData(5,EditN) = SQLArr(3,N)
				ElseIf SQLArr(6,N) <> "" Then
					SQL = "Update " & UploadTable & " Set " &_
					"Info='" & Replace(SQLArr(6,N),"'","''") & "'" &_
				 	" where ID=" & UploadListData(0,EditN)
				 	CALL LDExeCute(SQL,1)
				 End If
			End If
			EditN = EditN + 1
		Else '�������
			If SQLArr(4,N) <> "" Then
				Num = Num + 1
				UserID = SQLArr(0,N)
				SQL = "insert into " & UploadTable & "(UserID,PhotoDir,SPhotoDir,ndatetime,FileType,FileName,FileSize,Info,AnnounceID,BoardID) Values(" &_
					SQLArr(0,N) & ",'" & Replace(SQLArr(1,N),"'","''") & "','" & Replace(SQLArr(2,N),"'","''") & "'," &_
				 	GetTimeValue(DEF_Now) & "," & SQLArr(3,N) &_
				 	",'" & Replace(SQLArr(4,N),"'","''") & "'" &_
				 	"," & SQLArr(5,N) &_
				 	",'" & Replace(SQLArr(6,N),"'","''") & "'," & AncID & "," & GBL_Board_ID & "" &_
				 	")"
				CALL LDExeCute(SQL,1)
				
				SQL = sql_select("Select ID From " & UploadTable & " where UserID=" & SQLArr(0,N) & " order by ID DESC",1)
				Set Rs = LDExeCute(SQL,0)
				If Not Rs.Eof Then
					If UploadList = "0" Then UploadList = ""
					If UploadList = "" Then
						UploadList = Rs(0)
					Else
						UploadList = UploadList & "," & Rs(0)
					End If
					If inStr(Form_Content,"[upload=" & N & "]") Then
						If Upload_ViewType <> 1 and SQLArr(3,N) = 0 Then
							Form_Content = Replace(Form_Content,"[upload=" & N & "]","[img]" & TmpHome & DEF_BBS_UploadPhotoUrl & SQLArr(1,N) & "[/img]")
						Else
							Form_Content = Replace(Form_Content,"[upload=" & N & "]","[upload=" & Rs(0) & "," & SQLArr(3,N) & "]" & SQLArr(4,N) & "[/upload]")
						End If
					Else
						If Upload_ViewType <> 1 and SQLArr(3,N) = 0 Then
							SQL = "[img]" & TmpHome & DEF_BBS_UploadPhotoUrl & SQLArr(1,N) & "[/img]"
						Else
							SQL = "[upload=" & Rs(0) & "," & SQLArr(3,N) & "]" & SQLArr(4,N) & "[/upload]"
						End If
						Form_Content = Form_Content & VbCrLf & SQL
					End If
				End If
				Rs.Close
				Set Rs = Nothing
			End If
		End If
	Next
	If UserID > 0 Then
		CALL ChangeUploadNum(UserID,Num,1)
	End If

End Sub

Private Sub ChangeUploadNum(UserID,Num,Flag)

	If lcase(UploadTable) <> "leadbbs_upload" then exit sub
	Dim Temp,SQL
	If Flag = 1 Then 'add
		Temp = DEF_UploadSpendPoints
	Else
		Temp = DEF_UploadDeletePoints
	End If
	UserID = cCur(UserID)
	If UserID > 0 Then
		If Upd_SpendFlag = 1 and DEF_UploadSpendPoints <> 0 Then
			If Temp > 0 Then
				SQL = "Update LeadBBS_User Set UploadNum=UploadNum+" & Num & ",Points=Points-" & DEF_UploadSpendPoints*Num & " Where ID=" & UserID
			Else
				SQL = "Update LeadBBS_User Set UploadNum=UploadNum+" & Num & ",Points=Points+" & (0-DEF_UploadSpendPoints*Num) & " Where ID=" & UserID
			End If
			If UserID = GBL_UserID Then UpdateSessionValue 4,DEF_UploadSpendPoints*Num,1
		Else
			SQL = "Update LeadBBS_User Set UploadNum=UploadNum+" & Num & " Where ID=" & UserID
		End If
		CALL LDExeCute(SQL,1)
		SQL = "Update LeadBBS_SiteInfo Set UploadNum=UploadNum+" & Num
		CALL LDExeCute(SQL,1)
		UpdateStatisticDataInfo Num,5,1
	End If

End Sub

Private Sub ProcessFile

	Dim Temp,TmpName
	TmpName = "LeadBBS"
	If inStrRev(Upd_SaveName,".")>0 Then TmpName = Mid(Upd_SaveName,inStrRev(Upd_SaveName,".")+1)
	
	Temp = 0
	If DEF_EnableGFL = 1 Then
		Temp = SaveSmallPic(PhotoDir & Upd_SaveName,PhotoDir & Upd_SaveName_Small,DEF_UploadSwidth,DEF_UploadSheight,1)
		If Temp = 4 Then
			If inStrRev(UploadPhotoUrl_Small,".")>0 Then
				UploadPhotoUrl_Small = Left(UploadPhotoUrl_Small,inStrRev(UploadPhotoUrl_Small,".")) & "jpg"
			Else
				UploadPhotoUrl_Small = UploadPhotoUrl_Small & "jpg"
			End if
		ElseIf Temp = 3 Then
			If inStrRev(UploadPhotoUrl_Small,".")>0 Then
				UploadPhotoUrl_Small = Left(UploadPhotoUrl_Small,inStrRev(UploadPhotoUrl_Small,".")) & "gif"
			Else
				UploadPhotoUrl_Small = UploadPhotoUrl_Small & "gif"
			End if
		Else
			If Temp = 1 Then
				If inStrRev(UploadPhotoUrl_Small,".")>0 Then
					UploadPhotoUrl_Small = Left(UploadPhotoUrl_Small,inStrRev(UploadPhotoUrl_Small,".")) & "jpg"
				Else
					UploadPhotoUrl_Small = UploadPhotoUrl_Small & "jpg"
				End if
			End If
		End If
		If Temp = 1 or Temp = 3 or Temp = 4 Then
		Else
			If Temp <> 2 Then Save_Type = 2
			UploadPhotoUrl_Small = ""
		End If
	Else
		UploadPhotoUrl_Small = ""
	End If

End Sub

End Class
%>         