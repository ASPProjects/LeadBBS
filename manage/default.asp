<!-- #include file=../inc/BBSsetup.asp -->
<!-- #include file=../inc/Board_Popfun.asp -->
<!-- #include file=../inc/Limit_Fun.asp -->
<!-- #include file=inc/bbsmanage_Fun.asp -->
<%
DEF_BBS_HomeUrl = "../"
Dim GBL_ID
initDatabase
GBL_CHK_TempStr = ""
checkSupervisorPass
GBL_ID = GBL_UserID

Dim GBL_InPageFlag
If Request.QueryString <> "" Then
	GBL_InPageFlag = 1
Else
	GBL_InPageFlag = 0
End If

If GBL_InPageFlag = 1 Then
	Manage_sitehead DEF_SiteNameString & " - ����Ա",""
Else
	Manage_sitehead DEF_SiteNameString & " - ����Ա","frame_class"
End If

If GBL_CHK_Flag=1 Then
	LoginAccuessFul
Else
	DisplayLoginForm
End If
closeDataBase
Manage_Sitebottom("none")

Sub LoginAccuessFul


	Dim NewUrl
	NewUrl = DEF_BBS_HomeUrl
	If Left(NewUrl,3) = "../" or Left(NewUrl,3) = "..\" Then NewUrl = Mid(NewUrl,4)

	If GBL_InPageFlag = 1 Then
		Default_info
		Exit Sub
	End If
%>

<script>
	
	var nav_cursel = null;
	function nav_sel(obj)
	{
		if(nav_cursel && nav_cursel!=null && nav_cursel.parentNode)nav_cursel.parentNode.className="item";
		nav_cursel = obj;
		obj.parentNode.className="item_sel";
	}
	var nav_curassort = null;
	function nav_assortsel(n)
	{
		$id('nav_itemlist1').style.display="none";
		if(nav_curassort!=null)$id('nav_assort_' + nav_curassort).className="item";
		nav_curassort = n;
		$id('nav_assort_' + n).className="item_sel";
		$id('nav_itemlist0').innerHTML = $id('nav_itemlist' + n).innerHTML;
		nav_sel($id('nav_itemlist' + n + '_default'));
		$id('mainFrame').src = $id('nav_itemlist' + n + '_default').href;
	}
	//document.body.onselectstart = document.body.ondrag = function(){
    //return false;
	//}
</script>


	<div class="frame_top" id="topDataTd">
			<div class=managelogo><img src=pic/manage_title.gif></div>
			<div class=top_control>
				<a href="<%=NewUrl%>Default.asp?action=info" id="nav_assort_1" class="item_sel" target="mainFrame" onclick="nav_assortsel(1);">��ҳ</a>
				<a href="javascript:;" id="nav_assort_2" class="item" target="mainFrame" onclick="nav_assortsel(2);">�������</a>
				<a href="javascript:;" id="nav_assort_3" class="item" target="mainFrame" onclick="nav_assortsel(3);">����</a>
				<a href="javascript:;" id="nav_assort_4" class="item" target="mainFrame" onclick="nav_assortsel(4);">�û�</a>
				<a href="javascript:;" id="nav_assort_5" class="item" target="mainFrame" onclick="nav_assortsel(5);">���ݿ�</a>
				<a href="javascript:;" id="nav_assort_6" class="item" target="mainFrame" onclick="nav_assortsel(6);">���</a>
				<a href="javascript:;" id="nav_assort_7" class="item" target="mainFrame" onclick="nav_assortsel(7);">���</a>
				<a href="javascript:;" id="nav_assort_8" class="item" target="mainFrame" onclick="nav_assortsel(8);">����</a>
				<a href="javascript:;" id="nav_assort_9" class="item" target="mainFrame" onclick="nav_assortsel(9);">CMS</a>
			</div>
			
		    <div class=top_userinfo>
		    	&lt;<span class=item><b><%=GBL_CHK_User%></b></span>&gt;
		    	<span class="splitword"> | </span>
		    	<a href=<%=DEF_BBS_HomeUrl%>User/BoardMaster/Default.asp class=item target=_blank><%=DEF_PointsName(6)%></a>
		    	<span class="splitword"> | </span>
		    	<a href=<%=DEF_BBS_HomeUrl%>Boards.asp class=item>������ҳ</a>
		    	<span class="splitword"> | </span> 
		    	<a href=<%=DEF_BBS_HomeUrl%>User/login.asp?action=logout class=item>�˳�</a>
		    </div>
			
	</div>
	<div class="frame_topline">
		<div class="frame_topline1">
		</div>
			<div class="frame_topline2">				
			</div>
	</div>
	
	<div class="frame_leftbody" style="">
		<br />
		<div class="frame_leftcontent">
		<%Default_NavItem%>
		</div>
	</div>
	<div class="maincontent">
		<iframe src="Default.asp?action=info" name="mainFrame" id="mainFrame" hidefocus="" frameborder="no" scrolling="auto">
		</iframe>
	</div>
      
		
<%End Sub

Sub Default_info

	If CheckSupervisorUserName = 1 Then
		If LCase(Request.QueryString) <> "checkversion" Then
			DisplaySystemInfo
		Else
			Response.Clear
			Update_CheckVersion
			Response.End
		End If
	Else%>
		<p><br>
		�Ѿ��ɹ���¼��<br></p>
		<br><br>
	<%End If%>
	<br><br>

<%End Sub


Dim GBL_UpdateVersion '�ڲ��汾��
GBL_UpdateVersion = 0
Dim GBL_LeadBBS_Setup_Data '��ʱ��ȡ��SetupRID��¼��������

Function DisplaySystemInfo

	frame_TopInfo
	%>
	<div class=frametitlehead>��̳��Ϣһ��</div>
	<div class="frameline"><a href=default.asp?need=1773>����鿴�����װ���</a></div>
	<div class="frameline">������ʱ�䣺<%=now%>����̳(����ʱ��)ʱ�䣺<%=DEF_Now%></div>
	<div class="frameline">���������ͣ�<%=Request.ServerVariables("OS")%>[IP:<%=Request.ServerVariables("LOCAL_ADDR")%>]</div>
	<div class="frameline">�ű��������棺<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></div>
	<div class="frameline">վ������·����<%=request.ServerVariables("APPL_PHYSICAL_PATH")%></div>
	<div class="frameline">����IP��ַ��<%=GBL_IPAddress%></div>
	<div class="frameline"><%=ScriptEngine & " Version " & ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion & "." & ScriptEngineBuildVersion %></div>
	<%If Request.QueryString("need") = "1773" Then%>
	<div class="frameline">AspJpegͼ�������<%
	CheckObjInstalled("Persits.Jpeg")
	Response.write GBL_CHK_TempStr%></div>
	<div class="frameline">FSO�ı���д��<%
	CheckObjInstalled2(DEF_FSOString)
	Response.write GBL_CHK_TempStr%></div>
	<div class="frameline">���ݿ�ʹ�ã�<%
	CheckObjInstalled("adodb.connection")
	Response.write GBL_CHK_TempStr%></div>
	<div class="frameline">Jmail���֧�֣�<%
	CheckObjInstalled("JMail.SMTPMail")
	Response.write GBL_CHK_TempStr%></div>
	<div class="frameline">AspJpegͼ�������<%
	CheckObjInstalled("Persits.Jpeg")
	Response.write GBL_CHK_TempStr%></div>
	<div class="frameline">LeadBBSר�����֧�֣�<%
	CheckObjInstalled("leadbbs.bbsCode")
	Response.write GBL_CHK_TempStr%></div>
	<div class="frameline">������ϴ���Scripting.Dictionary <%
	CheckObjInstalled("Scripting.Dictionary")
	Response.write GBL_CHK_TempStr%>
	ADODB.Stream <%
	CheckObjInstalled("Scripting.Dictionary")
	Response.write GBL_CHK_TempStr%> (ȫ��֧�ֲ��������ϴ�)</div>
	<%End If%>
	<div class=frametitle>LeadBBS���¼��</div>
	<div class=frameline onclick="this.style.display='none';update_checkversion();"><a href="javascript:;" class="bluefont">���������</a></div>
	<div class=frameline id=checkversion></div>
	
	<div class=frametitle>LeadBBS�汾��Ϣ</div>
	<div class=frameline>����������LeadBBS�����ң�����SpiderMan(QQ:527274)</div>
	<div class=frameline>�汾��Ϣ��<a href=http://www.leadbbs.com target=_blank><b><span class=redfont><%=DEF_Version%>.<%=GBL_UpdateVersion%></span></b></a></div>
	

	<div class=frametitle>Ȩ�޲ο�</div>
	<ol class=listli>
		<li>�����û��޸�Ȩ�ޣ��������ͨ�û�����Ȩ�����Լ������ϼ����������<%=DEF_PointsName(8)%>��������Ա����������ƹ̶����̶ܹ����༭�������ӡ�</li>
		<li><%=DEF_PointsName(5)%>���κλ�Ա��������(��������Ա)��<%=DEF_PointsName(5)%>ר���ӵ����֤�ʸ����Ա�ſ��Խ��롣</li>
		<li>Ȩ�޵ȼ����򣺹���Ա-><%=DEF_PointsName(6)%>-><%=DEF_PointsName(8)%>->��ͨ��Ա������<%=DEF_PointsName(5)%>���������û���</li>
		<li>��ֹת�����ӹ��ܽ��԰���������Ч��</li>
		<li>Ĭ�ϰ���ӵ�б༭��������ɾ�����������̶�������Ȩ�ޣ��鿴�����ϴ��������������ɫ��������</li>
		<li><%=DEF_PointsName(6)%>��ӵ�а�����Ȩ���⣬��ӵ���̶ܹ���Ȩ�ޡ�</li>
		<li>����Աӵ��һ��Ȩ�ޣ����Է���html�﷨�����⼰�������ݣ������û�����̳һ�����ϡ�</li>
	</ol>
	<script>
	function update_checkversion()
	{
	$id('checkversion').innerHTML = "�����...";
	getAJAX("default.asp?checkversion","","checkversion",0);}
	</script>
	<%
	frame_BottomInfo

End Function


Sub Update_CheckVersion

	
	
	If Update_CheckSetupRIDExist(1002," and ClassNum=0") = 0 Then
		GBL_UpdateVersion = "20100101001"
	Else
		GBL_UpdateVersion = cCur(GBL_LeadBBS_Setup_Data(2,0))
	End If
	
Const NetFlag = 1
Const NetUrl = "http://update.u1.leadbbs.com/"
Const NativeDir = "Download/"
Const SplitString = "---NdetVeL---"
	Dim Update,CurFile,CurFile_Name,CurFile_Intro
	Dim FileList
	Dim m
	If NetFlag = 0 Then
		Update = ADODB_LoadFile(NativeDir & "update.txt")
	Else
		Update = BytesToBstr(Update_GetInternetFile(NetUrl & "update.txt"))
	End If
	If Update = "err" Then Exit Sub
	Update = Split(Update,VbCrLf)
	
	Dim UpdateFlag
	UpdateFlag = 0
	For M = 0 to Ubound(Update,1)
		If Trim(Update(M)) <> "" Then
			If inStr(Update(M),SplitString) > 0 Then
				CurFile = Split(Update(M),SplitString)
				CurFile_Name = CurFile(0)
				CurFile_Intro = " (" & CurFile(1) & ")"
			Else
				CurFile_Name = Update(M)
				CurFile_Intro = ""
			End If
			
			If isNumeric(CurFile_Name) = 0 Then CurFile_Name = 0
			CurFile_Name = cCur(CurFile_Name)
			If CurFile_Name > cCur(GBL_UpdateVersion) Then
				Response.Write "<div class=redfont>��⵽�²���<u>" & CurFile_Name & "</u>" & CurFile_Intro & "</div>"
				UpdateFlag = UpdateFlag + 1
			End If
		End If
	Next
	If UpdateFlag = 0 Then
		Response.Write "<div class=greenfont>������̳�������°汾��</div>"
	Else
		Response.Write "<div class=redfont>����" & UpdateFlag & "��������Ҫ���¡�</div>"
	End If

End Sub


Function Update_CheckSetupRIDExist(RID,extend)

	Dim Rs
	Set Rs = LDExeCute(sql_select("Select ID,RID,ValueStr,ClassNum,SaveData from LeadBBS_Setup where RID=" & RID & extend,1),0)
	If Rs.Eof Then
		Update_CheckSetupRIDExist = 0
		Set GBL_LeadBBS_Setup_Data = Nothing
		GBL_LeadBBS_Setup_Data = ""
	Else
		Update_CheckSetupRIDExist = 1
		GBL_LeadBBS_Setup_Data = Rs.GetRows(-1)
		GBL_LeadBBS_Setup_Data(2,0) = Trim(GBL_LeadBBS_Setup_Data(2,0))
	End If
	Rs.Close
	Set Rs = Nothing

End Function

Function BytesToBstr(body) 

	on error resume next
	If LenB(body) < 1 Then
		BytesToBstr = ""
		Exit Function
	End If
	dim objstream
	set objstream = Server.CreateObject("adodb.stream")
	with objstream
	.Type = 1
	.Mode = 3
	.Open
	.Write body 
	.Position = 0
	.Type = 2
	.Charset = "GB2312"
	
	'.Charset = "UTF-8"
	BytesToBstr = .ReadText
	.Close
	end with
	set objstream = nothing
	If Err and BytesToBstr = "" Then
		BytesToBstr = body
		Err.clear
	End If

End Function

Function Update_GetInternetFile(ur)

	Dim url
	Url = ur
	url = Left(url,5000)
	If url = "" Then Exit Function
	Dim xmlHttp
	Set xmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
	xmlHttp.setTimeouts 5000,5000,5000,15000
	xmlHttp.setOption 2, 13056
	xmlHttp.open "GET", url, False, "", "" 
	
	on error resume next
	xmlHttp.send()
	If Err Then
		Response.Write "<p>��������: <font color=red>" & err.description & "</font></p>"
		Err.clear
		Update_GetInternetFile = "err"
		Exit Function
	End If

	If xmlHttp.readystate = 4 then 
	'if xmlHttp.status=200 Then
		Update_GetInternetFile = xmlhttp.Responsebody
	'end if 
	Else 
		Update_GetInternetFile = "err"
	End If
	Set xmlHttp = Nothing

End Function

Sub Default_NavItem

	Dim NewUrl
	NewUrl = DEF_BBS_HomeUrl
	If Left(NewUrl,3) = "../" or Left(NewUrl,3) = "..\" Then NewUrl = Mid(NewUrl,4)
	%>
	<script language=javascript>
	function sss(obj)
	{
		if(obj.style.display == "none")
		{
			obj.style.display = "block";
		}
		else
		{
			obj.style.display = "none";
		}
	}
	</script>
	<div class="nav_itemlist" id="nav_itemlist0">
	</div>
	<div class="nav_itemlist" id="nav_itemlist1">
		<div class="item"><a href=<%=NewUrl%>Default.asp?action=info target="mainFrame" onclick="nav_sel(this);" id="nav_itemlist1_default"><span>����������ҳ</span></a></div>
		<div class="item"><a href=<%=NewUrl%>SiteManage/SiteSetup.asp target="mainFrame" onclick="nav_sel(this);"><span>ȫ�ֲ�������</span></a></div>
		<div class="item"><a href=<%=NewUrl%>SiteManage/UploadSetup.asp target="mainFrame" onclick="nav_sel(this);"><span>�ϴ���������</span></a></div>
		<div class="item"><a href=<%=NewUrl%>SiteManage/UbbcodeSetup.asp target="mainFrame" onclick="nav_sel(this);"><span>���ݱ����������</span></a></div>
		<div class="item"><a href=javascript:;><div class=nav_sepline></div></a></div>
		<div class="item"><a href=<%=NewUrl%>SiteManage/SiteInfo.asp?action=Side target="mainFrame" onclick="nav_sel(this);"><span>��������</span></a></div>
		<div class="item"><a href=<%=NewUrl%>User/IPManage.asp target="mainFrame" onclick="nav_sel(this);" title="��������IP�λ�ĳ��IP��ַ"><span>IP��ַ����</span></a></div>
		<div class="item"><a href=<%=NewUrl%>SiteManage/SiteLink.asp target="mainFrame" onclick="nav_sel(this);"><span>������̳����</span></a></div>
		<div class="item"><a href=javascript:;><div class=nav_sepline></div></a></div>
		<div class="item"><a href=<%=NewUrl%>User/SendMailList.asp target="mainFrame" onclick="nav_sel(this);"><span>�ʼ����ͼ�Ⱥ��</span></a></div>
		<div class="item"><a href=<%=NewUrl%>User/SendGroupMessage.asp target="mainFrame" onclick="nav_sel(this);"><span>��̳����ϢȺ��</span></a></div>
		<div class="item"><a href=javascript:;><div class=nav_sepline></div></a></div>
		
		
		<div class="item"><a href=<%=NewUrl%>SiteManage/SiteEditFileContent.asp?file=-1 target="mainFrame" onclick="nav_sel(this);"><span>�༭�û�ע��Э��</span></a></div>
		<div class="item"><a href=<%=NewUrl%>SiteManage/SiteEditFileContent.asp?file=-3 target="mainFrame" onclick="nav_sel(this);"><span>�༭��ϵ������Ϣ</span></a></div>
		<div class="item"><a href=javascript:;><div class=nav_sepline></div></a></div>
		
		<div class="item"><a href=<%=NewUrl%>SiteManage/SiteInfo.asp target="mainFrame" onclick="nav_sel(this);"><span>վ����Ϣ����վ�޸�</span></a></div>
		<div class="item"><a href=<%=NewUrl%>SiteManage/Space.asp target="mainFrame" onclick="nav_sel(this);" title="�鿴���ݿ⣬�ϴ��ļ�����̳��ռ�ÿռ��������Ҫʱ����ܽϳ�"><span>�鿴�ռ�ռ�����</span></a></div>
		<div class="item"><a href=<%=NewUrl%>update.asp?checkversion=checkversion&sure=1&submitflag=1 target="mainFrame" onclick="nav_sel(this);"><span>����Ƿ��а汾����</span></a></div>
		<div class="item"><a href=<%=NewUrl%>update.asp?sure=1 target="mainFrame" onclick="nav_sel(this);"><span>������չ����</span></a></div>
		<div class="item"><a href=<%=NewUrl%>update.asp?submitflag=1&sure=1 target="mainFrame" onclick="nav_sel(this);"><span>������չ����</span></a></div>
		<div class="item"><a href=<%=NewUrl%>update.asp?sure=1&checkversion=updateversion&submitflag=1 target="mainFrame" onclick="nav_sel(this);"><span>�������²���</span></a></div>
	</div>
	<div class="nav_itemlist" id="nav_itemlist2" style="display:none;">
		<div class="item"><a href=<%=NewUrl%>ForumCategory/ForumCategoryManage.asp target="mainFrame" onclick="nav_sel(this);" id="nav_itemlist2_default"><span>��̳�������</span></a></div>
		<div class="item"><a href=<%=NewUrl%>ForumCategory/ForumCategoryManage.asp?action=join target="mainFrame" onclick="nav_sel(this);"><span>�����̳����</span></a></div>
	</div>
	<div class="nav_itemlist" id="nav_itemlist3" style="display:none;">
		<div class="item"><a href=<%=NewUrl%>ForumBoard/ForumBoardManage.asp target="mainFrame" onclick="nav_sel(this);" id="nav_itemlist3_default"><span>��̳�������</span></a></div>
		<div class="item"><a href=<%=NewUrl%>ForumBoard/ForumBoardJoin.asp target="mainFrame" onclick="nav_sel(this);"><span>�����̳����</span></a></div>
		<div class="item"><a href=<%=NewUrl%>ForumBoard/ForumBoardAssort.asp target="mainFrame" onclick="nav_sel(this);"><span>��̳����ר������</span></a></div>
		<div class="item"><a href=<%=NewUrl%>ForumBoard/MakeBoardList.asp target="mainFrame" onclick="nav_sel(this);"><span>������̳�б��޸�</span></a></div>
		<div class="item"><a href=<%=NewUrl%>ForumBoard/RepairYesterdayAnc.asp target="mainFrame" onclick="nav_sel(this);"><span>���¼������շ�����</span></a></div>
	</div>
	<div class="nav_itemlist" id="nav_itemlist4" style="display:none;">
		<div class="item"><a href=<%=NewUrl%>User/UserManage.asp target="mainFrame" onclick="nav_sel(this);" id="nav_itemlist4_default" title="��̳�����û��б��ṩǿ���޸ļ�ǿ��ָ��Ȩ�޷���"><span>��̳�û�����</span></a></div>
		<div class="item"><a href=<%=NewUrl%>User/UserSpecial.asp target="mainFrame" onclick="nav_sel(this);" title="�����û������������������û���������û��ȵȣ���ϸ�����������"><span>�����û�����</span></a></div>
		<div class="item"><a href=<%=NewUrl%>User/UserSetup.asp target="mainFrame" onclick="nav_sel(this);" title="�趨���û�ע���ѡ��"><span>�û�ע���������</span></a></div>
		<div class="item"><a href=<%=NewUrl%>User/UserJoin.asp target="mainFrame" onclick="nav_sel(this);" title="ǿ�����һ�����û�����ʹǰ̨�ر���ע�Ṧ��"><span>������û�</span></a></div>
		<div class="item"><a href=<%=NewUrl%>User/DeleteForbidIPandUser.asp target="mainFrame" onclick="nav_sel(this);"  title="ĳЩ���Σɣе�ַ�����ε��û���<%=DEF_PointsName(5)%>�е������ޣ�ÿ�����ֹ�ִ��һ�Σ��������"><span>������������û���IP</span></a></div>
		<div class="item"><a href=<%=NewUrl%>User/ClearOnlineUser.asp target="mainFrame" onclick="nav_sel(this);" title=�����������û���ʱ������><span>���������û�/�û�����</span></a></div>
	</div>
	
	<div class="nav_itemlist" id="nav_itemlist5" style="display:none;">
		<div class="item"><a href=<%=NewUrl%>Database/ExecuteString.asp target="mainFrame" onclick="nav_sel(this);" id="nav_itemlist5_default"><span>ֱ��ִ��SQL���</span></a></div>
		<%If DEF_UsedDataBase = 0 Then%>
		<div class="item"><a href=<%=NewUrl%>Database/FullTextManage.asp target="mainFrame" onclick="nav_sel(this);"><span>ȫ�ļ��������ݿ����</span></a></div>
		<%End If%>
		<%If DEF_UsedDataBase = 1 Then%>
		<div class="item"><a href=<%=NewUrl%>Database/BackupDatabase.asp target="mainFrame" onclick="nav_sel(this);"><span>���ݿⱸ�ݼ�ѹ��</span></a></div>
		<%End If%>
	</div>
	<div class="nav_itemlist" id="nav_itemlist6" style="display:none;">
		<div class="item"><a href=<%=NewUrl%>SiteManage/TempletManage.asp target="mainFrame" onclick="nav_sel(this);" id="nav_itemlist6_default"><span>��̳ģ�����</span></a></div>
		<div class="item"><a href=<%=NewUrl%>SiteManage/SiteEditFile.asp target="mainFrame" onclick="nav_sel(this);"><span>���߱༭������趨</span></a></div>
		<div class="item"><a href=<%=NewUrl%>SiteManage/DefineStyleParameter.asp?action=extentskin_manage target="mainFrame" onclick="nav_sel(this);"><span>��չ����趨</span></a></div></a></div>
	</div>
	<div class="nav_itemlist" id="nav_itemlist7" style="display:none;">
		<div class="item"><a href=<%=NewUrl%>SiteManage/SiteLink.asp?SiteLink_Flag=10 target="mainFrame" onclick="nav_sel(this);" id="nav_itemlist7_default"><span>���������</span></a></div>
		<div class="item"><a href=<%=NewUrl%>SiteManage/SiteInfo.asp?action=admanage target="mainFrame" onclick="nav_sel(this);"><span>�ۺϹ����λ����</span></a></div>
	</div>
	
	<div class="nav_itemlist" id="nav_itemlist8" style="display:none;">
		<div class="item"><a href=<%=NewUrl%>BlockUpdate/BlockUpdate.asp target="mainFrame" onclick="nav_sel(this);" id="nav_itemlist8_default"><span>�����޸���̳����</span></a></div>
		<div class="item"><a href=<%=NewUrl%>SiteManage/DeleteAllTopAnnounce.asp target="mainFrame" onclick="nav_sel(this);"><span>����̶ܹ���(���޸�)</span></a></div>
		<div class="item"><a href=<%=NewUrl%>SiteManage/RepairSite.asp target="mainFrame" onclick="nav_sel(this);"><span>�ϴ�·��/�û��������޸�</span></a></div>
		<div class="item"><a href=<%=NewUrl%>BlockUpdate/BlockUpdate.asp?action=blockdelete target="mainFrame" onclick="nav_sel(this);"><span>��̳��������ɾ��</span></a></div>
		
		<%
		If application(DEF_MasterCookies & "SiteEnableFlagzoieiu") = 1 or application(DEF_MasterCookies & "SiteDisbleWhyszoieiu") = "" Then
		%>
		<div class="item"><a href=<%=NewUrl%>SiteManage/SiteOpenClose.asp?Flag=close target="mainFrame" onclick="nav_sel(this);"><span>��ͣ��̳����</span></a></div>
		<%
		Else
		%>
		<div class="item"><a href=<%=NewUrl%>SiteManage/SiteOpenClose.asp?Flag=open target="mainFrame" onclick="nav_sel(this);"><span>������̳����</span></a></div>
		<%
		End If%>
		<div class="item"><a href=<%=NewUrl%>SiteManage/SiteReset.asp?Flag=open target="mainFrame" onclick="nav_sel(this);"><span>����������̳</span></a></div>
		
		<div class="item"><a href=<%=NewUrl%>SiteManage/SiteInfo.asp?action=MoreSV target="mainFrame" onclick="nav_sel(this);"><span>��̳��չ����</span></a></div>
		
		<div class="item"><a href=javascript:;><div class=nav_sepline></div></a></div>
		<div class="item"><a href=<%=NewUrl%>SiteManage/ForumLog.asp target="mainFrame" onclick="nav_sel(this);"><span>��̳��־����</span></a></div>
		<div class="item"><a href=<%=NewUrl%>SiteManage/ForumLog.asp?clear=yes target="mainFrame" onclick="nav_sel(this);"><span>�������ǰ����̳��־</span></a></div>
		
		<div class="item">
			<a href=javascript:;><div class=nav_sepline></div></a>
		</div>
		
		<div class="item"><a href=<%=DEF_BBS_HomeUrl%>Search/UploadList.asp target="_blank" onclick="nav_sel(this);"><span>�ϴ���������</span></a></div>
		<div class="item"><a href=<%=DEF_BBS_HomeUrl%>User/MyInfoBox.asp?AllPrinting=Yesing target="_blank" onclick="nav_sel(this);"><span>����Ϣ����</span></a></div>
		<div class="item"><a href=<%=DEF_BBS_HomeUrl%>User/LookUserInfo.asp?Evol=bag target="_blank" onclick="nav_sel(this);"><span>�ղ����ӹ���</span></a></div>
		<div class="item"><a href=<%=DEF_BBS_HomeUrl%>User/SendMessage.asp?pub=1 target="_blank" onclick="nav_sel(this);"><span>��������</span></a></div>
	</div>
	
	
	
	<div class="nav_itemlist" id="nav_itemlist9" style="display:none;">
		<div class="item"><a href=<%=DEF_BBS_HomeUrl%>article/center.asp?action=newsclass&list=1 target="mainFrame" onclick="nav_sel(this);" id="nav_itemlist9_default"><span>�������</span></a></div>
		<div class="item"><a href=<%=DEF_BBS_HomeUrl%>article/center.asp?action=newsclass target="mainFrame" onclick="nav_sel(this);"><span>�������</span></a></div>
		<div class="item"><a href=javascript:;><div class=nav_sepline></div></a></div>
		<div class="item"><a href=<%=DEF_BBS_HomeUrl%>article/center.asp?action=newsarticle target="mainFrame" onclick="nav_sel(this);"><span>�������</span></a></div>
		<div class="item"><a href=<%=DEF_BBS_HomeUrl%>article/center.asp?action=newsmanage target="mainFrame" onclick="nav_sel(this);"><span>���¹���</span></a></div>
		<div class="item"><a href=javascript:;><div class=nav_sepline></div></a></div>
		
		<div class="item"><a href=<%=DEF_BBS_HomeUrl%>article/center.asp?action=setchannel target="mainFrame" onclick="nav_sel(this);"><span>������ҳ��Ŀ����</span></a></div>
		<div class="item"><a href=<%=DEF_BBS_HomeUrl%>article/center.asp?action=editfile&form_fileid=0 target="mainFrame" onclick="nav_sel(this);"><span>�༭��ҳͼƬ����</span></a></div>
		<div class="item"><a href=<%=DEF_BBS_HomeUrl%>article/center.asp?action=editfile&form_fileid=1 target="mainFrame" onclick="nav_sel(this);"><span>�Զ�����վ�ײ���Ϣ</span></a></div>
		<div class="item"><a href=<%=DEF_BBS_HomeUrl%>article/center.asp?action=editfile&form_fileid=2 target="mainFrame" onclick="nav_sel(this);"><span>CSS��ʽ��</span></a></div>
		<div class="item"><a href=javascript:;><div class=nav_sepline></div></a></div>
		<div class="item"><a href=<%=DEF_BBS_HomeUrl%>article/center.asp?action=updatecache target="mainFrame" onclick="nav_sel(this);"><span>��������ϵͳ����</span></a></div>
	</div>
	<script>
	var nav_curassort = 1;
	nav_sel($id("nav_itemlist1_default"));
	</script>
<%

End Sub

Function CheckObjInstalled2(strClassString)

	On Error Resume Next
	Dim Temp
	Err = 0
	Dim TmpObj
	Set TmpObj = CreateObject(strClassString)
	Temp = Err
	If Temp = 0 Then
		CheckObjInstalled2 = True
		GBL_CHK_TempStr = "<font color=green class=greenfont>��</font>"
	ElseIf Temp = -2147221005 Then
		GBL_CHK_TempStr = "<font color=red class=redfont>���δ��װ</font>"
		CheckObjInstalled2 = False
	ElseIf Temp = -2147221477 Then
		GBL_CHK_TempStr = "<font color=green class=greenfont>��֧�ִ����</font>"
		CheckObjInstalled2 = True
	ElseIf Temp = 1 Then
		GBL_CHK_TempStr = "<font color=red class=redfont>��δ֪�Ĵ����������δ��ȷ��װ</font>"
		CheckObjInstalled2 = False
	End If
	Err.Clear
	Set TmpObj = Nothing
	Err = 0

End Function

Function CheckObjInstalled(strClassString)

	On Error Resume Next
	Dim Temp
	Err = 0
	Dim TmpObj
	Set TmpObj = Server.CreateObject(strClassString)
	Temp = Err
	If Temp = 0 Then
		CheckObjInstalled = True
		GBL_CHK_TempStr = "<font color=green class=greenfont>��</font>"
	ElseIf Temp = -2147221005 Then
		GBL_CHK_TempStr = "<font color=red class=redfont>���δ��װ</font>"
		CheckObjInstalled = False
	ElseIf Temp = -2147221477 Then
		GBL_CHK_TempStr = "<font color=green class=greenfont>��֧�ִ����</font>"
		CheckObjInstalled = True
	ElseIf Temp = 1 Then
		GBL_CHK_TempStr = "<font color=red>��δ֪�Ĵ����������δ��ȷ��װ</font>"
		CheckObjInstalled = False
	End If
	Err.Clear
	Set TmpObj = Nothing
	Err = 0

End Function
%>