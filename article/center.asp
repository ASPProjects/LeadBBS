<!-- #include file=../inc/BBSsetup.asp -->
<!-- #include file=../inc/User_Setup.asp -->
<!-- #include file=../inc/Board_popfun.asp -->
<!-- #include file=../inc/Limit_fun.asp -->
<!-- #include file=inc/UserTopic.asp -->
<!-- #include file=../inc/Constellation2.asp -->
<!-- #include file=../a/inc/upload1_fun.asp -->
<%DEF_BBS_HomeUrl = "../"%>
<!-- #include file=../inc/Upload_Fun.asp -->
<!-- #include file=inc/popfun.asp -->
<!-- #include file=inc/cms_fun.asp -->
<!-- #include file=inc/center_editfile.asp -->
<!-- #include file=inc/center_setchannel.asp -->
<!-- #include file=../a/inc/Editor.asp -->
<!-- #include file=../a/inc/Editor_Fun.asp -->
<%
Dim Form_Submitflag,Form_Action,Form_ActionStr,GBL_AjaxFlag,Form_UpClass
Dim LMT_EnableUpload
LMT_EnableUpload = 1
Dim Form_EditAnnounceID
Form_EditAnnounceID = 0
UploadTable = "article_upload"
PhotoDirectory = DEF_BBS_HomeUrl & DEF_CMS_UploadPhotoUrl
UploadPhotoUrl = DEF_BBS_HomeUrl & DEF_CMS_UploadPhotoUrl
UploadOneDayMaxNum = DEF_CMSUploadOneDayMaxNum
upload_NoteLength = 255

Main

Sub Page_Expires

	Response.Expires = 0
	Response.ExpiresAbsolute = DEF_Now - 1
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private" 
	Response.CacheControl = "no-cache"

End Sub

Sub Main

	Page_Expires
	initDatabase
	User_GetStartValue

	If GBL_AjaxFlag = 0 Then
		article_center_Head(Form_ActionStr)
		%><div class="body_area_out">
		<%cms_manage_Navigate("<span class=navigate_string_step>" & Form_ActionStr & "</span>")
	End If
	UpdateOnlineUserAtInfo GBL_board_ID,Form_ActionStr
	GBL_CHK_TempStr=""
	If GBL_UserID = 0 Then GBL_CHK_TempStr = GBL_CHK_TempStr & "δ��¼���������.<br>" & VbCrLf

	If GBL_AjaxFlag = 0 Then UserTopicTopInfo("user")

	If GBL_CHK_Flag=1 Then
		If GBL_CHK_TempStr = "" Then
			Main_Action
		Else
			cms_DisplayLoginForm(GBL_CHK_TempStr)
		End If
	Else
		If Form_Submitflag = "" Then
			cms_DisplayLoginForm("���ȵ�¼")
		Else
			cms_DisplayLoginForm("<span class=redfont>" & GBL_CHK_TempStr & "</span>")
		End If
	End If

	closeDataBase
	If GBL_AjaxFlag = 0 Then cms_UserTopicBottomInfo
	If GBL_AjaxFlag = 0 Then cms_center_Bottom
	If Form_UpFlag = 1 Then Set Form_UpClass = Nothing

End Sub

Sub User_GetStartValue

	If Request.QueryString("dontRequestFormFlag") = "" Then
		Form_UpFlag = 0
	Else
		Form_UpFlag = 1
		init_Upload = 1
		Server.ScriptTimeOut=3000
		set Form_UpClass = new upload_Class
		Form_UpClass.ProgressID = Request.QueryString("Upload_ID")
		Form_UpClass.GetUpFile
	End If

	Form_Submitflag = Request.QueryString("submitflag")
	If Form_Submitflag = "" Then Form_Submitflag = GetFormData("submitflag")
	Form_Action = GetFormData("action")
	GBL_AjaxFlag = GetFormData("ajaxflag") 
	If GBL_AjaxFlag <> "1" Then 
		GBL_AjaxFlag = 0
	Else
		GBL_AjaxFlag = 1
	End If
	Select Case Form_Action
		case "newsclass":
			Form_ActionStr = "������·���(����Ա)"
			if GetFormData("form_modifyid") <> "" and GetFormData("form_modifyid") <> "0" Then Form_ActionStr = "�༭���·���(����Ա)"
		case "newsarticle":
			Form_ActionStr = "�����������"
			if GetFormData("form_modifyid") <> "" and GetFormData("form_modifyid") <> "0" Then Form_ActionStr = "�༭��������"
		case "newsmanage":
			Form_ActionStr = "��������"
		case "editfile":
			Form_ActionStr = "�༭������Ϣ"
		case "setchannel":
			Form_ActionStr = "������ҳ��Ŀ����"
		case "updatecache":
			Form_ActionStr = "���»���"
		Case Else
			Form_Action = "newsmanage"
			Form_ActionStr = "��������"
	End Select			

End Sub


Sub Main_Action

	If Check_jdsupervisor = 0 and (CheckUserAnnounceLimit = 0 or GBL_UserID < 1) Then
		Response.Write "<span class=cms_error>����Ȩ���д˲���,���ܴ��û�δ��֤,���ѱ��������Ȩ��.</span>"
		Exit Sub
	End if
	Select Case Form_Action
		case "newsclass":
			center_newsclass
		case "newsarticle":
			center_newsarticle
		case "newsmanage":
			center_newsmanage
		case "editfile":
			center_editfile
		case "setchannel":
			center_setchannel
		case "updatecache":
			center_updatecache
	End Select

End Sub

sub center_updatecache

	Response.Write "<div class=""cms_ok"">��ʼ���»��森����</div>"
	Response.Write "<div class=""cms_ok"">��ʼ��ȡ�������ݲ�չʾ������</div>"
	Response.Write "<div style=""zoom:0.8;max-height:600px;overflow:auto;"">"
	dim cmscacheClass
	set cmscacheClass = new cms_cache_Class
	cmscacheClass.updatecache
	set cmscacheClass = nothing
	response.Write "</div>"
	Response.Write "<div class=""clear""></div><div class=""cms_ok"" style=""width:100%;"">���������ɣ�</div>"

End sub
%>