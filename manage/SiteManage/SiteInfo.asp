<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Board_popfun.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<!-- #include file=inc/admanage_Fun.asp -->
<%
DEF_BBS_HomeUrl = "../../"
Const MaxLinkNum = 200
Dim GBL_ID
initDatabase
GBL_CHK_TempStr = ""
GBL_ID = checkSupervisorPass

Manage_sitehead DEF_SiteNameString & " - ����Ա",""
frame_TopInfo
DisplayUserNavigate("��չ����")

Dim GBL_FSOString
GBL_FSOString = DEF_FSOString
If GBL_FSOString = "" Then GBL_FSOString = "Scripting.FileSystemObject"

Dim Fs,FsFlag
FSFlag = 1
Set fs = Server.CreateObject(DEF_FSOString)
If Err Then
     FSFlag = 0
     Err.Clear
End If

Dim MoreSV_LineStr

If GBL_CHK_Flag=1 Then
	Main
Else
DisplayLoginForm
End If
frame_BottomInfo
closeDataBase
Manage_Sitebottom("none")

Sub Main

	Dim Action
	Action = Left(Request.QueryString("action"),10)
	Select Case Action
		Case "MoreSV":
			MoreSV_Main
		Case "Side":
			Side_Main
		Case "admanage":
			admanage_Main
		Case Else:
			SiteInfo
			%>
			<ol class=listli>
			<li><a href=RepairSite.asp>����ͳ���ϴ��ļ����û���������̳��������</a><br>
			<li><a href=DeleteAllTopAnnounce.asp>ȡ��һ���̶ܹ�����</a> <span class=grayfont>�̶ܹ������ܻ������������(�������)�����ô������</span>
			</ol>
			<%
	End Select

End Sub

Sub SiteInfo

	Dim OnlineTime,PageCount,YesterdayAnc,UploadNum,MaxAnnounce,MaxAncTime
	Dim MaxOnline,MaxolTime,DBWrite,DBNum
	Dim Temp1
	Set Rs = LDExeCute(sql_select("Select * from LeadBBS_SiteInfo",1),0)
	If Rs.Eof Then
		OnlineTime = 0
		PageCount = 0
	else
		OnlineTime = cCur(Rs("OnlineTime"))
		PageCount = cCur(Rs("PageCount"))
		YesterdayAnc = cCur(Rs("YesterdayAnc"))
		UploadNum = Rs("UploadNum")
		MaxAnnounce = Rs("MaxAnnounce")
		MaxAncTime = Rs("MaxAncTime")
		MaxOnline = Rs("MaxOnline")
		MaxolTime = Rs("MaxolTime")
		DBWrite = Rs("DBWrite")
		DBNum = Rs("DBNum")
	end If
	Rs.Close
	Set Rs = Nothing

	Dim Rs,Temp
	Response.Write "<ol class=listli><li>ʼ��ʱ��: " & month(application("SiteStartTimeszoieiu")) & "��" & day(application("SiteStartTimeszoieiu")) & "��" & hour(application("SiteStartTimeszoieiu")) & ":" & minute(application("SiteStartTimeszoieiu")) & "</li>" & VbCrLf

	Set Rs = LDExeCute("select count(*) from LeadBBS_User",0)
	If Rs.Eof Then
		Temp = 0
	Else
		Temp = Rs(0)
		if isNull(Temp) or Temp="" Then Temp=0
		Temp = cCur(Temp)
	End If
	Response.Write "<li>��վ�û�: " & Temp & "��</li>" & VbCrLf
	Rs.Close
	Set Rs = Nothing
	Response.Write "<li>��������: " & GetActiveUserNumber & "��<span class=grayfont>(����ͳ��)</span>" & VbCrLf
	Response.Write "<li>��������: " & Application("ActiveUserszoieiu") & "��<span class=grayfont>(Global.ASA)</span></li>" & VbCrLf
	
	Response.Write "<li>����ʱ��: "
	OnlineTime = OnlineTime + application(DEF_MasterCookies & "SiteOlTime")
	Temp = OnlineTime/(24*60*60)
	Response.Write Fix(Temp) & "��"
	OnlineTime=OnlineTime-Fix(Temp)*24*60*60
	Temp = OnlineTime/(60*60)
	Response.Write Fix(Temp) & "ʱ"
	OnlineTime=OnlineTime-Fix(Temp)*60*60
	Temp = OnlineTime/(60)
	Response.Write Fix(Temp) & "��</li>"
	Response.Write "<li>��������: " & PageCount+application(DEF_MasterCookies & "SitePageCount") & "" & "</li>" & VbCrLf
	Response.Write "<li>���շ���: " & YesterdayAnc & "</b>��&nbsp;" & "</li>" & VbCrLf
	Response.Write "<li>�������: " & MaxOnline & "</b>��&nbsp;������" & RestoreTime(MaxolTime) & "</li>" & VbCrLf
	Response.Write "<li>��߷���: " & MaxAnnounce & "</b>��&nbsp;������" & RestoreTime(MaxAncTime) & "</li>" & VbCrLf
	Response.Write "<li>���ݿ���д�����: " & DBWrite & "</b>��</li>" & VbCrLf
	Response.Write "<li>���ݿ����������: " & DBNum & "</b>��</li>" & VbCrLf
	Response.Write "<li>ͳ�ƽ�ֹ: " & year(DEF_Now) & "��" & month(DEF_Now) & "��" & day(DEF_Now) & "</li></ol>" & VbCrLf

End Sub

Function GetActiveUserNumber

	dim Rs
	Set Rs = LDExeCute("select count(*) from LeadBBS_onlineUser",0)
	If Rs.Eof Then
		GetActiveUserNumber = 0
	Else
		GetActiveUserNumber = ccur(Rs(0))
	End If
	Rs.Close
	Set Rs = Nothing

End Function

Sub MoreSV_Main

	%>
	<div class=frametitle>��̳��չ����</div>
	<%
	Dim SV
	SV = Left(Request.QueryString("SV"),10)
	
	Select Case SV
		Case Else:
			MoreSV_BoardCount
	End Select

End Sub

Sub MoreSV_BoardCount

	%>
	<div class=frameline>
	<span class=grayfont><b>1.CNZZ��վͳ����(��ҵý�����ݷ���ר��-WSS����ͳ��)</b></span>
	</div>
	<div class=frameline>��ϵͳ��: <u><a href=http://www.cnzz.com target=_blank>WSSͳ��ϵͳ</a></u> �����ṩ֧��
	</div>
	<%
	Dim User,Pass,Domain,Tmp
	User = 0
	Dim Rs
	Set Rs = LDExeCute(sql_select("Select ID,RID,ValueStr from LeadBBS_Setup where RID=10050",1),0)
	If Not Rs.Eof Then
		Tmp = Split(Trim(Rs(2)&""),"@")
		If Ubound(Tmp,1) + 1 >= 3 Then
			User = Tmp(0)
			Pass = Tmp(1)
			Domain = Tmp(2)
			If isNumeric(User) = 0 Then
				User = 0
				Pass = ""
				Domain = ""
			Else
				User = Fix(cCur(User))
			End If
		End If
	Else
		Tmp = ""
	End If
	Rs.Close
	Set Rs = Nothing
	Dim ID,NewStr
	ID = MoreSV_CheckFileInStr(DEF_BBS_HomeUrl & "inc/incHtm/Bottom_AD.asp","<center><script src=""http://w.cnzz.com/c.php?id=")
	
	NewStr = "<center><script src=""http://w.cnzz.com/c.php?id=" & User & "&amp;l=2"" type=""text/javascript"" charset=""gb2312""></script></center>" & VbCrLf
	If User > 0 and User <> ID Then
		If Request("SV") = "open" Then
			If MoreSV_LineStr <> "" Then
				CALL MoreSV_ReplaceFileStr(DEF_BBS_HomeUrl & "inc/incHtm/Bottom_AD.asp",MoreSV_LineStr,NewStr)
			Else
				CALL MoreSV_ReplaceFileStr(DEF_BBS_HomeUrl & "inc/incHtm/Bottom_AD.asp","",NewStr)
			End If
			ID = User
			%>
			<div class="alert">�ѳɹ�����ҳ�м���ͳ����.</div>
			<%
		End If
	End If
	
	If User = ID and User > 0 Then
		If Request("SV") = "close" Then
			If MoreSV_LineStr <> "" Then
				CALL MoreSV_ReplaceFileStr(DEF_BBS_HomeUrl & "inc/incHtm/Bottom_AD.asp",MoreSV_LineStr,"")
			Else
				CALL MoreSV_ReplaceFileStr(DEF_BBS_HomeUrl & "inc/incHtm/Bottom_AD.asp",NewStr,"")
			End If
			%>
			<div class="alert">�ѳɹ��Ƴ�����ҳ�е�ͳ����.</div>
			<%
			ID = 0
		End If
	End If

	If Request.QueryString("SV") = "counter" and User = 0 Then
		MoreSV_ApplyCounter
	Else
		%>
		<div class=frameline><span class=bluefont>״̬: 
		<%
		If (ID = 0 and User = "") or User = 0 Then
			Response.Write "<span class=redfont>δ��ͨ</span></span> "
			Response.Write "<a href=SiteInfo.asp?action=MoreSV&SV=counter>�������뿪ͨWSSͳ����</a>"
		Else
			Response.Write "<span class=greenfont>�ѿ�ͨ</span></span> "
			
			Response.Write "<p>�˺���Ϣ</p><p>�˺�: p" & User & "@" & Domain
			Response.Write "<br>����: " & Pass & "</p>"
			
			If User > 0 and User <> ID Then
				%>
				<div class="alert"><a href="SiteInfo.asp?action=MoreSV&SV=open">���Ѿ�������ͳ����, ��δ����ҳ�м���ͳ�ƴ���, ��Ҫ����ͳ�ƴ��뿪ʼͳ������.</a></div>
				<%
			Else
				%>
				<div class="alert"><a href="SiteInfo.asp?action=MoreSV&SV=close">���Ѿ���ͨ��ͳ�������Ҽ�����ҳ, ��Ҫɾ����ҳͳ�ƴ�������.</a></div>
				<%
			End If

			Response.Write "<br><b>�鿴ͳ��</b> <a href=http://wss.cnzz.com/user/companion/leadbbs_login.php?site_id=" & User & "&password=" & Pass & " target=_blank><u>�Զ���¼ͳ��ϵͳ</u></a>"
		End If
		%>
		</div>
		<%
	End If

End Sub

Function MoreSV_ApplyCounter

	Dim TmpObj,ResponseTxt,GetUrl,Domain
	
	Domain = Request.ServerVariables("server_name")
	GetUrl = "http://wss.cnzz.com/user/companion/leadbbs.php?domain=" & Domain & "&key=" & MD5(Domain & "J7MdLsaR")
	
	Set TmpObj = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
	TmpObj.setOption 2, 13056 
	TmpObj.open "GET", GetUrl, False, "", "" 
	TmpObj.send()
	ResponseTxt = TmpObj.ResponseText
	Set TmpObj = Nothing
	
	If inStr(ResponseTxt,"@") Then
		MoreSV_ApplyCounter = ResponseTxt
		Dim Rs,ID,Pass
		ID = Left(MoreSV_ApplyCounter,inStr(ResponseTxt,"@")-1)
		Pass = Mid(MoreSV_ApplyCounter,inStr(ResponseTxt,"@")+1)
		Response.Write "<div class=frameline><b>ͳ��ϵͳ�������:</b></div>"
		Response.Write "<div class=frameline> �˺�: p" & ID & "@" & Domain
		Response.Write "</div><div class=frameline>����: " & Pass
		Response.Write "</div><div class=frameline><a href=http://wss.cnzz.com/user/companion/leadbbs_login.php?site_id=" & ID & "&password=" & Pass & " target=_blank><u>�Զ���¼ͳ��ϵͳ</u></a>"
		Response.Write "</div>"
		Set Rs = LDExeCute(sql_select("Select ID,RID,ValueStr from LeadBBS_Setup where RID=10050",1),0)
		If Rs.Eof Then
			CALL LDExeCute("Insert into LeadBBS_Setup(RID,ValueStr,ClassNum) Values(10050,'" & Replace(ResponseTxt & "@" & Domain,"'","''") & "',0)",1)
			Rs.Close
			Set Rs = Nothing
		Else
			CALL LDExeCute("Update LeadBBS_Setup Set ValueStr='" & Replace(ResponseTxt & "@" & Domain,"'","''") & "' where RID=10050",1)
			Rs.Close
			Set Rs = Nothing
		End If
	Else
		MoreSV_ApplyCounter = ""
		Response.Write "<div class=alert>�ӿڴ���,ʧ�ܴ���: " & ResponseTxt & "</div>"
	End If

End Function

Sub MoreSV_ReplaceFileStr(FileName,OldStr,NewStr)

	Dim fs,WriteFile,fileContent
	If FSFlag = 1 Then
		Set fs = Server.CreateObject(GBL_FSOString)
		Set WriteFile = fs.OpenTextFile(Server.MapPath(FileName),1,True)
		If Not WriteFile.AtEndOfStream Then
			fileContent = WriteFile.ReadAll
		End If
		WriteFile.Close
		Set fs = Nothing

		If OldStr = "" Then
			fileContent = fileContent & NewStr
		Else
			fileContent = Replace(fileContent,OldStr,NewStr)
		End If
		Set fs = Server.CreateObject(GBL_FSOString)
		Set WriteFile = fs.CreateTextFile(Server.MapPath(FileName),True)
		WriteFile.Write fileContent
		WriteFile.Close
		Set fs = Nothing
	Else
		fileContent = ADODB_LoadFile(FileName)
		If OldStr = "" Then
			fileContent = fileContent & NewStr
		Else
			fileContent = Replace(fileContent,OldStr,NewStr)
		End If
		ADODB_SaveToFile fileContent,FileName
		Response.Write GBL_CHK_TempStr
	End If

End Sub

Function MoreSV_CheckFileInStr(FileName,Str)

	Dim fs,WriteFile,fileContent,ID,Tmp
	If FSFlag = 1 Then
		Set fs = Server.CreateObject(GBL_FSOString)
		Set WriteFile = fs.OpenTextFile(Server.MapPath(FileName),1,True)
		If Not WriteFile.AtEndOfStream Then
			fileContent = WriteFile.ReadAll
		End If
		WriteFile.Close
		Set fs = Nothing
	Else
		fileContent = ADODB_LoadFile(FileName)
	End If
	
	Tmp = InStr(fileContent,Str)
	If Tmp < 1 Then
		MoreSV_CheckFileInStr = 0
		Exit Function
	End If
	
	ID = Mid(fileContent,Tmp+Len(Str),30)
	ID = Left(ID,inStr(ID,"&") - 1)

	MoreSV_LineStr = Mid(fileContent,Tmp,3000)
	
	Dim BottomStr
	BottomStr = "&amp;l=2"" type=""text/javascript"" charset=""gb2312""></script></center>"
	MoreSV_LineStr = Left(MoreSV_LineStr,inStr(MoreSV_LineStr,BottomStr) + Len(MoreSV_LineStr))
	If isNumeric(ID) = 0 Then
		ID = 0
		MoreSV_LineStr = ""
	End If
	MoreSV_CheckFileInStr = Fix(cCur(ID))

End Function

Sub Side_Main

	If Request.Form("subside") = "1" Then
		Side_UpdateFormData
		Exit Sub
	End If

	Dim Side_Select
	Side_Select = Array("��������","���¾���","����ר��","����ͼƬ","�������")
	
	Dim Side_Data,Dn
	Dim Rs
	Set Rs = LDExeCute("Select ID,RID,ValueStr,ClassNum,saveData from LeadBBS_Setup where RID=01000 order by ClassNum ASC",0)
	If Not Rs.Eof Then
		Side_Data = Rs.GetRows(-1)
		Dn = Ubound(Side_Data,2)
	Else
		Dn = -1
	End If
	Rs.Close
	Set Rs = Nothing
	
	Dim Sn,m
	Dim CheckFlag,Title,RecordCount,OtherInfo,Sort,Tmp,SaveData
	%>
	<script src="<%=DEF_BBS_HomeUrl%>inc/js/jquery.easyui.js"></script>
	<div id="testinfo"></div>
	<div id=test_html style="display:none;"></div>
	<h2>
	��ҳ������������(���϶���������)</h2>
	<form action="SiteInfo.asp?action=Side" method="post" name="LeadBBSFm" onSubmit="return checksubmit(this);"">
	<input type="hidden" value="1" name="subside">
	<div id="home_side_form">
	<%
	For Sn = 0 To Ubound(Side_Select,1)
		CheckFlag = 0
		RecordCount = 10
		If Sn = 3 Then RecordCount = 4 '����ͼƬĬ��¼����
		Title = Side_Select(Sn)
		OtherInfo = ""
		Sort = 0
		For m = 0 to Dn
			If inStr("|" & Side_Data(2,m),"|" & Sn + 1 & "|") Then
				CheckFlag = 1
				Tmp = Split(Trim(Side_Data(2,m)&""),"|")
				If Ubound(Tmp,1) >= 2 Then
					Title = Tmp(1)
					RecordCount = Tmp(2)
				End If
				If Ubound(Tmp,1) >= 3 Then OtherInfo = Tmp(3)
				Sort = Side_Data(3,m)
				Exit For
			End If
		Next
	%>
		<div class="sortitems">
		<input type="checkbox" class=fmchkbox name="Side_Select<%=Sn%>" value="1"<%If CheckFlag = 1 Then
				Response.Write " checked>"
			Else
				Response.Write ">"
			End If%><span class="moveitem"><%=Side_Select(Sn)%></span>
		
		���� <input class='fminpt input_3' maxlength=50 name=Title<%=Sn%> value="<%=htmlencode(Title)%>">
		
		�������� <input name=RecordCount<%=Sn%> value="<%=RecordCount%>" maxLength="2" class="fminpt input_1">
		
		˳�� <input name=Sort<%=Sn%> onchange="$(this).next().html(this.value);sort_start();" value="<%=Sort%>" maxLength="2" class="sortinput fminpt input_1">
		<span style="display:none;" class="sorttxt"><%=Sort%></span>
		
		
		<%If Sn = 2 Then%>
		<br />
		ר����� <input name=OtherInfo<%=Sn%> value="<%=OtherInfo%>" maxLength="12" class="fminpt input_2">
		<a href="../ForumBoard/ForumBoardAssort.asp">��ϸ���������ר������鿴</a>
		
		<%End If%>
		</div>
	<%
	Next
	
	'�����ҳ�����Զ������ 999��ͷΪ�Զ�����
	Sn = Ubound(Side_Select,1)
	Dim MaxSort : MaxSort = 0
	For m = 0 To dn
		CheckFlag = 0
		RecordCount = 10
		If Sn = 3 Then RecordCount = 4 '����ͼƬĬ��¼����
		Title = "��ҳ�����Զ������"
		OtherInfo = ""
		Sort = 0
		If inStr("|" & Side_Data(2,m),"|999|") Then
			Sn = Sn + 1
			CheckFlag = 1
			Tmp = Split(Trim(Side_Data(2,m)&""),"|")
			If Ubound(Tmp,1) >= 2 Then
				Title = Tmp(1)
				RecordCount = Tmp(2)
			End If
			If Ubound(Tmp,1) >= 3 Then OtherInfo = Tmp(3)
			Sort = Side_Data(3,m)
			If MaxSort < Sort Then MaxSort = Sort
			SaveData = Side_Data(4,m) & ""
		%>
			<div class="sortitems">
			<input type="hidden" name="trueID<%=Sn%>" value="<%=Side_Data(0,m)%>">
			<input type="checkbox" class=fmchkbox name="Side_Select<%=Sn%>" value="1"<%If CheckFlag = 1 Then
					Response.Write " checked>"
				Else
					Response.Write ">"
				End If%><span class="moveitem">�Զ������<%=Sn%></span>
			
			���� <input class='fminpt input_3' maxlength=50 name=Title<%=Sn%> value="<%=htmlencode(Title)%>">
			
			<span style="display:none;">�������� <input name=RecordCount<%=Sn%> value="<%=RecordCount%>" maxLength="2" class="fminpt input_1"></span>
			
			˳�� <input name=Sort<%=Sn%> value="<%=Sort%>" onchange="$(this).next().html(this.value);sort_start();" maxLength="2" class="sortinput fminpt input_1">
			<span style="display:none;" class="sorttxt"><%=Sort%></span>
			<br />
			��������룬����ʹ��HTML��JavaScript 
			<textarea cols="80" name="SaveData<%=Sn%>" rows="6" tabindex="51" class="fmtxtra"><%If SaveData <> "" Then Response.Write VbCrLf & htmlEncode(SaveData)%></textarea>
			</div>
		<%
		end if
	Next
	%>
	</div>

	<script>
	var indicator = $('<div class="indicator">>></div>').appendTo('body');
	function initsort()
	{
			initsorted = 1;
			$('.sortitems').draggable({
				revert:true,
				deltaX:0,
				deltaY:0,
				handle:'.moveitem',
			}).droppable({
				onDragOver:function(e,source){
					indicator.css({
						display:'block',
						left:$(this).offset().left-10,
						top:$(this).offset().top+$(this).outerHeight()-5
					});
				},
				onDragLeave:function(e,source){
					indicator.hide();
				},
				onDrop:function(e,source){
					$(source).insertAfter(this);
					indicator.hide();
					sort_byorder();
				}
			});
	}
		$(function(){
			initsort();
		});
	
	
	(function($) {
	$.fn.sorted = function(customOptions) {

		var options = {
			reversed: false,
			by: function(a) { return a.text(); }
		};

		$.extend(options, customOptions);

		$data = $(this);
		arr = $data.get();
		arr.sort(function(a, b) {
			var valA = options.by($(a));
			var valB = options.by($(b));
			if (options.reversed) {
				return (valA < valB) ? 1 : (valA > valB) ? -1 : 0;				
			} else {		
				return (valA < valB) ? -1 : (valA > valB) ? 1 : 0;	
			}
		});
		
		var upDom = $(this).parent();
		$(upDom).empty();
		for(var n=0;n<arr.length;n++)
		{		
		$(arr[n]).find("input.sortinput").val(n);
		$(arr[n]).find("span.sorttxt").html(n);
		$(upDom).append($(arr[n]));}
		initsort();
		return $(arr);
	};
})(jQuery);

function sort_start()
{
	var arr=$("#home_side_form .sortitems").sorted(
		{
			by: function(v) {
				return parseInt(v.find('span.sorttxt').html());
			}
		}
	);
}

function sort_byorder()
{
	var arr=$("#home_side_form .sortitems");
	for(var n=0;n<arr.length;n++)
	{		
	$(arr[n]).find("input.sortinput").val(n);
	$(arr[n]).find("span.sorttxt").html(n);}
}

function sort_start()
{
	var arr=$("#home_side_form .sortitems").sorted(
		{
			by: function(v) {
				return parseInt(v.find('span.sorttxt').html());
			}
		}
	);
}
sort_start();

	function checksubmit(f)
	{
		var textarea = $('textarea').length;
		for(var n=0;n<textarea;n++)
		{
			if($('textarea').eq(n))
			{
				$("#test_html").html($('textarea').eq(n).val());
				$('textarea').eq(n).val($("#test_html").html());
				if($('textarea').eq(n).val().length>10240)
				{alert("�����Զ������"+(n+5)+" ���ȳ�����10240.");return false;}
			}
		}
		return true;
	}
	</script>
<script language=javascript>
var maxNumber=<%=MaxLinkNum%>;
var Number=<%=Sn%>;
var MaxSort=<%=MaxSort%>;

function additem()
{
	Number+=1;
	if(Number>maxNumber)
	{
		alert("�Ѿ��ﵽ�����Ŀ������������!");
	}
	else
	{
		
		var tmp="<table border=0 cellpadding=0 class=blanktable><tr><td><input type=hidden name=trueID" + Number + " value=999999>";
		tmp+="<input type=checkbox class=fmchkbox name=Side_Select" + Number + " value=1 checked><span class='moveitem'>�Զ������" + Number + "</span></td><td>";
		tmp+="���� <input class='fminpt input_3' maxlength=50 name=Title" + Number + " value=''></td><td>";
		tmp+="<span style='display:none;'>�������� <input name=RecordCount" + Number + " value='' maxLength=2 class='fminpt input_1'></span></td><td>";
		tmp+="˳�� <input name=Sort" + Number + " onchange='$(this).next().html(this.value);sort_start();' value='"+(MaxSort)+"' maxLength=2 class='sortinput fminpt input_1'><span style='display:none;' class='sorttxt'>"+(MaxSort)+"</span></td></tr><tr><td> </td><td colspan=3>";
		tmp+="��������룬����ʹ��HTML��JavaScript <textarea cols=80 name=SaveData" + Number + " rows=6 tabindex=51 class=fmtxtra></textarea></td></tr></table>";
		$id('home_side_form').innerHTML+=tmp;
		//this.scroll(0, 65000);
	}
}
</script>
<a href=javascript:; onclick="additem();" class=manage_submit>�������Զ������(���Բ������������HTML����)</a>
	<%
	

	Side_Select = Array("�Ӱ��","�������","��龫��")
	
	Set Rs = LDExeCute("Select * from LeadBBS_Setup where RID=01003 order by ClassNum",0)
	If Not Rs.Eof Then
		Side_Data = Rs.GetRows(-1)
		Dn = Ubound(Side_Data,2)
	Else
		Dn = -1
	End If
	Rs.Close
	Set Rs = Nothing
	
	%>
	<br />
	<hr class=splitline>
	<br />
	<p>
	<b>���������������</b>
	</p>
	<%
	For Sn = 0 To Ubound(Side_Select,1)
		CheckFlag = 0
		RecordCount = 10
		If Sn = 4 Then RecordCount = 5 '����ͼƬĬ��¼����
		Title = Side_Select(Sn)
		OtherInfo = ""
		Sort = 0
		For m = 0 to Dn
			If inStr("|" & Side_Data(2,m),"|" & Sn + 1 & "|") Then
				CheckFlag = 1
				Tmp = Split(Trim(Side_Data(2,m)&""),"|")
				If Ubound(Tmp,1) >= 2 Then
					Title = Tmp(1)
					RecordCount = Tmp(2)
				End If
				If Ubound(Tmp,1) >= 3 Then OtherInfo = Tmp(3)
				Sort = Side_Data(3,m)
				Exit For
			End If
		Next
	%>
		<table border=0 cellpadding="0" class="blanktable">
		<tr>
		<td>
		<input type="checkbox" class=fmchkbox name="board_Side_Select<%=Sn%>" value="1"<%If CheckFlag = 1 Then
				Response.Write " checked>"
			Else
				Response.Write ">"
			End If%><%=Side_Select(Sn)%>
		</td><td>
		���� <input class='fminpt input_3' maxlength=50 name=board_Title<%=Sn%> value="<%=htmlencode(Title)%>">
		</td><td>
		<%If Sn = 0 Then%>
		<input name=board_RecordCount<%=Sn%> value="0" maxLength="2" class="fminpt input_1" type="hidden">
		<%Else%>
		�������� <input name=board_RecordCount<%=Sn%> value="<%=RecordCount%>" maxLength="2" class="fminpt input_1">
		<%End If%>
		</td><td>
		˳�� <input name=board_Sort<%=Sn%> value="<%=Sort%>" maxLength="2" class="fminpt input_1">
		</td></tr>
		
		<%If Sn = 3 Then%>
		<tr><td> </td><td colspan="3">
		ר����� <input name=board_OtherInfo<%=Sn%> value="<%=OtherInfo%>" maxLength="12" class="fminpt input_2">
		<a href="../ForumBoard/ForumBoardAssort.asp">��ϸ���������ר������鿴</a>
		</td></tr>
		<%End If%>
		</table>
	<%
	Next
	%>
	
	<p>
	<input name=submit type=submit value="�������" class="fmbtn">
	</p>
	</form>
		<div class=frametitle>ע��:</div>
		<ol class=listli>
		<li>��Ҫ����ҳ��ʾ��Ӧ��Ϣ�빴ѡǰ�渴ѡ��</li>
		<li>����: ָ���ǵ�����Ŀ��TITLE</li>
		<li>��������: ������ʾ��Ӧ���ݵļ�¼����</li>
		<li>˳��: �ڲ�������ʾ˳��, ��������. ����ԽС������ʾ��Խ�����λ��</li>
		<li>ר�����: �����Ҫ����ר��,����Ҫ��д��Ӧ�İ���ר�����,��������ר����(������������)</li>
		<li>�Ӱ��: �����ò�����ʾ�Ӱ��,�򲻻����ظ�������������ʾ.</li>
		</ol>
	<%

End Sub

Sub Side_UpdateFormData

	Dim Sn,m,Rs
	Dim CheckFlag,Title,RecordCount,OtherInfo,Sort,Tmp,SaveData,trueID
	For Sn = 0 to MaxLinkNum
		CheckFlag = Request.Form("Side_Select" & Sn)
		Title = Request.Form("Title" & Sn)
		RecordCount = toNum(Request.Form("RecordCount" & Sn),0)
		Sort = Request.Form("Sort" & Sn)
		OtherInfo = Request.Form("OtherInfo" & Sn)
		SaveData = Request.Form("SaveData" & Sn)
		trueID = toNum(Request.Form("trueID" & Sn),0)
		If CheckFlag = "1" Then
			Title = Left(Replace(Title,"|",""),50)
			SaveData = Replace(Replace(Left(Replace(SaveData,"|",""),10240),"<" & "%","&lt;%"),"%" & ">","%&gt;")
			If Title = "" Then Title = "�ޱ���"
			If isNumeric(RecordCount) = 0 Then RecordCount = 10
			RecordCount = cCur(Fix(RecordCount))
			If RecordCount < 1 or RecordCount > 99 Then RecordCount = 10
			
			If isNumeric(Sort) = 0 Then Sort = 0
			Sort = cCur(Fix(Sort))
			
			OtherInfo = Left(Replace(OtherInfo,"|",""),50)
			If Sn = 2 Then
				If isNumeric(OtherInfo) = 0 Then OtherInfo = 54
				OtherInfo = cCur(Fix(OtherInfo))
			End If
			If Sn > 4 Then
				Tmp = "999|" & Title & "|" & RecordCount
			Else
				Tmp = Sn + 1 & "|" & Title & "|" & RecordCount
			End If
			If OtherInfo <> "" Then Tmp = Tmp & "|" & OtherInfo
			
			If trueID = 0 Then
				Set Rs = LDExeCute("Select * from LeadBBS_Setup where RID=01000 and ValueStr like '" & Sn + 1 & "|%'",0)
				If Not Rs.Eof Then
					CALL LDExeCute("Update LeadBBS_Setup Set ValueStr='" & Replace(Tmp,"'","''") & "',ClassNum=" & Sort & " where RID=01000 and ValueStr like '" & Sn + 1 & "|%'",1)
				Else
					CALL LDExeCute("insert into LeadBBS_Setup(RID,ValueStr,ClassNum,SaveData) Values(01000,'" & Replace(Tmp,"'","''") & "'," & Sort & ",'')",1)
				End If
			Else
				Set Rs = LDExeCute("Select * from LeadBBS_Setup where RID=01000 and id=" & trueID,0)
				If Not Rs.Eof Then
					CALL LDExeCute("Update LeadBBS_Setup Set ValueStr='" & Replace(Tmp,"'","''") & "',ClassNum=" & Sort & ",SaveData='" & Replace(SaveData,"'","''") & "' where RID=01000 and id=" & trueID,1)
				Else
					CALL LDExeCute("insert into LeadBBS_Setup(RID,ValueStr,ClassNum,SaveData) Values(01000,'" & Replace(Tmp,"'","''") & "'," & Sort & ",'" & Replace(SaveData,"'","''") & "')",1)
				End If
			End If
			Rs.Close
			Set Rs = Nothing
		Else
			If trueID = 0 Then
				if Sn <= 4 then CALL LDExeCute("delete from LeadBBS_Setup where RID=01000 and ValueStr like '" & Sn + 1 & "|%'",1)
			Else
				CALL LDExeCute("delete from LeadBBS_Setup where RID=01000 and id=" & trueID,1)
			End If
		End If
	Next
	Side_UpdateFileData
	
	
	For Sn = 0 to 4
		CheckFlag = Request.Form("board_Side_Select" & Sn)
		Title = Request.Form("board_Title" & Sn)
		RecordCount = Request.Form("board_RecordCount" & Sn)
		Sort = Request.Form("board_Sort" & Sn)
		OtherInfo = Request.Form("board_OtherInfo" & Sn)
		If CheckFlag = "1" Then
			Title = Left(Replace(Title,"|",""),50)
			If Title = "" Then Title = "�ޱ���"
			If isNumeric(RecordCount) = 0 Then RecordCount = 10
			RecordCount = cCur(Fix(RecordCount))
			If RecordCount < 1 or RecordCount > 99 Then RecordCount = 10
			
			If isNumeric(Sort) = 0 Then Sort = 0
			Sort = cCur(Fix(Sort))
			
			OtherInfo = Left(Replace(OtherInfo,"|",""),50)
			If Sn = 3 Then
				If isNumeric(OtherInfo) = 0 Then OtherInfo = 54
				OtherInfo = cCur(Fix(OtherInfo))
			End If
			Tmp = Sn + 1 & "|" & Title & "|" & RecordCount
			If OtherInfo <> "" Then Tmp = Tmp & "|" & OtherInfo
			
			Set Rs = LDExeCute("Select * from LeadBBS_Setup where RID=01003 and ValueStr like '" & Sn + 1 & "|%'",0)
			If Not Rs.Eof Then
				CALL LDExeCute("Update LeadBBS_Setup Set ValueStr='" & Replace(Tmp,"'","''") & "',ClassNum=" & Sort & " where RID=01003 and ValueStr like '" & Sn + 1 & "|%'",1)
			Else
				CALL LDExeCute("insert into LeadBBS_Setup(RID,ValueStr,ClassNum,SaveData) Values(01003,'" & Replace(Tmp,"'","''") & "'," & Sort & ",'')",1)
			End If
			Rs.Close
			Set Rs = Nothing
		Else
			CALL LDExeCute("delete from LeadBBS_Setup where RID=01003 and ValueStr like '" & Sn + 1 & "|%'",1)
		End If
	Next
	
	Board_Side_UpdateFileData

End Sub

Sub Side_UpdateFileData

	
	Dim Side_Data,Dn
	Dim Rs
	Set Rs = LDExeCute("Select * from LeadBBS_Setup where RID=01000 order by ClassNum",0)
	If Not Rs.Eof Then
		Side_Data = Rs.GetRows(-1)
		Dn = Ubound(Side_Data,2)
	Else
		Dn = -1
	End If
	Rs.Close
	Set Rs = Nothing
	
	Dim m
	Dim Title,RecordCount,OtherInfo,Tmp,SideType,SaveData
	
	Dim Str
	Str = "<" & "%" & VbCrLf
	For m = 0 to Dn
		Tmp = Split(Side_Data(2,m),"|")
		If Ubound(Tmp,1) >= 2 Then
			SideType = Tmp(0)
			Title = Tmp(1)
			RecordCount = Tmp(2)
		End If
		SaveData = Replace(Replace(Replace(Replace(Replace(Replace(Side_Data(4,m) & "",VbCrLf,""),chr(0),""),chr(13),""),"""",""""""),"<script","<sc"" & ""ript"),"/script","/sc"" & ""ript")
		If Ubound(Tmp,1) >= 3 Then OtherInfo = Tmp(3)
		Select Case cCur(SideType)
			Case 1:			
				Str = Str & "Str = Str & ""		<div class=""""content_side_box"""">"" & VbCrLf &_" & VbCrLf &_
				"""			<div class=""""title""""><b>" & htmlencode(Title) & "</b></div>"" & VbCrLf &_" & VbCrLf &_
				"""			"" & Topic_AnnounceList(0," & RecordCount & ",0,""yes"",""0"",""0"","""") & VbCrLf &_" & VbCrLf &_
				"""		</div>"" & VbCrLf" & VbCrLf
			Case 2:			
				Str = Str & "Str = Str & ""		<div class=""""content_side_box"""">"" & VbCrLf &_" & VbCrLf &_
				"""			<div class=""""title""""><b>" & htmlencode(Title) & "</b></div>"" & VbCrLf &_" & VbCrLf &_
				"""			"" & Topic_AnnounceList(0," & RecordCount & ",0,""yes"",""1"",""0"","""") & VbCrLf &_" & VbCrLf &_
				"""		</div>"" & VbCrLf" & VbCrLf
			Case 3:			
				Str = Str & "Str = Str & ""		<div class=""""content_side_box"""">"" & VbCrLf &_" & VbCrLf &_
				"""			<div class=""""title""""><b>" & htmlencode(Title) & "</b></div>"" & VbCrLf &_" & VbCrLf &_
				"""			"" & Topic_AnnounceList(0," & RecordCount & "," & OtherInfo & ",""yes"",""0"",""0"","""") & VbCrLf &_" & VbCrLf &_
				"""		</div>"" & VbCrLf" & VbCrLf
			Case 4:			
				Str = Str & "Str = Str & ""		<div class=""""content_side_box"""">"" & VbCrLf &_" & VbCrLf &_
				"""			<div class=""""title""""><b>" & htmlencode(Title) & "</b></div>"" & VbCrLf &_" & VbCrLf &_
				"""			"" & Topic_PicInfo(140,105," & RecordCount & ") & VbCrLf &_" & VbCrLf &_
				"""		</div>"" & VbCrLf" & VbCrLf
			Case 5:
				Str = Str & "Str = Str & ""		<div class=""""content_side_box"""">"" & VbCrLf &_" & VbCrLf &_
				"""			<div class=""""title""""><b>" & htmlencode(Title) & "</b></div>"" & VbCrLf &_" & VbCrLf &_
				"""			"" & Topic_AnnounceList(0," & RecordCount & ",0,""yes"",""2"",""0"","""") & VbCrLf &_" & VbCrLf &_
				"""		</div>"" & VbCrLf" & VbCrLf
			Case Else				
				Str = Str & "Str = Str & ""		<div class=""""content_side_box"""">"" & VbCrLf &_" & VbCrLf &_
				"""			<div class=""""content"""">" & SaveData & "</b></div>"" & VbCrLf &_" & VbCrLf &_
				"""		</div>"" & VbCrLf" & VbCrLf
		End Select
	Next
	Str = Str & "%" & ">"
	CALL ADODB_SaveToFile(Str,DEF_BBS_HomeUrl & "inc/IncHtm/Boards_Side_Setup.asp")
	Response.Write "<p>��ҳ�������������. ������ò���һ��������Ч,�´θ�ʱˢ�½����Զ���ɸ���.</p><p><a href=""SiteInfo.asp?action=Side"">��˷�������</a></p>"

End Sub

Sub Board_Side_UpdateFileData

	
	Dim Side_Data,Dn
	Dim Rs
	Set Rs = LDExeCute("Select * from LeadBBS_Setup where RID=01003 order by ClassNum",0)
	If Not Rs.Eof Then
		Side_Data = Rs.GetRows(-1)
		Dn = Ubound(Side_Data,2)
	Else
		Dn = -1
	End If
	Rs.Close
	Set Rs = Nothing
	
	Dim m
	Dim Title,RecordCount,OtherInfo,Tmp,SideType
	
	Dim Str,SubBoard_Flag
	SubBoard_Flag = 0
	Str = "<" & "%" & VbCrLf
	Str = Str & "Function SideBoard_GetContent()" & VbCrLf
	Str = Str & "Dim Str,Tmp" & VbCrLf
	For m = 0 to Dn
		Tmp = Split(Side_Data(2,m),"|")
		If Ubound(Tmp,1) >= 2 Then
			SideType = Tmp(0)
			Title = Tmp(1)
			RecordCount = Tmp(2)
		End If
		If Ubound(Tmp,1) >= 3 Then OtherInfo = Tmp(3)
		Select Case cCur(SideType)
			Case 1:	
				Str = Str & "Tmp = Topic_AnnounceList(GBL_Board_ID,10,0,""yes"",""3"",""0"","""")" & VbCrLf
				Str = Str & "If Tmp <> """" Then Str = Str & ""		<div class=""""content_side_box"""">"" & VbCrLf &_" & VbCrLf &_
				"""			<div class=""""title""""><b>" & htmlencode(Title) & "</b></div>"" & VbCrLf &_" & VbCrLf &_
				"""			"" & Topic_AnnounceList(GBL_Board_ID," & RecordCount & ",0,""no"",""3"",""0"","""") & VbCrLf &_" & VbCrLf &_
				"""		</div>"" & VbCrLf" & VbCrLf
				SubBoard_Flag  = 1
			Case 2:	
				Str = Str & "Str = Str & ""		<div class=""""content_side_box"""">"" & VbCrLf &_" & VbCrLf &_
				"""			<div class=""""title""""><b>" & htmlencode(Title) & "</b></div>"" & VbCrLf &_" & VbCrLf &_
				"""			"" & Topic_AnnounceList(GBL_Board_ID," & RecordCount & ",0,""yes"",""0"",""0"","""") & VbCrLf &_" & VbCrLf &_
				"""		</div>"" & VbCrLf" & VbCrLf
			Case 3:	
				Str = Str & "Tmp = Topic_AnnounceList(GBL_Board_ID," & RecordCount & ",0,""yes"",""1"",""0"","""")" & VbCrLf
				Str = Str & "If Tmp <> """" Then Str = Str & ""		<div class=""""content_side_box"""">"" & VbCrLf &_" & VbCrLf &_
				"""			<div class=""""title""""><b>" & htmlencode(Title) & "</b></div>"" & VbCrLf &_" & VbCrLf &_
				"""			"" & Tmp & VbCrLf &_" & VbCrLf &_
				"""		</div>"" & VbCrLf" & VbCrLf
			Case 4:	
				Str = Str & "Str = Str & ""		<div class=""""content_side_box"""">"" & VbCrLf &_" & VbCrLf &_
				"""			<div class=""""title""""><b>" & htmlencode(Title) & "</b></div>"" & VbCrLf &_" & VbCrLf &_
				"""			"" & Topic_AnnounceList(GBL_Board_ID," & RecordCount & "," & OtherInfo & ",""yes"",""0"",""0"","""") & VbCrLf &_" & VbCrLf &_
				"""		</div>"" & VbCrLf" & VbCrLf
		End Select
	Next
	Str = Str & "SideBoard_GetContent = Str" & VbCrLf
	Str = Str & "End Function" & VbCrLf
	Str = Str & "Const GBL_B_SubBoard_Flag = " & SubBoard_Flag & VbCrLf
	Str = Str & "%" & ">"
	CALL ADODB_SaveToFile(Str,DEF_BBS_HomeUrl & "inc/IncHtm/Boards_Side_Setup2.asp")
	Response.Write "<p>����������������. </p><p><a href=""SiteInfo.asp?action=Side"">��˷�������</a></p>"

End Sub
%>