<!-- #include file=../../inc/BBSSetup.asp -->
<!-- #include file=../../inc/Board_Popfun.asp -->
<!-- #include file=../../inc/Upload_Setup.asp -->
<%
'----------------------------------------------------------------------------------------
'----- LeadBBS�������ӵ���(2003-7-2)
'----- ���þ���: <script language=javascript src=http://��װĿ¼/other/Article/topic.asp?B=1&Number=20&StrLen=12&NewWindow=yes></script>
'----- ��ʾ��ʾ����ɣĺ�Ϊ1�İ��������20�����⣬������������12���ֽڵ���12�ֽ���ʾ(һ������Ϊ�����ֽ�)
'----- ������BoardID ����ɣĺţ������ָ�������ʾ������̳����10������
'----- Number����ʾ��¼�����������������100������ָ��Ĭ�ϵ���10������
'----- StrLen���е������������̫����Ϊ��������ʾЧ����������ʾ������ֽڳ���
'----- NewWindow�����´����д�������Ϊyes��������ʾ��
'----- PithFlag���Ƿ���þ�����1-��ʾ�����þ�����(���Լ���BoardID���Ƶ��ð���)
'----- newanc=1���Ƿ�ֻ�������»ظ�(ʹ�ô˲�������������ȫ��ʧЧ)
'----------------------------------------------------------------------------------------
'----- ����ĳר������������
'----- <script language=javascript src=http://��װĿ¼/other/Article/topic.asp?GoodAssort=10&Number=20&StrLen=12&NewWindow=yes></script>
'----- ������GoodAssort��ָר�������
'----------------------------------------------------------------------------------------


'----------------------------------------------------------------------------------------
'----- LeadBBS��̳��Ϣ���ó���(2003-09-30 by LeadBBS) for LeadBBS v2.88 v2.77 v3.14
'----- ���þ���: <script language=javascript src=http://��װĿ¼/other/Article/Topic.asp?img=/images/del.gif&info=1></script>
'----- ����img��ָ����ʾ��ǰ�������ͼƬ��ַ����ָ������ʾͼƬ��Ҫ���С16x16
'----- ����exectime���Ƿ���ʾִ��ʱ�䣬no:����ʾ��������ʾ��ʾ
'----------------------------------------------------------------------------------------

'--------------
'LeadBBS ����ͼƬ����
'ʹ�÷���: <script language=javascript src=http://��װĿ¼/other/Article/Topic.asp?info=pic&num=5></script>
'����num: ��ʾ���µ�ͼƬ����������Ϊ1-50��Ĭ��Ϊ6
'--------------



Dim HomeUrl
HomeUrl = "http://"&Request.ServerVariables("server_name")
If Request.ServerVariables("SERVER_PORT") <> "80" Then HomeUrl = HomeUrl & ":" & Request.ServerVariables("SERVER_PORT")
HomeUrl = Lcase(HomeUrl & Request.Servervariables("SCRIPT_NAME"))
HomeUrl = Replace(HomeUrl,"other/article/topic.asp","")

DEF_BBS_HomeUrl = "../../"
Const LMT_MaxDisplayAnnounce = 100
OpenDatabase
If Request.QueryString("info") = "pic" Then
	Topic_PicInfo
ElseIf Request.QueryString("info") <> "" Then
	Topic_BBSInfo
Else
	Topic_AnnounceList
End If
closeDataBase

Sub Topic_AnnounceList

	Dim BoardID,ListNum,StrLen,GoodAssort
	BoardID = Left(Request.QueryString("BoardID"),14)
	ListNum = Left(Request.QueryString("Number"),14)
	StrLen = Left(Request.QueryString("StrLen"),14)
	GoodAssort = Left(Request.QueryString("GoodAssort"),14)
	If isNumeric(StrLen) = False or StrLen = "" Then StrLen = 20
	StrLen = Fix(cCur(StrLen))
	If isNumeric(ListNum) = False or ListNum = "" Then ListNum = 10
	If isNumeric(BoardID) = False Then BoardID = 0
	BoardID = Fix(cCur(BoardID))
	If isNumeric(GoodAssort) = False Then GoodAssort = 0
	GoodAssort = Fix(cCur(GoodAssort))

	ListNum = Fix(cCur(ListNum))
	If ListNum < 1 or ListNum > 100 Then ListNum = 10
	If StrLen < 1 or StrLen > 255 Then StrLen = 255
	If GoodAssort < 1 Then GoodAssort = 0
	
	Dim NewWindow
	NewWindow = Request.QueryString("NewWindow")
	If NewWindow <> "yes" Then
		NewWindow = ""
	Else
		NewWindow = " target=_blank"
	End If

	Dim Rs,SQL
	If DEF_UsedDataBase = 1 Then
		SQL = "select T1.ID,T1.Title,T1.TitleStyle,T1.BoardID,T2.ForumPass,T2.BoardLimit,T2.OtherLimit,T2.HiddenFlag from LeadBBS_Topic as T1 left join LeadBBS_Boards as T2 on T1.BoardID=T2.BoardID"
		If GoodAssort > 0 Then
			SQL = SQL & " where T1.GoodAssort=" & GoodAssort & " Order by ID DESC"
		Else
			If Request.QueryString("PithFlag") = "1" Then
				If BoardID = 0 Then
					SQL = SQL & " Where T1.GoodFlag=1 Order by T1.ID DESC"
				Else
					SQL = SQL & " Where T1.GoodFlag=1 and T1.BoardID=" & BoardID & " Order by T1.ID DESC"
				End If
			Else
				If BoardID = 0 Then
					SQL = SQL & " Order by T1.ID DESC"
				Else
					SQL = SQL & " Where T1.BoardID=" & BoardID & " Order by T1.RootID DESC"
				End If
			End If
		End If
		sql = sql_select(sql,ListNum)
	Else
	If Request.QueryString("newanc") = "1" Then
		SQL = sql_select("select T1.ID,T1.Title,T1.TitleStyle,T1.BoardID,T2.ForumPass,T2.BoardLimit,T2.OtherLimit,T2.HiddenFlag from LeadBBS_Announce as T1 left join LeadBBS_Boards as T2 on T1.BoardID=T2.BoardID Order by T1.ID DESC",ListNum)
	Else
		SQL = "select T1.ID,T1.Title,T1.TitleStyle,T1.BoardID,T2.ForumPass,T2.BoardLimit,T2.OtherLimit,T2.HiddenFlag from LeadBBS_Announce as T1 left join LeadBBS_Boards as T2 on T1.BoardID=T2.BoardID Where"
		If GoodAssort > 0 Then
			SQL = SQL & " T1.GoodAssort=" & GoodAssort & " Order by ID DESC"
		Else
			If Request.QueryString("PithFlag") = "1" Then
				If BoardID = 0 Then
					SQL = SQL & " T1.GoodFlag=1 Order by T1.ID DESC"
				Else
					SQL = SQL & " T1.GoodFlag=1 and T1.BoardID=" & BoardID & " Order by T1.ID DESC"
				End If
			Else
				If BoardID = 0 Then
					SQL = SQL & " T1.ParentID=0 Order by T1.RootIDBak DESC"
				Else
					SQL = SQL & " T1.ParentID=0 and T1.BoardID=" & BoardID & " Order by T1.RootIDBak DESC"
				End If
			End If
		End If
		sql = sql_select(sql,ListNum)
	End If
	End If
	Set Rs = LDExeCute(SQL,0)
	Dim Num
	Dim GetData
	If Not rs.Eof Then
		GetData = Rs.GetRows(-1)
		Num = Ubound(GetData,2)
	Else
		Num = -1
	End If
	Rs.close
	Set Rs = Nothing
	Response.Write "document.write("""
	If Request.QueryString("exectime") = "1" Then Response.Write "ִ��ʱ��" & FormatNumber(cCur(Timer - DEF_PageExeTime1),4,True)*1000 & "ms ���ݿ�" & GBL_DBNum & "��<br>"


	Dim img
	img = Replace(Replace(Left(Request.QueryString("img"),100),"\",""),"""","")
	If img <> "" Then img = "<img src=" & Replace(img,"\","\\") & " align=middle> "
	If img = "" Then img = "<li>"

	For SQL = 0 to Num
		If GetData(2,SQL) = 1 Then GetData(1,SQL) = KillHTMLLabel(GetData(1,SQL))
		If Len(GetData(1,SQL)) > StrLen/2 Then
			If StrLength(GetData(1,SQL)) > StrLen Then
				GetData(1,SQL) = LeftTrue(GetData(1,SQL),StrLen - 3) & "..."
			End If
		End If
		GetData(1,SQL) = Replace(htmlencode(GetData(1,SQL)),"\","\\")
		'GetData(1,SQL) = DisplayAnnounceTitle(GetData(1,SQL),GetData(2,SQL))
		If GBL_CheckLimitTitle(GetData(4,SQL),GetData(5,SQL),GetData(6,SQL),GetData(7,SQL)) = 1 Then
				GetData(1,SQL) = "�����ӱ���������Ϊ����"
				GetData(2,SQL) = 1
		End If
		Response.Write img & "<a href=" & HomeUrl & "a/a.asp?B=" & GetData(3,SQL) & "&id=" & GetData(0,SQL) & NewWindow & ">" & GetData(1,SQL) & "</a><br>"
	Next
	Response.Write """);"

End Sub

Sub Topic_BBSInfo

	Dim GBL_OnlineUserNum,GBL_UserCount,GBL_MaxOnline,GBL_MaxolTime,GBL_OnlineTime,GBL_PageCount,GBL_UploadNum
	Dim GBL_MaxAnnounce,GBL_MaxAncTime,GBL_YesterdayAnc

	Dim TmpData
	If isArray(application(DEF_MasterCookies & "StatisticData")) = False Then ReloadStatisticData
	TmpData = Application(DEF_MasterCookies & "StatisticData")
	GBL_OnlineTime = cCur(TmpData(0,0))
	GBL_UserCount = cCur(TmpData(1,0))
	GBL_MaxOnline = cCur(TmpData(2,0))
	GBL_MaxolTime = cCur(TmpData(3,0))
	GBL_PageCount = cCur(TmpData(4,0))
	GBL_UploadNum = cCur(TmpData(5,0))
	GBL_MaxAnnounce = cCur(TmpData(6,0))
	GBL_MaxAncTime = cCur(TmpData(7,0))
	GBL_YesterdayAnc = cCur(TmpData(8,0))

	Dim GBL_TodayAnnounce,GBL_AnnounceNum,GBL_TopicNum	
	GBL_AnnounceNum = cCur(TmpData(9,0))
	GBL_TopicNum = cCur(TmpData(10,0))
	GBL_TodayAnnounce = cCur(TmpData(11,0))

	Dim img,Temp
	img = Replace(Replace(Left(Request.QueryString("img"),100),"\",""),"""","")
	If img <> "" Then img = "<img src=" & Replace(img,"\","\\") & " align=middle> "

	Response.Write "document.write("""
	Response.Write img & "����������" & GBL_AnnounceNum & "��"
	Response.Write "<br>" & img & "�������ӣ�" & GBL_TopicNum & "��"
	Response.Write "<br>" & img & "ע���û���" & GBL_UserCount & "��"
	Response.Write "<br>" & img & "�û����ߣ�"

	If isNumeric(application(DEF_MasterCookies & "SiteOlTime")) = False Then
		Application.Lock
		application(DEF_MasterCookies & "SiteOlTime") = 0
		Application.UnLock
	End If
	GBL_OnlineTime = GBL_OnlineTime + application(DEF_MasterCookies & "SiteOlTime")
	Temp = (GBL_OnlineTime)/(24*60*60)
	Response.Write Fix(Temp) & "��"
	'GBL_OnlineTime=GBL_OnlineTime-Fix(Temp)*24*60*60
	'Temp = GBL_OnlineTime/(60*60)
	'Response.Write Fix(Temp) & "ʱ"
	'GBL_OnlineTime=GBL_OnlineTime-Fix(Temp)*60*60
	'Temp = GBL_OnlineTime/60
	'Response.Write Fix(Temp) & "��"
	
	Response.Write "<br>" & img & "ҳ����ʣ�" & (GBL_PageCount+application(DEF_MasterCookies & "SitePageCount")) & "��"
	Response.Write "<br>" & img & "��ǰ���ߣ�" & application(DEF_MasterCookies & "ActiveUsers") & "��"
	Response.Write "<br>" & img & "������ߣ�" & GBL_MaxOnline & "��"
	'Response.Write "<br>" & img & "����շ�����" & GBL_MaxAnnounce & "��"
	Response.Write "<br>" & img & "���շ�����<b><font color=Red class=redfont>" & GBL_TodayAnnounce & "</font></b>��"
	Response.Write "<br>" & img & "�ϴ�������" & GBL_UploadNum & "��"

	If Request.QueryString("exectime") = "1" Then Response.Write "<br>" & img & "ִ��ʱ�䣺" & FormatNumber(cCur(Timer - DEF_PageExeTime1),4,True)*1000 & "����"
	Response.Write """);"

End Sub

Sub Topic_PicInfo

	Dim Num,Height,Width
	Num = Request.QueryString("Num")
	If isNumeric(Num) = 0 Then Num = 0
	Num = Fix(cCur(Num))
	If Num < 1 or Num > 50 Then Num = 6
	
	
	Height = Request.QueryString("h")
	If isNumeric(Height) = 0 Then Height = 140
	Height = Fix(cCur(Height))
	
	Width = Request.QueryString("w")
	If isNumeric(Width) = 0 Then Width = 140
	Width = Fix(cCur(Width))
	
	If Height < 1 Then Height=105
	If Width < 1 Then Width=140

	Dim Rs,SQL,GetData
	SQL = sql_select("Select U.ID,U.PhotoDir,U.SPhotoDir,U.NdateTime,A.Title,U.AnnounceID,A.BoardID from LeadBBS_Upload as U left Join LeadBBS_Announce As A on U.AnnounceID=A.ID where U.FileType=0 Order by U.ID DESC",Num)
	Set Rs = Con.ExeCute(SQL)
	If Not Rs.Eof Then
		GetData = Rs.GetRows(-1)
		Rs.Close
		Set Rs = Nothing
	Else
		Rs.Close
		Set Rs = Nothing
		Exit Sub
	End If
	SQL = Ubound(GetData,2)
	%>

document.write("<a href=<%
If cCur(GetData(5,0)) <> 0 Then
	Response.Write HomeUrl & "a/a.asp?b=" & GetData(6,0) & "&id=" & GetData(5,0)
Else
	Response.Write HomeUrl & DEF_BBS_UploadPhotoUrl & Replace(GetData(1,0),"\","/")
End If%> id=topic_imgUrl name=topic_imgUrl target=_blank><img src=\"<%Response.Write HomeUrl & DEF_BBS_UploadPhotoUrl
	If GetData(2,0) <> "" Then
		Response.Write Replace(GetData(2,0),"\","/")
	Else
		Response.Write Replace(GetData(1,0),"\","/")
	End If%>\"<%If DEF_EnableGFL = 1 Then
		Else%> width=<%=Width%> height=<%=Height%><%End If%> id=topic_img name=topic_img style=\"border:1 gold ridge\"></a>");
var topic_time = 5000;//�趨��ʱ����
var topic_index;//Ŀǰ��ʾ��ͼƬ���
var topic_n;//�ܹ���ͼƬ��
var duration = 3;//blendTrans�˾�ʹ�õ�һ������ֵ���������˵��

//ͼƬ���������
  function ImgArray(len)
  {
   this.length=len;
  }

//�趨��������ͼƬ��
var ImgName=new ImgArray(<%=SQL+1%>);
var ImgName_Open=new ImgArray(<%=SQL+1%>);
<%For Rs = 0 To SQL%>
  ImgName[<%=Rs%>]="<%Response.Write HomeUrl & DEF_BBS_UploadPhotoUrl
	If GetData(2,Rs) <> "" Then
		Response.Write Replace(GetData(2,Rs),"\","/")
	Else
		Response.Write Replace(GetData(1,Rs),"\","/")
	End If%>";
	ImgName_Open[<%=Rs%>]="<%
	If cCur(GetData(5,Rs)) <> 0 Then
		Response.Write HomeUrl & "a/a.asp?b=" & GetData(6,Rs) & "&id=" & GetData(5,Rs)
	Else
		Response.Write HomeUrl & DEF_BBS_UploadPhotoUrl & Replace(GetData(1,Rs),"\","/")
	End If%>";
<%Next%>
 
topic_index=0;
topic_n=<%=SQL%>;
$ = function(id){
	return document.getElementById(id);
}
var is_ie = (navigator.userAgent.indexOf('msie') != -1);
  

function topic_playImg()
{
	if(topic_index==topic_n)
	{
	topic_index=0;
	}
	else
	{
		topic_index++;
	}
	if(is_ie)$id('topic_img').style.filter="blendTrans(Duration=duration)";
	if(is_ie)$id('topic_img').filters[0].apply();
	$id('topic_img').src=ImgName[topic_index];
	if(is_ie)$id('topic_img').filters[0].play();
	$id('topic_imgUrl').href=ImgName_Open[topic_index];
	mytimeout=setTimeout("topic_playImg()",topic_time);
}
setTimeout("topic_playImg()",topic_time);
	<%

End Sub%>