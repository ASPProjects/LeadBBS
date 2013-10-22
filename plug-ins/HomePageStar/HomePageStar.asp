<!--#include file="Inc/StarSetup.asp"-->
<%
Dim GBL_PLUG_HPS_DataOne,GBL_PLUG_HPS_DataTwo
GBL_PLUG_HPS_DataOne = Application(DEF_MasterCookies & "_PLUG_HPS_DAY")
GBL_PLUG_HPS_DataTwo = Application(DEF_MasterCookies & "_PLUG_HPS_OTHER")
Dim GBL_PLUG_HPS_Str1,GBL_PLUG_HPS_Str2
Dim GBL_PLUG_HPS_RefreshEnable
GBL_PLUG_HPS_RefreshEnable = 0

Sub FUN_PLUG_HPS_GetDayStarData

	Dim TimeM
	TimeM = Application(DEF_MasterCookies & "_PLUG_HPS_M")
	If isNull(TimeM) or isNumeric(TimeM) = False Then
		TimeM = 0
		TimeM = Minute(DEF_Now)
		Application.Lock
		Application(DEF_MasterCookies & "_PLUG_HPS_M") = TimeM
		Application.UnLock
	Else
		If (isArray(GBL_PLUG_HPS_DataTwo) = False and GBL_PLUG_HPS_Str2 <> "") or (isArray(GBL_PLUG_HPS_DataOne) = False and GBL_PLUG_HPS_Str1 <> "") or (Minute(DEF_Now) - TimeM) >= GBL_PLUG_HPS_RefreshSpace or (Minute(DEF_Now) - TimeM) < 0 Then
		Else
			If GBL_PLUG_HPS_RefreshEnable = 0 Then Exit Sub
		End If
	End If

	GBL_PLUG_HPS_RefreshEnable = 1

	Dim SQL,Rs,N
	If GBL_PLUG_HPS_Str1 <> "" Then
		Set Rs = LDExeCute(GBL_PLUG_HPS_Str1,0)
		If Not Rs.Eof Then
			GBL_PLUG_HPS_DataOne = Rs.GetRows(GBL_PLUG_HPS_TopMax)
			Rs.Close
			Set Rs = Nothing
			For N = 0 To Ubound(GBL_PLUG_HPS_DataOne,2)
				Set Rs = LDExeCute(sql_select("Select UserName from LeadBBS_User where ID=" & GBL_PLUG_HPS_DataOne(0,N),1),0)
				If Not Rs.Eof Then
					GBL_PLUG_HPS_DataOne(0,N) = Rs(0)
				Else
					GBL_PLUG_HPS_DataOne(0,N) = ""
				End If
				Rs.Close
				Set Rs = Nothing
			Next
			Application.Lock
			Application(DEF_MasterCookies & "_PLUG_HPS_DAY") = GBL_PLUG_HPS_DataOne
			Application.UnLock
		Else
			Rs.Close
			Set Rs = Nothing
		End If
	End If

	If (GBL_PLUG_HPS_Str2 <> "") Then
		Set Rs = LDExeCute(GBL_PLUG_HPS_Str2,0)
		If Not Rs.Eof Then
			GBL_PLUG_HPS_DataTwo = Rs.GetRows(GBL_PLUG_HPS_TopMax)
			Rs.Close
			Set Rs = Nothing
			For N = 0 To Ubound(GBL_PLUG_HPS_DataTwo,2)
				Set Rs = LDExeCute(sql_select("Select UserName from LeadBBS_User where ID=" & GBL_PLUG_HPS_DataTwo(0,N),1),0)
				If Not Rs.Eof Then
					GBL_PLUG_HPS_DataTwo(0,N) = Rs(0)
				Else
					GBL_PLUG_HPS_DataTwo(0,N) = ""
				End If
				Rs.Close
				Set Rs = Nothing
			Next
			Application.Lock
			Application(DEF_MasterCookies & "_PLUG_HPS_OTHER") = GBL_PLUG_HPS_DataTwo
			Application.UnLock
		Else
			Rs.Close
			Set Rs = Nothing
		End If
	End If

	Application.Lock
	Application(DEF_MasterCookies & "_PLUG_HPS_M") = Minute(DEF_Now)
	Application.UnLock

End Sub

Sub LeadBBSHomePageStar()

	If GBL_PLUG_HPS_ShowType < 1 and CheckSupervisorUserName = 0 Then Exit Sub
	
	If ((GBL_PLUG_HPS_ShowType = 1 or GBL_PLUG_HPS_ShowType = 3) and (GBL_PLUG_HPS_LineSecondType > 0)) or ((GBL_PLUG_HPS_ShowType = 1 or GBL_PLUG_HPS_ShowType = 2) and (GBL_PLUG_HPS_LineFirstType > 0)) Then
	Else
		If CheckSupervisorUserName = 0 Then Exit Sub
	End If

	Dim CFlag
	If inStr(OpenAssort,",bstar,") > 0 or (GBL_PLUG_HPS_Collapse = 0 and inStr(CloseAssort,",bstar,") = 0) Then
		CFlag = 0
	Else
		CFlag = 1
	End If%>
	<div class="contentbox">
	<table border="0" cellspacing="0" cellpadding="0" width="100%" class="tablebox">
	<tr class="tbhead">
		<td><script>var bstar=<%
			If CFlag = 1 Then
				Response.Write "0"
			Else
				Response.Write "1"
			End If%>;</script>
			<div class="b_assort">
			<div class="b_assort_title"><span class="clicktext" title="�ر�/չ��" onclick="bstar=(bstar==0)?1:0;LD.blist.assort_click('bstar',bstar,1);"><img src="<%=DEF_BBS_HomeUrl%>images/blank.gif" id="b_assort_img_bstar" class="b_assort_close<%
			If CFlag = 1 Then Response.Write "_swap"%>" alt="�ر�/չ��" /></span>
				<b>��������</b>
			</div>
			</div>
		</td>
	</tr>
	</table>
	<div id="b_assort_bstar"<%If CFlag = 1 Then  Response.Write " style=""display:none"""%>>
	<% 
	If Request("action1") <> "HomePageStar" and (GBL_PLUG_HPS_ShowType = 1 or GBL_PLUG_HPS_ShowType = 2) and (GBL_PLUG_HPS_LineFirstType > 0) Then
		%><table border="0" cellspacing="0" cellpadding="0" width="100%" class="tablebox">
		<tr><td class="tdbox">
			<div class="b_list_box">
				<%Call FUN_PLUG_HPS_HomePageStar()%>
			</div>
		</td>
		</tr></table>
		<%
	End If
	If (GBL_PLUG_HPS_ShowType = 1 or GBL_PLUG_HPS_ShowType = 3) and (GBL_PLUG_HPS_LineSecondType > 0) Then
	%>
				<%FUN_PLUG_HPS_HomePageStarTop%>
	<%
	End If%>
	<%
	If CheckSupervisorUserName = 1 Then%>
		<table border="0" cellspacing="0" cellpadding="0" width="100%" class="tablebox">
		<tr>
			<td class="tdbox" align="right">
				<div class="b_list_box">
				<a href="plug-ins/HomePageStar/admin_HomePageStar.asp">��ʾ��ʽ����</a>
				</div>
			</td>
		</tr>
		</table>
	<%End If%>
	</div>
	</div>
    <%

End Sub

Sub FUN_PLUG_HPS_HomePageStar()

	'ͷ���ļ�
	%>
	<table border="0" cellspacing="0" cellpadding="0" class="blanktable" width="100%"><tr>
	<%

	Dim SQL,Rs,F_or_M,UserName,AnnounceNum,TempLineStr,UserID,UserID2,AnnounceNum2,UserName2
	Dim NTime,WTime,YTime,MTime

	'ÿ�շ���
	NTime = cCur(Left(GetTimeValue(DEF_Now),8) & "000000")

	'ÿ�ܹ�ˮ
	WTime = cCur(Left(GetTimeValue(DateAdd("d",0-WeekDay(DEF_Now,2),DEF_Now)),8) & "000000")

	'���귢��
	YTime = cCur(Left(GetTimeValue(DEF_Now),4) & "0000000000")

	'���·���
	MTime = cCur(Left(GetTimeValue(DEF_Now),6) & "00000000")
	'��ʾ��һ�е�ͷ��
	Select Case GBL_PLUG_HPS_LineFirstType
		Case 1
			TempLineStr = "����"
			SQL = sql_select("Select UserID,Count(UserID) from LeadBBS_Announce Where NDateTime>=" & NTime & " Group By UserID Order by Count(UserID) DESC",GBL_PLUG_HPS_TopMax)
		Case 2
			TempLineStr = "����"
			SQL = sql_select("Select UserID,Count(UserID) from LeadBBS_Announce Where NDateTime>=" & WTime & " Group By UserID Order by Count(UserID) DESC",GBL_PLUG_HPS_TopMax)
		Case 3
			TempLineStr = "����"
			SQL = sql_select("Select UserID,Count(UserID) from LeadBBS_Announce Where NDateTime>=" & MTime & " Group By UserID Order by Count(UserID) DESC",GBL_PLUG_HPS_TopMax)
		Case 4
			TempLineStr = "����"
			SQL = sql_select("Select UserID,Count(UserID) from LeadBBS_Announce Where NDateTime>=" & YTime & " Group By UserID Order by Count(UserID) DESC",GBL_PLUG_HPS_TopMax)
		Case 5
			TempLineStr = "����"
			SQL = sql_select("Select ID,AnnounceNum from LeadBBS_User Order By AnnounceNum DESC",GBL_PLUG_HPS_TopMax)
		'Case 6
		'	TempLineStr = "˧��"
		'		sql_select(SQL = "Select ID,AnnounceNum ID from LeadBBS_User Where Sex='��' Order By AnnounceNum DESC",GBL_PLUG_HPS_TopMax)
		'Case 7
		'	TempLineStr = "����"
		'	SQL = sql_select("Select ID,AnnounceNum from LeadBBS_User Where Sex='Ů' Order By AnnounceNum DESC",GBL_PLUG_HPS_TopMax)
		Case Else
			GBL_PLUG_HPS_LineFirstType = 1
			TempLineStr = "����"
			SQL = sql_select("Select UserID,Count(UserID) from LeadBBS_Announce Where NDateTime>=" & NTime & " Group By UserID Order by Count(UserID) DESC",GBL_PLUG_HPS_TopMax)
	End Select
	GBL_PLUG_HPS_Str1 = SQL
	FUN_PLUG_HPS_GetDayStarData

	Dim GetData,GetDataUserData1,GetDataUserData2
	GetData = GBL_PLUG_HPS_DataOne

	If isArray(GetData) = False Then
		%>
			<td valign="middle" align="center" width="20%"><img src="<%=DEF_BBS_HomeUrl%>images/face/0000.gif" alt="ͷ��" /></td>
			<td valign="top" width="30%"><strong>
			<%=TempLineStr%>��ˮ״Ԫ
			</strong><br /><br />
			�û�������<span class="bluefont">��������д</span><br />
			
			����<%=DEF_PointsName(3)%>��<img src="<%=DEF_BBS_HomeUrl%>images/lvstar/level0.gif" alt="�ȼ�" />
			<br /><%=TempLineStr%>������0 ƪ
			<br />����<%=DEF_PointsName(0)%>����
			<br />����<%=DEF_PointsName(1)%>����
	        	<br />����<%=DEF_PointsName(2)%>����
			<br />����<%=DEF_PointsName(4)%>����
			<br />E&nbsp;-&nbsp;Mail����
			</td>
		<%
	Else
		Dim Flag,LoopN
		Flag = 0

		For LoopN = 0 to 2
			If LoopN > Ubound(GetData,2) or Flag = 2 Then Exit For
			If Flag = 1 Then
				UserName2 = GetData(0,LoopN)
				AnnounceNum2 = cCur("0" & GetData(1,LoopN))
				SQL = sql_select("Select Mail,Sex,UserPhoto,UserLevel,Points,OnlineTime,AnnounceNum,FaceWidth,FaceHeight,CharmPoint,FaceUrl,CachetValue,UserName,ID from LeadBBS_User Where UserName='" & Replace(UserName2,"'","''") & "'",1)
				Set Rs = LDExeCute(SQL,0)
				If Not Rs.Eof Then
					GetDataUserData2 = Rs.GetRows(1)
					Flag = 2
				Else
					Flag = 1
				End If
			ElseIf Flag = 0 Then
				UserName = GetData(0,LoopN)
				AnnounceNum = cCur("0" & GetData(1,LoopN))
				SQL = sql_select("Select Mail,Sex,UserPhoto,UserLevel,Points,OnlineTime,AnnounceNum,FaceWidth,FaceHeight,CharmPoint,FaceUrl,CachetValue,UserName,ID from LeadBBS_User Where UserName='" & Replace(UserName,"'","''") & "'",1)
				Set Rs = LDExeCute(SQL,0)
				If Not Rs.Eof Then
					GetDataUserData1 = Rs.GetRows(1)
					Flag = 1
				Else
					Flag = 0
				End If
			End If
		Next

		Dim Mail,Sex,UserPhoto,UserLevel,Points,OnlineTime,FaceWidth,FaceHeight,CharmPoint,FaceUrl,CachetValue
		If isArray(GetDataUserData1) Then
			Mail = GetDataUserData1(0,0)
			Sex = GetDataUserData1(1,0)
			UserPhoto = GetDataUserData1(2,0)
			UserLevel = GetDataUserData1(3,0)
			Points = GetDataUserData1(4,0)
			OnlineTime = GetDataUserData1(5,0)
			'AnnounceNum = GetDataUserData1(6,0)
			FaceWidth = GetDataUserData1(7,0)
			FaceHeight = GetDataUserData1(8,0)
			CharmPoint = GetDataUserData1(9,0)
			FaceUrl = GetDataUserData1(10,0)
			CachetValue = cCur(GetDataUserData1(11,0))
		Else
			UserName = "��"
			Mail = "��"
			Sex = "��"
			UserPhoto = ""
			UserLevel = 0
			Points = 0
			OnlineTime = 0
			AnnounceNum = 0
			FaceWidth = DEF_AllFaceMaxWidth
			FaceHeight = DEF_AllFaceMaxWidth*2
			CharmPoint = 0
			FaceUrl = ""
			CachetValue = 0
		End If

		If Sex = "��" Then
			F_or_M = DEF_BBS_HomeUrl & "images/sxmg/Male.gif"
		ElseIf Sex = "Ů" Then
			F_or_M = DEF_BBS_HomeUrl & "images/sxmg/FeMale.gif"
        Else
			F_or_M = DEF_BBS_HomeUrl & "images/sxmg/Male.gif"
		End If
		%>
			<td valign="middle" align="center" width="20%">
				<a href="User/LookUserInfo.asp?name=<%=urlEncode(UserName)%>" target="_blank">
				<img src="<%
		If FaceWidth > DEF_AllFaceMaxWidth Then FaceWidth = DEF_AllFaceMaxWidth
		If FaceHeight > DEF_AllFaceMaxWidth*2 Then FaceHeight = DEF_AllFaceMaxWidth
		If DEF_AllDefineFace <> 0 and FaceUrl <> "" Then
			If Lcase(Left(FaceUrl,5)) <> "http:" Then
				FaceUrl = DEF_BBS_HomeUrl & Replace(htmlencode(FaceUrl),"../","")
			Else
				FaceUrl = htmlencode(FaceUrl)
			End If
			%><%=FaceUrl%>" width="<%=FaceWidth%>" height="<%=FaceHeight%>"<%
		Else
			%>images/face/<%=String(4-len(CStr(UserPhoto)),"0") & UserPhoto%>.gif"
		<%End If%> title="��ʲô��,�������ǣ���" alt="ͷ��" /></a>
			</td>
			<td valign="top" width="30%">
				<strong>
        			<%=TempLineStr%>��ˮ״Ԫ
        			</strong>
        			<br /><br />
				�û�������<span class="bluefont"><%=UserName%></span><br />
				����<%=DEF_PointsName(3)%>��<img src="images/<%=GBL_DefineImage%>lvstar/level<%=UserLevel%>.gif" align="middle" alt="�ȼ�" /><br />
				<%=TempLineStr%>������<%=AnnounceNum%> ƪ
				<br />����<%=DEF_PointsName(0)%>��<%=cCur(Points)%>
				<br />����<%=DEF_PointsName(1)%>��<%=CharmPoint%>
		<%
        	If CachetValue <> 0 Then
			If CachetValue > 0 Then
				CachetValue = "<span class=""bluefont"">��" & CachetValue & "</span>"
			End If
		Else
			CachetValue = "������Ŭ����" 
		End If
		%>
			<br />����<%=DEF_PointsName(2)%>��<%=CachetValue%>
			<br />����<%=DEF_PointsName(4)%>��<%=Fix(cCur(OnlineTime)/60)%>	
			<br />E&nbsp;-&nbsp;Mail��
		<%
		If Trim(Mail) <> "" Then
			Response.Write("<a href=""Mailto:" & htmlencode(Mail) & """>�ɸ봫��</a>")
		End If
		%>
			</td><%
	End If

	If isArray(GetDataUserData2) = False Then
		%>
			<td valign="middle" align="center" width="20%">
				<img src="<%=DEF_BBS_HomeUrl%>images/face/0000.gif" alt="ͷ��" />
			</td>
			<td valign="top" width="30%">
				<strong>
				<%=TempLineStr%>��ˮ����
				</strong>
				<br /><br />
				�û�������<span class="bluefont">��������д</span> <br />
			����<%=DEF_PointsName(3)%>��<img src="<%=DEF_BBS_HomeUrl%>images/lvstar/level0.gif" alt="�ȼ�" />
			<br /><%=TempLineStr%>������0 ƪ
			<br />����<%=DEF_PointsName(0)%>����
			<br />����<%=DEF_PointsName(1)%>����
			<br />����<%=DEF_PointsName(2)%>����
			<br />����<%=DEF_PointsName(4)%>����
			<br />E&nbsp;-&nbsp;Mail����
			</td>
		<%
  	Else
		AnnounceNum = AnnounceNum2
		UserName = UserName2
		If isArray(GetDataUserData2) Then
			Mail = GetDataUserData2(0,0)
			Sex = GetDataUserData2(1,0)
			UserPhoto = GetDataUserData2(2,0)
			UserLevel = GetDataUserData2(3,0)
			Points = GetDataUserData2(4,0)
			OnlineTime = GetDataUserData2(5,0)
			'AnnounceNum = GetDataUserData2(6,0)
			FaceWidth = GetDataUserData2(7,0)
			FaceHeight = GetDataUserData2(8,0)
			CharmPoint = GetDataUserData2(9,0)
			FaceUrl = GetDataUserData2(10,0)
			CachetValue = cCur(GetDataUserData2(11,0))
		Else
			UserName = "��"
			Mail = "��"
			Sex = "��"
			UserPhoto = ""
			UserLevel = 0
			Points = 0
			OnlineTime = 0
			AnnounceNum = 0
			FaceWidth = DEF_AllFaceMaxWidth
			FaceHeight = DEF_AllFaceMaxWidth*2
			CharmPoint = 0
			FaceUrl = ""
			CachetValue = 0
		End If

		If Sex = "��" Then
			F_or_M = DEF_BBS_HomeUrl & "images/sxmg/Male.gif"
		ElseIf Sex = "Ů" Then
			F_or_M = DEF_BBS_HomeUrl & "images/sxmg/FeMale.gif"
		Else
			F_or_M = DEF_BBS_HomeUrl & "images/sxmg/Male.gif"
		End If
		%>
			<td valign="middle" align="center" width="20%">
				<a href="User/LookUserInfo.asp?name=<%=urlEncode(UserName)%>" target="_blank">
				<img src="<% 
		If FaceWidth > DEF_AllFaceMaxWidth Then FaceWidth = DEF_AllFaceMaxWidth
		If FaceHeight > DEF_AllFaceMaxWidth*2 Then FaceHeight = DEF_AllFaceMaxWidth
		If DEF_AllDefineFace <> 0 and FaceUrl <> "" Then
			If Lcase(Left(FaceUrl,5)) <> "http:" Then
				FaceUrl = DEF_BBS_HomeUrl & Replace(htmlencode(FaceUrl),"../","")
			Else
				FaceUrl = htmlencode(FaceUrl)
			End If
			Response.Write FaceUrl & Chr(34) & " width=""" & FaceWidth & """ height=""" & FaceHeight & """"
		Else
			Response.Write "images/face/" & String(4-len(CStr(UserPhoto)),"0") & UserPhoto & ".gif"""
		End If%> title="��ʲô��,�������ǣ���" alt="ͷ��" /></a>
			</td>
			<td valign="top" width="30%">
				<strong>
				<%=TempLineStr%>��ˮ����
				</strong>
				<br /><br />
				�û�������<span class="bluefont"><%=UserName%></span>
				<br />
				����<%=DEF_PointsName(3)%>��<img src="images/<%=GBL_DefineImage%>lvstar/level<%=UserLevel%>.gif" align="middle" alt="�ȼ�" />
				<br />
				<%=TempLineStr%>������<%=AnnounceNum%> ƪ
				<br />����<%=DEF_PointsName(0)%>��<%=Points%>
				<br />����<%=DEF_PointsName(1)%>��<%=CharmPoint%>
		<%
		If CachetValue <> 0 Then
			If CachetValue > 0 Then
				CachetValue = "<span class=""bluefont"">��" & CachetValue & "</span>"
			End If
		Else
			CachetValue = "������Ŭ����"
		End If%>
			<br />����<%=DEF_PointsName(2)%>��<%=CachetValue%>
			<br />����<%=DEF_PointsName(4)%>��<%=Fix(cCur(OnlineTime)/60)%>	
			<br />E&nbsp;-&nbsp;Mail��
		<%
		If Trim(Mail) <> "" Then
	  		Response.Write("<a href=""Mailto:" & htmlencode(Mail) & """>�ɸ봫��</a>")
		End If
		%>
			</td><%
	End If

	'β���ļ�
	Response.Write("</tr></table>")

End Sub 

Sub FUN_PLUG_HPS_HomePageStarTop

	Dim NTime,WTime,YTime,MTime,SQL

	'ÿ�շ���
	NTime = cCur(Left(GetTimeValue(DEF_Now),8) & "000000")

	'ÿ�ܹ�ˮ
	WTime = cCur(Left(GetTimeValue(DateAdd("d",0-WeekDay(DEF_Now,2),DEF_Now)),8) & "000000")

	'���귢��
	YTime = cCur(Left(GetTimeValue(DEF_Now),4) & "0000000000")

	'���·���
	MTime = cCur(Left(GetTimeValue(DEF_Now),6) & "00000000")
	'��ʾ��һ�е�ͷ��
	Dim TempLineStr
	Select Case GBL_PLUG_HPS_LineSecondType
		Case 1
			TempLineStr = "����"
			SQL = sql_select("Select UserID,Count(UserID) from LeadBBS_Announce Where NDateTime>=" & NTime & " Group By UserID Order by Count(UserID) DESC",GBL_PLUG_HPS_TopMax)
		Case 2
			TempLineStr = "����"
			SQL = sql_select("Select UserID,Count(UserID) from LeadBBS_Announce Where NDateTime>=" & WTime & " Group By UserID Order by Count(UserID) DESC",GBL_PLUG_HPS_TopMax)
		Case 3
			TempLineStr = "����"
			SQL = sql_select("Select UserID,Count(UserID) from LeadBBS_Announce Where NDateTime>=" & MTime & " Group By UserID Order by Count(UserID) DESC",GBL_PLUG_HPS_TopMax)
		Case 4
			TempLineStr = "����"
			SQL = sql_select("Select UserID,Count(UserID) from LeadBBS_Announce Where NDateTime>=" & YTime & " Group By UserID Order by Count(UserID) DESC",GBL_PLUG_HPS_TopMax)
		Case 5
			TempLineStr = "����"
			SQL = sql_select("Select ID,AnnounceNum from LeadBBS_User Order By AnnounceNum DESC",GBL_PLUG_HPS_TopMax)
		Case Else
			GBL_PLUG_HPS_LineSecondType = 1
			TempLineStr = "����"
			SQL = sql_select("Select UserID,Count(UserID) from LeadBBS_Announce Where NDateTime>=" & NTime & " Group By UserID Order by Count(UserID) DESC",GBL_PLUG_HPS_TopMax)
	End Select
	GBL_PLUG_HPS_Str2 = SQL
	FUN_PLUG_HPS_GetDayStarData

	If isArray(GBL_PLUG_HPS_DataTwo) = False Then Exit Sub
	Dim N
	%>
	<table border="0" cellspacing="0" cellpadding="0" width="100%" class="tablebox">
	<tr>
	<td class="tdbox">
		<div class="b_list_box">
			<table border="0" cellspacing="0" cellpadding="0" width="100%">
			<tr>
			<td width="60">
			<b><%=TempLineStr%>��ˮ</b>
			</td>
			<td>
				<div class="b_assortlist">
<script type="text/javascript">
<!--
document.write("<ul>");
function hps(n,m){if(n=="")n="�ο�";document.write("<li><span class=\"bluefont\">��</span><a href=\"User/LookUserInfo.asp?name=" + escape(n) + "\" target=\"_blank\" title=\"�鿴�û�����\">" + n + "[<span class=\"redfont\">" + m + "</span>]</a></li>");}
<%
	For N = 0 to Ubound(GBL_PLUG_HPS_DataTwo,2)
		Response.Write "hps(""" & GBL_PLUG_HPS_DataTwo(0,N) & """," & GBL_PLUG_HPS_DataTwo(1,N) & ");" & VbCrLf
	Next%>
document.write("</ul>");
-->
</script>
				</div>
				</td>
			</tr>
			</table>
			</div>
		</td>
	</tr>
	</table>
	<%

End Sub%>