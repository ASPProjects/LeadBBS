<!-- #include file=../inc/BBSsetup.asp -->
<!-- #include file=../inc/Board_Popfun.asp -->
<%
'--------------------------------------------------
'LEADBBS RSS FOR 4.0 
'MODIFY TIME 2007-03-13
'--------------------------------------------------
Const RSS_ViewNumer = 12 '���������ʾ��RSS��¼����
DEF_BBS_HomeUrl = "../"

RSS_View

Sub RSS_View

	Dim MyHomeUrl
	MyHomeUrl = LCase(Request.Servervariables("SCRIPT_NAME"))
	If Right(MyHomeUrl,14) = "/other/rss.asp" Then
		If Request.ServerVariables("SERVER_PORT") <> "80" Then MyHomeUrl = ":" & Request.ServerVariables("SERVER_PORT") & MyHomeUrl
		MyHomeUrl = Lcase("http://"&Request.ServerVariables("server_name") & MyHomeUrl)
		MyHomeUrl = Replace(MyHomeUrl,"other/rss.asp","")
	Else
		MyHomeUrl = ""
	End If

	Dim BoardID
	BoardID = Request.QueryString("ID")
	If BoardID = "" Then BoardID = Request.QueryString("BoardID")
	If BoardID = "" Then BoardID = Request.QueryString("b")
	If isNumeric(BoardID) = 0 Then BoardID = 0
	BoardID = Fix(cCur(BoardID))

	OpenDatabase
	
	Dim Temp
	If BoardID > 0 Then
		Temp = Application(DEF_MasterCookies & "BoardInfo" & BoardID)
		If isArray(Temp) = False Then
			ReloadBoardInfo(BoardID)
			Temp = Application(DEF_MasterCookies & "BoardInfo" & BoardID)
		End If
		If isArray(Temp) = False Then BoardID = 0
	End If

	Dim SQLEndString

	Dim Rs,GetData,RssNum

	IF BoardID > 0 Then
		select case DEF_UsedDataBase
		case 0,2:
			SQLEndString = "and TA.BoardID=" & BoardID
		case Else
			SQLEndString = "where TA.BoardID=" & BoardID
		End select
	Else
		SQLEndString = ""
	End If
	select case DEF_UsedDataBase
		case 0,2:
			Set Rs = LDExeCute(sql_select("select TA.ID,TA.BoardID,TA.Title,TA.Content,TA.ndatetime,TA.LastTime,TA.UserName,TA.LastUser,TA.TitleStyle,TB.BoardName,TA.HTMLFlag,TB.BoardLimit,TB.ForumPass,TB.OtherLimit,TB.HiddenFlag from LeadBBS_Announce as TA left join LeadBBS_Boards as TB on TA.BoardID=TB.BoardID where TA.ParentID = 0 " & SQLEndString & " Order by TA.RootIDBak DESC",RSS_ViewNumer),0)
		case Else
			Set Rs = LDExeCute(sql_select("select TA.ID,TA.BoardID,TA.Title,'',TA.ndatetime,TA.LastTime,TA.UserName,TA.LastUser,TA.TitleStyle,TB.BoardName,0,TB.BoardLimit,TB.ForumPass,TB.OtherLimit,TB.HiddenFlag from LeadBBS_Topic as TA left join LeadBBS_Boards as TB on TA.BoardID=TB.BoardID " & SQLEndString & " Order by TA.ID DESC",RSS_ViewNumer),0)
	End select

	If Not rs.Eof Then
		GetData = Rs.GetRows(-1)
		RssNum = Ubound(GetData,2)
	Else
		RssNum = -1
	End If
	Rs.close
	Set Rs = Nothing
	CloseDatabase
	
	Dim PostTime
	Response.ContentType="application/xml"
	Response.Write "<?xml version=""1.0"" encoding=""gb2312""?>"
	'<?xml-stylesheet type="text/css" href="rss.css"?>
	'<?xml-stylesheet type="text/xsl" href="viewforfeed.xslt"?>
	%>
<rss version="2.0">
<channel>
	<%
	IF BoardID = 0 or RssNum = -1 Then
		Temp = ""
	Else
		Temp = " " & HtmlEncode(KillHTMLLabel(GetData(9,0)))
	End If
	Response.Write "<title><![CDATA[ " & HtmlEncode(DEF_SiteNameString & " " & DEF_BBS_Name) & Temp & " ]]></title>" & VbCrLf
	%>
<link><%=MyHomeUrl%></link>
<description><![CDATA[ <%
	If Temp = "" Then
		Response.Write "���а���"
	Else
		Response.Write "���棺" & Temp
	End IF%> ����<%=RSS_ViewNumer%>������ ]]></description>
<language>zh-cn</language>
<copyright>Copyright(C)LeadBBS.COM</copyright>
<webMaster>Info@LeadBBS.COM</webMaster>
<generator>LeadBBS.COM</generator>
<lastBuildDate><%=RestoreTime(GetTimeValue(DEF_Now))%></lastBuildDate>
<ttl>30</ttl>
<image>
<url><%=MyHomeUrl%>images/logo.gif</url>
<title><![CDATA[ <%=HtmlEncode(DEF_SiteNameString)%> ]]></title>
<link><%=MyHomeUrl%></link>
</image>
	<%
	IF RssNum = -1 Then
		Response.Write "<item></item>"
	Else
		Dim N
		For n = 0 to RssNum
			If GBL_CheckLimitTitle(GetData(12,N),GetData(11,N),GetData(13,N),GetData(14,N)) = 1 Then
				GetData(2,N) = "�鿴�����ӱ�����Ҫ����Ȩ��."
				GetData(7,N) = "����"
				GetData(6,N) = "����"
				GetData(10,N) = 1
			End If
			If GBL_CheckLimitContent(GetData(12,N),GetData(11,N),GetData(13,N),GetData(14,N)) = 1 Then GetData(3,N) = "�鿴������������Ҫ����Ȩ��"
			If GetData(8,N) = 1 Then GetData(2,N) = KillHTMLLabel(HtmlEncode(GetData(2,N)))
			if GetData(7,N) <> "" then GetData(7,N) = "���ظ���" & HtmlEncode(GetData(7,N)) & " at " & RestoreTime(GetData(5,N)) & VbCrLf
			Response.Write "<item>" & VbCrLf
			Response.Write "<title><![CDATA[ " & HtmlEncode(GetData(2,N)) & " ]]></title>" & VbCrLf
			Response.Write "<link>" & MyHomeUrl & "a/a.asp?b=" & GetData(1,N) & "&amp;ID=" & GetData(0,N) & "</link>" & VbCrLf
			Response.Write "<author><![CDATA[ " & HtmlEncode(GetData(6,N)) & " ]]></author>" & VbCrLf
			Response.Write "<category><![CDATA[ " & HtmlEncode(KillHTMLLabel(GetData(9,N))) & " ]]></category>" & VbCrLf
			Response.Write "<pubDate>" & RestoreTime(GetData(4,N)) & "</pubDate>" & VbCrLf
			Response.Write "<description><![CDATA[ " & GetData(7,N)
			Response.Write "<br>���ڰ��棺<a href=" & MyHomeUrl & "b/b.asp?b=" & GetData(1,N) & ">" & HtmlEncode(KillHTMLLabel(GetData(9,N))) & "</a>" & VbCrLf
			Response.Write "<br>�������ߣ�" & HtmlEncode(GetData(6,N)) & VbCrLf
			Response.Write "<br>������Ҫ��"
			If isNull(GetData(3,N)) Then GetData(3,N) = ""
			GetData(3,N) = Left(GetData(3,N),200)
			Select Case GetData(10,N)
				Case 1
					Response.Write Server.HtmlEncode(KillHTMLLabel(GetData(3,N)))
				Case 2
					Response.Write Server.HtmlEncode(clearUbbcode(GetData(3,N)))
				Case Else
					Response.Write Server.HtmlEncode(GetData(3,N))
			End Select
			Response.Write " ]]></description>" & VbCrLf
			Response.Write "</item>"
		Next
	End IF
	%>
<LeadBBS>
<ExeCuteTime>��ʱ<%=FormatNumber(cCur(Timer - DEF_PageExeTime1),3,True)%>��</ExeCuteTime>
<Query>����<%=GBL_DBNum%>��</Query>
</LeadBBS>
</channel>
</rss>
	<%

End Sub

Function clearUbbcode(str)

	Dim n,m,str2
	n = inStr(1,str,"[",0)
	if n > 0 Then
		m = inStr(n + 1,str,"]",0)
	Else
		m = 0
	End If
	str2 = str
	Do while n > 0 and n < m and m > 0
		str2 = Left(str2,n-1) & Mid(str2,m+1)
		n = inStr(1,str2,"[",0)
		if n > 0 Then
			m = inStr(n + 1,str2,"]",0)
		Else
			m = 0
		End If
	Loop
	clearUbbcode = str2

End Function
%>