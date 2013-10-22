<!-- #include file=../inc/BBSsetup.asp -->
<!-- #include file=../inc/Board_popfun.asp -->
<%DEF_BBS_HomeUrl = "../"%>
<!-- #include file=inc/Mini_Board.asp -->
<!-- #include file=inc/Mini_Announce.asp -->
<%
Class Mini_Parameter

	Public BoardID,ID,R,UpFlag,p,q,Num,Act,AFirst

	'初始化全局变量，从地址栏获取全部参数
	Private Sub Class_Initialize
		BoardID = 0
		ID = 0
		R = 0
		UpFlag = 0
		p = 0
		q = 0
		Num = 0
		AFirst = 0

		Dim Str,N
		Str = Left(Request.QueryString,150)
		Str = Split(Str,"-")
		If isArray(Str) Then
			N = Ubound(Str,1)
			If N >= 7 Then
				BoardID = Str(0)
				If isNumeric(BoardID) = 0 Then BoardID = 0
				BoardID = Fix(cCur(BoardID))
				
				ID = Str(1)
				If isNumeric(ID) = 0 Then ID = 0
				ID = Fix(cCur(ID))
				
				R = Str(2)
				If isNumeric(R) = 0 Then R = 0
				R = Fix(cCur(R))
				
				UpFlag = Str(3)
				If isNumeric(UpFlag) = 0 Then UpFlag = 0
				UpFlag = Fix(cCur(UpFlag))
				
				p = Str(4)
				If isNumeric(p) = 0 Then p = 0
				p = Fix(cCur(p))
				
				q = Str(5)
				If isNumeric(q) = 0 Then q = 0
				q = Fix(cCur(q))
				
				Num = Str(6)
				If isNumeric(Num) = 0 Then Num = 0
				Num = Fix(cCur(Num))
				
				Act = LCase(Str(7))
				If Act <> "b" and Act <> "a" and Act <> "h" Then Act = "h"
			End If
			If N >= 8 Then
				AFirst = LCase(Str(8))
				If isNumeric(AFirst) = 0 Then AFirst = 0
				AFirst = Fix(cCur(AFirst))
			End If
		End If
		If BoardID > 0 Then
			GBL_board_ID = BoardID
			Borad_GetBoardIDValue(GBL_board_ID)
		End If
	
	End Sub
	
	'根据变量重新生成新的地址栏参数
	Public Function GetPar(BoardID,ID,R,UpFlag,p,q,Num,Act)
	
		GetPar = BoardID & "-" & ID & "-" & R & "-" & UpFlag & "-" & p & "-" & q & "-" & Num & "-" & Act & "-.htm"
	
	End Function

End Class

Class Mini_DisplayBoard

	Private LoopN,LowBoardString,LastAssosrt,CurrentAssosrt

	Private Sub Class_Initialize
	
		LoopN = 0
		LowBoardString = ""
		LastAssosrt = 0
		CurrentAssosrt = 0
	
	End Sub
	
	'显示论坛列表
	Public Sub BoardList()
	
		Dim Rs,GetData,BoardNum
		Set Rs = LDExeCute("Select BoardID,BoardAssort,BoardName,BoardIntro,LastWriter,LastWriteTime,TopicNum,AnnounceNum,ForumPass,HiddenFlag,LastAnnounceID,LastTopicName,MasterList,BoardLimit,LeadBBS_Assort.AssortID,LeadBBS_Assort.AssortName,LowerBoard from LeadBBS_Boards left join LeadBBS_Assort on LeadBBS_Assort.AssortID=LeadBBS_Boards.BoardAssort where LeadBBS_Boards.ParentBoard=0 and LeadBBS_Boards.HiddenFlag = 0 order by LeadBBS_Boards.BoardAssort,LeadBBS_Boards.OrderID ASC",0)
		If Not Rs.Eof Then
			GetData = Rs.GetRows(-1)
			BoardNum = Ubound(GetData,2)
		Else
			BoardNum = -1
		End If
		Rs.Close
		Set Rs = Nothing
	
		'on error resume next
		Dim TempStr
		TempStr = ""
	
		Response.Write "	<ul style='line-height:1.0;'>" & VbCrLf
		Response.Write "	<b>" & DEF_BBS_Name & "</b><br>" & VbCrLf
	
		If BoardNum = -1 Then
		Else
			Dim CurrentAssosrt,N
			CurrentAssosrt = -1183
			Dim LastAssosrt,WriteStr
			LastAssosrt = cCur(GetData(1,BoardNum))
			Dim LastFlag
			For N = 0 to BoardNum
				WriteStr = ""
				If CurrentAssosrt<>cCur(GetData(1,N)) Then
					CurrentAssosrt = cCur(GetData(1,N))
					If LastAssosrt = CurrentAssosrt Then
						WriteStr = "└┬"
					Else
						WriteStr = "├┬"
					End If
					Response.Write "		<span class=TBBG1>" & WriteStr & KillHTMLLabel(GetData(15,N)) & "</span><br>" & VbCrLf
				End If
				If N >= BoardNum Then
					If LastAssosrt = CurrentAssosrt Then
						If GetData(16,n) & ""  = "" Then
							WriteStr = "　└"
						Else
							WriteStr = "　├"
						End if
					Else
						WriteStr = "│└"
					End If
				Else
					If CurrentAssosrt<>cCur(GetData(1,N+1)) Then
						If LastAssosrt = CurrentAssosrt Then
							WriteStr = "　└"
						Else
							WriteStr = "│└"
						End If
					Else
						If LastAssosrt = CurrentAssosrt Then
							WriteStr = "　├"
						Else
							WriteStr = "│├"
						End If
					End If
				End If
				WriteStr = WriteStr & KillHTMLLabel(GetData(2,N))
				'If StrLength(WriteStr) > 21 Then
				'	WriteStr = LeftTrue(WriteStr,18) & "..."
				'End If
				Response.Write "		<a href=Default.asp?" & M_Par.GetPar(GetData(0,N),0,0,0,0,0,0,"b") & ">" & WriteStr & "</a><br>" & VbCrLf
				LowBoardString = ""
				LoopN = 0
				GetLowBoardString_Move GetData(16,n)
				If LowBoardString <> "" Then Response.Write LowBoardString
				
			Next
		End If
	
		Response.Write "	</ul>" & VbCrLf
	
	End Sub
	
	Public Function GetLowBoardString_Move(LowBoardStr)
	
		If LowBoardStr = "" or isNull(LowBoardStr) or LoopN > 100 Then Exit Function
		LoopN = LoopN + 1
		Dim BoardNum,LowArray,N
		LowArray = Split(LowBoardStr,",")
		BoardNum = Ubound(LowArray,1)
	
		Dim Temp
		Dim WriteStr
		For N = 0 to BoardNum
			Temp = Application(DEF_MasterCookies & "BoardInfo" & LowArray(N))
			If isArray(Temp) = False Then
				ReloadBoardInfo(LowArray(N))
				Temp = Application(DEF_MasterCookies & "BoardInfo" & LowArray(N))
			End If
			If isArray(Temp) = True Then
				If Temp(8,0) = 0 Then
					If N >= BoardNum Then
						If LastAssosrt = CurrentAssosrt Then
							WriteStr = "│" & String(LoopN, "│") & "├"
						Else
							WriteStr = "│" & String(LoopN, "│") & "├"
						End If
					Else
						If LastAssosrt = CurrentAssosrt Then
							WriteStr = "│├"
						Else
							WriteStr = "│" & String(LoopN, "│") & "├"
						End If
					End If
					'WriteStr = String(LoopN, "　") & WriteStr
					WriteStr = WriteStr & KillHTMLLabel(Temp(0,0))
					'If StrLength(WriteStr) > 21 Then
					'	WriteStr = LeftTrue(WriteStr,18) & "..."
					'End If
					LowBoardString = LowBoardString & "		<a href=Default.asp?" & M_Par.GetPar(LowArray(N),0,0,0,0,0,0,"b") & ">" & WriteStr & "</a><br>" & VbCrLf
					GetLowBoardString_Move Temp(27,0)
				End If
			End If
		Next
			
		LoopN = LoopN - 1
		
	End Function

End Class

Class Mini_PageDefine

	'显示页面头部代码
	Public Sub PageHead(BoardID,Str,headStr)
%>
<html>
<head>
	<meta HTTP-EQUIV="Content-Type" content="text/html; charset=gb2312">
	<meta name="description" content="<%=htmlencode(DEF_GBL_Description)%>" />
	<title>
		<%=headStr%>
	</title>
	<link rel="stylesheet" type="text/css" href="inc/STYLE_MINI.CSS">
	<style>table {WORD-BREAK: break-all;}</style>
	<script src="<%=DEF_BBS_HomeUrl%>inc/js/jquery.js" type="text/javascript"></script>
	<script src="<%=DEF_BBS_HomeUrl%>inc/js/common.js" type="text/javascript"></script>
</head>

<body bgcolor="#f7f7f7" leftmargin=30 topmargin=30 marginwidth=30 marginheight=30 class=TBBGbody>
<a href="<%

			If GBL_Board_BoardAssort = "" Then
				Response.Write DEF_SiteHomeUrl & """><font class=NavColor>" & DEF_SiteNameString & "</font></a> &gt;&gt; <a href=""Default.asp?" & M_Par.GetPar(0,0,0,0,0,0,0,"h") & """><font class=NavColor>" & DEF_BBS_Name & "</font></a>"
			Else
				Response.Write DEF_SiteHomeUrl & """><font class=NavColor>" & DEF_SiteNameString & "</font></a> &gt;&gt; <a href=""Default.asp?" & M_Par.GetPar(0,0,0,0,0,0,0,"h") & """><font class=NavColor>" & DEF_BBS_Name & "</font></a>"
				If GBL_Board_BoardName = "" Then 
					If GBL_Board_AssortName<>"" Then Response.Write " &gt;&gt; " & GBL_Board_AssortName
				Else
					If GBL_Board_AssortName<>"" Then Response.Write " &gt;&gt; <font class=NavColor>" & GBL_Board_AssortName & "</font>"
					If BoardID > 0 Then
						Dim Temp,TempStr,N
						Temp = cCur(Application(DEF_MasterCookies & "BoardInfo" & GBL_Board_ID)(26,0))
						Do While Temp > 0
							If isArray(Application(DEF_MasterCookies & "BoardInfo" & Temp)) = False Then
								ReloadBoardInfo(Temp)
								If isArray(Application(DEF_MasterCookies & "BoardInfo" & Temp)) = False Then Exit Do
							End If
							TempStr = " &gt;&gt; <a href=Default.asp?" & M_Par.GetPar(Temp,0,0,0,0,0,0,"b") & "><font class=NavColor>" & Application(DEF_MasterCookies & "BoardInfo" & Temp)(0,0) & "</font></a>" & TempStr
							Temp = cCur(Application(DEF_MasterCookies & "BoardInfo" & Temp)(26,0))
							N = N + 1
							If N > 10 Then Exit Do
						Loop
						Response.Write TempStr
						Response.Write " &gt;&gt; <a href=""Default.asp?" & M_Par.GetPar(GBL_Board_ID,0,0,0,0,0,0,"b") & """><font class=NavColor>" & GBL_Board_BoardName & "</font></a>"
					End If
				End If
			End If
			Response.write Str
			%>
<hr size=1>
<%

	End Sub

	'显示页面尾部代码
	Public Sub PageBottom%>
<hr size=1>
<%
		Response.Write "<p align=center><b>[" & FullStr & "]</b></p>"
		SiteBottom_Spend

	End Sub

End Class

Dim M_Par
Dim MiniPageDefine
Dim FullStr

Sub Main

	OpenDatabase
	Set M_Par = New Mini_Parameter
	Set MiniPageDefine = New Mini_PageDefine
	
	Select Case M_Par.Act

		Case "a"
			Dim M_Anc
			Set M_Anc = New Mini_Announce
			M_Anc.GetTopicInfo
			MiniPageDefine.PageHead M_Par.BoardID," >> " & M_Anc.LMT_TopicName,M_Anc.LMT_TopicName
			M_Anc.DisplayTopic
			Set M_Anc = Nothing
			
			FullStr = "<a href=../a/a.asp?B=" & M_Par.BoardID
			If M_Par.ID > 0 Then FullStr = FullStr & "&ID=" & M_Par.ID
			If M_Par.AFirst > 0 Then FullStr = FullStr & "&Afirst=" & M_Par.AFirst
			If M_Par.UpFlag > 0 Then FullStr = FullStr & "&AUpFlag=" & M_Par.UpFlag
			If M_Par.Num > 0 Then FullStr = FullStr & "&ANum=" & M_Par.Num
			If M_Par.p > 0 Then FullStr = FullStr & "&Ap=" & M_Par.p
			If M_Par.q > 0 Then FullStr = FullStr & "&Aq=" & M_Par.q
			If M_Par.r > 0 Then FullStr = FullStr & "&Ar=" & M_Par.r
			FullStr = FullStr & ">查看完整模式</a>"
				
		Case "b"
			MiniPageDefine.PageHead M_Par.BoardID,"",GBL_Board_BoardName 
			Dim M_Board
			Set M_Board = New Mini_Board
			M_Board.List
			Set M_Board = Nothing
			
			FullStr = "<a href=../b/b.asp?B=" & M_Par.BoardID
			If M_Par.UpFlag > 0 Then FullStr = FullStr & "&UpFlag=" & M_Par.UpFlag
			If M_Par.Num > 0 Then FullStr = FullStr & "&Num=" & M_Par.Num
			If M_Par.p > 0 Then FullStr = FullStr & "&p=" & M_Par.p
			If M_Par.q > 0 Then FullStr = FullStr & "&q=" & M_Par.q
			If M_Par.r > 0 Then FullStr = FullStr & "&r=" & M_Par.r
			FullStr = FullStr & ">查看完整模式</a>"
		Case Else
			MiniPageDefine.PageHead M_Par.BoardID,"",DEF_SiteNameString
			Dim MiniBoard
			Set MiniBoard = New Mini_DisplayBoard
			MiniBoard.BoardList
			Set MiniBoard = Nothing
			
			FullStr = "<a href=../Boards.asp>查看完整模式</a>"
	End Select
	
	CloseDatabase
	Set M_Par = Nothing
	MiniPageDefine.PageBottom
	Set MiniPageDefine = Nothing

End Sub

Main
%>