<%
Const DEF_BBS_MaxListPage = 10 '������������ʾҳ��(�������Ӱ�����ܣ�Ĭ������Ϊ10)
Const DEF_BBS_MaxWords = 300 '�������������������Ҫ��ʾ����(�����ʾ�ֽ�)
Dim LMT_Key

Sub DisplayAnnouncesSplitPages

	Dim Page,First,ExcuteErr

	Page = Left(Request("Page"),14)
	If isNumeric(Page)=0 Then Page=0
	Page = cCur(Page)
	If Page > DEF_BBS_MaxListPage Then Page = DEF_BBS_MaxListPage

	Dim FullTextKey,Key,SessionKey
	Key = Left(Request("key"),100)
	Dim N
	FullTextKey = key
	LMT_Key = key
	SessionKey = ""
	If DEF_UsedDataBase = 0 and DEF_BBS_SearchMode = 2 and (key <> "") Then
		Dim Noise_Chinese
		Noise_Chinese = Array("~","!","@","#","$","%","^","&","*","(",")","_","+","=","`","[","]","{","}",";",":","""","'",",","<",">",".","/","\","|","?","_","about","1","2","3","4","5","6","7","8","9","0","a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z","after","all","also","an","and","another","any","are","as","at","be","because","been","before","being","between","both","but","by","came","can","come","could","did","do","each","for","from","get","got","had","has","have","he","her","here","him","himself","his","how","if","in","into","is","it","like","make","many","me","might","more","most","much","must","my","never","now","of","on","only","or","other","our","out","over","said","same","see","should","since","some","still","such","take","than","that","the","their","them","then","there","these","they","this","those","through","to","too","under","up","very","was","way","we","well","were","what","where","which","while","who","with","would","you","your","��","һ","��","��","��","��","��","Ϊ","��","��","��","��","��","��","֮","��","��","��","��","��","��","��","��","��","��","Ҳ","��","��","��","��","��","��","��","С","��","��","��","��","��","��","ȥ","��","ֻ","��","��","��","��","ĳ","��","��","��","��","��")

		FullTextKey = Replace(FullTextKey,"!"," ")
		FullTextKey = Replace(FullTextKey,"]"," ")
		FullTextKey = Replace(FullTextKey,"["," ")
		FullTextKey = Replace(FullTextKey,")"," ")
		FullTextKey = Replace(FullTextKey,"("," ")
		FullTextKey = Replace(FullTextKey,"��"," ")
		FullTextKey = Replace(FullTextKey,"-"," ")
		FullTextKey = Replace(FullTextKey,"/"," ")
		FullTextKey = Replace(FullTextKey,"+"," ")
		FullTextKey = Replace(FullTextKey,"="," ")
		FullTextKey = Replace(FullTextKey,","," ")
		FullTextKey = Replace(FullTextKey,"'"," ")
		LMT_Key = FullTextKey
	
		For N = 0 To Ubound(Noise_Chinese,1)
			If FullTextKey=Noise_Chinese(N) Then
				GBL_CHK_TempStr = "�������������<font color=ff0000 class=redfont>" & htmlencode(Noise_Chinese(N)) & "</font>���ں��Դʣ�ϵͳ�Ѿ������������̡�"
				Exit for
			End If
		Next
		
		Dim I,IFlag,TempKey
		If key<>"" then
			TempKey = split(FullTextKey," ")
			Redim LMT_Key(ubound(TempKey))
			LMT_Key = TempKey
			FullTextKey = ""
			for N = 0 to ubound(TempKey)
				IFlag = 1
				For I = 0 To Ubound(Noise_Chinese,1)
					If TempKey(N)=Noise_Chinese(I) Then
						'Response.Write "<br>����" & TempKey(N)
						IFlag = 0
						Exit for
					End If
				Next
				If IFlag = 1 and TempKey(N)<>"" Then FullTextKey = FullTextKey & TempKey(N) & " and "
			Next
			If Right(FullTextKey,5) = " and " Then FullTextKey = Left(FullTextKey,Len(FullTextKey) - 5)
		End If
	End If
	
	GBL_CHK_TempStr = ""
	Dim Mode,BoardID2
	Mode = Left(Request("Mode"),14)
	BoardID2 = ""
	If Mode <> "0" and Mode <> "1" and Mode <> "2" and Mode <> "3" Then Mode = "0"
	Mode = cCur(Mode)
	If FullTextKey<>"" and DEF_UsedDataBase = 0 and DEF_BBS_SearchMode = 2 Then
		If Sch_AllContent = 0 and Mode = 0 Then Mode = 1
		If Sch_AncContent = 0 and Mode = 2 Then Mode = 1
		If Sch_AncTitle = 0 and Mode = 1 Then Mode = 3
		If Sch_AncTitle <> 3 Then SessionKey = Left(Replace(FullTextKey," and "," "),20)
		Select Case Mode
			Case 0: SQLendString = "*"
			Case 1: SQLendString = "Title"
			Case 2: SQLendString = "Content"
			Case 3: CloseDatabase
				Response.Redirect "../User/LookUserInfo.asp?name=" & UrlEncode(key) & "&Evol=n"
			'Case Else: SQLendString = " right join containstable(LeadBBS_Announce,*,'" & Replace(FullTextKey,"'","''") & "'," & First + DEF_TopicContentMaxListNum + 1 & ") as T1 ON TT.id = T1.[KEY]"
		End Select
	ElseIf Key <> "" and (DEF_BBS_SearchMode = 1) Then
		If Sch_AncTitle = 0 and Mode = 0 Then Mode = 1
		If Mode <> 1 Then SessionKey = Left(key,20)
		Select Case Mode
			'������
			Case 1: CloseDatabase
					Response.Redirect "../User/LookUserInfo.asp?name=" & UrlEncode(key) & "&Evol=n"
			'����������
			Case Else: 
					BoardID2 = Left(Request.QueryString("BoardID2"),14)
					If BoardID2 = "" Then BoardID2 = Left(Request.QueryString("bd"),14)
					If isNumeric(BoardID2) = 0 Then BoardID2 = 0
					BoardID2 = cCur(Fix(BoardID2))
					If BoardID2 > 0 Then
						SQLendString = " where BoardID=" & BoardID2 & " and title like'%" & Replace(key,"'","''") & "%'"
					Else
						SQLendString = " where title like'%" & Replace(key,"'","''") & "%'"
					End If
					WhereFlag = 1
					If BoardID2 > 0 Then
						BoardID2 = "&bd=" & BoardID2
					Else
						BoardID2 = ""
					End If
		End Select
	Else
		If DEF_BBS_SearchMode = 0 Then
			GBL_CHK_TempStr = "��̳������������"
		ElseIf key <> "" Then
			GBL_CHK_TempStr = "�����������û�б������ѯ��Χ��"
		Else
			GBL_CHK_TempStr = "�����������ؼ��ʣ�"
		End If
	End If

	Dim Rs,SQL,Temp,All_Count

	Dim SQLEndString,WhereFlag,OrderColumn,BoardUrlString
	BoardUrlString = "Search.asp"

	Dim RecordCount,ThisPageNum,MaxPage,GetData_2
	RecordCount = -1
	'on error resume next
	Con.CommandTimeout = 120
	If FullTextKey<>"" and DEF_UsedDataBase = 0 and DEF_BBS_SearchMode = 2 Then
		All_Count = Request.QueryString("c")
		If All_Count = "" Then
			SQL = "select count(*) from containstable(LeadBBS_Announce," & SQLEndString & ",'" & Replace(FullTextKey,"'","''") & "'," & DEF_BBS_MaxListPage * (DEF_TopicContentMaxListNum + 1) & ") as T1"
			Set Rs = LDExeCute(SQL,0)
			If Not Rs.Eof Then
				All_Count = cCur("0" & Rs(0))
			Else
				All_Count = 0
			End If
			Rs.Close
			Set Rs = Nothing
		Else
			If isNumeric(All_Count) = 0 then All_Count = 0
			All_Count = Fix(cCur(All_Count))
		End if
		
		MaxPage = Fix(All_Count / DEF_TopicContentMaxListNum)
		If (All_Count mod DEF_MaxListNum) <> 0 Then MaxPage = MaxPage + 1
		MaxPage = MaxPage - 1
		If Page > MaxPage Then Page = MaxPage
	
		If Page < 1 Then
			First = 0
		Else
			First = Page*DEF_TopicContentMaxListNum+1
		End If
		

		SQL = sql_select("select TT.id,TT.ParentID,TT.ChildNum,TT.Layer,TT.Title,TT.FaceIcon,TT.NDateTime,TT.Hits,TT.Length,TT.UserName,TT.UserID,TT.RootIDBak,TT.TopicSortID,TT.LastUser,TT.NotReplay,TT.GoodFlag,TT.BoardID,TT.TopicType,TT.PollNum,TT.TitleStyle,TT.Content from LeadBBS_Announce as TT " & " right join containstable(LeadBBS_Announce," & SQLEndString & ",'" & Replace(FullTextKey,"'","''") & "'," & DEF_BBS_MaxListPage * (DEF_TopicContentMaxListNum + 1) & ") as T1 ON TT.id = T1.[KEY] order by tt.id desc",First + DEF_TopicContentMaxListNum + 1)
	ElseIf Key <> "" and (DEF_BBS_SearchMode = 1 or DEF_BBS_SearchMode = 2) Then
		If Replace(key,"��","") = "" Then
			GBL_CHK_TempStr = "�����������ؼ��ʣ�"
		End If
		SQL = "select id,ParentID,ChildNum,Layer,Title,FaceIcon,NDateTime,Hits,Length,UserName,UserID,RootIDBak,TopicSortID,LastUser,NotReplay,GoodFlag,BoardID,TopicType,PollNum,TitleStyle,Content from LeadBBS_Announce " & SQLEndString
		If First > 0 Then SQL = SQL & " and ID>" & First
		sql = sql_select(sql,DEF_TopicContentMaxListNum * DEF_BBS_MaxListPage)
	End If

	If GBL_CHK_TempStr = "" Then
		If FullTextKey<>"" and DEF_BBS_SearchMode = 2 and All_Count > 0 Then
			Set Rs = LDExeCute(SQL,0)
		Else
			Set Rs = LDExeCute(SQL,0)
		End If
		If err Then
			Rs.Close
			Set Rs = Nothing
			Err.Clear
			ExcuteErr = "�����������û�б������ѯ��Χ����������"
		End if
	End If

	Dim LastSearchTime
	If Session(DEF_MasterCookies & "Schtime") & "" = "" Then Session(DEF_MasterCookies & "Schtime") = GetTimeValue("2000-03-03 00:00:00")
	LastSearchTime = DateDiff("s",RestoreTime(Session(DEF_MasterCookies & "Schtime")),DEF_Now)
	If DEF_UsedDataBase = 0 and DEF_BBS_SearchMode = 2 Then
		If SessionKey <> "" and Session(DEF_MasterCookies & "Sch") & "" = "" Then Session(DEF_MasterCookies & "Sch") = SessionKey
		If LastSearchTime < Sch_LimitTime Then
			If SessionKey <> "" and Session(DEF_MasterCookies & "Sch") <> SessionKey Then
				GBL_CHK_TempStr = "���������ƣ�" & Sch_LimitTime & "����ֻ��������һ�Σ����Ժ����ԡ�"
			Else
				If SessionKey <> "" Then Session(DEF_MasterCookies & "Sch") = SessionKey
			End If
		Else
			Session(DEF_MasterCookies & "Schtime") = GetTimeValue(DEF_Now)
			Session(DEF_MasterCookies & "Sch") = SessionKey
		End If
	Else
		If DEF_BBS_SearchMode = 1 or DEF_BBS_SearchMode = 2 Then
			If LastSearchTime < Sch_LimitTime Then
				GBL_CHK_TempStr = "���������ƣ�" & Sch_LimitTime & "����ֻ��������һ�Σ����Ժ����ԡ�"
			Else
				Session(DEF_MasterCookies & "Schtime") = GetTimeValue(DEF_Now)
			End If
		End If
	End If
			

	LMT_WidthStr = DEF_BBS_ScreenWidth
	If GBL_CHK_TempStr <> "" or ExcuteErr <> "" Then
		If GBL_ShowBottomSure = 0 Then GBL_SiteBottomString = ""
		Global_ErrMsg GBL_CHK_TempStr & ExcuteErr
		Exit Sub
	End If

	If DEF_UsedDataBase = 0 and DEF_BBS_SearchMode = 2 and All_Count = 0 Then
		ThisPageNum = 0
	Else
		If Not Rs.Eof Then
			If RecordCount = -1 Then
				'RecordCount = Rs.RecordCount
				'If RecordCount = -1 Then
					For N = 1 to First
						If Not Rs.Eof Then
							Rs.MoveNext
						Else
							Exit For
						End If
					Next
				'Else
				'	If First > RecordCount Then First = RecordCount - DEF_TopicContentMaxListNum
				'	If First>0 and First <= RecordCount Then Rs.absoluteposition = First
				'End If
			Else
				If First > RecordCount Then First = RecordCount - DEF_TopicContentMaxListNum
				For N = 1 to First
					Rs.MoveNext
				Next
			End If
			GetData_2 = Rs.GetRows(DEF_TopicContentMaxListNum+1)
			Rs.Close
			Set Rs = Nothing
			ThisPageNum = Ubound(GetData_2,2) + 1
		Else
			Rs.Close
			Set Rs = Nothing
			RecordCount = 0
			ThisPageNum = 0
		End If
	End If
	If ThisPageNum > RecordCount Then RecordCount = ThisPageNum

	'If Page > 0 Then
	'	PageSplitString = PageSplitString & "<a href=Search.asp" & SQL & ">��ҳ</a>"
	'Else
	'	PageSplitString = PageSplitString & "<font color=888888 class=grayfont>��ҳ</font>"
	'End If
	'If RecordCount > DEF_TopicContentMaxListNum Then
	'	PageSplitString = PageSplitString & " <a href=Search.asp" & SQL & "&Page=" & Page + 1 & ">��һҳ</a>"
	'End If

	Dim Search_MaxID,Search_MinID
	Dim PageSplitString,PageSplitString2
	
	SQL = "?key=" & urlencode(Request("key"))
	SQL = SQL & "&mode=" & urlencode(Request.QueryString("mode")) & BoardID2
	PageSplitString = "<table border=0 cellspacing=0 cellpadding=0><tr><td><div class=j_page>"

	If DEF_UsedDataBase = 0 and DEF_BBS_SearchMode = 2 Then
		If All_Count > 0 Then SQL = SQL & "&c=" & All_Count
		If Page > 0 and MaxPage > 0 Then
			PageSplitString = PageSplitString & "<a href=Search.asp" & SQL & "&Page=0>1</a>"
		
			if Page <> 1 Then
				PageSplitString = PageSplitString & "<a href=Search.asp" & SQL & "&Page=" & Page-1 & ">��ҳ"
				If (Page - DEF_DisplayJumpPageNum) > 0 Then PageSplitString = PageSplitString & "��"
				PageSplitString = PageSplitString & "</a>"
			End If
		Else
			'PageSplitString = PageSplitString & "��ҳ"
			'PageSplitString = PageSplitString & "��ҳ"
		End If
		
		Dim DN
		DN = DEF_DisplayJumpPageNum
		Dim For1,For2,DotFlag
		DotFlag = 0

		If MaxPage > 0 Then
			For1 = Page - DN
			For2 = Page + DN
			If For1 < 0 Then
				For1 = 0
			End If
			If For2 >= MaxPage Then For2 = MaxPage
			If For2 > MaxPage Then For2 = MaxPage
			If For2 - For1 < DEF_MaxJumpPageNum and For1 > 1 and For2 > For1 Then For1 = For1 - (DEF_MaxJumpPageNum - ((For2 - For1) + 1))
			
			If For1 < 0 Then
				For1 = 0
			ElseIf For1 > 0 Then
				'PageSplitString = PageSplitString & " ..."
			End If
	
			If For2 - For1 < DEF_MaxJumpPageNum and For2 < MaxPage and For2 > For1 Then For2 = For2 + (DEF_MaxJumpPageNum - ((For2 - For1) + 1))
			If For2 >= MaxPage Then For2 = MaxPage
	
			If For2 > 998 Then
				If Page > For1 Then For1 = For1 + 1
				If Page > For1 Then For1 = For1 + 1
				If Page < For2 - DN + 1 Then For2 = For2 - 1
			End If
			For N = For1 to For2
				If N - Page < -5 or N-Page > 5 Then
				Else
					If N = Page Then
						PageSplitString = PageSplitString & "<b>" & N + 1 & "</b></span>"
					Else
						If N <> MaxPage and N <> 0 Then
								PageSplitString = PageSplitString & "<a href=Search.asp" & SQL & "&Page=" & N & " class=j_page>" & N + 1 & "</a>"
						End If
					End If
				End If
				DotFlag = 2
			Next
			If For2 < MaxPage Then
				'PageSplitString = PageSplitString & "..."
				DotFlag = 1
			End If
		Else
			PageSplitString = PageSplitString & " <b>1</b>"
		End If
	
		If Page >= MaxPage Then
			'PageSplitString = PageSplitString & "��ҳ"
			'PageSplitString = PageSplitString & "βҳ"
		Else
			If Page <> MaxPage-1 Then
				PageSplitString = PageSplitString & "<a href=Search.asp" & SQL & "&Page=" & Page + 1 & ">"
				If (Page + DN) < MaxPage Then PageSplitString = PageSplitString & "��"
				PageSplitString = PageSplitString & "��ҳ</a>"
			End If
			PageSplitString = PageSplitString & "<a href=Search.asp" & SQL & "&Page=" & MaxPage & ">" & MaxPage + 1 & "</a>"
		End If
		'PageSplitString = PageSplitString & " ��ҳ<b>" & ThisPageNum & "</b>�� ÿҳ<b>" & DEF_TopicContentMaxListNum + 1 & "</b>��"
		
		If cCur(All_Count) = DEF_BBS_MaxListPage * (DEF_TopicContentMaxListNum + 1) Then
			PageSplitString = PageSplitString & "<b>�������Խ����Χ�������¼�Ѻ���</b>"
		End If
	Else
		PageSplitString = PageSplitString & ""
		Page = fix(Page)
		If RecordCount > DEF_TopicContentMaxListNum Then
			Temp = Page + 2
		Else
			Temp = Page + 1
		End If
		If Temp >= DEF_BBS_MaxListPage Then Temp = DEF_BBS_MaxListPage
		For N = 1 to Temp
				If N = Page+1 Then
					PageSplitString = PageSplitString & " <b>" & N & "</b>"
				Else
					PageSplitString = PageSplitString & " <a href=Search.asp" & SQL & "&Page=" & N-1 & ">" & N & "</a>"
				End If
		Next
		'PageSplitString = PageSplitString & " ҳ"
	
		'PageSplitString = PageSplitString & " ��ҳ<b>" & ThisPageNum & "</b>�� ÿҳ<b>" & DEF_TopicContentMaxListNum + 1 & "</b>��"
	End If
	PageSplitString = PageSplitString & "</div></td></tr></table>"
	If ThisPageNum < DEF_TopicContentMaxListNum and GBL_ShowBottomSure = 0 Then GBL_SiteBottomString = ""

	If isArray(GetData_2) = False Then
		Response.Write "<div class=alert>��ѯ������޷������ݣ�[<a href=Search.asp>����</a>]</div>"
	Else
		Response.Write PageSplitString
		DisplayAnnounceData 0,ThisPageNum-1,1,GetData_2,1
		Response.Write PageSplitString
	End If
	Set Rs = Nothing

End Sub

Sub DisplayAnnounceData(For1,For2,StepValue,GetData,AllFlag)
	
	Dim ReAncStr

	Dim N,Temp,Temp1,Temp2,ForumPass,BoardLimit,OtherLimit,HiddenFlag,BoardName,Temp3,Temp4,Temp5
	
	dim re
	set re = New RegExp
	re.Global = True
	re.IgnoreCase = True
	LMT_WidthStr = "100%"
	For N = For1 to For2 Step StepValue
		Response.Write "<table width=100% border=0 cellspacing=0 cellpadding=0 class=table_in>"
		Response.Write "<tr><td>"
		Response.Write "<hr class=splitline>"
		Response.Write "</td></tr>"
		Response.Write "<tr><td>"

		GetData(8,N) = cCur(GetData(8,N))
		If GetData(8,N) > 1024 Then
			GetData(8,N) = cLng(GetData(8,N)/1024) & " KB"
		Else
			GetData(8,N) = GetData(8,N) & " �ֽ�"
		End If

		If cCur(GetData(1,n)) = 0 Then
			ReAncStr = ""
		Else
			ReAncStr = "&RID=" & GetData(0,N) & "#F" & GetData(0,N)
		End If
		Response.Write "<span class=fontzi><a href=../a/a.asp?B=" & GetData(16,N) & "&ID=" & GetData(11,N) & ReAncStr & " title=""" & GetData(8,N) & """"
		If AllFlag = 1 Then
			Response.Write " target=_blank>"
		Else
			Response.Write ">"
		End If

		GetData(2,N) = cCur(GetData(2,N))
		Temp1 = Fix((GetData(2,N)+1)/DEF_TopicContentMaxListNum)
		If ((GetData(2,N)+1) mod DEF_TopicContentMaxListNum) > 0 Then Temp1 = Temp1 + 1
		If GetData(2,N)>=DEF_TopicContentMaxListNum Then
			Temp = DEF_BBS_DisplayTopicLength + 20 - (Len(Temp1) + 4)
		Else
			Temp = DEF_BBS_DisplayTopicLength + 20 - 1
		End If
		
		'If ccur(GetData(15,n)) = 1 Then Temp = Temp - 3
		GetData(16,N) = cCur(GetData(16,N))
		Temp2 = Application(DEF_MasterCookies & "BoardInfo" & GetData(16,N))
		If isArray(Temp2) = False Then
			ReloadBoardInfo(GetData(16,N))
			Temp2 = Application(DEF_MasterCookies & "BoardInfo" & GetData(16,N))
		End If
		If isArray(Temp2) = False Then
			'Response.Write "������̳��������,����ϵ����Ա!<br>" & VbCrLf
			ForumPass = "a"
			BoardLimit = 0
			OtherLimit = 0
			HiddenFlag = 0
			BoardName = ""
		Else
			ForumPass = Temp2(7,0)
			BoardLimit = cCur(Temp2(9,0))
			OtherLimit = cCur(Temp2(36,0))
			HiddenFlag = Temp2(8,0)
			BoardName = Temp2(0,0)
		End If
		
		
		'If GBL_NoneLimitFlag = 0 and (ForumPass <> "" or GetBinarybit(BoardLimit,7) = 1 or GetBinarybit(BoardLimit,2) = 1 or GetBinarybit(BoardLimit,15) = 1 or OtherLimit > 0) Then
		'	GetData(4,n) = "<font color=gray calss=grayfont>�����ӱ���������Ϊ����</font>"
		'	GetData(19,n) = 1
		'	GetData(20,n) = "<font color=gray calss=grayfont>�����������������ư��棬��������鿴...</font>"
		'End If

		If GBL_NoneLimitFlag = 0 and GBL_CheckLimitTitle(ForumPass,BoardLimit,OtherLimit,HiddenFlag) = 1 Then
			GetData(4,n) = "<span calss=grayfont>�����ӱ���������Ϊ����</span>"
			GetData(19,n) = 1
		End If
		If GBL_CheckLimitContent(ForumPass,BoardLimit,OtherLimit,HiddenFlag) = 1 Then GetData(20,n) = "<span calss=grayfont>�����������������ư��棬��������鿴</span>"

		If GetData(17,n) <> 80 and GetData(17,n) <> 0 Then GetData(20,n) = "<span calss=grayfont>�����������������ƣ���������鿴...</span>"

		If left(GetData(4,N),3) = "re:" and GetData(4,N) <> "re:" Then GetData(4,N) = Mid(GetData(4,N),4)
		If GetData(19,n) <> 1 and strLength(GetData(4,N))>Temp Then GetData(4,N) = LeftTrue(GetData(4,N),Temp-3) & "..."
		
		If GetData(19,n) <> 1 Then GetData(4,n) = htmlEncode(GetData(4,n)) '��html����

		If isArray(LMT_Key) Then
			Temp4 = LMT_Key(0)
		Else
			Temp4 = LMT_Key
		End If

		Temp5 = inStr(1,GetData(20,N),Temp4,0)
		If Temp5 > 0 Then
			Temp2 = inStrRev(GetData(20,N),">",Temp5,0)
		Else
			Temp2 = 0
		End If
		If Temp2 > 0 Then
			If Temp5 - Temp2 > DEF_BBS_MaxWords/2 Then
				Temp3 = Temp5 - DEF_BBS_MaxWords/2-1
			Else
				Temp3 = Temp2 + 1
			End If
		Else
			Temp3 = Temp5 - DEF_BBS_MaxWords/2
			If Temp3 < 1 Then Temp3 = 1
		End If
		GetData(20,N) = Mid(GetData(20,N),Temp3,DEF_BBS_MaxWords)
		GetData(20,N) = HtmlEncode(KillHTMLLabel(GetData(20,N)))
		
		If isArray(LMT_Key) Then
			Temp4 = LMT_Key(0)
			For Temp3 = 0 to Ubound(LMT_Key)
				If Len(LMT_Key(Temp3)) > 0 Then
					re.Pattern="(" & ConverStr(LMT_Key(Temp3)) & ")"
					GetData(4,n)=re.Replace(GetData(4,n),"<span class=redfont>$1</span>")
					GetData(20,n)=re.Replace(GetData(20,n),"<span class=redfont>$1</span>")
				End If
			Next
		Else
			Temp4 = LMT_Key
			If Len(Temp4) > 0 Then
				re.Pattern="(" & ConverStr(Temp4) & ")"
				GetData(4,n)=re.Replace(GetData(4,n),"<span class=redfont>$1</span>")
				GetData(20,n)=re.Replace(GetData(20,n),"<span class=redfont>$1</span>")
			End If
		End If
		Response.Write GetData(4,n)
		Response.Write "</a></span>"

		If GetData(2,N)>=DEF_TopicContentMaxListNum Then
			Response.Write " [<a href=../a/a.asp?B=" & GetData(16,N) & "&ID=" & GetData(0,N) & "&AUpflag=1&ANum=1>" & Temp1 & "</b></a>]"
		End If

		If ccur(GetData(15,n)) = 1 Then
			Response.Write "<img src=../images/" & GBL_DefineImage & "jh1.GIF border=0 title=�������� align=absbottom>"
		End If

		If BoardName <> "" Then Response.Write " <span class=grayfont>-</span> <a href=../b/b.asp?b=" & GetData(16,N) & "><span class=greenfont>" & BoardName & "</span></a>"
		Response.Write "<br><img src=../images/null.gif width=2 height=5><br>"

		Response.Write GetData(20,N)
		GetData(10,N) = cCur(GetData(10,N))
		Response.Write "<div class=value2><span class=grayfont>���ߣ�"
		If GetData(10,N) > 0 Then
			Response.Write "<a href=../User/LookUserInfo.asp?ID=" & GetData(10,N) & "><span class=greenfont>" & htmlencode(GetData(9,N)) & "</span></a>"
		Else
			Response.Write htmlencode(GetData(9,N))
		End If

		'If isNull(GetData(18,N)) Then GetData(18,N) = 0
		'If GetData(17,n) = 80 Then
		'	Response.Write "��" & cCur(GetData(18,N)) & "Ʊ"
		'Else
		'	Response.Write GetData(2,N) & "/" & GetData(7,N)
		'End If
		Temp = RestoreTime(GetData(6,N))
		If DateDiff("d",Temp,DEF_Now)<1 Then
			Response.Write " ������ <span class=redfont>" & Temp & "</span>"
		Else
			Response.Write " ������ " & Temp
			Response.Write " " & DateDiff("d",Temp,DEF_Now) & "��ǰ"
		End If
		Response.Write "</span></div></td></tr>"
		If N = For2 Then Response.Write "<tr><td><hr class=splitline></td></tr>"
		Response.Write "</table>" & VbCrLf
	Next

End Sub

Function ConverStr(s)

	ConverStr = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(s,"*","\*"),"\","\\"),"?","\?"),"%","\%"),"^","\^"),")","\)"),"(","\("),"]","\]"),"[","\["),"+","\+")
	ConverStr = Replace(Replace(Replace(ConverStr,"{","\{"),"}","\}"),".","\.")

End Function
%>