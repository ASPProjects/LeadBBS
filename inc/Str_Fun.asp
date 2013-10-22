<%
Dim BinaryData,BinaryDataNum
BinaryData = Array(1,2,4,8,16,32,64,128,256,512,1024,2048,4096,8192,16384,32768,65536,131072,262144,524288,1048576,2097152,4194304,8388608,16777216,33554432,67108864,134217728,268435456,536870912,1073741824,2147483648)
BinaryDataNum = 32

Function GetBinarybit(Number,bit)

	if isNull(Number) Then
		GetBinarybit = 0
		Exit Function
	Else
		Number = cCur(Number)
	End If
	If bit = BinaryDataNum Then
		If Number = BinaryData(bit) Then
			GetBinarybit = 1
		Else
			GetBinarybit = 0
		End If
	Else
		If (cCur(Number) mod BinaryData(bit)) >= BinaryData(bit-1) Then
			GetBinarybit = 1
		Else
			GetBinarybit = 0
		End If
	End if

End Function

Function CodeCookie(str)

	Dim i
	Dim StrRtn
	For i = Len(Str) to 1 Step -1
		StrRtn = StrRtn & Ascw(Mid(Str,i,1))
		If (i <> 1) Then StrRtn = StrRtn & "a"
	Next
	CodeCookie = StrRtn

End Function

Function DecodeCookie(Str)

	Dim i
	Dim StrArr,StrRtn
	StrArr = Split(Str,"a")
	For i = UBound(StrArr) - LBound(StrArr) to 0 Step -1
		If isNumeric(StrArr(i)) = True Then
			StrRtn = StrRtn & Chrw(StrArr(i))
		Else
			StrRtn = Str
			Exit Function
		End If
	Next
	DecodeCookie = StrRtn

End Function

Function RestoreTime(DateString)

	If isNull(DateString) Then Exit Function
	DateString = cstr(DateString)
	If len(DateString)<8 then
		RestoreTime=DateString
	Else
		If len(DateString)<14 then
			RestoreTime = Mid(DateString,1,4) & "-" & Mid(DateString,5,2) & "-" & Mid(DateString,7,2)
		Else
			RestoreTime = Mid(DateString,1,4) & "-" & Mid(DateString,5,2) & "-" & Mid(DateString,7,2) & " " & Mid(DateString,9,2) & ":" & Mid(DateString,11,2) & ":" & Mid(DateString,13,2)
		End If
	End If

End Function

Function StrLength(str)

	If isNull(str) or Str = "" Then
		StrLength = 0
		Exit function
	End If
	If len("例子") = 2 then
		Dim l,t,c,i
		l=len(str)
		t=l
		for i=1 to l
			c=asc(mid(str,i,1))
			If c<0 then c=c+65536
			If c>255 then
				t=t+1
			End If
		next
		StrLength=t
	Else 
		StrLength=len(str)
	End If
End Function

Function GetTimeValue(DateString)

	Dim Temp,TempStr
	If isNull(DateString) or isTrueDate(DateString) = 0 Then
		GetTimeValue = 0
		Exit Function
	End If
	Temp = CsTr(Year(DateString))
	If Len(temp)<3 Then
		Temp = Left(year(DEF_Now),2) & Temp
	End If
	TempStr = Temp
	
	Temp = CsTr(month(DateString))
	If Len(temp)<2 Then Temp = "0" & Temp
	TempStr = TempStr & Temp

	Temp = CsTr(day(DateString))
	If Len(Temp) < 2 Then Temp = "0" & Temp
	TempStr = TempStr & Temp

	Temp = csTr(Hour(DateString))
	If Len(Temp) < 2 Then Temp = "0" & Temp
	TempStr = TempStr & Temp

	Temp = CsTr(Minute(DateString))
	If Len(Temp) < 2 Then Temp = "0" & Temp
	TempStr = TempStr & Temp

	Temp = CsTr(Second(DateString))
	If Len(Temp) < 2 Then Temp = "0" & Temp
	TempStr = TempStr & Temp

	GetTimeValue = cCur(TempStr)

End Function

Function htmlEncode(str)

	If str & "" <> "" Then
		htmlEncode=Replace(Replace(Replace(str,">","&gt;"),"<","&lt;"),"""","&quot;")
	Else
		htmlEncode=str
	End If

End Function

Function UrlEncode(str)

	If str & "" <> "" Then
		urlencode = Server.UrlEncode(str)
	Else
		UrlEncode = str
	End If

End Function



rem 显示左边的n个字符(自动识别汉字)
Function LeftTrue(str,n)

	If len(str) <= n/2 Then
		LeftTrue = str
	Else
		Dim TStr,l,t,c,i
		l = len(str)
		TStr = ""
		t = 0
		For i=1 To l
			c = asc(mid(str,i,1))
			If c < 0 then c=c+65536
			If c > 255 then
				t = t+2
			Else
				t = t+1
			End If
			If t > n Then exit for
			TStr = TStr&(mid(str,i,1))
		Next
		LeftTrue = TStr
	End If

End Function

Function isTrueDate(TStr)

	Dim T
	T = TStr
	If isNull(T) Then T = ""
	T = Replace(Replace(Replace(Replace(Replace(Replace(Replace(T,"年","-"),"月","-"),"日"," "),"上午"," "),"下午"," "),"  "," "),"  "," ")
	
	Dim N1,N2
	N1 = inStr(T,"-")
	If N1>0 Then N2 = inStrRev(T,"-")
	If N1 = N2 and N1 >0 Then
		isTrueDate = 0
		Exit Function
	End If

	N1 = inStr(T,":")
	If N1>0 Then N2 = inStrRev(T,"-")
	If N1 = N2 and N1 >0 Then
		isTrueDate = 0
		Exit Function
	End If

	If isDate(TStr) Then
		isTrueDate = 1
	Else
		isTrueDate = 0
	End If

End Function



Function KillHTMLLabel(str)

	Dim n,m,str2
	m = 0
	n = inStr(str,"<")
	if n > 0 Then m = inStr(n,str,">")
	str2 = str
	Do while n > 0 and n < m
		str2 = Left(str2,n-1) & Mid(str2,m+1)
		n = inStr(str2,"<")
		if n > 0 Then m = inStr(n,str2,">")
	Loop
	KillHTMLLabel = str2

End Function

Function LeftTrueHTML(str,Ln)

	Dim n,m,str2,str3,htm,tmp,flag,tmp2,tmp3
	str3 = ""
	htm = ""
	tmp = ""
	flag = 0
	tmp2 = ""
	tmp3 = ""
	n = inStr(Str,"<")
	m = inStr(Str,">")
	str2 = str
	Dim s
	s = 0
	do while(n >= 1 and n < m)
		s=s+1
		if s>100 then exit do
		tmp = Mid(str2,1,n-1)
		If flag = 0 Then
			If StrLength(str3 & tmp) > Ln Then
				flag = 1
				tmp2 = LeftTrue(tmp,Ln-strlength(str3))
				tmp2 = tmp2 & "..."
			Else
				tmp2 = tmp
				str3 = str3 & tmp
			End If
		Else
			tmp2 = ""
		End If
		If flag = 0 Then
			htm = htm & tmp & Mid(str2,n,m-n+1)
		Else
			htm = htm & tmp2 & Mid(str2,n,m-n+1)
		End If
		tmp3 = Mid(str2,m+1)
		str2 = tmp3
		n = inStr(Str2,"<")
		m = inStr(Str2,">")
	Loop
	
	If flag = 0 Then
		If strlength(str3 & tmp3)>Ln Then
			flag = 1
			tmp2 = LeftTrue(tmp3,Ln-strlength(str3))
			tmp2 = tmp2 & "..."
		Else
			tmp2 = tmp3
		End If
	Else
		tmp2 = ""
	End If
	htm = htm + tmp2
	LeftTrueHTML = htm

End Function

Function ADODB_LoadFile(ByVal File)

	On Error Resume Next
	Dim objStream,FSFlag,fs,WriteFile
	FSFlag = 1
	If DEF_FSOString <> "" Then
		Set fs = Server.CreateObject(DEF_FSOString)
		If Err Then
			FSFlag = 0
			Err.Clear
			Set fs = Nothing
		End If
	Else
		FSFlag = 0
	End If

	If FSFlag = 1 Then
		Set WriteFile = fs.OpenTextFile(Server.MapPath(File),1,True)
		If Err Then
			GBL_CHK_TempStr = "<br>读取文件失败：" & err.description & "<br>其它可能：确定是否对此文件有读取权限."
			err.Clear
			Set Fs = Nothing
			Exit Function
		End If
		If Not WriteFile.AtEndOfStream Then
			ADODB_LoadFile = WriteFile.ReadAll
			If Err Then
				GBL_CHK_TempStr = "读取文件失败：<p>" & err.description & "</p> 其它可能：确定是否对此文件有读取权限."
				err.Clear
				Set Fs = Nothing
				Exit Function
			End If
		End If
		WriteFile.Close
		Set Fs = Nothing
	Else
		Set objStream = Server.CreateObject("ADODB.Stream")
		If Err.Number=-2147221005 Then 
			GBL_CHK_TempStr = "您的主机不支持ADODB.Stream，无法完成操作，请手工进行"
			Err.Clear
			Set objStream = Nothing
			Exit Function
		End If
		With objStream
			.Type = 2
			.Mode = 3
			.Open
			.LoadFromFile Server.MapPath(File)
			.Charset = "gb2312"
			.Position = 2
			ADODB_LoadFile = .ReadText
			.Close
		End With
		Set objStream = Nothing
	End If
	If Err Then
		GBL_CHK_TempStr = "错误信息：<p>" & err.description & "</p>其它可能：确定是否对此文件有读取权限."
		err.Clear
		Set Fs = Nothing
		Exit Function
	End If

End Function

'存储内容到文件
Sub ADODB_SaveToFile(ByVal strBody,ByVal File)

	'On Error Resume Next
	Dim objStream,FSFlag,fs,WriteFile
	FSFlag = 1
	If DEF_FSOString <> "" Then
		Set fs = Server.CreateObject(DEF_FSOString)
		If Err Then
			FSFlag = 0
			Err.Clear
			Set fs = Nothing
		End If
	Else
		FSFlag = 0
	End If
	If FSFlag = 1 Then
		Set WriteFile = fs.CreateTextFile(Server.MapPath(File),True)
		WriteFile.Write strBody
		WriteFile.Close
		Set Fs = Nothing
	Else
		Set objStream = Server.CreateObject("ADODB.Stream")
		If Err.Number=-2147221005 Then 
			GBL_CHK_TempStr = "您的主机不支持ADODB.Stream，无法完成操作，请手工进行"
			Err.Clear
			Set objStream = Nothing
			Exit Sub
		End If
		With objStream
			.Type = 2
			.Open
			.Charset = "gb2312"
			.Position = objStream.Size
			.WriteText = strBody
			.SaveToFile Server.MapPath(File),2
			.Close
		End With
		Set objStream = Nothing
	End If
	If Err Then
		GBL_CHK_TempStr = "错误信息：<p>" & err.description & "</p>其它可能：确定是否对此文件有写入权限."
		err.Clear
		Set Fs = Nothing
		Exit Sub
	End If

End Sub

Function GetSBInfo(Flag)

	Dim Brs,Sys,I,N,Tmp,Str
	Sys = "Unknown"
	Brs = "Unknown"
	Str = Request.ServerVariables("HTTP_USER_AGENT")
	Tmp = LCase(Str)
	'If inStr(Tmp,"http://") > 0 Then
	'	Brs = "Spider"
	'	Sys = "Spider"
	'Else
		I = inStr(Tmp,"msie")
		If I > 0 Then
			N = inStr(I,Tmp,";")
			If N > 0 Then
				Brs = Mid(Str,I,N-i)
				I = inStr(N+1,Tmp,";")
				If I > 0 Then
					Sys = Trim(Mid(Str,N + 1,I-N-1))
				End If
			End If
		Else
			I = inStr(Tmp,"opera")
			If I > 0 Then
				N = inStr(i,Tmp," ")
				If N > 0 Then Brs = Replace(Mid(Str,i,n-i),"/"," ")
				I = inStr(Tmp,"(")
				N = inStr(Tmp,";")
				If N > I and I > 0 Then
					Sys = Mid(Str,I+1,N-I-1)
				End If
			ElseIf inStr(Tmp,"safari") > 0 Then
				I = inStr(Tmp,"version")
				If I > 0 Then
					If inStr(i,Tmp," ")-I-7 > 0 Then Brs = "Safari " & Replace(Mid(Tmp,I + 7,inStr(I,Tmp," ")-I-7),"/","")
				Else
					I = inStr(Tmp,"chrome")
					If I > 0 Then
						If inStr(I,Tmp," ") > I Then
							Brs = Replace(Mid(Tmp,I,inStr(I,Tmp," ")-I),"/"," ")
						End If
					End If
				End If
			ElseIf inStr(Tmp,"wap") > 0 Then
				Brs = "Wap"
				Sys = "Wap"
			Else
				If inStr(Tmp,";")>0 then
					Dim T
					N = split(Str,";")
					
					I = inStr(Tmp,"firefox")
					If I > 0 and Ubound(N) >=2 Then
						Sys = Trim(N(2))
						Brs = Replace(Mid(Str,I,20),"/"," ")
					Else
						If Ubound(N) >=2 Then
							N(2) = Trim(replace(N(2),")",""))
							Brs = Replace(N(2),"/"," ")
						End If
						If Ubound(N) >=1 Then
							N(1) = Trim(N(1))
							Sys = N(1)
						End If
					End If
				End If
			End If
		End If
	'End If
	If Brs = "Unknown" and inStr(Tmp,"http://") > 0 Then Brs = "Spider"
	If Sys <> "" Then
		If inStr(Str,"Windows NT 5.0") Then
			Sys = "Windows 2000" 
		Elseif inStr(Str,"Windows NT 5.1") Then
			Sys = "Windows XP" 
		Elseif inStr(Str,"Windows NT 5.2") Then
			Sys = "Windows 2003"
		Elseif inStr(Str,"Windows NT 6.0") Then
			Sys = "Windows Vista" 
		Elseif inStr(Str,"Windows NT 6.1") Then
			Sys = "Windows 7" 
		Elseif inStr(Str,"Windows NT 6.2") Then
			Sys = "Windows 8" 
		Elseif inStr(Str,"Windows vista") Then
			Sys = "Windows Vista" 
		Elseif inStr(Str,"Windows 4.10") Then
			Sys = "Windows 98" 
		Elseif inStr(Str,"Windows 98") Then
			Sys = "Windows 98" 
		Elseif inStr(Str,"Windows me") Then
			Sys = "Windows Me" 
		Elseif inStr(Str,"ipad") Then
			Sys = "iPad" 
		Elseif inStr(Str,"Windows 3.") Then
			Sys = "Windows 3.1" 
		elseif inStr(Tmp,"mac os x") Then	
			I = inStr(Tmp,"mac os")	
			N = inStr(I,Tmp,";")
			If N > 0 Then
				Sys = Mid(Str,I,N-i)
				Sys = Replace(Replace(Sys,"_","."),";","")
			Else
				Sys = "Mac OS" 
			End If
		elseif inStr(Tmp,"android") Then	
			I = inStr(Tmp,"android")	
			N = inStr(I,Tmp,";")
			If N > 0 Then
				Sys = Mid(Str,I,N-i)
				Sys = Replace(Replace(Sys,"_","."),";","")
			Else
			Sys = "Android" 
			End If
		End If		
	End If
	
	If Flag = 1 Then
		GetSBInfo = Brs
	Else
		GetSBInfo = Sys
	End If

End Function

Function ConvertTimeString(t)

	Dim Tmp,M
	M = Datediff("n",t,DEF_Now)
	If M > 2880 Then
	ElseIf M > 720 Then
		Select Case Datediff("d",t,DEF_Now)
			Case 0: Tmp = "今天 " & Mid(t,12,5)
			Case 1: Tmp = "昨天 " & Mid(t,12,5)
			Case 2: Tmp = "前天 " & Mid(t,12,5)
			Case Else: Tmp = t
		End Select
	ElseIf M >= 60 Then
		Dim M1
		M1 = M mod 60
		If M1 = 0 Then
			Tmp = Fix(M/60) & "小时前"
		Else
			Tmp = Fix(M/60) & "小时" & M1 & "分前"
		End If
	ElseIf M >= 1 Then
		Tmp = M & "分前"
	Else
		M = Datediff("s",t,DEF_Now)
		If M >= 0 Then Tmp = M & "秒前"
	End If

	If Tmp = "" Then Tmp = t		
	ConvertTimeString = Tmp

End Function

Function toNum(s,default)

	if isNumeric(s) = 0 Then
		toNum = default
	else
		toNum = ccur(s)
	end if

End Function
%>