<!-- #include file=../inc/BBSsetup.asp -->
<!-- #include file=../inc/Board_popfun.asp -->
<!-- #include file=inc/UserTopic.asp -->
<%
DEF_BBS_HomeUrl = "../"
GBL_CHK_PWdFlag = 0

Class SelStyle_Class

Private AjaxFlag

Private Sub Class_Initialize

	Dim af
	af = Request("AjaxFlag")
	If af <> "" and af <> "0" Then
		AjaxFlag = 1
	Else
		AjaxFlag = 0
	End If

End Sub

Public Sub Main_Style

	OpenDatabase
	GBL_CHK_TempStr = ""

	Dim HomeUrl,u
	HomeUrl = "http://"&Request.ServerVariables("server_name")
	u = filterUrlstr(Request("u"))
	
	If Left(u,1) <> "/" and Left(u,1) <> "\" and Left(u,Len(HomeUrl)) <> HomeUrl and LCase(Left(u,9)) <> "frame.asp" and LCase(Left(u,12)) <> "../frame.asp" Then u = ""
	
	If LCase(Left(u,12)) = "frame.asp?u=" Then
		u = DEF_BBS_HomeUrl & "Frame.asp?u=" & UrlEncode(Mid(u,13))
	Else
		If inStr(u,"rnd=") Then
			u = Replace(u,"rnd=" & Mid(u,inStr(u,"rnd=")+4,2),"rnd=" & (Fix(Rnd*90) + 10))
		End If
		If u <> "" and inStr(u,"rnd=") = 0 Then
			If inStr(u,"?") = 0 Then
				u = u & "?rnd=00" & Fix(Rnd*1314)
			Else
				u = u & "&rnd=00" & Fix(Rnd*1314)
			End If
		End If
	End If
	
	If u = "" Then
		u = filterUrlstr(Lcase(Request.ServerVariables("HTTP_REFERER")))
	
		If Request.ServerVariables("SERVER_PORT") <> "80" Then HomeUrl = HomeUrl & ":" & Request.ServerVariables("SERVER_PORT")
	
		If Left(u,Len(HomeUrl)) <> Lcase(HomeUrl) Then u = ""
		If inStr(u,"/user/boardstyle.asp") > 0 Then u = ""
	End If
	
	If AjaxFlag = 0 Then BBS_SiteHead DEF_SiteNameString & " - ѡ����",0,"<span class=navigate_string_step>ѡ����</span>"
	dim jsflag
	jsflag = left(request("jsflag"),1)
	If AjaxFlag = 0 Then
		Boards_Body_Head("")
		%>
		<div class='alertbox fire'>
		<%
	Else
		if jsflag = "" then
		%>
		<div class="ajaxbox" id="stylelist">
	<%
		end if
	end if
	If Request("b") <> "" and Request("s") <> "" Then
		SetBoardStyle(u)
	Else
		DisplayBoardStyleList(u)
	End If
	closeDataBase
	if jsflag = "" then
	%>
	</div>
	<%
	end if
	If AjaxFlag = 0 Then
		Boards_Body_Bottom
		sitebottom
	End If

End Sub

Private Function DisplayBoardStyleList(u)

	Dim N,SetBoardID
	SetBoardID = Left(Request("b"),14)
	
	Dim Temp
	Temp = Application(DEF_MasterCookies & "BoardInfo" & SetBoardID)
	If isArray(Temp) = True Then
		If Temp(24,0) <> "" or Temp(25,0) <> "" Then
			Response.Write "<p><b><font color=red class=redfont>�˰��治��������Ϊ������������</font></b>"
			
			If u <> "" Then Response.Write "<p>-- ������ҳ<a href=" & u & ">" & u & "</a>"
			Exit Function
		End If
	End If
	
	If Request("action") <> "extended" then
	%>
	<div class=title>�����¼��ַ��ɹ�ѡ��</div>
	<ul>
	<%
	for N = 0 to DEF_BoardStyleStringNum
	%>
		<li><a href="BoardStyle.asp?b=<%=htmlencode(SetBoardID)%>&s=<%=N%>&u=<%=urlencode(u)%>&AjaxFlag=<%=AjaxFlag%>" target=_top><%=DEF_BoardStyleString(N)%></a><br>
	<%
	Next
	%>
	</ul>
	<%
	else
		
		Dim Url
		Url = filterUrlstr(Left(Request("dir"),100))
		If Url = "" and inStr(Request.QueryString,"dir=")=0 Then
			Url = DEF_BBS_HomeUrl
		End If
		dim rs,sql,num,getdata,startid
		
		dim extendcount,p,listNum
		extendcount = 0
		listNum = 4
		sql = "select count(*) from LeadBBS_Skin where styleid>=1000"
		Set rs = LDExeCute(sql,0)
		if not rs.eof then
			extendcount = rs(0)
			if isnumeric(extendcount) = false then extendcount = 0
			extendcount = ccur(extendcount)
		end if
		rs.close
		set rs = nothing
		
		dim maxpage
		maxpage = Fix(extendcount/listNum)
		if (extendcount mod listNum) > 0 then maxpage = maxpage + 1
		p = request("page")
		if isnumeric(p) = false then p = 0
		p = clng(ccur(p))
		if p >= maxpage then p = maxpage - 1
		if p < 0 then p = 0

		startid = 1000
		dim selpage
		selpage = p
		
		dim movenum
		if p > 100 then
			select case def_useddatabase
				case 0,1:
					sql = sql_select("selectStyleID from LeadBBS_Skin where styleid>=" & startid & " order by styleid asc",(p+1)*listNum)
					Set rs = LDExeCute(sql,0)
					If Not Rs.Eof Then
						Rs.Move p*listNum
						If Not Rs.Eof Then
							startid = Rs(0)
							selpage = 0
						end if
					end if
				case 2:
					movenum = p*listNum
					sql = sql_select("selectStyleID from LeadBBS_Skin where styleid>=" & startid & " order by styleid asc",movenum & "," & (p+1)*listNum)
					Set rs = LDExeCute(sql,0)
					If Not Rs.Eof Then
						startid = Rs(0)
						selpage = 0
					end if
			end select
			rs.close
			set rs = nothing
		end if
		
		select case def_useddatabase
		case 0,1:
			sql = sql_select("select StyleID,ScreenWidth,SmallTableBottom from LeadBBS_Skin where styleid>=" & startid & " order by styleid asc",(selpage+1)*listNum)
			Set rs = LDExeCute(sql,0)
			if not rs.eof then
				if p<=100 and p>0 then
					rs.move p*listNum
				end if
				getdata = rs.getrows(-1)
				num = ubound(getdata,2)
			else
				num = -1
			end if
		case 2:
			movenum = 0
			
			if p<=100 and p>0 then
				movenum = p*listNum
			end if
			sql = sql_select("select StyleID,ScreenWidth,SmallTableBottom from LeadBBS_Skin where styleid>=" & startid & " order by styleid asc",movenum & "," & (selpage+1)*listNum)
			Set rs = LDExeCute(sql,0)
			if not rs.eof then
				getdata = rs.getrows(-1)
				num = ubound(getdata,2)
			else
				num = -1
			end if
		end select
		rs.close
		set rs = nothing
		Dim cleanStyle
		cleanStyle = 0
		if GBL_SiteHeadString = "" and GBL_SiteBottomString = "" and GBL_TableHeadString = "" and GBL_TableBottomString = "" then cleanStyle = 1
		%>
		<div class=title>�������·��ɹ�ѡ��</div>
		<ul class="skinlist">
		<%
		for N = 0 to num
		%>
			<li><a href="<%=Url%>User/BoardStyle.asp?b=<%=htmlencode(SetBoardID)%>&s=<%=getdata(0,n)%>&u=<%=urlencode(u)%>&AjaxFlag=<%=AjaxFlag%>&dir=<%=Url%>&SureFlag=E72ksiOkw2&action=extended"<%
			if cleanStyle = 1 then
				response.write " onclick=""setStyle('" & htmlencode(Url) & "inc/css/"
				If cCur(getdata(0,n)) < 10000 Then Response.Write "0"
				Response.Write getdata(0,n)
				Response.Write ".css','cssfile');LD.Cookie.Add('" & DEF_MasterCookies & "style','border=" & getdata(0,n) & "',1);return false;"""
			end if
			%> target="_top">
			<%=getdata(1,n)%></a></li>
		<%
		Next
		%>
		</ul>
		<div class="clear"></div> 
		<%
		CALL viewextendstylepage(Url & "User/BoardStyle.asp?b=" & htmlencode(SetBoardID) & "&u=" & urlencode(u) & "&dir=" & Url & "&action=extended",maxpage,p,"stylelist")
	end if

End Function

Private function viewextendstylepage(url,num,curp,ajaxobj)

	dim n
	%>
	<div class="j_page" style="float:left;">
	<%
	if curp-4 > 0 then Response.Write "<b>...</b>"
	for n = curp-4 to curp+4
		if n >=0 and n < num then
			if n <> curp then
	%><a href="<%=url%>&page=<%=n%>"<%
		if ajaxobj <> "" Then
			%> onclick="getAJAX(this.href+'&AjaxFlag=1&jsflag=1','','<%=ajaxobj%>');return(false);"<%
		end if
	%>><%=n+1%></a><%
			else
	%><b><%=n+1%></b><%
			end if
		end if
	next
	if curp+4 < num-1 then Response.Write "<b>...</b>"
	%>
	</div>
	<%

end function

Private Function SetBoardStyle(u)

	Dim SetBoardID,BoardStyle

	SetBoardID = Left(Request("b"),14)
	If isNumeric(SetBoardID)=0 Then SetBoardID=0
	SetBoardID = cCur(SetBoardID)
	
	BoardStyle = Left(Request("s"),14)
	If isNumeric(BoardStyle)=0 Then BoardStyle=0
	BoardStyle = Fix(cCur(BoardStyle))
		
	If BoardStyle < 0 or (BoardStyle > DEF_BoardStyleStringNum and BoardStyle < 1000) Then BoardStyle = 0

	If Request("SureFlag") <> "E72ksiOkw2" and BoardStyle < 1000 Then
		%>
		<script language=javascript>
		<!--
		var ValidationPassed = true;
		function submitonce(theform)
		{
			if (document.all||document.getElementById)
			{
				for (i=0;i<theform.length;i++)
				{
					var tempobj=theform.elements[i];
					if(tempobj.type.toLowerCase()=="submit"||tempobj.type.toLowerCase()=="reset")
					tempobj.disabled=true;
				}
			}
		}
		//-->
		</script>
			<form action=BoardStyle.asp method=post onSubmit="submitonce(this);return ValidationPassed;">
			<div class=title>��ȷ�����õ�ǰ���Ϊ:  <%=DEF_BoardStyleString(BoardStyle)%></div>
			<div class=value2>
			<img src=<%=DEF_BBS_HomeUrl%>images/skin/<%=BoardStyle%>/preview.jpg>
			</div>
			<br>
			<div class=value2>
			<input type=hidden name=SureFlag value="E72ksiOkw2">
			<input type=hidden name=b value="<%=SetBoardID%>">
			<input type=hidden name=s value="<%=BoardStyle%>">
			<input type=hidden name=u value="<%=HtmlEncode(u)%>">
			<input type=submit value=ȷ������ class="fmbtn btn_3">
			</div>
			</form>
		<%
	Else
		If SetBoardID < 1 Then
			Response.Cookies(DEF_MasterCookies & "style").Expires = Date + 365
			'��Զ�����治ͬ������� Response.Cookies(DEF_MasterCookies & "style")("border0") = BoardStyle
			Response.Cookies(DEF_MasterCookies & "style")("border") = BoardStyle
			Response.Cookies(DEF_MasterCookies & "style").Domain = DEF_AbsolutHome
			If AjaxFlag = 0 or BoardStyle >= 1000 Then
				If u <> "" or BoardStyle >= 1000 Then Response.Redirect u
				Response.Write "<br>"
				If u <> "" Then Response.Write "<br>-- ������ҳ<a href=" & u & ">" & u & "</a>"
				Response.Write "<br>-- <a href=../Boards.asp>������ҳ</a>"
			End If
			Exit Function
		End If

		Dim Rs,BoardName
		Set Rs = LDExeCute(sql_select("Select BoardName,BoardStyle from LeadBBS_Boards where BoardID=" & SetBoardID,1),0)
		If Not Rs.Eof Then
			Response.Cookies(DEF_MasterCookies & "style").Expires = Date + 365
			'��Զ�����治ͬ������� Response.Cookies(DEF_MasterCookies & "style")("border" & SetBoardID) = cCur(BoardStyle)
			Response.Cookies(DEF_MasterCookies & "style")("border") = cCur(BoardStyle)
			Response.Cookies(DEF_MasterCookies & "style").Domain = DEF_AbsolutHome
			BoardName = Rs(0)
		End If
		Rs.Close
		Set Rs = Nothing
		If AjaxFlag = 0 or BoardStyle >= 1000 Then
			If u <> "" or BoardStyle >= 1000 Then Response.Redirect u
			Response.write "<br>"
			If u <> "" Then Response.Write "<br>-- ������ҳ<a href=" & u & ">" & u & "</a>"
			Response.Write "<br>-- ���ذ���<a href=../b/b.asp?B=" & SetBoardID & ">" & BoardName & "</a>"
			Response.Write "<br>-- <a href=../Boards.asp>������ҳ</a>"
		End If
	End If

End Function

Private Sub Processor_styleMsg(str,obj,evl)

	If AjaxFlag = 0 Then
		Response.Write str
	Else
		If AjaxFlag = 1 and Request("JsFlag")="1" Then%>
		<script>parent.layer_outmsg("<%=obj%>","<span class=\"redfont\"><%=Replace(Replace(Replace(Str,"\","\\"),"""","\"""),VbCrLf,"\n")%></span>","","<%=Replace(Replace(Replace(evl,"\","\\"),"""","\"""),VbCrLf,"\n")%>");</script>
		<%
		Else%>
		<span class="redfont">
			<%=Str%>
		</span>
	<%	End If
	End If

End Sub

End Class

Dim SelStyleClass
Set SelStyleClass = New SelStyle_Class
SelStyleClass.Main_Style
Set SelStyleClass = Nothing
%>