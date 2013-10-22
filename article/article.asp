<!-- #include file=../inc/BBSsetup.asp -->
<!-- #include file=../inc/User_Setup.ASP -->
<!-- #include file=../inc/Board_Popfun.asp -->
<!-- #include file=inc/popfun.asp -->
<!-- #include file=../inc/Upload_Fun.asp -->
<%
DEF_BBS_homeUrl = "../"
dim Form_ActionStr,Form_ActionCommand

Sub Main

	initdatabase
	cms_article_getAction
	article_SiteHead("")
	main_body
	Closedatabase

End Sub

Sub main_body
%>
<div class="body_area_out">
<%
cms_DisplayBBSNavigate("<span class=navigate_string_step>" & Form_ActionStr & "</span>")
cms_bodyhead_index("inpage")
cms_bodyBottom%>


</div>
<%
cms_SiteBottom

End Sub

dim readid,article_title,article_content,article_classid,article_modifytime,article_ndatetime,article_author,article_fromauthor
sub cms_article_getAction

	dim tmp,rs,sql,tmp2,newsclassname
	tmp = requestFormData("classid")
	Form_ActionCommand = ""
		tmp = FormClass_CheckFormValue(tmp,"","int","0","<~~~0|>~~~10000000000",12)
		if tmp > 0 then
			sql = sql_select("select t1.id,t1.title,t2.id,t2.classname,t1.content,t1.modifytime,t1.ndatetime,t1.author,t1.fromauthor from article_newsarticle as t1 left join article_newsclass as t2 on t1.classid=t2.id where t1.classid=" & cms_sql(tmp),2)
			set rs = ldexecute(sql,0)
			if not rs.eof then
				readid = rs(0)
				article_title = rs(1)
				article_content = rs(4)
				article_modifytime = rs(5)
				article_ndatetime = rs(6)
				article_author = rs(7)
				article_fromauthor = rs(8)
				Form_ActionStr = "<a href=article.asp?classid=" & tmp & ">" & rs(3) & "</a>"
				rs.movenext
				if not rs.eof then
					article_classid = tmp
					Form_ActionCommand = "listnews"
					Form_ActionStr = "<a href=article.asp?classid=" & tmp & ">" & rs(3) & "</a>"
				Else
					Form_ActionCommand = "readarticle"
				End if
			end if
			rs.close
			set rs = nothing
		end if
	if Form_ActionCommand = "" then
		tmp = requestFormData("articleid")
		tmp = FormClass_CheckFormValue(tmp,"","int","0","<~~~0|>~~~10000000000",12)
		if tmp > 0 then
			sql = sql_select("select t1.id,t1.title,t2.id,t2.classname,t1.content,t1.modifytime,t1.ndatetime,t1.author,t1.fromauthor from article_newsarticle as t1 left join article_newsclass as t2 on t1.classid=t2.id where t1.id=" & cms_sql(tmp),1)
			set rs = ldexecute(sql,0)
			if not rs.eof then
				Form_ActionCommand = "readarticle"
				readid = rs(0)
				article_title = rs(1)
				article_content = rs(4)
				article_modifytime = rs(5)
				article_ndatetime = rs(6)
				article_author = rs(7)
				article_fromauthor = rs(8)
				Form_ActionStr = "<a href=article.asp?classid=" & rs(2) & ">" & rs(3) & "</a>"
			end if
			rs.close
			set rs = nothing
		end if
	end if

end sub

sub cms_bodyhead_index(sideinfo)%>

<div class="area">
<div class="cms_body_box">
<div class="cms_body">
<div class="main">
	<div class="content_side_right" id="p_side">
		<%
		dim cmscacheClass
		select case sideinfo
			case "homepage":
				set cmscacheClass = new cms_cache_Class
				cmscacheClass.CMS_HOMESIDE
				set cmscacheClass = nothing
			case "inpage":
				set cmscacheClass = new cms_cache_Class
				cmscacheClass.CMS_INSIDE
				set cmscacheClass = nothing
		end select		
		%>
	</div>
	<div class="content_main_right">
		<div class="content_main_2_right">
		<div class="content_main_body">
		<%
		select case Form_ActionCommand
			case "listnews":
				call cms_body_listNews(article_classid)
			case "readarticle":
				cms_body_readarticle(readid)
		
		end select
		
		
		%>
		<div style="height:20px;"></div>
		<%

		
End Sub

sub cms_body_listNews(classid)

	dim class_sql,class_idname,class_selcolumn,class_page,sql_extend
	sql_extend = " where t1.classid=" & classid
	class_page = 0
	class_sql = "select {~~~} from article_newsarticle as t1 " & sql_extend
	class_idname = "t1.id"
	class_selcolumn = "t1.ID,t1.title"
	
	class_page = requestFormData("page")
	class_page = FormClass_CheckFormValue(class_page,"","int","0","<~~~0|>~~~10000000000",12)
	
	splitpage_listNum = 30
	CALL splitpage_returnData(class_sql,class_idname,class_page,class_selcolumn,0)
%>
<div class=cms_listtop>
<div class=cms_main_info_left><div class="title cms_listtitle"><%=Form_ActionStr%></div>
	<ul class=cms_listnews>
		<%dim n
		for n = 0 to splitpage_num
		%>
					<li>
					◇ <a href=article.asp?articleid=<%=splitpage_getdata(0,n)%>><%=splitpage_getdata(1,n)%></a>
					</li>
		<%next%>
		<li>
		<div class=clear></div>
<%

		dim extendurl
		extendurl = ""
		CALL splitpage_viewpagelist("article.asp?classid=" & classid & extendurl,splitpage_maxpage,splitpage_page,"")
		%>
		</li>
		</ul>
</div>
</div>
<div class=clear></div>
		<%

end sub


Function PrintTrueText(tempString)

	If tempString<>"" Then
		PrintTrueText=Replace(Replace(Replace(Replace(Replace(Replace(htmlEncode(tempString),VbCrLf & " ","<br />" & "&nbsp;"),"[P] ","[P]&nbsp;"),VbCrLf,"<br />" & VbCrLf),"   "," &nbsp; "),"  "," &nbsp;"),chr(9)," &nbsp; &nbsp; &nbsp;")
		If Left(PrintTrueText,1) = chr(32) Then
			PrintTrueText = "&nbsp;" & Mid(PrintTrueText,2)
		End If
	Else
		PrintTrueText=""
	End If

End Function

sub cms_body_readarticle(classid)

dim LMTDEF_ConvetType : LMTDEF_ConvetType = GetBinarybit(DEF_Sideparameter,7)
%>

<div class=cms_listtop>
<div class=cms_main_info_left><div class="title cms_articletitle"><%=article_title%></div>

<div class="cms_article_note"><%=restoretime(article_modifytime)%><%
	If article_author <> "" Then
	%>，作者：<%=article_author%>
	<%End If
	If article_fromauthor <> "" Then
	%>，来自：<%=article_fromauthor%><%
	End If%></div>
	<script src="<%=DEF_BBS_HomeUrl%>a/inc/leadcode.js<%=DEF_Jer%>" type="text/javascript"></script>
	<div class=cms_listnews>
					<div>
					<%
	article_content = PrintTrueText(article_content)
	if LMTDEF_ConvetType = 1 then
		dim bbsObj,outstr
		Set bbsObj = CreateObject("leadbbs.bbsCode")
		
		if inStr(lcase(Request.ServerVariables("HTTP_USER_AGENT")),"msie") then
			Response.Write bbsObj.convertcode(article_content,DEF_BBS_HomeUrl,DEF_DownKey & "&type=1","|all|",outstr,"msie")
		else
			Response.Write bbsObj.convertcode(article_content,DEF_BBS_HomeUrl,DEF_DownKey & "&type=1","|all|",outstr,"other")
		end if
		set bbsObj = nothing
	else%>
		<script type="text/javascript">
		var GBL_domain="|all|";
		var DEF_DownKey="<%=UrlEncode(DEF_DownKey)%>&type=1";
		HU="<%=DEF_BBS_HomeUrl%>";
		</script>
	<%
		Response.Write "<div id=""articlecontent"">" & VbCrLf
		Response.Write article_content
		Response.Write "</div>"
		%>
		<script type="text/javascript">
		<!--
		leadcode('articlecontent');
		-->
		</script>
		<%
	End If
					%>
					</div>
		</div>
</div>
</div>
<div class=clear></div>
		<%

end sub

Main
%>
