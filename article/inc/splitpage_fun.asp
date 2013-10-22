<%
dim splitpage_getdata,splitpage_num,splitpage_page,splitpage_maxpage,splitpage_listNum,splitpage_orderstr
splitpage_listNum = 0

sub splitpage_returnData(class_sql,class_idname,class_page,class_selcolumn,class_recordcount)

		dim rs,sql,startid
		
		dim extendcount,p,listNum
		extendcount = 0
		if splitpage_listNum > 0 then 
			listNum = splitpage_listNum
		else
			listNum = 20
		end if
		If class_recordcount = 0 then
			sql = replace(class_sql,"{~~~}","count(*)")
			Set rs = LDExeCute(sql,0)
			if not rs.eof then
				extendcount = rs(0)
				if isnumeric(extendcount) = false then extendcount = 0
			end if
			rs.close
			set rs = nothing
		Else
			extendcount = class_recordcount
		End If
		
		dim maxpage
		maxpage = Fix(extendcount/listNum)
		if (extendcount mod listNum) > 0 then maxpage = maxpage + 1
		p = class_page
		if p >= maxpage then p = maxpage - 1
		if p < 0 then p = 0

		startid = 0
		dim selpage
		selpage = p
		
		if p > 100 then
			if splitpage_orderstr <> "" then
				'sql = replace(class_sql,"{~~~}","top " & (p+1)*listNum & " " & class_idname) & " order by " & splitpage_orderstr
				sql = sql_select(replace(class_sql,"{~~~}",class_idname) & " order by " & splitpage_orderstr,(p+1)*listNum)
			else
				'sql = replace(class_sql,"{~~~}","top " & (p+1)*listNum & " " & class_idname) & " order by " & class_idname & " desc"
				sql = sql_select(replace(class_sql,"{~~~}",class_idname) & " order by " & class_idname & " desc",(p+1)*listNum)
			end if
			Set rs = LDExeCute(sql,0)
			If Not Rs.Eof Then
				Rs.Move p*listNum
				If Not Rs.Eof Then
					startid = ccur(Rs(0))
					selpage = 0
				end if
			end if
			rs.close
			set rs = nothing
		end if
		
		'sql = replace(class_sql,"{~~~}","top " & ((selpage+1)*listNum) & " " & class_selcolumn)
		'sql = replace(class_sql,"{~~~}","top " & ((selpage+1)*listNum) & " " & class_selcolumn)
		sql = class_sql
		If startid > 0 Then
			If inStr(lcase(splitpage_orderstr)," desc") or splitpage_orderstr = "" then
				if inStr(sql," where ") > 0 Then
					sql = sql & " and " & class_idname & "<=" & startid
				Else
					sql = sql & " where " & class_idname & "<=" & startid
				end if
			else
				if inStr(sql," where ") > 0 Then
					sql = sql & " and " & class_idname & ">=" & startid
				Else
					sql = sql & " where " & class_idname & ">=" & startid
				end if
			end if
		end if
		if splitpage_orderstr <> "" then
			sql = sql & " order by " & splitpage_orderstr
		else
			sql = sql & " order by " & class_idname & " desc"
		end if
		
		sql = sql_select(replace(sql,"{~~~}"," " & class_selcolumn),(selpage+1)*listNum)
		Set rs = LDExeCute(sql,0)
		if not rs.eof then
			if p<=100 and p>0 then
				rs.move p*listNum
			end if
			splitpage_getdata = rs.getrows(-1)
			splitpage_num = ubound(splitpage_getdata,2)
		else
			splitpage_num = -1
		end if
		rs.close
		set rs = nothing
		
		splitpage_page = p
		splitpage_maxpage = maxpage

end Sub

sub splitpage_viewpagelist(url,num,curp,ajaxobj)

	dim n
	%>
	<div class=clear></div>
	<div class="j_page">
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
	<div class=clear></div>
	<%

end sub

%>