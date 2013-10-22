<%
sub center_newsarticle

		dim centernewsarticleClass
		set centernewsarticleClass = new center_newsarticle_Class
		set centernewsarticleClass = nothing
	
End sub

sub center_newsmanage

		dim centermanagenewsArticleClass
		set centermanagenewsArticleClass = new center_managenewsArticle_Class
		set centermanagenewsArticleClass = nothing
	
End sub

class center_newsArticle_Class

	Private form_modifyid,form_title,form_classid,form_fromauthor,form_author

	
	Private Sub Class_Initialize
	
		if cms_checkdeleteform("cms_newsarticle",1) = 1 then
			exit sub
		end if
		dim submitflag
		form_modifyid = GetFormData("form_modifyid")
		form_modifyid = FormClass_CheckFormValue(form_modifyid,"","int",0,"",0)
		submitflag = GetFormData("submitflag")
		form_author = GBL_CHK_User

		If form_modifyid > 0 Then
			if private_getarticleclassinfo(form_modifyid) = 0 Then
				response.write "<span class=cms_error>����Ȩ���д˲���.</span>"
				exit sub
			End if
		End If
		If form_modifyid > 0 Then
			EditFlag = 1
		Else
			EditFlag = 0
		End If
		Form_EditAnnounceID = form_modifyid
		GetAncUploaInfo

		if submitflag = "" then
			If form_modifyid > 0 Then
			Else
			End If
			center_articleclass_Form
		else
			private_getformdata
		end if
	
	End Sub
	
	private sub private_getformdata
	
		form_title = GetFormData("form_title")
		form_content = Getformdata("form_content")
		form_classid = getformdata("form_classid")
		form_fromauthor = left(getformdata("form_fromauthor"),50)
		form_author = left(getformdata("form_author"),50)
		
		form_title = FormClass_CheckFormValue(form_title,"���±���","string","none","=~~~",255)
		If CheckErrorStr = "" Then form_content = FormClass_CheckFormValue(form_content,"��������","string","none","=~~~",65535)
		If CheckErrorStr = "" Then form_classid = FormClass_CheckFormValue(form_classid,"���·���","int","none","<~~~1|>~~~10000000",12)
		
		If CheckErrorStr <> "" Then
			Response.Write "<span class=cms_error>" & CheckErrorStr & "</span>"
			center_articleclass_Form
		Else
			If Form_UpFlag = 1 Then
				if form_modifyid > 0 then
					Form_EditAnnounceID = form_modifyid
				Else
					Form_EditAnnounceID = 0
				End If
				Dim Upd_FileInfo,UploadSave
				Set UploadSave = New Upload_Save
				UploadSave.Upload_File
				Upd_FileInfo = UploadSave.Upd_FileInfo
				Upd_ErrInfo = UploadSave.Upd_ErrInfo
			End If
			
			private_Saveformdata

			If Form_UpFlag = 1 Then
				if EditFlag = 0 then
					Form_EditAnnounceID = private_getMaxArticlID
				End If
				UploadSave.UpdateUpload(Form_EditAnnounceID)
				Set UploadSave = Nothing
			End If
		End If 
	
	End Sub
	
	private function private_getMaxArticlID
	
		dim sql,rs
		sql = "select max(id) from article_newsarticle"
		set rs = ldexecute(sql,0)
		if rs.eof then
			private_getMaxArticlID = 0
		else
			private_getMaxArticlID = ccur(rs(0))
		end if
		rs.close
		set rs = nothing
	
	End function
	
	private sub private_Saveformdata
	
		dim sql
		if form_modifyid > 0 then
			sql = "update article_newsarticle set"&_
				" classid=" & cms_sql(form_classid) & ""&_
				",title='" & cms_sql(form_title) & "'"&_
				",content='" & cms_sql(form_content) & "'"&_
				",modifytime=" & gettimevalue(DEF_Now) & ""&_
				",author='" & cms_sql(form_author) & "'" &_
				",fromauthor='" & cms_sql(form_fromauthor) & "'" &_
				" where id=" & form_modifyid
			call ldexecute(sql,1)
			Response.Write "<span class=cms_ok>�ɹ��༭��Ϣ.</span>"
		else
			sql = "insert into article_newsarticle(title,content,classid,ndatetime,modifytime,author,fromauthor)" &_
				" values('" & cms_sql(form_title) & "'" &_
				",'" & cms_sql(form_content) & "'" &_
				"," & cms_sql(form_classid) & "" &_
				"," & gettimevalue(DEF_Now) & "" &_
				"," & gettimevalue(DEF_Now) & "" &_
				",'" & cms_sql(form_author) & "'" &_
				",'" & cms_sql(form_fromauthor) & "'" &_
				")"
			call ldexecute(sql,1)
			Response.Write "<span class=cms_ok>�ɹ��������Ϣ.</span>"
		end if	

	End Sub
	
	private function private_getarticleclassinfo(UID)
	
		Dim RS,SQL,userid
		sql = "select title,content,classid,fromauthor,author from article_newsarticle where id=" & UID
		Set rs  = LDexecute(sql,0)
		If Not Rs.Eof Then
			form_title = Rs("title")
			form_classid = rs("classid")
			form_content = rs("content")
			form_fromauthor = rs("fromauthor")
			form_author = rs("author")
			private_getarticleclassinfo = 1
		else
			private_getarticleclassinfo = 0
		End If
		Rs.Close
		Set Rs = Nothing
		
	end function
	
	Public Sub center_articleclass_Form
	
		CALL FormClass_Head(Form_ActionStr,1,"center.asp?action=newsarticle")
		CALL FormClass_ItemPring("","hidden","form_modifyid",form_modifyid,"","","","","")
		CALL FormClass_ItemPring("","hidden","submitflag","yes","","","","","")
		CALL FormClass_ItemPring("���±��⣺","input","form_title",form_title,4,255,"����","","")
		CALL FormClass_ItemPring("���ߣ�","input","form_author",form_author,3,50,"","","")
		CALL FormClass_ItemPring("���ԣ�","input","form_fromauthor",form_fromauthor,3,50,"","","")
		call cms_selectnewsclass
		%>
				<div class="itemline">
				<div class="iteminfo cms_article">
		<%
				if form_modifyid > 0 then
					EditFlag = 1
					Form_EditAnnounceID = form_modifyid
				Else
					EditFlag = 0
				End If
		call DisplayLeadBBSEditor1(2,Form_Content,1,0)
		%>
			</div>
		</div>
		<%
		'CALL FormClass_ItemPring("�������ݣ�","textarea","form_content",form_content,"500px;",15,"","","")
		FormClass_End
	
	End Sub
	
	

	private sub cms_selectnewsclass
	
	
			dim sql,rs,getdata
			sql = "select id,classname from article_newsclass order by orderflag asc"
			set rs = ldexecute(sql,0)
			if rs.eof then
				rs.close
				set rs = nothing
				exit sub
			end if
			getdata = rs.getrows(-1)
			rs.close
			set rs = nothing
			dim n,count
			count = ubound(getdata,2)
			
			dim str
			str = "0~~~ѡ�����"
			for n = 0 to count
				str = str & "|" & getdata(0,n) & "~~~" & replace(htmlencode(getdata(1,n)),"|","��")
			next		
			CALL FormClass_ItemPring("���·���","select","form_classid",form_classid,"","","",str,"")
	
	end sub
	
End Class

class center_managenewsArticle_Class

	Private class_page,class_sql,class_idname,class_selcolumn
	
	Private Sub Class_Initialize
	
		Dim sql_extend,classid
		sql_extend = ""

		class_page = GetFormData("page")
		class_page = FormClass_CheckFormValue(class_page,"","int","0","<~~~0|>~~~10000000000",12)
		classid = GetFormData("classid")
		classid = FormClass_CheckFormValue(classid,"","int",0,"<~~~0|>~~~10000000000",12)
		if classid <> "0" and classid > 0 then sql_extend = " where classid=" & classid
		
		class_sql = "select {~~~} from article_newsarticle " & sql_extend
		class_idname = "id"
		class_selcolumn = "id,title"
		CALL splitpage_returnData(class_sql,class_idname,class_page,class_selcolumn,0)
		
		
		dim paraextend
		if classid>0 then paraextend = "&classid=" & classid
		
		dim sql,rs,getdata
		sql = "select id,classname from article_newsclass order by orderflag asc"
		set rs = ldexecute(sql,0)
		if rs.eof then
			rs.close
			set rs = nothing
			exit sub
		end if
		getdata = rs.getrows(-1)
		rs.close
		set rs = nothing
		dim n,count
		count = ubound(getdata,2)
		Response.Write "<span class=grayfont>��ѡ�����: </span>"
		%>
		<a href=center.asp?action=newsmanage<%if classid=0 then response.write " style='font-weight:bold'"%>>ȫ��</a>
		<%
		for n = 0 to count
			%>
			<a href=center.asp?action=newsmanage&classid=<%=getdata(0,n)%><%if classid=ccur(getdata(0,n)) then response.write " style='font-weight:bold'"%>><%=getdata(1,n)%></a>
			<%
		next
		
		private_managelist
		
		CALL splitpage_viewpagelist("center.asp?action=newsmanage" & paraextend,splitpage_maxpage,splitpage_page,"")
			
	End sub
	
	private sub private_managelist
	
		cms_selectFormScript("center.asp?action=newsarticle")
		%>
		<div class="title"><div class="titlebg">�������� <a href=center.asp?action=newsarticle>����������</a></div></div>
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class=table_in>
				<tr class="tbinhead cms_tbinhead">
					<td width=60><div class=cms_value>���</div></td>
					<td><div class=cms_value>��������</div></td>
					<td width=60><div class=cms_value>�༭</div></td>
					
				</tr>
				<%dim n
				for n = 0 to splitpage_num%>
				<tr>
					<td class=tdbox><%=splitpage_getdata(0,n)%></td>
					<td class=tdbox><span class="layerico"><input class="fmchkbox" type="checkbox" name="ids" id="ids<%=n%>" value="<%=splitpage_getdata(0,n)%>" onclick="delbody_view(this);" /></span><a href=center.asp?action=newsarticle&form_modifyid=<%=splitpage_getdata(0,n)%>><%=splitpage_getdata(1,n)%></a></td>
					<td class=tdbox><a href=center.asp?action=newsarticle&form_modifyid=<%=splitpage_getdata(0,n)%>>�༭</a></td>
				</tr>
				<%next%>
		</table>
		<br />
		<hr class=splitline>
		<%

	end sub

End Class
%>