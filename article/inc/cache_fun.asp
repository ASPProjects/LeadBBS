<!-- #include file=cache/CACHE_CMS_Announcement.asp -->
<!-- #include file=cache/CACHE_CMS_HOMECONTENT.asp -->
<!-- #include file=cache/CACHE_CMS_HOMESIDE.asp -->
<!-- #include file=cache/CACHE_CMS_INSIDE.asp -->
<!-- #include file=cache/CACHE_CMS_NAVIGATECLASS.asp -->
<%
const CMS_AnnouncementClassID = 1

class cms_cache_Class

	private forceRefresh

	Private Sub Class_Initialize
	
		Server.ScriptTimeOut = 600
		forceRefresh = 0
	
	end sub

	public sub Announcement

		Dim t
		'on error resume next
		t = DateDiff("s",CMS_Announcement_UpdateTime,DEF_Now)
		dim tmptime
		If Err Then
			tmptime = GetTimeValue(DEF_NOw)
		Else
			tmptime = GetTimeValue(CMS_Announcement_UpdateTime)
		End If
		If forceRefresh = 1 or ((t < 0 or t > DEF_UpdateInterval or Err) and Application(DEF_MasterCookies & "_CMS_Announcement") & "" <> "yes") Then
			'防止多重写入
			Application.Lock
			Application(DEF_MasterCookies & "_CMS_Announcement") = "yes"
			Application.UnLock
			Announcement_MakeFile(tmptime)
			If Err Then
				Err.clear
			End If
			Application.Contents.Remove(DEF_MasterCookies & "_CMS_Announcement")
		Else
			CMS_Announcement_View
		End If
	
	end sub
	
	private sub Announcement_MakeFile(tmptime)
	
		dim listnum,liststyle,classname
	
		dim sql,rs,getdata
		sql = "select id,classname,liststyle,listNum from article_newsclass where id=" & CMS_AnnouncementClassID
		set rs = ldexecute(sql,0)
		if rs.eof then
			rs.close
			set rs = nothing
			exit sub
		end if
		listnum = rs("listnum")
		classname = rs(1)
		liststyle = rs("liststyle")
		rs.close
		set rs = nothing
		
		dim str
		str = cms_listClass(CMS_AnnouncementClassID,listnum,classname,liststyle,tmptime)
		response.Write str
		Str = "<" & "%" & VbCrLf &_
		"Dim CMS_Announcement_UpdateTime" & VbCrLf &_
		"CMS_Announcement_UpdateTime = """ & htmlencode(DEF_Now) & """" & VbCrLf &_
		"" & VbCrLf &_
		"Sub CMS_Announcement_View" & VbCrLf &_
		"" & VbCrLf &_
		"%" & ">" & VbCrLf &_
		str &_
		"<" & "%" & VbCrLf &_
		"" & VbCrLf &_
		"End Sub" & VbCrLf &_
		"%" & ">" & VbCrLf
		CALL ADODB_SaveToFile(Str,DEF_BBS_HomeUrl & "article/inc/cache/CACHE_CMS_Announcement.asp")
	
	end sub
	
	
	public sub CMS_HOMECONTENT

		Dim t
		'on error resume next
		t = DateDiff("s",CMS_HOMECONTENT_UpdateTime,DEF_Now)
		
		dim tmptime
		If Err Then
			tmptime = GetTimeValue(DEF_NOw)
		Else
			tmptime = GetTimeValue(CMS_HOMECONTENT_UpdateTime)
		End If
	
		If forceRefresh = 1 or ((t < 0 or t > DEF_UpdateInterval or Err) and Application(DEF_MasterCookies & "_CMS_HOMECONTENT") & "" <> "yes") Then
			'防止多重写入
			Application.Lock
			Application(DEF_MasterCookies & "_CMS_HOMECONTENT") = "yes"
			Application.UnLock
			CMS_HOMECONTENT_MakeFile(tmptime)
			If Err Then
				Err.clear
			End If
			Application.Contents.Remove(DEF_MasterCookies & "_CMS_HOMECONTENT")
		Else
			CMS_HOMECONTENT_View
		End If
	
	end sub
	
	private sub CMS_HOMECONTENT_MakeFile(tmptime)

		dim str
		dim form_content
		form_content = ADODB_LoadFile(DEF_BBS_HomeUrl & "article/inc/home_channellist.asp")
		dim tmp,n,tmp2,existn
		dim form_type,form_title,form_listnum,form_id,form_extendflag,form_style
		tmp = split(form_content,VbCrLf)
		for n = 0 to ubound(tmp)
			tmp2 = split(tmp(n),"#~#^#")
			if ubound(tmp2) >= 4 then
				form_type = cstr(tmp2(0))
				form_title = tmp2(1)
				form_listnum = tmp2(2)
				form_id = tmp2(3)
				form_extendflag = tmp2(4)
				form_style = tmp2(5)
				select case form_type
					case "0": form_type = 0
					case "1": form_type = 1
					case "2": form_type = 2
					case "3": form_type = 3
					case "4": form_type = 4
					case else
							form_type = 999
				end select
			end if
			if form_type <> 999 then
				dim tmpurl : tmpurl = DEF_BBS_HomeUrl
				select case form_type								
					case 0:
						tmpurl = tmpurl & "search/list.asp?1"" target=""_blank"
					case 1:
						tmpurl = tmpurl & "search/list.asp?2"" target=""_blank"
					case 2:
						tmpurl = tmpurl & "b/b.asp?E=1&EID=" & form_ID & """ target=""_blank"
					case 3:
						if cstr(form_extendflag) = "1" then
							tmpurl = tmpurl & "b/b.asp?b=" & form_id & "&E=0"" target=""_blank"
						else
							tmpurl = tmpurl & "b/b.asp?b=" & form_id & """ target=""_blank"
						end if
					case 4:
						tmpurl = tmpurl & "article/article.asp?classid=" & form_id
				End select
				str = str & "<div class=""cell"">" & VbCrLf
						If form_title <> "" Then
							str = str & "<div class=""cms_index_channelhead"">" & VbCrLf
							str = str & "<a href=""" & tmpurl & """ class=""title"">" & form_title & "</a>" & VbCrLf
							str = str & "</div>" & VbCrLf
						End If
						str = str & "<div class=""cms_index_channelcontent"">" & VbCrLf
							select case form_type
								case 0,1,2,3:
									str = str & topic_listClass(form_id,form_listnum,"channel",form_style,form_type,form_extendflag,tmptime)
									'response.write Topic_AnnounceList(0,form_listnum,0,"yes","0","0","none")
								case 4:
									str = str & cms_listClass(form_id,form_listnum,"channel",form_style,tmptime)
							end select
						str = str & "</div>" & VbCrLf
					str = str & "</div>" & VbCrLf
			end if
		next
		response.Write str
		Str = "<" & "%" & VbCrLf &_
		"Dim CMS_HOMECONTENT_UpdateTime" & VbCrLf &_
		"CMS_HOMECONTENT_UpdateTime = """ & htmlencode(DEF_Now) & """" & VbCrLf &_
		"" & VbCrLf &_
		"Sub CMS_HOMECONTENT_View" & VbCrLf &_
		"" & VbCrLf &_
		"%" & ">" & VbCrLf &_
		str &_
		"<" & "%" & VbCrLf &_
		"" & VbCrLf &_
		"End Sub" & VbCrLf &_
		"%" & ">" & VbCrLf
		CALL ADODB_SaveToFile(Str,DEF_BBS_HomeUrl & "article/inc/cache/CACHE_CMS_HOMECONTENT.asp")
	
	end sub
	
	
	public sub CMS_HOMESIDE

		Dim t
		'on error resume next
		t = DateDiff("s",CMS_HOMESIDE_UpdateTime,DEF_Now)
		dim tmptime
		If Err Then
			tmptime = GetTimeValue(DEF_NOw)
		Else
			tmptime = GetTimeValue(CMS_HOMESIDE_UpdateTime)
		End If
		If forceRefresh = 1 or ((t < 0 or t > DEF_UpdateInterval or Err) and Application(DEF_MasterCookies & "_CMS_HOMESIDE") & "" <> "yes") Then
			'防止多重写入
			Application.Lock
			Application(DEF_MasterCookies & "_CMS_HOMESIDE") = "yes"
			Application.UnLock
			CMS_HOMESIDE_MakeFile(tmptime)
			If Err Then
				Err.clear
			End If
			Application.Contents.Remove(DEF_MasterCookies & "_CMS_HOMESIDE")
		Else
			CMS_HOMESIDE_View
		End If
	
	end sub
	
	private sub CMS_HOMESIDE_MakeFile(tmptime)
	
		dim str : str = ""
		dim classdata,classdatanum
	
		dim sql,rs,getdata
		sql = "select id,classname,liststyle,listNum from article_newsclass where listflag=1 or listflag=3 or listflag=4 order by orderflag asc"
		set rs = ldexecute(sql,0)
		if rs.eof then
			rs.close
			set rs = nothing
			exit sub
		end if
		classdata = rs.getrows(-1)
		rs.close
		set rs = nothing
		classdatanum = ubound(classdata,2)
		
		dim n
	
		for n = 0 to classdatanum
			str = str & cms_listClass(classdata(0,n),classdata(3,n),classdata(1,n),classdata(2,n),tmptime)
		next
		response.Write str
		Str = "<" & "%" & VbCrLf &_
		"Dim CMS_HOMESIDE_UpdateTime" & VbCrLf &_
		"CMS_HOMESIDE_UpdateTime = """ & htmlencode(DEF_Now) & """" & VbCrLf &_
		"" & VbCrLf &_
		"Sub CMS_HOMESIDE_View" & VbCrLf &_
		"" & VbCrLf &_
		"%" & ">" & VbCrLf &_
		str &_
		"<" & "%" & VbCrLf &_
		"" & VbCrLf &_
		"End Sub" & VbCrLf &_
		"%" & ">" & VbCrLf
		CALL ADODB_SaveToFile(Str,DEF_BBS_HomeUrl & "article/inc/cache/CACHE_CMS_HOMESIDE.asp")
	
	end sub
	
	public sub CMS_INSIDE

		Dim t
		'on error resume next
		t = DateDiff("s",CMS_INSIDE_UpdateTime,DEF_Now)
		dim tmptime
		If Err Then
			tmptime = GetTimeValue(DEF_NOw)
		Else
			tmptime = GetTimeValue(CMS_INSIDE_UpdateTime)
		End If
		If forceRefresh = 1 or ((t < 0 or t > DEF_UpdateInterval or Err) and Application(DEF_MasterCookies & "_CMS_INSIDE") & "" <> "yes") Then
			'防止多重写入
			Application.Lock
			Application(DEF_MasterCookies & "_CMS_INSIDE") = "yes"
			Application.UnLock
			CMS_INSIDE_MakeFile(tmptime)
			If Err Then
				Err.clear
			End If
			Application.Contents.Remove(DEF_MasterCookies & "_CMS_INSIDE")
		Else
			CMS_INSIDE_View
		End If
	
	end sub
	
	private sub CMS_INSIDE_MakeFile(tmptime)
	
		dim str : str = ""
		dim classdata,classdatanum
		
		dim sql,rs,getdata
		sql = "select id,classname,liststyle,listNum from article_newsclass where listflag=1 or listflag=3 or listflag=5 order by orderflag asc"
		set rs = ldexecute(sql,0)
		if rs.eof then
			rs.close
			set rs = nothing
			exit sub
		end if
		classdata = rs.getrows(-1)
		rs.close
		set rs = nothing
		classdatanum = ubound(classdata,2)
		
		dim n
		for n = 0 to classdatanum
			'call cms_listClass(classdata(0,n),5,classdata(1,n),7)
			str = str & cms_listClass(classdata(0,n),classdata(3,n),classdata(1,n),classdata(2,n),tmptime)
		next

		response.Write str
		Str = "<" & "%" & VbCrLf &_
		"Dim CMS_INSIDE_UpdateTime" & VbCrLf &_
		"CMS_INSIDE_UpdateTime = """ & htmlencode(DEF_Now) & """" & VbCrLf &_
		"" & VbCrLf &_
		"Sub CMS_INSIDE_View" & VbCrLf &_
		"" & VbCrLf &_
		"%" & ">" & VbCrLf &_
		str &_
		"<" & "%" & VbCrLf &_
		"" & VbCrLf &_
		"End Sub" & VbCrLf &_
		"%" & ">" & VbCrLf
		
		CALL ADODB_SaveToFile(Str,DEF_BBS_HOMEUrl & "article/inc/cache/CACHE_CMS_INSIDE.asp")
	
	end sub

	rem 生成顶部分类导航缓存	
	public sub CMS_NAVIGATECLASS

		Dim t
		'on error resume next
		t = DateDiff("s",CMS_NAVIGATECLASS_UpdateTime,DEF_Now)
		If forceRefresh = 1 or ((t < 0 or t > DEF_UpdateInterval or Err) and Application(DEF_MasterCookies & "_CMS_NAVIGATECLASS") & "" <> "yes") Then
			'防止多重写入
			Application.Lock
			Application(DEF_MasterCookies & "_CMS_NAVIGATECLASS") = "yes"
			Application.UnLock
			CMS_NAVIGATECLASS_MakeFile
			If Err Then
				Err.clear
			End If
			Application.Contents.Remove(DEF_MasterCookies & "_CMS_NAVIGATECLASS")
		Else
			CMS_NAVIGATECLASS_View
		End If
	
	end sub
	
	private sub CMS_NAVIGATECLASS_MakeFile

		Dim str : str = ""
		dim classid
		classid = tonum(request.querystring("classid"),0)
		str = article_view_newsClass("listflag=1 or listflag=2",classid)
		CMS_NAVIGATECLASS_View
		Str = "<" & "%" & VbCrLf &_
		"Dim CMS_NAVIGATECLASS_UpdateTime" & VbCrLf &_
		"CMS_NAVIGATECLASS_UpdateTime = """ & htmlencode(DEF_Now) & """" & VbCrLf &_
		"" & VbCrLf &_
		"Sub CMS_NAVIGATECLASS_View" & VbCrLf &_
		"" & VbCrLf &_
		"%" & ">" & VbCrLf &_
		str &_
		"<" & "%" & VbCrLf &_
		"" & VbCrLf &_
		"End Sub" & VbCrLf &_
		"%" & ">" & VbCrLf
		
		CALL ADODB_SaveToFile(Str,DEF_BBS_HOMEUrl & "article/inc/cache/CACHE_CMS_NAVIGATECLASS.asp")
	
	end sub	

	private function article_view_newsClass(flag,classid)
	
		dim sql,rs,getdata
		dim str : str = ""
		sql = "select id,classname from article_newsclass where " & flag & " order by orderflag asc"
		set rs = ldexecute(sql,0)
		if rs.eof then
			rs.close
			set rs = nothing
			exit function
		end if
		getdata = rs.getrows(-1)
		rs.close
		set rs = nothing
		dim n,count
		count = ubound(getdata,2)
		str = str & "<" & "%" & VbCrLf
		str = str & "dim classid" & VbCrLf
		str = str & "classid = tonum(request.querystring(""classid""),0)" & VbCrLf
		str = str & "%" & ">" & VbCrLf
		dim tmp
		for n = 0 to count
			str = str & "<a class="""
			str = str & "cms_top_item"
			if lcase(left(getdata(1,n),5)) = "http:" and instr(getdata(1,n),"|") then
				tmp = split(getdata(1,n),"|")
				str = str & """ href=""" & tmp(0) & """ id=""cmstopitem" & getdata(0,n) & """>" & tmp(1) & "</a>" & VbCrLf
			else
				str = str & """ href=""<" & "%=DEF_BBS_HomeUrl" & "%" & ">article/article.asp?classid=" & getdata(0,n) & """ id=""cmstopitem" & getdata(0,n) & """>" & getdata(1,n) & "</a>" & VbCrLf
			end if
		next
		str = str & "<" & "script>" & VbCrLf
		str = str & "$(""#cmstopitem<" & "%=classid%" & ">"").attr(""class"",""cms_top_sel"");" & VbCrLf
		str = str & "</" & "script>" & VbCrLf
		article_view_newsClass = str
	
	end function
	
	public sub updatecache
	
		forceRefresh = 1
		Announcement
		CMS_HOMECONTENT
		CMS_HOMESIDE
		CMS_INSIDE
		CMS_NAVIGATECLASS
	
	End sub

end class
%>