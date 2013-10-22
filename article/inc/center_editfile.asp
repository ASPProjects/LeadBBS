<%

dim Form_Content,Upd_ErrInfo

sub center_editfile

		dim centereditfileClass
		set centereditfileClass = new center_editfileClass_Class
		set centereditfileClass = nothing
	
End sub

class center_editfileClass_Class

	Private form_fileid,FileName
	
	Private Sub Class_Initialize
	
		form_fileid = GetFormData("form_fileid")
		form_fileid = FormClass_CheckFormValue(form_fileid,"","int",0,"",0)
		select case form_fileid
			case 0:
				FileName = "inc/home_bannerlist.asp"
				EditFlag = 1
				Form_EditAnnounceID = -100
				GetAncUploaInfo
			case 1:
				FileName = "inc/sitebottom_info.asp"
				LMT_EnableUpload = 0
			case 2:
				FileName = "inc/default.css"
				LMT_EnableUpload = 0
		end select

		dim submitflag
		submitflag = GetFormData("submitflag")
		if submitflag = "" then
			private_getClassinfo
			center_Class_Form
		else
			private_getformdata
		end if
	
	End Sub
	
	
	private sub private_getformdata
	
		form_content = GetFormData("form_content")
		
		CALL FormClass_CheckFormValue(form_content,"内容","string","none","=~~~",65535)
		
		If CheckErrorStr <> "" Then
			Response.Write "<span class=cms_error>" & CheckErrorStr & "</span>"
			center_Class_Form
		Else
			select case form_fileid
			case 0:
				If Form_UpFlag = 1 Then
					Dim Upd_FileInfo,UploadSave
					Set UploadSave = New Upload_Save
					UploadSave.Upload_File
					Upd_FileInfo = UploadSave.Upd_FileInfo
					Upd_ErrInfo = UploadSave.Upd_ErrInfo
				End If
				form_content = Topic_HomePicInfo(659,171,-10)
				CALL Update_InsertSetupRID(1051,"article/" & FileName,7,form_content," and ClassNum=" & 7)
			case 1:
				CALL Update_InsertSetupRID(1051,"article/" & FileName,9,form_content," and ClassNum=" & 9)
			end select
			private_Saveformdata
		End If 
	
	End Sub
	
	private sub private_Saveformdata
	
		ADODB_SaveToFile form_content,FileName
		Response.Write "<span class=cms_ok>成功编辑信息.</span>"

	End Sub
	
	private function private_getClassinfo
	
		form_content = ADODB_LoadFile(FileName)
		
	end function
	
	Public Sub center_Class_Form
	
	%>
		<ul>
		<li><a href=center.asp?action=editfile&form_fileid=0>编辑首页图片新闻</a></li>
		<li><a href=center.asp?action=editfile&form_fileid=1>自定义网站底部信息</a></li>
		<li><a href=center.asp?action=editfile&form_fileid=2>CSS样式表</a></li>
		</ul>
	<%
		select case form_fileid
				case 0:
					Form_ActionStr = "首页图片新闻"
					CALL FormClass_Head(Form_ActionStr,1,"center.asp?action=editfile")
				case 1:
					Form_ActionStr = "网站底部信息"
					CALL FormClass_Head(Form_ActionStr,0,"center.asp?action=editfile")
				case 2:
					Form_ActionStr = "CSS样式表"
					CALL FormClass_Head(Form_ActionStr,0,"center.asp?action=editfile")
		end select
		CALL FormClass_ItemPring("","hidden","form_fileid",form_fileid,"","","","","")
		CALL FormClass_ItemPring("","hidden","submitflag","yes","","","","","")
		
		select case form_fileid
			case 0:
		%>
		<div class="itemline">
				<div class="iteminfo homeimagesfornews">
		<%
		call DisplayLeadBBSEditor1(2,Form_Content,1,0)
		%>
			</div>
		</div>
		<%
			case else
				CALL FormClass_ItemPring("","textarea","form_content",form_content,"500px;",15,"","","")
		end select
		FormClass_End
		%>
		<br /><br />
		<hr class=splitline>
		<b>编辑首页图片新闻说明: </b>
		<ol>
		<li>注释中可以填写说明及图片网址,格式为: 链接地址|注释(以|号分隔)</li>
		</ol>
		
		<%
	
	End Sub
	
	private Function Topic_HomePicInfo(Width,Height,tNum)

		dim DEF_IMG_PlayWidth : DEF_IMG_PlayWidth = Width
		dim DEF_IMG_PlayHeight : DEF_IMG_PlayHeight = Height
	
		Dim RewriteFlag,url
		If GetBinarybit(DEF_Sideparameter,16) = 0 Then
			RewriteFlag = 0
		else
			RewriteFlag = 1
		end if
		
		Dim Num : Num = abs(tNum)
		If isNumeric(Num) = 0 Then Num = 0
		If Num < 1 or Num > 50 Then Num = 6
		
		If isNumeric(Height) = 0 Then Height = Fix(DEF_UploadSwidth * (DEF_IMG_PlayHeight/DEF_IMG_PlayWidth))
		If isNumeric(Width) = 0 Then Width = DEF_UploadSwidth
		If Height < 1 Then Height=DEF_IMG_PlayHeight
		If Width < 1 Then Width=DEF_IMG_PlayWidth
	
		Dim Rs,SQL,GetData
		
		SQL = sql_select("Select U.ID,U.PhotoDir,U.SPhotoDir,U.NdateTime,U.Info,0,0 from Article_Upload as U where U.AnnounceID=-100 Order by U.ID DESC",Num)
	
		Set Rs = Con.ExeCute(SQL)
		If Not Rs.Eof Then
			GetData = Rs.GetRows(-1)
			Rs.Close
			Set Rs = Nothing
		Else
			Rs.Close
			Set Rs = Nothing
			Exit Function
		End If
		SQL = Ubound(GetData,2)
	
		Dim Str
		Dim UrlData,UrlLink,TitleList
		
		str = "<div class=""playimages"">" &_
			"<div class=""playimages_bg""></div>" &_
			"<div class=""playimages_info""></div>" &_
			"<ul>" &_
			"<li class=""on"">1</li>"
		For Rs = 1 To SQL
			str = str & "<li>" & Rs+1 & "</li>"
		Next
		str = str & "</ul>"
		str = str & "<div class=""playimages_list"">"
		
		Dim udir,tmp,infotxt
		For Rs = 0 To SQL
			If cCur(GetData(5,Rs)) <> 0 Then
			
				If RewriteFlag = 0 Then
					url = "a/a.asp?B=" & GetData(6,Rs) & "&id=" & GetData(5,Rs)
				else
					url = "a/topic-" & GetData(6,Rs) & "-" & GetData(5,Rs) & "-1.html"
				end if
				UrlLink = DEF_BBS_HomeUrl & url
			Else
				UrlLink = DEF_BBS_HomeUrl & DEF_CMS_UploadPhotoUrl & Replace(GetData(1,Rs),"\","/")
			End If
			
			udir = DEF_BBS_HomeUrl & DEF_CMS_UploadPhotoUrl
			If GetData(2,Rs) <> "" and DEF_EnableGFL = 0 Then
				UrlData = udir & Replace(GetData(2,Rs),"\","/")
			Else
				UrlData = udir & Replace(GetData(1,Rs),"\","/")
			End If
			if DEF_EnableGFL = 1 then
					call SaveSmallPic(Server.Mappath(UrlData),Server.Mappath(DEF_BBS_HomeUrl & "images/temp/NewsPic_Home_" & Rs & ".jpg"),DEF_IMG_PlayWidth,DEF_IMG_PlayHeight,-1)
					UrlData = "images/temp/NewsPic_Home_" & Rs & ".jpg?ver=" & Timer
			end if
			TitleList = htmlencode(GetData(4,Rs) & "")
			tmp = split(TitleList,"|")
			If ubound(tmp) >= 1 then
				UrlLink = tmp(0)
				infotxt = tmp(1)
			else
				UrlLink = TitleList
				infotxt = ""
			end if
			If Left(TitleList,3) = "re:" and len(TitleList) > 4 Then TitleList = Mid(TitleList,4)
	        	str = str & "<a href=""" & UrlLink & """ target=""_blank"""
	        	If Rs = 0 Then
	        		str = str & " style='z-index:2;background: url(" & UrlData & ") center no-repeat" & "'"
	        	else
	        		str = str & " style='z-index:1;background: url(" & UrlData & ") center no-repeat" & "'"
	        	end if
	        	str = str & " title=""" & infotxt & """"
	        	'str = str & "><img src=""" & UrlData & """ title=""" & TitleList & """ alt=""" & TitleList & """ /></a>"
	        	str = str & "></a>"
	        Next
		str = str & "</div></div>"
		str = str & "<" & "script src=""" & "inc/js/img.js" & DEF_Jer & """ type=""text/javascript""></script" & ">" & VbCrLf

		Topic_HomePicInfo = Str
	
	End Function
	
End Class

%>