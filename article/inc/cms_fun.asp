<!-- #include file=cms_article_fun.asp -->
<%
sub center_newsclass

		dim centernewsClassClass
		set centernewsClassClass = new center_newsClass_Class
		set centernewsClassClass = nothing
	
End sub


class center_newsClass_Class

	Private form_modifyid,form_classname,form_listflag,form_orderflag,form_liststyle,form_listNum,StyleItem
	Private GBL_LowClassString,GBL_LoopN

	Private Sub Class_Initialize
	
		StyleItem = Array("1|����Ӵ�","2|չʾ������Ҫ","3|չʾ���ͼƬ","4|��ͼƬʱ���ر���","5|��������¼����Ӵ�չʾͼƬ","6|���ͼƬ��ʾΪ��ͼƬ")
		If getformdata("action2") = "delete" then
			dim tmpidlist
			tmpidlist = getformdata("idlist")
			if instr("," & tmpidlist & ",",",1,") then
				response.Write "<div class=ajaxbox><div class=cms_error>���Ϊ1�ķ����ֹɾ����.</div></div>"
				exit sub
			else
				if cms_checkdeleteform("article_newsclass",1) = 1 then
					exit sub
				end if
			end if
		end if
		dim submitflag
		dim list
		list = left(GetFormData("list"),1)
		If list <> "1" then
			form_modifyid = GetFormData("form_modifyid")
			form_modifyid = FormClass_CheckFormValue(form_modifyid,"","int",0,"",0)
			submitflag = GetFormData("submitflag")
			If form_modifyid > 0 Then
				if private_getnewsClassinfo(form_modifyid) = 0 Then
					response.write "<span class=cms_error>����Ȩ���д˲���.</span>"
					exit sub
				End if
			End If
			if submitflag = "" then
				If form_modifyid > 0 Then
				Else
				End If
				center_newsClass_Form
			else
				private_getformdata
			end if
		else
			dim centermanagenewsClassClass
			set centermanagenewsClassClass = new center_managenewsClass_Class
			set centermanagenewsClassClass = nothing
		end if
	
	End Sub
	
	
	private sub private_getformdata
	
		form_classname = GetFormData("form_classname")
		form_listflag = GetFormData("form_listflag")
		form_orderflag = GetFormData("form_orderflag")
		form_listNum = GetFormData("form_listNum")
		
		form_classname = FormClass_CheckFormValue(form_classname,"��������","string","none","=~~~",255)
		
		If CheckErrorStr = "" Then form_listflag = FormClass_CheckFormValue(form_listflag,"�Ƿ������ҳ�г��˷��ࣺ","int","none","0|1|2|3|4|5",2)
		If CheckErrorStr = "" Then form_orderflag = FormClass_CheckFormValue(form_orderflag,"����˳��","int",0,"<~~~1|>~~~10000000",12)
		If CheckErrorStr = "" Then form_listNum = FormClass_CheckFormValue(form_listNum,"����չʾ��¼��Ŀ","int",1,"<~~~1|>~~~20",12)
		
		
		dim N,Temp2,TempN
		form_liststyle = 0
		Temp2 = 1
		For TempN = 0 to Ubound(StyleItem,1)
			N = Request("form_liststyle" & TempN+1)
			If N <> "1" Then N = "0"
			If N = "1" Then form_liststyle = form_liststyle+cCur(Temp2)
			Temp2 = Temp2*2
		Next
		
		If CheckErrorStr <> "" Then
			Response.Write "<span class=cms_error>" & CheckErrorStr & "</span>"
			center_newsClass_Form
		Else
			private_Saveformdata
		End If 
	
	End Sub
	
	private sub private_Saveformdata
	
		dim sql
		if form_modifyid > 0 then
			sql = "update article_newsclass set"&_
				" classname='" & cms_sql(form_classname) & "'"&_
				",listflag=" & cms_sql(form_listflag) & ""&_
				",orderflag=" & cms_sql(form_orderflag) & ""&_
				",liststyle=" & form_liststyle & "" &_
				",listNum=" & form_listNum & "" &_				
				" where id=" & form_modifyid
			call ldexecute(sql,1)
			Response.Write "<span class=cms_ok>�ɹ��༭��Ϣ.</span>"
		else
			sql = "insert into article_newsclass(classname,listflag,orderflag,liststyle,listNum)" &_
				" values('" & cms_sql(form_classname) & "'," & form_listflag & "" &_
				"," & form_orderflag & "" &_
				"," & form_liststyle & "" &_
				"," & form_listNum & "" &_
				")"
			call ldexecute(sql,1)
			Response.Write "<span class=cms_ok>�ɹ��������Ϣ.</span>"
		end if
		UpdateCacheData("data_artileclass.asp")

	End Sub
	
	private function private_getnewsClassinfo(UID)
	
		Dim RS,SQL,city,userid
		sql = "select * from article_newsclass where id=" & UID
		Set rs  = LDexecute(sql,0)
		If Not Rs.Eof Then
			form_classname = Rs("classname")
			form_listflag = Rs("listflag")
			form_orderflag = Rs("orderflag")
			form_liststyle = rs("liststyle")
			form_listNum = rs("listNum")
			private_getnewsClassinfo = 1
		else
			private_getnewsClassinfo = 0
		End If
		Rs.Close
		Set Rs = Nothing
		
	end function
	
	Public Sub center_newsClass_Form
	
		CALL FormClass_Head(Form_ActionStr,0,"center.asp?action=newsclass")
		CALL FormClass_ItemPring("","hidden","form_modifyid",form_modifyid,"","","","","")
		CALL FormClass_ItemPring("","hidden","submitflag","yes","","","","","")
		CALL FormClass_ItemPring("�Ƿ������ҳ�г��˷��ࣺ","select","form_listflag",form_listflag,"","","","0~~~��ȫ����ʾ|1~~~��ȫ��ʾ|2~~~ֻ��ʾ�ڶ���(����ʾ�ڲ���)|3~~~ֻ��ʾ�����в���|4~~~ֻ��ʾ����ҳ����|5~~~��ʾ����ҳ����","")
		CALL FormClass_ItemPring("��ʾ˳��","input","form_orderflag",form_orderflag,2,8,"����ԽС����Խǰ","","")
		CALL FormClass_ItemPring("���·������ƣ�","input","form_classname",form_classname,3,255,"����","","")
		CALL FormClass_ItemPring("����չʾ��ʽ","splitchecked","form_liststyle",form_liststyle,"","","",StyleItem,"")
		CALL FormClass_ItemPring("����չʾ��¼��Ŀ","input","form_listNum",form_listNum,2,2,"����չʾ�˷�����������ʾ������","","")
		FormClass_End
	
	End Sub

	private Function UpdateCacheData(savefile)

		Dim Rs,GetData,Num
		Set Rs = LDExeCute("Select id,classname from article_newsclass where ParentID=0 order by orderflag asc",0)
	
		If Not Rs.Eof Then
			GetData = Rs.GetRows(-1)
			Num = Ubound(GetData,2)
		Else
			Num = -1
		End If
		Rs.Close
		Set Rs = Nothing
		
		'on error resume next
		Dim TempStr
		TempStr = ""
	
		Dim N,WriteStr
		TempStr = TempStr & "["
	
		If Num = -1 Then
		Else
			For N = 0 to Num
				WriteStr = ""
				WriteStr = WriteStr & KillHTMLLabel(GetData(1,N))
				If StrLength(WriteStr) > 21 Then
					WriteStr = LeftTrue(WriteStr,18) & "..."
				End If	
				
				If N = 0 Then
					TempStr = TempStr & "{" & VbCrLf
				Else
					TempStr = TempStr & ",{" & VbCrLf
				End If
				TempStr = TempStr & "	""id"":" & GetData(0,N) & "," & VbCrLf
				TempStr = TempStr & "	""text"":""" & GetData(0,N) & "." & htmlencode(WriteStr) & """" & VbCrLf & "}"
				GBL_LowClassString = ""
				GBL_LoopN = 0
				GetLowClassString_Json GetData(0,n)
				If GBL_LowClassString <> "" Then TempStr = TempStr & GBL_LowClassString				
			Next
		End If
	
		TempStr = DEF_pageHeader & TempStr & "]"
		
		ADODB_SaveToFile TempStr,DEF_BBS_HomeUrl & "inc/IncHtm/" & savefile & ""
		If GBL_CHK_TempStr = "" Then
			Response.Write "<br><span class=cms_ok>2.�ɹ������ļ�../../inc/IncHtm/" & savefile & "��</span>"
		Else
			%><p><%=GBL_CHK_TempStr%><br>��������֧������д���ļ����ܣ���ʹ��FTP�ȹ��ܣ�<br>��<span Class=cms_error>inc/IncHtm/<%=savefile%></span>�ļ��滻���¿�������(ע�ⱸ��)<p>
			<textarea name="fileContent" cols="80" rows="20" class=fmtxtra><%=Server.htmlencode(TempStr)%></textarea><%
			GBL_CHK_TempStr = ""
		End If
	
	End Function
	
	Private sub GetLowClassString_Json(classid)

		If classid = "" or isNull(classid) or GBL_LoopN > 100 Then Exit sub
		Dim Rs,GetData,Num
		
		GBL_LoopN = GBL_LoopN + 1

		Set Rs = LDExeCute("Select id,classname from article_newsclass where ParentID=" & classid & " order by orderflag asc",0)
		If Not Rs.Eof Then
			GetData = Rs.GetRows(-1)
			Num = Ubound(GetData,2)
		Else
			Num = -1
		End If
		Rs.Close
		Set Rs = Nothing
	
		Dim Temp
		Dim WriteStr,N
		For N = 0 to Num
			WriteStr = ""
			WriteStr = WriteStr & KillHTMLLabel(GetData(1,N))
			If StrLength(WriteStr) > 21 Then
				WriteStr = LeftTrue(WriteStr,18) & "..."
			End If
			GBL_LowClassString = GBL_LowClassString & ",{" & VbCrLf
			GBL_LowClassString = GBL_LowClassString & "	""id"":" & GetData(0,N) & "," & VbCrLf
			GBL_LowClassString = GBL_LowClassString & "	""text"":""" & GetData(0,N) & "." & htmlencode(WriteStr) & """" & VbCrLf & "}"
			GetLowBoardString_Json GetData(0,N)
		Next			
		GBL_LoopN = GBL_LoopN - 1
	
	end sub
	
End Class

class center_managenewsClass_Class

	Private class_page,class_sql,class_idname,class_selcolumn
	
	Private Sub Class_Initialize
	
		Dim sql_extend
		sql_extend = ""

		class_page = GetFormData("page")
		class_page = FormClass_CheckFormValue(class_page,"","int","0","<~~~0|>~~~10000000000",12)
		
		class_sql = "select {~~~} from article_newsclass " & sql_extend
		class_idname = "id"
		class_selcolumn = "id,classname"
		CALL splitpage_returnData(class_sql,class_idname,class_page,class_selcolumn,0)
		
		private_managelist
		
		CALL splitpage_viewpagelist("center.asp?action=newsclass",splitpage_maxpage,splitpage_page,"")
			
	End sub
	
	private sub private_managelist
	
		cms_selectFormScript("center.asp?action=newsclass")
		%>
		<div class="title"><div class="titlebg">�������·��� <a href=center.asp?action=newsclass>�������·���</a></div></div>
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class=table_in>
				<tr class="tbinhead cms_tbinhead">
					<td width=60><div class=cms_value>���</div></td>
					<td><div class=cms_value>��������</div></td>
					<td>�༭</td>
				</tr>
				<%dim n
				for n = 0 to splitpage_num%>
				<tr>
					<td class=tdbox><%=splitpage_getdata(0,n)%></td>
					<td class=tdbox><span class="layerico"><%
					If ccur(splitpage_getdata(0,n)) <> 1 then
					%><input class="fmchkbox" type="checkbox" name="ids" id="ids<%=n%>" value="<%=splitpage_getdata(0,n)%>" onclick="delbody_view(this);" /><%
					else
						Response.Write "<span class=""grayfont"">[��ҳ������ר�÷���]</span>"
					end if%></span><a href=center.asp?action=newsclass&form_modifyid=<%=splitpage_getdata(0,n)%>><%=splitpage_getdata(1,n)%></a></td>
					<td class=tdbox><a href=center.asp?action=newsclass&form_modifyid=<%=splitpage_getdata(0,n)%>>�༭<a></td>
				</tr>
				<%next%>
		</table>
		<br />
		<hr class=splitline>
		<%
	
	end sub

End Class




%>