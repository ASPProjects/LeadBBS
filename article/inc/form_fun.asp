<%

Sub FormClass_Head(title,enableuploadflag,actionurl)
%>
	<script>
	var ValidationPassed = true;
	function submitonce(theform)
	{
		if(ValidationPassed == false)return;
		if(typeof edt_checkContent != "undefined")
		{
			edt_checkContent();
			lg = edt_getdoclen();
			if(lg > <%=DEF_MaxTextLength%>)
			{
				alert("��������ݳ�����<%=DEF_MaxTextLength%>���֣�Ŀǰ����" + lg + "����\n");
				ValidationPassed = false;
				submitflag = 0;
				return;
			}
		}
		submit_disable(theform);
	}
	</script>
	<div class="title"><div class="titlebg"><%=title%></div></div>
			<div class="itemtable">
			<%If enableuploadflag = 0 Then%>
			<form action=<%=actionurl%> method=post name=LeadBBSFm id="LeadBBSFm" onSubmit="submitonce(this);return ValidationPassed">
			<%Else%>
			<form action=<%=actionurl%><%
			if instr(actionurl,"?") then
				response.write "&dontRequestFormFlag=1"
			else
				response.write "?dontRequestFormFlag=1"
			end if%> method=post name=LeadBBSFm id=LeadBBSFm enctype="multipart/form-data" onSubmit="submitonce(this);if(ValidationPassed)Upl_submit();return ValidationPassed">
			<%End If%>
			<%

End Sub

' Remark ��Ŀע��,maxlength,������󳤶� inputClass,���򳤶���,���ֱ�ʾ,ItemName,������,ItemValue,���ֵ,PrintType,��ӡ�����,SelectStr,�����select ���,�������ֵ��,��|�ŷָ�
Sub FormClass_ItemPring(title,PrintType,ItemName,Item_Value,inputClass,maxLength,reMark,SelectStr,moreEvent)

	dim ItemValue
	ItemValue = Item_Value
	if isnull(ItemValue) then ItemValue = ItemValue & ""
	Dim N,Tmp,Count,Tmp2
	If title <> "single" Then
%>
			<div class="itemline"<%If PrintType = "hidden" Then Response.Write " style=""display:none"""%>>
				<span class="itemtitle">
					<%=title%>
				</span>
				<span class="iteminfo">
					<%
	end if
					Select Case PrintType
						Case "printvalue":
							Response.Write ItemValue
						Case "input_notzero":
							%><input class='fminpt input_<%=inputClass%>' maxLength=<%=maxLength%> name="<%=ItemName%>" type=input Value="<% If ItemValue<>"" and cstr(ItemValue) <> "0" Then Response.Write Server.HtmlEncode(ItemValue)%>"<%=moreEvent%>><%If reMark<>"" Then%> <span class=cms_remark><%=reMark%></span><%End If
							
						Case "input":
							%><input class='fminpt input_<%=inputClass%>' maxLength=<%=maxLength%> name="<%=ItemName%>" type=input Value="<% If ItemValue<>"" Then Response.Write Server.HtmlEncode(ItemValue)%>"<%=moreEvent%>><%If reMark<>"" Then%> <span class=cms_remark><%=reMark%></span><%End If%><%
						Case "hidden":
							%><input name="<%=ItemName%>" type=hidden Value="<% If ItemValue<>"" Then Response.Write Server.HtmlEncode(ItemValue)%>"<%=moreEvent%>>
							<%
						Case "password":	
							%><input class='fminpt input_<%=inputClass%>' maxLength=<%=maxLength%> name="<%=ItemName%>" type=password Value="<% If ItemValue<>"" Then Response.Write Server.HtmlEncode(ItemValue)%>"<%=moreEvent%>><%If reMark<>"" Then%> <span class=cms_remark><%=reMark%></span><%End If
						Case "select":
							dim sss
							if instr(moreEvent," class=""") then
								tmp = replace(moreEvent," class="""," class=""")
							else
								tmp = moreEvent & " class=""easyui-combobox"""
							end if
							Response.Write "<select data-options=""panelHeight: 'auto',editable:false"" name=""" & ItemName & """" & tmp & ">"
							Tmp = Split(SelectStr,"|")
							Count = Ubound(Tmp,1)
							For N = 0 To Count
								Tmp2 = Split(Tmp(N),"~~~")
								If cstr(ItemValue) <> cstr(Tmp2(0)) Then
									Response.Write "<option value=""" & Tmp2(0) & """>" & Tmp2(1) & "</option>"
								Else
									Response.Write "<option value=""" & Tmp2(0) & """ selected='selected'>" & Tmp2(1) & "</option>"
								End If
							Next
							Response.Write "</select>"
						Case "textarea":
							%>
							<textarea class=fmtxtra name="<%=ItemName%>" rows=<%=maxLength%> cols=34 <%If inputClass<>"" then%>style="width:<%=inputClass%>"<%End If%><%=moreEvent%>><%If ItemValue <> "" Then Response.Write VbCrLf & htmlEncode(ItemValue)%></textarea>
							<%
						Case "splitchecked":						
							Dim indexN,InfoText,tmpArr
							%>
							<ul class="splitchecked">
							<%
							for n = 0 to Ubound(SelectStr,1)
								If inStr(SelectStr(n),"|") <= 0 Then
									%>
									</ul><b><%=SelectStr(n)%></b><ul>
									<%
								Else
									tmpArr = Split(SelectStr(n),"|")
									IndexN = tmpArr(0)
									InfoText = tmpArr(1)
									If instr(InfoText,"<span") = 0 Then%>
									<li><span class="grayfont"><%
									If IndexN <= 9 Then Response.Write "0"
									Response.Write IndexN%></span><input type="checkbox" class=fmchkbox name="<%=ItemName%><%=IndexN%>" value="1"<%
									If instr(InfoText,"<span") Then Response.Write " disabled=""disabled"""
									If GetBinarybit(Item_Value,IndexN) = 1 Then
										Response.Write " checked>"
									Else
										Response.Write ">"
									End If%><%=InfoText%></li>
									<%
									End If
								End If
							Next
							%>
							</ul>
							<%
						case "other":
							Response.Write ItemName
					End Select
			If title <> "single" Then%>
				</span>
			</div>
<%
			end if

End Sub

Sub FormClass_End

%>
	
			<div class="itemline">
				<span class="itemtitle">&nbsp;</span>
				<span class="iteminfo itembottom">
					<input name=submit2 type=submit value="�ύ" class="fmbtn btn_2">
					<input name=b1 type=reset value="��д" class="fmbtn btn_2">
				</span>
			</div>
			</div>
		</form>
<%

End Sub

'Formitem Ҫ���Ե�ֵ checktype �������� default ����,�趨Ϊ��Ĭ��ֵ inValue ����İ���ֵ��
Dim CheckErrorStr
Function FormClass_CheckFormValue(Formitem,ItemName,checktype,default,inValue,maxlength)

	CheckErrorStr = ""
	Dim Tmp_Formitem
	Tmp_Formitem = Formitem
	Select Case checktype	
		Case "string":
			FormClass_CheckFormValue = Tmp_Formitem
		Case "numeric":
			if isNumeric(Tmp_Formitem) = 0 Then
				Tmp_Formitem = 0
				FormClass_CheckFormValue = Tmp_Formitem
				CheckErrorStr = CheckErrorStr & "Error: " & ItemName & "����1,��ȷ��."
			Else
				FormClass_CheckFormValue = cCur(Tmp_Formitem)
			End If
		Case "int":
			if isNumeric(Tmp_Formitem) = 0 or inStr(Tmp_Formitem,".") > 0 Then
				Tmp_Formitem = 0
				FormClass_CheckFormValue = Tmp_Formitem
				CheckErrorStr = CheckErrorStr & "Error: " & ItemName & "����2,��ȷ��."
			Else
				FormClass_CheckFormValue = Fix(cCur(Tmp_Formitem))
			End If
	End Select
	
	If maxlength > 0 Then
		If strLength(Tmp_Formitem) > maxlength Then
			If checktype = "int" Then
				CheckErrorStr = CheckErrorStr & "Error: " & ItemName & "����ֵ�������д����."
			Else
				CheckErrorStr = CheckErrorStr & "Error: " & ItemName & "����,���ܳ���" & maxlength & " �ֽ�."
			End If
		End If
	End If
	
	
	Dim Tmp,n,count,typeTmp,ValueTmp,tmp2
	If inValue <> "" and CheckErrorStr = "" Then
		If inStr(inValue,"~~~") > 0 Then
			Tmp = Split(inValue,"|")
			count = ubound(tmp)
			for n = 0 to count
				tmp2 = Split(tmp(n),"~~~")
				typeTmp = tmp2(0)
				ValueTmp = tmp2(1)
				select case typeTmp
					case ">":
						If cCur(Tmp_Formitem) > cCur(ValueTmp) Then
							CheckErrorStr = CheckErrorStr & "Error: " & ItemName & "����,��ȷ��."
						End If
					case "<":
						If cCur(Tmp_Formitem) < cCur(ValueTmp) Then
							CheckErrorStr = CheckErrorStr & "Error: " & ItemName & "����,��ȷ��."
						End If
					case "=":
						if ValueTmp = "" and Cstr(Tmp_Formitem) = Cstr(ValueTmp) Then
							CheckErrorStr = CheckErrorStr & "Error: " & ItemName & "������д,��ȷ��."
						Else
							If Cstr(Tmp_Formitem) = Cstr(ValueTmp) Then
								CheckErrorStr = CheckErrorStr & "Error: " & ItemName & "����5,��ȷ��."
							End If
						End If
				end select
				If CheckErrorStr <> "" Then exit for
			next
		Else
			if inStr("|" & inValue & "|","|" & Tmp_Formitem & "|") = 0 Then
				CheckErrorStr = CheckErrorStr & "Error: " & ItemName & "����6,��ȷ��."
			End if
		End If	
	End If
	
	If default <> "none" and CheckErrorStr <> "" Then
		FormClass_CheckFormValue = default
		CheckErrorStr = ""
	End If

End Function

Function cms_SQL(str)

	cms_SQL = Replace(str,"'","''")

End Function

sub cms_selectFormScript(url)
%>

		<script type="text/javascript">
		function a_command(cstr,obj,action)
		{
			layer_view(cstr,obj,'','','anc_delbody','<%=url%>','',1,'AjaxFlag=1&action2=' + action,1);return(false);
		}
		function delbody_view(obj,check)
		{
			layer_create("anc_msgbody");
			var tmp="";
			if(check==1)
			{
				tmp=" <a href=\"javascript:;\" onclick=\"a_command('����ͨ�����',$id('" + obj.id + "'),'check&idlist='+p_getselected());\">��ͨ�����</a>";
				tmp+=" <a href=\"javascript:;\" onclick=\"a_command('����ȡ�����',$id('" + obj.id + "'),'uncheck&idlist='+p_getselected());\">��ȡ�����</a>";
			}
				
			$id('anc_msgbody').innerHTML="<div class=ajaxbox>��ѡ�� <b id=layer_selectnum>" + p_getnum() + "</b> ����¼��<br>��ѡ��<b><a href=\"javascript:;\" onclick=\"a_command('ɾ����¼',$id('" + obj.id + "'),'delete&idlist='+p_getselected());\">����ɾ��</a>" + tmp + "</b><br><input class=\"fmchkbox\" type=\"checkbox\" name=\"selmsg\" id=\"selmsg\" value=\"1\" onclick=\"achoose();\" />ѡ��ȫ��</div>";
			layer_view('',obj,'','','anc_msgbody','','',0,'',0,20);
		}
		</script>
		<script src="../inc/js/p_list.js?ver=20090601.2" type="text/javascript"></script>
	<%	
end sub

function cms_checkdeleteform(table,superflag)
	
		dim cityid,checkLevelsql
		checkLevelsql = ""
		
		if superflag = 1 and Check_jdsupervisor = 0 then
			response.Write "<div class=ajaxbox><div class=cms_error>Ȩ�޲���.</div></div>"
			cms_checkdeleteform = 0
			exit function
		end if

		dim action2,idlist
		action2 = getformdata("action2")
		idlist = getformdata("idlist")
		if action2 <> "delete" then
			cms_checkdeleteform = 0
			exit function
		end if
		dim listtemp
		listtemp = ""
		idlist = split(idlist,",")
		dim n,count,val,sql
		count = ubound(idlist,1)
		for n = 0 to count
			val = idlist(n)
			val = FormClass_CheckFormValue(val,"","int",0,"",0)
			if val > 0 Then
				sql  = "delete from " & table & " where id=" & cms_SQL(val) & checkLevelsql
				listtemp = listtemp & val & ","
				call ldexecute(sql,1)
			end if
		next
		if listtemp = "" then listtemp = " ��(δѡ���κμ�¼)."
		response.Write "<div class=ajaxbox><div class=cms_ok>���¼�¼�ɹ�ɾ��: " & listtemp & "</div></div>"
		cms_checkdeleteform = 1

end function


function cms_changeCheckedFlag(table,superflag)
	
		dim cityid,checkLevelsql
		checkLevelsql = ""
		
		if superflag = 1 and Check_jdsupervisor = 0 then
			response.Write "<div class=ajaxbox><div class=cms_error>Ȩ�޲���.</div></div>"
			cms_changeCheckedFlag = 0
			exit function
		end if
		
		
		dim action2,idlist
		action2 = getformdata("action2")
		idlist = getformdata("idlist")
		if action2 <> "check" and action2 <> "uncheck" then
			cms_changeCheckedFlag = 0
			exit function
		end if
		if action2 = "check" Then
			action2 = 1
		else
			action2 = 0
		end if
		dim listtemp
		listtemp = ""
		idlist = split(idlist,",")
		dim n,count,val,sql
		count = ubound(idlist,1)
		for n = 0 to count
			val = idlist(n)
			val = FormClass_CheckFormValue(val,"","int",0,"",0)
			if val > 0 Then
				sql  = "update " & table & " set checkedflag=" & action2 & " where id=" & cms_SQL(val) & checkLevelsql
				listtemp = listtemp & val & ","
				call ldexecute(sql,1)
			end if
		next
		if listtemp = "" then listtemp = " ��(δѡ���κμ�¼)."
		response.Write "<div class=ajaxbox><div class=cms_ok>���¼�¼�ɹ�����: " & listtemp & "</div></div>"
		cms_changeCheckedFlag = 1
	
end function
%>