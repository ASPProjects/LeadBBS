<%
sub admanage_Main

	If Request.Form("subside") = "1" Then
		admanage_UpdateFormData
		Exit Sub
	End If

	Dim Side_Select
	Side_Select = Array("��ҳ-����","��ҳ-β��","���-����","���-β��","��������-����","��������-β��")
	
	Dim Side_Data,Dn
	Dim Rs
	Set Rs = LDExeCute("Select ID,RID,ValueStr,ClassNum,saveData from LeadBBS_Setup where RID=01004 order by ClassNum ASC",0)
	If Not Rs.Eof Then
		Side_Data = Rs.GetRows(-1)
		Dn = Ubound(Side_Data,2)
	Else
		Dn = -1
	End If
	Rs.Close
	Set Rs = Nothing
	
	Dim Sn,m,n
	Dim CheckFlag,Title,ClassNum,SaveData
	%>
	<script>
	function checksubmit(f)
	{
		var textarea = $('textarea').length;
		for(var n=0;n<textarea;n++)
		{
			if($('textarea').eq(n))
			{
				$("#test_html").html($('textarea').eq(n).val());
				$('textarea').eq(n).val($("#test_html").html());
				if($('textarea').eq(n).val().length>10240)
				{alert("���󣺴��볤�ȳ�����10240.");$('textarea').eq(n).focus();return false;}
			}
		}
		return true;
	}
	</script>
	<div id=test_html style="display:none;"></div>
	<p>
	<b>��̳���λ��Ͷ�Ŵ�������</b>
	</p>
	<br />
	<div class=grayfont>�����첽����˵����������Ĵ��뺬��script��ǩ��js�����ļ�������js�ڲ����б���ͬ�������document.write֮��Ĵ��룬�빴ѡ����
	<ul>
	<li>��1����������Ĺ����룬�й�վ��վ��ͳ��������������ڲ����� document.write��ͬ����������빴ѡ���������޷���ʾ�������</li>
	<li>��2���ٶȷ�������ڲ���ͬ���������Ҫ��ѡ����</li>
	<li>��3���������� js ���룬Ҳ����û��&lt;script src="...">֮��ı�ǩ�����ù�ѡ����</ul>
	</ul>
	</div>
	<div class="frameline greenfont">�ڿ��������Զ�����룬��������ʹ������HTML��JavaScript</div>
	<form action="SiteInfo.asp?action=admanage" method="post" name="LeadBBSFm" onSubmit="return checksubmit(this);"">
	<input type="hidden" value="1" name="subside">
	<div id="home_side_form">
	<%
	
	dim trueID
	Sn = Ubound(Side_Select,1)
	For m = 0 To Sn
		trueID = 9999
		CheckFlag = 1
		Title = "0"
		ClassNum = m
		SaveData = ""
		For n = 0 to Dn
			If ccur(Side_Data(3,m)) = m Then
				trueID = Side_Data(0,m)
				Title = Side_Data(2,m)
				If Title <> "0" and Title <> "1" Then Title = "0"
				SaveData = Side_Data(4,m)
				Exit For
			End If
		Next
		%>
			<hr class=splitline>
			<table border=0 cellpadding="0" class="blanktable">
			<tr>
			<td width=100 valign=top>
			<input type="hidden" name="trueID<%=m%>" value="<%=trueID%>">
			<b><%=Side_Select(m)%></b>
			<br />
			<br />
			<input type="checkbox" class=fmchkbox name="Side_Select<%=m%>" value="1"<%If Title = "1" Then
					Response.Write " checked>"
				Else
					Response.Write ">"
				End If%>�첽����
			</span>
			</td><td>
			<textarea cols="120" name="SaveData<%=m%>" rows="12" tabindex="51" class="fmtxtra"><%If SaveData <> "" Then Response.Write VbCrLf & htmlEncode(SaveData)%></textarea>
			</td></tr>
			</table>
		<%
	Next
	%>
	</div>
	<p>
	<input name=submit type=submit value="�ύ����" class="fmbtn">
	</p>
	</form>
	<%

end sub

sub admanage_UpdateFormData

	Dim Sn,m,Rs
	Dim CheckFlag,Title,SaveData,trueID,ClassNum
	For Sn = 0 to 5
		CheckFlag = Request.Form("Side_Select" & Sn)
		Title = Request.Form("Title" & Sn)
		SaveData = Request.Form("SaveData" & Sn)
		trueID = toNum(Request.Form("trueID" & Sn),0)
		If CheckFlag = "1" Then
			Title = "1"
		Else
			Title = "0"
		End If
		SaveData = Replace(Replace(Left(Replace(SaveData,"|",""),10240),"<" & "%","&lt;%"),"%" & ">","%&gt;")
			
		Set Rs = LDExeCute("Select * from LeadBBS_Setup where RID=01004 and ClassNum=" & Sn,0)
		If Not Rs.Eof Then
			CALL LDExeCute("Update LeadBBS_Setup Set ValueStr='" & Replace(Title,"'","''") & "',SaveData='" & Replace(SaveData,"'","''") & "' where RID=01004 and ClassNum=" & Sn,1)
		Else
			CALL LDExeCute("insert into LeadBBS_Setup(RID,ValueStr,ClassNum,SaveData) Values(01004,'" & Replace(Title,"'","''") & "'," & Sn & ",'" & Replace(SaveData,"'","''") & "')",1)
		End If
		Rs.Close
		Set Rs = Nothing
	Next
	admanage_UpdateFileData
	
	
end sub

Sub admanage_UpdateFileData

	
	Dim Side_Data,Dn
	Dim Rs
	Set Rs = LDExeCute("Select * from LeadBBS_Setup where RID=01004 order by ClassNum",0)
	If Not Rs.Eof Then
		Side_Data = Rs.GetRows(-1)
		Dn = Ubound(Side_Data,2)
	Else
		Dn = -1
	End If
	Rs.Close
	Set Rs = Nothing
	
	Dim m
	Dim Title,ClassNum,SaveData
	
	Dim Str
	Str = str & "function ad_start()" & VbCrLf
	Str = str & "{" & VbCrLf
	Str = str & "	var ad_select,ad_content;" & VbCrLf
	Str = str & "	var bbsad_html = new Array();" & VbCrLf
	Str = Str & "	var ad_idArray = ""ad_hometop|ad_homebottom|ad_boardtop|ad_boardbottom|ad_topictop|ad_topicbottom|bottom_ad"".split(""|"");" & VbCrLf
	Str = Str & "	var adsync_data=[];" & VbCrLf
	
	For m = 0 to Dn
		Title = Side_Data(2,m)
		SaveData = replace(replace(replace(replace(replace(replace(Side_Data(4,m),"\","\\"),"""","\"""),"script","s\x63ript"),VbCrLf,"\n"),chr(10),""),chr(13),"")
		Str = str & "bbsad_html[" & m & "] = """ & SaveData & """" & VbCrLf
		if Title = "1" Then
			Str = str & "adsync_data[" & m & "] = 1;" & VbCrLf
		Else
			Str = str & "adsync_data[" & m & "] = 0;" & VbCrLf
		End If
	Next

	Str = Str & "	if($id(""bottom_ad""))" & VbCrLf
	Str = Str & "	bbsad_html[6] = $id(""bottom_ad"").innerHTML.replace(/<!--/,"""").replace(/-->/,"""");" & VbCrLf
	Str = Str & "	else" & VbCrLf
	Str = Str & "	bbsad_html[6] = """";" & VbCrLf
	Str = Str & "	$id('bottom_ad').innerHTML="""";" & VbCrLf
	Str = Str & "	for(var i=0;i<=ad_idArray.length;i++)" & VbCrLf
	Str = Str & "	if($id(ad_idArray[i]))" & VbCrLf
	Str = Str & "	{" & VbCrLf
	Str = Str & "	ad_select = bbsad_html[i].split(""------leadbbs-split--------"");" & VbCrLf
	Str = Str & "	ad_content = ad_select[parseInt(Math.random()*ad_select.length)];" & VbCrLf
	Str = Str & "	if(adsync_data[i]==0)" & VbCrLf
	Str = Str & "	$('#'+ad_idArray[i]).html(ad_content);" & VbCrLf
	Str = Str & "	else" & VbCrLf
	Str = Str & "	$('#'+ad_idArray[i]).writeCapture().html(ad_content);" & VbCrLf
	Str = Str & "	if($.trim($('#'+ad_idArray[i]).html())!="""")$('#'+ad_idArray[i]).show();" & VbCrLf
	Str = Str & "	}" & VbCrLf
	Str = Str & "	}" & VbCrLf
	Str = Str & "	ad_start();" & VbCrLf
	
	CALL ADODB_SaveToFile(Str,DEF_BBS_HomeUrl & "inc/js/ad.js")
	Response.Write "<p>����������������. </p><p><a href=""SiteInfo.asp?action=admanage"">��˷�������</a></p>"

End Sub
%>