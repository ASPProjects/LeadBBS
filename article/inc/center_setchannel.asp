<%
sub center_setchannel

		dim centersetchannelClass
		set centersetchannelClass = new center_setchannelClass_Class
		set centersetchannelClass = nothing
	
End sub


class center_setchannelClass_Class

	Private form_content,form_fileid,FileName
	Private form_type(16),form_title(16),form_listnum(16),form_id(16),form_extendflag(16),form_style(16)	
	Private StyleItem 
	
	rem type: 0.�г��������� 1.�г����¾��� 2.�г�ĳ(����)ר������ 3.�г�ĳ�������� 4.��������
	rem form_title: Ƶ������
	rem form_listnum: �г��ļ�¼����
	rem form_id: ר���Ż��ǰ���ı��
	rem form_extendflag: ������г���������(tppeΪ3)�˲�������Ч��0: �г�����ľ��� 1: �г��������������
	rem form_style: ��ʾ��ʽ,�����Ƿ�չʾ���ݻ���ͼƬ

	Private Sub Class_Initialize
	
		StyleItem = Array("1|����Ӵ�","2|չʾ������Ҫ","3|չʾ���ͼƬ","4|��ͼƬʱ���ر���","5|��������¼����Ӵ�չʾͼƬ","6|���ͼƬ��ʾΪ��ͼƬ")
		form_fileid = GetFormData("form_fileid")
		form_fileid = FormClass_CheckFormValue(form_fileid,"","int",0,"",0)
		select case form_fileid
			case 0:
				FileName = "inc/home_channellist.asp"
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
	
	
	rem Private form_type,form_title,form_listnum,form_id,form_extendflag,form_style
	
	rem type: 999.�����ͣ��ر�״̬ 0.�г��������� 1.�г����¾��� 2.�г�ĳ(����)ר������ 3.�г�ĳ�������� 
	rem form_title: Ƶ������
	rem form_listnum: �г��ļ�¼����
	rem form_id: ר���Ż��ǰ���ı��
	rem form_extendflag: ������г���������(tppeΪ3)�˲�������Ч��0: �г�����ľ��� 1: �г��������������
		dim indexn
		form_content = ""
		Dim Temp2,TempN,N
		for indexn = 0 to 15
			form_type(indexn) = GetFormData("form_type" & indexn)
			select case form_type(indexn)
				case "0": form_type(indexn) = 0
					     form_id(indexn) = 0
					     form_extendflag(indexn) = 0
				case "1": form_type(indexn) = 1
					     form_id(indexn) = 0
					     form_extendflag(indexn) = 0
				case "2": form_type(indexn) = 2
					     form_id(indexn) = GetFormData("form_id" & indexn)
					     form_extendflag(indexn) = 0
				case "3": form_type(indexn) = 3
					     form_id(indexn) = GetFormData("form_id" & indexn)
					     form_extendflag(indexn) = GetFormData("form_extendflag" & indexn)
				case "4": form_type(indexn) = 4
					     form_id(indexn) = GetFormData("form_id" & indexn)
					     form_extendflag(indexn) = 0
				case else
					     form_type(indexn) = 999
					     form_id(indexn) = 0
					     form_extendflag(indexn) = 0
			end select
			form_id(indexn) = FormClass_CheckFormValue(form_id(indexn),"","int",0,"",0)
			form_extendflag(indexn) = FormClass_CheckFormValue(form_extendflag(indexn),"","int",0,"",0)
			form_title(indexn) = lefttrue(GetFormData("form_title" & indexn),1024)
			form_listnum(indexn) = GetFormData("form_listnum" & indexn)
			form_listnum(indexn) = FormClass_CheckFormValue(form_listnum(indexn),"","int",0,"",0)
			
			
			form_style(indexn) = 0
			Temp2 = 1
			For TempN = 0 to Ubound(StyleItem,1)
				N = Request("form_style" & indexn & TempN+1)
				If N <> "1" Then N = "0"
				If N = "1" Then form_style(indexn) = form_style(indexn)+cCur(Temp2)
				Temp2 = Temp2*2
			Next
	
			form_title(indexn) = replace(form_title(indexn),"<" & "%","")
			form_title(indexn) = replace(form_title(indexn),"%" & ">","")
			form_title(indexn) = replace(form_title(indexn),"<" & "script","")
			form_title(indexn) = replace(form_title(indexn),"</" & "script","")
			form_title(indexn) = replace(form_title(indexn),chr(10),"")
			form_title(indexn) = replace(form_title(indexn),chr(13),"")
			form_title(indexn) = replace(form_title(indexn),"#~#^#","")
			if indexn > 0 then form_content = form_content & VbCrLf
			form_content = form_content & form_type(indexn) & "#~#^#" & form_title(indexn) & "#~#^#" & form_listnum(indexn) & "#~#^#" & form_id(indexn) & "#~#^#" & form_extendflag(indexn) & "#~#^#" & form_style(indexn)
		next
		
		private_Saveformdata
		CALL Update_InsertSetupRID(1051,"article/" & FileName,8,form_content," and ClassNum=" & 8)
	
	End Sub
	
	private sub private_Saveformdata
	
		ADODB_SaveToFile form_content,FileName
		Response.Write "<span class=cms_ok>�ɹ��༭��ҳ��Ŀ��Ϣ.</span>"

	End Sub
	
	private function private_getClassinfo
	
		form_content = ADODB_LoadFile(FileName)
		dim tmp,n,tmp2
		tmp = split(form_content,VbCrLf)
		for n = 0 to ubound(tmp)
			tmp2 = split(tmp(n),"#~#^#")
			form_style(n) = 0
			if ubound(tmp2) >= 4 then
				form_type(n) = tmp2(0)
				form_title(n) = tmp2(1)
				form_listnum(n) = tmp2(2)
				form_id(n) = tmp2(3)
				form_extendflag(n) = tmp2(4)
				if ubound(tmp2) >= 5 then form_style(n) = tmp2(5)
			else
				form_type(n) = 999
				form_title(n) = "null"
				form_listnum(n) = 0
				form_id(n) = 0
				form_extendflag(n) = 0
			end if
		next
		
	end function
	
	Public Sub center_Class_Form
	%>
	<div id=testttt></div>
		<script>
		$(document).ready(function(){
		$("select.itemtype").combobox({
		onChange: function (n,o){
		checkselect(this);
		//alert($(this).change);
		}
		});
		}); 

		function checkselect(obj,clickflag)
		{
			var sel = $(obj).parent().parent();
			var numitem = $(sel).next(".itemline").next(".itemline").next(".itemline").find('.itemtitle');
			if(clickflag!=-1){
			}
			switch($(obj).combobox('getValue'))
			{
				case "999":
					$(sel).next(".itemline").hide().next(".itemline").hide().next(".itemline").hide().next(".itemline").hide().next(".itemline").hide();
					$(numitem).html("��ر��");
					break;
				case "0":
					$(sel).next(".itemline").show().next(".itemline").show().next(".itemline").hide().next(".itemline").hide().next(".itemline").show();
					$(numitem).html("��ر��");
					break;
				case "1":
					$(sel).next(".itemline").show().next(".itemline").show().next(".itemline").hide().next(".itemline").hide().next(".itemline").show();
					$(numitem).html("��ر��");
					break;
				case "2":
					$(sel).next(".itemline").show().next(".itemline").show().next(".itemline").show().next(".itemline").hide().next(".itemline").show();
					$(numitem).html("ר����");
					if(clickflag!=-1)
					{
					var index = $(obj).attr("comboname").replace(/form_type/,"");
					$('#form_id'+index).combobox('reload', '<%=DEF_BBS_HomeUrl%>inc/inchtm/data_goodassort.asp')
					}
					break;
				case "3":
					$(sel).next(".itemline").show().next(".itemline").show().next(".itemline").show().next(".itemline").show().next(".itemline").show();
					$(numitem).html("�����");
					if(clickflag!=-1)
					{
					var index = $(obj).attr("comboname").replace(/form_type/,"");
					$('#form_id'+index).combobox('reload', '<%=DEF_BBS_HomeUrl%>inc/inchtm/data_boardlist.asp?1')
					}
					break;
				case "4":
					$(sel).next(".itemline").show().next(".itemline").show().next(".itemline").show().next(".itemline").hide().next(".itemline").show();
					$(numitem).html("���·�����");
					if(clickflag!=-1)
					{
					var index = $(obj).attr("comboname").replace(/form_type/,"");
					$('#form_id'+index).combobox('reload', '<%=DEF_BBS_HomeUrl%>inc/inchtm/data_artileclass.asp')
					}
					break;
				default:					
			}
		}
		function initItem()
		{
			//return;
			var arr = $(".itemline select.itemtype");
			for(var n=0;n<arr.length;n++)
			{
			checkselect(arr[n],-1);
			}
		}
		$(document).ready(function() {initItem();});
		function formatItem(row){
			if(row.id==0)
			return('<span style="font-weight:bold" class="grayfont">' + row.text + '</span>');
			else
			return(row.text);
		}
		</script>
		<b>˵��: </b>
		<ol>
		<li>ר�������ţ�����Ŀ����Ϊר�⣬���������дר���ţ����Ϊ���棬����д�����ţ���Ϊ���£�����д��Ӧ�����·����š�������Ժ���</li>
		<li>����ѡ�����ֻ�е���Ŀ����Ϊ����ʱ����Ч</li>
		</ol>
		<hr class=splitline>
		<div class="definehome">
		<%
		dim n
		CALL FormClass_Head(Form_ActionStr,0,"center.asp?action=setchannel")
		CALL FormClass_ItemPring("","hidden","form_fileid",form_fileid,"","","","","")
		CALL FormClass_ItemPring("","hidden","submitflag","yes","","","","","")
		
		dim str,datafile,arr
		for n = 0 to 15
			response.Write "<div class='itemline'><span class='iteminfo'><b>��Ŀ" & n+1 & "</b></span></div>"
			CALL FormClass_ItemPring("����","select","form_type" & n,form_type(n),"","","","999~~~��-�رմ���|0~~~�г���̳��������|1~~~�г���̳���¾���|2~~~�г�ĳ(����)ר������|3~~~�г�ĳһ�����������|4~~~�г�ĳ���·����µ����� "," class=""easyui-combobox itemtype""")
			if form_title(n) = "null" then form_title(n) = ""
			CALL FormClass_ItemPring("����","input","form_title" & n,form_title(n),4,255,"Ϊ������Ŀ�б��������,����Ϊ��","","")
			CALL FormClass_ItemPring("��ʾ����","input","form_listnum" & n,form_listnum(n),2,2,"����Ŀ��ʾ������¼����","","")
			'CALL FormClass_ItemPring("��ر��","input","form_id" & n,form_id(n),2,12,"","","")
			
			select case form_type(n)
				case 2:
					datafile = "data_goodassort.asp"
				case 3:
					datafile = "data_boardlist.asp"
				case 4:
					datafile = "data_artileclass.asp"
			end select
			str = "<input class=""easyui-combobox"" id=""form_id" & n & """ name=""form_id" & n & """ value=""" & form_id(n) & """ data-options=""" & VbCrLf &_
					"	url: '" & DEF_BBS_HomeUrl & "inc/inchtm/" & datafile & "'," & VbCrLf &_
					"	valueField: 'id'," & VbCrLf &_
					"	textField: 'text'," & VbCrLf &_
					"	panelWidth: 250," & VbCrLf &_
					"	panelHeight: 'auto'," & VbCrLf &_
					"	formatter: formatItem" & VbCrLf &_
					""">" & VbCrLf
			'str = "<input class=""easyui-combobox"" id=""form_id" & n & """ name=""form_id" & n & """ data-options=""valueField:'id',textField:'text'"">"
			CALL FormClass_ItemPring("��ر��","other",str,form_id(n),2,12,"","","")
			CALL FormClass_ItemPring("����ѡ��","select","form_extendflag" & n,form_extendflag(n),"","","","0~~~�������������|1~~~������¾�����"," class=""easyui-combobox""")
			CALL FormClass_ItemPring("���ƽ���","splitchecked","form_style" & n,form_style(n),"","","",StyleItem,"")
			response.Write "<div class='itemline'><span class='iteminfo'><hr class=""splitline"" /></span></div>"
		next
		FormClass_End
		%>
		</div>
		<br /><br />
		<%
	
	End Sub
	
End Class

%>