	function user_setface(newText)
	{
		$id("LeadBBSFm").Form_userphoto.value=newText;
		$id("faceimg").src=user_DEF_BBS_HomeUrl + 'images/face/'+newText+'.gif'
		layer_hidelayer($id('anc_delbody'));
	}
	ValidationPassed = true;
	function isnum(str)
	{
		rset="";
		for(i=0;i<str.length;i++)
		{
			if(str.charAt(i)>="0" && str.charAt(i)<="9")
			{
			}
			else
			{
				return 0;
			}
		}
		return 1;
	}

	function changeface()
	{
		var temp;
		temp=$id("LeadBBSFm").Form_userphoto.value;
		if (temp!="" && isnum(temp)==1 && temp.length==4)
		{
			if (temp > 0 && temp <= user_DEF_faceMaxNum)
			{
				$id("faceimg").src=user_DEF_BBS_HomeUrl + 'images/face/'+temp+'.gif';
			}
			else
			{
				alert("����!��ͼ����Ų�����!");
				$id("faceimg").src=user_DEF_BBS_HomeUrl + 'images/blank.gif';
				$id("LeadBBSFm").Form_userphoto.value='';
				ValidationPassed = false;
			}
		}
		else
		{
			alert("����!��ͼ����Ų�����!\nͼ����ű�����4λ��" + (user_DEF_faceMaxNum.toString().length>4?"������":"") + ",���� 0001 ,���Ϊ" + user_DEF_faceMaxNum);
			$id("faceimg").src=user_DEF_BBS_HomeUrl+'images/blank.gif';
			$id("LeadBBSFm").Form_userphoto.value='';
			ValidationPassed = false;
		}
	}
	function changeface2()
	{
		var temp,obj;
		obj=$id("LeadBBSFm");
		if(obj.Form_FaceWidth.value!="")
		{
			if (! isnum(obj.Form_FaceWidth.value))
			{
				obj.Form_FaceWidth.value = 120;
				return;
			}
			else
			{
				if(obj.Form_FaceWidth.value<20 || obj.Form_FaceWidth.value>user_DEF_AllFaceMaxWidth)
				{
					obj.Form_FaceWidth.value = 120;
					return;
				}
			}
		}

		if(obj.Form_FaceHeight.value!="")
		{
			if (! isnum(obj.Form_FaceHeight.value))
			{
				obj.Form_FaceHeight.value = 120;
				return;
			}
			else
			{
				if(obj.Form_FaceHeight.value<20 || obj.Form_FaceHeight.value>user_DEF_AllFaceMaxWidth*2)
				{
					obj.Form_FaceHeight.value = 120;
					return;
				}
			}
		}

		temp=$id("LeadBBSFm").Form_FaceUrl.value;
		if (temp!="")
		{
			$id("faceimg").src=temp;
			$id("faceimg").width=obj.Form_FaceWidth.value;
			$id("faceimg").height=obj.Form_FaceHeight.value;
		}
	}

	function form_onsubmit(obj)
	{
		if(obj.file)
		{
			if(obj.upload_step&&obj.upload_step.value!=""){}
		else
			{$('#file').attr("disabled",true);}
		}
		if(obj.oldpass&&obj.oldpass.value=="")
		{
			alert("������д�ɵ�����!\n");
			ValidationPassed = false;
			obj.oldpass.focus();
			return;
		}

		if(obj.Form_password1&&obj.Form_password1.value!=obj.Form_password2.value)
		{
			alert("��������������벻��ͬ��\n");
			ValidationPassed = false;
			obj.Form_password1.focus();
			return;
		}

		if(obj.Form_icq&&obj.Form_icq.value!="")
		{
			if (! isnum(obj.Form_icq.value))
			{
				alert("ι,��������ICQ���������˶���,�����ICQ������ô�������֣�\n");
				ValidationPassed = false;
				obj.Form_icq.focus();
				return;
			}
		}

		if(obj.Form_oicq&&obj.Form_oicq.value!="")
		{
			if (! isnum(obj.Form_oicq.value))
			{
				alert("ι,�������ˣѣѿ��������˶���,����ģѣѺ�����ô��������?\n");
				ValidationPassed = false;
				obj.Form_oicq.focus();
				return;
			}
		}

		if(obj.Form_byear&&obj.Form_byear.value!="")
		{
			if (! isnum(obj.Form_byear.value))
			{
				alert("ι,����������ĳ�����,����������ô�������֣�\n");
				ValidationPassed = false;
				obj.Form_byear.focus();
				return;
			}
		}

		if(obj.Form_bmonth&&obj.Form_bmonth.value!="")
		{
			if (! isnum(obj.Form_bmonth.value))
			{
				alert("ι,����������ĳ�����,������·���ô�������֣�\n");
				ValidationPassed = false;
				obj.Form_bmonth.focus();
				return;
			}
		}

		if(obj.Form_bday&&obj.Form_bday.value!="")
		{
			if (! isnum(obj.Form_bday.value))
			{
				alert("ι,����������ĳ�����,����ĳ�������ô�������֣�\n");
				ValidationPassed = false;
				obj.Form_bday.focus();
				return;
			}
		}

		if(obj.Form_userphoto&&obj.Form_userphoto.value!="")
		{
			if (obj.Form_bday&&! isnum(obj.Form_bday.value))
			{
				alert("�û�ͼ��,ֻ����0001-0318֮������֣�\n");
				ValidationPassed = false;
				obj.Form_bday.focus();
				return;
			}
		}
		
		if(obj.Form_Underwrite&&obj.Form_Underwrite.value.length>255)
		{
			alert("�û�ǩ������ҪС��255���ַ�!\n");
			ValidationPassed = false;
			obj.Form_Underwrite.focus();
			return;
		}

		if(user_DEF_AllDefineFace!=0 && user_DEF_AllDefineFace != 2)
		{
			if(obj.Form_FaceWidth&&obj.Form_FaceWidth.value!="")
			{
				if (! isnum(obj.Form_FaceWidth.value))
				{
					alert("�Զ���ͷ���ȱ��������֣�\n");
					ValidationPassed = false;
					obj.Form_FaceWidth.focus();
					return;
				}
				else
				{
					if(obj.Form_FaceWidth.value<20 || obj.Form_FaceWidth.value>user_DEF_AllFaceMaxWidth)
					{
						alert("�Զ���ͷ���ȱ�����20-" + user_DEF_AllFaceMaxWidth + "֮�䣡\n");
						ValidationPassed = false;
						obj.Form_FaceWidth.focus();
						return;
					}
				}
			}
	
			if(obj.Form_FaceHeight&&obj.Form_FaceHeight.value!="")
			{
				if (! isnum(obj.Form_FaceHeight.value))
				{
					alert("�Զ���ͷ��߶ȱ��������֣�\n");
					ValidationPassed = false;
					obj.Form_FaceHeight.focus();
					return;
				}
				else
				{
					if(obj.Form_FaceHeight.value<20 || obj.Form_FaceHeight.value>user_DEF_AllFaceMaxWidth*2)
					{
						alert("�Զ���ͷ��߶ȱ�����20-" + user_DEF_AllFaceMaxWidth*2 + "֮�䣡\n");
						ValidationPassed = false;
						obj.Form_FaceHeight.focus();
						return;
					}
				}
			}
		}
		
		ValidationPassed = true;
		return true;
	}

	function submitonce(theform)
	{
		if(ValidationPassed == false)return;
		submit_disable(theform);
	}
	

function init_uploadform()
{
	var upload = $$("fminpt uninit_upload","input"),fun,dis;
	for (i = 0; i != upload.length; i++)
	{
		if(upload[i].id==""){upload[i].id="upload_id_" + LD.mnu_n;LD.mnu_n++;}
		upload[i].style.cssText = "*margin-left:-6px;filter:alpha(opacity=0);-moz-opacity:0.0;opacity:0.0; cursor:pointer;";
		upload[i].parentNode.className=upload[i].className="btn_upload";
		upload[i].onclick=function(){if(this.childNodes[0])this.childNodes[0].click();};
		upload[i].parentNode.style.cssText="vertical-align:middle;display: inline-block;*display: inline;zoom:1;";
		fun = upload[i].onchange;
		if(fun!=""&&fun!=null)
		fun = fun.toString().replace(/[\s\r\n]*function[\s\r\n]*[a-z]*[\s]*\([a-z\s\n\r\s]*\)[\s\n\r\s]*\{([\\\s\\\S]*?)\}[\s\n\r\s]*/gim,"$1");
		else
		fun="";
		upload[i].onchange="function(){$id(this.id+\"_inpt\").value = this.value;" + fun + "};"
		setTimeout("$id('" + upload[i].id + "').onchange=function(){$id(this.id+\"_inpt\").value = this.value;" + fun + "};",1);
		var iobj = document.createElement('input');
		iobj.className = "fminpt input_2";
		iobj.type = "text";
		iobj.readonly = "readonly";
		iobj.name = "_file9385a6";
		iobj.id = upload[i].id + "_inpt";
		upload[i].parentNode.parentNode.insertBefore(iobj, upload[i].parentNode);
	}
}


	function reg_checkinfo(item,str)
	{
		if(str.replace(/(^\s*)|(\s*$)/g,"")==""||str=="")
		{
			$id("reg_check_" + item).innerHTML="<span class=redfont>���������д��</span>"
			return;
		}
		getAJAX(user_DEF_RegisterFile,"checkflag=1&checkitem=" + item + "&checkvalue=" + escape(str),"reg_check_" + item);
	}