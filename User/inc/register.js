	function user_setface(newText)
	{
		$id("LeadBBSFm").Form_userphoto.value=newText;
		$id("faceimg").src=user_DEF_BBS_HomeUrl + 'images/face/'+newText+'.gif'
		layer_hidelayer($id('anc_delbody'));
	}
	var ValidationPassed = true;
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
			if (parseInt(temp) > 0 && parseInt(temp) <= user_DEF_faceMaxNum)
			{
				$id("faceimg").src=user_DEF_BBS_HomeUrl+'images/face/'+temp+'.gif';
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
			$id("faceimg").src=user_DEF_BBS_HomeUrl + 'images/blank.gif';
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
				alert("�Զ���ͷ���ȱ��������֣�\n");
				obj.Form_FaceWidth.focus();
				return;
			}
			else
			{
				if(obj.Form_FaceWidth.value<20 || obj.Form_FaceWidth.value>user_DEF_AllFaceMaxWidth)
				{
					alert("�Զ���ͷ���ȱ�����20-" + user_DEF_AllFaceMaxWidth + "֮�䣡\n");
					obj.Form_FaceWidth.focus();
					return;
				}
			}
		}

		if(obj.Form_FaceHeight.value!="")
		{
			if (! isnum(obj.Form_FaceHeight.value))
			{
				alert("�Զ���ͷ��߶ȱ��������֣�\n");
				obj.Form_FaceHeight.focus();
				return;
			}
			else
			{
				if(obj.Form_FaceHeight.value<20 || obj.Form_FaceHeight.value>user_DEF_AllFaceMaxWidth*2)
				{
					alert("�Զ���ͷ��߶ȱ�����20-" + user_DEF_AllFaceMaxWidth*2 + "֮�䣡\n");
					obj.Form_FaceHeight.focus();
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
		if(obj.Form_username.value=="")
		{
			alert("����������û���!\n");
			ValidationPassed = false;
			obj.Form_username.focus();
			return;
		}
		
		if(obj.Form_username.value.length<(user_DEF_ShortestUserName/2))
		{
			alert("�û�������������Ҫ" + user_DEF_ShortestUserName + "���ַ�!\n");
			ValidationPassed = false;
			obj.Form_username.focus();
			return;
		}

		if(obj.Form_mail.value=="")
		{
			alert("�����������ʵ�����ַ��\n");
			ValidationPassed = false;
			obj.Form_mail.focus();
			return;
		}

		if(obj.Form_password1.value=="")
		{
			alert("�������������!\n");
			ValidationPassed = false;
			obj.Form_password1.focus();
			return;
		}

		if(obj.Form_password2.value=="")
		{
			alert("�����������֤���룡\n");
			ValidationPassed = false;
			obj.Form_password2.focus();
			return;
		}

		if(obj.Form_password1.value!=obj.Form_password2.value)
		{
			alert("��������������벻��ͬ��\n");
			ValidationPassed = false;
			obj.Form_password1.focus();
			return;
		}

		if(obj.Form_Question.value=="")
		{
			alert("������������ʾ!\n");
			ValidationPassed = false;
			obj.Form_Question.focus();
			return;
		}

		if(obj.Form_Answer.value=="")
		{
			alert("��������ʾ��!\n");
			ValidationPassed = false;
			obj.Form_Answer.focus();
			return;
		}
		if(obj.Form_icq.value!="")
		{
			if (! isnum(obj.Form_icq.value))
			{
				alert("ι,��������ICQ���������˶���,�����ICQ������ô�������֣�\n");
				ValidationPassed = false;
				obj.Form_icq.focus();
				return;
			}
		}

		if(obj.Form_oicq.value!="")
		{
			if (! isnum(obj.Form_oicq.value))
			{
				alert("ι,�������ˣѣѿ��������˶���,����ģѣѺ�����ô��������?\n");
				ValidationPassed = false;
				obj.Form_oicq.focus();
				return;
			}
		}

		if(obj.Form_byear.value!="")
		{
			if (! isnum(obj.Form_byear.value))
			{
				alert("ι,����������ĳ�����,����������ô�������֣�\n");
				ValidationPassed = false;
				obj.Form_byear.focus();
				return;
			}
		}

		if(obj.Form_bmonth.value!="")
		{
			if (! isnum(obj.Form_bmonth.value))
			{
				alert("ι,����������ĳ�����,������·���ô�������֣�\n");
				ValidationPassed = false;
				obj.Form_bmonth.focus();
				return;
			}
		}

		if(obj.Form_bday.value!="")
		{
			if (! isnum(obj.Form_bday.value))
			{
				alert("ι,����������ĳ�����,����ĳ�������ô�������֣�\n");
				ValidationPassed = false;
				obj.Form_bday.focus();
				return;
			}
		}

		if(obj.Form_userphoto.value!="")
		{
			if (! isnum(obj.Form_bday.value))
			{
				alert("�û�ͼ��,ֻ����001-318֮������֣�\n");
				ValidationPassed = false;
				obj.Form_bday.focus();
				return;
			}
		}
		
		if(obj.Form_Underwrite.value.length>255)
		{
			alert("�û�ǩ������ҪС��255���ַ�!\n");
			ValidationPassed = false;
			obj.Form_Underwrite.focus();
			return;
		}
		if(user_DEF_AllDefineFace!=0 && user_DEF_AllDefineFace != 2)
		{
			if(obj.Form_FaceWidth.value!="")
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
	
			if(obj.Form_FaceHeight.value!="")
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
		
		if(user_ShowTestNumber>2)
		if(obj.ForumNumber.value=="")
		{
			alert("��������֤��!\n");
			ValidationPassed = false;
			obj.ForumNumber.focus();
			return;
		}
		ValidationPassed = true;
		return true;
	}

	function submitonce(theform)
	{
		form_onsubmit(theform);
		if(ValidationPassed == false)return;
		if (document.all||document.getElementById)
		{
			for (i=0;i<theform.length;i++)
			{
				var tempobj=theform.elements[i];
				if(tempobj.type.toLowerCase()=="submit"||tempobj.type.toLowerCase()=="reset")
				tempobj.disabled=true;
			}
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