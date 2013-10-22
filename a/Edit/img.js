
var sAction = "INSERT";
var sTitle = "����";

var oControl;
var oSeletion;
var sRangeType;

var sFromUrl = "http://";
var sAlt = "";
var sBorder = "0";
var sBorderColor = "#000000";
var sFilter = "";
var sAlign = "";
var sWidth = "";
var sHeight = "";
var sVSpace = "";
var sHSpace = "";

var sCheckFlag = "file";

// ֻ������������
function IsDigit(e){
	var evt = window.event?window.event:e,target=evt.srcElement||evt.target;
  return ((evt.keyCode >= 48) && (evt.keyCode <= 57));
}

// ����������ֵ��ָ��ֵƥ�䣬��ѡ��ƥ����
function SearchSelectValue(o_Select, s_Value){
	for (var i=0;i<o_Select.length;i++){
		if (o_Select.options[i].value == s_Value){
			o_Select.selectedIndex = i;
			return true;
		}
	}
	return false;
}

// ȥ�ո�left,right,all��ѡ
function BaseTrim(str){
	  lIdx=0;rIdx=str.length;
	  if (BaseTrim.arguments.length==2)
	    act=BaseTrim.arguments[1].toLowerCase()
	  else
	    act="all"
      for(var i=0;i<str.length;i++){
	  	thelStr=str.substring(lIdx,lIdx+1)
		therStr=str.substring(rIdx,rIdx-1)
        if ((act=="all" || act=="left") && thelStr==" "){
			lIdx++
        }
        if ((act=="all" || act=="right") && therStr==" "){
			rIdx--
        }
      }
	  str=str.slice(lIdx,rIdx)
      return str
}

// תΪ�����ͣ�����ǰ��0������ת�򷵻�""
function ToInt(str){
	str=BaseTrim(str);
	if (str!=""){
		var sTemp=parseFloat(str);
		if (isNaN(sTemp)){
			str="";
		}else{
			str=sTemp;
		}
	}
	return str;
}

// ��ʼֵ
function editor_InitimgDocument(){

	sAction = "INSERT";
	sTitle = "����";
	
	sFromUrl = "http://";
	sAlt = "";
	sBorder = "0";
	sBorderColor = "#000000";
	sFilter = "";
	sAlign = "";
	sWidth = "";
	sHeight = "";
	sVSpace = "";
	sHSpace = "";
	
	sCheckFlag = "file";
	

	if(!isUndef(edt_doc) && edt_doc!=null)
	{
		if(Browser.ie)
		{
		oSelection = edt_doc.selection.createRange();
		sRangeType = edt_doc.selection.type;
		
		if (sRangeType == "Control") 
		{
			if (oSelection.item(0).tagName == "IMG")
			{
				sAction = "MODI";
				sTitle = "�޸�";
				sCheckFlag = "url";
				oControl = oSelection.item(0);
				sFromUrl = oControl.src;
				sAlt = oControl.alt;
				sBorder = oControl.border;
				sBorderColor = oControl.style.borderColor;
				sFilter = oControl.style.filter;
				sAlign = oControl.align;
				sWidth = oControl.width;
				sHeight = oControl.height;
				sVSpace = oControl.vspace;
				sHSpace = oControl.hspace;
			}
		}
		}
	
		SearchSelectValue($id("img_d_align"), sAlign.toLowerCase());
	
		$id("img_d_fromurl").value = sFromUrl;
		//img_d_alt.value = sAlt;
		$id("img_d_border").value = sBorder;
		//img_d_bordercolor.value = sBorderColor;
		//s_bordercolor.style.backgroundColor = sBorderColor;
		$id("img_d_width").value = sWidth;
		$id("img_d_height").value = sHeight;
		//img_d_vspace.value = sVSpace;
		//img_d_hspace.value = sHSpace;
	}
	else
	setTimeout("editor_InitimgDocument()",100);
}

editor_InitimgDocument();

// �����ڷ���ֵ
function editor_imgReturnValue(){
	sFromUrl = $id("img_d_fromurl").value;
	//sAlt = img_d_alt.value;
	sBorder = $id("img_d_border").value;
	//sBorderColor = img_d_bordercolor.value;
	//sFilter = img_d_filter.value;
	sAlign = $id("img_d_align").value;
	sWidth = $id("img_d_width").value;
	sHeight = $id("img_d_height").value;
	//sVSpace = img_d_vspace.value;
	//sHSpace = img_d_hspace.value;

	if (sAction == "MODI") {
		oControl.src = sFromUrl;
		//oControl.alt = sAlt;
		oControl.border = sBorder;
		//oControl.style.borderColor = sBorderColor;
		//oControl.style.filter = sFilter;
		oControl.align = sAlign;
		oControl.width = sWidth;
		oControl.height = sHeight;
		//oControl.vspace = sVSpace;
		//oControl.hspace = sHSpace;
	}else{
		if(sBorder=="")sBorder="0";
		if(sAlign=="")sAlign="absmiddle";
		if (sWidth!=""&&sHeight!="")
		{
			var sHTML = '<img src="'+sFromUrl+'" border="'+sBorder+'" align="'+sAlign+'" width="'+sWidth+'" height="'+sHeight+'">';
			var ubbHTML = "[IMG=" + sBorder + "," + sAlign + "," + sHeight + "," + sWidth + "]" + sFromUrl + "[/IMG]"
		}
		else
		{
			var sHTML = '<img src="'+sFromUrl+'" border="'+sBorder+'" align="'+sAlign+'">';
			var ubbHTML = "[IMG=" + sBorder + "," + sAlign + "]" + sFromUrl + "[/IMG]"
		}
		if(sFromUrl==""||sFromUrl=="http://")
		{
			alert("������ͼƬ��ַ.");
			return;
		}
		if(!edt_mode)
		{addcontent(2,sHTML);}
		else
		{addcontent(1,ubbHTML);}
	}
}

// ��ȷ��ʱִ��
function editor_imgok(){
	// �������������Ч��
	$id("img_d_border").value = ToInt($id("img_d_border").value);
	$id("img_d_width").value = ToInt($id("img_d_width").value);
	$id("img_d_height").value = ToInt($id("img_d_height").value);
	//img_d_vspace.value = ToInt(img_d_vspace.value);
	//img_d_hspace.value = ToInt(img_d_hspace.value);
	
	editor_imgReturnValue();
}