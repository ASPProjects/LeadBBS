
// ֻ������������
function IsDigit(evt){
	//var evt = window.event?window.event:evt,target=evt.srcElement||evt.target;
  return ((evt.keyCode >= 48) && (evt.keyCode <= 57));
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

function editor_Initmedia()
{
	$id("media_d_fromurl").value = "http://";
	$id("media_d_width").value = "";
	$id("media_d_height").value = "";
}

// ��ȷ��ʱִ��
function editor_mediasubmit(){
	// �������������Ч��
	$id("media_d_width").value=ToInt($id("media_d_width").value);
	$id("media_d_height").value=ToInt($id("media_d_height").value);
	
	
	var sFromUrl = $id("media_d_fromurl").value;
	var sWidth = $id("media_d_width").value;
	var sHeight = $id("media_d_height").value;
	var sType = $id("media_d_type").options[$id("media_d_type").selectedIndex].value;
	
	if((sType=="FLV" || sType=="MP") && (sWidth=="" || sHeight==""))
	{
		alert("����FLV��Media�ļ�����ָ�����Ϳ�");
		return;
	}

	if(sFromUrl!="http://" && sFromUrl!="")
	{
		var sHTML = "[" + sType
		if (sWidth!="" && sHeight!="")
		{
			sHTML+="="+sWidth+","+sHeight+"]";
		}
		else{sHTML+="]";}
		sHTML+=sFromUrl+"[/" + sType + "]";
	
		addcontent(1,sHTML);
	}
	else
		alert("������ý���ַ");
}