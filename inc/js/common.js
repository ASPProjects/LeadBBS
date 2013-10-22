var fadeflag = 1,HU="../";
var layer_loadstr = "<div class=\"ajaxloading\"></div>";

function jsloading()
{
	var cssText = "\n select { display: none;}";
	try
	{
		var styleSheet = document.createStyleSheet();
		styleSheet.cssText = cssText;
	}
	catch (e)
	{
		var style = document.createElement("style");
		style.type = "text/css";
		document.getElementsByTagName("head").item(0).appendChild(style);
	}
}

function isUndef(obj)
{
	return (typeof obj == 'undefined' || obj==null)? true : false;
}
//menu layer
var Browser = {
	a : navigator.userAgent.toLowerCase()
}
Browser = {
	ie : /*@cc_on!@*/false,
	ie6 : Browser.a.indexOf('msie 6') != -1,
	ie7 : Browser.a.indexOf('msie 7') != -1,
	ie8 : Browser.a.indexOf('msie 8') != -1,
	ie9 : Browser.a.indexOf('msie 9') != -1,
	ff : Browser.a.indexOf('firefox') != -1,
	opera : !!window.opera,
	safari : Browser.a.indexOf('safari') != -1,
	safari3 : Browser.a.indexOf('applewebkit/5') != -1,
	mac : Browser.a.indexOf('mac') != -1,
	is_ie: false,
	is_ie_lower: false
}

Browser.is_ie = (Browser.ie6 || Browser.ie7 || Browser.ie8 || Browser.ie9);
Browser.is_ie_lower = (!Browser.is_ie && Browser.ie);
jsloading();

function page_preexecuted()
{
	try {
	document.execCommand('BackgroundImageCache', false, true);
	}
	catch(e) {}
}
if(Browser.ie6||Browser.is_ie_lower)page_preexecuted();

$id = function(id){
	return document.getElementById(id);
}

$$ = document.getElementsByClassName = function(classname,tag){
	var elems,alls,ta;
	if(tag!="undefined"&&!isUndef(tag)&&tag!=null&&tag!="")
	ta = tag;
	else
	ta = "div"
	elems = [];
	alls = document.getElementsByTagName(ta); 
	for(var i=0; i<alls.length; i++){
		if(alls[i].className && alls[i].className == classname){
			elems.push(alls[i]);
		}
	}
	if(ta==tag)return elems;
	alls = document.getElementsByTagName('span'); 
	for(var i=0; i<alls.length; i++){
		if(alls[i].className && alls[i].className == classname){
			elems.push(alls[i]);
		}
	}
	return elems;
}

LayerMenu = function(menus, layers,tp,atitle){
	this.menus = $$(menus);
	if(layers!="none"){this.layers = $$(layers);if(menus=="layer_item2")this.atitle = $$("layer_item_title");}
	if(tp=="option")this.option = $$("layer_option_title");
	this.addEvent(tp);
}

var LD = {
	mnu_curele:"",
	mun_clickcur:"",
	mun_oldclickcur:"",
	cflag:1,

	mnu_n:1,
	clearOldLayer: function()
	{	
		if(LD.mun_oldclickcur&&LD.mun_oldclickcur!="")LD.mun_oldclickcur.style.display="none";
		if(LD.mnu_curele.info&&LD.mnu_curele.info!="")LD.mnu_curele.info.style.display="none";
	},
	hide: function(d)
	{
		var i=$id(d);
		if(i)
		{
			LD.clearOldLayer();
			i.style.display="none";
		}
	},
	
	curSty: function(a)
		{
			if(a.currentStyle)
			return(a.currentStyle);
			else
			return(document.defaultView.getComputedStyle(a,null));
			//Browsers
		},

	classstr : "|layer_item|layer_item2|layer_ajaxitem|layer_option|light_box|layer_onclickitem|",
	checkabs: function(o,s)
	{
		var eee=o.parentNode;
		var layerflag = 0;
		while(eee&&s!="abs") {
			if(LD.classstr.indexOf("|" + eee.className + "|")!=-1)
			{layerflag=1;break;}
			eee=eee.parentNode;
		}
		var ar= new Array(layerflag,eee) 
		return ar;
	},
	getX: function(e,s) {
		var ee = e;
		var x = 0;
		var layerflag = LD.checkabs(e,s)[0];

		while(ee) {
			if(s!="abs")
			if(layerflag==1)
			if(LD.curSty(ee).position=="absolute")return(x);
			x += ee.offsetLeft;
			ee = ee.offsetParent;
		}
		return x;
	},

	getY: function(e,s) {
		var ee = e;
		var h = 0;
		var layerflag = LD.checkabs(e,s)[0];

		while(ee) {
			if(s!="abs")
			if(layerflag==1)
			if(LD.curSty(ee).position=="absolute")return(h);
			h += parseInt(ee.offsetTop);
			ee = ee.offsetParent;
		}
		return h;
	},
	
	getMX: function(e) {
		var evt = window.event?window.event:e;
		var x = evt.pageX;
		if (!x && 0 !== x)
		{
			x = evt.clientX || 0;
			if ( Browser.ie )
			{
				x += (document.compatMode && document.compatMode!="BackCompat") ? document.documentElement.scrollLeft : document.body.scrollLeft;
			}
		}
		
		var elem = evt.target || evt.srcElement;
		if(elem)
		{
			var lflag=LD.checkabs(elem,"");
			if(lflag[0]==1)return(x-LD.getX(lflag[1],"abs"));
		}
		return x;
	},
	
	getMY: function(e) {
		var evt = window.event?window.event:e;
		var y = evt.pageY;
		if (!y && 0 !== y)
		{
			y = evt.clientY || 0;
			if ( Browser.ie )
			{
				y += (document.compatMode && document.compatMode!="BackCompat") ? document.documentElement.scrollTop : document.body.scrollTop;
			}
		}
		
		
		var elem = evt.target || evt.srcElement;
		if(elem)
		{
			var lflag=LD.checkabs(elem,"");
			if(lflag[0]==1)return(y-LD.getY(lflag[1],"abs"));
		}
		return y;
	},
	
	mnu_hide : function(obj)
	{
		var iobj = $id(obj);
		//alert(iobj.info.style.display);
		if(iobj.id!=LD.mnu_curele.id || isUndef(LD.mnu_curele.id))
		{
			iobj.info.style.display = 'none';
		}
		else
		{
			//iobj.info.style.display = 'none';
			if(iobj.info.style.display != 'none')setTimeout("(function(){LD.mnu_hide('"+obj+"')})()",3500);
		}
	},

	setpos: function(obj,menuobj,pos_x,tp,event)
	{
		var x,y,oldx,oldy;
		var posx = (!isUndef(pos_x)&&pos_x!="undefined")?pos_x:0;
		var promp = (isUndef(tp) || tp=="undefined")?false:true;
		var promptflag = (!promp&&obj!=""&&obj!=null)?true:false;
		var absstr = (LD.checkabs(menuobj,"")[0]==1)?"":"abs";
		if(promptflag)
		{
		x = LD.getX(obj,absstr);
		y = LD.getY(obj,absstr)-1;
		oldx = x,oldy = y;
		var distHeight = obj.clientHeight;
		distWidth = 0;
		
		var tmp = obj.parentNode,ajaxitem=0,tmpHeight,tmp2;
		while(tmp)
		{
			if(tmp.className == "ajaxitembody")tmp2 = tmp;
			if(tmp.className == "layer_ajaxitem")
			{
				ajaxitem = 1;
				tmpHeight = tmp.scrollHeight.toString().replace(/px/g,'');
				break;
			}
			tmp = tmp.parentNode;
		}
		
		if(!distHeight)distHeight+=obj.offsetHeight;
		x += distWidth;
		y += distHeight;
		x += parseInt(posx);
		}
		else
		{			
		x = LD.getMX(event);
		y = LD.getMY(event);
		oldx = x,oldy = y;
		if(posx==-1314||posx==-1315)
		{x -= parseInt(menuobj.offsetWidth/2);
		if(posx==-1314)y -= parseInt(menuobj.offsetHeight/2);
		}
		else
		{x += parseInt(posx);y += parseInt(posx);}
		}
		menuobj.style.top = y + 'px';
		menuobj.style.left = x + 'px';
		if(LD.cflag==1||(menuobj.className!="layer_optioninfo"&&menuobj.className!="layer_iteminfo"))menuobj.style.display = 'block';

		if (document.compatMode == "BackCompat")
		{
		W = document.body.clientWidth;
		H = document.body.clientHeight;
		WS = document.body.scrollLeft;
		HS = document.body.scrollTop;
		}
		else { //CSS1Compat
		W = document.documentElement.clientWidth;
		H = document.documentElement.clientHeight;
		WS = document.documentElement.scrollLeft == 0 ? document.body.scrollLeft : document.documentElement.scrollLeft;
		HS = document.documentElement.scrollTop == 0 ? document.body.scrollTop : document.documentElement.scrollTop;
		}
		var thisY = LD.getY(menuobj,"abs");
		var topY = thisY - HS;
		var bottomY = (HS+H-y);
		
		var thisX = LD.getX(menuobj,"abs");
		var thisX2 = thisX + menuobj.offsetWidth;
		var headX = thisX - WS;
		var bottomX = WS+W-thisX;
		var viewMaxX = WS + W;
		if(!Browser.is_ie_lower && bottomY<menuobj.offsetHeight && topY>bottomY)
		{
			y = !promptflag?LD.getMY(event)-((posx==-1315||posx==-1314)?0:posx)-menuobj.offsetHeight+1:LD.getY(obj,absstr,1)-menuobj.offsetHeight+2;
			menuobj.style.top = y + 'px';
		}
		if(thisX < WS) 
		{
			x = WS + 1;
			menuobj.style.left = x + 'px';
		}
		else if(!Browser.is_ie_lower && thisX2 > viewMaxX && !(posx==-1314||posx==-1315))
		{
			x = viewMaxX-(menuobj.offsetWidth);
			menuobj.style.left = x + 'px';
		}
		
		if(ajaxitem==1)
		{
			if(!isNaN(tmpHeight))
			{
				if(parseInt(tmpHeight)<menuobj.offsetHeight+135)
				{
					if (menuobj.offsetHeight>200)
					{menuobj.style.height="200px";
					tmp2.style.height=(200+135)+"px";
					}
					else
					tmp2.style.height=(menuobj.offsetHeight+135)+"px";
				}
			}
		}
	},
	stope:function(evt)
	{
		e = evt?evt:window.event;
		if(Browser.ie) {
			e.returnValue = false;
			e.cancelBubble = true;
		} else if(e) {
			e.stopPropagation();
			e.preventDefault();
		}
	},
	Cookie: {}

}

LayerMenu.prototype.addEvent = function(tp){
	for(var i=0; i<this.menus.length; i++){
		var item = this.menus[i];
		if(this.layers)
		{
			if(item.id=="")item.id = "menu_" + LD.mnu_n;
			LD.mnu_n++;
			item.info = this.layers[i];
			if(this.atitle)item.atitle = this.atitle[i];
			if(item.info.id=="")item.info.id = "menu_" + LD.mnu_n + "_info";
			item.type = tp;
			if(item.info.onclick==null)
			{	
				item.onmouseover = this.showInfo;
				item.onmouseout = this.hiddenInfo;
				if(tp=="prompt")item.onmousemove = this.showInfo;
			}
			else
			{
				var tmpp = item.getElementsByTagName("div")
				for(var ii=0; ii<tmpp.length; ii++){
					if(tmpp[ii].className && (tmpp[ii].className.substring(tmpp[ii].className.length-6,tmpp[ii].className.length) == "_title"||tmpp[ii].className.indexOf("_title ")!=-1)){
						tmpp[ii].onclick=function(){
							var k=0;
							if(this==LD.mun_clickcur)k=1;
							if(k==0&&this.parentNode)
							if(this.parentNode==LD.mun_clickcur)k=1;
							if(k==0&&this.parentNode.parentNode)
							if(this.parentNode.parentNode==LD.mun_clickcur)k=1;
							LD.cflag=1;
							if(k==0)return;
							var c=0;
							if(LD.mun_clickcur)
							{
								if(LD.mun_clickcur.info.style.display!="none")
								{
									//setTimeout("(function(){LD.cflag=0;LD.mun_clickcur.info.style.display=\"none\";})()",1100);//delay to fixed ie viewed bug
									LD.cflag=0;LD.mun_clickcur.info.style.display="none";
									c=1;
								}
							}
							if(c==0)
							if(LD.mnu_curele)
							{
								if(LD.mnu_curele.info.style.display!="none")
								{
									LD.cflag=0;
									LD.mnu_curele.info.style.display="none";
								}
								else
								LD.cflag=1;
							}
						};break;
					}}
				tmpp = item.getElementsByTagName("ul")
				for(var ii=0; ii<tmpp.length; ii++)
				{
					if(tmpp[ii].innerHTML != "")
					{
						tmpp[ii].onclick=function(){LD.cflag=0;};break;
					}
				}
				item.type = "option";
				item.onclick = this.showInfo;
			}
		}
		else
		{
			item.onmouseover = this.showInfo;
			item.onmouseout = this.hiddenInfo;
		}
	}
}

LayerMenu.prototype.showInfo = function(e){
	LD.mnu_curele = this;
	if(this.info)
	{
		if(this.atitle)this.atitle.parentNode.className = "hover";
		var m=-1,mm=-1,t=this;
		for(var n=0;n<10;n++)
		{
			if(this.childNodes[n].nodeName!="#text")
			{
				m=n;
				break;
			}
		}
		if(m>=0)
		for(var n=0;n<10;n++)
		{
			if(isUndef(this.childNodes[m].childNodes[n]))break;
			if(this.childNodes[m].childNodes[n].nodeName!="#text")
			{
				mm=n;
				break;
			}
		}
		if(mm>=0)
		{
			t=this.childNodes[m].childNodes[mm];
		}
		else
		{
			if(m>=0)t=this.childNodes[m];
		}
		//fixed ie 6.0 5.5 max-height start
		if(Browser.ie&&!Browser.ie8&&!Browser.ie7&&!Browser.ie9)
		if(this.info.className=="layer_optioninfo")
		{
			if(LD.cflag==1||(this.info.className!="layer_optioninfo"&&this.info.className!="layer_iteminfo"))this.info.style.display = 'block';
			if(this.info.offsetHeight>400)this.info.style.height="400px";
		}
		//fixed ie 6.0 5.5 max-height end

		if(this.type != "prompt")
		LD.setpos(t,this.info);
		else
		LD.setpos(t,this.info,15,this.type,e);
	}
	//if(this.className.indexOf('active')==-1)this.className += 'active'; 
	if(this.onclick!=null && this.type != "prompt" && this.type != "option")
	{
		this.info.onmouseover = function(){if(LD.mnu_curele.info&&LD.cflag==1)LD.mnu_curele.info.style.display='block';};
		this.onmouseout = function(){if(LD.mnu_curele.info)LD.mnu_curele.info.style.display='none';};
		//
	}
	if(this.type=="option")
	{
		LD.mun_clickcur = this;
		try {
			if(edt_doc)
			addListener(edt_doc,"click",function(evt)
			{
				if(LD.mun_clickcur.info)LD.mun_clickcur.info.style.display='none';
			})
		}catch(e){}
		addListener(document,"click",function(evt)
		{
			var evt = window.event?window.event:evt,target=evt.srcElement||evt.target;
			//var classtr = "|layer_option|";
			var classtr = "|layer_option|ico|layerico|layer_item|";
			var tar1 = target,tar2 = target.parentNode,flag=1;
			if(tar1.nodeName.toLowerCase() == "div")
			if(tar1==LD.mun_clickcur.info||tar1.id=="editor_selcolor")flag=0;
			if(classtr.indexOf("|" + tar1.className + "|")!=-1)flag=0;
			while(tar2 && flag==1 && tar1 != tar2 && tar2.nodeName.toLowerCase() != "html")
			{
				if(classtr.indexOf("|" + tar2.className + "|")!=-1)
				{
					flag=0;
					if(LD.mun_oldclickcur!=""&&LD.mun_clickcur.info)
					if(LD.mun_oldclickcur.id!=LD.mun_clickcur.info.id)
					{LD.mun_oldclickcur.style.display="none";}
					LD.mun_oldclickcur = LD.mun_clickcur.info;
					if(LD.mun_clickcur.info&&Browser.ie&&!Browser.ie8&&!Browser.ie9)
					{
						//delay to fixed ie 6&7 viewed bug
						if(LD.mun_clickcur.info.offsetHeight>400)setTimeout("(function(){LD.mun_clickcur.info.style.overflowY = 'scroll';})()",300);
					}
					
				}
				tar1 = tar2;
				tar2 = tar2.parentNode;
			}
			if(LD.mun_clickcur.info)
			if(LD.mun_clickcur.info.style.display=="none")flag=0;
			if(flag == 1)
			{
				if(LD.mun_clickcur.info)LD.mun_clickcur.info.style.display='none';
				LD.mun_oldclickcur="";
			}
		});
	}
}


LayerMenu.prototype.hiddenInfo = function(e){
	var evt = window.event || e;
	var elem = evt.target || evt.srcElement;
	if(this.info)
	{
		//if(elem.className != this && elem != this.info){
		//hidden later
		if(this.info.parentNode==this)
		{
			LD.mnu_curele = "";
			setTimeout("(function(){LD.mnu_hide('"+this.id+"')})()",100);
		}
		//hidden at once
		else
		{
			this.info.style.display = 'none';
		}
		//}
		//this.info.style.display = 'none';
		if(this.atitle)this.atitle.parentNode.className = "mouseout";
	}
	//this.className = this.className.replace(/active/g,'');
}

//cookie
LD.Cookie={
	StrToDate:function(s) {
	    var d = new Date(); 
	    var g = s.length;
	    if(g>=4)d.setYear(parseInt(s.substring(0,4),10));
	    if(g>=6)d.setMonth(parseInt(s.substring(4,6)-1,10)); 
	    if(g>=8)d.setDate(parseInt(s.substring(6,8),10)); 
	    if(g>=10)d.setHours(parseInt(s.substring(8,10),10));
	    if(g>=12)d.setMinutes(parseInt(s.substring(10,12),10));
	    if(g>=14)d.setSeconds(parseInt(s.substring(12,14),10));
	    return(d);
	},

	Add:function(name,value,esc)
	{
		var expiredate = LD.Cookie.Get(DEF_MasterCookies,"expires");
		if(expiredate.length<8)expiredate=30;
		var cookieString=name+"="+(isUndef(esc)?escape(value):value);
		if(expiredate.length>=8)
		{
			cookieString=cookieString+"; expires="+(this.StrToDate(expiredate).toUTCString());
		}
		else
		{
			var date=new Date();
			date.setTime(date.getTime()+expiredate*3600000*24);
			cookieString=cookieString+"; expires="+date.toUTCString();
		
		}
		cookieString = cookieString+";path=/; domain="+this.Domain();
		document.cookie=cookieString;
	},
	Domain: function(flag)
	{
		var tmp = window.location.host;
		tmp=tmp.indexOf(":")>0?tmp.substring(0,tmp.indexOf(":")):tmp;
		if(flag!=1)
		{
			if(isNaN(tmp.replace(/\./g,'')) == true)
			if(tmp.split(".").length>2)tmp=tmp.substring(tmp.indexOf(".")+1,tmp.length);
		}
		return(tmp);
	},

	Get: function(name,name2)
	{
		var str=document.cookie;
		var arrCookie=str.split("; ");
		str = "";
		for(var i=0;i<arrCookie.length;i++)
		{
			if(arrCookie[i].indexOf(name+"=")==0)
			{
				str = arrCookie[i].substring(name.length+1,arrCookie[i].length);
			}
		}
		if(!name2)return(unescape(str));
		if(str=="")return("");
		arrCookie=str.split("&");
		for(var i=0;i<arrCookie.length;i++)
		{
			if(arrCookie[i].indexOf(name2+"=")==0)
			{
				str = arrCookie[i].substring(name2.length+1,arrCookie[i].length);
			}
		}
		return(unescape(str));
	},

	Del:function(name,flag){
	        var date=new Date();
	        date.setTime(date.getTime()-10000);	        
	        document.cookie=name+"=; expires="+date.toUTCString()+";domain="+this.Domain();
	},
	Clear:function()
	{
		var temp=document.cookie.split(";");
		var n,s;
		for (n=0;n<temp.length;n++)
		{
			if(temp[n].indexOf("="))
			{
				s=temp[n].split("=")[0];
				this.Del(s);
				this.Del(s,1);
			}
		}
		alert("Successfully cleared.");
	}
};

LD.move = {
	posX:0,
	posY:0,
	fdiv:null,
	mousedown:function(div,e)
	{
		if(!e) e = window.event;
		var tar=e.srcElement||e.target;
		if(tar)
		if(tar.tagName)
		if("|input|a|img|".indexOf("|" + tar.tagName.toLowerCase() + "|")!=-1)return;
		LD.move.fdiv = div;
		LD.move.id=div.id;
		LD.move.posX = e.clientX - LD.getX(div,"abs");
		LD.move.posY = e.clientY - LD.getY(div,"abs");
		document.onmousemove = LD.move.mousemove;
		document.onmouseup = function()
		{
			document.onmousemove = null;
		}
	},

	mousemove:function(ev)
	{
		if(ev==null) ev = window.event;
		LD.move.fdiv.style.left = (ev.clientX - LD.move.posX) + "px";
		LD.move.fdiv.style.top = (ev.clientY - LD.move.posY) + "px";
	}
};
//import file

function $import_loadexecute(s,execut)
{
	if(typeof(execut)=="function")
	{				
		s.doneState = { loaded: true, complete: true};
		if(s.onreadystatechange !== undefined)
		{
			s.onreadystatechange = function()
			{
				if(s.doneState[s.readyState] )
				{
					execut();
				}
			};
		}
		else
		{
			s.onload = function()
			{
			    	execut();
			};
		}
	}
}
function $import(path,type,title,execut,must)
{
	var s,i;
	if(type=="js")
	{
		var ss=document.getElementsByTagName("script");
		if($id("tempjs"+must))return;
		if(!must)
		for(i=0;i<ss.length;i++)
		{
			if(ss[i].src && ss[i].src.indexOf(path)!=-1)return;
		}
		s=document.createElement("script");
		s.type="text/javascript";
		s.src=path;
		if(must)s.id="tempjs"+must;
		$import_loadexecute(s,execut);
	}
	else if(type=="css")
	{
		s=document.createElement("link");
		s.rel="stylesheet";
		s.type="text/css";
		s.href=path;
		s.title=title;
		s.disabled=false;
	}
	else return;
	var head=document.getElementsByTagName("head")[0];
	head.appendChild(s);
}

function setStyle2(file,title)
{
	var i=0, Links;
	$import(file,"css",title);
	Links = document.getElementsByTagName("link")[i]
	while(Links)
	{
		if(Links.title==title)
		{document.getElementsByTagName("link")[0].disabled = true;
		document.getElementsByTagName("link")[0].disabled = false;
		}
		i++;
		Links = document.getElementsByTagName("link")[i]
	}
}

function setStyle(file,title)
{
	var i, links,eflag=false;
	var Links=document.getElementsByTagName("link")[0];
	while(Links)
	{
		//if(Links.title==title)return;
		Links.parentNode.removeChild(Links);
		Links=document.getElementsByTagName("link")[0]
	}
	$import(file,"css",title);
	if(document.getElementsByTagName("link")[0])
	{
		document.getElementsByTagName("link")[0].disabled = true;
		document.getElementsByTagName("link")[0].disabled = false;
	}
}

//ajax

function getHttp()
{
	var oT = false;
	try {
		oT = new ActiveXObject('Msxml2.XMLHTTP');
		} catch(e) {
	try {
		oT = new ActiveXObject('Microsoft.XMLHTTP');
		} catch(e) {
			oT = new XMLHttpRequest();
		}
	}
	return oT;
}


function getAJAX(url,str,lb,ty,execut)
{
	delete HR;
	var HR = getHttp();
	function processAJAX(lb)
	{
		if (HR.readyState == 4)
		{
			if (HR.status == 200)
			{
				var tmp = HR.responseText;
				if(ty==1)
					eval(lb);
				else
					$id(lb).innerHTML = tmp;
				if(!isUndef(execut))eval(execut);
			}
			else
			{
				var tmp = "<p>Page error: " + HR.statusText +"<\/p>";
				$id('errorstr').innerHTML=HR.responseText;
				if(ty==1)
					eval(lb);
				else
					$id(lb).innerHTML=tmp;
			}
			layer_initselect();
			delete HR;
			HR=null;
			delete HR;
			if(Browser.ie)CollectGarbage;
		}
	}
	HR.onreadystatechange = function() {processAJAX(lb);};
	if(str!="")
		HR.open("POST", url, true);
	else
		HR.open("GET", url, true);
	HR.setRequestHeader("Content-Type","application/x-www-form-urlencoded;charset=gb2312");
	HR.send(str);
}

//string

function $replace(str,str1,str2)
{
	var re = new RegExp(str1,"g");
	return(str.replace(re,str2));
}

function submit_disable(theform,tp)
{
	if (document.all||document.getElementById)
	{
		for (var i=0;i<theform.length;i++)
		{
			var tempobj=theform.elements[i];
			if(tempobj.type && (tempobj.type.toLowerCase()=="submit"||tempobj.type.toLowerCase()=="reset"))
			{
				if(tp==1)
				tempobj.disabled=false;
				else
				tempobj.disabled=true;
			}
		}
	}
}

if(!Browser.ie){ //firefox innerText define
    HTMLElement.prototype.__defineGetter__(    "innerText", 
        function(){ 
            return this.textContent.replace(/(^\s*)|(\s*$)/g, "");
        } 
    ); 
    HTMLElement.prototype.__defineSetter__(    "innerText", 
        function(sText){ 
            this.textContent=sText; 
        } 
    ); 
}

function relocationxy(ob,menuob,edt,center,timeout,posx,event)
{	
	document.body.style.cursor = "wait";
	var obj=$id(ob);
	var menuobj = $id(menuob);

	var menuobj_get,tmp;
	tmp = menuobj.childNodes;
	menuobj_get=menuobj;
	if(tmp.length==2)
	{
		if(tmp[0].nodeName!="#text")
		if(tmp[0].className.substring(0,13)=="ajaxitemtitle")
		menuobj_get=tmp[1];
	}
	if(menuobj_get.innerHTML=="loading..."||menuobj_get.innerHTML=="")menuobj_get.innerHTML=layer_loadstr;
	if(menuobj_get.innerHTML.replace(/\'/g, "").replace(/\"/g, "").toLowerCase()!=layer_loadstr.replace(/\'/g, "").replace(/\"/g, "").toLowerCase())
	{
		if(center>=1)
		//setCenterDiv(menuobj);
		{}
		else
		LD.setpos(obj,menuobj,posx,"undefined",event);
		document.body.style.cursor = "";

		if(center<1)
		{
			var classtr;
			if(menuobj.className=="layer_alertmsg")
			classtr = "|layer_alertclick|";
			else
			classtr = "|layer_alertmsg|ico|layerico|layer_ajaxitem|layer_ajaxmsg|";
			addListener(document,"click",function(evt)
			{
				var evt = window.event?window.event:evt,target=evt.srcElement||evt.target;
					//while(target.nodeName.toLowerCase() != "div" && target.nodeName.toLowerCase() != "html")
					//{
					//	target = target.parentNode;
					//}
					var tar1 = target,tar2 = target.parentNode,flag=1;
					if(tar1.nodeName.toLowerCase() == "div")
					if(tar1.id==menuob||tar1.id=="editor_selcolor")flag=0;
					if(classtr.indexOf("|" + tar1.className + "|")!=-1)flag=0;
					while(tar2 && flag==1 && tar1 != tar2 && tar2.nodeName.toLowerCase() != "html")
					{
						if(classtr.indexOf("|" + tar2.className + "|")!=-1)flag=0;
						if(tar2.nodeName.toLowerCase() == "div")
						if(tar2.id==menuob||tar2.id=="editor_selcolor")flag=0;
						tar1 = tar2;
						tar2 = tar2.parentNode;
					}
					if($id(menuob).style.display=="none")flag=0;
					if(flag == 1)
					{
						$id(menuob).style.display = "none";
					}
			})
		}
		
		if(edt==1)
		addListener(edt_doc,"click",function(evt)
		{
			$id(menuob).style.display = "none";
		})
	}
	else
	{
		if(timeout!=1)
		if(center>=1)
		{}
		//setCenterDiv(menuobj);
		else
		LD.setpos(obj,menuobj,posx,"undefined",event);
		setTimeout("relocationxy('" + ob + "','" + menuob+ "','" + edt + "','" + center +"',1,'" + posx +"');",100);
	}
}

function addListener(element,e,fn,del)
{
	if(element.addEventListener)
	{
		if(del==1)
		element.removeEventListener(e,fn,false);
		else
		element.addEventListener(e,fn,false);
	}
	else
	{
		if(del==1)
		element.detachEvent("on" + e,fn);
		else
		element.attachEvent("on" + e,fn);
	}
}

var pub_clipdata = "";
function set_clipboarddata() {
	window.document.obj_clipboard.SetVariable('str', pub_clipdata);
}

function copyClipboard(data,value,alertmsg,dir,obj,event)
{
	if (window.clipboardData)
	{
		window.clipboardData.setData(data,value);
		alert(alertmsg);
	}else
	{
		pub_clipdata = value;
		layer_create("win_clipboard");
		$id('win_clipboard').innerHTML='<div style=margin-top:10px;text-align:center;font-weight:bold;><a href=#copy>再次点击复制到剪帖板。</a></div><embed name="obj_clipboard" devicefont="false" src="/images/pub/_clipboard.swf" menu="false" allowscriptaccess="sameDomain" swliveconnect="true" wmode="transparent" style="margin-top:-20px;z-index:11;" type="application/x-shockwave-flash" width="150" height="30">';
		layer_view("",obj,'','','win_clipboard','','',0,'',0,0);

	}   
}

function getPageScroll(){

	var yScl;

	if (self.pageYOffset) {
		yScl = self.pageYOffset;
	} else if (document.documentElement && document.documentElement.scrollTop){
		yScl = document.documentElement.scrollTop;
	} else if (document.body) {
		yScl = document.body.scrollTop;
	}

	arrayPageScroll = new Array('',yScl) 
	return arrayPageScroll;
}

function getPageSize(){
	
	var xScl, yScl;
	
	if (window.innerHeight && window.scrollMaxY) {	
		xScl = document.body.scrollWidth;
		yScl = window.innerHeight + window.scrollMaxY;
	} else if (document.body.scrollHeight > document.body.offsetHeight){
		xScl = document.body.scrollWidth;
		yScl = document.body.scrollHeight;
	} else {
		xScl = document.body.offsetWidth;
		yScl = document.body.offsetHeight;
	}
	
	var wW, wH;
	if (self.innerHeight) {
		wW = self.innerWidth;
		wH = self.innerHeight;
	} else if (document.documentElement && document.documentElement.clientHeight) {
		wW = document.documentElement.clientWidth;
		wH = document.documentElement.clientHeight;
	} else if (document.body) {
		wW = document.body.clientWidth;
		wH = document.body.clientHeight;
	}	

	if(yScl < wH){
		pageHeight = wH;
	} else { 
		pageHeight = yScl;
	}

	if(xScl < wW){	
		pageWidth = wW;
	} else {
		pageWidth = xScl;
	}


	arrayPageSize = new Array(pageWidth,pageHeight,wW,wH) 
	return arrayPageSize;
}

function setCenterDiv(obj)
{
	var arrayPageSize = getPageSize();
	var arrayPageScroll = getPageScroll();

	if (obj) {
		h=obj.offsetHeight;
		w=obj.offsetWidth;
		obj.style.top = (arrayPageScroll[1] + ((arrayPageSize[3] - 35 - h) / 2) + 'px');
		obj.style.left = (((arrayPageSize[0] - 20 - w) / 2) + 'px');
	}
}

function layer_outmsg(obj,str,url,evl)
{
	var menuobj_get,tmp;
	menuobj_get=$id(obj);
	if(menuobj_get.style.display=="none")menuobj_get=$id("anc_msgbody");
	tmp = menuobj_get.childNodes;
	if(tmp.length==2)
	if(tmp[0].nodeName!="#text")
	if(tmp[0].className.substring(0,13)=="ajaxitemtitle")
	menuobj_get=tmp[1];
	
	menuobj_get.innerHTML=str;
	if(url!=""&&!isUndef(url))
	this.location=url;
	else
	{
		if(evl!=""&&!isUndef(evl))
		eval(evl)
		else
		setTimeout("layer_hidelayer($id('" + obj + "'));",3000);
	}
}

function layer_hidelayer(obj)
{
	if(obj.style.display=="none")return;
	var overlay = $id('layer_overlay')
	/*
	if(fadeflag)
	{
		$(id).fadeOut(300);
		setOpacity(overlay,0);
	}
	else
	{*/
		obj.style.display="none";
		setOpacity(overlay,0);
		if(overlay)overlay.style.display="none";
	/*}*/
	selects = document.getElementsByTagName("select");
	for (i = 0; i != selects.length; i++) {
		selects[i].style.visibility = "visible";
	}
	addListener(Browser.ie?document.body:window,"keydown",getKeys,1);
}

function layer_createoverlay(obj,center)
{
	if($id('layer_overlay'))
	{
		if(obj&&center!=1000)overlay.onclick = function () {layer_hidelayer(obj); return false;}
		return($id('layer_overlay'));
	}
	
	var overlay = document.createElement("div");
	if(obj&&center!=1000)overlay.onclick = function () {layer_hidelayer(obj); return false;}
	overlay.setAttribute('id','layer_overlay');
	overlay.className = "overlay";
	document.getElementsByTagName("body").item(0).insertBefore(overlay, document.getElementsByTagName("body").item(0).firstChild);
	return(overlay);
}

function setOpacity(obj,n)
{
	if(fadeflag==0)return;
	//if(n=="end"){if(obj)Browser.ie?obj.style.filter="":obj.style.opacity = "";return;} //clear filter bug
	if(obj)Browser.ie?obj.style.filter="Alpha(Opacity=" + n + ")":obj.style.opacity = n/100;
}

function getOpacity(obj)
{
	var t=(Browser.ie?obj.style.filter:obj.style.opacity);
	if(t==""||isUndef(t))return(100);
	return(Browser.ie?t.substring(14,t.length-1):t*100);
}

function layer_viewlayer(obj,center,show)
{
	var msgflag=(obj.className=="layer_ajaxmsg")?1:0;
	if(msgflag==0)setOpacity(obj,0);
	if(show!="none")obj.style.display="block";
	//clearTimeout(viewTimer);
	if(fadeflag&&msgflag==0&&show!="none")
	$(obj).fadeIn(300);
	else
	if(msgflag==0)setOpacity(obj,100);

	if(center<1)return;
	var overlay = $id('layer_overlay');
	if(!$id('layer_overlay'))
	{
		overlay=layer_createoverlay(obj,center);
	}
	selects = document.getElementsByTagName("select");
	for (i = 0; i != selects.length; i++) {
		selects[i].style.visibility = "hidden";
	}
	
	var arrayPageSize = getPageSize();
	overlay.style.height = (arrayPageSize[1] + 'px');
	
	setOpacity(overlay,0);
	overlay.style.display="block";
	setOpacity(overlay,40);
	if(center!=1000)addListener(Browser.ie?document.body:window,"keydown",getKeys);
}

function getKeys(e){
	if (Browser.ie) {
		keycode = event.keyCode;
	} else {
		keycode=e.keyCode;
	}
	if(keycode == 27){ layer_hidelayer($id('anc_delbody')); }
}

var layer_num=0,layer_olditem;
/*
obj hited object
w view layer width
h layer height
menu string of objectid
file loading fileurl
js loading js file url
fresh forced to fresh
filepara ajax post data
center layer center: >=1.center <1.object position  1000.cancel onclick event
posx layer horizontal shifting
*/
function layer_create(menu,center,title,file)
{
	if(isUndef($id(menu))||menu=='anc_delbody')
	{		
		var tj = document.createElement("div");
		tj.setAttribute('id',menu);
		var ttl=(!isUndef(title)&&title!="")?title:"";
		var fl=(!isUndef(file)&&file!="")?file:"";
		
		if(menu=="anc_delbody"||ttl!="")
		{
			if(fl==""&&menu!="anc_delbody")
			{
				tj.innerHTML="<div><span>" + ttl + "</span></div>";
				tj.className = "layer_alertmsg";
			}
			else
			{tj.innerHTML="<div unselectable=on class=\"ajaxitemtitle fire unsel\" style=\"cursor: move;\" onmousedown=\"LD.move.mousedown(this.parentNode,event);\"><span id=\"" + menu + "title\" unselectable=on class=\"title unsel\"></span><div class=\"layer_close\"><a href=\"javascript:;\" onclick=\"if(1000==" + center + "){if(confirm('\u63D0\u9192\uFF1A\u6B64\u64CD\u4F5C\u5C06\u5173\u95ED\u5F53\u524D\u7A97\u53E3\u3002')){layer_hidelayer($id('" + menu + "'));return false;}}else{layer_hidelayer($id('" + menu + "'));return false;}\" class=\"unsel\" hidefocus=\"true\" title=\"close\" /></a></div></div><div class=\"ajaxitembody\">loading...</div>";
			tj.className = "layer_ajaxitem";
			}
		}
		else
		{
			tj.className = "layer_ajaxmsg";
		}
		document.getElementsByTagName("body").item(0).insertBefore(tj, document.getElementsByTagName("body").item(0).firstChild);
		$(tj).hide();
	}
	else
	{
		if($id(menu).className=="layer_alertmsg")$id(menu).innerHTML="<div><span>" + title + "</span></div>";
	}
}

function layer_view(title,obj,w,h,menu,file,js,fresh,filepara,center,posx,evl,event)
{
	var objexist = (obj&&obj!="")?true:false;
	layer_create(menu,center,title,file);
	if(menu=="anc_msgbody")
	{layer_create("anc_delbody",center);layer_hidelayer($id("anc_delbody"));}
	else
	{layer_create("anc_msgbody",center);$id("anc_msgbody").style.display="none";}

	if(title!=""&&$id(menu+"title"))$id(menu+"title").innerHTML=title;
	var menuobj=$id(menu);
	
	if(w!="")
	menuobj.style.width = w+"px";
	else
		menuobj.style.width = "";
	if(h!="")menuobj.style.height = h+"px";
	
	var menuobj_get,tmp,menu_get;
	tmp = menuobj.childNodes;
	
	menuobj_get=menuobj;
	menu_get = menu;
	if(tmp.length==2)
	{
		if(tmp[0].nodeName!="#text")
		if(tmp[0].className.substring(0,13)=="ajaxitemtitle")
		{
			menuobj_get=tmp[1];
			menu_get=menuobj_get.id;
			if(menuobj_get.id=="")
			{
				menuobj_get.id="ajaxitem_body"+layer_num;
				menu_get = "ajaxitem_body"+layer_num;
				layer_num++;
			}
		}
	}

	if(objexist)
	if(obj.id==""){obj.id="ajaxitem_view"+layer_num;layer_num++;}

	if(menuobj_get.innerHTML=="loading..."||menuobj_get.innerHTML=="")menuobj_get.innerHTML=layer_loadstr;
	if(fresh==1 || (file && menuobj_get.innerHTML.replace(/\'/g, "").replace(/\"/g, "").toLowerCase()==layer_loadstr.replace(/\'/g, "").replace(/\"/g, "").toLowerCase()))
	{
		if(fresh==1)menuobj_get.innerHTML=layer_loadstr;
		var evlajax = "";
		if(fadeflag==1)
		{
			evlajax="$('#" + menu + "').fadeIn(300);"
		}
		else{
			evlajax="$('" + menu + "').style.display='block';"
		}
		if(center==1)evlajax+="setCenterDiv($id('" + menu + "'));";
		if(!isUndef(filepara))
		{
			if(isUndef(evl))
			{
				evl="$id('" + menu_get + "').innerHTML = tmp;";
				evl+=evlajax;
				getAJAX(file,filepara,evl,1);
			}
			else
			{
				evl+="$id('" + menu_get + "').innerHTML = tmp;";
				evl+=evlajax;
				getAJAX(file,filepara,evl,1);
			}
		}
		else
		{
			if(isUndef(evl))
			{
			getAJAX(file,"",menu_get,evl);
			}
			else
			{
				evl+="$id('" + menu_get + "').innerHTML = tmp;";
				evl+=evlajax;
				getAJAX(file,"",evl,1);
			}
		}
		if(js)editor_inputjs(menu_get,js);
	}
	if(fresh==1 || file=="" && js)
	{$import(js,"js");}
	
	if(fresh!=1)
	{
		if(menuobj.style.display != 'none')
		{
			//menuobj.style.display = 'none';
			if(objexist)
			if(layer_olditem==obj.id)return;
			else
			if(layer_olditem=="")return;
		}
	}
	if(center==1)
	{
		//setCenterDiv(menuobj);
	}
	else
	{
		var x,y,distHeight;
		if(objexist)
		{
			x = LD.getX(obj);
			y = LD.getY(obj)-1;
			distHeight = obj.clientHeight;
			if(!distHeight)distHeight+=obj.offsetHeight;
			y += distHeight;
			if(!isUndef(posx)&&posx!="undefined"){x += posx;}
		}
		else
		{
			x = LD.getMX(event);
			y = LD.getMY(event)-1;
			if(!isUndef(posx)&&posx!="undefined")x += posx;
		}
		menuobj.style.top = y + 'px';
		menuobj.style.left = x + 'px';
	}
	if(center!=1)
	{menuobj.style.display = 'block';
	layer_viewlayer(menuobj,center,"none");
	}
	else
		layer_viewlayer(menuobj,center,"none");
	if(objexist)
	{
		relocationxy(obj.id,menuobj.id,0,center,0,posx,event);
		layer_olditem = obj.id;
	}
	else
	{
		relocationxy("",menuobj.id,0,center,0,posx,event);
		layer_olditem = "";
	}

}

function pub_command(str,obj,div,para,evl)
{
	layer_view(str,obj,'','',div,obj.href,'',1,'AjaxFlag=1' + para,1,0,evl);return(false);
}

function pub_msg(obj,div,para,evl)
{
	layer_view('',obj,'','',div,obj.href,'',1,'AjaxFlag=1' + para,0,0,evl);return false;
}

function insertAfter(newElement,targetElement)
{
	var parent = targetElement.parentNode;    
	if(parent.lastChild == targetElement)
	{
		parent.appendChild(newElement);    
	}
	else
	{
		parent.insertBefore(newElement,targetElement.nextSibling);    
	}
}

function layer_initselect()
{
	var selects = document.getElementsByTagName("select"),l,i,j,layerstr,tmp,c,tmp2;
	for (i = 0; i != selects.length; i++) {
	if(selects[i].style.display != "none")
	{
		if(selects[i].id==""){selects[i].id="selectmenu" + LD.mnu_n;LD.mnu_n++;}
		
		layerstr = "<div class=\"layer_option\"><div><div class=\"unsel layer_option_title\" unselectable=\"on\"><div id=\"layer_" + selects[i].id + "\" class=\"unsel\" unselectable=\"on\">" + selects[i].options[selects[i].selectedIndex].innerHTML + "</div></div></div></div>";
		tmp = document.createElement("div");
		tmp.style.cssText = "vertical-align:middle;display: inline-block;*display: inline;zoom:1;";
		tmp.innerHTML = layerstr;
		insertAfter(tmp,selects[i]);
		layerstr = "";

		layerstr += "<ul class=\"menu_list\">";
		
		l = selects[i].options.length;
		tmp2 = (selects[i].onchange!=null)?"$id('" + selects[i].id + "').onchange();":"";
		for (j=0; j < l; j++)
		{
			layerstr += "<li";
			c = selects[i].options[j].style.cssText;
			if(c!="")layerstr += " style=\"" + c + "\"";
			layerstr += "><a href=\"javascript:;\" onclick=\"$id('" + selects[i].id + "').options[" + j + "].selected=true;$id('layer_" + selects[i].id + "').innerHTML=this.innerHTML;$id('layerinfo_" + selects[i].id + "').style.display='none';"
			layerstr += tmp2 + "return false;\">" + selects[i].options[j].innerHTML + "</a></li>";
		}
		layerstr += "</ul>";
		tmp = document.createElement("div");
		tmp.className = "layer_optioninfo";
		tmp.id = "layerinfo_" + selects[i].id;
		tmp.onclick = function(){this.style.display="none";};
		tmp.innerHTML = layerstr;
		insertAfter(tmp,selects[i]);
		selects[i].style.display = "none";
		if(selects[i].style.width!="" && selects[i].style.width!="auto")
		{
		$id("layer_" + selects[i].id).style.width = (parseInt($replace(selects[i].style.width,"px",""))-30) + "px";
		$id("layer_" + selects[i].id).style.overflow = "hidden";
		}
	}
	}
	new LayerMenu('layer_option','layer_optioninfo','option');
}

function verify_load(r,u)
{
	var verifyload = ($id('verify_click').style.display=='none');
	if(verifyload&&r!=1)return;
	$id('verify_click').style.display='none';
	var ur=isUndef(u)?HU:u
	$id('verifycode').src=ur+'User/number.asp?r=1&' + Math.random();
	$id('verifycode').className="verifycode";
}

function sendprivatemsg(obj,u)
{
	layer_view('发送消息',obj,'625','','anc_delbody',obj.href,'',1,'AjaxFlag=1&dir='+u,1000,0,'');
	return(false);
}


function SelectItemByValue(objSelect, objItemText)
{
	for (var i = 0; i < objSelect.options.length; i++)
	{
		if (objSelect.options[i].value+"" == objItemText)
		{
			objSelect.options[i].selected = true;
			break;
		}
	}
}

var maxjsreload = 6,jsreloadcount=0;
// rerun script command and js file
function js_Reload(obj)
{
	var child=obj.getElementsByTagName("SCRIPT");
	var importFlag = 0;
	for(var i=0;i<child.length;i++)
	{
		if(child[i].src)
		{
			$import(child[i].src+" ","js","ld_reloadjs"+i,function(){
				jsreloadcount++;
				if(jsreloadcount<maxjsreload)js_Reload(obj);
			},i);
			importFlag = 1;
		}
		else
		{
			if(child[i].innerHTML.replace(/(\n|\r|\ |\t|\f)/g, "")=="")
			{
			}
			else
			eval(child[i].innerHTML);
		}
	}
	if(importFlag == 0)
	{
		jsreloadcount++;
		if(jsreloadcount<maxjsreload)js_Reload(obj);
	}
}

$(document).ready(function() {
       // focus on the first text input field in the first field on the page
       if(typeof edt_win!="undefined")edt_win.focus();
	else
        {
        	var s=$("input[type='text']:first", document.forms[0]);
        	if($(s)&&($(s).attr("class"))&&($(s).attr("class").indexOf("notfocus")==-1))
        	$(s).focus();
        }
});