//File:cnsidepl.html
//Writen by hoja
//via:the opener document.
//Tmv Tgasa Tnew replace the images of the list menu title(only for the access version).
//elmABlock:how many songs in one page.
//cookie_path:save the list cookie used of cookie_path.
//cookie_domain:save the list cookie used of cookie_domain.
//you can also change the day to change the cookies expdate

	var via = opener;
	var write_via = "opener";
	var iLoc= self.location.href;

	function playSel(){via.wmpStop();via.startExobud();}
	function refreshPl(){ self.location=iLoc;}
	function chkSel(){via.chkAllSel();refreshPl();}
	function chkDesel(){via.chkAllDesel();refreshPl();}

	function dspList(n){
		var elmABlock= 9;
		var totElm = via.intMmCnt;
		var totBlock= Math.floor((via.intMmCnt -1) / elmABlock)+1;
		var cblock;
		if(n==null){cblock=1;}
		else{cblock=n;}
		var seed;
		var limit;
		if(cblock < totBlock){seed= elmABlock * (cblock-1); limit =  cblock*elmABlock -1}
		else{seed=elmABlock * (cblock-1); limit= totElm-1;}

	if(via.intMmCnt >0 ){
		var list_num=0;
		mmList.innerHTML='<p>';
		pageList.innerHTML='<br>页 ';
		for (var i=seed; i <= limit; i++)
		{
			var Tmv='<img src=img/mv.gif alt=ＭＶ width=14 height=14 border=0 align=absbottom>';
			var TGasa='<img src=img/gasa3.gif alt=有歌词 width=14 height=14 border=0 align=absbottom>';
			var Tnew='<img src=img/up.gif alt=新歌速递 width=12 height=9 border=0 align=absbottom>';

			var TitleInfo=via.objMmInfo[i].mmTit.replace(Tmv," ").replace(TGasa," ").replace(Tnew," ");

			list_num = i + 1;
			elm = list_num + '.';
			if(via.objMmInfo[i].selMm=="t"){elm=elm+'<input type=checkbox class=fmchkbox style="cursor:hand;" onclick='+ write_via + '.chkItemSel('+ i +'); checked>' ;}
			else{elm =elm+ '<input type=checkbox class=fmchkbox style="cursor:hand;" onclick='+ write_via + '.chkItemSel('+ i +');>' ;}
			elm = elm + '<a href=javascript:' + write_via + '.selPlPlay(' + i + ');'
			if(strlength(via.objMmInfo[i].mmTit)>35)
			{
				elm = elm + ' onclick=\"this.blur();\" title=\"' + TitleInfo + '\">' + LeftTrue(via.objMmInfo[i].mmTit,32) + '...</a><br>';
			}
			else
			{
				elm = elm + ' onclick=\"this.blur();\" title=\"' + TitleInfo + '\">' + via.objMmInfo[i].mmTit + '</a><br>';
			}
			mmList.innerHTML=mmList.innerHTML+elm;
		}
		var pmin=cblock-3;
		var pmax=cblock+3;
		var ppre=cblock-1;
		var pnext=cblock+1;
		if(pmin<=3){pmin=1;pmax=7;}
		if(pmax>totBlock){pmax=totBlock;}
		if(ppre<=1){ppre=1;}
		if(pnext>totBlock){pnext=totBlock;}

		for(var j=pmin; j<=pmax; j++){
			page='<a href=javascript:dspList('+j+') title=第'+j+'页>['+j+']</a> ';
			pageList.innerHTML=pageList.innerHTML+page;
		}
		
		pageInfo.innerHTML=''
		+'<br>'
		+'<a href=javascript:dspList(1) title=首页><font face=webdings title=首页>9</font></a><a href=javascript:dspList('+ppre+') title=上一页><font face=webdings title=上页>7</font></a>&nbsp;第<font color=blue class=bluefont>'+cblock+ '</font>页&nbsp;共'+ totBlock+'页&nbsp;共'+totElm+'首&nbsp;<a href=javascript:dspList('+pnext+') title=下一页><font face=webdings title=下页>8</font></a><a href=javascript:dspList('+totBlock+') title=尾页><font face=webdings title=尾页>:</font></a>';
		//list_top.innerHTML = '<img src=img/scope_on.gif border=0>';
	}
	else {
	 //mmList.innerHTML='<br><p align=Left>&nbsp; &nbsp; &nbsp; 没有可播放的音乐</div>'; 
	 }
	 }
	 
//Cookie Function
function setCookies(n){
	var cookie_path = location.pathname.replace('cnsidepl.html','');
	var cookie_domain = location.hostname;
	var days = 7;//设定Cookies过期时间
	var expdate = new Date();
	expdate.setTime (expdate.getTime() + (86400 * 1000 * days));
	
	switch (n){
		case 1:
	var songlist = "";
	document.cookie = "mylist=" + songlist + ";expires="+ expdate +"; path="+cookie_path+"; domain="+cookie_domain+"";
	document.cookie = "mytype=" + songlist + ";expires="+ expdate +"; path="+cookie_path+"; domain="+cookie_domain+"";
	changeList();
	break;
	
		case 2:
	var songlist = "1";

	document.cookie = "mytype=" + songlist + ";expires="+ expdate +"; path="+cookie_path+"; domain="+cookie_domain+"";
	changeList();
	break;
	
		case 3:
	var songlist = "2";
	document.cookie = "mytype=" + songlist + ";expires="+ expdate +"; path="+cookie_path+"; domain="+cookie_domain+"";
	changeList();
	break;
	
		case 4:
	var songlist = "3";
	document.cookie = "mytype=" + songlist + ";expires="+ expdate +"; path="+cookie_path+"; domain="+cookie_domain+"";
	changeList();
	break;
	
		case 5:
	var songlist = "4";
	document.cookie = "mytype=" + songlist + ";expires="+ expdate +"; path="+cookie_path+"; domain="+cookie_domain+"";
	changeList();
	break;
	
		case 6:
	var songlist = "5";
	document.cookie = "mytype=" + songlist + ";expires="+ expdate +"; path="+cookie_path+"; domain="+cookie_domain+"";
	changeList();
	break;
	}
}

function strlength(str)
{
	var mx=String.fromCharCode(127);
	if(str.length==0)
	{return(0);}
	else
	{
		var TStr="",l,t=0,i;
		l=str.length;
		t=l;
		for(i=0;i<l;i++)
		{
			if(str.charAt(i)>mx)
			{t+=1;}
		}
		return(t);
	}
}

function LeftTrue(str,n)
{
	var mx=String.fromCharCode(127);
	if(str.length<=n/2)
	{return(str);}
	else
	{
		var TStr="",l,t=0,i;
		l=str.length;
		for(i=0;i<l;i++)
		{
			if(str.charAt(i)>mx)
			{t=t+2;}
			else
			{t=t+1;}
			if(t>n)break;
			TStr=TStr+str.charAt(i);
		}
		return(TStr);
	}
}

function changeList(){
	via.location.reload();
	setTimeout(self.location.reload,3000);
}

function returnfalse()
    {return false; }

    document.oncontextmenu = returnfalse
    document.ondragstart = returnfalse
    document.onselectstart = returnfalse
    
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);