<%

Sub Editor_View(Edt_MiniMode,Form_Content)
%>
<script>
var editFile_dir = "<%=DEF_BBS_HomeUrl%>a/";
</script>
<!-- #include file=post_layer.asp -->
<table border=0 cellpadding=0 cellspacing=0 width="500"><tr><td>
<table border=0 cellpadding=0 cellspacing=0 width="100%">
<tr><td>
<%If Edt_MiniMode = 0 Then%>
<%Else%>
<table border=0 cellpadding=0 cellspacing=0 class="editor_table" width=100%>
<tr><td>
<table border=0 cellpadding=0 cellspacing=0><tr>
<td>
	<div class="layer_item">
	<div class="layer_icon_title"><a href="javascript:;" onclick="return false"><img src="../images/blank.gif" width="22" height="22" title="字体" class="a_pic" style="background-position:0px -44px" /></a></div>
	<div class="layer_iteminfo" id="menu_info_face" onclick="this.style.display='none';">
	<ul class="menu_list">
	<li unselectable=on onclick="addcontent(0,'GLOW','/GLOW','255,RED,2');">发光</li>
	<li unselectable=on onclick="addcontent(0,'FLY','/FLY');">飞行</li>
	<li unselectable=on onclick="addcontent(0,'SHADOW','/SHADOW','255,RED,2');">阴影</li>
	<li unselectable=on onclick="insert('宋体');" style="font-family:宋体">宋体</li>
	<li unselectable=on onclick="insert('黑体');" style="font-family:黑体">黑体</li>
	<li unselectable=on onclick="insert('微软雅黑');" style="font-family:微软雅黑">微软雅黑</li>
	<li unselectable=on onclick="insert('Arial');" style="font-family:Arial">Arial</li>
	<li unselectable=on onclick="insert('Arial Black');" style="font-family:Arial Black">Arial Black</li>
	<li unselectable=on onclick="insert('Century Gothic');" style="font-family:Century Gothic">Century Gothic</li>
	<li unselectable=on onclick="insert('Comic Sans MS');" style="font-family:Comic Sans MS">Comic Sans MS</li>
	<li unselectable=on onclick="insert('Courier');" style="font-family:Courier">Courier</li>
	<li unselectable=on onclick="insert('Courier New');" style="font-family:Courier New">Courier New</li>
	<li unselectable=on onclick="insert('Times New Roman');" style="font-family:Times New Roman">Times New Roman</li>
	<li unselectable=on onclick="insert('Verdana');" style="font-family:Verdana">Verdana</li>
	<li unselectable=on onclick="insert('Impact');" style="font-family:Impact">Impact</li>
	<li unselectable=on onclick="insert('Wingdings');" style="font-family:Wingdings">Wingdings</li>
	</ul>
	</div>
	</div>
</td>
<td>
	<div class="layer_item" unselectable=on>
	<div class="layer_icon_title"><a href="javascript:;" onclick="return false"><img src="../images/blank.gif" class="a_pic" style="background-position:0px -462px;" title="字号" /></a></div>
	<div class="layer_iteminfo" onclick="this.style.display='none';">
	<ul class="menu_list">
	<li unselectable=on onclick="addcontent(0,'size','/size',1);" style="font-size:xx-small;">1</li>
	<li unselectable=on onclick="addcontent(0,'size','/size',2);" style="font-size:x-small">2</li>
	<li unselectable=on onclick="addcontent(0,'size','/size',3);" style="font-size:small">3</li>
	<li unselectable=on onclick="addcontent(0,'size','/size',4);" style="font-size:medium">4</li>
	<li unselectable=on onclick="addcontent(0,'size','/size',5);" style="font-size:large">5</li>
	<li unselectable=on onclick="addcontent(0,'size','/size',6);" style="font-size:x-large">6</li>
	<li unselectable=on onclick="addcontent(0,'size','/size',7);" style="font-size:xx-large;">7</li>
	</ul>
	</div>
	</div>
	</td>
<td width=23 class=ico><a href=#ic title=加粗 onclick="addcontent(0,'B','/B');" class="a_pic" style="background-position:0px -242px;"></td>
<td width=23 class=ico><a href=#ic title=斜体 onclick="addcontent(0,'I','/I');" class="a_pic" style="background-position:0px -726px;"></a></td>
<td width=23 class=ico><a href=#ic title=下划线 onclick="addcontent(0,'U','/U');" class="a_pic" style="background-position:0px -330px;"></a></td>
<td width=23 class=ico><a href=#ic title=中划线 onclick="addcontent(0,'STRIKE','/STRIKE');" class="a_pic" style="background-position:0px -440px;"></a></td>
<td><div class="a_pic" style="background-position:0px -638px;width:10px;height:16px;"></div></td><td>
<td width=23 class=ico><a href=#ic title=剪切 onclick="addcontent(2,'cut');" class="a_pic" style="background-position:0px -132px"></a></td>
<td width=23 class=ico><a href=#ic title=复制 onclick="addcontent(2,'copy');" class="a_pic" style="background-position:0px -176px"></a></td>
<td width=23 class=ico><a href=#ic title=常规粘贴 onclick="addcontent(2,'paste');" class="a_pic" style="background-position:0px -572px;"></a></td>
<td width=23 class=ico><a href=#ic title=删除 onclick="addcontent(2,'delete');" class="a_pic" style="background-position:0px -110px"></a></td>
<td width=23 class=ico><a href=#ic title=清除样式 onclick="addcontent(2,'RemoveFormat');" class="a_pic" style="background-position:0px -506px;"></a></td>
<td><div class="a_pic" style="background-position:0px -638px;width:10px;height:16px;"></div></td><td>
<td width=23 class=ico><a href=#ic title=撤消 onclick="addcontent(2,'undo');" class="a_pic" style="background-position:0px -308px;"></a></td>
<td width=23 class=ico><a href=#ic title=恢复 onclick="addcontent(2,'redo');" class="a_pic" style="background-position:0px -528px;"></a></td>
<td><div class="a_pic" style="background-position:0px -638px;width:10px;height:16px;"></div></td><td>
<td width=23 class=ico><a href=#ic title=引用 onclick="insert('quote');" class="a_pic" style="background-position:0px -550px;"></a></td>
<td width=23 class=ico><a href=#ic title=代码 onclick="insert('code');" class="a_pic" style="background-position:0px -220px;"></a></td>
<td><div class="a_pic" style="background-position:0px -638px;width:10px;height:16px;"></div></td><td>
<td width=23 class=ico><a href=#ic title=插入特殊字符 onclick="editor_view(this,'editor_symbol','symbol.asp?id=56','symbol.js?id=31');" class="a_pic" style="background-position:0px -374px;"></a></td>
<td width=23 class=ico><a href=#ic title=插入分隔线 onclick="addcontent(0,'hr');" class="a_pic" style="background-position:0px 0px"></a></td>
</td></tr></table></td></tr></table>
<table border=0 cellpadding=0 cellspacing=0 width=100%>
<tr height=29>
<td>
<table border=0 cellpadding=0 cellspacing=0><tr height=29 align=center>
<td width=23 class=ico><a href=#ic title=插入表情 onclick="editor_view(this,'editor_icon','','icon.asp');" class="a_pic" style="background-position:0px -88px"></a></td>
<td width=23 class=ico><a href=#ic title=左对齐 onclick="addcontent(0,'ALIGN','/ALIGN','left');" class="a_pic" style="background-position:0px -682px;"></a></td>
<td width=23 class=ico><a href=#ic title=居中对齐 onclick="addcontent(0,'ALIGN','/ALIGN','center');" class="a_pic" style="background-position:0px -704px;"></a></td>
<td width=23 class=ico><a href=#ic title=右对齐 onclick="addcontent(0,'ALIGN','/ALIGN','right');" class="a_pic" style="background-position:0px -660px;"></a></td>
<td><div class="a_pic" style="background-position:0px -638px;width:10px;height:16px;"></div></td><td>
<td width=23 class=ico><a href=#ic title=编号 onclick="addcontent(2,'insertorderedlist');" class="a_pic" style="background-position:0px -748px;"></a></td>
<td width=23 class=ico><a href=#ic title=项目符号 onclick="addcontent(2,'insertunorderedlist');" class="a_pic" style="background-position:0px -770px;"></a></td>
<td width=23>
	<div class="layer_item">
	<div class="layer_icon_title"><a href="javascript:;" onclick="return false"><img src="../images/blank.gif" class="a_pic" style="background-position:0px -616px;" title="行距" /></a></div>
	<div class="layer_iteminfo" id="menu_info_lineheight" onclick="this.style.display='none';">
	<ul class="menu_list">
	<li unselectable=on onclick="addcontent(3,'line-height','1');">100%</li>
	<li unselectable=on onclick="addcontent(3,'line-height','0.5');">50%</li>
	<li unselectable=on onclick="addcontent(3,'line-height','1.5');">150%</li>
	<li unselectable=on onclick="addcontent(3,'line-height','2');">200%</li>
	<li unselectable=on onclick="addcontent(3,'line-height','2.5');">250%</li>
	<li unselectable=on onclick="addcontent(3,'line-height','3');">300%</li>
	<li unselectable=on onclick="addcontent(3,'line-height','4');">400%</li>
	</ul>
	</div>
	</div>

</td>
<td><div class="a_pic" style="background-position:0px -638px;width:10px;height:16px;"></div></td><td>
<td width=23 class=ico><a href=#ic title=上标 onclick="addcontent(0,'SUP','/SUP');" class="a_pic" style="background-position:0px -396px;"></a></td>
<td width=23 class=ico><a href=#ic title=下标 onclick="addcontent(0,'SUB','/SUB');" class="a_pic" style="background-position:0px -418px;"></a></td>
<td><div class="a_pic" style="background-position:0px -638px;width:10px;height:16px;"></div></td><td>
<td width=23 class=ico><a href=#ic title=字体颜色 onclick="editor_sAction = 'forecolor';editor_view(this,'editor_selcolor','','selcolor.js');" class="a_pic" style="background-position:0px -66px"></a></td>
<td width=23 class=ico><a href=#ic title=字体背景颜色 onclick="editor_sAction = 'backcolor';editor_view(this,'editor_selcolor','','selcolor.js');" class="a_pic" style="background-position:0px -264px;"></a></td>
<td><div class="a_pic" style="background-position:0px -638px;width:10px;height:16px;"></div></td><td>
<td width=23 class=ico><a href=#ic title=插入或修改超级链接 onclick="edt_link();" class="a_pic" style="background-position:0px -154px"></a></td>
<td width=23 class=ico><a href=#ic title=取消超级链接或标签 onclick="addcontent(2,'UnLink');" class="a_pic" style="background-position:0px -286px;"></a></td>
<td><div class="a_pic" style="background-position:0px -638px;width:10px;height:16px;"></div></td><td>
<td width=23 class=ico><a href=#ic unselectable=on title=插入或修改表格 onclick="if(typeof(editor_inittable)=='function')editor_inittable();editor_view(this,'editor_insttable','table.asp','table.js');" class="a_pic" style="background-position:0px -352px;"></a></td>
<td width=23 class=ico><a href=#ic title=插入或修改图片 onclick="if(typeof(editor_InitimgDocument)=='function')editor_InitimgDocument();editor_view(this,'editor_img','img.asp','img.js');" class="a_pic" style="background-position:0px -22px"></a></td>
<td width=23 class=ico><a href=#ic title=插入多媒体 onclick="if(typeof(editor_Initmedia)=='function')editor_Initmedia();editor_view(this,'editor_media','media.asp','media.js');" class="a_pic" style="background-position:0px -594px;"></a></td>
<td><div class="a_pic" style="background-position:0px -638px;width:10px;height:16px;"></div></td><td>
<td width=23 class=ico><a href=#ic title=另存为文件 onclick="addcontent(2,'SaveAs');" class="a_pic" style="background-position:0px -484px;"></a></td>
</tr></table></td></tr></table>

<%End If%>
</td></tr>
</table>
</td></tr></table>

<div class=editor><textarea cols=80 style="width: 100%;height:220px; word-break: break-all;" id=Form_Content name=Form_Content rows=16 ONSELECT="storeCaret(this);" onclick="storeCaret(this);" ONKEYUP="storeCaret(this);" onkeydown="if(ctlkey(event)==false)return(false);"><%If Form_Content <> "" Then Response.Write VbCrLf & Server.htmlEncode(Form_Content)%></textarea></div>
<div class=editor_choose><table border="0" cellPadding="0" cellSpacing="0" width="100%">
	<tr>
	<td align="left" valign="top">
	<div id="LEADEDT_TXT" style="display:block;">
	<map name="LEADEDT_Map1">
	<area shape="polygon" coords="5, 3, 12, 14, 43, 14, 49, 6, 43, 0" alt="纯文本和编码编辑模式" onclick="edt_setmode(1);">
	<area shape="polygon" coords="87, 14, 91, 5, 87, 0, 50, 0, 46, 9, 49, 14" alt="HTML代码编辑模式" onclick="edt_setmode(0);">
	</map> <img src="../images/blank.gif" class="a_editmode a_modeedit" usemap="#LEADEDT_Map1" border="0"></div>

	<div id="LEADEDT_EDIT" style="display:none">
	<map name="LEADEDT_Map2">
	<area shape="polygon" coords="5, 3, 12, 14, 43, 14, 49, 6, 43, 0" alt="纯文本和编码编辑模式" onclick="edt_setmode(1);">
	<area shape="polygon" coords="87, 14, 91, 5, 87, 0, 50, 0, 46, 9, 49, 14" alt="HTML代码编辑模式" onclick="edt_setmode(0);">
	</map> <img src="../images/blank.gif" class="a_editmode a_modetext" usemap="#LEADEDT_Map2" border="0"></div>
	</td>
	<td width=100 align=right>
	<a href=#icon onclick="edt_htsub();" title=缩短编缉区><b>-</b></a>
	<a href=#icon onclick="edt_htresume();" title=恢复编缉区><b>=</b></a>
	<a href=#icon onclick="edt_htadd();" title=增长编缉区>+</a>
	</td>
	</tr>
	</table>
</div>
<%End Sub%>
