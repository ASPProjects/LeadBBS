<!-- #include file=../../inc/BBSsetup.asp -->
<table border=0 cellpadding=0 cellspacing=0 style="cursor: move;" onmousedown="LD.move.mousedown(this.parentNode,event);">
<tr>
	<td>
<div align="right" style="width:100%;float:left">
<div class="layer_close"><a href="javascript:;" onclick="LD.hide('editor_media');return false;" class="unsel" hidefocus="true" title="close" /></a></div>
</div>
<br>
	<fieldset style="padding:5px;margin:2px 0px 5px 0px">
	<legend>插入媒体文件</legend>
	<table border=0 cellpadding=5 cellspacing=0>
	<tr>
		<td noWrap>
		媒体类型: <select id="media_d_type" style="width:72px">
			<option value='FLASH'>FLASH</option>
			<option value='MP'>MEDIA文件</option>
			<option value='RM'>RM文件</option>
			<option value='FLV'>FLV文件</option>
		</select> 允许toudu,youku,youtube等站点
		</td>
	</tr>
	<tr>
		<td noWrap>网络地址:
		<input type=text id="media_d_fromurl" style="width:243px" size=30 value="http://" class="fminpt input_3"></td>
	</tr>
	<tr>
		<td noWrap>显示宽度: <input type=text id=media_d_width size=10 value="" onkeypress="event.returnValue=IsDigit(event);" maxlength=4 class="fminpt input_1">
		显示高度: <input type=text id=media_d_height size=10 value="" onkeypress="event.returnValue=IsDigit(event);" maxlength=4 class="fminpt input_1"></td>
	</tr>
	</table>
	</fieldset>
	</td>
</tr>
<tr><td align=right>
<span class="clicktext" onclick="editor_mediasubmit();">确定</span></td></tr>
</table>