<!-- #include file=../../inc/BBSsetup.asp -->
<table border=0 cellpadding=0 cellspacing=0 style="cursor: move;" onmousedown="LD.move.mousedown(this.parentNode,event);">
<tr>
	<td>
<div align="right" style="width:100%;float:left">
<div class="layer_close"><a href="javascript:;" onclick="LD.hide('editor_media');return false;" class="unsel" hidefocus="true" title="close" /></a></div>
</div>
<br>
	<fieldset style="padding:5px;margin:2px 0px 5px 0px">
	<legend>����ý���ļ�</legend>
	<table border=0 cellpadding=5 cellspacing=0>
	<tr>
		<td noWrap>
		ý������: <select id="media_d_type" style="width:72px">
			<option value='FLASH'>FLASH</option>
			<option value='MP'>MEDIA�ļ�</option>
			<option value='RM'>RM�ļ�</option>
			<option value='FLV'>FLV�ļ�</option>
		</select> ����toudu,youku,youtube��վ��
		</td>
	</tr>
	<tr>
		<td noWrap>�����ַ:
		<input type=text id="media_d_fromurl" style="width:243px" size=30 value="http://" class="fminpt input_3"></td>
	</tr>
	<tr>
		<td noWrap>��ʾ���: <input type=text id=media_d_width size=10 value="" onkeypress="event.returnValue=IsDigit(event);" maxlength=4 class="fminpt input_1">
		��ʾ�߶�: <input type=text id=media_d_height size=10 value="" onkeypress="event.returnValue=IsDigit(event);" maxlength=4 class="fminpt input_1"></td>
	</tr>
	</table>
	</fieldset>
	</td>
</tr>
<tr><td align=right>
<span class="clicktext" onclick="editor_mediasubmit();">ȷ��</span></td></tr>
</table>