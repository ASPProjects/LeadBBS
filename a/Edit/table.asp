<!-- #include file=../../inc/BBSsetup.asp -->
<table border=0 cellspacing=0 unselectable=on style="cursor: move;padding:5px 0px 5px 0px;" onmousedown="LD.move.mousedown(this.parentNode,event);">
<tr><td align=right unselectable=on>
<div class="layer_close"><a href="javascript:;" onclick="LD.hide('editor_insttable');return false;" class="unsel" hidefocus="true" title="close" /></a></div>
</td></tr>
<tr>
	<td>
	<fieldset unselectable=on>
	<legend unselectable=on>����С</legend>
	<table border=0 cellpadding=5 cellspacing=0 unselectable=on>
	<tr>
	<td noWrap unselectable=on>
		����
		</td><td unselectable=on><input type=text id=d_row size=5 value="" onkeypress="event.returnValue=IsDigit(event);" maxlength=3 class="fminpt input_1">
		</td><td noWrap unselectable=on>����
		</td><td unselectable=on><input type=text id=d_col size=5 value="" onkeypress="event.returnValue=IsDigit(event);" maxlength=3 class="fminpt input_1">
		</td><td noWrap unselectable=on>����
		</td><td unselectable=on><input type=text id="d_bgurl" style="width:98px" size=13 value="" class="fminpt input_2">
	</td>
	</tr>
	</table>
	</fieldset>
	</td>
</tr>
<tr>
	<td>
	<fieldset>
	<legend>��񲼾�</legend>
	<table border=0 cellpadding=5 cellspacing=0>
	<tr>
	<td>
		���뷽ʽ:
		<select id="d_align">
			<option value=''>Ĭ��</option>
			<option value='left'>�����</option>
			<option value='center'>����</option>
			<option value='right'>�Ҷ���</option>
			</select>
		�߿��ϸ:
		<input type=text id=d_border size=10 value="" onkeypress="event.returnValue=IsDigit(event);" class="fminpt input_1">
	</td>
	</tr>
	<tr>
		<td>��Ԫ���:
		<input type=text id=d_cellspacing size=10 value="" onkeypress="event.returnValue=IsDigit(event);" maxlength=3 class="fminpt input_1">
		��Ԫ�߾�:
		<input type=text id=d_cellpadding size=10 value="" onkeypress="event.returnValue=IsDigit(event);" maxlength=3 class="fminpt input_1">
	</td>
	</tr>
	</table>
	</fieldset>
	</td>
</tr>
<tr>
	<td>
	<fieldset>
	<legend>�����</legend>
	<table border=0 cellpadding=5 cellspacing=0 width='100%'>
	<tr>
		<td><input id="d_check" type="checkbox" onclick="d_widthvalue.disabled=(!this.checked);d_widthunit.disabled=(!this.checked);" value="1" class="fmchkbox">ָ�����Ŀ��
			<input name="d_widthvalue" id="d_widthvalue" type="text" value="" size="5" onkeypress="event.returnValue=IsDigit(event);" maxlength="4" class="fminpt input_1">
			<select name="d_widthunit" id="d_widthunit">
			<option value='px'>����</option><option value='%'>�ٷֱ�</option>
			</select>
		</td>
	</tr>
	</table>
	</fieldset>
	</td>
</tr>
<tr>
	<td>
	<fieldset>
	<legend>�����ɫ</legend>
	<table border=0 cellpadding=0 cellspacing=0 style="padding:5px;">
	<tr>
		<td>�߿���ɫ:
		<input type=text id=d_bordercolor size=7 value="" class="fminpt input_1">
		<a href=#ic onclick="editor_sAction = 'bordercolor';editor_view(this,'editor_selcolor','','selcolor.js');" class=ico>
		<img src="../images/blank.gif" class="a_pic" style="background-position:0px -198px;width:18px;height:17px;" style="cursor:pointer" align=absmiddle id=s_bordercolor></a>
		������ɫ:
		<input type=text id=d_bgcolor size=7 value="" class="fminpt input_1">
		<a href=#ic onclick="editor_sAction = 'bgcolor';editor_view(this,'editor_selcolor','','selcolor.js');" class=ico>
		<img src="../images/blank.gif" class="a_pic" style="background-position:0px -198px;width:18px;height:17px;" align=absmiddle id=s_bgcolor></a>
		</td>
	</tr>
	</table>
	</fieldset>
	</td>
</tr>
<tr><td align=right><span class="clicktext" onclick="editor_insttablesubmit();">����/���±��</span></td></tr>
</table>