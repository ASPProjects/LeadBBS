<!-- #include file=../../inc/BBSsetup.asp -->
<table border=0 cellpadding=0 cellspacing=0 style="cursor: move;" onmousedown="LD.move.mousedown(this.parentNode,event);">
<tr>
	<td style="padding:0px 0px 5px 0px;">
	<div align="right" style="width:100%;float:left">
	<div class="layer_close"><a href="javascript:;" onclick="LD.hide('editor_img');return false;" class="unsel" hidefocus="true" title="close" /></a></div>
	</div>
	<br>
	<fieldset>
	<legend>ͼƬ��Դ</legend>
	<table border=0 cellpadding=5 cellspacing=0>
	<tr>
		<td noWrap>�����ַ:</td>
		<td><input type=text id="img_d_fromurl" style="width:243px" size=255 value="" class="fminpt input_3">
	</tr>
	</table>
	</fieldset>
	</td>
</tr>
<tr>
	<td style="padding:0px 0px 5px 0px;">
	<fieldset>
	<legend>��ʾЧ��</legend>
	<table border=0 cellpadding=5 cellspacing=0>
	<tr>
		<td noWrap>�߿��ϸ:</td>
		<td><input type=text id=img_d_border size=10 value="" ONKEYPRESS="event.returnValue=IsDigit(event);" class="fminpt input_1"></td>
		<td noWrap>���뷽ʽ:</td>
		<td>
			<select id=img_d_align>
			<option value='absmiddle' selected>���Ծ���</option>
			<option value='left'>����</option>
			<option value='right'>����</option>
			<option value='top'>����</option>
			<option value='middle'>�в�</option>
			<option value='bottom'>�ײ�</option>
			<option value='absbottom'>���Եײ�</option>
			<option value='baseline'>����</option>
			<option value='texttop'>�ı�����</option>
			</select></td>
	</tr>
	</table>
	</fieldset>
	</td>
</tr>


<tr>
	<td style="padding:0px 0px 5px 0px;">
	<fieldset>
	<legend>ͼƬ��С</legend>
	<table border=0 cellpadding=5 cellspacing=0>
	<tr>
		<td noWrap>���:</td>
		<td><input type=text id=img_d_width size=10 value="" ONKEYPRESS="event.returnValue=IsDigit(event);" class="fminpt input_1"></td>
		<td noWrap>�߶�:</td>
		<td><input type=text id=img_d_height size=10 value="" ONKEYPRESS="event.returnValue=IsDigit(event);" class="fminpt input_1"></td>
	</tr>
	</table>
	</fieldset>
	</td>
</tr>

<tr><td align=right>
<span class="clicktext" onclick="editor_imgok();">ȷ��</span>
</td></tr>
</table>