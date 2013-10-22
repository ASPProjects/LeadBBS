var sAction = "INSERT";
var sTitle = "����";

var oControl;
var oSeletion;
var sRangeType;

var sRow = "2";
var sCol = "2";
var sAlign = "";
var sBorder = "1";
var sCellPadding = "3";
var sCellSpacing = "2";
var sWidth = "";
var sBorderColor = "#000000";
var sBgColor = "#FFFFFF";

var sWidthUnit = "%"
var bCheck = true;
var bWidthDisable = false;
var sWidthValue = "100";
var sBackGround = "";


function editor_inittable()
{
	sAction = "INSERT";
	sTitle = "����";
	
	sRow = "2";
	sCol = "2";
	sAlign = "";
	sBorder = "1";
	sCellPadding = "3";
	sCellSpacing = "2";
	sWidth = "";
	sBorderColor = "#000000";
	sBgColor = "#FFFFFF";
	
	sWidthUnit = "%"
	bCheck = true;
	bWidthDisable = false;
	sWidthValue = "100";
	sBackGround = "";
	if(!isUndef(edt_doc) && edt_doc!=null)
	{
		if(Browser.ie)
		{
			oSelection = edt_doc.selection.createRange();
			sRangeType = edt_doc.selection.type;
		}
		else
		{
			oSelection = edt_win.getSelection();
			sRangeType = oSelection.type;
		}
		
		if (sRangeType == "Control")
		{
			if (oSelection.item(0).tagName == "TABLE")
			{
				sAction = "MODI";
				sTitle = "�޸�";
				oControl = oSelection.item(0);
				sRow = oControl.rows.length;
				sCol = getColCount(oControl);
				sAlign = oControl.align;
				sBorder = oControl.border;
				sCellPadding = oControl.cellPadding;
				sCellSpacing = oControl.cellSpacing;
				sWidth = oControl.width;
				sBorderColor = oControl.borderColor;
				sBgColor = oControl.bgColor;
				sBackGround = oControl.background;
			}
		}
		editor_tableInitDocument();
	}
	else
	{
		setTimeout("editor_inittable()",100);
	}
}
editor_inittable();

// ��ʼֵ
function editor_tableInitDocument(){
	SearchSelectValue($id("d_align"), sAlign.toLowerCase());

	// �޸�״̬ʱȡֵ
	if (sAction == "MODI"){
		if (sWidth == ""){
			bCheck = false;
			bWidthDisable = true;
			sWidthValue = "100";
			sWidthUnit = "%";
		}else{
			bCheck = true;
			bWidthDisable = false;
			if (sWidth.substr(sWidth.length-1) == "%"){
				sWidthValue = sWidth.substring(0, sWidth.length-1);
				sWidthUnit = "%";
			}else{
				sWidthUnit = "";
				sWidthValue = parseInt(sWidth);
				if (isNaN(sWidthValue)) sWidthValue = "";
			}
		}
	}

	switch(sWidthUnit){
	case "%":
		$id("d_widthunit").selectedIndex = 1;
		break;
	default:
		sWidthUnit = "";
		$id("d_widthunit").selectedIndex = 0;
		break;
	}

	$id("d_row").value = sRow;
	$id("d_col").value = sCol;
	$id("d_border").value = sBorder;
	$id("d_cellspacing").value = sCellSpacing;
	$id("d_cellpadding").value = sCellPadding;
	$id("d_widthvalue").value = sWidthValue;
	$id("d_widthvalue").disabled = bWidthDisable;
	$id("d_widthunit").disabled = bWidthDisable;
	$id("d_bordercolor").value = sBorderColor;
	$id("s_bordercolor").style.backgroundColor = sBorderColor;
	$id("d_bgcolor").value = sBgColor;
	$id("s_bgcolor").style.backgroundColor = sBgColor;
	$id("d_check").checked = bCheck;
	$id("d_bgurl").value = sBackGround;


}

// �ж�ֵ�Ƿ����0
function MoreThanOne(obj, sErr){
	var b=false;
	if (obj.value!=""){
		obj.value=parseFloat(obj.value);
		if (obj.value!="0"){
			b=true;
		}
	}
	if (b==false){
		BaseAlert(obj,sErr);
		return false;
	}
	return true;
}

// �õ��������
function getColCount(oTable) {
	var intCount = 0;
	if (oTable != null) {
		for(var i = 0; i < oTable.rows.length; i++){
			if (oTable.rows[i].cells.length > intCount) intCount = oTable.rows[i].cells.length;
		}
	}
	return intCount;
}

// ������
function InsertRows( oTable ) {
	if ( oTable ) {
		var elRow=oTable.insertRow();
		for(var i=0; i<oTable.rows[0].cells.length; i++){
			var elCell = elRow.insertCell();
			elCell.innerHTML = "&nbsp;";
		}
	}
}

// ������
function InsertCols( oTable ) {
	if ( oTable ) {
		for(var i=0; i<oTable.rows.length; i++){
			var elCell = oTable.rows[i].insertCell();
			elCell.innerHTML = "&nbsp;"
		}
	}
}

// ɾ����
function DeleteRows( oTable ) {
	if ( oTable ) {
		oTable.deleteRow();
	}
}

// ɾ����
function DeleteCols( oTable ) {
	if ( oTable ) {
		for(var i=0;i<oTable.rows.length;i++){
			oTable.rows[i].deleteCell();
		}
	}
}


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

// ������Ϣ��ʾ���õ����㲢ѡ��
function BaseAlert(theText,notice){
	alert(notice);
	theText.focus();
	theText.select();
	return false;
}
function editor_insttablesubmit()
{
	// �߿���ɫ����Ч��
	sBorderColor = $id("d_bordercolor").value;
	sBgColor = $id("d_bgcolor").value;
	// ��������Ч��
	if (!MoreThanOne($id("d_row"),'��Ч������������Ҫ1�У�')) return;
	// ��������Ч��
	if (!MoreThanOne($id("d_col"),'��Ч������������Ҫ1�У�')) return;
	// ���ߴ�ϸ����Ч��
	if ($id("d_border").value == "") $id("d_border").value = "0";
	if ($id("d_cellpadding").value == "") $id("d_cellpadding").value = "0";
	if ($id("d_cellspacing").value == "") $id("d_cellspacing").value = "0";
	// ȥǰ��0
	$id("d_border").value = parseFloat($id("d_border").value);
	$id("d_cellpadding").value = parseFloat($id("d_cellpadding").value);
	$id("d_cellspacing").value = parseFloat($id("d_cellspacing").value);
	// �����Чֵ��
	var sWidth = "";
	if ($id("d_check").checked){
		if (!MoreThanOne($id("d_widthvalue"),'��Ч�ı���ȣ�')) return;
		sWidth = $id("d_widthvalue").value + $id("d_widthunit").value;
	}

	sRow = $id("d_row").value;
	sCol = $id("d_col").value;
	sAlign = $id("d_align").options[$id("d_align").selectedIndex].value;
	sBorder = $id("d_border").value;
	sCellPadding = $id("d_cellpadding").value;
	sCellSpacing = $id("d_cellspacing").value;
	sBackGround = $id("d_bgurl").value;

	if (sAction == "MODI") {
		// �޸�����
		var xCount = sRow - oControl.rows.length;
  		if (xCount > 0)
	  		for (var i = 0; i < xCount; i++) InsertRows(oControl);
  		else
	  		for (var i = 0; i > xCount; i--) DeleteRows(oControl);
		// �޸�����
  		var xCount = sCol - getColCount(oControl);
  		if (xCount > 0)
  			for (var i = 0; i < xCount; i++) InsertCols(oControl);
  		else
  			for (var i = 0; i > xCount; i--) DeleteCols(oControl);

		try {
			oControl.width = sWidth;
		}
		catch(e) {
			//alert("�Բ�������������Ч�Ŀ��ֵ��\n���磺90%  200  300px  10cm��");
		}

		oControl.align			= sAlign;
  		oControl.border			= sBorder;
  		oControl.cellSpacing	= sCellSpacing;
  		oControl.cellPadding	= sCellPadding;
  		oControl.borderColor	= sBorderColor;
  		oControl.bgColor		= sBgColor;
  		oControl.background     = sBackGround;

	}else{
		if(sBorderColor=="")sBorderColor="#000000";
		if(sCellPadding=="")sBorderColor="0";
		if(sCellSpacing=="")sCellSpacing="0";
		if(sWidth=="")sWidth="0";
		if(sAlign=="")sAlign="center";
		if(sBgColor=="")sBgColor="#FFFFFF";
		if(sBorder=="")sBorder="0";
		if(sBackGround=="")sBackGround="#";
		var sTable = "<table align='"+sAlign+"' border='"+sBorder+"' cellpadding='"+sCellPadding+"' cellspacing='"+sCellSpacing+"' width='"+sWidth+"' bordercolor='"+sBorderColor+"' bgcolor='"+sBgColor+"'";
		var ubbTable = "[TABLE=" + sBorderColor + "," + sCellSpacing + "," + sCellPadding + "," + sWidth + "," + sAlign + "," + sBgColor + "," + sBorder + "," + sBackGround + "]"
		if(sBackGround != "")sTable += " background='" + sBackGround + "'>";
		for (var i=1;i<=sRow;i++){
			sTable = sTable + "<tr>";
			ubbTable = ubbTable + "[TR]"
			for (var j=1;j<=sCol;j++){
				ubbTable = ubbTable + "[TD] [/TD]"
				sTable = sTable + "<td>&nbsp;</td>";
			}
		}
		sTable = sTable + "</table>";
		ubbTable = ubbTable + "[/TABLE]"
		if(!edt_mode)
		{addcontent(2,sTable);}
		else
		{addcontent(1,ubbTable);}
	}
}