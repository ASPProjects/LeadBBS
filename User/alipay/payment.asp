<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Board_Popfun.asp -->
<!-- #include file=../../inc/Limit_Fun.asp -->
<!-- #include file=../inc/UserTopic.asp -->
<!--#include file="alipayto/alipay_payto.asp"-->
<%
DEF_BBS_HomeUrl = "../../"
Main

Sub Main

	initDatabase
	GBL_CHK_TempStr = ""
	
	BBS_SiteHead DEF_SiteNameString & " - " & DEF_PointsName(1) & "��ֵ",0,"<span class=navigate_string_step>" & DEF_PointsName(1) & "��ֵ</span>"
	
	UserTopicTopInfo("user")
	If GBL_CHK_Flag = 1 Then
		LoginAccuessFul
	Else
		If Request("submitflag")="" Then
			DisplayLoginForm("���ȵ�¼")
		Else
			DisplayLoginForm(GBL_CHK_TempStr)
		End If
	End If
	closeDataBase
	UserTopicBottomInfo
	SiteBottom

End Sub

Function LoginAccuessFul

              Dim act
              act = Left(Request.QueryString("act"),20)
              Payment_nav(act)
              
              Select Case act
              	Case "done":
              		Response.Write "<b>" & Request.QueryString("str") & "</b>"
              	Case "list":
              		Alipay_SellList
              	Case Else
              		DisplayPaymentForm
              End Select
    
End Function

Sub Payment_nav(Evol)

	Response.Write "<div class='user_item_nav fire'><ul>"
	Response.Write "<li><div class=name>" & DEF_PointsName(1) & "��ֵ</div></li>"
	If Evol = "list" Then
		%><li><div class=navactive>��ʷ����</div></li>
		<li><a href=payment.asp>������ֵ</a></li><%
	Else%>
		<li><a href=payment.asp?act=list>��ʷ����</a></li>
		<li><div class=navactive>������ֵ</div></li><%
	End If
	Response.Write "</ul></div>"
	

End Sub

Dim shijian,dingdan,subject,body,out_trade_no,price,quantity,discount,AlipayObj,itemUrl

Sub PaymentSubmit

	If DEF_seller_email = "" Then
		Response.Write "<div class=alert>��վδ��ͨ�������.</div>"
		Exit Sub
	End If

	GBL_CHK_TempStr = ""
	If CheckUserAnnounceLimit = 0 Then
		Response.Write "<div class=alert>" & GBL_CHK_TempStr & "</div>"
		Exit Sub
	End If
	Dim Rs,PreSellTime
	Set Rs = LdExeCute(sql_select("Select PID,SellTime,UserName from LeadBBS_SellList where UserName='" & Replace(GBL_CHK_User,"'","''") & "' and PayFlag = 0 Order by ID DESC",1),0)
	If Not Rs.Eof Then
		PreSellTime = cCur(Rs(1))
	Else
		PreSellTime = 0
	End If
	Rs.Close
	Set Rs = Nothing
	
	If PreSellTime > 0 Then
		If DateDiff("s",RestoreTime(GBL_UDT(13)),DEF_Now) < 5 Then
			GBL_CHK_TempStr = "�����ύ̫Ƶ,���Ժ����ύ."
			Exit Sub
		End If
	End If
	
	'ɾ����δ�����(7��ǰ)
	CALL LdExeCute("Delete from LeadBBS_SellList where UserName='" & Replace(GBL_CHK_User,"'","''") & "' and PayFlag = 0 and SellTime<" & GetTimeValue(DateAdd("d",-7,DEF_Now)) & "",1)

	Dim PayPoints
	PayPoints = Left(Request.Form("PayPoints"),14)
	If isNumeric(PayPoints) = False Then PayPoints = 0
	PayPoints = Fix(cCur(PayPoints))
	If PayPoints < DEF_seller_minpoints Then
		GBL_CHK_TempStr = "���γ�ֵ��������� " & DEF_seller_minpoints & " ��"
		Exit Sub
	End If
	
	shijian=now()
	dingdan= GetTimeValue(DEF_Now)
	'�ͻ���վ�����ţ�����ȡϵͳʱ�䣬�ɸĳ���վ�Լ��ı�����
	
	Dim LoopN
	LoopN = 0
	For LoopN = 0 To 9
		Set Rs = LdExeCute(sql_select("Select ID from LeadBBS_SellList where PID=" & dingdan & Right("00" & LoopN,1),1),0)
		If Rs.Eof Then
			dingdan = cCur(dingdan & Right("00" & LoopN,1))
			Rs.Close
			Set Rs = Nothing
			Exit For
		Else
			Rs.Close
			Set Rs = Nothing
		End If
	Next
		
	
	'subject			=	DEF_PointsName(1) & "��ֵ"		'��Ʒ����
	'body			=	"��̳" & DEF_PointsName(1) & "��ֵ����" & PayPoints*DEF_seller_exchangescale & "��"		'body			��Ʒ����
	subject			=	dingdan		'��Ʒ����
	body			=	"value" & PayPoints*DEF_seller_exchangescale & ""		'body			��Ʒ����
	out_trade_no    =   dingdan
	price		    =	PayPoints				'price��Ʒ����			0.01��50000.00
	quantity        =   "1"               '��Ʒ����,����߹��ﳵĬ��Ϊ1
	discount        =   "0"               '��Ʒ�ۿ�
	Set AlipayObj	= New creatAlipayItemURL
	itemUrl=AlipayObj.creatAlipayItemURL(subject,body,out_trade_no,price,quantity,seller_email)
	
	'��ӱ��涩��
	CALL LdExeCute("insert into LeadBBS_SellList(PID,UserName,PayPoints,GetPoints,SellTime,PayFlag) Values(" &_
		dingdan &_
		",'" & Replace(GBL_CHK_User,"'","''") & "'" &_
		"," & PayPoints & "" &_
		"," & PayPoints*DEF_seller_exchangescale & "" &_
		"," & GetTimeValue(DEF_Now) & "" &_
		",0)",1)
	%>
	<form id="PaymentForm" method="post" action="<%=itemUrl%>" method=get target="_blank">
	<table cellspacing="0" cellpadding="0" class=blanktable>
	<tr>
	<td><b>�����˻�</b></td>
	<td><%=htmlencode(GBL_CHK_User)%></td>
	</tr>
	<tr>
	<td><b>��ֵ���</b></td>
	<td><%=PayPoints%>Ԫ</td>
	</tr>
	<tr>
	<td><b>����<%=DEF_PointsName(1)%>����</b></td>
	<td><%=PayPoints*DEF_seller_exchangescale%>��</td>
	</tr>
	<tr>
	<td><b>������</b></td>
	<td><%=dingdan%></td>
	</tr>
	<tr>
	<td colspan=2><br/>��ȷ��������Ϣ��������°�ť�ύ����֧��
	<br><br>
	<input type="submit" value=����֧�� class="fmbtn btn_3" />
	</td></tr></table>
	</form>
	<%

End Sub

Sub DisplayPaymentForm

	GBL_CHK_TempStr = ""
	If Request.Form("action") = "submitpayment" Then
		PaymentSubmit
		If GBL_CHK_TempStr <> "" Then
			Response.Write "<div class=alert>������ʾ��" & GBL_CHK_TempStr & "</div>"
		Else
			Exit Sub
		End If
	End If
	%>
	<form id="PaymentForm" method="post" action="payment.asp">
	<input type="hidden" name="action" value="submitpayment">
	<table cellspacing="0" cellpadding="0" class=blanktable>
	<tr>
	<td><b>��ֵ<%=DEF_PointsName(1)%>����</b></td>
	<td>
	������ֽ� <b>1</b> Ԫ = <%=DEF_PointsName(1)%> <b><%=1*DEF_seller_exchangescale%></b> ��	<br>������ͳ�ֵ <%=DEF_PointsName(1)%> <b><%=DEF_seller_minpoints%></b> ��
	</td>
	</tr>
	<tr>
	<td><b>��ֵ����ҽ��</b></td>
	<td>
	<script>
	function $id(id)
	{
		return(document.getElementById(id));
	}
	function refreshMoney()
	{
		var PayPoints = parseInt($id('PayPoints').value);
		if(isNaN(PayPoints))
		{
			PayPoints = 0;
			alert("��ֵ�����ʹ������.");
			$id('PayPoints').value="";
		}
		else
		{
			$id('PayPoints').value=PayPoints;
		}
		$id('payment_money').innerHTML = PayPoints;
		$id('payment_points').innerHTML = PayPoints*<%=DEF_seller_exchangescale%>;
	}
	</script>
	<input name=PayPoints id=PayPoints onchange="refreshMoney()" class="fminpt input_2" /> Ԫ</td>
	</tr>

	<tr>
	<td><b>������Ҫ����֧��</b></td>
	<td>����� <span id="payment_money">0</span> Ԫ �ɹ���<%=DEF_PointsName(1)%>���� <span id="payment_points">0</span> ��</td>
	</tr>

	<tr>
	<td></td>
	<td>
	<div class=value2>
	�����ֵ������������ֽ�����֧����������̳������ʹ�õ�����ң�<%=DEF_PointsName(1)%><br>
	���ֳ�ֵ���ܳ������˿�����ڳ�ֵǰȷ������ϸ�˶Գ�ֵ�Ľ�
	</div>
	<br />
	<div class=value2>
	���ɹ�֧����ϵͳ������Ҫ�����ӵ�ʱ��ȴ�֧���������˿����޷�˲�����ˡ�<br>
	�������48Сʱ��δ�յ�֪ͨ����Ϣ��������̳����Ա��ϵ��
	</div></td>
	</tr>
	<tr class="btns">
		<td>&nbsp;</td>
		<td>
			<input type="submit" value=�ύ class="fmbtn btn_2" /></td>
	</tr>
	</table>

	</form>
	<%

End Sub

Sub Alipay_SellList

	Dim Rs,payflag,SQL
	payflag = Request.QueryString("payflag")
	If CheckSupervisorUserName = 0 Then
		If payflag = "0" Then
			payflag = 0
			SQL = " where UserName='" & Replace(GBL_CHK_User,"'","''") & "' and PayFlag=0"
		ElseIf payflag = "1" Then
			payflag = 1
			SQL = " where UserName='" & Replace(GBL_CHK_User,"'","''") & "' and PayFlag=1"
		Else
			payflag = -1
			SQL = " where UserName='" & Replace(GBL_CHK_User,"'","''") & "'"
		End If
	Else
		If payflag = "0" Then
			payflag = 0
			SQL = " where PayFlag=0"
		ElseIf payflag = "1" Then
			payflag = 1
			SQL = " where PayFlag=1"
		Else
			payflag = -1
			SQL = ""
		End If
	End If
	Dim GetData,Count
	SQL = sql_select("Select ID,PID,UserName,PayPoints,GetPoints,SellTime,PayFlag From LeadBBS_SellList" & SQL & " Order by ID DESC",150)
	Set Rs = LdExeCute(SQL,0)
	If Not Rs.Eof Then
		GetData = Rs.GetRows(-1)
		Count = Ubound(GetData,2)
	Else
		Count = -1
	End If
	Rs.Close
	Set Rs = Nothing
	%>
	<table width="100%" border="0" cellspacing="0" cellpadding="0" class=table_in>
	<tr class=tbinhead>
		<td><div class=value>������</div></td>
		<td><div class=value>�û�</div></td>
		<td><div class=value>��(Ԫ)</div></td>
		<td><div class=value>�һ�ֵ</div></td>
		<td><div class=value>ʱ��</div></td>
		<td><div class=value>״̬</div></td>
	</tr>
	<%
	Dim N
	If Count >= 0 Then
	For N = 0 to Count
		%>
	<tr>
		<td class=tdbox><%=GetData(1,N)%></td>
		<td class=tdbox><%=htmlencode(GetData(2,N))%></td>
		<td class=tdbox><%=GetData(3,N)%></td>
		<td class=tdbox><%=GetData(4,N)%></td>
		<td class=tdbox><%=RestoreTime(GetData(5,N))%></td>
		<td class=tdbox><%If GetData(6,N) = 0 Then
			Response.Write "δ����"
		Else
			Response.Write "<span class=greenfont>�Ѹ���</span>"
		End If%></td>
	</tr>
		<%
	Next
	End If

	Response.Write "<tr><td class=tdbox colspan=6><div class=j_page>"
	If payflag = -1 Then
		Response.Write "<b>ȫ������</b>"
	Else
		Response.Write "<a href=payment.asp?act=list>ȫ������</a>"
	End If
	If payflag = 1 Then
		Response.Write "<b>�Ѹ���</b>"
	Else
		Response.Write "<a href=payment.asp?act=list&payflag=1>�Ѹ���</a>"
	End If
	If payflag = 0 Then
		Response.Write "<b>δ����</b>"
	Else
		Response.Write "<a href=payment.asp?act=list&payflag=0>δ����</a>"
	End If
	If CheckSupervisorUserName = 1 Then Response.Write " <b><span class=gray>���ǹ���Ա,������ʾΪȫ���û�������Ϣ</span></b>"
	%>
	</div>
	</td>
	</tr>
	</table>
	<%

End Sub
%>