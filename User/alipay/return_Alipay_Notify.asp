<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Board_Popfun.asp -->
<%
DEF_BBS_HomeUrl = "../../"
Dim Key,partner,out_trade_no,total_fee,receive_name
Dim receive_address,receive_zip,receive_phone,receive_mobile
Dim alipayNotifyURL,Retrieval,ResponseTxt,varItem,mystr
Dim Count,i,j,md5str,mysign
Dim minmax,minmaxSlot,temp,mark,Value

 key="o48habnndc8yr4jtyf9g1p02hlt7fs7h"    ' ֧������ȫ������
 partner="2088002030498170"  '֧��������id


	out_trade_no		= DelStr(Request("out_trade_no")) '��ȡ������
    total_fee		    = DelStr(Request("total_fee")) '��ȡ֧�����ܼ۸�
    receive_name    =DelStr(Request("receive_name"))   '��ȡ�ջ�������
	receive_address =DelStr(Request("receive_address")) '��ȡ�ջ��˵�ַ
	receive_zip     =DelStr(Request("receive_zip"))   '��ȡ�ջ����ʱ�
	receive_phone   =DelStr(Request("receive_phone")) '��ȡ�ջ��˵绰
	receive_mobile  =DelStr(Request("receive_mobile")) '��ȡ�ջ����ֻ�

'******************************************�ж���Ϣ�ǲ���֧��������
alipayNotifyURL = "http://notify.alipay.com/trade/notify_query.do?"
alipayNotifyURL = alipayNotifyURL &"partner=" & partner & "&notify_id=" & request("notify_id")
	Set Retrieval = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
    Retrieval.setOption 2, 13056 
    Retrieval.open "GET", alipayNotifyURL, False, "", "" 
    Retrieval.send()
    ResponseTxt = Retrieval.ResponseText
	Set Retrieval = Nothing
'*****************************************
'��ȡ֧����GET����֪ͨ��Ϣ,�ж���Ϣ�ǲ��Ǳ��޸Ĺ�
For Each varItem in Request.QueryString
mystr=varItem&"="&Request(varItem)&"^"&mystr
Next 
If mystr<>"" Then 
mystr=Left(mystr,Len(mystr)-1)
End If 

mystr = SPLIT(mystr, "^")
Count=ubound(mystr)
'�Բ�������
For i = Count TO 0 Step -1
minmax = mystr( 0 )
minmaxSlot = 0
For j = 1 To i
mark = (mystr( j ) > minmax)
If mark Then 
minmax = mystr( j )
minmaxSlot = j
End If 
Next
    
If minmaxSlot <> i Then 
temp = mystr( minmaxSlot )
mystr( minmaxSlot ) = mystr( i )
mystr( i ) = temp
End If
Next
'����md5ժҪ�ַ���
 For j = 0 To Count Step 1
 value = SPLIT(mystr( j ), "=")
 If  value(1)<>"" And value(0)<>"sign" And value(0)<>"sign_type"  Then
 If j=Count Then
 md5str= md5str&mystr( j )
 Else 
 md5str= md5str&mystr( j )&"&"
 End If 
 End If 
 Next
md5str=md5str&key
 mysign=md5(md5str)
'********************************************************

If mysign=Request("sign") and ResponseTxt="true"   Then 	

	'response.write "����ɹ�ҳ��"        '�������ָ������Ҫ��ʾ������
	Alipay_UpdateSellList
	Response.Redirect "payment.asp?act=done&str=" & server.urlencode("<font color=green class=greenfont>����ɹ�,�ɲ鿴������ʷ�����ٴ�ȷ�ϳ�ֵ���!</font>")

Else
	'response.write "��תʧ��"          '�������ָ������Ҫ��ʾ������
	Response.Redirect "payment.asp?act=done&str=" & server.urlencode("<font color=red class=redfont>����ʧ��,�뷵����ʼҳ�������ύ���ѯ��ʷ�����˶����!</font>")
End If

Function DelStr(Str)

	If IsNull(Str) Or IsEmpty(Str) Then
		Str	= ""
	End If
	DelStr	= Replace(Str,";","")
	DelStr	= Replace(DelStr,"'","")
	DelStr	= Replace(DelStr,"&","")
	DelStr	= Replace(DelStr," ","")
	DelStr	= Replace(DelStr,"��","")
	DelStr	= Replace(DelStr,"%20","")
	DelStr	= Replace(DelStr,"--","")
	DelStr	= Replace(DelStr,"==","")
	DelStr	= Replace(DelStr,"<","")
	DelStr	= Replace(DelStr,">","")
	DelStr	= Replace(DelStr,"%","")

End Function

Sub Alipay_UpdateSellList

	InitDatabase
	If isNumeric(out_trade_no) = 0 Then out_trade_no = 0
	out_trade_no = Fix(cCur(out_trade_no))
	If isNumeric(total_fee) = 0 Then total_fee = 0
	total_fee = Fix(cCur(total_fee))
	
	Dim Rs,UserName,PayPoints,GetPoints,PayFlag
	Set Rs = LdExeCute(sql_select("Select UserName,PayPoints,GetPoints,PayFlag From LeadBBS_SellList Where PID=" & out_trade_no,1),0)
	If Not Rs.Eof Then
		UserName = Rs(0)
		PayPoints = cCur(Rs(1))
		GetPoints = cCur(Rs(2))
		PayFlag = cCur(Rs(3))
		Rs.Close
		Set Rs = Nothing
	Else
		Rs.Close
		Set Rs = Nothing
	End If
	
	If PayFlag = 0 and PayPoints = total_fee Then '����Ӧ,����δ֧��������¸��¶������û�״̬
		CALL LdExeCute("Update LeadBBS_SellList Set PayFlag=1 Where PID=" & out_trade_no,1)
		CALL LdExeCute("Update LeadBBS_User Set CharmPoint=CharmPoint+" & GetPoints & " Where UserName='" & Replace(UserName,"'","''") & "'",1)
	End If
	
	CloseDatabase

End Sub
%>