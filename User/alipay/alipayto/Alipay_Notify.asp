<%
	'���ܣ���������з�����֪ͨҳ��
	'�汾��2.0
	'���ڣ�2008-10-24
	'���ߣ�֧������˾���۲�����֧���Ŷ�
	'��ϵ��0571-26888888
	'��Ȩ��֧������˾
%>

<!--#include file="alipayto/Alipay_md5.asp"-->
<%
    key="o48habnndc8yr4jtyf9g1p02hlt7fs7h"         '֧������ȫ������
    partner="2088002030498170"     '֧��������id 
 
	out_trade_no	=DelStr(Request.Form("out_trade_no"))      '��ȡ������
    total_fee		=DelStr(Request.Form("total_fee"))         '��ȡ֧�����ܼ۸�
	'�����ȡ��������������д ���� =DelStr(Request.Form("��ȡ������"))
	
'*******************�ж���Ϣ�ǲ���֧��������***********************
alipayNotifyURL = "http://notify.alipay.com/trade/notify_query.do?"
alipayNotifyURL = alipayNotifyURL &"partner=" & partner & "&notify_id=" & request.Form("notify_id")
	Set Retrieval = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
    Retrieval.setOption 2, 13056 
    Retrieval.open "GET", alipayNotifyURL, False, "", "" 
    Retrieval.send()
    ResponseTxt = Retrieval.ResponseText
	Set Retrieval = Nothing
'*******************************************************************

'*******************��ȡ֧����POST����֪ͨ��Ϣ**********************
For Each varItem in Request.Form
	mystr=varItem&"="&Request.Form(varItem)&"^"&mystr
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

'*************************����״̬���ش���*************************
If mysign=request.Form("sign") And ResponseTxt="true" Then 	
	If request.Form("trade_status") = "TRADE_FINISHED" Then 
		'�ڴ˴����ӣ�����ɹ�,�������ݿ����  
	 	returnTxt	= "success"
	Else
		returnTxt	= "fail"
	End If
	Response.Write returnTxt
Else
response.write "fail"
End If 
'*******************************************************************
 'д�ı���������ԣ�����վ����Ҳ���Ըĳɴ������ݿ⣩
TOEXCELLR=TOEXCELLR&md5str&"MD5���:"&mysign&"="&request.Form("sign")&"--ResponseTxt:"&ResponseTxt
set fs= createobject("scripting.filesystemobject") 
set ts=fs.createtextfile(server.MapPath("alipayto/Notify_DATA/"&replace(now(),":","")&".txt"),true)

ts.writeline(TOEXCELLR)
ts.close
set ts=Nothing
set fs=Nothing

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
%>