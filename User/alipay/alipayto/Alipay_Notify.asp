<%
	'功能：付款过程中服务器通知页面
	'版本：2.0
	'日期：2008-10-24
	'作者：支付宝公司销售部技术支持团队
	'联系：0571-26888888
	'版权：支付宝公司
%>

<!--#include file="alipayto/Alipay_md5.asp"-->
<%
    key="o48habnndc8yr4jtyf9g1p02hlt7fs7h"         '支付宝安全教研码
    partner="2088002030498170"     '支付宝合作id 
 
	out_trade_no	=DelStr(Request.Form("out_trade_no"))      '获取定单号
    total_fee		=DelStr(Request.Form("total_fee"))         '获取支付的总价格
	'如需获取其它参数，可填写 参数 =DelStr(Request.Form("获取参数名"))
	
'*******************判断消息是不是支付宝发出***********************
alipayNotifyURL = "http://notify.alipay.com/trade/notify_query.do?"
alipayNotifyURL = alipayNotifyURL &"partner=" & partner & "&notify_id=" & request.Form("notify_id")
	Set Retrieval = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
    Retrieval.setOption 2, 13056 
    Retrieval.open "GET", alipayNotifyURL, False, "", "" 
    Retrieval.send()
    ResponseTxt = Retrieval.ResponseText
	Set Retrieval = Nothing
'*******************************************************************

'*******************获取支付宝POST过来通知消息**********************
For Each varItem in Request.Form
	mystr=varItem&"="&Request.Form(varItem)&"^"&mystr
Next 
If mystr<>"" Then 
	mystr=Left(mystr,Len(mystr)-1)
End If 
mystr = SPLIT(mystr, "^")
Count=ubound(mystr)
'对参数排序
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
'构造md5摘要字符串
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

'*************************交易状态返回处理*************************
If mysign=request.Form("sign") And ResponseTxt="true" Then 	
	If request.Form("trade_status") = "TRADE_FINISHED" Then 
		'在此处添加：付款成功,更新数据库语句  
	 	returnTxt	= "success"
	Else
		returnTxt	= "fail"
	End If
	Response.Write returnTxt
Else
response.write "fail"
End If 
'*******************************************************************
 '写文本，方便测试（看网站需求，也可以改成存入数据库）
TOEXCELLR=TOEXCELLR&md5str&"MD5结果:"&mysign&"="&request.Form("sign")&"--ResponseTxt:"&ResponseTxt
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
	DelStr	= Replace(DelStr,"　","")
	DelStr	= Replace(DelStr,"%20","")
	DelStr	= Replace(DelStr,"--","")
	DelStr	= Replace(DelStr,"==","")
	DelStr	= Replace(DelStr,"<","")
	DelStr	= Replace(DelStr,">","")
	DelStr	= Replace(DelStr,"%","")
End Function
%>