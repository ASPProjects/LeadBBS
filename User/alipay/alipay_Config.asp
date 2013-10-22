<!-- #include file=../../inc/User_Setup.ASP -->
<%
Dim alipay_MyHomeUrl
alipay_MyHomeUrl = LCase(Request.Servervariables("SCRIPT_NAME"))
If Right(alipay_MyHomeUrl,24) = "/user/alipay/payment.asp" Then
	If Request.ServerVariables("SERVER_PORT") <> "80" Then alipay_MyHomeUrl = ":" & Request.ServerVariables("SERVER_PORT") & alipay_MyHomeUrl
	alipay_MyHomeUrl = Lcase("http://"&Request.ServerVariables("server_name") & alipay_MyHomeUrl)
	alipay_MyHomeUrl = Replace(alipay_MyHomeUrl,"/user/alipay/payment.asp","/user/")
Else
	alipay_MyHomeUrl = ""
End If

Dim show_url,seller_email,partner,key,notify_url,return_url
show_url = Lcase("http://"&Request.ServerVariables("server_name")) & "/"                   '网站的网址
seller_email = DEF_seller_email				'请设置成您自己的支付宝帐户
partner = "2088002030498170" '支付宝获取id，您先需要一个支付宝账号，再从相应网址获取id(<a href=https://www.alipay.com/himalayas/practicality_customer.htm?customer_external_id=C4335329546596834111&market_type=from_agent_contract&pro_codes=F7F62F29651356BB target=_blank>点此获取</a>)
key = "o48habnndc8yr4jtyf9g1p02hlt7fs7h" '支付宝获取的密钥，您先需要一个支付宝账号，再从相应网址获取密钥(<a href=https://www.alipay.com/himalayas/practicality_customer.htm?customer_external_id=C4335329546596834111&market_type=from_agent_contract&pro_codes=F7F62F29651356BB target=_blank>点此获取</a>)

notify_url = alipay_MyHomeUrl & "alipay/Alipay_Notify.asp"	'付完款后服务器通知的页面 要用 http://格式的完整路径
return_url = alipay_MyHomeUrl & "alipay/return_Alipay_Notify.asp"	'付完款后跳转的页面 要用 http://格式的完整路径

	 
'登陆 www.alipay.com 后, 点商家服务,可以看到支付宝安全校验码和合作id,导航栏的下面

%>