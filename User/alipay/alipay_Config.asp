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
show_url = Lcase("http://"&Request.ServerVariables("server_name")) & "/"                   '��վ����ַ
seller_email = DEF_seller_email				'�����ó����Լ���֧�����ʻ�
partner = "2088002030498170" '֧������ȡid��������Ҫһ��֧�����˺ţ��ٴ���Ӧ��ַ��ȡid(<a href=https://www.alipay.com/himalayas/practicality_customer.htm?customer_external_id=C4335329546596834111&market_type=from_agent_contract&pro_codes=F7F62F29651356BB target=_blank>��˻�ȡ</a>)
key = "o48habnndc8yr4jtyf9g1p02hlt7fs7h" '֧������ȡ����Կ��������Ҫһ��֧�����˺ţ��ٴ���Ӧ��ַ��ȡ��Կ(<a href=https://www.alipay.com/himalayas/practicality_customer.htm?customer_external_id=C4335329546596834111&market_type=from_agent_contract&pro_codes=F7F62F29651356BB target=_blank>��˻�ȡ</a>)

notify_url = alipay_MyHomeUrl & "alipay/Alipay_Notify.asp"	'�����������֪ͨ��ҳ�� Ҫ�� http://��ʽ������·��
return_url = alipay_MyHomeUrl & "alipay/return_Alipay_Notify.asp"	'��������ת��ҳ�� Ҫ�� http://��ʽ������·��

	 
'��½ www.alipay.com ��, ���̼ҷ���,���Կ���֧������ȫУ����ͺ���id,������������

%>