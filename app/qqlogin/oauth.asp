<%
'==================================
'=类 名 称：QqConnet
'=功    能：QQ登录 For ASP
'=作    者：㊣FireFox㊣
'=Q      Q: 63572063
'=日    期：2012-01-02
'==================================
'转载时请保留以上内容！！
Const apiKey = "" 'APP ID,您需要从腾讯平台申请获取资料：(<a href=http://connect.qq.com/ target=_blank>点此申请</a>)
Const secretKey = "" 'APP KEY,您需要从腾讯平台申请获取
Const callback = "" 'CALL BACK,回调地址，注意只需要填写域名，不包括http及目录。

Class QqConnet
    Private QQ_OAUTH_CONSUMER_KEY
    Private QQ_OAUTH_CONSUMER_SECRET
	Private QQ_CALLBACK_URL
	Private QQ_SCOPE
        
    Private Sub Class_Initialize
        QQ_OAUTH_CONSUMER_KEY = apiKey
        QQ_OAUTH_CONSUMER_SECRET = secretKey
        QQ_CALLBACK_URL = callback & Request.Servervariables("SCRIPT_NAME")
		QQ_SCOPE ="get_user_info,add_t,add_share,get_info,add_topic" '授权项 例如：QQ_SCOPE=get_user_info,list_album,upload_pic,do_like,add_t 
                                                '不传则默认请求对接口get_user_info进行授权。
                                                '建议控制授权项的数量，只传入必要的接口名称，因为授权项越多，用户越可能拒绝进行任何授权。
    End Sub
    Property Get APP_ID()    
        APP_ID = QQ_OAUTH_CONSUMER_KEY    
    End Property

	'生成Session("State")数据.
	Public Function MakeRandNum()
		Randomize
		Dim width : width = 6 '随机数长度,默认6位
		width = 10 ^ (width - 1)
		MakeRandNum = Int((width*10 - width) * Rnd() + width)
	End Function
	
	Private Function CheckXml()
        Dim oxml,Getxmlhttp
        On Error Resume Next
        oxml=array("Microsoft.XMLHTTP","Msxml2.ServerXMLHTTP.6.0","Msxml2.ServerXMLHTTP.5.0","Msxml2.ServerXMLHTTP.4.0","Msxml2.ServerXMLHTTP.3.0","Msxml2.ServerXMLHTTP","Msxml2.XMLHTTP.6.0","Msxml2.XMLHTTP.5.0","Msxml2.XMLHTTP.4.0","Msxml2.XMLHTTP.3.0","Msxml2.XMLHTTP")
        For i=0 to ubound(oxml)
           Set Getxmlhttp = Server.CreateObject(oxml(i))
           If Err Then
              Err.Clear
              CheckXml = False
           Else
              CheckXml = oxml(i) :Exit Function
           End if
       Next
     End Function

	
	'Get方法请求url,获取请求内容
	Private Function RequestUrl(url)
		dim XmlObj
		Set XmlObj = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
		XmlObj.open "GET",url, false
		XmlObj.send
		RequestUrl = XmlObj.responseText
		Set XmlObj = nothing
	End Function
	
	'Post方法请求url,获取请求内容
	Private Function RequestUrl_post(url,data)
		dim XmlObj
		'Set XmlObj = Server.CreateObject(CheckXml())
		Set XmlObj = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
		XmlObj.open "POST", url, false
		XmlObj.setrequestheader "POST"," /t/add_t HTTP/1.1"
		XmlObj.setrequestheader "Host"," graph.qq.com"
		XmlObj.setrequestheader "content-length",len(data)  
      XmlObj.setRequestHeader "Content-Type"," application/x-www-form-urlencoded "
		XmlObj.setrequestheader "Connection"," Keep-Alive"
        XmlObj.setrequestheader "Cache-Control"," no-cache"
        XmlObj.send(data)
		RequestUrl_post = XmlObj.responseText
		Set XmlObj = nothing
	End Function
	
	
	Private Function CheckData(data,str)
		If Instr(data,str)>0 Then
		   CheckData = True
		Else
		   CheckData = False
		End If
	End Function
	

	
	'生成登录地址
	Public Function GetAuthorization_Code()
		Dim url, params
		url = "https://graph.qq.com/oauth2.0/authorize"
		params = "client_id=" & QQ_OAUTH_CONSUMER_KEY
		params = params & "&redirect_uri=" & QQ_CALLBACK_URL
		params = params & "&response_type=code"
		params = params & "&scope="&QQ_SCOPE
		params = params & "&state="&Session("State")
		url = url & "?" & params
		GetAuthorization_Code = (url)
	End Function
	
	
	'获取 access_token
	'获取到的access token具有3个月有效期，用户再次登录时自动刷新。
	'第三方网站可存储access token信息，以便后续调用OpenAPI访问和修改用户信息时使用。
	Public Function GetAccess_Token()
		Dim url, params,Temp
		Url="https://graph.qq.com/oauth2.0/token"
	    params = "client_id=" & QQ_OAUTH_CONSUMER_KEY
		params = params & "&client_secret=" & QQ_OAUTH_CONSUMER_SECRET
		params = params & "&redirect_uri=" & QQ_CALLBACK_URL
		params = params & "&grant_type=authorization_code"
		params = params & "&code="&Session("Code")
		params = params & "&state="&Session("State")
		url = Url & "?" & params
		Temp=RequestUrl(url)
		If CheckData(Temp,"access_token=") = True Then
           GetAccess_Token=CutStr(Temp,"access_token=","&")
		Else
		   ErrorJump("获取 access_token 时发生错误，错误代码："&CutStr(Temp,"{""error"":",",")) 
		End If
	End Function
	
	'检测是否合法登录！
	Public Function CheckLogin()
		Dim Code,mState
		Code=Trim(Request.QueryString("code"))
		'mState=Trim(Request.QueryString("state"))
		If Code<>"" Then
			CheckLogin = True
			Session("Code")=Code
		Else
			CheckLogin = False
		End If
	End Function
	
	'获取openid
	Public Function Getopenid()
		Dim url, params,Temp
		url = "https://graph.qq.com/oauth2.0/me"
		params = "access_token="&Session("Access_Token")
		url = Url & "?" & params
		Temp=RequestUrl(url)
		If Instr(Temp,"openid")>0 Then
		   Getopenid=CutStr(Temp,"openid"":""","""}")
		Else
		   ErrorJump("获取 Openid 时发生错误，错误代码："&CutStr(Temp,"{""error"":",",")) 
		End If
	End Function
	
	'发送一条微博
	Public Function Post_Webo(content,Access_Token,Access_Openid)
		Dim url, params, Tk,Oid
		Tk = Access_Token
		Oid = Access_Openid
		if Tk = "" then Tk = Session("Access_Token")
		if Oid = "" then Oid = Session("Openid")
		url = "https://graph.qq.com/t/add_t"
		params = "oauth_consumer_key=" & QQ_OAUTH_CONSUMER_KEY
		params = params & "&access_token=" & Tk
		params = params & "&openid=" & Oid
		params = params & "&content="&sim_urlencode(content)
        params = params & "&format=json"
		Post_Webo = RequestUrl_post(url,params)
	End Function
	'发送一条说说
	Public Function Post_add_topic(content)
		Dim url, params
		url = "https://graph.qq.com/shuoshuo/add_topic"
		params = "oauth_consumer_key=" & QQ_OAUTH_CONSUMER_KEY
		params = params & "&access_token=" & Session("Access_Token")
		params = params & "&openid=" & Session("Openid")
		params = params & "&con="&content
        params = params & "&format=json"
		Post_add_topic = RequestUrl_post(url,params)
	End Function
	
	private function sim_urlencode(str)
	
		sim_urlencode = replace(replace(replace(str,"&","%26"),"=","%3D"),VbCrLf," ")
	
	end function
	
	'分享内容到QQ空间
	Public Function Post_Share(title,turl,comment,summary,images,nswb,Access_Token,Access_Openid)
		Dim url, params, Tk,Oid
		Tk = Access_Token
		Oid = Access_Openid
		if Tk = "" then Tk = Session("Access_Token")
		if Oid = "" then Oid = Session("Openid")
		url = "https://graph.qq.com/share/add_share"
		params = "oauth_consumer_key=" & QQ_OAUTH_CONSUMER_KEY
		params = params & "&access_token=" & Tk
		params = params & "&openid=" & Oid
		params = params & "&title="&sim_urlencode(title)
		params = params & "&url="&sim_urlencode(turl)
		params = params & "&comment="&sim_urlencode(comment)
		params = params & "&summary="&sim_urlencode(summary)
		params = params & "&images="&sim_urlencode(images)
		params = params & "&nswb="&sim_urlencode(nswb)
		params = params & "&format=json"
		Post_Share = RequestUrl_post(url,params)
	End Function
	
	'获取用户信息,得到一个json格式的字符串
	Public Function GetUserInfo()
		Dim url, params, result
		url = "https://graph.qq.com/user/get_user_info"
		params = "oauth_consumer_key=" & QQ_OAUTH_CONSUMER_KEY
		params = params & "&access_token=" & Session("Access_Token")
		params = params & "&openid=" & Session("Openid")
		url = url & "?" & params
		GetUserInfo = RequestUrl(url)
		If CheckData(GetUserInfo,"nickname") = False Then
		   ErrorJump("获取用户信息时发生错误，错误代码："&CutStr(GetUserInfo,"{""ret"":",",")) 
		End If
	End Function
	
	'获取腾讯微博登录用户的用户资料,得到一个json格式的字符串
	Public Function Get_Info()
		Dim url, params, result
		url = "https://graph.qq.com/user/get_info"
		params = "oauth_consumer_key=" & QQ_OAUTH_CONSUMER_KEY
		params = params & "&access_token=" & Session("Access_Token")
		params = params & "&openid=" & Session("Openid")
		params = params & "&format=json"
		url = url & "?" & params
		Get_Info = RequestUrl(url)
	End Function

	
	'获取用户名字,性别,从json字符串里截取相关字符
	Public Function GetUserName(json)
	    Dim nickname,sex
		nickname = CutStr(json, "nickname"":""",""",")
		sex=CutStr(json, "gender"":""","""")
	    GetUserName = Array(nickname,sex)
	End Function
	
	Public Function CutStr(data,s_str,e_str)
	    If Instr(data,s_str)>0 and Instr(data,e_str)>0 Then
		   CutStr = Split(data,s_str)(1)
		   CutStr = Split(CutStr,e_str)(0)
		Else
		   CutStr = ""
		End If
	End Function
	
End Class
%>