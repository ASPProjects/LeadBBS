<%@ LANGUAGE=VBScript CodePage=936%>
<%Option Explicit
Response.Charset = "gb2312"
Session.CodePage=936

Class Proxy_Class

Public Sub GetBody(url)

	url = Left(url,5000)
	If url = "" Then Exit Sub
	Dim xmlHttp
	Set xmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
	xmlHttp.setTimeouts 5000,5000,5000,15000
	xmlHttp.setOption 2, 13056
	xmlHttp.open "GET", url, False, "", "" 
	
	on error resume next
	xmlHttp.send()
	If Err Then
		Exit Sub
	End If

	If xmlHttp.readystate = 4 then 
	'if xmlHttp.status=200 Then
		'Response.Write xmlHttp.ResponseText
		'Response.binaryWrite xmlhttp.Responsebody
		Response.Write BytesToBstr(xmlhttp.Responsebody)
	'end if 
	Else 
		Response.Write ""
	End If
	Set xmlHttp = Nothing

End Sub

private Function BytesToBstr(body) 

	'on error resume next
	dim objstream
	set objstream = Server.CreateObject("adodb.stream")
	with objstream
	.Type = 1
	.Mode = 3
	.Open
	.Write body 
	.Position = 0
	.Type = 2
	.Charset = "GB2312"
	
	'.Charset = "UTF-8"
	BytesToBstr = .ReadText
	.Close
	end with
	set objstream = nothing

End Function

End Class

Sub Proxy_Main

	Dim MyProxy
	Set MyProxy = New Proxy_Class
	MyProxy.GetBody(Request.QueryString("u"))
	Set MyProxy = Nothing

End Sub

Proxy_Main
%>