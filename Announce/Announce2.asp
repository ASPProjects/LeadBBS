<%
	If Request.QueryString = "" Then
		Response.Redirect "../a/a2.asp"
	Else
		Response.Redirect "../a/a2.asp?" & Request.QueryString
	End If
%>