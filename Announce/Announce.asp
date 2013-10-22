<%
	If Request.QueryString = "" Then
		Response.Redirect "../a/a.asp"
	Else
		Response.Redirect "../a/a.asp?" & Request.QueryString
	End If
%>