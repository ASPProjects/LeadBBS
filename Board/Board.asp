<%
	If Request.QueryString = "" Then
		Response.Redirect "../b/b.asp"
	Else
		Response.Redirect "../b/b.asp?" & Request.QueryString
	End If
%>