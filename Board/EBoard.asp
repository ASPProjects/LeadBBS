<%
	If Request.QueryString = "" Then
		Response.Redirect "../b/eb.asp"
	Else
		Response.Redirect "../b/eb.asp?" & Request.QueryString
	End If
%>