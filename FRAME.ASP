<!-- #include file=inc/BBSsetup.asp -->
<%
Dim GBL_FrameUrl
GBL_FrameUrl = filterUrlstr(Request.QueryString("u"))
If GBL_FrameUrl = "" Then
	GBL_FrameUrl = "http://" & LCase(Request.ServerVariables("server_name")) & Replace(LCase(Request.Servervariables("SCRIPT_NAME")),"frame.asp","")
	If Left(LCase(Request.ServerVariables("HTTP_REFERER")),Len(GBL_FrameUrl)) = GBL_FrameUrl Then
		GBL_FrameUrl = filterUrlstr(Request.ServerVariables("HTTP_REFERER"))
	Else
		GBL_FrameUrl = "Boards.asp"
	End If
End If
//response.Write GBL_FrameUrl
//Response.End
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
	"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"  xml:lang="zh-CN" lang="zh-CN">
<title><%=DEF_SiteNameString%></title>
</head>
<script type="text/javascript">   
if(top.location != self.location)   
top.location=self.location;   
</script>
<frameset id="forum" name="forum" cols="150,*" frameborder="NO" border="0" framespacing="0">
	<frameset name="left" rows="0,*" cols="*" framespacing="0" frameborder="NO" border="0">
		<frame src="" name="topFrame" scrolling="NO" noresize marginwidth="0" marginheight="0">
		<frame src="BoardNav.asp" name="leftFrame" scrolling="NO" noresize marginwidth="0" marginheight="1">
	</frameset>
	<frameset cols="10,*" border="0" framespacing="0">
		<frame src="frame_button.asp" noresize name="button" scrolling="NO">
		<frame src="<%=GBL_FrameUrl%>" scrolling="auto" name="r_top" scrolling="yes">
	</frameset>
</frameset>
<noframes><body></body></noframes>
</html>