<%
Dim CMS_NAVIGATECLASS_UpdateTime
CMS_NAVIGATECLASS_UpdateTime = "2013/5/7 17:46:50"

Sub CMS_NAVIGATECLASS_View

%>
<%
dim classid
classid = tonum(request.querystring("classid"),0)
%>
<a class="cms_top_item" href="<%=DEF_BBS_HomeUrl%>article/article.asp?classid=1" id="cmstopitem1">¹«¸æ</a>
<script>
$("#cmstopitem<%=classid%>").attr("class","cms_top_sel");
</script>
<%

End Sub
%>
