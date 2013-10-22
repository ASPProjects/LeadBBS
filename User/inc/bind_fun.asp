<!-- #include file=../../app/qqlogin/oauth.asp -->
<%
Sub DisplayBind

	Dim ConnetBind
	set ConnetBind = New Connet_Bind
	ConnetBind.BindList
	set ConnetBind = nothing

End Sub

Sub Unbind

	dim apptype
	apptype = toNum(request.querystring("apptype"),0)
	if apptype > 0 then
		call ldexecute("delete from leadbbs_applogin where userid=" & GBL_UserID & " and apptype=" & apptype,1)
		Response.Write "成功解除绑定."
	end if

End Sub

Class Connet_Bind

	private BindType 
	private BindName
	private BindLoginUrl
	private BindMethod

	Private Sub Class_Initialize
	
		BindType = array(1)
		BindName = array("腾讯QQ空间")
		BindLoginUrl = array("app/qqlogin/login.asp")
		BindMethod = array("发表帖子","回复帖子","评论","收藏")
	
	End Sub
	
	Public Sub BindList
	
		dim rs,sql,getdata,num,Retention1
		sql = "select id,appid,Token,appType,Retention1 from leadbbs_applogin where userid=" & GBL_UserID & " order by apptype asc"
		set rs = ldexecute(sql,0)
		If rs.eof then
			num = -1
		else
			getdata = rs.getrows(-1)
			num = Ubound(GetData,2)
		end if
		rs.close
		set rs = nothing
		dim n,i,exist,m
		%>
		<script>
		function unbind(n,url)
		{
			getAJAX('LookUserInfo.asp?Evol=unbind&apptype='+n+'&t='+Math.random(),'','$id("unbind'+n+'").innerHTML="<a href=<%=DEF_BBS_HomeUrl%>'+url+' class=\'fmbtn btn_3 inline\' target=_blank>立即绑定</a>";',1);return false;
		}
		</script>
		<%
		for n = 0 to ubound(BindType)
			exist = 0
							%>
							<div class="value">
							<img src="<%=DEF_BBS_HomeUrl%>images/app/<%=BindType(n)%>.gif" class="absmiddle" />
							<%=BindName(n)%>
							<%
			for i = 0 to num
				if BindType(n) = GetData(3,i) then
					exist = 1
					
					Retention1 = ccur(toNum(GetData(4,i),0))
					select case BindType(n)
						Case 1:
							%>
							<span id="unbind<%=BindType(n)%>">
							<%
							for m = 0 to ubound(BindMethod)
								%>
								<input type="checkbox" class=fmchkbox name="Limit_<%=BindType(n)%>_<%=m%>" value="1"<%
								If GetBinarybit(Retention1,m+1) = 0 Then
									Response.Write " checked>"
								Else
									Response.Write ">"
								End If%><%=BindMethod(m)%>
									<%
							next
							%>
							<a href="javascript:;" onclick="unbind(<%=BindType(n)%>,'<%=BindLoginUrl(N)%>');" class="fmbtn btn_3 inline">解除绑定</a>
							</span>
							<%
					end select
				end if
			next
 						If exist = 0 then%>
							<a href="<%=DEF_BBS_HomeUrl%><%=BindLoginUrl(N)%>" class="fmbtn btn_3 inline" target=_blank>立即绑定</a>
							</div>
							<%
						end if
		next
	
	End Sub

	Public Sub BindAnnounceList
	
		dim rs,sql,getdata,num,Retention1
		sql = "select id,appid,Token,appType,Retention1 from leadbbs_applogin where userid=" & GBL_UserID & " order by apptype asc"
		set rs = ldexecute(sql,0)
		If rs.eof then
			num = -1
		else
			getdata = rs.getrows(-1)
			num = Ubound(GetData,2)
		end if
		rs.close
		set rs = nothing
		if num = -1 then exit sub
		dim n,i,exist,m
		%>
		<script>
		function toggleBind(n)
		{
			if($(n).prev().attr("value")=="1")
			{
				$(n).prev().val("0")
				$(n).next().attr("class","");
			}
			else
			{			
				$(n).prev().val("1")
				$(n).next().attr("class","select");
			}
		}
		</script>
		<div class="value bindpost"><span>同步到</span>
		<%
		for n = 0 to ubound(BindType)
			exist = 0
			for i = 0 to num
				if BindType(n) = GetData(3,i) and GetData(2,i) & "" <> "" then
					exist = 1					
					Retention1 = ccur(toNum(GetData(4,i),0))
					select case BindType(n)
						Case 1:
							%>
							<span class="item">
							<input type="hidden" name="bindpost_<%=BindType(n)%>" value="1">
							<a href="javascript:;" title="同步到<%=BindName(n)%>" onclick="toggleBind(this);return false;" class="bindpost_<%=BindType(n)%>"></a><em class="select"></em>
							</span>
							
							<span class="item">
							<input type="hidden" name="bindpost_<%=BindType(n)%>_1" value="1">
							<a href="javascript:;" title="同步到腾讯微博" onclick="toggleBind(this);return false;" class="bindpost_<%=BindType(n)%>_weibo"></a><em class="select"></em>
							</span>
							<%
						case else
					end select
				end if
			next
 						If exist = 0 then%>
 							<input type="hidden" name="bindpost_<%=BindType(n)%>" value="0">
							<a href="<%=DEF_BBS_HomeUrl%><%=BindLoginUrl(N)%>" title="连接到<%=BindName(n)%>" target=_blank class="bindpost_<%=BindType(n)%>_dis"></a>
							<%
						end if
		next
		%>
		</div>
		<%
	
	End Sub
	
	Public Sub PostShare(appType,title,turl,comment,summary,images,weibo)
	
		dim rs,sql,getdata,num,Token,appid
		sql = "select id,appid,Token,appType,Retention1 from leadbbs_applogin where userid=" & GBL_UserID & " and apptype=" & appType
		set rs = ldexecute(sql,0)
		If rs.eof then
			num = -1
		else
			num = 1
			Token = Rs(2)
			appid = Rs(1)
		end if
		rs.close
		set rs = nothing
		If num = -1 Then exit Sub
		Select case appType
			Case 1:
				dim qc
				set qc = New QqConnet
				if weibo = 2 then
					call qc.Post_Webo(title & " / " & turl & " / " & comment & " / " & summary,Token,appid)
				else
					call qc.Post_Share(title,turl,comment,summary,images,weibo,Token,appid)
				end if
				set qc = nothing
		End Select
	
	End Sub

End Class
%>