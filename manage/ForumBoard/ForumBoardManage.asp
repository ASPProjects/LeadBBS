<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Board_Popfun.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<!-- #include file=inc/ForumBoard_fun.asp -->
<%
DEF_BBS_HomeUrl = "../../"
Dim GBL_ID
initDatabase
GBL_CHK_TempStr = ""
GBL_ID = checkSupervisorPass

Manage_sitehead DEF_SiteNameString & " - ����Ա",""
frame_TopInfo
DisplayUserNavigate("��̳�������")
If GBL_CHK_Flag=1 Then
	LoginAccuessFul
Else
DisplayLoginForm
End If
frame_BottomInfo
closeDataBase
Manage_Sitebottom("none")

Function LoginAccuessFul

	GBL_CHK_TempStr = ""
	Dim Rs
	Set Rs = LDExeCute("Select T1.BoardID,T1.BoardAssort,T1.BoardName,T1.OrderID,T2.AssortName,T1.ParentBoard from LeadBBS_Boards as T1 left join LeadBBS_Assort as T2 on T1.BoardAssort=T2.AssortID order by T1.BoardAssort,T1.ParentBoard,T1.OrderID",0)
	If Rs.Eof Then
		Response.Write "��û���κ���̳��������Ӱ�!"
		Rs.Close
		Set Rs = Nothing
		Exit Function
	Else
		GBL_GetData = Rs.GetRows(-1)
		Rs.Close
		Set Rs = Nothing
	End If
	If GBL_CHK_TempStr<>"" then
		Response.Write GBL_CHK_TempStr
	Else
%>
<script language=javascript>
function opw(f,r,id)
{
	document.location.href = f+'?B=<%=GBL_board_ID%>&'+r+'='+id;
}
</script>
<a href=ForumBoardJoin.asp>�������������̳</a> <br>
<span class=redfont>[<u>ɾ������ȫ������</u>��<u>�ϲ�����</u>����<b>������̳�б��޸�</b>�Ա�֤��̳��������ͳ����ȷ]</span>
<table border=0 cellpadding=0 cellspacing=0 width="100%" class=frame_table>
	<tr class=frame_tbhead>
		<td width="8%"><div class=value>ID</div></td>
		<td width="62%"><div class=value>��̳����</div></td>
		<td width="20%"><div class=value>����</div></td>
		<td width="10%"><div class=value>˳��</div></td>
	</tr>
	<%
	Dim CountN,TempN
	CountN = Ubound(GBL_GetData,2)
	Dim Old_Assort
	Old_Assort = 0
	for TempN=0 to CountN
		If cCur(GBL_GetData(1,TempN)) <> Old_Assort Then
			Response.Write "       <tr>" & VbCrLf
			Response.Write "          <td class=tdbox colspan=4><a href=../ForumCategory/ForumCategoryManage.asp?action=edit&GBL_MODIFYID=" & GBL_GetData(1,TempN) & ">" & GBL_GetData(4,TempN) & "</a></td>" & VbCrLf
			Response.Write "       </td></tr>" & VbCrLf
			Old_Assort = cCur(GBL_GetData(1,TempN))
		End If
		Response.Write "        <tr bgcolor=#FFFFFF class=TBBG9>" & VbCrLf
		Response.Write "          <td class=tdbox>"
		Response.Write GBL_GetData(0,TempN) & "</td>" & VbCrLf
		Response.Write "          <td class=tdbox>"
		If cCur(GBL_GetData(5,TempN)) > 0 and cCur(GBL_GetData(5,TempN)) <> cCur(GBL_GetData(0,TempN)) Then Response.Write "<font color=Red class=redfont title=""�Ӱ���"">��</font>"
		Response.Write GBL_GetData(2,TempN) & " <a href=ForumBoardModify.asp?GBL_MODIFYID=" & GBL_GetData(0,TempN) & ">�޸�</a> <a href=ForumBoardDelete.asp?GBL_DELETEID=" & GBL_GetData(0,TempN) & " title=����ɾ��һ�����κη����İ���>ɾ����</a>" & VbCrLf
		Response.Write "          <a href='javascript:opw(""ForumBoardDeleteAnnounce.asp"",""DelBoardID""," & GBL_GetData(0,TempN) & ");' title=��մ˰�����������ӣ������û�" & DEF_PointsName(0) & ">ɾ����</a>"
		Response.Write "          <a href='javascript:opw(""BoardMoveAnnounce.asp"",""MoveFromBoardID""," & GBL_GetData(0,TempN) & ");' title=�ϲ�(ת��)�˰�ȫ�����ӵ���������>�ϲ�</a>"
		%>
		<a href=../BlockUpdate/UpdateRootMaxMinAnnounceID.asp?ID=<%=GBL_GetData(0,TempN)%>&BlockType=2 title="�������⽫�������ظ�ʱ�����������������ڰ���ϲ���������޸�">��������</a>
		<%
		Response.write "			</td>"
		Response.Write "          <td class=tdbox><a href=../ForumCategory/ForumCategoryManage.asp?action=edit&GBL_MODIFYID=" & GBL_GetData(1,TempN) & ">" & GBL_GetData(4,TempN) & "</a></td>" & VbCrLf
		Response.Write "<td class=tdbox>" & GBL_GetData(3,TempN) & "</td>" & VbCrLf
		Response.Write "        </tr>" & VbCrLf
	next
	%>
	</table>
<%
	End If

End Function%>