<!-- #include file=../../inc/BBSSetup.asp -->
<!-- #include file=../../inc/Board_Popfun.asp -->
<!-- #include file=../inc/bbsmanage_Fun.asp -->
<%
DEF_BBS_HomeUrl = "../../"
Dim GBL_ID
initDatabase
GBL_CHK_TempStr = ""
GBL_ID = checkSupervisorPass
Server.ScriptTimeOut = 600

Dim GBL_EXEString

Manage_sitehead DEF_SiteNameString & " - ����Ա",""

'GBL_CHK_TempStr = "��̳�Ѿ���ֹ��Σ�չ���."

frame_TopInfo
DisplayUserNavigate("ֱ��ִ��SQL���")
If GBL_CHK_Flag=1 and GBL_CHK_TempStr = "" Then
	LoginAccuessFul
Else
	Response.Write "<div class=alert>" & GBL_CHK_TempStr & "</div>"
End If
frame_BottomInfo
closeDataBase
Manage_Sitebottom("none")

Function LoginAccuessFul

	If Request.Form("submitflag")="Dieos9xsl29LO_8" Then
		GBL_EXEString = Request("GBL_EXEString")
		If GBL_EXEString <> "" Then
			On Error Resume Next
			Dim RowCount,Rs
			Dim Time1,Time2
			Time1=Timer
			GBL_EXEString = Request("GBL_EXEString")
			If inStr(Lcase(GBL_EXEString),"leadbbs_log") Then
				Response.Write "<p><br>���󣬲��ܶ���̳��־�����κβ�����"
				Exit Function
			End If
			Con.CommandTimeout = 600
			CALL LDExeCute(GBL_EXEString,0)
			Time2=Timer

			select case DEF_UsedDataBase
				case 0:
					Set Rs = LDExeCute("select @@rowcount",0)
					RowCount = Rs(0)
					Rs.Close
				case 2:
					Set Rs = LDExeCute("select ROW_COUNT()",0)
					RowCount = Rs(0)
					Rs.Close
				case Else
					RowCount = "<font color=ff0000>δ֪</font>"
			End select
			Set Rs = Nothing
			If err.number<>0 Then
				Response.Write "<p><br><span style=""FONT-FAMILY: ����; FONT-SIZE: 12px;""><font color=ff0000><b>���ݿ��������ʧ�ܣ�</b></font><p>"&err.description & "</span>"
				err.clear
			Else
				Response.Write "<p><br><span style=""FONT-FAMILY: ����; FONT-SIZE: 12px;""><font color=008800><b>�������ݿ���������ɹ�����Ӱ��<font color=ff0000>" & RowCount & "</font>�����ݣ���ʱ" & (Time2-Time1)*1000 & "����!</b></font></span><hr size=1>" & PrintTrueText(GBL_EXEString) & "<hr size=1>" & VbCrLf
			End If
		Else
			Response.Write "<p><br><font color=ff0000><b>�����Ϊ��!</b></font>"
		End If
		DisplayStringForm
	Else
		DisplayStringForm
	End If

End Function

Function DisplayStringForm
%>
<p>
<form action=ExecuteString.asp method="post">
	��ִ��SQL���(���棺ִ�����Ҫ���С��!) <p>
	<textarea name=GBL_EXEString rows=8 cols=61 class=fmtxtra><%If GBL_EXEString <> "" Then Response.Write VbCrLf & htmlEncode(GBL_EXEString)%></textarea>
	<input name=submitflag type=hidden value="Dieos9xsl29LO_8">
	<p>
	<input type=submit value="ִ��" class=fmbtn> <input type=reset value="ȡ��" class=fmbtn>
</form>
<%
End Function

Function PrintTrueText(tempString)

	If tempString<>"" Then
		PrintTrueText=Replace(Replace(Replace(Replace(Replace(htmlEncode(tempString),VbCrLf & " ","<br>" & "&nbsp;"),VbCrLf,"<br>" & VbCrLf),"   "," &nbsp; "),"  "," &nbsp;"),chr(9)," &nbsp; &nbsp; &nbsp;")

		If Left(PrintTrueText,1) = chr(32) Then
			PrintTrueText = "&nbsp;" & Mid(PrintTrueText,2)
		End If
	Else
		PrintTrueText=""
	End If

End Function%>