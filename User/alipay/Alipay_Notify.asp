<!-- #include file=../../inc/BBSSetup.asp -->
<!-- #include file=../../inc/Board_Popfun.asp -->
<%
DEF_BBS_HomeUrl = "../../"

Main

Sub Main

	Dim Key,partner,out_trade_no,total_fee,receive_name
	Dim receive_address,receive_zip,receive_phone,receive_mobile
	Dim alipayNotifyURL,Retrieval,ResponseTxt,varItem,mystr
	Dim Count,i,j,md5str,mysign,trade_status,returnTxt
	Dim TOEXCELLR
	dim strsss
	Dim minmax,minmaxSlot,mark,temp,value


	'���ܣ���������з�����֪ͨҳ��
	'�汾��2.0
	'���ڣ�2008-10-24
	'���ߣ�֧������˾���۲�����֧���Ŷ�
	'��ϵ��0571-26888888
	'��Ȩ��֧������˾

	key="o48habnndc8yr4jtyf9g1p02hlt7fs7h"         '֧������ȫ������
	partner="2088002030498170"     '֧��������id 
 
	out_trade_no	=DelStr(Request.Form("out_trade_no"))      '��ȡ������
	total_fee		=DelStr(Request.Form("total_fee"))         '��ȡ֧�����ܼ۸�
	'�����ȡ��������������д ���� =DelStr(Request.Form("��ȡ������"))
	
	'*******************�ж���Ϣ�ǲ���֧��������***********************
	alipayNotifyURL = "http://notify.alipay.com/trade/notify_query.do?"
	alipayNotifyURL = alipayNotifyURL &"partner=" & partner & "&notify_id=" & request.Form("notify_id")
		Set Retrieval = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
	    Retrieval.setOption 2, 13056 
	    Retrieval.open "GET", alipayNotifyURL, False, "", "" 
	    Retrieval.send()
	    ResponseTxt = Retrieval.ResponseText
		Set Retrieval = Nothing
	'*******************************************************************
	
	'*******************��ȡ֧����POST����֪ͨ��Ϣ**********************
	For Each varItem in Request.Form
		mystr=varItem&"="&Request.Form(varItem)&"^"&mystr
	Next 
	If mystr<>"" Then 
		mystr=Left(mystr,Len(mystr)-1)
	End If 
	mystr = SPLIT(mystr, "^")
	Count=ubound(mystr)
	'�Բ�������
	For i = Count TO 0 Step -1
		minmax = mystr( 0 )
		minmaxSlot = 0
		For j = 1 To i
			mark = (mystr( j ) > minmax)
			If mark Then 
				minmax = mystr( j )
				minmaxSlot = j
			End If 
		Next
		If minmaxSlot <> i Then 
			temp = mystr( minmaxSlot )
			mystr( minmaxSlot ) = mystr( i )
			mystr( i ) = temp
		End If
	Next
	'����md5ժҪ�ַ���
	For j = 0 To Count Step 1
		value = SPLIT(mystr( j ), "=")
		If  value(1)<>"" And value(0)<>"sign" And value(0)<>"sign_type"  Then
			If j=Count Then
				md5str= md5str&mystr( j )
			Else 
				md5str= md5str&mystr( j )&"&"
			End If 
		End If 
	Next
	md5str=md5str&key
	mysign=md5(md5str)
	
	'*************************����״̬���ش���*************************
	If mysign=request.Form("sign") And ResponseTxt="true" Then 	
		If request.Form("trade_status") = "TRADE_FINISHED" Then 
			'�ڴ˴���ӣ�����ɹ�,�������ݿ����  
		 	returnTxt	= "success"
		 	CALL Alipay_UpdateSellList(out_trade_no,total_fee)
		Else
			returnTxt	= "fail"
		End If
		Response.Write returnTxt
	Else
	response.write "fail"
	End If 
	'*******************************************************************
	 'д�ı���������ԣ�����վ����Ҳ���Ըĳɴ������ݿ⣩
	TOEXCELLR=TOEXCELLR&md5str&"MD5���:"&mysign&"="&request.Form("sign")&"--ResponseTxt:"&ResponseTxt
	
	strsss = VbCrLf & "ResponseTxt:" & ResponseTxt & VbCrLf
	strsss = strsss & "mysign:" & mysign & VbCrLf
	strsss = strsss & "mysign_form:" & request.Form("sign") & VbCrLf
	strsss = strsss & "trade_status:" & request.Form("trade_status") & VbCrLf
	strsss = strsss & "time:" & DEF_Now & VbCrLf
	strsss = strsss & "out_trade_no:" & out_trade_no & VbCrLf
	strsss = strsss & "total_fee:" & total_fee & VbCrLf

	'CALL ADODB_SaveToFile(TOEXCELLR&strsss,"alipayto/Notify_DATA/"&replace(now(),":","")&".txt")

End Sub

Function DelStr(Str)
	If IsNull(Str) Or IsEmpty(Str) Then
		Str	= ""
	End If
	DelStr	= Replace(Str,";","")
	DelStr	= Replace(DelStr,"'","")
	DelStr	= Replace(DelStr,"&","")
	DelStr	= Replace(DelStr," ","")
	DelStr	= Replace(DelStr,"��","")
	DelStr	= Replace(DelStr,"%20","")
	DelStr	= Replace(DelStr,"--","")
	DelStr	= Replace(DelStr,"==","")
	DelStr	= Replace(DelStr,"<","")
	DelStr	= Replace(DelStr,">","")
	DelStr	= Replace(DelStr,"%","")
End Function


Sub Alipay_UpdateSellList(out_trade_no,total_fee)

	OpenDatabase
	If isNumeric(out_trade_no) = 0 Then out_trade_no = 0
	out_trade_no = Fix(cCur(out_trade_no))
	If isNumeric(total_fee) = 0 Then total_fee = 0
	total_fee = Fix(cCur(total_fee))
	
	Dim Rs,UserName,PayPoints,GetPoints,PayFlag
	Set Rs = LdExeCute(sql_select("Select UserName,PayPoints,GetPoints,PayFlag From LeadBBS_SellList Where PID=" & out_trade_no,1),0)
	If Not Rs.Eof Then
		UserName = Rs(0)
		PayPoints = cCur(Rs(1))
		GetPoints = cCur(Rs(2))
		PayFlag = cCur(Rs(3))
		Rs.Close
		Set Rs = Nothing
	Else
		Rs.Close
		Set Rs = Nothing
		GetPoints = 0
	End If
	
	If PayFlag = 0 and PayPoints = total_fee Then '����Ӧ,����δ֧��������¸��¶������û�״̬
		CALL LdExeCute("Update LeadBBS_SellList Set PayFlag=1 Where PID=" & out_trade_no,1)
		CALL LdExeCute("Update LeadBBS_User Set CharmPoint=CharmPoint+" & GetPoints & " Where UserName='" & Replace(UserName,"'","''") & "'",1)
	End If
	
	CloseDatabase

End Sub
%>