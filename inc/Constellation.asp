<%
Function Constellation(Births)

	Dim Birth
	Birth = Births
	dim BirthDay,BirthMonth
	BirthDay=day(Birth)
	BirthMonth=month(Birth)
	Constellation = "<img src=""" & DEF_BBS_HomeUrl & "images/" & GBL_DefineImage & "cnstl/z"
	Select Case BirthMonth
	case 1
		if BirthDay>=21 then
			Constellation = Constellation & "11.gif"" title=""ˮƿ��"" align=middle />"
		else
			Constellation = Constellation & "10.gif"" title=""ħ����"" align=middle />"
		end if
	case 2
		if BirthDay>=20 then
			Constellation = Constellation & "12.gif"" title=""˫����"" align=middle />"
		else
			Constellation = Constellation & "11.gif"" title=""ˮƿ��"" align=middle />"
		end if
	case 3
		if BirthDay>=21 then
			Constellation = Constellation & "1.gif"" title=""������"" align=middle />"
		else
			Constellation = Constellation & "12.gif"" title=""˫����"" align=middle />"
		end if
	case 4
		if BirthDay>=21 then
			Constellation = Constellation & "2.gif"" title=""��ţ��"" align=middle />"
		else
			Constellation = Constellation & "1.gif"" title=""������"" align=middle />"
		end if
	case 5
		if BirthDay>=22 then
			Constellation = Constellation & "3.gif"" title=""˫����"" align=middle />"
		else
			Constellation = Constellation & "2.gif"" title=""��ţ��"" align=middle />"
		end if
	case 6
		if BirthDay>=22 then
			Constellation = Constellation & "4.gif"" title=""��з��"" align=middle />"
		else
			Constellation = Constellation & "3.gif"" title=""˫����"" align=middle />"
		end if
	case 7
		if BirthDay>=23 then
			Constellation = Constellation & "5.gif"" title=""ʨ����"" align=middle />"
		else
			Constellation = Constellation & "4.gif"" title=""��з��"" align=middle />"
		end if
	case 8
		if BirthDay>=24 then
			Constellation = Constellation & "6.gif"" title=""��Ů��"" align=middle />"
		else
			Constellation = Constellation & "5.gif"" title=""ʨ����"" align=middle />"
		end if
	case 9
		if BirthDay>=24 then
			Constellation = Constellation & "7.gif"" title=""�����"" align=middle />"
		else
			Constellation = Constellation & "6.gif"" title=""��Ů��"" align=middle />"
		end if
	case 10
		if BirthDay>=24 then
			Constellation = Constellation & "8.gif"" title=""��Ы��"" align=middle />"
		else
			Constellation = Constellation & "7.gif"" title=""�����"" align=middle />"
		end if
	case 11
		if BirthDay>=23 then
			Constellation = Constellation & "9.gif"" title=""������"" align=middle />"
		else
			Constellation = Constellation & "8.gif"" title=""��Ы��"" align=middle />"
		end if
	case 12
		if BirthDay>=22 then
			Constellation = Constellation & "10.gif"" title=""ħ����"" align=middle />"
		else
			Constellation = Constellation & "9.gif"" title=""������"" align=middle />"
		end if
	case else
		Constellation=""
	end select

End Function

Function DisplayBirthAnimal(BirthYear)

	Dim Temp,tmp
	Temp = BirthYear mod 12
	tmp = "<img src=""" & DEF_BBS_HomeUrl & "images/" & GBL_DefineImage & "snxa/sx"
	Select Case Temp
		Case 0: tmp=tmp & "9s.gif"" align=middle title=""���"" />"
		Case 1: tmp=tmp & "10s.gif"" align=middle title=""�ϼ�"" />"
		Case 2: tmp=tmp & "11s.gif"" align=middle title=""�繷"" />"
		Case 3: tmp=tmp & "12s.gif"" align=middle title=""����"" />"
		Case 4: tmp=tmp & "1s.gif"" align=middle title=""����"" />"
		Case 5: tmp=tmp & "2s.gif"" align=middle title=""��ţ"" />"
		Case 6: tmp=tmp & "3s.gif"" align=middle title=""����"" />"
		Case 7: tmp=tmp & "4s.gif"" align=middle title=""î��"" />"
		Case 8: tmp=tmp & "5s.gif"" align=middle title=""����"" />"
		Case 9: tmp=tmp & "6s.gif"" align=middle title=""����"" />"
		Case 10: tmp=tmp & "7s.gif"" align=middle title=""����"" />"
		Case 11: tmp=tmp & "8s.gif"" align=middle title=""δ��"" />"
		Case else: tmp = ""
	End Select
	
	DisplayBirthAnimal = tmp

End Function

%>