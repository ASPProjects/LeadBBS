<!-- #include file=../inc/BBSsetup.asp -->
<!-- #include file=inc/Canvassafecode.asp -->
<%
Response.Expires=0
DEF_BBS_HomeUrl = "../"
Dim RndNumber,RndNumberStr,Rs

Function Code_RndNumber(tp)

	Dim sCode
	Randomize
	Select Case tp
	Case 0:
		sCode = Fix(Rnd*9999)+1
	Case 1,2:
		Dim i,cCode,cAmount
		If tp = 1 Then
			cAmount = 23
			cCode = "ABCDEFGHJKLMNPQRSTUVWXY"
		Else
			cAmount = 30
			cCode = "3456789ABCDEFGHJKLMNPQRSTUVWXY"
		End If
		For i = 0 To 3
			sCode = sCode &""& Mid(cCode, Int(Rnd * cAmount) + 1, 1) 
		Next
	Case 3:
		sCode = Rnd_cn(2)
	End Select
	
	Session(DEF_MasterCookies & "RndNum") = sCode
	Session(DEF_MasterCookies & "RndNum_par") = ""
	Code_RndNumber = sCode

End Function

If (DEF_EnableAttestNumber = 1 or DEF_EnableAttestNumber = 3) Then
	'DisplayGflSDK_num
	'DisplayGflSDK
	'DisplayGflSDK_cn
	'DisplayBMP
	
	'DisplayCanvasCode
	RndNumber_Jpeg
Else
	DisplayCanvasCode
End If


Sub DisplayGflSDKgif

	Dim rndStr
	rndStr = Code_RndNumber(1)
	Dim MyObj
	Set MyObj = Server.CreateObject("GflAx190.GflAx")
	MyObj.EnableLZW = True
	MyObj.LoadBitmap(server.mappath(DEF_BBS_HomeUrl & "images/null.gif"))

	MyObj.NewBitmap 32, 12
	MyObj.FontBold = True
	MyObj.FontName = "����"
	MyObj.FontSize = 12
	
	Dim c,N
	For N = 1 to 4
		c = Rnd*255
		MyObj.TextOut Mid(rndStr,N,1),2+7*N-7,-1,c
	Next
	
	MyObj.UseTransparency = true
	MyObj.ChangeColorDepth 8
	Response.ContentType="image/gif"
	MyObj.SaveFormat = 2
	MyObj.MaskColor = 255*256*256+255*256+255
	Response.BinaryWrite MyObj.SendBinary
	Set MyObj = Nothing

End Sub

Sub DisplayGflSDK

	Dim rndStr
	rndStr = Code_RndNumber(1)
	Dim MyObj
	Set MyObj = Server.CreateObject("GflAx190.GflAx")
	MyObj.EnableLZW = True
	MyObj.LoadBitmap(server.mappath(DEF_BBS_HomeUrl & "images/null.jpg"))

	Dim fontName
	fontName = Array("Comic Sans MS","Arial Black","Century Gothic","Courier","Courier New","Times New Roman","Impact","Verdana")

	MyObj.NewBitmap 32, 12
	MyObj.FontBold = True
	Randomize
	MyObj.FontName = fontName(Fix(Rnd*7)+1)
	MyObj.FontSize = 15
	
	Dim c,N
	For N = 1 to 4
		c = Rnd*255
		MyObj.TextOut Mid(rndStr,N,1),2+7*N-7,-1,c
	Next
	
	MyObj.UseTransparency = true
	MyObj.ChangeColorDepth 8
	MyObj.SaveJPEGQuality = 20
	Response.ContentType="image/jpg"
	MyObj.SaveFormat = 1
	'MyObj.MaskColor = 255*256*256+255*256+255
	Response.BinaryWrite MyObj.SendBinary
	Set MyObj = Nothing

End Sub




Function Rnd_cn(num)

dim str
str = "���������������������������������������������������������������°ðİŰǰȰɰʰ˰̰ͰΰϰаѰӰ԰հְװذٰڰ۰ܰݰ߰���������������������������������������������������������������������������������������������������±ıűƱǱȱɱʱ˱̱ͱαϱбұӱԱձֱױرٱڱ۱ܱ߱���������������������������������������������������������������������������������������²òĲŲƲǲȲɲʲ˲̲ͲβϲвѲҲӲԲղֲײزٲڲ۲ܲݲ޲߲��������������������������������������������������������������������������������������������������³óĳųƳǳȳɳʳ˳̳ͳγϳгѳҳԳճֳ׳سٳڳ۳ܳݳ޳߳�����������������������������������������������������������������������������������������´ŴƴǴȴɴʴ˴̴ʹδϴдҴӴԴմִ״ٴڴܴݴ޴ߴ����������������������������������������������������������������������������������������������������µõĵŵƵǵȵɵʵ˵̵͵εϵеѵҵֵ׵صٵڵ۵ܵݵ޵ߵ���������������������������������������������������������������������������������������������¶öĶŶǶȶɶʶ˶̶Ͷζ϶жѶҶӶԶֶ׶ضٶڶܶݶ޶���������������������������������������������������������������������������������������������·÷ķŷƷǷȷɷʷ˷̷ͷηϷзѷҷԷַ׷طٷ۷ܷݷ޷߷�����������������������������������������������������������������������������������������������¸øĸƸǸȸɸʸ˸̸͸θϸиѸҸԸոָ׸ظٸڸ۸ܸݸ߸�����������������������������������������������������������������������������������������������¹ùĹŹǹȹɹʹ˹̹͹ιϹйѹҹӹԹչֹ׹عٹڹ۹ܹݹ޹߹�����������������������������������������������������������������������������������������������úĺźƺǺɺʺ˺̺ͺκϺкӺ׺غٺںۺܺݺ޺ߺ���������������������������������������������������������������������������������������������������ûĻŻƻɻʻ˻̻ͻλлѻһӻԻջֻ׻ػٻڻۻݻ޻߻������������������������������������������������������������������������������������������������¼üļżƼǼȼɼʼ˼̼ͼμмѼҼӼ׼ؼټڼۼܼݼ޼߼�������������������������������������������������������������������������������������������������ĽŽƽǽȽɽʽ̽ͽνϽнҽӽԽֽ׽ؽٽھ������������������������������������������������¾þľžƾȾɾ˾̾;ξϾоӾԾվ־׾ؾپھ۾ܾݾ޾߾��������������������۽ܽݽ޽߽�������������������������������������������������������������������������������������������������������������������������������¿ÿĿſƿǿȿɿʿ˿̿ͿοϿпѿҿӿԿտֿ׿ؿٿڿۿܿݿ޿߿�������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������¢£¤¥¦§©ª«¬­®¯³µ¶·¹¼½¿��������������������������������������������������������������������������������������������������������������������áâãäåæçèéëìðñòóôõö÷øùúûüýÿ����������������������������������������������������������������������������������������������������������������ġĢģĤĥĦħĨĩĪīĬĭĮįİıĲĳĴĵĶķĸĹĺĻļĽľĿ������������������������������������������������������������������������������������������������������������������������������šŢţŤťŦŧŨũŪūŬŭŮůŰűŲųŴŵŶŷŸŹźŻżŽžſ������������������������������������������������������������������������������������������������������������������������������ơƢƣƤƥƦƧƨƩƪƫƬƭƮƯưƱƲƳƴƵƶƷƸƹƺƻƼƽƾƿ������������������������������������������������������������������������������������������������������������������������������ǢǣǤǥǦǧǨǩǪǫǬǭǮǯǰǱǲǳǴǵǶǷǸǹǺǻǼǽǾǿ������������������������������������������������������������������������������������������������������������������������������ȡȢȣȤȥȦȧȨȩȪȫȬȭȮȯȰȱȲȳȴȵȶȷȸȹȺȻȼȽȾȿ������������������������������������������������������������������������������������������������������������������������������ɡɢɣɤɥɦɧɨɩɪɫɬɭɮɯɰɱɲɳɴɵɶɷɸɹɺɻɼɽɾɿ������������������������������������������������������������������������������������������������������������������������������ʡʢʣʤʥʦʧʨʩʪʫʬʭʮʯʰʱʲʳʴʵʶʷʸʹʺʻʼʽʾʿ������������������������������������������������������������������������������������������������������������������������������ˡˢˣˤ˥˦˧˨˩˪˫ˬ˭ˮ˯˰˱˲˳˴˵˶˷˸˹˺˻˼˽˾˿������������������������������������������������������������������������������������������������������������������������������̴̵̶̷̸̡̢̧̨̣̤̥̦̩̪̫̬̭̮̯̰̱̲̳̹̺̻̼̽̾̿������������������������������������������������������������������������������������������������������������������������������ͣͤͥͦͧͨͩͪͫͬͭͮͯ͢͡ͰͱͲͳʹ͵Ͷͷ͸͹ͺͻͼͽ;Ϳ������������������������������������������������������������������������������������������������������������������������������Ρ΢ΣΤΥΦΧΨΩΪΫάέήίΰαβγδεζηθικλμνξο������������������������������������������������������������������������������������������������������������������������������ϡϢϣϤϥϦϧϨϩϪϫϬϭϮϯϰϱϲϳϴϵ϶ϷϸϹϺϻϼϽϾϿ������������������������������������������������������������������������������������������������������������������������������СТУФХЦЧШЩЪЫЬЭЮЯабвгдежзийклмноп������������������������������������������������������������������������������������������������������������������������������ѡѢѣѤѥѦѧѨѩѪѫѬѭѮѯѰѱѲѳѴѵѶѷѸѹѺѻѼѽѾѿ������������������������������������������������������������������������������������������������������������������������������ҡҢңҤҥҦҧҨҩҪҫҬҭҮүҰұҲҳҴҵҶҷҸҹҺһҼҽҾҿ������������������������������������������������������������������������������������������������������������������������������ӡӢӣӤӥӦӧӨөӪӫӬӭӮӯӰӱӲӳӴӵӶӻӼӽӾӿ������������������������������������������������������������������������������������������������������������������������������ԡԢԣԤԥԦԧԨԩԪԫԬԭԮԯ԰ԱԲԳԴԵԶԷԸԹԺԻԼԽԾԿ������������������������������������������������������������������������������������������������������������������������������աբգդեզէըթժիլխծկհձղճմյնշոչջռսվտ������������������������������������������������������������������������������������������������������������������������������ְֱֲֳִֵֶַָֹֺֻּֽ֢֣֤֥֦֧֪֭֮֡֨֩֫֬֯־ֿ������������������������������������������������������������������������������������������������������������������������������סעףפץצקרשת׬׭׮ׯװױײ׳״׵׶׷׸׹׺׻׼׽׾׿����������������������������������������������������������������������������������������������������������"

Randomize
Dim R,N,L
L = len(str)
For N = 1 to num
R = R & Mid(Str,Fix(Rnd*L),1)
Next
Rnd_cn = R


End Function

Sub DisplayGflSDK_cn

	Dim rndStr
	rndStr = Code_RndNumber(2)
	Dim MyObj
	Set MyObj = Server.CreateObject("GflAx190.GflAx")
	MyObj.EnableLZW = True
	MyObj.LoadBitmap(server.mappath(DEF_BBS_HomeUrl & "images/null.jpg"))

	Dim fontName
	fontName = Array("����","����")

	MyObj.NewBitmap 32, 26
	MyObj.FontBold = True
	Randomize
	MyObj.FontName = fontName(Fix(Rnd*1))
	MyObj.FontSize = 25
	
	Dim c,N
	For N = 1 to Len(rndStr)
		c = Rnd*255
		MyObj.TextOut Mid(rndStr,N,1),2+25*N-25,-1,c
	Next
	
	MyObj.UseTransparency = true
	MyObj.ChangeColorDepth 8
	MyObj.SaveJPEGQuality = 20
	Response.ContentType="image/jpg"
	MyObj.SaveFormat = 1
	'MyObj.MaskColor = 255*256*256+255*256+255
	Response.BinaryWrite MyObj.SendBinary
	Set MyObj = Nothing

End Sub

Sub DisplayGflSDK_num

	Dim MyObj
	Set MyObj = Server.CreateObject("GflAx190.GflAx")
	MyObj.EnableLZW = True
	MyObj.LoadBitmap(server.mappath(DEF_BBS_HomeUrl & "images/null.jpg"))


	Dim SafeCode,a,b,b1,tp
	Randomize
	tp = Fix(Rnd*5)
	Randomize
	a = Fix(Rnd*13)
	Randomize
	b = Fix(Rnd*9)
	b1 = Fix(Rnd*9)
	
	Dim numName
	fontName = Array("��","��","��","��","��","5","��","��","��","��","ʮ","����","12","����")

	Select Case tp
	case 0:
		SafeCode = "" & (a * b +b1)
		Session(DEF_MasterCookies & "RndNum") = SafeCode
		Session(DEF_MasterCookies & "RndNum_par") = ""
		RndNumberStr = fontName(a) & "��" & fontName(b) & "��" & fontName(b1) & "����"
	Case 1:
		SafeCode = "" & (a*a)
		Session(DEF_MasterCookies & "RndNum") = SafeCode
		Session(DEF_MasterCookies & "RndNum_par") = ""
		RndNumberStr = "����" & fontName(a) & "��ƽ������"
	Case 2:
		SafeCode = "" & (a*a+b1)
		Session(DEF_MasterCookies & "RndNum") = SafeCode
		Session(DEF_MasterCookies & "RndNum_par") = ""
		RndNumberStr = "" & fontName(a) & "ƽ����" & fontName(b1) & "����"
	case 3:
		SafeCode = "" & (a * b)
		Session(DEF_MasterCookies & "RndNum") = SafeCode
		Session(DEF_MasterCookies & "RndNum_par") = ""
		RndNumberStr = fontName(b) & "��" & fontName(a) & "����"
	case 4:
		SafeCode = "" & Rnd_cn(1)
		Session(DEF_MasterCookies & "RndNum") = SafeCode
		Session(DEF_MasterCookies & "RndNum_par") = ""
		RndNumberStr = "������:" & SafeCode
		
	End Select

	Dim fontName
	fontName = Array("����","����")

	MyObj.NewBitmap len(RndNumberStr)*21, 26
	MyObj.FontBold = True
	Randomize
	MyObj.FontName = fontName(Fix(Rnd*2))
	MyObj.FontSize = 18+rnd*3
	
	
	Dim r,g,blue,N,pos,color
	For N = 1 to len(RndNumberStr)
		
		Randomize
		tp = Fix(Rnd*2)
		Select Case tp
		Case 0:
			Randomize
			r = Rnd*55
			Randomize
			g=rnd*55
			Randomize
			blue=rnd*55
			pos = rnd*5
			color = r*256*256+g*256+blue
			'MyObj.TextOut Mid(RndNumberStr,N,1),2+18*N-18,-1+pos+rnd*2-rnd*2,222*256*256+222*256+255
			MyObj.TextOut Mid(RndNumberStr,N,1),2+18*N-18,-1+pos+1,color
			MyObj.TextOut Mid(RndNumberStr,N,1),2+18*N-18+1,-1+pos,color
			MyObj.TextOut Mid(RndNumberStr,N,1),2+18*N-18+1,-1+pos+1,color
			MyObj.TextOut Mid(RndNumberStr,N,1),2+18*N-18-1,-1+pos,color
			MyObj.TextOut Mid(RndNumberStr,N,1),2+18*N-18,-1+pos-1,color
			MyObj.TextOut Mid(RndNumberStr,N,1),2+18*N-18-1,-1+pos-1,color
			MyObj.TextOut Mid(RndNumberStr,N,1),2+18*N-18+1,-1+pos-1,color
			MyObj.TextOut Mid(RndNumberStr,N,1),2+18*N-18-1,-1+pos+1,color
			MyObj.TextOut Mid(RndNumberStr,N,1),2+18*N-18,-1+pos,255*256*256+255*256+255
		Case 1:
			Randomize
			r = Rnd*55
			Randomize
			g=rnd*55
			Randomize
			blue=rnd*55
			pos = rnd*5
			color = r*256*256+g*256+blue
			MyObj.TextOut Mid(RndNumberStr,N,1),2+18*N-18,-1+pos,color
		End Select
	Next
	
	MyObj.UseTransparency = true
	MyObj.ChangeColorDepth 8
	MyObj.SaveJPEGQuality = 58
	Response.ContentType="image/jpg"
	MyObj.SaveFormat = 1
	'MyObj.MaskColor = 255*256*256+255*256+255
	Response.BinaryWrite MyObj.SendBinary
	Set MyObj = Nothing

End Sub

Sub RndNumber_Jpeg

	Dim Jpeg
	Set Jpeg = Server.CreateObject("Persits.Jpeg")
	Jpeg.New 50,21,0
	Jpeg.Canvas.Font.Color = &HFFFFFF
	Jpeg.Canvas.Font.Family = "Impact"
	'Jpeg.Canvas.Font.Family = "SharKbait Regular"
	Jpeg.Canvas.Font.Bold = false
	Jpeg.Canvas.Font.Size = 19
	Jpeg.Canvas.Font.Quality = 4
	Jpeg.Canvas.PrintText 6, 1, Code_RndNumber(1)
	jpeg.SendBinary
	jpeg.Close
	Set Jpeg = Nothing

End Sub

%>