<%
'jmail.message方式发送
Function SendJmail_Message(Email,Topic,MailBody)

	Dim msg
	set msg = Server.CreateOBject( "JMail.Message" )
	msg.Logging = true
	msg.silent = true
	msg.From = "sender@leadbbs.com"
	msg.FromName = "LeadBBS.COM"
	msg.AddRecipient Email
	msg.Subject = Topic
	msg.Charset="gb2312"
	msg.ContentType = "text/html"
	msg.Body = MailBody
	if not msg.Send( "1.1.1.4" ) then 'mail server
		Response.write "<pre>" & msg.log & "</pre>"
		SendJmail_Message = 0
	else
		SendJmail_Message = 1
	end if
	Set msg = Nothing

End Function


'jmail.smtpmail发送
Function SendJmail(Email,Topic,MailBody)

	Dim JMail
	'on error resume next
	Set JMail = Server.CreateObject("JMail.SMTPMail")
	JMail.LazySend = true
	JMail.silent = true
	JMail.Charset = "gb2312"
	JMail.ContentType = "text/html"
	JMail.Sender = "mail377234@yourmail.com" '改为你的邮箱
	JMail.ReplyTo = "mail377234@yourmail.com" '改为你的邮箱
	JMail.SenderName = "LeadBBS邮件发送"
	JMail.Subject = Topic
	JMail.SimpleLayout = true
	JMail.Body = MailBody
	JMail.Priority = 1
	JMail.AddRecipient Email
	JMail.AddHeader "Originating-IP", GBL_IPAddress
	If JMail.Execute() = false Then
		SendJmail = 0
	Else
		SendJmail = 1
	End If
	JMail.Close
	Set JMail = Nothing

End Function

Function SendEasyMail(Email,Topic,MailBody,TextBody)

	'on error resume next
	dim Mailsend
	set Mailsend = Server.CreateObject("easymail.Mailsend")
	Dim Tid,Un
	Un = "qfy@yp.cn"  '您的邮件服务器登录名，不需要密码

	Dim EI
	Set EI = server.CreateObject("easymail.Users")
	Tid = EI.Login(un)
	Set EI = Nothing
	Mailsend.createnew Un,Tid '邮箱账号,临时ID
	Mailsend.CharSet = "gb2312"  '编码
	Mailsend.MailName = "LeadBBS"  '发件人名

	Mailsend.EM_BackAddress = "" '邮件回复地址
	Mailsend.EM_Bcc = "" '暗送地址
	Mailsend.EM_Cc = "" '抄送地址
	Mailsend.EM_OrMailName = "" '原邮件名
	Mailsend.EM_Priority = "Normal" '邮件重要度	
	Mailsend.EM_ReadBack = false '是否读取确认,挂号信(限本系统内用户)	
	Mailsend.EM_SignNo = -1  '使用签名的序号
	
	Mailsend.EM_Subject = Topic '主题
	Mailsend.EM_Text = TextBody '内容
	Mailsend.EM_HTML_Text = MailBody 'HTML邮件内容
	Mailsend.useRichEditer = true '发送的是否为HTML格式邮件

	Mailsend.EM_TimerSend = ""  '定时发送的时间
	Mailsend.EM_To = Email '收件人地址
	Mailsend.ForwardAttString = "" '转发邮件时的原附件

	Mailsend.AddFromAttFileString = "" '添加自网络存储中的文件名

	Mailsend.SystemMessage = false '是否是系统邮件

	Mailsend.SendBackup = false '是否保存发送邮件
	
	If Mailsend.Send() = false Then
		SendEasyMail = 0
	Else
		SendEasyMail = 1
	End If
	Set Mailsend = nothing

End Function

Function SendCDOMail(Email,Topic,TextBody)

	dim  objCDOMail
	Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
	objCDOMail.From ="mail377234@yourmail.com" '改为你的邮箱
	objCDOMail.To = Email
	objCDOMail.Subject = Topic

	objCDOMail.Body = TextBody

	objCDOMail.Send
	Set objCDOMail = Nothing
	SendCDOMail = 1

End Function
%>