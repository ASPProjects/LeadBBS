<%
'jmail.message��ʽ����
const DEF_MAIL_smtpUser = "" '���ʼ�������ʹ��SMTP������֤ʱ���õĵ�¼�ʻ���
const DEF_MAIL_smtpPass = "" 'ʹ��SMTP������֤ʱ���õĵ�¼���롣
const DEF_MAIL_smtpHost = "" '�ʼ���������ַ(IP������)
const DEF_MAIL_FromName = "LeadBBS" '�����˵����ƣ�������д����վ������

Function SendJmail_Message(Email,Topic,MailBody)

	Dim msg
	set msg = Server.CreateOBject( "JMail.Message" )
	msg.Logging = true
	msg.silent = false
	msg.From = DEF_MAIL_smtpUser
	msg.FromName = DEF_MAIL_FromName
	msg.AddRecipient Email
	msg.Subject = Topic
	msg.Charset="gb2312"
	msg.ContentType = "text/html"
	msg.Body = MailBody
	msg.Priority = 1
	msg.MailServerUserName = DEF_MAIL_smtpUser
	msg.MailServerPassword = DEF_MAIL_smtpPass
	if not msg.Send( DEF_MAIL_smtpHost ) then 'mail server
		Response.write "<pre>" & msg.log & "</pre>"
		SendJmail_Message = 0
	else
		SendJmail_Message = 1
	end if
	Set msg = Nothing

End Function


'jmail.smtpmail����
Function SendJmail(Email,Topic,MailBody)

	If DEF_MAIL_smtpUser <> "" Then
		SendJmail = SendJmail_Message(Email,Topic,MailBody)
		exit function
	end if
	Dim JMail
	'on error resume next
	Set JMail = Server.CreateObject("JMail.SMTPMail")
	JMail.LazySend = true
	JMail.silent = false
	JMail.Charset = "gb2312"
	JMail.ContentType = "text/html"
	JMail.Sender = "mail377234@yourmail.com" '��Ϊ�������
	JMail.ReplyTo = "mail377234@yourmail.com" '��Ϊ�������
	JMail.SenderName = "LeadBBS�ʼ�����"
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
	Un = "qfy@yp.cn"  '�����ʼ���������¼��������Ҫ����

	Dim EI
	Set EI = server.CreateObject("easymail.Users")
	Tid = EI.Login(un)
	Set EI = Nothing
	Mailsend.createnew Un,Tid '�����˺�,��ʱID
	Mailsend.CharSet = "gb2312"  '����
	Mailsend.MailName = "LeadBBS"  '��������

	Mailsend.EM_BackAddress = "" '�ʼ��ظ���ַ
	Mailsend.EM_Bcc = "" '���͵�ַ
	Mailsend.EM_Cc = "" '���͵�ַ
	Mailsend.EM_OrMailName = "" 'ԭ�ʼ���
	Mailsend.EM_Priority = "Normal" '�ʼ���Ҫ��	
	Mailsend.EM_ReadBack = false '�Ƿ��ȡȷ��,�Һ���(�ޱ�ϵͳ���û�)	
	Mailsend.EM_SignNo = -1  'ʹ��ǩ�������
	
	Mailsend.EM_Subject = Topic '����
	Mailsend.EM_Text = TextBody '����
	Mailsend.EM_HTML_Text = MailBody 'HTML�ʼ�����
	Mailsend.useRichEditer = true '���͵��Ƿ�ΪHTML��ʽ�ʼ�

	Mailsend.EM_TimerSend = ""  '��ʱ���͵�ʱ��
	Mailsend.EM_To = Email '�ռ��˵�ַ
	Mailsend.ForwardAttString = "" 'ת���ʼ�ʱ��ԭ����

	Mailsend.AddFromAttFileString = "" '���������洢�е��ļ���

	Mailsend.SystemMessage = false '�Ƿ���ϵͳ�ʼ�

	Mailsend.SendBackup = false '�Ƿ񱣴淢���ʼ�
	
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
	objCDOMail.From ="mail377234@yourmail.com" '��Ϊ�������
	objCDOMail.To = Email
	objCDOMail.Subject = Topic

	objCDOMail.Body = TextBody

	objCDOMail.Send
	Set objCDOMail = Nothing
	SendCDOMail = 1

End Function
%>