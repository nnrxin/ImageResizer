/* 函数:利用CDO.Message发送邮件
 * param emailConfig  对象
 *       .from        设置发信人的邮箱
 *       .to          设置收信人的邮箱，多个邮箱之间用分号隔开
 *       .bcc         设置密送人的邮箱
 *       .cc          设置抄送人的邮箱
 *       .subject     设定邮件的主题
 *       .htmlBody    使用html格式的内容
 *       .textBody    使用文本格式的内容
 *       .attachments 附件地址(数组)
 * param cfgFields       对象
 *       .smtpserver     发件人SMTP服务器地址
 *       .smtpserverport 发件人SMTP服务器端口
 *       .sendusername   发件邮箱账号
 *       .sendpassword   发件邮箱密码
 * returns
 */
SendEmail(emConfigs, cfgFields) {
	try {
		pmsg := ComObject("CDO.Message")
		pmsg.From := emConfigs.from ; 设置发信人的邮箱
		pmsg.To := emConfigs.to ; 设置收信人的邮箱，多个邮箱之间用分号隔开
		pmsg.CC := emConfigs.HasProp("cc") ? emConfigs.cc : "" ; 设置抄送人的邮箱
		pmsg.BCC := emConfigs.HasProp("bcc") ? emConfigs.bcc : "" ; 设置密送人的邮箱
		pmsg.bodyPaRT.Charset := "UTF-8" ; 定义国际通用字体
	
		if emConfigs.HasProp("subject") and emConfigs.subject != "" ; 邮件标题
			pmsg.Subject := emConfigs.subject
		else if emConfigs.HasProp("attachments") and emConfigs.attachments.Length >= 1 { ; 无标题时使用附件名称
			SplitPath emConfigs.attachments[1],,,, &nameNoExt
			pmsg.Subject := nameNoExt
		}
		else
			pmsg.Subject := "untitled"
	
		if emConfigs.HasProp("htmlBody") ; 邮件内容
			pmsg.htmlbody := emConfigs.htmlBody ; 使用html格式
		else
			pmsg.TextBody := emConfigs.HasProp("textBody") ? emConfigs.textBody : "" ; 邮件正文，使用文本格式发送邮件
	
		if emConfigs.HasProp("attachments") { ; 邮件附件
			for i, attachment in emConfigs.attachments {
				if FileExist(attachment) and !DirExist(attachment)
					pmsg.AddAttachment(attachment) ; 添加附件
			}
		}
		fields := {}
		fields.smtpserver := cfgFields.smtpserver ; 发件人SMTP服务器地址
		fields.smtpserverport := cfgFields.HasProp("smtpserverport") ? cfgFields.smtpserverport : 25 ; 发件人SMTP服务器端口
		fields.smtpusessl := cfgFields.HasProp("smtpusessl") ? cfgFields.smtpusessl : true ; False
		fields.sendusing := cfgFields.HasProp("sendusing") ? cfgFields.sendusing : 2 ; 发送端口
		fields.smtpauthenticate := 1 ; 需要提供用户名和密码，0是不提供
		fields.sendusername := cfgFields.sendusername ; 发件邮箱账号
		fields.sendpassword := cfgFields.sendpassword ; 发件邮箱密码
		fields.smtpconnectiontimeout := 45
		schema := "http://schemas.microsoft.com/cdo/configuration/" ; 微软服务器网址
		pfld := pmsg.Configuration.Fields
		for field, value in fields.OwnProps()
			pfld.Item[schema . field] := value
		pfld.Update()
		pmsg.Send() ; 执行发送
		pmsg := ""
	} catch as e {
		pmsg := ""
		return "邮件发送失败:`nwhat: " e.what "`nfile: " e.file . "`nline: " e.line "`nmessage: " e.message "`nextra: " e.extra
	}	
}



/*
sina.com:
POP3服务器地址:pop3.sina.com.cn（端口：110）
SMTP服务器地址:smtp.sina.com.cn（端口：25） 


sina.cn：
POP3服务器地址:pop3.sina.com（端口：110）                                     ------- > pop.sina.com
SMTP服务器地址:smtp.sina.com（端口：25） 

sinaVIP：
POP3服务器:pop3.vip.sina.com （端口：110）
SMTP服务器:smtp.vip.sina.com （端口：25）

sohu.com:
POP3服务器地址:pop3.sohu.com（端口：110）
SMTP服务器地址:smtp.sohu.com（端口：25）

126邮箱：
POP3服务器地址:pop.126.com（端口：110）
SMTP服务器地址:smtp.126.com（端口：25）

139邮箱：
POP3服务器地址：POP.139.com（端口：110）
SMTP服务器地址：SMTP.139.com(端口：25)

163.com:
POP3服务器地址:pop.163.com（端口：110）
SMTP服务器地址:smtp.163.com（端口：25）

QQ邮箱
POP3服务器地址：pop.qq.com（端口：110）
SMTP服务器地址：smtp.qq.com（端口：25）

QQ企业邮箱
POP3服务器地址：pop.exmail.qq.com （SSL启用 端口：995）
SMTP服务器地址：smtp.exmail.qq.com（SSL启用 端口：587/465）

yahoo.com:
POP3服务器地址:pop.mail.yahoo.com
SMTP服务器地址:smtp.mail.yahoo.com

yahoo.com.cn:
POP3服务器地址:pop.mail.yahoo.com.cn（端口：995）
SMTP服务器地址:smtp.mail.yahoo.com.cn（端口：587

HotMail
POP3服务器地址：pop3.live.com（端口：995）
SMTP服务器地址：smtp.live.com（端口：587）

gmail(google.com)
POP3服务器地址:pop.gmail.com（SSL启用端口：995）
SMTP服务器地址:smtp.gmail.com（SSL启用 端口：587）

263.net:
POP3服务器地址:pop3.263.net（端口：110）
SMTP服务器地址:smtp.263.net（端口：25）

263.net.cn:
POP3服务器地址:pop.263.net.cn（端口：110）
SMTP服务器地址:smtp.263.net.cn（端口：25）

x263.net:
POP3服务器地址:pop.x263.net（端口：110）
SMTP服务器地址:smtp.x263.net（端口：25）

21cn.com:
POP3服务器地址:pop.21cn.com（端口：110）
SMTP服务器地址:smtp.21cn.com（端口：25）

Foxmail：
POP3服务器地址:POP.foxmail.com（端口：110）
SMTP服务器地址:SMTP.foxmail.com（端口：25）

china.com:
POP3服务器地址:pop.china.com（端口：110）
SMTP服务器地址:smtp.china.com（端口：25）

tom.com:
POP3服务器地址:pop.tom.com（端口：110）
SMTP服务器地址:smtp.tom.com（端口：25）

etang.com:
POP3服务器地址:pop.etang.com
SMTP服务器地址:smtp.etang.com
*/