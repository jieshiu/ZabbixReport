#!/usr/bin/python
# -*- coding: utf-8 -*-

import socket
import email
import mimetypes
from email.MIMEBase import MIMEBase
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMEImage import MIMEImage
from email import Utils
import base64
import smtplib
         
def sendEmail(authInfo, fromAdd, toAdd, subject, plainText, htmlText, attpath ,attName,attName2):

        strFrom = fromAdd
        strTo = toAdd

        server = authInfo.get('server')        
        user = authInfo.get('user')
        passwd = authInfo.get('password')

        if not (server and user and passwd) :
                print 'incomplete login info, exit now'
                return

        msgRoot = MIMEMultipart('related')
        msgRoot['Subject'] = subject
        msgRoot['From'] = strFrom
        msgRoot['To'] = ', '.join(strTo)
        msgRoot.preamble = 'This is a multi-part message in MIME format.'

        msgAlternative = MIMEMultipart('alternative')
        msgRoot.attach(msgAlternative)

        msgText = MIMEText(plainText, 'plain', 'utf-8')
        msgAlternative.attach(msgText)

        msgText = MIMEText(htmlText, 'html', 'utf-8')
        msgAlternative.attach(msgText)

        fp = open(attpath+attName, 'rb')
        msgImage1 = MIMEBase('application',"octet-stream")
        msgImage1.set_payload(fp.read())
        fp.close()
        msgImage1.add_header('Content-Disposition', 'attachment; filename='+attName)
        email.Encoders.encode_base64(msgImage1)
        msgRoot.attach(msgImage1)
        
        if attName2 != '':
           fp = open(attpath+attName2, 'rb')
           msgImage = MIMEBase('application',"octet-stream")
           msgImage.set_payload(fp.read())
           fp.close()
           msgImage.add_header('Content-Disposition', 'attachment; filename='+attName2)
           email.Encoders.encode_base64(msgImage)
           msgRoot.attach(msgImage)
          
        #smtp = smtplib.SMTP()
        #smtp.connect(server)
        #smtp.ehlo()        
        #smtp.login(user, passwd)
        #smtp.sendmail(strFrom, strTo, msgRoot.as_string())
        #smtp.quit()
        
        
        smtp=smtplib.SMTP_SSL(server,465)        
        #smtp.set_debuglevel(1)
        smtp.login(user,passwd)#认证
        smtp.sendmail(strFrom,strTo,msgRoot.as_string())
        smtp.quit()
        return