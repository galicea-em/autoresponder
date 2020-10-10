#!/usr/bin/python3
# -*- encoding: utf-8 -*-
#
# author: Jerzy Wawro

import smtplib
import imaplib
import ssl
#import socket
#import getpass
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email.encoders import encode_base64
from email.mime.application import MIMEApplication
from pprint import pprint
import configparser
import re
import sys
import os

config = None

class Config():

    config_path = "autoresponder.ini"
    config = None

    def __init__(self, debug=False):
      self.config=self.read_config(self.get_config_path())
      if debug:
        self.debug()

    def get_config_path(self):
        if "--config-path" in sys.argv and len(sys.argv) >= 3:
            self.config_path = sys.argv[2]
        if "--help" in sys.argv or "-h" in sys.argv:
            self.display_help_text()
            exit(0)
        if not os.path.isfile(self.config_path):
            self.shutdown_with_error(
                "Configuration file not found. Expected it at '" + self.config_path + "'.")
        return self.config_path

    def debug(self):
      for sec in self.config.sections():
        pprint(sec)
        for par in self.config[sec]:
          pprint('%s = %s' % (par,self.config[sec][par]))

    def read_config(self, file_name):
        try:
            config = configparser.ConfigParser()
            config.read(file_name, encoding="UTF-8")
        except KeyError as e:
            self.shutdown_with_error(
                "Configuration file is invalid! (Key not found: " + str(e) + ")")
        return config

    def shutdown_with_error(self, message):
        message = "Error! " + str(message)
        message += "\nCurrent configuration file path: '" + \
            str(self.config_path) + "'."
        if config is not None:
            message += "\nCurrent configuration: " + str(config)
        print(message)
        exit(-1)

    def log_warning(self, message):
        print("Warning! " + message)

    def display_help_text(self):
        print("Options:")
        print("\t--help: Display this help information")
        print("\t--config-path <path/to/config/file>: "
              "Override path to config file (defaults to same directory as the script is)")
        exit(0)


class SMTP:
    host = ""
    fromWho = ""
    
    login = ""
    password =  ""
    
    server = False
    body=html=jpeg=pdf=doc=None
    jpeg_fn=''
    jpeg_len=0
    pdf_fn=''
    doc_fn=''

    def __init__(self, config):
      self.host=config['SMTP']['smtp.host']
      self.login=config['SMTP']['smtp.username']
      self.password=config['SMTP']['smtp.password']
      self.port=config['SMTP']['smtp.port']
      self.fromWho=config['mail']['mail.from']
      self.subject=config['mail']['mail.subject.prefix']
      self.body=config['mail']['mail.body']

    def read(self,vascii='',vhtml='',vjpg='',vpdf='',vdoc=''):
      if vtresc != '':
        self.body=open(vascii).read() #.decode('utf8')
      if vhtml != '':
        self.html=open(vhtml).read() # .decode('utf8')
      if vjpg != '':
        self.jpeg_fn=vjpg
        f=open(vjpg, 'rb')
        self.jpeg=f.read()
        self.jpeg_len=len(self.jpeg)
        f.close()
      if vpdf != '':
        self.pdf_fn=vpdf
        f=open(vpdf, 'rb')
        self.pdf=f.read()
     #  f.close()
      if vdoc != '':
        self.doc_fn=vdoc
        f=open(vdoc, 'rb')
        self.doc=f.read()
#       f.close()

    
    def connect(self):
        self.server = smtplib.SMTP(self.host)
        context = ssl.SSLContext (ssl.PROTOCOL_TLSv1_1)
#        connection.starttls(context = context)

        self.server.ehlo()
        self.server.starttls(context = context)
        self.server.ehlo()
        self.server.login(self.login, self.password)
    
    def sendMail(self, toWho, subject,  filePaths=[]):
        if not self.server:
            self.connect()

        if self.html==None:
          msg = MIMEMultipart()
        else:
          msg = MIMEMultipart('alternative') # gdy tylko tekst (ASCII):MIMEMultipart()
        msg['To'] = toWho
        msg['From'] = self.fromWho
        msg['Subject'] = subject

        if self.body != None:
          msg.attach( MIMEText(self.body.encode('utf-8'), 'plain', 'utf-8') )
        if self.html != None:
          msg_alt = MIMEMultipart('related')
          msg_alt.attach( MIMEText(self.html.encode('utf-8'), 'html', 'utf-8') )
          msg.attach(msg_alt)
        if self.pdf != None:
          attachFile = MIMEApplication(self.pdf,_subtype="pdf")
#          attachFile = MIMEBase('application', 'pdf')
#          attachFile.set_payload(self.pdf)
#          encode_base64(attachFile)
          attachFile.add_header('Content-Disposition', 'attachment', filename=self.pdf_fn)		
          msg.attach(attachFile)
        if self.doc != None:
          attachFile = MIMEApplication(self.doc,_subtype="msword")
          attachFile.add_header('Content-Disposition', 'attachment', filename=self.doc_fn)		
          msg.attach(attachFile)
        if self.jpeg != None:
          img_iter=0
          attachFile = MIMEImage(self.jpeg, 'jpeg')
          attachFile.add_header('Content-ID', '<image%d>' % img_iter)
          attachFile.add_header('Content-Disposition', 'form-data', name='media', filename=self.jpeg_fn)		
          attachFile.add_header('Content-Length', str(self.jpeg_len))
          msg.attach(attachFile)
        self.server.sendmail(self.fromWho, toWho, msg.as_string())

    def disconnect(self):
        self.server.quit()
        self.server = False


class Imap():

  def __init__(self, config):
    self.body=None
    self.subject=None
    self.returnPath=''
    self.adrFrom=''
    self.host=config['IMAP']['imap.host']
    self.port=config['IMAP']['imap.port']
    self.username=config['IMAP']['imap.username']
    self.password=config['IMAP']['imap.password']

  def fetch(self, debug=False):
    ctx = ssl.SSLContext(ssl.PROTOCOL_TLSv1_2)
    imap_server = imaplib.IMAP4(self.host, self.port)
    imap_server.starttls(ssl_context=ctx)
    imap_server.login(self.username, self.password)
    imap_server.select('INBOX', 1)
#    typ, uids = imap_server.uid("SEARCH", None, "ALL")
    typ, uids = imap_server.search(None, '(UNSEEN)')
    uids = uids[0].split()
    c=0
    for u in uids:
        if debug:
          print("found uid ", u)
        (retcode, data) = imap_server.fetch(u, '(RFC822)')
        if retcode == 'OK':
            c+=1
            raw_email = data[0][1]
            msg = email.message_from_bytes(raw_email)
            tsubject = email.header.decode_header(msg.get('Subject'))
            self.subject=''
            for sub in tsubject:
              if isinstance(sub[0], bytes):
                self.subject += sub[0].decode('UTF-8')
              else:
                self.subject += sub[0]
            self.body = ""
            if msg.is_multipart():
                for payload in msg.get_payload():
                    if payload.get_content_type() == "text/plain":
                        self.body = payload.get_payload()
            else:
                if msg.get_content_type() == "text/plain":
                    self.body = msg.get_payload()
            self.returnPath=msg.get('Return-Path')
            self.adrFrom=msg.get('From')
            if debug:
              print('*** Subject')
              print(self.subject)
#              print('*** Body')
#              print(self.body)
              print('*** Type')
              print(msg.get_content_type())
              print('*** From')
              print(self.adrFrom)
              print('*** Return-Path')
              print(self.returnPath)
              print('* ITEMS')
              for (h,v) in msg.items():
                pprint(h)
              print('* PARTS')
              for part in msg.walk():
                print(part.get_content_type())
              print('----')
            typ, data = imap_server.store(u, '+FLAGS', '\\Seen')
            assert typ=='OK'
        yield self
    if debug:
      print("count=%s" % len(uids))
      print('filtered=%s' % c)
    imap_server.close()
    imap_server.logout()
    del imap_server

config = Config(False).config
filter_domain = config['mail']['mail.sender.domain']
body = config['mail']['mail.body']
body_reject = config['mail']['mail.body.reject']
subject_prefix = config['mail']['mail.subject.prefix']
sender = SMTP(config)
try:
  sender.connect()
except Exception as e:
  print("Connecting error: %s" % e)

for imap in Imap(config).fetch(False):
  print(imap.returnPath)
  print(imap.subject)
  print('----')
  try:
    domain = imap.returnPath.split("@")[1]
    if domain[-1:]=='>':
      domain=domain[:-1]
    print(domain)
    if domain==filter_domain:
      sender.body=body
      sender.sendMail(imap.returnPath, subject_prefix + imap.subject)
    else:
      sender.body = body_reject
      sender.sendMail(imap.returnPath, "Re: " + imap.subject)
  except Exception as e:
    print("Sending error: %s" % e)

sender.disconnect()
