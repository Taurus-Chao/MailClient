from email import encoders
from email.header import Header
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
from tkinter import *
import tkinter.messagebox as messagebox
import smtplib
import poplib

#图形窗口
class Useritfc(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.pack()
        self.createWidgets()

    def createWidgets(self):
        #发送者邮箱
        self.emailLable = Label(self, text='Email:')
        self.emailLable.pack()
        self.emailInput = Entry(self)
        self.emailInput.pack()
        #邮箱密码
        self.passwordLable = Label(self, text='Authorization Key:')
        self.passwordLable.pack()
        self.passwordInput = Entry(self, show='*')
        self.passwordInput.pack()
        #接受者邮箱
        self.recieverLable = Label(self, text='Reciever:')
        self.recieverLable.pack()
        self.recieverInput = Entry(self)
        self.recieverInput.pack()
        #发送smtp
        self.smtpLable = Label(self, text='SMTP/POP:')
        self.smtpLable.pack()
        self.smtpInput = Entry(self)
        self.smtpInput.pack()
        #发送内容
        self.sendtextLable2 = Label(self, text='Subject:')
        self.sendtextLable2.pack()
        self.sendsubjectInput = Entry(self)
        self.sendsubjectInput.pack()
        self.sendtextLable = Label(self, text='Text')
        self.sendtextLable.pack()
        self.sendtextInput = Entry(self)
        self.sendtextInput.pack()
        #确认按钮
        self.submitButton = Button(self, text='SEND', command=self.submit)
        self.submitButton.pack()
        #读取内容
        self.readtextLable = Label(self, text='Number of Mails')
        self.readtextLable.pack()
        self.readNumInput = Entry(self)
        self.readNumInput.pack()
        self.readButton = Button(self, text='READ', command=self.read)
        self.readButton.pack()
    def submit(self):
        s_email = self.emailInput.get()
        s_password = self.passwordInput.get()
        s_reciever = self.recieverInput.get()
        s_smtp = self.smtpInput.get()
        s_sendsubject=self.sendsubjectInput.get()
        s_sendtext = self.sendtextInput.get()
        if s_email and s_password and s_reciever and s_smtp and s_sendtext:
            startsend(s_smtp, s_email, s_password, s_reciever, s_sendtext,s_sendsubject)
            messagebox.showinfo('Message', 'OK!')
            self.sendsubjectInput.delete(0, END)
            self.sendtextInput.delete(0, END)
        else:
            #填表出错弹窗
            messagebox.showinfo('Message', 'Please input all item correctly!')
    def read(self):
        s_email = self.emailInput.get()
        s_password = self.passwordInput.get()
        s_smtp = self.smtpInput.get()
        s_num=self.readNumInput.get()
        if s_email and s_password and s_smtp and  s_num:
            try:
                startread(s_smtp, s_email, s_password,s_num)
            except ValueError as identifier:
                messagebox.showinfo('Message', 'Please input the number correctly!')
            else:
                messagebox.showinfo('Message', 'OK!')
                self.readNumInput.delete(0, END)
        else:
            #填表出错弹窗
            messagebox.showinfo('Message', 'Please input all item correctly!')

def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name, 'utf-8').encode(), addr))


def startsend(smtp, email, pswd, reciever, sendtext,sendsubject=''):
    msg = MIMEText(sendtext, 'plain', 'utf-8')
    msg['From'] = _format_addr('Chao <%s>' % email)
    msg['To'] = _format_addr('Admin <%s>' % reciever)
    msg['Subject'] = Header(sendsubject, 'utf-8').encode()

    server = smtplib.SMTP(smtp, 587) # SMTP协议默认端口是25
    server.starttls()
    server.set_debuglevel(1)
    server.login(email, pswd)
    server.sendmail(email, [reciever], msg.as_string())
    server.quit()

def startread(s_pop, s_email, s_password,s_num):
    try:
        s_num=int(s_num)
    except ValueError as identifier:
        raise 
    else:
        if s_num<=0:
            raise ValueError
        server = poplib.POP3(s_pop)
        server.set_debuglevel(1)
        print(server.getwelcome().decode('utf-8'))
        # 身份认证:
        server.user(s_email)
        server.pass_(s_password)
        # stat()返回邮件数量和占用空间:
        print('Messages: %s. Size: %s' % server.stat())
        # list()返回所有邮件的编号:
        resp, mails, octets = server.list()
        print(mails)
        # 获取邮件, 注意索引号从1开始:

        count=len(mails)
        s=''
        for index in range(count,count-s_num,-1):
            resp, lines, octets = server.retr(index)
            msg_content = b'\r\n'.join(lines).decode('utf-8')
            msg = Parser().parsestr(msg_content)
            #s=print_info(msg)
            #messagebox.showinfo('Content', s)
            s=s+'\n'+'Mail_Num %i'%(count-index+1)+'\n'+print_info(msg)
        messagebox.showinfo('Content', s)
        server.quit()


def print_info(msg, indent=0,resstr=''):
    if indent == 0:
        for header in ['From', 'To', 'Subject']:
            value = msg.get(header, '')
            if value:
                if header=='Subject':
                    value = decode_str(value)
                else:
                    hdr, addr = parseaddr(value)
                    name = decode_str(hdr)
                    value = u'%s <%s>' % (name, addr)
            #print('%s%s: %s' % ('  ' * indent, header, value))
            resstr=resstr+'\n'+'%s%s: %s' % ('  ' * indent, header, value)
    if (msg.is_multipart()):
        parts = msg.get_payload()
        for n, part in enumerate(parts):
            #print('%spart %s' % ('  ' * indent, n))
            resstr=resstr+'\n'+'%spart %s' % ('  ' * indent, n)
            #print('%s--------------------' % ('  ' * indent))
            resstr=resstr+'\n'+'%s--------------------' % ('  ' * indent)
            print_info(part, indent + 1,resstr)
    else:
        content_type = msg.get_content_type()
        if content_type=='text/plain' or content_type=='text/html':
            content = msg.get_payload(decode=True)
            charset = guess_charset(msg)
            if charset:
                content = content.decode(charset)
            #print('%sText: %s' % ('  ' * indent, content + '...'))
            resstr=resstr+'\n'+'%sText: %s' % ('  ' * indent, content + '...')
        else:
            #print('%sAttachment: %s' % ('  ' * indent, content_type))
            resstr=resstr+'\n'+'%sAttachment: %s' % ('  ' * indent, content_type)
    return resstr

def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value

def guess_charset(msg):
    charset = msg.get_charset()
    if charset is None:
        content_type = msg.get('Content-Type', '').lower()
        pos = content_type.find('charset=')
        if pos >= 0:
            charset = content_type[pos + 8:].strip()
    return charset
        
#启动窗口程序
if __name__=='__main__':
    app = Useritfc()
    app.master.title('SMTP email sender')
    app.mainloop()