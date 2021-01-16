# Date    2020/3/20
# Code 'UTF-8'
# Author lbw

import re
import pickle
import smtplib
import tkinter as tk
import tkinter.messagebox
from email.header import Header
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr
from tkinter import filedialog


class App:
    isUploadFile = 0  # 判断是否上传附件的全局变量
    fileNames = None  # 上传附件的文件名

    def __init__(self, window):
        self.window = window
        self.window.title('Python 邮件客户端(by lbw)')
        # 设定窗口的大小(长 * 宽)
        self.window.geometry('765x605')
        # x，y轴不可改变
        self.window.resizable(0, 0)

        # 创建菜单栏
        self.menubar = tk.Menu(self.window)
        # 创建文件菜单默认不下拉
        self.filemenu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label='文件', menu=self.filemenu)
        self.filemenu.add_command(label='配置', command=self.config)  # 在文件小菜单
        self.filemenu.add_separator()  # 添加一条分隔线
        self.filemenu.add_command(label='退出', command=self.window.quit)  # 用tkinter里面自带的quit()函数
        self.window.config(menu=self.menubar)  # 配置让菜单栏menubar显示出来

        tk.Label(self.window, text="收件人:").place(x=10, y=25)
        tk.Label(self.window, text="主  题:").place(x=10, y=65)
        # 输入框
        self.var_addressee = tk.StringVar()
        self.entry_addressee = tk.Entry(self.window, textvariable=self.var_addressee, width=60)
        self.entry_addressee.place(x=80, y=25)
        self.var_subject = tk.StringVar()
        self.entry_subject = tk.Entry(self.window, textvariable=self.var_subject, width=60)
        self.entry_subject.place(x=80, y=65)

        # 添加附件部分
        self.file_button = tk.Button(self.window, text='添加附件', command=self.upload_file)
        self.file_button.place(x=15, y=105)
        self.var_filename = tk.StringVar()
        tk.Entry(self.window, textvariable=self.var_filename, state='readonly', fg='red', width='50').place(x=75, y=110)

        # 邮件正文输入框
        self.entry_body = tk.Text(self.window, width=104)
        self.entry_body.place(x=15, y=165)

        # 再下层显示配置文件中的发件人信息
        tk.Label(self.window, text="发件人:").place(x=10, y=525)
        self.var_sender = tk.StringVar()  # 存放发件人地址
        tk.Label(self.window, text="发件人信息请在[文件]->[配置]中设置！", fg='red').place(x=220, y=525)

        # 发送和重置button
        self.b = tk.Button(self.window, text="发送", command=self.send, width=10, height=1)  # 调用send()
        self.b.place(x=10, y=560)
        self.b1 = tk.Button(self.window, text="取消", command=self.clear, width=10, height=1)
        self.b1.place(x=140, y=560)
        self.show_sender()

    def show_sender(self):
        """这个方法会将配置文件的信息显示在下层发件人信息处"""
        try:
            with open('Appconfig', 'rb') as conf_file:
                data = pickle.load(conf_file)
            uid = data['UID']
            self.var_sender.set(uid)
            tk.Label(self.window, textvariable=self.var_sender).place(x=70, y=525)
        except FileNotFoundError:
            data = {'UID': '', 'AUTHCODE': '', 'SERVER': '', 'PORT': ''}
            with open('Appconfig', 'wb') as conf_file:
                pickle.dump(data, conf_file)

    def save(self):
        """定义配置文件时的保存方法"""
        uid = self.var_uid.get()
        authcode = self.var_authcode.get()
        server = self.var_server.get()
        port = self.var_port.get()
        data = {'UID': uid, 'AUTHCODE': authcode, 'SERVER': server, 'PORT': port}
        print(data)
        # 打开数据文并从头覆盖写入，如果没有config文件则自动创建并写入
        if (uid == '') or (authcode == '') or (server == '') or (port == ''):
            tk.messagebox.showinfo("提示", "请正确输入,或使用默认值!")
        else:
            try:
                with open('Appconfig', 'wb') as conf_file:
                    pickle.dump(data, conf_file)
                    # conf_file.close()
                tk.messagebox.showinfo("提示", "保存成功!")
                self.show_again()
            except Exception:
                msg = Exception
                tk.messagebox.showerror("Error!", msg)

    def show_again(self):
        """关闭子窗体时重新显示主窗体"""
        self.show_sender()
        self.window.update()
        self.window.deiconify()
        self.window_config.destroy()

    def config(self):
        """ 需要再这个方法中定义顶层窗体，用以配置用户名授权码smtp服务器地址等信息"""
        # 定义子窗体
        self.window_config = tk.Toplevel(self.window)
        self.window_config.geometry('305x225')
        self.window_config.title("配置")
        self.window_config.resizable(0, 0)
        self.window.withdraw()  # 隐藏主窗体
        self.window_config.protocol('WM_DELETE_WINDOW', self.show_again)  # 拦截子窗体关闭事件
        # 用户名输入框
        tk.Label(self.window_config, text="用户名:").place(x=10, y=10)
        self.var_uid = tk.StringVar()
        self.entry_uid = tk.Entry(self.window_config, textvariable=self.var_uid, width=30)
        self.entry_uid.place(x=60, y=10)
        # 授权码输入框
        tk.Label(self.window_config, text="授权码:").place(x=10, y=50)
        self.var_authcode = tk.StringVar()
        self.entry_authcode = tk.Entry(self.window_config, textvariable=self.var_authcode, width=30)
        self.entry_authcode.place(x=60, y=50)
        # SMTP服务器输入框
        tk.Label(self.window_config, text="SMTP服务器:").place(x=10, y=90)
        self.var_server = tk.StringVar()
        self.var_server.set('smtp.163.com')
        self.entry_server = tk.Entry(self.window_config, textvariable=self.var_server, width=24)
        self.entry_server.place(x=100, y=90)
        # 端口号
        tk.Label(self.window_config, text="端口号:").place(x=10, y=120)
        self.var_port = tk.StringVar()
        self.var_port.set('25')
        self.entry_port = tk.Entry(self.window_config, textvariable=self.var_port, width=24)
        self.entry_port.place(x=100, y=120)
        # 写入配置中的服务器地址和端口号如果不存在键值则向文件中追加
        try:
            with open('Appconfig', 'rb') as conf_file:
                data = pickle.load(conf_file)
            if data['SERVER'] != 'smtp.163.com':
                self.var_server.set(data['SERVER'])
        except KeyError:
            data = {'SERVER': 'smtp.163.com', 'PORT': '25'}
            with open('config', 'ab') as cf:
                pickle.dump(data, cf)

        # 保存button 取消button
        tk.Button(self.window_config, text="保存", command=self.save, width=10, height=1).place(x=25, y=160)
        tk.Button(self.window_config, text="取消", command=self.show_again, width=10, height=1).place(x=200, y=160)

    def upload_file(self):
        """打开附件方法"""
        global fileNames
        global isUploadFile
        fileNames = filedialog.askopenfilenames()
        self.var_filename.set(fileNames)
        if len(fileNames):
            self.isUploadFile = 1
        else:
            self.isUploadFile = 0

    def send(self):
        """在这里进行消息的构造和发送"""
        def _format_addr(s):
            # 格式化发件人信息
            name, addr = parseaddr(s)
            return formataddr((Header(name, 'utf-8').encode(), addr))
        msg = "用" + self.var_sender.get() + "这个账户发送邮件?发送时填好主题(若主题为test等可能会被退信导致发送失败)和收件人地址哦!"
        is_send = tk.messagebox.askyesno("Are you sure?", msg)  # 最后确认发送信息
        if is_send:
            # 开始发送
            # 打开并读取配置文件信息
            with open('Appconfig', 'rb') as conf_file:
                data = pickle.load(conf_file)
            print(data)
            uid = data['UID']
            authcode = data['AUTHCODE']
            server = data['SERVER']
            port = data['PORT']

            # 收集输入框信息
            addressee = self.var_addressee.get()
            subject = self.var_subject.get()
            body = self.entry_body.get('0.0', 'end')
            # print(body)
            if addressee == ''or subject == '':
                tk.messagebox.showinfo('注意！','收件人地址和标题为必填！')
            if re.match("^.+\\@(\\[?)[a-zA-Z0-9\\-\\.]+\\.([a-zA-Z]{2,3}|[0-9]{1,3})(\\]?)$", addressee) == None:
                tk.messagebox.showinfo('注意!', '收件人格式错误!')
            else:
                # 生成消息对象并构建邮件头部
                # 三个参数：第一个为文本内容body，第二个 plain 设置文本格式，第三个 utf-8 设置编码
                message = MIMEMultipart()
                message['From'] = _format_addr('Python邮件客户端 <%s>' % uid)  # 因为包含中文需格式化
                message['To'] = Header(addressee)
                message['Subject'] = Header(subject, 'utf-8').encode()
                msg = MIMEText(body.encode('utf-8'), 'html', 'utf-8')
                message.attach(msg)
                print(message)

                # 判断是否需要上传附件
                if self.isUploadFile:
                    for fname in fileNames:
                        with open(fname, 'rb') as f:
                            attachfiles = MIMEApplication(f.read())
                            attachfiles.add_header('Content-Disposition', 'attachment', filename=fname)
                            message.attach(attachfiles)
                try:
                    if int(port) == 25:
                        smtpObj = smtplib.SMTP(server, int(port))
                    else:
                        # 如果为非默认端口，建立加密连接
                        smtpObj = smtplib.SMTP_SSL(server, int(port))
                    smtpObj.login(uid, authcode)
                    smtpObj.sendmail(uid, addressee, message.as_string())
                    tk.messagebox.showinfo("成功！", "发送成功！")
                    smtpObj.close()
                except smtplib.SMTPException as e:
                    tk.messagebox.showerror("发送失败!", e)

    def clear(self):
        """清除输入框中的全部内容"""
        self.entry_body.delete(1.0, 'end')
        self.entry_subject.delete(0, 'end')
        self.entry_addressee.delete(0, 'end')


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
