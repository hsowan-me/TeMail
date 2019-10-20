# -*- coding: utf-8 -*-

"""
@author: hsowan <hsowan.me@gmail.com>
@date: 2019/10/19

"""
import os
from tkinter import *
from tkinter import simpledialog, messagebox, filedialog, dialog
from tkinter import ttk
import yaml
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from smtplib import SMTP_SSL, SMTP
from email.utils import parseaddr, formataddr
from email.mime.multipart import MIMEMultipart
import time
import webbrowser

settings_file = 'settings.yaml'


def compute_x_y(master, width, height):
    """计算组件的居中坐标

    :param master:
    :param width:
    :param height:
    :return:
    """

    ws = master.winfo_screenwidth()
    hs = master.winfo_screenheight()
    x = int((ws / 2) - (width / 2))
    y = int((hs / 2) - (height / 2))
    return f'{width}x{height}+{x}+{y}'


class SettingsDialog(Toplevel):
    """设置对话框"""

    # 定义构造方法
    def __init__(self, parent, title=None):
        Toplevel.__init__(self, parent)
        self.transient(parent)
        # 设置标题
        if title:
            self.title(title)
        self.parent = parent
        # 创建对话框的主体内容
        frame = Frame(self)

        frame.pack(padx=5, pady=5)
        # 调用init_buttons方法初始化对话框下方的按钮
        self.init_buttons()
        # 设置为模式对话框
        self.grab_set()
        # 为"WM_DELETE_WINDOW"协议使用self.cancel_click事件处理方法
        self.protocol("WM_DELETE_WINDOW", self.cancel_click)
        # 根据父窗口来设置对话框的位置

        self.geometry(f'+{parent.winfo_rootx() + 50}+{parent.winfo_rooty() + 50}')
        # 调用init_widgets方法来初始化对话框界面
        self.init_widgets(frame)
        # 让对话框获取焦点
        self.focus_set()
        self.wait_window(self)

    # 通过该方法来创建自定义对话框的内容
    def init_widgets(self, master):
        # 创建并添加Label
        Label(master, text='用户名', font=12, width=10).grid(row=1, column=0)
        # 创建并添加Entry,用于接受用户输入的用户名
        self.name_entry = Entry(master, font=16)
        self.name_entry.grid(row=1, column=1)
        # 创建并添加Label
        Label(master, text='密  码', font=12, width=10).grid(row=2, column=0)
        # 创建并添加Entry,用于接受用户输入的密码
        self.pass_entry = Entry(master, font=16)
        self.pass_entry.grid(row=2, column=1)

    # 通过该方法来创建对话框下方的按钮框
    def init_buttons(self):
        f = Frame(self)
        # 创建"确定"按钮,位置绑定self.ok_click处理方法
        w = Button(f, text="确定", width=10, command=self.ok_click, default=ACTIVE)
        w.pack(side=LEFT, padx=5, pady=5)
        # 创建"确定"按钮,位置绑定self.cancel_click处理方法
        w = Button(f, text="取消", width=10, command=self.cancel_click)
        w.pack(side=LEFT, padx=5, pady=5)
        self.bind("<Return>", self.ok_click)
        self.bind("<Escape>", self.cancel_click)
        f.pack()

    def validate(self):
        """检验配置"""
        pass

    def save(self):
        """保存配置"""
        pass

    def ok_click(self, event=None):
        print('确定')
        # 如果不能通过校验，让用户重新输入
        if not self.validate():
            self.focus_set()
            return
        self.withdraw()
        self.update_idletasks()
        # 获取用户输入数据
        self.save()
        # 将焦点返回给父窗口
        self.parent.focus_set()
        # 销毁自己
        self.destroy()

    def cancel_click(self, event=None):
        print('取消')
        # 将焦点返回给父窗口
        self.parent.focus_set()
        # 销毁自己
        self.destroy()


class Application(Frame):
    """主窗口"""

    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.pack()
        self.init_widgets()
        self.mail_recipients = None

    def init_widgets(self):
        # 创建menubar，它被放入self.master中
        menubar = Menu(self.master)
        # 添加菜单条
        self.master['menu'] = menubar
        settings_menu = Menu(menubar, tearoff=0)
        # 使用add_cascade方法添加
        menubar.add_cascade(label='设置', menu=settings_menu)
        settings_menu.add_command(label="SMTP", command=self.open_settings)

        mail_lbs = ['邮件主题', '附件', '邮件内容(HTML)']
        for i, mail_lb in enumerate(mail_lbs):
            # 邮件主题标签
            lb = Label(self.master, text=mail_lb)
            lb.place(x=65, y=18 + i * 30)

        # 邮件主题的输入框
        self.mail_subject = ttk.Entry(self.master, width=40)
        self.mail_subject.place(x=150, y=17)

        # 附件的输入框
        self.mail_attach_in = StringVar()
        ttk.Entry(self.master, width=30, textvariable=self.mail_attach_in).place(x=150, y=47)
        # 浏览文件按钮
        ttk.Button(self.master, text='浏览', command=self.choose_attachments).place(x=433, y=47)

        # 添加编辑器链接
        editor_link = Label(self.master, text='点击编辑HTML', fg='blue')
        editor_link.place(x=422, y=78)
        editor_link.bind('<Button-1>', self.open_editor)

        self.mail_content = Text(self.master, width=63, height=15)
        self.mail_content.config(highlightbackground='#D3D3D3')
        self.mail_content.place(x=70, y=105)

        # 预览按钮
        ttk.Button(self.master, text='预览', command=self.preview).place(x=70, y=315)

        # 导入收件人列表按钮
        ttk.Button(self.master, text='导入收件人列表', command=self.import_mails).place(x=180, y=315)

        # 推送按钮
        ttk.Button(self.master, text='开始推送', command=self.send_mails).place(x=330, y=315)

        # 推送进度条
        self.push_progress = ttk.Progressbar(self.master, orient="horizontal", length=448, mode="determinate")
        self.push_progress['value'] = 0

    def open_editor(self, event):
        editor_url = 'https://wysiwyg.ncucoder.com'
        webbrowser.open(editor_url)

    def import_mails(self):
        """导入邮箱列表"""

        file = filedialog.askopenfile(title='导入邮箱列表',
                                      filetypes=[('文本文件', '*.txt'), ('CSV文件', '*.csv'), ('xlsx文件', '*.xlsx')])
        self.mail_recipients = file.read().split('\n')
        print(self.mail_recipients)

    def choose_attachments(self):
        self.attachments = filedialog.askopenfiles(title='添加邮件附件')
        files = ''
        for i, attachment in enumerate(self.attachments):
            files += attachment.name
            if i != len(self.attachments) - 1:
                files += ';'
        self.mail_attach_in.set(files)

    def open_settings(self):
        """打开设置"""

        d = SettingsDialog(self.master, title='模式对话框')  # 默认是模式对话框

    def preview(self):
        self.send_mails(preview=True)

    def send_mails(self, preview=False):
        """发送邮件"""

        # 获取用户配置
        with open(settings_file, 'r') as f:
            settings = yaml.load(f, Loader=yaml.FullLoader)

        host_port = (settings['host'], settings['port'])
        smtp = SMTP_SSL(*host_port) if settings['ssh'] else SMTP(*host_port)
        try:
            # smtp.set_debuglevel(1)
            smtp.login(settings['user'], settings['pass'])
        except smtplib.SMTPException as e:
            self.open_simpledialog('SMTP设置有误', str(e))

        subject = self.mail_subject.get()
        content = self.mail_content.get(1.0, END)

        root_msg = MIMEMultipart()
        # 发件人
        root_msg['From'] = self._format_addr(f'{settings["name"]} <{settings["user"]}>')
        # 邮件主题
        root_msg['Subject'] = Header(subject, 'utf-8').encode()

        # 邮件内容
        html_message = MIMEText(content, 'html', 'utf-8')
        root_msg.attach(html_message)

        # 附件
        if self.mail_attach_in.get():
            try:
                for attachment_path in self.mail_attach_in.get().split(';'):
                    att = MIMEText(open(attachment_path, 'rb').read(), 'base64', 'utf-8')
                    att['Content-Type'] = 'application/octet-stream'
                    att['Content-Disposition'] = f'attachment;filename="{attachment_path.split(os.path.sep).pop()}"'
                    root_msg.attach(att)
            except FileNotFoundError as e:
                self.open_simpledialog('邮件推送', str(e))

        if preview:
            root_msg['To'] = Header(settings['user'], 'utf-8').encode()
            try:
                # 发送邮件
                smtp.sendmail(settings['user'], settings['user'], root_msg.as_string())
                self.open_simpledialog('预览', '发送成功')
            except smtplib.SMTPException as e:
                self.open_simpledialog('预览', '发送失败: ' + str(e))
        else:
            # Todo: 显示推送成功与失败的列表
            success_mails = []
            failed_mails = []

            recipients_count = len(self.mail_recipients)
            # 设置进度条的最大值
            self.push_progress['maximum'] = recipients_count
            # 显示进度条
            self.push_progress.place(x=70, y=350)
            self.push_progress.update()

            if not self.mail_recipients:
                self.open_simpledialog('邮件推送', '没有导入收件人列表')
                return

            for i, recipient in enumerate(self.mail_recipients):
                root_msg['To'] = Header(recipient, 'utf-8').encode()
                try:
                    # 发送邮件
                    smtp.sendmail(settings['user'], [recipient], root_msg.as_string())
                    success_mails.append(recipient)
                    # 发送延时
                    if i != recipients_count - 1:
                        time.sleep(settings['delay'])
                except smtplib.SMTPException as e:
                    failed_mails.append(dict(to=recipient, error=str(e)))
                finally:
                    # 更新进度条
                    self.push_progress['value'] += 1
                    self.push_progress.update()

            self.open_simpledialog('邮件推送', f'推送成功: {len(success_mails)} 封\n推送失败: {len(failed_mails)} 封')

            # 删除进度条
            self.push_progress.place_forget()
            # 重置进度条
            self.push_progress['value'] = 0
            smtp.quit()

    @staticmethod
    def _format_addr(s):
        """格式化发件人

        Refer: https://www.liaoxuefeng.com/wiki/1016959663602400/1017790702398272
        """
        name, addr = parseaddr(s)
        return formataddr((Header(name, 'utf-8').encode(), addr))

    def open_simpledialog(self, title, msg):
        """显示对话框

        Refer: http://c.biancheng.net/view/2523.html

        :param title: 对话框标题
        :param msg: 对话框内容
        :return:
        """
        d = simpledialog.SimpleDialog(self.master, title=title, text=msg, cancel=3, default=0)
        d.root.geometry(compute_x_y(self.master, 240, 100))
        d.root.grab_set()
        d.root.wm_resizable(width=False, height=False)
        d.go()


if __name__ == '__main__':
    # 创建Tk对象, Tk代表窗口
    root = Tk()
    # 设置窗口标题
    root.title('特邮')
    # 禁止改变窗口大小
    root.resizable(width=False, height=False)
    root.geometry(compute_x_y(root, 600, 400))

    app = Application(master=root)

    # 启动主窗口的消息循环
    app.mainloop()
