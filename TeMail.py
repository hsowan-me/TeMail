# -*- coding: utf-8 -*-

"""
@author: hsowan <hsowan.me@gmail.com>
@date: 2019/10/19

"""

from tkinter import *
from tkinter import simpledialog, messagebox, filedialog, dialog
from tkinter import ttk
import yaml
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from smtplib import SMTP_SSL, SMTP
from email.utils import parseaddr, formataddr
import time


class SettingsDialog(Toplevel):
    """应用设置对话框"""

    # 定义构造方法
    def __init__(self, parent, title=None):
        Toplevel.__init__(self, parent)
        self.transient(parent)
        # 设置标题
        if title:
            self.title(title)
        self.parent = parent
        self.result = None
        # 创建对话框的主体内容
        frame = Frame(self)

        frame.pack(padx=5, pady=5)
        # 调用init_buttons方法初始化对话框下方的按钮
        self.init_buttons()
        # 设置为模式对话框
        self.grab_set()
        if not self.initial_focus:
            self.initial_focus = self
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
            self.initial_focus.focus_set()
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
    """主应用类"""

    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.pack()
        self.init_widgets()

    def init_widgets(self):
        # 创建menubar，它被放入self.master中
        menubar = Menu(self.master)
        # 添加菜单条
        self.master['menu'] = menubar
        settings_menu = Menu(menubar, tearoff=0)
        # 使用add_cascade方法添加
        menubar.add_cascade(label='设置', menu=settings_menu)
        settings_menu.add_command(label="SMTP", command=self.open_settings)

        # 创建Entry组件
        self.entry = ttk.Entry(self.master, width=44)
        self.entry.pack()

        quit_button = ttk.Button(self, text='quit')

        quit_button['command'] = self.quit
        quit_button.pack({"side": "left"})

        settings_button = ttk.Button(self, text='settings', command=self.open_settings)
        settings_button.pack({"side": "left"})


        # 创建2个按钮，并为之绑定事件处理函数
        ttk.Button(self.master, text='设置', command=self.open_settings).pack(side=LEFT, ipadx=5, ipady=5, padx=10)

    def import_mails(self):
        file = filedialog.askopenfile(title='导入邮箱列表', filetypes=[('文本文件', '*.txt'), ('CSV文件', '*.csv'), ('xlsx文件', '*.xlsx')])
        print(file.read())

    def open_settings(self):
        print('open settings')
        d = SettingsDialog(self.master, title='模式对话框')  # 默认是模式对话框

    def send_mails(self, recipients, subject, content, settings=None):
        """发送邮件

        :param recipients: set, 收件人列表
        :param subject: str, 邮件主题
        :param content: str, 邮件内容
        :param settings: dict, 配置信息
        :return:
        """

        # mail content
        message = MIMEText(content, 'html', 'utf-8')
        # mail recipient
        message['From'] = self._format_addr('%s <%s>' % (settings['mail_name'], settings['mail_user']))
        # mail subject
        message['Subject'] = Header(subject, 'utf-8').encode()

        host_port = (settings['mail_host'], settings['mail_port'])
        smtp = SMTP_SSL(*host_port) if settings['use_ssh'] else SMTP(*host_port)

        try:
            # smtp.set_debuglevel(1)
            smtp.login(settings['mail_user'], settings['mail_pass'])
            success_mails = []
            failed_mails = []

            for recipient in recipients:
                message['To'] = Header(recipient, 'utf-8').encode()
                try:
                    # send mail
                    smtp.sendmail(settings['mail_user'], [recipient], message.as_string())
                    success_mails.append(recipient)
                except smtplib.SMTPException as e:
                    failed_mails.append(dict(to=recipient, error=str(e)))
                # delay time before send the next mail
                time.sleep(settings['mail_delay'])
        except smtplib.SMTPException as e:
            print(e)
            # 跳出对话框显示错误
        finally:
            smtp.quit()

    @staticmethod
    def _format_addr(s):
        """格式化发件人

        Refer: https://www.liaoxuefeng.com/wiki/1016959663602400/1017790702398272
        """
        name, addr = parseaddr(s)
        return formataddr((Header(name, 'utf-8').encode(), addr))

    def center_window(self, width, height, root=None):
        """设置窗口大小并居中

        :param width: int
        :param height: int
        :param root: Tkinter.Tk()
        :return:
        """

        if root is None:
            root = self.master
        # 获取屏幕 宽、高
        ws = root.winfo_screenwidth()
        hs = root.winfo_screenheight()
        # 计算 x, y 位置
        x = (ws / 2) - (width / 2)
        y = (hs / 2) - (height / 2)
        # 设置窗口的大小和位置
        # width x height + x_offset + y_offset
        root.geometry('%dx%d+%d+%d' % (width, height, x, y))


if __name__ == '__main__':
    # 创建Tk对象, Tk代表窗口
    win = Tk()
    # 设置窗口标题
    win.title('特邮')
    # 禁止改变窗口大小
    win.resizable(width=False, height=False)

    app = Application(master=win)
    app.center_window(600, 400)
    # 启动主窗口的消息循环
    app.mainloop()
