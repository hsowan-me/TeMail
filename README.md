# TeMail 特邮

基于 `Tkinter` 开发的邮件推送桌面应用, 支持 `Mac/ Windows`.

## 开发

```bash
# 虚拟环境
# pip install -U pipenv # Add to PATH
pipenv install --python 3.7

# 运行
./start.sh # ./start.bat

```

## 打包

```bash
# Mac
./build2app.sh

# Windows
./build2exe.bat

```

```yaml
mail:
  name: hsowan # 用户名
  user: hsowan.me@gmail.com # 发件人邮箱
  pass: qwerasdf # 密码/授权码
  port: 465 # 端口
  ssl: true # 加密
  delay: 3 # 发送延时, 单位秒

```

## 参考文档

* [py2app](https://py2app.readthedocs.io/en/latest/index.html)
* [pyinstaller](http://www.pyinstaller.org/)
* [Python Tkinter教程（GUI图形界面开发教程）](http://c.biancheng.net/python/tkinter/)