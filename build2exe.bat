:: 关闭回显
@echo off

:: 删除打包目录
rd/s/q build dist

:: 打包
:: --clean -y
pipenv run pyinstaller -F -w TeMail.py

:: 窗口等待
:: @pause