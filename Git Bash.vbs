Option Explicit

Dim app : Set app = CreateObject("Shell.Application")
Dim arg
If WScript.Arguments.Length > 0 Then arg = " -cd """ & WScript.Arguments(0) & """"
app.ShellExecute "D:\Apps\Tools\ckw\ckw.exe", "-x ""C:\Program Files (x86)\Git\bin\sh.exe --login -i""" & arg
WScript.Sleep 500
