Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

Const TemporaryFolder = 2
linkfile = fso.BuildPath(fso.GetSpecialFolder(TemporaryFolder), "Git Bash.lnk")

' Dynamically create a shortcut with the current directory as the working directory.
Set link = shell.CreateShortcut(linkfile)
link.TargetPath = "C:\Program Files (x86)\Git\bin\mintty.exe"
link.Arguments = "-"
link.IconLocation = fso.BuildPath(gitdir, "etc\git.ico")
If WScript.Arguments.Length > 0 Then link.WorkingDirectory = WScript.Arguments(0)
link.Save

Set app = CreateObject("Shell.Application")
app.ShellExecute linkfile
