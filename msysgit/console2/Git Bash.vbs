Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

Const TemporaryFolder = 2
linkfile = fso.BuildPath(fso.GetSpecialFolder(TemporaryFolder), "Git Bash.lnk")
gitdir = fso.GetParentFolderName(WScript.ScriptFullName)

' Dynamically create a shortcut with the current directory as the working directory.
Set link = shell.CreateShortcut(linkfile)
link.TargetPath = "D:\Apps\Develop\Console2\Console.exe"
link.Arguments = " -t ""Git Bash"" -d """ & WScript.Arguments(0) & """"
link.IconLocation = fso.BuildPath(gitdir, "etc\git.ico")
link.Save

Set app = CreateObject("Shell.Application")
app.ShellExecute linkfile
