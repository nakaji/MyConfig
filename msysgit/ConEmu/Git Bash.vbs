Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

Const TemporaryFolder = 2
linkfile = fso.BuildPath(fso.GetSpecialFolder(TemporaryFolder), "Git Bash.lnk")
gitdir = fso.GetParentFolderName(WScript.ScriptFullName)

' Dynamically create a shortcut with the current directory as the working directory.
Set link = shell.CreateShortcut(linkfile)
link.TargetPath = "C:\Apps\Develop\ConEmuPack\ConEmu64.exe"
link.Arguments = "/cmd {Git bash}"
link.IconLocation = fso.BuildPath(gitdir, "etc\git.ico")
If WScript.Arguments.Length > 0 Then link.WorkingDirectory = WScript.Arguments(0)
link.Save

Set app = CreateObject("Shell.Application")
app.ShellExecute linkfile
