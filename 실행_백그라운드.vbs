Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
' pythonw = 콘솔 창 없이 실행 (종료는 작업 관리자에서 pythonw 프로세스 종료)
WshShell.Run "pythonw app.py", 0, False
WScript.Sleep 2500
WshShell.Run "http://localhost:5000", 1, False
