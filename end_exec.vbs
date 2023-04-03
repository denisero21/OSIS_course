Set TaskKill = CreateObject("WScript.Shell")
MsgBox "You are saved. Be careful next time.", 64+4096+262144, "It's okay"
TaskKill.Run "TaskKill /f /im wscript.exe", 0
