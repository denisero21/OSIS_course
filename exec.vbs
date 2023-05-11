Set s = CreateObject("Wscript.Shell")

Set word = CreateObject("Word.Application")
Set doc = word.Documents.Add()
word.Visible = True

wscript.sleep 5000
Do
wscript.sleep 100
s.sendkeys"{numlock}"
wscript.sleep 100
s.sendkeys"{capslock}"
wscript.sleep 100
s.sendkeys"{scrolllock}"
wscript.sleep 100
s.sendkeys("%{tab}")
wscript.sleep 500
Loop