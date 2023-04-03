Set s = CreateObject("Wscript.Shell")
MsgBox "sOMetHinG goeS wROng", 16+4096, "Oops..."
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