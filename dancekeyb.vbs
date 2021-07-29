Set wshShell =wscript.CreateObject("WScript.Shell")
do
wscript.sleep 89
wshshell.sendkeys "{NUMLOCK}"
wscript.sleep 92
wshshell.sendkeys "{CAPSLOCK}"
wscript.sleep 110
wshshell.sendkeys "{SCROLLLOCK}"
loop