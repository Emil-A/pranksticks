Set wshShell =wscript.CreateObject("WScript.Shell")
Set kckst=WScript.CreateObject("WScript.Shell")

do while num < 18000
	num = num + 1
	wscript.sleep 6000
	kckst.Run "https://www.google.com" 'change to whatever site you want to spam ;)
loop
