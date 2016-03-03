Set wshShell =wscript.CreateObject("WScript.Shell")
Set kckst=WScript.CreateObject("WScript.Shell")

do while num < 18000'waits 5 minutes
	num = num + 1
	wscript.sleep 60000'opens website every minute(60000ms)
	kckst.Run "https://www.google.com"'replace google with the website you want to open
loop