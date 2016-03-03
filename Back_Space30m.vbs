'Cool you figured out how to open hidden files
Set wshShell =wscript.CreateObject("WScript.Shell") 

'Wait 5 mins
wscript.sleep 300000
do while num < 360
	num = num + 1
	wscript.sleep 5000
	'toggle backspace
	wshshell.sendkeys "{bs}" 
loop
'Loop backspace prank every 5 sec 360 times - 30 mins

wscript.Echo "A friend has just pranked you via PrankSticks"
'Follow me on twitter @TheRipeLime