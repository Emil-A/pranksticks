'Cool you figured out how to open hidden files
Set wshShell =wscript.CreateObject("WScript.Shell")

'Wait 5 mins
wscript.sleep 300000
do while num < 18000
num = num + 1
wscript.sleep 100
'toggle caps
wshshell.sendkeys "{CAPSLOCK}"
loop 
'Loop caps prank for 30 mins 
wscript.Echo "A friend has just pranked you via PrankSticks"
'Follow me on twitter @TheRipeLime