'Cool you figured out how to open hidden files
Set Excel = WScript.CreateObject("Excel.Application")
maxX=2050
minX=1

maxY=750
minY=1

'Wait 5 mins
wscript.sleep 300000
do while num < 360 
	num = num + 1
	wscript.sleep 5000

	randomize 
	randNumX = (Int((maxX-minX+1)*rnd+minX))
	x = randNumX

	randomize 
	randNumY = (Int((maxY-minX+1)*rnd+minX))
	y = randNumY
	Excel.ExecuteExcel4Macro ("CALL(""user32"",""SetCursorPos"",""JJJ""," & x & "," & y & ")")
	'Launch mouse to new position
loop 
'Loop every 5 sec 360 times - 1 day

wscript.Echo "A friend has just pranked you via PrankSticks"
'Follow me on twitter @TheRipeLime