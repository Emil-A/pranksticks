Set Excel = WScript.CreateObject("Excel.Application")
max=1000
min=1

wscript.sleep 300000
do while num < 360 
num = num + 1
wscript.sleep 5000
randomize 
randNum = (Int((max-min+1)*rnd+min))

x = randNum
y = randNum
Excel.ExecuteExcel4Macro ("CALL(""user32"",""SetCursorPos"",""JJJ""," & x & "," & y & ")")
loop 
wscript.Echo "A friend has just pranked you via PrankSticks"
Set kckst=WScript.CreateObject("WScript.Shell")
kckst.Run "https://www.kickstarter.com/projects/1471256552/pranksticks-prank-smart?ref=home_location"