Set oWMP = CreateObject("WMPlayer.OCX.7")
Set colCDROMs = oWMP.cdromCollection
max=300000
min=10000

wscript.sleep 300000
do while num < 100
randomize 
randNum = (Int((max-min+1)*rnd+min))

x = randNum

num = num + 1
if colCDROMs.Count >= 1 then
For i = 0 to colCDROMs.Count - 1
colCDROMS.Item(i).Eject
Next
For i = 0 to colCDROMs.Count - 1
colCDROMs.Item(i).Eject
Next
End If
wscript.sleep x
loop
wscript.Echo "A friend has just pranked you via PrankSticks"
Set kckst=WScript.CreateObject("WScript.Shell")
kckst.Run "https://www.kickstarter.com/projects/1471256552/pranksticks-prank-smart?ref=home_location"
