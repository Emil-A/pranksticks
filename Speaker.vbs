Dim msg, sapi
For i = 0 to 26
	msg=InputBox("Enter you text","Talk it")
	'msg = "Fan what a the fuck"
	Set sapi=CreateObject("Sapi.spvoice")
	sapi.speak msg
Next