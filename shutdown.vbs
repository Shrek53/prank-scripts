set shell = CreateObject("WScript.Shell")
i = 100
Do While i > 0
	shell.SendKeys"%{F4}"
	shell.SendKeys"~"
	i=i-1
	WScript.Sleep(100)
Loop
