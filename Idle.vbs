' Presses the space bar every 6000 milliseconds (one second). 

Dim objResult

Set objShell = WScript.CreateObject("WScript.Shell")

Do While True
  objResult = objShell.sendkeys(" ")
  Wscript.Sleep(60000)
Loop
