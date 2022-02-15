Dim a
Do
Do
a=InputBox("Please set the timer(seconds).","Little Egg Timer",1)
Loop Until IsNumeric(a)
Loop Until a>0
Wscript.sleep 1000*a
CreateObject("SAPI.SpVoice").Speak"Time is up!"
MsgBox"Time is up!",,"Little Egg Timer"