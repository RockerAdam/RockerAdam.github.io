Dim a,b,c
MsgBox"Click to start.",,"Little Stopwatch"
a=Now()
MsgBox"Click to end.",,"Little Stopwatch"
b=Now()
c=DateDiff("s",a,b)
MsgBox c&" second(s).",,"Little Stopwatch"