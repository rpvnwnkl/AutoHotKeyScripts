
#R::
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 
Sleep, 25
Send, {F5}
WinWait, Go To ..., 
IfWinNotActive, Go To ..., , WinActivate, Go To ..., 
WinWaitActive, Go To ..., 
Send, rel{ENTER}
return