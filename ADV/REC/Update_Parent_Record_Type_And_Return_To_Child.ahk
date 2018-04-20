
#u::
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 
;MouseClick, left,  196,  124
Sleep, 100
Send, pa{TAB}
Sleep, 100
Send, 2018{TAB}
Sleep, 100
Send, fn{F8}
Sleep, 100
Send, ^{F4}^{F4}
Send, {ALTDOWN}
Sleep, 50
Send, {ALTUP}
Sleep, 50
Send, b
Sleep, 50
Send, n
return