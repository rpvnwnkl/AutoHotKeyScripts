
#u::
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 

;From the parent rec window

;first type PA
Sleep, 100
Send, pa{TAB}

;Then type 2018
Sleep, 100
Send, 2018{TAB}

;Now FN for friedman, then save with f8
Sleep, 100
Send, fn{F8}

;Not sure what this is
Sleep, 100
Send, {ALTDOWN}
Sleep, 50
Send, {ALTUP}
Sleep, 50
Send, b
Sleep, 50
Send, n
return