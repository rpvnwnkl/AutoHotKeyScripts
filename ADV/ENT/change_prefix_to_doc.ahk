#d::
SetTitleMatchMode, 1
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 
MouseClick, left,  251,  129
Sleep, 100
Send, {DEL}{TAB}{DEL}{TAB}{DEL}
Send, {TAB}{TAB}{TAB}{TAB}
Send, {SHIFTDOWN}d{SHIFTUP}r.
Send, {F8}{ENTER}{ENTER}
;turn vcr mode on/off with the next two lines
MouseClick, left,  33,  466
Sleep, 100
return