

#g::
SetTitleMatchMode, 2
clipboard = ;
WinWait, Excel, 
IfWinNotActive, Excel, , Excel, 
WinWaitActive, Excel,
Sleep, 100
MouseClick, left,  352,  209
MouseClick, left,  352,  209
Sleep, 100
Send, {CTRLDOWN}c{CTRLUP}
Send, {DOWN}
Send, {ALTDOWN}{TAB}{ALTUP} 
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 
Sleep, 100
Send, {CTRLDOWN}{F4}{CTRLUP}{SHIFTDOWN}{F4}{SHIFTUP}{TAB}{TAB}{TAB}
Sleep, 100
Send, {CTRLDOWN}v{CTRLUP}
Sleep, 100
Send, {ENTER}
return