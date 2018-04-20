#f::
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 
MouseClick, left,  303,  126
Sleep, 100
Send, {F5}
WinWait, Go To ..., 
IfWinNotActive, Go To ..., , WinActivate, Go To ..., 
WinWaitActive, Go To ..., 
Send, mailc{ENTER}
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 
Send, {TAB}{DOWN}{DOWN}{TAB}09/12/2017{TAB}{TAB}{TAB}{DOWN}{TAB}{DOWN}{TAB}{DOWN}{TAB}{DOWN}{TAB}{DOWN}{TAB}{DOWN}{TAB}{DOWN}{TAB}{DOWN}{TAB}{TAB}{TAB}{SHIFTDOWN}p{SHIFTUP}er{SPACE}{SHIFTDOWN}t{SHIFTUP}elefund{SPACE}{SHIFTDOWN}cr{SHIFTUP}{SHIFTDOWN}l{SHIFTUP}{F8}{CTRLDOWN}{F4}{F4}{CTRLUP}{SHIFTDOWN}{F4}{SHIFTUP}
return
