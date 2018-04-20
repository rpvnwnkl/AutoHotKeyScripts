#t::
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 
MouseClick, left,  487,  125
Sleep, 100
Send, {F5}
WinWait, Go To ..., 
IfWinNotActive, Go To ..., , WinActivate, Go To ..., 
WinWaitActive, Go To ..., 
Send, mailc{ENTER}
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 
Send, {TAB}{DOWN}{DOWN}{DOWN}
MouseClick, left,  680,  185
Sleep, 100
Send, 09/12/2017
MouseClick, left,  353,  386
Sleep, 100
Send, {SHIFTDOWN}p{SHIFTUP}er{SPACE}{SHIFTDOWN}t{SHIFTUP}elefund{SPACE}{SHIFTDOWN}c{SHIFTUP}{SHIFTDOWN}r{SHIFTUP}{SHIFTDOWN}l{SHIFTUP}
MouseClick, left,  55,  258
Sleep, 100
Send, ^{F4}
Send, ^{F4}
Send, {SHIFTDOWN}{F4}{SHIFTUP}
return