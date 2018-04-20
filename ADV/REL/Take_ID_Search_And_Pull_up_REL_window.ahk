
#R::

SetTitleMatchMode, 2

WinWait, Excel, 
IfWinNotActive, Excel, , Excel, 
WinActivate, Excel, ,WinWaitActive, Excel 

;capture first Id
clipboard = ;
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;

;clean up from excel
advID = %clipboard%
StringTrimRight, advID, advID, 2 ; trim two char from end of string

;move to Advance
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 
Sleep, 100

;open Entity search window
Send, {SHIFTDOWN}{F4}{SHIFTUP}
clipboard = %advID%
ClipWait ;
;MsgBox Clipboard is:`n`n%clipboard%

;Paste id into first text box
Send, ^v
Sleep, 100
Send, {ENTER}
Sleep, 25

;pull up REL window
Send, {F5}
WinWait, Go To ..., 
IfWinNotActive, Go To ..., , WinActivate, Go To ..., 
WinWaitActive, Go To ..., 
Send, rel{ENTER}

return