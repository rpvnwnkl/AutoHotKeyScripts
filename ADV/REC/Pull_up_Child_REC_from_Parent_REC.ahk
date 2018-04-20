/*
From the parent's REC window, this will open the child dialog box, click the first child, and then open their record window
*/

#y::
SetTitleMatchMode, 2
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 

;Pull up GoTo window
Send, {F5}
WinWait, Go To ..., 
IfWinNotActive, Go To ..., , WinActivate, Go To ..., 
WinWaitActive, Go To ..., 

;Enter CHILD for child window
Send, child{ENTER}
Sleep, 100

;Click first child
;MouseClick, left,  102,  294

;Get first child ID
Send, {TAB}

;Copy ID to clipboard
Send, ^c

;Close child window
Send, {CTRLDOWN}{F4}{CTRLUP}

;Pull up the GoTo window
Send, {F5}
WinWait, Go To ..., 
IfWinNotActive, Go To ..., , WinActivate, Go To ..., 
WinWaitActive, Go To ..., 

;Enter REC for record
Send, REC
Send, {TAB}
;Paste in child id
Send, ^v{ENTER}

return