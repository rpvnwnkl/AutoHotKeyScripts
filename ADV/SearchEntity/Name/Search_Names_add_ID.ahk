
;Context: You are using first and last names in advance using a spreadsheet as source

;Instructions: have advance and excel open
;               select the cell of the first last name you'd like to search
;               press Windows Key plus G


#G::
;MsgBox F started
SetTitleMatchMode, 2

WinWait, Excel, 
IfWinNotActive, Excel, , WinActivate, Excel, 
WinWaitActive, Excel,

;capture last name
clipboard = ;
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;

;clean up from excel
lastName = %clipboard%
StringTrimRight, lastName, lastName, 2 ; trim two char from end of string

;move lft, Capture first name
clipboard = ;

Send, {LEFT}
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;

;clean up from excel
firstName = %clipboard%
StringTrimRight, firstName, firstName, 2 ; trim two char from end of string

;Empty clipboard
clipboard = ;

;move to Advance
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 
Sleep, 100

;pull up entity search window, tab to last name
Send, {SHIFTDOWN}{F4}{SHIFTUP}{TAB}{TAB}{TAB}
clipboard = lastName
ClipWait ;
;MsgBox Clipboard is:`n`n%clipboard%
Sleep, 50

;paste last name in
Send, %lastName%
Sleep, 50

;empty clipboard, add first name to it
clipboard = ;
Sleep, 50
clipboard = %firstName%
ClipWait ;
;MsgBox Clipboard is:`n`n%clipboard%
Sleep, 50

;past first name in
Send, {Tab}^v
Sleep, 50

;submit search
Send, {ENTER}

;send Y, just in case last name is too short
;Send, {Y}
Sleep, 50
Send, {ENTER}{ENTER}


return