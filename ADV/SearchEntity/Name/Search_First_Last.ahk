
;Context: You are search first and last names in advance
;	Your source is a spreadsheet with first and last side by side
;	the script copies last, then first and tabs over to advance 
;	any oopen window in advance is closed, and then an entity search
;	window is opened
;	Currently this script deletes two invisible characters copied from 
;	The excel file. I haven't found another way to get around it, yet
;Instructions: have advance and excel open, click advance, then excel, so they
;	are the last two programs you've accessed. select the cell of the first 
;	last name you'd like to search, and press Windows Key plus F
;	alt tab back to excel when the new search is to begin
#f::
SetTitleMatchMode, 2
WinWait, Excel, 
IfWinNotActive, Excel, , Excel, 
WinWaitActive, Excel, 
clipboard = ;
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;
lastName = %clipboard%
realLastName = %lastName%
clipboard = ;
Send, {LEFT}
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;
firstName = %clipboard%
realFirstName = %firstName%
Sleep, 100
clipboard = ;
Sleep, 100
Send, {RIGHT}{DOWN}
;MsgBox Control-C copied the following contents to the clipboard:`n`n%firstName% and %lastName%
Send, {ALTDOWN}{TAB}{ALTUP}
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 
Sleep, 100
Send, {CTRLDOWN}{F4}{CTRLUP}
Sleep, 100
Send, {CTRLDOWN}{F4}{CTRLUP}
Sleep, 100
Send, {SHIFTDOWN}{F4}{SHIFTUP}{TAB}{TAB}{TAB}
clipboard = %realLastName%
ClipWait ;
;MsgBox Clipboard is:`n`n%clipboard%
Send, ^v
Sleep, 100
Send, {BACKSPACE}{BACKSPACE}{TAB}
clipboard = ;
Sleep, 100
clipboard = %realFirstName%
ClipWait ;
;MsgBox Clipboard is:`n`n%clipboard%
Send, {CTRLDOWN}v{CTRLUP}
Sleep, 100
Send, {BACKSPACE}{BACKSPACE}{ENTER}
return