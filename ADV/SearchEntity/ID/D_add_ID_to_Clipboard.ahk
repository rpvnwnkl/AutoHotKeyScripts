
;Context: You are searching IDs in advance and want to add them to Clipboard
;	Your source is a spreadsheet with IDs in a column
;	the script copies ID and tabs over to advance 
;	any open window in advance is closed, and then an entity search
;	window is opened
;	Currently this script deletes two invisible characters copied from 
;	The excel file. I haven't found another way to get around it, yet
;Instructions: have advance and excel open, click advance, then excel, so they
;	are the last two programs you've accessed. select the cell of the first 
;	ID you'd like to search, and press Windows Key plus D
;	
#A::
Reload
return

#D::
;MsgBox D started
SetTitleMatchMode, 2

;Loop 
;{
;MsgBox Loop started

WinWait, Excel, 
IfWinNotActive, Excel, , Excel, 
WinActivate, Excel, ,WinWaitActive, Excel 

;capture first Id
clipboard = ;
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;

;StringLen, clipLength, clipboard
;IfLess, StringLen, 5
;	{
;	MsgBox Caught a break
;	break
;	}

;clean up from excel
entityID = %clipboard%
StringTrimRight, entityID, entityID, 2 ; trim two char from end of string


;reset for next row
Sleep, 100
clipboard = ;
Sleep, 100
Send, {DOWN}

;move to Advance
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 
Sleep, 100

;open Entity search window
Send, {SHIFTDOWN}{F4}{SHIFTUP}
clipboard = %entityID%
ClipWait ;
;MsgBox Clipboard is:`n`n%clipboard%
;Paste id into first text box
Send, ^v
Sleep, 100
Send, {ENTER}
Sleep, 100

;Add to clipboard
Sleep, 150
Send, {F12}

;close current windows, twice for safety
Send, {CTRLDOWN}{F4}{CTRLUP}
Send, {CTRLDOWN}{F4}{CTRLUP}
Sleep, 100

;Send, {ALTDOWN}{TAB}{ALTUP}
WinWait, Excel, 
IfWinNotActive, Excel, , Excel, 
WinActivate, Excel, 
entityID = ;

;}
return