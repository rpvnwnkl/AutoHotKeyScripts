
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

#b::
MsgBox B started
SetTitleMatchMode, 2

Loop 
{
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
parentID = %clipboard%
StringTrimRight, parentID, parentID, 2 ; trim two char from end of string

;MsgBox copied parent ID: %parentID%xxxxx

;capture studentID
clipboard = ;
Send, {LEFT}
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;

;clean up from excel
studentID = %clipboard%
StringTrimRight, studentID, studentID, 2 ; trim two char from end of string

;MsgBox Copied studet ID: %studentID%xxxxx

;reset for next row
Sleep, 100
clipboard = ;
Sleep, 100
Send, {RIGHT}{DOWN}

;move to Advance
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 
Sleep, 100

;open Entity search window
Send, {SHIFTDOWN}{F4}{SHIFTUP}
clipboard = %parentID%
ClipWait ;
;MsgBox Clipboard is:`n`n%clipboard%
;Paste id into first text box
Send, ^v
Sleep, 100
Send, {ENTER}
Sleep, 100

;Add child relationship
Send, {F5}
WinWait, Go To ..., 
IfWinNotActive, Go To ..., , WinActivate, Go To ..., 
WinWaitActive, Go To ..., 
Send, CHILD{ENTER}
Sleep, 100
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 

;Select details of relationship and put in ID number
Send, {F6}ch{TAB}
clipboard = %studentID%
ClipWait ;
Send, ^v
Sleep, 100
Send, {TAB}
MsgBox Add child?
Send, {F8}

;close current windows, twice for safety
Send, {CTRLDOWN}{F4}{CTRLUP}
Send, {CTRLDOWN}{F4}{CTRLUP}
Sleep, 100

;Send, {ALTDOWN}{TAB}{ALTUP}
WinWait, Excel, 
IfWinNotActive, Excel, , Excel, 
WinActivate, Excel, 
parentID = ;
studentID = ;
}
return