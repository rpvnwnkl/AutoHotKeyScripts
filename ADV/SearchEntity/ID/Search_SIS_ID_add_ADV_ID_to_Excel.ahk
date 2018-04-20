
;Context: You have SIS ID and are looking to put Advance ID beside it
;	Your source is a spreadsheet with SIS ID col next to empty Col
;	the script copies SIS ID and tabs over to advance 
;	an entity search is run with alt id box checked
;	record is added to clipboard, clipboard is pulled up, and then the id
;	is copied and then pasted into excel
;	Currently this script deletes two invisible characters copied from 
;	The excel file. I haven't found another way to get around it, yet
;Instructions: have advance and excel open, click advance, then excel, so they
;	are the last two programs you've accessed. Make sure the clipboard 
;	is open and empty in ADvance. select the cell of the first 
;	ID you'd like to search, and press Windows Key plus F
;	

#F::
;MsgBox F started
SetTitleMatchMode, 2


WinWait, Excel, 
IfWinNotActive, Excel, , Excel, 
WinActivate, Excel, ,WinWaitActive, Excel 

;capture first Id
clipboard = ;
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;



;clean up from excel
sisID = %clipboard%
StringTrimRight, sisID, sisID, 2 ; trim two char from end of string


;Move to next col for next row
Sleep, 100
clipboard = ;
Sleep, 100
Send, {RIGHT}

;move to Advance
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 
Sleep, 100

;open Entity search window
Send, {SHIFTDOWN}{F4}{SHIFTUP}
clipboard = %sisID%
ClipWait ;
;MsgBox Clipboard is:`n`n%clipboard%

;Paste id into first text box
Send, ^v
;tab and press space to check alt ID box
Send, {TAB}{SPACE}
Sleep, 100
Send, {ENTER}
Sleep, 100

;Add to clipboard
Send, {F12}

;pull up clipboard
Send, {F4}

;capture SIS Id
clipboard = ;
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;

entityID = %clipboard%

;clear clipboard
Send, {ALTDOWN}{ALTUP}ed{ENTER}
Sleep, 25

;switch to entity window
Send, {CTRLDOWN}{F6}{CTRLUP}
Sleep, 25

;close entity window
Send, {CTRLDOWN}{F4}{CTRLUP}

;Move back to Excel
WinWait, Excel, 
IfWinNotActive, Excel, , Excel, 
WinActivate, Excel, 

;should have right cell selected, so paste
clipboard =;
clipboard = %entityID%
ClipWait ;
Send, ^v

;arrow to next cell
Send, {DOWN}{LEFT}

return