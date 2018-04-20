/*This script is designed to take receipt numbers from an excel spreadsheet
	and then update the source code for that receipt to OSH in Advance.
*/

; This script starts with windows Key + O


#o:: 
SetTitleMatchMode, 2

WinWait, Excel, 
IfWinNotActive, Excel, , Excel, 
WinActivate, Excel, ,WinWaitActive, Excel 

;capture receipt #
clipboard = ;
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;

;clean up from excel
receiptNum = %clipboard%
StringTrimRight, receiptNum, receiptNum, 2 ; trim two char from end of string

;MsgBox You've copied reciept number: %receiptNum%

;reset for next row
Sleep, 100
clipboard = ;
Sleep, 100
Send, {DOWN}

;move to Advance

WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, , 
;WinWaitActive, Production Database, , 
Sleep, 100

;open gift lookup window
Send, {ALTDOWN}{ALTUP}
Send, a
Send, g
Send, g

WinWait, Primary Gift, 
IfWinNotActive, Primary Gift, , WinActivate, Primary Gift, 
WinWaitActive, Primary Gift,
clipboard = %receiptNum%
Sleep, 100

Send, ^v{ENTER}{ENTER}

WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 
Send, {TAB}{TAB}{TAB}{TAB}{TAB}{TAB}
Sleep, 100
Send, osh
Sleep, 100
MsgBox Is this OK?
Send, {F8}
Sleep, 100
Send, {CTRLDOWN}{F4}{CTRLUP}
Sleep, 100
Send, {CTRLDOWN}{F4}{CTRLUP}
Sleep, 100

WinWait, Excel, 
IfWinNotActive, Excel, , Excel, 
WinActivate, Excel, 
receiptNum = ;
clipboard = ;

return