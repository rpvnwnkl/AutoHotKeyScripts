;Context: You have an ID highlighted and would like to search it in Advance
;	
;	the script copies ID and moves over to advance 
;	an entity search window is opened and the ID pasted in
;	

;Instructions: have advance open, select the ID you'd like to search 
;	execute this script
;	
Sleep, 100
SetTitleMatchMode, 2
 
;Copy highlighted ID
clipboard = ;
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;

;clean up from excel
entityID = %clipboard%
StringTrimRight, entityID, entityID, 2 ; trim two char from end of string


;move to Advance
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 
Sleep, 100

;open Entity search window
Send, {SHIFTDOWN}{F4}{SHIFTUP}
clipboard = %entityID%
ClipWait ;
Sleep, 100

;Paste id into first text box
Send, ^v
Sleep, 100
Send, {ENTER}
Sleep, 100

return