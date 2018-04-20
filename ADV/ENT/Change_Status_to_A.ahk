
;Context: You have a clipboard list in Advance and want to mark all as Active
;	
;	You have the email or summary window open on the first entry
;	The script opens the entity window and marks status as A
;	Saves, and then closes the entity window. it then advances the clipboard
;	 
;	
;Instructions: have advance open with the clipboard populated 
;		and email window open 
;		Press windows key and the letter A

#A::
;MsgBox A started
SetTitleMatchMode, 2

;Loop 
;{
;MsgBox Loop started

WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 

;PUll up GoTo window
Send, {F5}
WinWait, Go To ..., 
IfWinNotActive, Go To ..., , WinActivate, Go To ..., 
WinWaitActive, Go To ..., 

;PUll up Entity window
Send, ent{ENTER}
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 
Sleep, 50

;Tab to Status field
Send, {TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}

;Enter letter A
Send, A
;Save entry and press enter at confirmation dialog
Send, {F8}
Sleep, 25
Send, {ENTER}{ENTER}
Sleep, 50
;Close Entity window
Send, {CTRLDOWN}{F4}{CTRLUP}
Sleep, 50
;Close Bio Summary window
Send, {CTRLDOWN}{F4}{CTRLUP}
Sleep, 50
;Move to next entry in clipboard
Send, {ALTDOWN}{ALTUP}bn


;PUll up GoTo window
Send, {F5}
WinWait, Go To ..., 
IfWinNotActive, Go To ..., , WinActivate, Go To ..., 
WinWaitActive, Go To ..., 

;PUll up BSUM window
Send, BSUM{ENTER}
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 
Sleep, 50




;}
return