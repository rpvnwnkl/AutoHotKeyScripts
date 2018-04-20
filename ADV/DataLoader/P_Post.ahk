
;Context: You are in DataLoader and are ready to Post
;	 
;	
;Instructions: have advance open with DataLoader window selected
;		Press windows key and the letter P
;       Pressing Windows Key and Letter A will reset in case of error
#A::
Reload
Return

#P::
SetTitleMatchMode, 2

;move to advance
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 
Sleep, 50

;pull up menu, key to Edit and then Post
Send, {AltDown}{AltUp}et

return