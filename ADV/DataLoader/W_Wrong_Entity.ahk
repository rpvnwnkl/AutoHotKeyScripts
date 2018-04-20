
;Context: You are in DataLoader and are ready to mark as wrong
;           entity afer discarding
;	 
;	
;Instructions: have advance open with DataLoader window selected and already
;       pressed keys to discard
;		Press windows key and the letter w
;       Pressing Windows Key and Letter R will reset in case of error
#R::
Reload
Return

#W::
SetTitleMatchMode, 2

;Key down twice to select Insufficient data
Send, {Down}{Down}
Sleep, 10

;Tab and type 'Wrong entity'
Send, {Tab}
Sleep, 10
Send, Wrong entity
Sleep, 10
Msgbox OK?
;Tab and press 'O' to select OK button
Send, {Tab}O

Return
