
;Context: You are in DataLoader and are ready to Discard 
;           and mark insufficient data
;	 
;	
;Instructions: have advance open with DataLoader window selected
;		Press windows key and the letter I
;       Pressing Windows Key and Letter R will reset in case of error
#R::
Reload
Return

#I::
SetTitleMatchMode, 2

;pull up menu, key to Edit and then Post
Send, {AltDown}er{AltUp}

;Key down twice to select Insufficient data
Send, {Down}{Down}
Sleep, 10

;Tab and type 'Insufficient data'
Send, {Tab}
Sleep, 10
Send, bad data
Sleep, 10
;Msgbox OK?
;Tab and press 'O' to select OK button
Send, {Tab}O
return