
;Context: You are in DataLoader and are ready to Discard
;	 
;	
;Instructions: have advance open with DataLoader window selected
;		Press windows key and the letter D
;       Pressing Windows Key and Letter R will reset in case of error
#R::
Reload
Return

#S::
SetTitleMatchMode, 2

;release windows key
Send, {LWinDown}

;pull up menu, key to Edit and then Post
Send, {AltDown}er{AltUp}

;Key down once to select Duplicate data
Send, {Down}
Sleep, 10

;Tab and type 'Same Address'
Send, {Tab}
Sleep, 10
Send, Same Address
Sleep, 10
;Msgbox OK?
;Tab and press 'O' to select OK button
Send, {Tab}O
return