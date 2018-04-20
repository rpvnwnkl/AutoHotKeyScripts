
;Context: You are in DataLoader and are ready to Discard
;	 
;	
;Instructions: have advance open with DataLoader window selected
;		Press windows key and the letter D
;       Pressing Windows Key and Letter R will reset in case of error
#A::
Reload
Return

#D::
SetTitleMatchMode, 2

;pull up menu, key to Edit and then Post
Send, {AltDown}er{AltUp}
;choose duplicate and exit
Send, {Down}
Send, {Tab 2}
Send, O

return