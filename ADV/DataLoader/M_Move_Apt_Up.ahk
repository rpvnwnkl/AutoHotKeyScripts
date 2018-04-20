
;Context: You want to move apt info to first line
;	 
;	
;Instructions: 
;		Press windows key and the letter M
;       Pressing Windows Key and Letter A will reset in case of error
#A::
Reload
Return

#M::
SetTitleMatchMode, 2


;Move down to street name line 2
Send, {Tab 6}

;copy to clipboard
clipboard = ;
Send, ^c{Delete}
ClipWait ;
Sleep, 10

;up to previous line, over to clear highlight, add comma
Send, {ShiftDown}{Tab}{ShiftUp}{Right},{Space}

;paste clipboard
Send, %clipboard%
Sleep, 15

;back to add period
;Send, {CtrlDown}{Left}{CtrlUp}{Left}.


return