
;Context: You are in DataLoader and are ready to Discard
;	 
;	
;Instructions: have advance open with DataLoader window selected
;		Press windows key and the letter B
;       Pressing Windows Key and Letter R will reset in case of error
#A::
Reload
Return

~#S::
SetTitleMatchMode, 2

;collect street address
Send, {TAB 9}

;copy street line one to clipboard
clipboard = ;
Send, ^c
ClipWait ;
streetLineOne = %clipboard%

;move to next line
Send, {TAB}

;copy street line two to clipboard
clipboard = ;
Send, ^c
ClipWait ;
streetLineTwo = %clipboard%

;move to city name
Send, {TAB 3}

;copy city to clipboard
clipboard = ;
Send, ^c
ClipWait ;
cityName = %clipboard%

;move to state 
Send, {TAB}

;copy state abbr to clipboard
clipboard = ;
Send, ^c
ClipWait ;
stateAbbr = %clipboard%

;Msgbox, street: %streetLineOne%, street2: %streetLineTwo%, city: %cityName%, statename: %stateAbbr%

;move to last name 
Send, {TAB 13}

;copy last name to clipboard
clipboard = ;
Send, ^c
ClipWait ;
lastName = %clipboard%

;move to first name 
Send, {TAB}

;copy first name to clipboard
clipboard = ;
Send, ^c
ClipWait ;
firstName = %clipboard%

;search for name
;pull up ent search
Send, {SHIFTDOWN}{F4}{SHIFTUP}{TAB}{TAB}{TAB}

;past last name
Send, %lastName%

;move to next field and enter first name
Send, {TAB}
Send, %firstName%
Send, {ENTER}

return