
;Context: You are in DataLoader and want to check whether a number \
;           is a cell or a landline
;	 
;	
;Instructions: have advance open with DataLoader window selected and
;               be in the operation window with the operation highlighted
;		Press windows key and the letter C
;       Pressing Windows Key and Letter R will reset in case of error
#R::
Reload
Return

#C::
SetTitleMatchMode, 2

;Tab once to select Telephone field
Send, {Tab}
Sleep, 10

;copy field
clipboard = ;
Send, ^c
ClipWait ;

numberEnd = %clipboard%
clipboard = ;

;Tab to Area Code
Send, {TAB 4}
Sleep, 10

;copy area code
Send, ^c
ClipWait ;

numberFront = %clipboard%
clipboard = ;

wholeNumber = %numberFront%%numberEnd%
fronturl = https://lookups.twilio.com/v1/PhoneNumbers/
backurl = ?CountryCode=US&Type=carrier
url = %fronturl%%wholeNumber%%backurl%
;MsgBox, url equals %url%

;Move to firefox and past number into address bar
WinWait, Mozilla Firefox, 
IfWinNotActive, Mozilla Firefox, , WinActivate, Mozilla Firefox, 
WinWaitActive, Mozilla Firefox, 
Sleep, 100

;click on the address bar
MouseClick, left,  582,  77
Sleep, 25

;delete address
Send, {BackSpace}

;paste address in
clipboard = %url%
Send, ^v

;Msgbox OK?

Send, {Enter}

Return
