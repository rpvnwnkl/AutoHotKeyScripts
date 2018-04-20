
;Context: You are in the TEL window and want to check whether a number \
;           is a cell or a landline
;	 
;	
;Instructions: have advance open with cursor in the area code field and
;              
;		Press windows key and the letter C
;       Pressing Windows Key and Letter A will reset in case of error
#A::
Reload
Return

#C::
SetTitleMatchMode, 2

;select area-code field
Send, {ShiftDown}{CtrlDown}{Right}
Send, {CtrlUp}{ShiftUp}
Sleep, 10

;copy area code
clipboard = ;
Send, ^c
ClipWait ;

areaCode = %clipboard%
clipboard = ;

;Tab to rest of numbers
Send, {TAB}
Sleep, 10

;select number field
Send, {ShiftDown}{CtrlDown}{Right 3}
Send, {CtrlUp}{ShiftUp}
Sleep, 25

;copy number body
Send, ^c
ClipWait ;

numberBody = %clipboard%
clipboard = ;

wholeNumber = %areaCode%%numberBody%
fronturl = https://lookups.twilio.com/v1/PhoneNumbers/
;backurl = ?CountryCode=US&Type=carrier&Type=caller-name
backurl = ?&Type=carrier&Type=caller-name
url = %fronturl%%wholeNumber%%backurl%
;MsgBox, url equals %url%

;Move to firefox and past number into address bar
WinWait, Mozilla Firefox, 
IfWinNotActive, Mozilla Firefox, , WinActivate, Mozilla Firefox, 
WinWaitActive, Mozilla Firefox, 
Sleep, 250

;key to the address bar
Send, {Alt}
Sleep, 25
Send, F
Sleep, 25
Send, T
Sleep, 250
/*
Send, {ShiftUp}{CtrlDown}T
Send, {CtrlUp}
Sleep, 100
*/
;select all just in case
/*Send, {Shiftup}
Send, {CtrlDown}
Send, A{CtrlUp}
Sleep, 25
*/
;delete address
;Send, {Delete}

;paste address in
clipboard = %url%
ClipWait ;
Send, ^v
Sleep, 25
;Send, %url%

;Msgbox OK?

Send, {Enter}

Return
