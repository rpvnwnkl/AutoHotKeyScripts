
;Context: You have a clipboard list in Advance and want to mark up to 2
;	emails as Last known on all of them.
;	You have the email window open on the first entry
;	The script chagnes the first email and tries to change a second one
;	It then moves on to the next clipboard entry
;	 
;	
;Instructions: have advance open witht the clipboard populated. press windows key 
;		and the letter K

#K::
;MsgBox K started
SetTitleMatchMode, 2

;Loop 
;{
;MsgBox Loop started

WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 

Sleep, 50

;Change first email to Last Known
Send, {TAB}{TAB}k{F8}
Sleep, 25

;After saving, move down to next email
Send, {DOWN}

;Mark this email Last Known
Send, {TAB}{TAB}k{F8}
Sleep, 25

;Go to next entry on clipboard
Send, {ALTDOWN}{ALTUP}bn


;}
return