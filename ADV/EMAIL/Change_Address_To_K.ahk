
;Context: You have a clipboard list in Advance and want to mark 1
;	email as Last known on all of them.
;	You have the email window open on the first entry
;	The script chagnes the selected email
;	
;	 
;	
;Instructions: have advance open witht the clipboard populated 
;		and email row selected. press windows key 
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



;}
return