
;Context: You want to mark phone as mobile instead of home
;	 
;	
;Instructions: 
;		Press windows key and the letter M
;       Pressing Windows Key and Letter R will reset in case of error
#R::
Reload
Return

#M::
SetTitleMatchMode, 2

;move to advance
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 
Sleep, 50

;reverse tab back up to Type field
Send, {ShiftDown}{Tab 3}{ShiftUp}

;go up twice to select HC
Send, {UP 2}

;send menu commands to post
;pull up menu, key to Edit and then Post
Send, {AltDown}et{AltUp}


return