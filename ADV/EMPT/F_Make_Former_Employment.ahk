
;Context: You have a clipboard list in Advance and want to mark an
;           employment record as past, and paste in a team comment
;	
;	 
;	
;Instructions: have advance open witht the clipboard populated 
;		and empt row selected. press windows key 
;		and the letter F
;       Pressing Windows Key and Letter R will reset in case of error
#R::
Reload
Return

#F::
SetTitleMatchMode, 2

WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 

;Tab to Job Type row (When employer has ID)
Send, {Tab}

;Down to Former Employer
Send, {Down 7}
Sleep, 10

;uncheck Primary Employment box
Send, {Tab 15}{Space}
Sleep, 10

;tab to comment box
Send, {Tab 8}
Sleep, 10

;paste comment of choice
Send, Team 237107

;save changes
Send, {F8}


;navigate to next entry in clipboard


return