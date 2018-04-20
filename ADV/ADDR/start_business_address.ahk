
;Context: You are on entity main window and want to 
;           create a new business address
;	
;	 
;	
;Instructions: have entity open, press hotkey

#W::
SetTitleMatchMode, 2

WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 

;Open GoTo
Send, {F5}
Sleep, 25

;call up address window
Send, ADDR 
Send, {ENTER}
Sleep, 25

;new address
Send, {F6}
Sleep, 25

;change to business type
Send, {AltDown}Y{AltUp}
Send, B
Sleep, 25

;add SIP code to source field
Send, {Tab}{Tab}{Tab}{Tab}{Tab}{Tab}{Tab}{Tab}{Tab}{Tab}
Send, SIP

;move to zip, then back up to address line 1
Send, {AltDown}Z{AltUp}
Send, {ShiftDown}{Tab}{Tab}{Tab}{Tab}{Tab}{Tab}{ShiftUp}
Sleep, 25

;call up instant address
Send, {AltDown}3{AltUp}

return
