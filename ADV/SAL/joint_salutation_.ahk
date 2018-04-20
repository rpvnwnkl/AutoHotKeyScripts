;Context: You need to create joint informal salutations in the SAL windows for two entities
; You are in the MAR window for the entity whose name should come first,
; and the entities should not have any joint informal salutations in the SAL window. You should
; delete them first.
;Instructions: press WindowsKey+j
;Currently the SAL windows stays open after finishing, to have it close itself just delete the first
; semi-colon in the third to last line, before CTRLDOWN

#j::
SetTitleMatchMode, 1
clipboard = ;Start off empty to allow ClipWait to detect when the text has arrived
WinWait, Production, 
IfWinNotActive, Production, , WinActivate, Production, 
WinWaitActive, Production, 
;MouseClick, left,  430,  131
Sleep, 100
Send, {F5}
WinWait, Go To ..., 
IfWinNotActive, Go To ..., , WinActivate, Go To ..., 
WinWaitActive, Go To ..., 
Send, sal{ENTER}
Send, {SHIFTDOWN}{F3}{SHIFTUP} ; go to spouse 2 record
Send ^c ; copy spouses informal sal name
ClipWait  ; Wait for the clipboard to contain text.
spouse2 = %clipboard%
clipboard = 
Send, {SHIFTDOWN}{F3}{SHIFTUP} ; go to spouse 1 record
Send ^c ; copy spouses informal sal name
ClipWait  ; Wait for the clipboard to contain text.
spouse1 = %clipboard%
formalSal = %spouse1% and %spouse2%
;MsgBox Control-C copied the following contents to the clipboard:`n`n%formalSal%
;return
clipboard = %formalSal%
Send, {F6}^v
Send, {TAB  2}ji{F8}
Send, {SHIFTDOWN}{F3}{SHIFTUP}
Send, {F6}^v
Send, {TAB  2}ji{F8}
Send, {CTRLDOWN}{F4}{CTRLUP} ;Remove or add the semi-colon before CTRLDOWN to have the SAL window close itself
clipboard =
return
