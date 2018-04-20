
;Context: You have just done an entity search, need to pick a row
;press windows key plus P

#R::
Reload
Return


#P::
;MsgBox F started
SetTitleMatchMode, 2

;select row, and then send enter
Send, {ENTER}
Sleep, 300

;pull up id window
Send, {F3}{F3}
Sleep, 100
MsgBox, Hello
Sleep, 100
;copy id to clipboard
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;

return