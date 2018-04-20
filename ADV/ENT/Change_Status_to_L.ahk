
;Context: You have a clipboard list in Advance and want to mark all as Active
;	
;	You have the email or summary window open on the first entry
;	The script opens the entity window and marks status as A
;	Saves, and then closes the entity window. it then advances the clipboard
;	 
;	
;Instructions: have advance open with the clipboard populated 
;		and email window open 
;		Press windows key and the letter A

#A::
;MsgBox A started
SetTitleMatchMode, 2

callGoToBlank(tableText, entID)
{
    SetTitleMatchMode, 2
    Send, {F5} ;call up GoTo
    Sleep, 200
    WinWait, Go To, 
    IfWinNotActive, Go To, , WinActivate, Go To, 
    WinWaitActive, Go To, 
    
    ;MsgBox, %tableText%
    Send, %tableText%
    Send, {TAB}
    Send, %entID%{ENTER} ;Call up given window
    Sleep, 250
    ;MsgBox, I've just entered the info
    WinWait, Production Database, 
    IfWinNotActive, Production Database, , WinActivate, Production Database, 
    WinWaitActive, Production Database, 
    
    
    return
}

pullupEntity(entID)
{
    ;pull up entity window
    Send, {ShiftDown}{F4}{ShiftUp}
    Sleep, 25
    Send, %entID%
    Send, {Enter}
}

CopyCell()
{
    clipboard = ;
    Send, {CTRLDOWN}c{CTRLUP}
    ClipWait ;
    ;clean up id from excel
    ;check if you are in excel so the newline can be dropped
    IfWinActive, Excel, , , 
    {
        ;clean up id from excel
        StringTrimRight, clipboard, clipboard, 2
        Send, {Esc}
    }
    cellOfInterest = %clipboard%
    clipboard = ;
    return cellOfInterest
}

+^L::

;capture ID
entityID := CopyCell()

;move to Advance
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 

;call up entity window
callVar = ENT
callGoToBlank(callVar, entityID)


;Tab to Status field
Send, {TAB 8}
/*
Send, {TAB}
Send, {TAB}
Send, {TAB}
Send, {TAB}
Send, {TAB}
Send, {TAB}
Send, {TAB}
Send, {TAB}
Send, {TAB}
*/
;Enter letter L
Send, L
;add comment
Send, {Tab 2}
;send comment here if you have one
;Save entry and press enter at confirmation dialog
Send, {F8}

return