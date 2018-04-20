;Context: You are in the ENT window for an Entity, 
;and you wish to create an informal individual salutation in the SAL window
;Instructions: Press WindowsKey+i 
;The SAL window will stay open afterwards unless you delete the first semi-colon
;in the second to last line of this script

#A::
Reload
Return

callGoTo(tableText)
{
    SetTitleMatchMode, 2
    Send, {F5} ;call up GoTo
    Sleep, 200
    WinWait, Go To, 
    IfWinNotActive, Go To, , WinActivate, Go To, 
    WinWaitActive, Go To, 
    
    ;MsgBox, %tableText%
    Send, %tableText%
    ;Send, {TAB}
    ;Send, %entID%
    ;Call up given window
    Send, {ENTER} 
    Sleep, 250
    ;MsgBox, I've just entered the info
    WinWait, Production Database, 
    IfWinNotActive, Production Database, , WinActivate, Production Database, 
    WinWaitActive, Production Database, 
    
    
    return
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


#i::
SetTitleMatchMode, 2
If WinActive("Production Database")
{
    ;Go to "Last" name category, TAB to first
    Send, {ALTDOWN}L{ALTUP}{TAB}
    ;copy first name
    firstName := copyCell()
    ;pullup SAL window
    callVar = SAL
    callGoTo(callVar)
    ;Now create a new entry row
    Send, {F6}
    ;paste the first name in
    Send, %firstName%
    ;select II
    Send, {TAB}{TAB}ii
    ;save
    Send, {F8}
    ;close the sal window
    Send, {CTRLDOWN}{F4}{CTRLUP} 

}
return
