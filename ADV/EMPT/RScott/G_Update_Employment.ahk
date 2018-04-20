/*This script is designed to enter new parents for 
medical students 
*/

#A::
Reload
Return

#MaxMem 256 ; Allow 256 MB per variable

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

CopyCell()
{
    clipboard = ;
    Send, {CTRLDOWN}c{CTRLUP}
    ClipWait ;
    ;clean up id from excel
    StringTrimRight, clipboard, clipboard, 2 ; trim two char from end of string
    cellOfInterest = %clipboard%
    Send, {Esc}
    return cellOfInterest
}



pullupEntity(entID)
{
    ;pull up entity window
    Send, {ShiftDown}{F4}{ShiftUp}
    Sleep, 25
    Send, %entID%
    Send, {Enter}
}
/*
#####################################
SCRIPT START
*/
#G::
SetTitleMatchMode, 2

Loop {
;#####################################
;COLLECT DATA FROM EXCEL
;#####################################
WinWait, Excel, 
IfWinNotActive, Excel, , Excel, 
WinActivate, Excel, ,WinWaitActive, Excel 

;###############################
; Capture Info from cells
entityID = ;
industryDesc = ;
empName = ;

;entityID
;capture entityID (col 1)
;should already be selected
entityID := CopyCell()
;MsgBox, %entityID%

;check to see if the row is empty/without primary donor id
if (entityID < 1) {
    MsgBox, No Entity ID
    return
}

;industryDesc
;capture Industry Description (col 2)
Send, {Tab 1}
industryDesc := CopyCell()

;empName
;capture Employer Name (col 4)
Send, {Tab 2}
empName := CopyCell()

;########################################
;MOVE TO ADVANCE
;########################################

;move to Advance
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, , 
WinWaitActive, Production Database, , 
Sleep, 200

;pullupEntity(entityID)
;close previous windows
Send, {CtrlDown}{F4}{CtrlUp}
Sleep, 100
callVar = EMPT
callGoToBlank(callVar, entityID)

/*
;gather window details to see
;how many jobs entity has listed
WinGetText, emptWindow,,,,
Sleep, 200
;MsgBox, The text is:`n%emptWindow%
;FileReadLine, line, emptWindow, 2
;MsgBox, %line%
Loop, parse, emptWindow, `n, `record
{
    IfEqual, A_Index, 2
    {
        emptInfo = %A_LoopField%
    }
    ;MsgBox, Text row %A_Index% is %A_LoopField%.`n
}
;Now we parse the row to get the actual number
StringGetPos, lefParen, emptInfo, (
StringGetPos, rigParen, emptInfo, )
StringMid, emptNums, emptInfo, (lefParen+2), (rigParen-lefParen-1)
;MsgBox, %emptNums%
*/
;Prompt user to choose primary employment
MsgBox, Select the Primary Employment Row.

clipboard = ;
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;
employerID = %clipboard%
clipboard = ;

If StrLen(employerID) > 2
{
    ;If employer is an entity
    Send, {TAB 23}
} else
{
    ;If employer isn't entity
    Send, {TAB 25}
}

;paste industry code
Send, %industryDesc%

;save
Send, {F8}
Send, {Enter}
Send, {Enter}
Send, {Enter}


;back to excel
WinWait, Excel, 
IfWinNotActive, Excel, , Excel, 
WinActivate, Excel,

;shift tab back to id field
Send, {ShiftDown}{Tab 3}{ShiftUp}
;Alt down to color field
Send, {AltDown}{AltUp}HH{Right 5}
Send, {Enter}


;Enter to go to next row
Send, {Enter}
}
return