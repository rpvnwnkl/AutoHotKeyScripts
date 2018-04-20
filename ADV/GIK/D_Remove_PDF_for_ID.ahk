/*This script is designed to remove pdf pages containing ids 
    from an excel spreadhseet
Start with excel spreadsheet giving id num, 
start with cursor on id num

PDF doc should be open

Press Windows Key + D

*/

#R::
Reload
Return


#D::
SetTitleMatchMode, 2

WinWait, Excel, 
IfWinNotActive, Excel, , Excel, 
WinActivate, Excel, ,WinWaitActive, Excel 

;idNum
;capture entity ID (col 1)
clipboard = ;
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;
;clean up id from excel
idNum = %clipboard%
StringTrimRight, idNum, idNum, 2 ; trim two char from end of string
;MsgBox, IdNum is %idNum%, OK?

;move to Acrobat
WinWait, Acrobat, 
IfWinNotActive, Acrobat, , WinActivate, Acrobat, , 
WinWaitActive, Acrobat, , 
Sleep, 100

;open search dialog
Send, {Alt}
Send, E
Send, F
Sleep, 100

;Enter ID Num and press enter
Send, %idNum%
Send, {Enter}
Sleep, 1000
MsgBox, 4,, Look ok?
IfMsgBox, No
    return
;should have found it by now

;close Find dialog box
Send, {Alt}
Send, E
Send, F
Send, {Esc}
Sleep, 50

;delete page
Send, {ShiftDown}{CtrlDown}D{CtrlUp}{ShiftUp}
Sleep, 50

;press enter twice
Send, {Enter}
Sleep, 50
Send, {Enter}
Sleep, 50

;back to square one

WinWait, Excel, 
IfWinNotActive, Excel, , Excel, 
WinActivate, Excel, 

;color cell
Send, {Alt}
Send, H
Send, H
Send, {Right 7}
Send, {Enter}
Sleep, 50

;reset for next row
Sleep, 100
Send, {DOWN}

idNum = ;
clipboard = ;

return