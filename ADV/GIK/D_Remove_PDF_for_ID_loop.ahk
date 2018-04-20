/*This script is designed to remove pdf pages containing ids 
    from an excel spreadhseet
Start with excel spreadsheet giving id num, 
start with cursor on id num

PDF doc should be open
*/

#R::
Reload
Return


#D::
SetTitleMatchMode, 2
Loop {
WinWait, Excel, 
IfWinNotActive, Excel, , Excel, 
WinActivate, Excel, ,WinWaitActive, Excel 

;idNum
;capture entity ID (col 1)
idNum = ;
clipboard = ;
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;
;clean up id from excel
idNum = %clipboard%
StringTrimRight, idNum, idNum, 2 ; trim two char from end of string
;MsgBox, IdNum is %idNum%, OK?

;color cell
Send, {Alt}
Send, H
Send, H
Send, {Right 7}
Send, {Enter}
Sleep, 50

;reset for next row
Sleep, 100
clipboard = ;
Sleep, 100
Send, {DOWN}

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
MsgBox, Look ok?

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

}

return