/*This script is designed to enter GIK from the art sale
Start with excel spreadsheet giving id num, gift ratio, and gift total
start with cursor on id num

Press Windows Key +G to start
*/

#R::
Reload
Return


#G::
SetTitleMatchMode, 2
Loop{
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

;giftRatio
;clear clipboard, get percentage given
clipboard = ;
Send, {Right 3}
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;
;clean up percentage from excel
giftRatio = %clipboard%
StringTrimRight, giftRatio, giftRatio, 2 ; trim two char from end of string

;giftTotal
;clear clipboard, get total given
clipboard = ;
Send, {Right}
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;
;clean up number from excel
giftTotal = %clipboard%
StringTrimRight, giftTotal, giftTotal, 2 ; trim two char from end of string

;reset for next row
Sleep, 100
clipboard = ;
Sleep, 100
Send, {DOWN}{Left 4}

;MsgBox, idNum is %idNum%, giftRatio is %giftRatio%, giftTotal is %giftTotal%

;move to Production Database
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, , 
WinWaitActive, Production Database, , 
Sleep, 200

;batch and ledge are already open, with defaults set
;open new transaction
Send, {F6}
Sleep, 100

;Enter ID Num
Send, %idNum%
Send, {Tab}
Sleep, 100
;Press 'O' for OK
Send, O
Sleep, 100

;Enter total donated
Send, %giftTotal%
Sleep, 100

;Save
Send, {F8}
Sleep, 200

;Open Tender Window
Send, {AltDown}{AltUp}
Send, E
Send, T
Sleep, 100

;Tab to GIK Type
Send, {Tab 7}
Sleep, 50

;Enter A
Send, A
Sleep, 50

;Tab to Desc window
Send, {Tab}
Sleep, 50

;Enter GIK percentage info
Send, {ENTER}
Send, %giftRatio%+5 of the value of artwork sold during the SMFA at Tufts Art Sale, Nov. 2017
Sleep, 50

;Save and exit tender windwo
Send, {F8}
Sleep, 200
Send, {CtrlDown}{F4}{CtrlUp}
Sleep, 300

;back to square one

idNum = ;
giftRatio = ;
giftTotal = ;
clipboard = ;
Sleep, 200
}
return