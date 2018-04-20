/*This script is designed to enter GIK from the art sale
Start with excel spreadsheet giving id num, gift ratio, and gift total
start with cursor on id num
*/

#R::
Reload
Return


#G::
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

;move to Advance
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, , 
WinWaitActive, Production Database, , 
Sleep, 100

;batch and ledge are already open, with defaults set
;open new transaction
Send, {F6}

;Enter ID Num
Send, %idNum%
Send, {Tab}
Sleep, 50
;Press 'O' for OK
Send, O
Sleep, 50

;Enter total donated
Send, %giftTotal%

;Save
Send, {F8}
Sleep, 100

;Open Tender Window
Send, {AltDown}{AltUp}
Send, E
Send, T
Sleep, 10

;Tab to GIK Type
Send, {Tab 7}

;Enter A
Send, A

;Tab to Desc window
Send, {Tab}

;Enter GIK percentage info
Send, {ENTER}
Send, %giftRatio%+5 of the value of artwork sold during the SMFA at Tufts Art Sale, Nov. 2017

;Save and exit tender windwo
Send, {F8}
Sleep, 100
Send, {CtrlDown}{F4}{CtrlUp}

;back to square one


WinWait, Excel, 
IfWinNotActive, Excel, , Excel, 
WinActivate, Excel, 
idNum = ;
giftRatio = ;
giftTotal = ;
clipboard = ;

return