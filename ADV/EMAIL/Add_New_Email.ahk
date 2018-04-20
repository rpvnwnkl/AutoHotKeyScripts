;Context: You have an ID highlighted and would like to search it in Advance
;	
;	the script copies ID and moves over to advance 
;	an entity search window is opened and the ID pasted in
;	

;Instructions: have advance open, select the ID you'd like to search 
;	execute this script by pressing the middle mouse button
;	
;v 1.0
;Pressing windows key + A reloads the script in case it gets wonky
#A::
Reload
Return


CopyCell()
{
    clipboard = ;
    Send, {CTRLDOWN}c{CTRLUP}
    ClipWait ;
    cellOfInterest = %clipboard%
    clipboard = ;
    return cellOfInterest
}

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

#E::
SetTitleMatchMode, 2

;capture entity id
entityID := CopyCell()
Sleep, 25
;capture team id
Send, {ShiftDown}{Tab 15}{ShiftUp}
teamID := CopyCell()
Sleep, 25
;extract email from text field
;move to text field
Send, {TAB 20}
Sleep, 100
Send, {CtrlDown}A{CtrlUp}

requestText := CopyCell()
MsgBox, copied text
;parsing text
searchString := "[\_]*([a-zA-Z0-9]+(?:\.|\_*)?)+@((?:[a-z][a-z0-9\-]*(?:\.|\-*\.))+([a-z]{2,6}))"
Loop
{
foundpos := RegExMatch(requestText, searchString, match), nextpos := foundpos+strlen(match)
;MsgBox, %foundpos%
StringTrimLeft, requestText, requestText, %nextpos%
if (match != "")
    result .= match . "`n"
Else
    Break
}
MsgBox, %result%
entEmail = %result%

IfWinNotExist, Production Database
{
    MsgBox, Advance isn't open
    return
}

;move to Advance
WinWait, Production Database,
IfWinNotActive, Production Database, , WinActivate, Production Database,
WinWaitActive, Production Database,
Sleep, 100



;open Entity search window
pullUpEntity(entityID)

;open email window
callVar = EMAIL
callGoToBlank(callVar, entityID)

;add new email
Send, {F6}
Send, %entEmail%

;set preferred
Send, {TAB}{SPACE}
;or
;don't set preferred
;Send, {TAB}
;set status date
FormatTime, timeVar, ,ShortDate
Send, {TAB 2}%timeVar%

;prompt for P or B email type
emailType = P

#SingleInstance
SetTimer, ChangeButtonNames, 50
MsgBox, 4, Personal or Business, Choose an email type:
IfMsgBox, Yes
    emailType = P
Else
    emailType = B


;set email type
Send, {TAB}%emailType%

;Set BIO source code
Send, {TAB 3}BIO

;add teamID to comment
Send, {TAB 6}
Send, Team %teamID%


;save
Send, {F8}

return

ChangeButtonNames:
IfWinNotExist, Personal or Business
    return ; keep waiting
SetTimer, ChangeButtonNames, off
WinActivate
ControlSetText, Button1, &Personal
ControlSetText, Button2, &Business
return