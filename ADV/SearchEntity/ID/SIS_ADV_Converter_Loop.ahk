;Context: You have an ID highlighted and would like to search it in Advance
;	
;	the script copies ID and moves over to advance 
;	an entity search window is opened and the ID pasted in
;	

;Instructions: have advance open, select the ID you'd like to search 
;	execute this script
;
#A::
Reload
return




;#IfWinActive Excel
^+g::  ; Control + Shift + g  Hotkey
;MsgBox G started
SetTitleMatchMode, 2
;SetTitleMatchMode, slow
Loop {
WinWait, Excel, 
IfWinNotActive, Excel, , WinActivate, Excel, 
WinWaitActive, Excel,

;MsgBox you're it



;Copy highlighted ID
clipboard = ;
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;
sisID = %clipboard%
;clean up from excel
StringTrimRight, sisID, sisID, 2 ; trim two char from end of string
Sleep, 100
clipboard = ;
Send, {Esc}

;check for end of column
if (sisID = "") {
    ;MsgBox, Cooked
    Break
}

;Move to cursor to next col
Sleep, 100
Send, {ShiftDown}{tab}{ShiftUp}

;move to Advance
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 
Sleep, 100

;open Entity search window
Send, {SHIFTDOWN}{F4}{SHIFTUP}
clipboard = %sisID%
ClipWait ;
Sleep, 100

;Paste id into first text box
Send, ^v
Sleep, 50
;tab and press space to check alt ID box
Send, {TAB}{SPACE}
Sleep, 50
Send, {ENTER}
Sleep, 100

;boolean value to help determine search results
idFound = True

;see if query gets no results
idFound := Test_for_Nothing(idFound)
;MsgBox idFound = %idFound%

;if no results, close dialogs, otherwise proceed
if not idFound
{
    ;go right back to excel
    ;MsgBox idFound = %idFound%
    Send, {Enter}
    ;close search box
    Send, {CTRLDOWN}{F4}{CTRLUP}
} else {
    idFound := Test_for_Multiple(idFound)
    }

;capture id
if idFound
{
GoSub, Capture_ID
} else {
    ;set advance id to generic default
    advID = 0
}

;Move back to Excel
WinWait, Excel, 
IfWinNotActive, Excel, , Excel,
WinActivate, Excel, ,WinWaitActive, Excel,

;arrow to next column to paste ID
;Send, {LEFT}

;paste ID from clipboard
If idFound {
    Send, %advID%
    ;MsgBox real thing
} else {
    Send, 0
}
Sleep, 50
Send, {Enter}{Tab}
}
return

;sub routines
;#####################################
;case for no results from lookup
;#####################################
Test_for_Nothing(idFound)
{
;get window text to check No Entities warning
Sleep, 500
DetectHiddenText, On
WinGetText, windowCheck 
;MsgBox, The text is:`n%windowCheck%
;MsgBox 1 idFound = %idFound%

;term to check for
windowText = No Entities
;Search for windowText in current windows
IfInString, windowCheck, %windowText%
{
    idFound := False
    Send, {Enter}
    return %idFound%
}
return %idFound%
}

;#####################################
;case for multiple results from lookup
;#####################################
Test_for_Multiple(idFound)
{
;get window text to check for multiple hits
;DetectHiddenText, On
WinGetText, windowCheck 
;MsgBox, The text is:`n%windowCheck%

;term to check for
windowText = Criteria
;Search for windowText in current windows
IfInString, windowCheck, %windowText%
{
    ;WinSetTitle, SIS Converter, , NewTitle
    ControlSetText, Static1, User Input Required..., SIS Converter
    MsgBox, 262145,, Please open the correct entity record            
    
    IfMsgBox Cancel 
    {
        idFound = False
        return idFound
    }
    ;check if profile was opened or not
    WinGetText, windowCheck
    ;Search for windowText in current windows
    windowText = Custom
    IfNotInString, windowCheck, %windowText%
    {
       Send, {Enter} 
    }
}
return idFound
}

;ID capture subroutine
Capture_ID:
;pull up id window
Send, {F3}{F3}
Sleep, 250
;MsgBox, Hello
;Sleep, 100

;copy id to clipboard
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;
advID = %clipboard%
;MsgBox, clipboard is %advID%
;close the Id window
Send, {Esc}

;close the entity window
Send, {CTRLDOWN}{F4}{CTRLUP}

;close the entity lookup window
Send, {CTRLDOWN}{F4}{CTRLUP}
return


;get window text to check for alerts
WinGetText, alertCheck 
;MsgBox, The text is:`n%alertCheck%

;If found, alert popped up on record and we close it
IfInString, alertCheck, alert
{
    ;MsgBox, Closing this window
    ;close window
    Send, {CtrlDown}{F4}{CtrlUp}
    ;move to donor info window
    WinWait, Donor Information, 
    IfWinNotActive, Donor Information, , WinActivate, Donor Information, , 
    WinWaitActive, Donor Information, , 
    Sleep, 100
    ;close window
    Send, O
    ;move back to advance
    WinWait, Production Database, 
    IfWinNotActive, Production Database, , WinActivate, Production Database, , 
    WinWaitActive, Production Database, , 
    Sleep, 100
}