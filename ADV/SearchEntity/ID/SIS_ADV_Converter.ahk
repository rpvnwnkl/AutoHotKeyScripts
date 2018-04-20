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

;Move to cursor to next col
clipboard = ;
Send, {Esc}
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

;#####################################
;case for multiple results from lookup
;#####################################
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
        return
    ;check if profile was opened or not
    WinGetText, windowCheck
    ;Search for windowText in current windows
    windowText = Custom
    IfNotInString, windowCheck, %windowText%
    {
       Send, {Enter} 
    }
}
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

;Move back to Excel
WinWait, Excel, 
IfWinNotActive, Excel, , Excel,
WinActivate, Excel, ,WinWaitActive, Excel,

;arrow to next column to paste ID
;Send, {LEFT}

;paste ID from clipboard
Send, %advID%
Sleep, 50
Send, {Enter}{Tab}
return