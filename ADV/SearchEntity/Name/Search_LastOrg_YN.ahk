
;Context: You are search first and last names in advance
;	Your source is a spreadsheet with first and last side by side
;	the script copies last, then first and tabs over to advance 
;	any oopen window in advance is closed, and then an entity search
;	window is opened
;	Currently this script deletes two invisible characters copied from 
;	The excel file. I haven't found another way to get around it, yet
;Instructions: have advance and excel open, click advance, then excel, so they
;	are the last two programs you've accessed. select the cell of the first 
;	last name you'd like to search, and press Windows Key plus F
;	alt tab back to excel when the new search is to begin

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

Test_for_Nothing(idFound)
{
    ;get window text to check No Entities warning
    Sleep, 1000
    DetectHiddenText, On
    WinGetText, windowCheck 
    ;MsgBox, The text is:`n%windowCheck%
    ;MsgBox 1 idFound = %idFound%

    ;term to check for
    ;windowText = No Entities
    windowText = Looking
    ;Search for windowText in current windows
    IfInString, windowCheck, %windowText%
    {
        idFound := False
        ;Send, {Enter}
        return %idFound%
    }
    return %idFound%
}

Test_for_Multiple(idFound)
{
    ;get window text to check for multiple hits
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
            Exit
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
Capture_ID()
{
    
    ;pull up id window
    Send, {F3}{F3}
    Sleep, 250
    businessID := CopyCell()
    ;close the Id window
    Send, {Esc}
    ;close the entity window
    Send, {CTRLDOWN}{F4}{CTRLUP}
    ;close the entity lookup window
    Send, {CTRLDOWN}{F4}{CTRLUP}
    return businessID
}

searchLastOrg(lastName)
{
    ;open search box, tab to last/org
    Send, {SHIFTDOWN}{F4}{SHIFTUP}
    Send, {TAB 3}
    Send, %lastName%
    Send, {ENTER}
}

searchAgain()
{
    MsgBox, 262145,, Record Found?            
    
        IfMsgBox Cancel 
        {
            MsgBox, Search Again?
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
+^f::
;MsgBox F started
SetTitleMatchMode, 2
Loop 
{
;MsgBox Loop started
WinWait, Excel, 
IfWinNotActive, Excel, , Excel, 
WinWaitActive, Excel, 

businessName := CopyCell()

if not businessName
{
	MsgBox Finished
	exit
}
;prepare cell to have id pasted
Send, {LEFT}

;move to advance
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 

searchLastOrg(businessName)

;search again?
MsgBox, 262145,,Record Selected
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
} else 
{
    idFound := Test_for_Multiple(idFound)
}

;capture id
if idFound
{
    businessID := Capture_ID()
    MsgBox, %businessID%
} else 
{
    ;set advance id to generic default
    businessID = 0
}

;Move back to Excel
WinWait, Excel, 
IfWinNotActive, Excel, , Excel,
WinActivate, Excel, ,WinWaitActive, Excel,

;arrow to next column to paste ID
;Send, {LEFT}

;paste ID from clipboard
If idFound {
    Send, %businessID%
    ;MsgBox real thing
} else {
    Send, 0
}
Sleep, 50
Send, {Enter}{Tab}
}
return
