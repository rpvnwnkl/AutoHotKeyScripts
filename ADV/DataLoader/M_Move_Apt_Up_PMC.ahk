
;Context: You want to move apt info to first line
;	 
;	
;Instructions: 
;		Press windows key and the letter M
;       Pressing Windows Key and Letter A will reset in case of error
~#A::
Reload
Return

#M::
SetTitleMatchMode, 2


;Move down to street name line 2
Send, {Tab 10}

;copy to clipboard
clipboard = ;
Send, ^c{Delete}
ClipWait ;
Sleep, 10
;move up to second line
Send, {ShiftDown}{Tab}{ShiftUp}{Right}
;add comma, space
Send, ,{Space}

;MsgBox, clipboard = %clipboard%
;now we have an opportunity to evaluate the second line of the address
IfInString, clipboard, %A_Space%
{
    StringSplit, lineTwo, clipboard , %A_Space%, %A_Space%%A_Tab%
    ;MsgBox, lineTwo1: %lineTwo1%, lineTwo2: %lineTwo2%
    IfInString, lineTwo1, Apt
    {
        ;MsgBox, this was found
        ;ensure apt is abbr., capitalized, and has a period
        lineTwo1 = Apt.
        ;MsgBox, lineTwo0 = %lineTwo0%
        If %lineTwo3%
        {
            ;MsgBox, True
            lineRest = %lineTwo2%%A_Space%%lineTwo3%
        } else
        {
            lineRest = %lineTwo2%
        }
        ;send second new second line to end of first line
        Send, %lineTwo1%%A_Space%%lineRest%


    }
    else
    {
        ;it is a unit or otherwise not apt
        Send, %clipboard%
    }
    /*
    ;for now, just paste it upstairs
    ;up to previous line, over to clear highlight, add comma
    Send, {ShiftDown}{Tab}{ShiftUp}{Right}
    Send, ,{Space}
    ;paste clipboard
    Send, %clipboard%
    */
    Sleep, 15
    
    
} else {
    ;there is not qualifier (apt, unit, etc)
    ;add apt and call it a day
    ;up to previous line, over to clear highlight, add comma
    ;Send, {ShiftDown}{Tab}{ShiftUp}{Right}
    ;Send, ,{Space}
    Send, Apt.
    Send, {Space}

    ;paste clipboard
    Send, %clipboard%
    Sleep, 15
    }
;back to add period
;Send, {CtrlDown}{Left}{CtrlUp}{Left}.


return