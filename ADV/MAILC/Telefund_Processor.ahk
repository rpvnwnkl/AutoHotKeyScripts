;
; AutoHotkey Version: 1.x
; Language:       English
; Platform:       Win9x/NT
; Author:         A.N.Other <myemail@nowhere.com>
;
; Script Function:
;	Template script (you can customize this template by editing "ShellNew\Template.ahk" in your Windows folder)
;

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;Context: You are inactivating a preferred address
;
;v 1.0
;Pressing windows key + A reloads the script in case it gets wonky


#A::
Reload
Return

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
    WinWait, ADVTRN, 
    IfWinNotActive, ADVTRN, , WinActivate, ADVTRN, 
    WinWaitActive, ADVTRN, 
    
    
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

GetRow()
{
    global ;
    ;excel
    WinWait, Excel, 
    IfWinNotActive, Excel, , WinActivate, Excel, 
    WinWaitActive, Excel,
    ;assuem correct cell selected
    callID := CopyCell()
    ;check to see if the row is empty/without primary donor id
    if (callID < 1) 
    {
        exit
    }
    ;tab over twice to get call result
    Send, {TAB 2}
    callResult := CopyCell()
    ;tab over and get comment
    Send, {TAB}
    callComment := CopyCell()
    ;reset to next row
    Send, {ENTER}
    return
}

MarkNoTelefund(callID)
{
    global ;
    pullupEntity(callID)
    callVar = MAILC
    callGoToBlank(callVar, callID)
    MsgBox, No Telefund
    Send, {TAB}
    currVal := CopyCell()
    if not currVal
    {
        ;it's empty
        ;all other values are more exclusive
        Send, 3   
    } 
    ;Tab to No Call field
    Send, {TAB 11}
    Send, Y
    ;Tab to comment
    Send, {TAB 3}
    CommentCRL()
}

MarkNoSolicit(callID)
{
    global ;
    pullupEntity(callID)
    callVar = MAILC
    callGoToBlank(callVar, callID)
    MsgBox, No Solicitation
    Send, {TAB}
    currVal := CopyCell()
    IfNotEqual, currVal, 1
    {
        ;it's not 1
        ;change to 1 as it's most selective
        Send, 1
    }
    ;ensure proper controls are set
    ;move to S Mail field
    Send, {TAB 6}
    Send, Y
    ;move to S email
    Send, {TAB 3}
    Send, Y
    ;move to S Calls
    Send, {TAB 2}
    Send, Y

    ;check for comment
    Send, {TAB 3}
    CommentCRL()
    return
    
}
CommentCRL()
{
    ;select All
    Send, {Alt}
    Send, el
    currVal := CopyCell()
    IfEqual, currVal,
    {
        ;it's empty
        Send, Per CRL
    }Else
    {
        ;it's not empty
        ;does it contain CRL already?
        IfNotInString, currVal, CRL
        {
            Send, {RIGHT}
            Send, `; Per CRL
        }
        
    }
}

SavePrompt()
{
    MsgBox, 4, Save Changes, Save Modifications?
    IfMsgBox, Yes
    {
        ;save mailc
        Send, {F8}
        ;close mailc
        Send, {CtrlDown}{F4}{CtrlUp}
        ;close entity window
        Send, {CtrlDown}{F4}{CtrlUp}
    } Else
    {
        Send, {F7}
        ;close mailc
        Send, {CtrlDown}{F4}{CtrlUp}
        ;close entity window
        Send, {CtrlDown}{F4}{CtrlUp}
    }
    return
}
+^G::
SetTitleMatchMode 2 
Loop
{

;Capture Excel Row info
callID = ;
callResult = ;
callComment = ;
GetRow()

;move to advance
WinWait, ADVTRN, 
IfWinNotActive, ADVTRN, , WinActivate, ADVTRN, 
WinWaitActive, ADVTRN,

changeMade = False
;Call result codes
DoNotCall = Do Not Call
RemoveList = Remove From List
Deceased = Deceased
IfEqual, callResult, %Deceased%
{
    MsgBox, This person is Deceased.
}
IfEqual, callResult, %DoNotCall%
{
    MarkNoTelefund(callID)
    changeMade = True
} Else IfEqual, callResult, %RemoveList%
{
    MarkNoSolicit(callID)
    changeMade = True
}
IfNotEqual, callComment,
{
    ;a comment exists
    MsgBox, 1, Comment in CRL, `nComment: `n`n%callComment%
    IfMsgBox, Cancel
    {
        exit
    }        
}
IfEqual, changeMade, True
{
    SavePrompt()
}
}
return