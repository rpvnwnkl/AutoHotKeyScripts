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

#IfWinActive, Production Database
RAlt::
{
    ;open AGISML Row
    Send, {ENTER}
    Sleep, 10
    ;maximize window
    Send, {AltDown}{AltUp}WL
    return
}


;close window
#IfWinActive, Production Database
Escape::
{
    Send, {CtrlDown}{F4}{CtrlUp}
    Sleep, 10
    Send, {TAB 2}
    return
}

