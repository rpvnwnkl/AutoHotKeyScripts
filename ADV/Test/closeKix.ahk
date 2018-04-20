/*>>>>>>>>>>( Window Title & Class )<<<<<<<<<<<
Open File - Security Warning
ahk_class #32770

>>>>>>>>>>>>( Mouse Position )<<<<<<<<<<<<<
On Screen:	1353, 437  (less often used)
In Active Window:	667, 80

>>>>>>>>>( Now Under Mouse Cursor )<<<<<<<<

Color:	0xF0F0F0  (Blue=F0 Green=F0 Red=F0)

>>>>>>>>>>( Active Window Position )<<<<<<<<<<
left: 686     top: 357     width: 548     height: 332

>>>>>>>>>>>( Status Bar Text )<<<<<<<<<<

>>>>>>>>>>>( Visible Window Text )<<<<<<<<<<<
We can’t verify who created this file. Are you sure you want to run this file?
Name:
\\tufts.ad.tufts.edu\netlogon\KIX32.EXE
Type:
Application
From:
\\tufts.ad.tufts.edu\netlogon\KIX32.EXE
&Run
Cancel
This file is in a location outside your local network. Files from locations you don’t recognize can harm your PC. Only run this file if you trust the location. <A>What’s the risk?</A>

>>>>>>>>>>>( Hidden Window Text )<<<<<<<<<<<
Publisher:
&Save
Al&ways ask before opening this file

>>>>( TitleMatchMode=slow Visible Text )<<<<

>>>>( TitleMatchMode=slow Hidden Text )<<<<
*/

/*
IfWinExist, Untitled - Notepad
{
    WinActivate  ; Automatically uses the window found above.
    WinMaximize  ; same
    Send, Some text.{Enter}
    return
}

if WinExist("ahk_class Notepad") or WinExist("ahk_class" . ClassName)
    WinActivate  ; Uses the last found window.

MsgBox % "The active window's ID is " . WinExist("A")
*/

Loop
{
Sleep, 100
;MsgBox running
if WinExist("ahk_class #32770")
	{
	;MsgBox, Window found
	WinActivate
	WinGetText, windowText
	Needle = KIX32.EXE
	IfInString, windowText, %Needle%
		{
		;MsgBox, The string was found.
		;Sleep, 100
		WinActivate
		Send, {LEFT}{ENTER}
		;Send, {ALT}{R}
		;return
		}
	}
	
}
	 