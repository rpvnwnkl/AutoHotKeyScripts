#U::
month = 10
day = 23
year = 2017

WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 
;select month window
MouseClick, left,  460,  670
Sleep, 100

Send, %month%
Send, {TAB}
Send, %day%
Send, {TAB}
Send, %year%
Send, {ENTER}
Send, {F8}
/*clipboard =  ; Start off empty to allow ClipWait to detect when the text has arrived.
Send ^c
ClipWait  ; Wait for the clipboard to contain text.
Clip1 = %clipboard%;
Send ^c
ClipWait ;
Clip2 = %Clip1% and %clipboard% ;
MsgBox Control-C copied the following contents to the clipboard:`n`n%Clip2%

*/
return