#t::
clipboard =  ; Start off empty to allow ClipWait to detect when the text has arrived.
Send ^c
ClipWait  ; Wait for the clipboard to contain text.
Clip1 = %clipboard%;
Send ^c
ClipWait ;
Clip2 = %Clip1% and %clipboard% ;
MsgBox Control-C copied the following contents to the clipboard:`n`n%Clip2%
return