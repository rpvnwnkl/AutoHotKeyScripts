
;Context: You are in the TEL window and want to check whether a number \
;           is a cell or a landline
;	 
;	
;Instructions: have advance open with cursor in the area code field and
;              
;		Press windows key and the letter C
;       Pressing Windows Key and Letter A will reset in case of error
#A::
Reload
Return

#C::
SetTitleMatchMode, 2

UrlDownloadToFile, "https://lookups.twilio.com/v1/PhoneNumbers/617548-5907?CountryCode=US&Type=carrier&Type=caller-name", realVar

MsgBox, %realVar%


Return
