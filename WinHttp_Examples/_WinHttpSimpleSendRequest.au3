

#include "WinHttp.au3"

Opt("MustDeclareVars", 1)

; Initialize and get session handle
Global $hOpen = _WinHttpOpen()
; Get connection handle
Global $hConnect = _WinHttpConnect($hOpen, "w3schools.com")
; Make a request
Global $hRequest = _WinHttpSimpleSendRequest($hConnect, Default, "tags/tag_input.asp")

If $hRequest Then
	; Simple-read...
	ConsoleWrite(_WinHttpSimpleReadData($hRequest) & @CRLF)
	MsgBox(64, "Okey do!", "Returned source is print to concole. Check it.")
Else
	MsgBox(48, "Error", "Error ocurred for _WinHttpSimpleSendRequest, Error number is " & @error)
EndIf

; Close handles
_WinHttpCloseHandle($hRequest)
_WinHttpCloseHandle($hConnect)
_WinHttpCloseHandle($hOpen)