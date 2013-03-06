

#include "WinHttp.au3"

Opt("MustDeclareVars", 1)

; Initialize and get session handle
Global $hOpen = _WinHttpOpen()
If @error Then
	MsgBox(48, "Error", "Error initializing the usage of WinHTTP functions.")
	Exit 1
EndIf

; Close it
_WinHttpCloseHandle($hOpen)