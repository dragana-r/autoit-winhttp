

#include "WinHttp.au3"

Opt("MustDeclareVars", 1)

If Not _WinHttpCheckPlatform() Then
	MsgBox(48, "Caution", "WinHTTP not available on your system!")
	Exit 1
EndIf

;... The rest of the code
