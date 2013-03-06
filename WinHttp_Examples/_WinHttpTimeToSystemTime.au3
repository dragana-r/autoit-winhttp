

#include "WinHttp.au3"
#include <Array.au3>

Opt("MustDeclareVars", 1)

; Time
Global $aTime = _WinHttpTimeToSystemTime("Sat, 21 Aug 2010 22:04:43 GMT")
_ArrayDisplay($aTime, "_WinHttpTimeToSystemTime()")

