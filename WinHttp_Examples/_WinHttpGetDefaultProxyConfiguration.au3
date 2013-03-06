

#include "WinHttp.au3"
#include <Array.au3>

Opt("MustDeclareVars", 1)

; Current WinHTTP proxy configuration:
Global $aProxy = _WinHttpGetDefaultProxyConfiguration()
_ArrayDisplay($aProxy, "Current WinHTTP proxy configuration")