

#include "WinHttp.au3"
#include <Array.au3>

Opt("MustDeclareVars", 1)

; Internet Explorer proxy configuration for the current user:
Global $aIEproxy = _WinHttpGetIEProxyConfigForCurrentUser()
_ArrayDisplay($aIEproxy, "Internet Explorer proxy configuration")