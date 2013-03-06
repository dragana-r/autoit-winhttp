

#include "WinHttp.au3"
#include <Array.au3>

Opt("MustDeclareVars", 1)

; Cracking URL
Global $aUrl = _WinHttpCrackUrl("http://www.autoitscript.com/forum/index.php?showforum=9")
_ArrayDisplay($aUrl, "_WinHttpCrackUrl()")
