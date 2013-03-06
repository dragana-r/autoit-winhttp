

#include "WinHttp.au3"

Opt("MustDeclareVars", 1)

; Creating URL out of array of components:
Global $aURL[8] = ["http", 1, "www.autoitscript.com", 80, "Jon", "deadPiXels", "admin.php"]
MsgBox(0, "Created URL", _WinHttpCreateUrl($aURL))
