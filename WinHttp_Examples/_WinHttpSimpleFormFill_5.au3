#include "WinHttp.au3"

; Example 5:
; 1. Open try.coderlearner.com form-action page (http://try.coderlearner.com/html5/form/form_formaction_ex_2.html)
; 2. Fill form on that page with these values/conditins:
; - form is default one
; - set -User- and -Password- data to input boxes. Locate input boxes by their names -loginName- and -loginPass-
; - click third button (register action)

; Initialize and get session handle
$hOpen = _WinHttpOpen()
; Get connection handle
$hConnect = _WinHttpConnect($hOpen, "http://try.coderlearner.com")
; Fill form on this page
$aRead = _WinHttpSimpleFormFill($hConnect, "html5/form/form_formaction_ex_2.html", _
		Default, _
		"name:loginName", "User", _
		"name:loginPass", "Password", _
		"type:submit", 2 _ ; third button (zero-based counting scheme)
		)
; Close connection handle
_WinHttpCloseHandle($hConnect)
; Close session handle
_WinHttpCloseHandle($hOpen)

MsgBox(4096, "Returned", $aRead)
