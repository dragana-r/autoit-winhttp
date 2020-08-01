#include "WinHttp.au3"

; Example 5:
; 1. Open paratus.hr form-action page (https://paratus.hr/software/testing/html5form/)
; 2. Fill form on that page with these values/conditins:
; - form is default one
; - set -User- and -Password- data to input boxes. Locate input boxes by their names -loginName- and -loginPass-
; - click third button (register action)

; Initialize and get session handle
$hOpen = _WinHttpOpen()
; Get connection handle
$hConnect = _WinHttpConnect($hOpen, "https://paratus.hr")
; Fill form on this page
$aRead = _WinHttpSimpleFormFill($hConnect, "/software/testing/html5form/", _
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
