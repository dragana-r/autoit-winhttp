
#include "WinHttp.au3"

Opt("MustDeclareVars", 1)

Dim $hOpen, $hConnect
Dim $sRead

; Example 3:
; 1. Open cs.tut.fi forms page (http://www.cs.tut.fi/~jkorpela/forms/testing.html)
; 2. Fill form on that page with these values/conditins:
; - form is to be identified by its index, It's first form on the page, i.e. index is 0
; - set -Johnny B. Goode- data to textarea. Locate it by its name -Comments-.
; - check the checkbox. Locate it by name -box-. Checked value is -yes-.
; - set -This is hidden, so what?- data to input field identified by name -hidden field-.
; - gather data

; Initialize and get session handle
$hOpen = _WinHttpOpen()
; Get connection handle
$hConnect = _WinHttpConnect($hOpen, "www.cs.tut.fi")
; Fill form on this page
$sRead = _WinHttpSimpleFormFill($hConnect, _ ; connection handle
		"~jkorpela/forms/testing.html", _ ; target page
		"index:0", _ ; form identifier
		"name:Comments", "Johnny B. Goode", _ ; first field identifier paired with field data
		"name:box", "yes", _ ; second field identifier paired with data
		"name:hidden field", "This is hidden, so what?") ; third field identifier paired with data
; Close connection handle
_WinHttpCloseHandle($hConnect)
; Close session handle now that's no longer needed
_WinHttpCloseHandle($hOpen)

If $sRead Then
	ConsoleWrite($sRead & @CRLF)
	MsgBox(64 + 262144, "Web Page says", $sRead)
EndIf
