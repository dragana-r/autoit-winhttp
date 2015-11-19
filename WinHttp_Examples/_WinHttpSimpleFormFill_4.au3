#include <Array.au3>
#include "WinHttp.au3"

; Example 5:
; 1. Open www.cs.tut.fi forms page (https://www.cs.tut.fi/~jkorpela/forms/testing.html)
; 2. Fill form on that page with these values/conditins:
; - form is to be identified by its index -1-
; - set -Something- and -this script- data to input boxes. Locate input boxes by their names -textfield- and -filefield-
; 2. Return array

; Initialize and get session handle
$hOpen = _WinHttpOpen()
; Get connection handle
$hConnect = _WinHttpConnect($hOpen, "https://www.cs.tut.fi")
; Fill form on this page
$aRead = _WinHttpSimpleFormFill($hConnect, "~jkorpela/forms/testing.html", _
		"index:1", _
		"name:textfield", "Testing file upload", _
		"name:filefield", @ScriptFullPath, _
		"[RETURN_ARRAY]")

; Close connection handle
_WinHttpCloseHandle($hConnect)
; Close session handle
_WinHttpCloseHandle($hOpen)

_ArrayDisplay($aRead)
