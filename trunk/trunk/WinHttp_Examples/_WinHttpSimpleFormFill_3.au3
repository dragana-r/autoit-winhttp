
#include "WinHttp.au3"

Opt("MustDeclareVars", 1)

Global $hOpen, $hConnect
Global $sRead

; Example 4:
; 1. Open yahoo mail login page (https://login.yahoo.com/config/login_verify2?&.src=ym)
; 2. Fill form on that page with these values/conditins:
; - form is to be identifide by its name, Name is -login_form-
; - set -MyUserName- data to user-name input box. Locate input box by its Id -username-
; - set -MyPassword- data to password input box. Locate input box by its Id -passwd-
; - gather data

; Initialize and get session handle
$hOpen = _WinHttpOpen()
; Get connection handle
$hConnect = _WinHttpConnect($hOpen, "login.yahoo.com")
; Fill form on this page
$sRead = _WinHttpSimpleFormFill($hConnect, "config/login_verify2?&.src=ym", "name:login_form", "username", "MyUserName", "passwd", "MyPassword")
;Close connection handle
_WinHttpCloseHandle($hConnect)
; Close session handle
_WinHttpCloseHandle($hOpen)

;Print returned:
ConsoleWrite($sRead & @CRLF)
MsgBox(262144, "The End", "Source of the last example is printed to console." & @CRLF & _
 "In case valid login data was passed it will be user's mail page on yahoo.mail")
