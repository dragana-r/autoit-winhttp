
#include "WinHttp.au3"

Opt("MustDeclareVars", 1)

Global $hOpen, $hConnect
Global $sRead

; Example 4:
; 1. Open google accounts login page (https://accounts.google.com/ServiceLoginAuth)
; 2. Fill form on that page with these values/conditins:
; - form is to be identified by its id. Id is -gaia_loginform-
; - set -MyUserName- data to user-name input box. Locate input box by its id -Email-
; - set -MyPassword- data to password input box. Locate input box by its id -Passwd-
; - set -js_enabled- data to hidden field. Locate field by its id -bgresponse-
; - gather data

; Initialize and get session handle
$hOpen = _WinHttpOpen("Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.2; Trident/6.0)")
; Get connection handle
$hConnect = _WinHttpConnect($hOpen, "accounts.google.com")
; Fill form on this page
$sRead = _WinHttpSimpleFormFill($hConnect, "ServiceLogin", "gaia_loginform", "Email", "MyUserName", "Passwd", "MyPassword", "bgresponse", "js_enabled")
;Close connection handle
_WinHttpCloseHandle($hConnect)
; Close session handle
_WinHttpCloseHandle($hOpen)

;Print returned:
ConsoleWrite($sRead & @CRLF)
MsgBox(262144, "The End", "Source of the last example is printed to console." & @CRLF & _
 "In case valid login data was passed it will be user's page on google accounts")
