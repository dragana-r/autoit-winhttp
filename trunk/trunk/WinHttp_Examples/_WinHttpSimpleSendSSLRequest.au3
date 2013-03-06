

#include "WinHttp.au3"

Opt("MustDeclareVars", 1)

; Windows Live ID. JavaScript required to sign in, remember that.
Global $sDomain = "accountservices.passport.net"
Global $sPage = "uiresetpw.srf"

; Initialize and get session handle
Global $hOpen = _WinHttpOpen()
; Get connection handle
Global $hConnect = _WinHttpConnect($hOpen, $sDomain)

; Make a SimpleSSL request
Global $hRequestSSL = _WinHttpSimpleSendSSLRequest($hConnect, Default, $sPage)

; Read...
Global $sReturned = _WinHttpSimpleReadData($hRequestSSL)
; Close handles
_WinHttpCloseHandle($hRequestSSL)
_WinHttpCloseHandle($hConnect)
_WinHttpCloseHandle($hOpen)

; See what's returned
ConsoleWrite($sReturned & @CRLF)
MsgBox(64 + 262144, "Done", "Page source is printed to console")