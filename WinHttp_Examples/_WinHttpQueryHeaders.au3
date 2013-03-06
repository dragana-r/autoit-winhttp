

#include "WinHttp.au3"

Opt("MustDeclareVars", 1)

; Open needed handles
Global $hOpen = _WinHttpOpen()
Global $hConnect = _WinHttpConnect($hOpen, "msdn.microsoft.com")
; Specify the reguest:
Global $hRequest = _WinHttpOpenRequest($hConnect, Default, "en-us/library/aa384101(VS.85).aspx")

; Send request
_WinHttpSendRequest($hRequest)

; Wait for the response
_WinHttpReceiveResponse($hRequest)

; Get full header
Global $sHeader = _WinHttpQueryHeaders($hRequest)

; Close handles
_WinHttpCloseHandle($hRequest)
_WinHttpCloseHandle($hConnect)
_WinHttpCloseHandle($hOpen)

; Display retrieved header
MsgBox(0, "Header", $sHeader)