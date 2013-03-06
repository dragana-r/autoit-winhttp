

#include "WinHttp.au3"

Opt("MustDeclareVars", 1)

; Initialize and get session handle
Global $hOpen = _WinHttpOpen()
; Get connection handle
Global $hConnect = _WinHttpConnect($hOpen, "thetimes.co.uk")
; Request
Global $hRequest = _WinHttpSimpleSendRequest($hConnect)

; Simple-read...
Global $sRead = _WinHttpSimpleReadData($hRequest)
MsgBox(64, "Returned (first 1100 characters)", StringLeft($sRead, 1100) & "...")

; Close handles
_WinHttpCloseHandle($hRequest)
_WinHttpCloseHandle($hConnect)
_WinHttpCloseHandle($hOpen)