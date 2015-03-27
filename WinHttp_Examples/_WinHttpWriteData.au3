

#include "WinHttp.au3"

Opt("MustDeclareVars", 1)

; Initialize and get session handle
Global $hOpen = _WinHttpOpen()
; Get connection handle
Global $hConnect = _WinHttpConnect($hOpen, "httpbin.org")
; Specify the reguest
Global $hRequest = _WinHttpOpenRequest($hConnect, "POST", "/post")

Global $sPostData = "Additional data to send"
; Send request
_WinHttpSendRequest($hRequest, Default, Default, StringLen($sPostData))

; Write additional data to send
_WinHttpWriteData($hRequest, $sPostData)

; Wait for the response
_WinHttpReceiveResponse($hRequest)

; Check if there is data available...
If _WinHttpQueryDataAvailable($hRequest) Then
    MsgBox(64, "OK", _WinHttpReadData($hRequest))
Else
    MsgBox(48, "Error", "Site is experiencing problems (or you).")
EndIf

; Close handles
_WinHttpCloseHandle($hRequest)
_WinHttpCloseHandle($hConnect)
_WinHttpCloseHandle($hOpen)
