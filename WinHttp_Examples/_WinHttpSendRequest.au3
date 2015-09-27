

#include "WinHttp.au3"

; http://www.w3schools.com/php/demo_form_validation_escapechar.php

$sUserName = "SomeUserName"
$sEmail = "some.email@something.com"

$sDomain = "www.w3schools.com"
$sPage = "/php/demo_form_validation_escapechar.php"

; Data to send
$sAdditionalData = "name=" & $sUserName & "&email=" & $sEmail

; Initialize and get session handle
$hOpen = _WinHttpOpen("Mozilla/5.0 (Windows NT 6.1; WOW64; rv:40.0) Gecko/20100101 Firefox/40.0")

; Get connection handle
$hConnect = _WinHttpConnect($hOpen, $sDomain)

; Make a request
$hRequest = _WinHttpOpenRequest($hConnect, "POST", $sPage)

; Send it. Specify additional data to send too. This is required by the Google API:
_WinHttpSendRequest($hRequest, "Content-Type: application/x-www-form-urlencoded", $sAdditionalData)

; Wait for the response
_WinHttpReceiveResponse($hRequest)

; See what's returned
Dim $sReturned
If _WinHttpQueryDataAvailable($hRequest) Then ; if there is data
	Do
		$sReturned &= _WinHttpReadData($hRequest)
	Until @error
EndIf

; Close handles
_WinHttpCloseHandle($hRequest)
_WinHttpCloseHandle($hConnect)
_WinHttpCloseHandle($hOpen)

; See what's returned
MsgBox(4096, "Returned", $sReturned)
ConsoleWrite($sReturned & @CRLF)