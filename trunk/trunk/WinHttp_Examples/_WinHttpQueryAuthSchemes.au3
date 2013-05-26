
#include "WinHttp.au3"

Opt("MustDeclareVars", 1)

Global Const $sAddress = "192.168.1.1" ; some address that requires authorization

; Credentials (incorrect values here - the example will fail therefore!)
Global Const $sUserName = "user", $sPassword = "password"

; Login and print the result to the console
ConsoleWrite(_LoginExample($sAddress, $sUserName, $sPassword) & @CRLF)


; The function
Func _LoginExample($sAddress, $sUserName, $sPassword)
	Local $sOut ; this variable data will be returned

	; Initialize and get session handle
	Local $hOpen = _WinHttpOpen()
	If $hOpen Then
		; Get connection handle
		Local $hConnect = _WinHttpConnect($hOpen, $sAddress)
		If $hConnect Then
			; Specify the reguest
			Local $hRequest = _WinHttpOpenRequest($hConnect)
			If $hRequest Then
                ; Send the request
				_WinHttpSendRequest($hRequest)

				; Wait for the response
				_WinHttpReceiveResponse($hRequest)

				; Query status code
				Local $iStatusCode = _WinHttpQueryHeaders($hRequest, $WINHTTP_QUERY_STATUS_CODE)
				; Check status code
				If $iStatusCode = $HTTP_STATUS_DENIED Or $iStatusCode = $HTTP_STATUS_PROXY_AUTH_REQ Then
					; Query Authorization scheme
                    Local $iSupportedSchemes, $iFirstScheme, $iAuthTarget
					If _WinHttpQueryAuthSchemes($hRequest, $iSupportedSchemes, $iFirstScheme, $iAuthTarget) Then
						; Set passed credentials
						_WinHttpSetCredentials($hRequest, $iAuthTarget, $iFirstScheme, $sUserName, $sPassword)
						; Send request again now
						_WinHttpSendRequest($hRequest)
						; And wait for the response again
						_WinHttpReceiveResponse($hRequest)
					EndIf
				EndIf

				; Check if there is any data available and read if yes
				Local $sChunk
				If _WinHttpQueryDataAvailable($hRequest) Then
                    While 1
						$sChunk = _WinHttpReadData($hRequest)
						If @error Then ExitLoop
						$sOut &= $sChunk
					WEnd
				EndIf

				; Close handles when they are not needed any more
				_WinHttpCloseHandle($hRequest)
			EndIf
			_WinHttpCloseHandle($hConnect)
		EndIf
		_WinHttpCloseHandle($hOpen)
	EndIf

	; Return whatever is read
	Return $sOut
EndFunc
