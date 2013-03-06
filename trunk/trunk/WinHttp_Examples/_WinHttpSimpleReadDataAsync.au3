

#include "WinHttp.au3"

Opt("MustDeclareVars", 1)

;XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
;   Reserve memory space for asynchronous work
;XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
; The size
Global Const $iSizeBufferAsync = 1048576 ; 1MB for example, should be enough
; The buffer
Global Static $tBufferAsync = DllStructCreate("byte[" & $iSizeBufferAsync & "]")
; Get pointer to this memory space
Global $pBufferAsync = DllStructGetPtr($tBufferAsync)



; Register Callback function
Global $hWINHTTP_STATUS_CALLBACK = DllCallbackRegister("__WINHTTP_STATUS_CALLBACK", "none", "handle;dword_ptr;dword;ptr;dword")

; Initialize and get session handle. Asynchronous flag.
Global $hOpen = _WinHttpOpen(Default, Default, Default, Default, $WINHTTP_FLAG_ASYNC)

; Assign callback function
_WinHttpSetStatusCallback($hOpen, $hWINHTTP_STATUS_CALLBACK)

; Get connection handle
Global $hConnect = _WinHttpConnect($hOpen, "msdn.microsoft.com")

; Make request
Global $hRequest = _WinHttpOpenRequest($hConnect, Default, "en-us/library/aa383917(v=vs.85).aspx")

; Send it
_WinHttpSendRequest($hRequest)

; Some dummy code for waiting
MsgBox(64 + 262144, "Wait...", "Wait for the results if they are not shown already.")


; Close handles
_WinHttpCloseHandle($hRequest)
_WinHttpCloseHandle($hConnect)
_WinHttpCloseHandle($hOpen)
; Free callback. Redundant here
DllCallbackFree($hWINHTTP_STATUS_CALLBACK)



;XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
;         Define callback function
;XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Func __WINHTTP_STATUS_CALLBACK($hInternet, $iContext, $iInternetStatus, $pStatusInformation, $iStatusInformationLength)
	#forceref $hInternet, $iContext, $pStatusInformation, $iStatusInformationLength
	ConsoleWrite(">> ")
	; Interpret the status
	Local $sStatus
	Switch $iInternetStatus
		Case $WINHTTP_CALLBACK_STATUS_CLOSING_CONNECTION
			$sStatus = "Closing the connection to the server"
		Case $WINHTTP_CALLBACK_STATUS_CONNECTED_TO_SERVER
			$sStatus = "Successfully connected to the server."
		Case $WINHTTP_CALLBACK_STATUS_CONNECTING_TO_SERVER
			$sStatus = "Connecting to the server."
		Case $WINHTTP_CALLBACK_STATUS_CONNECTION_CLOSED
			$sStatus = "Successfully closed the connection to the server."
		Case $WINHTTP_CALLBACK_STATUS_DATA_AVAILABLE
			$sStatus = "Data is available to be retrieved with WinHttpReadData."
			ConsoleWrite($sStatus & @CRLF)

			;*************************************
			;        Read asynchronously
			;*************************************
			_WinHttpSimpleReadDataAsync($hInternet, $pBufferAsync, $iSizeBufferAsync)
			Return

		Case $WINHTTP_CALLBACK_STATUS_HANDLE_CREATED
			$sStatus = "An HINTERNET handle has been created: " & $hInternet
		Case $WINHTTP_CALLBACK_STATUS_HANDLE_CLOSING
			$sStatus = "This handle value has been terminated: " & $hInternet
		Case $WINHTTP_CALLBACK_STATUS_HEADERS_AVAILABLE
			$sStatus = "The response header has been received and is available with WinHttpQueryHeaders."
			ConsoleWrite($sStatus & @CRLF)

			;*************************************
			;          Print header
			;*************************************
			ConsoleWrite(_WinHttpQueryHeaders($hInternet) & @CRLF)

			;*************************************
			; Check if there is any data available
			;*************************************
			_WinHttpQueryDataAvailable($hInternet)

			Return

		Case $WINHTTP_CALLBACK_STATUS_INTERMEDIATE_RESPONSE
			$sStatus = "Received an intermediate (100 level) status code message from the server."
		Case $WINHTTP_CALLBACK_STATUS_NAME_RESOLVED
			$sStatus = "Successfully found the IP address of the server."
		Case $WINHTTP_CALLBACK_STATUS_READ_COMPLETE
			$sStatus = "Data was successfully read from the server."
			ConsoleWrite($sStatus & @CRLF)

			;*************************************
			;          Print read data
			;*************************************
			Local $sRead = DllStructGetData(DllStructCreate("char[" & $iStatusInformationLength & "]", $pStatusInformation), 1)
			ConsoleWrite($sRead & @CRLF)

			Return

		Case $WINHTTP_CALLBACK_STATUS_RECEIVING_RESPONSE
			$sStatus = "Waiting for the server to respond to a request."
		Case $WINHTTP_CALLBACK_STATUS_REDIRECT
			$sStatus = "An HTTP request is about to automatically redirect the request."
		Case $WINHTTP_CALLBACK_STATUS_REQUEST_ERROR
			$sStatus = "An error occurred while sending an HTTP request."
		Case $WINHTTP_CALLBACK_STATUS_REQUEST_SENT
			$sStatus = "Successfully sent the information request to the server."
		Case $WINHTTP_CALLBACK_STATUS_RESOLVING_NAME
			$sStatus = "Looking up the IP address of a server name."
		Case $WINHTTP_CALLBACK_STATUS_RESPONSE_RECEIVED
			$sStatus = "Successfully received a response from the server."
		Case $WINHTTP_CALLBACK_STATUS_SECURE_FAILURE
			$sStatus = "One or more errors were encountered while retrieving a Secure Sockets Layer (SSL) certificate from the server."
		Case $WINHTTP_CALLBACK_STATUS_SENDING_REQUEST
			$sStatus = "Sending the information request to the server."
		Case $WINHTTP_CALLBACK_STATUS_SENDREQUEST_COMPLETE
			$sStatus = "The request completed successfully."
			ConsoleWrite($sStatus & @CRLF)

			;*************************************
			;          Receive Response
			;*************************************
			_WinHttpReceiveResponse($hInternet)
			Return

		Case $WINHTTP_CALLBACK_STATUS_WRITE_COMPLETE
			$sStatus = "Data was successfully written to the server."
	EndSwitch
	; Print it
	ConsoleWrite($sStatus & @CRLF)
EndFunc   ;==>__WINHTTP_STATUS_CALLBACK