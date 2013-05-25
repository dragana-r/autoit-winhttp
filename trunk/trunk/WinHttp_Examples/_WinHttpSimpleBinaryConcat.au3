

#include "WinHttp.au3"

Opt("MustDeclareVars", 1)

Global $sHost = "images.fanpop.com"
Global $sTarget = "images/image_uploads/Hello-Kitty-hello-kitty-181509_1024_768.jpg"
Global $sDestination = @ScriptDir & "\HelloKitty.jpg"

; Initialize and get session handle
Global $hHttpOpen = _WinHttpOpen()
If @error Then
	MsgBox(48, "Error", "Error initializing the usage of WinHTTP functions.")
	Exit 1
EndIf
; Get connection handle
Global $hHttpConnect = _WinHttpConnect($hHttpOpen, $sHost)
If @error Then
	MsgBox(48, "Error", "Error specifying the initial target server of an HTTP request.")
	_WinHttpCloseHandle($hHttpOpen)
	Exit 2
EndIf
; Specify the reguest
Global $hHttpRequest = _WinHttpOpenRequest($hHttpConnect, Default, $sTarget)
If @error Then
	MsgBox(48, "Error", "Error creating an HTTP request handle.")
	_WinHttpCloseHandle($hHttpConnect)
	_WinHttpCloseHandle($hHttpOpen)
	Exit 3
EndIf
; Send request
_WinHttpSendRequest($hHttpRequest)
If @error Then
	MsgBox(48, "Error", "Error sending specified request.")
	_WinHttpCloseHandle($hHttpConnect)
	_WinHttpCloseHandle($hHttpOpen)
	Exit 4
EndIf

; Wait for the response
_WinHttpReceiveResponse($hHttpRequest)
; Read if available
Global $bChunk, $bData, $hFile
If _WinHttpQueryDataAvailable($hHttpRequest) Then
	While 1
		$bChunk = _WinHttpReadData($hHttpRequest, 2) ; read binary
		If @error Then ExitLoop
		$bData = _WinHttpSimpleBinaryConcat($bData, $bChunk) ; concat two binary data
	WEnd
    ; Save it to the file
	$hFile = FileOpen($sDestination, 26)
	FileWrite($hFile, $bData)
	FileClose($hFile)
Else
	MsgBox(48, "Error occurred", "No data available. " & @CRLF)
EndIf

; Close handles
_WinHttpCloseHandle($hHttpRequest)
_WinHttpCloseHandle($hHttpConnect)
_WinHttpCloseHandle($hHttpOpen)

; See what's downloaded
ShellExecute($sDestination)