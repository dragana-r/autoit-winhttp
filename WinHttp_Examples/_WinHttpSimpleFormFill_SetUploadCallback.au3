
#include "WinHTTP.au3"

$sAddress = "https://posttestserver.com/post.php?dir=WinHttp" ; the address of the target (https or http, makes no difference - handled automatically)

; Select some file
$sFileToUpload = FileOpenDialog("Select file to upload...", "", "All Files (*)")
If Not $sFileToUpload Then Exit 5 ; check if the file is selected and exit if not

$sForm = _
		'<form action="' & $sAddress & '" method="post" enctype="multipart/form-data">' & _
		' <input type="file" name="upload"/>' & _
		'</form>'

; Initialize and get session handle
$hOpen = _WinHttpOpen()

$hConnect = $sForm ; will pass form as string so this is for coding correctness because $hConnect goes in byref

; Creates progress bar window
ProgressOn("UPLOADING", $sFileToUpload, "0%")

; Register callback function
_WinHttpSimpleFormFill_SetUploadCallback(UploadCallback)

; Fill form
$sHTML = _WinHttpSimpleFormFill($hConnect, $hOpen, _
		Default, _
		"name:upload", $sFileToUpload)
; Collect error number
$iErr = @error

; Unregister callback function
_WinHttpSimpleFormFill_SetUploadCallback(0)

; Kill progress bar window
ProgressOff()

; Close handles
_WinHttpCloseHandle($hConnect)
_WinHttpCloseHandle($hOpen)

If $iErr Then
	MsgBox(4096, "Error", "Error number = " & $iErr)
Else
	ConsoleWrite($sHTML & @CRLF)
	MsgBox(4096, "Success", $sHTML)
EndIf



; Callback function. For example, update progress control
Func UploadCallback($iPrecent)
	If $iPrecent = 100 Then
		ProgressSet(100, "Done", "Complete")
		Sleep(800) ; give some time for the progress bar to fill-in completely
	Else
		ProgressSet($iPrecent, $iPrecent & "%")
	EndIf
EndFunc
