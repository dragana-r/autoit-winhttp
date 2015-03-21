
#include "WinHttp.au3"

$sAddress = "https://posttestserver.com/post.php?dump&dir=WinHttp" ; the address of the target  (https or http, makes no difference - handled automatically)

$sFileToUpload = @ScriptFullPath ; upload itself

$sForm = _
		'<form action="' & $sAddress & '" method="post" enctype="multipart/form-data">' & _
		'    <input type="file" name="upload"/>' & _ ;
		'   <input type="text" name="someparam" />' & _
		'</form>'

; Initialize and get session handle
$hOpen = _WinHttpOpen()

$hConnect = $sForm ; will pass form as string so this is for coding correctness because $hConnect goes in byref

; Fill form
$sHTML = _WinHttpSimpleFormFill($hConnect, $hOpen, _
		Default, _
		"name:upload", $sFileToUpload, _
		"name:someparam", "Candy")

If @error Then
	MsgBox(4096, "Error", "Error number = " & @error)
Else
	ConsoleWrite($sHTML & @CRLF)
EndIf

; Close handles
_WinHttpCloseHandle($hConnect)
_WinHttpCloseHandle($hOpen)
