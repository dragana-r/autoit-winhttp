
#include "WinHttp.au3"

Opt("MustDeclareVars", 1)

Global $hOpen, $hConnect
Global $sRead, $hFileHTM, $sFileHTM = @ScriptDir & "\Form.htm"

; Example 2:
; 1. Open w3schools forms page (http://www.w3schools.com/html/html_forms.asp)
; 2. Fill form on that page with these values/conditins:
; - form is to be identifide by its name -input0-
; - set -OMG!!!- data to input box. Locate input box by its name -user-
; - gather data

; Initialize and get session handle
$hOpen = _WinHttpOpen()
; Get connection handle
$hConnect = _WinHttpConnect($hOpen, "w3schools.com")
; Fill form on this page
$sRead = _WinHttpSimpleFormFill($hConnect, "html/html_forms.asp", "name:input0", "name:user", "OMG!!!")
; Close connection handle
_WinHttpCloseHandle($hConnect)
; Close session handle
_WinHttpCloseHandle($hOpen)

If $sRead Then
    MsgBox(64 + 262144, "Done!", "Will open returned page in your default browser now." & @CRLF & _
            "You should see 'OMG!!!' or 'OMG%21%21%21' (encoded version) somewhere on that page.")
    $hFileHTM = FileOpen($sFileHTM, 2)
    FileWrite($hFileHTM, $sRead)
    FileClose($hFileHTM)
    ShellExecuteWait($sFileHTM)
EndIf
