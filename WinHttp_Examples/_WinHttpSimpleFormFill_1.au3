

#include "WinHttp.au3"

Opt("MustDeclareVars", 1)

Global $hOpen, $hConnect
Global $sRead, $hFileHTM, $sFileHTM = @ScriptDir & "\Form.htm"

; Example 2:
; 1. Open w3schools forms page (http://www.w3schools.com/html/html_forms.asp)
; 2. Fill form on that page with these values/conditins:
; - form is to be identified by its index -0-
; - set -Miyake- and -Issey- data to input boxes. Locate input boxes by their names -firstname- and -lastname-

; Initialize and get session handle
$hOpen = _WinHttpOpen()
; Get connection handle
$hConnect = _WinHttpConnect($hOpen, "w3schools.com")
; Fill form on this page
$sRead = _WinHttpSimpleFormFill($hConnect, "html/html_forms.asp", "index:0", "name:firstname", "Miyake", "name:lastname", "Issey")
; Close connection handle
_WinHttpCloseHandle($hConnect)
; Close session handle
_WinHttpCloseHandle($hOpen)

If $sRead Then
    MsgBox(64 + 262144, "Done!", "Will open returned page in your default browser now." & @CRLF & _
            "You should see 'Miyake Issey' somewhere on that page.")
    $hFileHTM = FileOpen($sFileHTM, 2)
    FileWrite($hFileHTM, $sRead)
    FileClose($hFileHTM)
    ShellExecuteWait($sFileHTM)
EndIf