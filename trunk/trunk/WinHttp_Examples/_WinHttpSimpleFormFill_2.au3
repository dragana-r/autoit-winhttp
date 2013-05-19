
#include "WinHttp.au3"

Opt("MustDeclareVars", 1)

Global $hOpen, $hConnect
Global $sRead, $hFileHTM, $sFileHTM = @ScriptDir & "\Form.htm"

; Example 3:
; 1. Open cs.tut.fi forms page (http://www.cs.tut.fi/~jkorpela/forms/testing.html)
; 2. Fill form on that page with these values/conditins:
; - form is to be identified by its index, It's first form on the page, i.e. index is 0
; - set -Johnny B. Goode- data to textarea. Locate it by its name -Comments-.
; - check the checkbox. Locate it by name -box-. Checked value is -yes-.
; - set -This is hidden, so what?- data to input field identified by name -hidden field-.
; - gather data

; Initialize and get session handle
$hOpen = _WinHttpOpen()
; Get connection handle
$hConnect = _WinHttpConnect($hOpen, "www.cs.tut.fi")
; Fill form on this page
$sRead = _WinHttpSimpleFormFill($hConnect, "~jkorpela/forms/testing.html", "index:0", "name:Comments", "Johnny B. Goode", "name:box", "yes", "name:hidden field", "This is hidden, so what?")
; Close connection handle
_WinHttpCloseHandle($hConnect)
; Close session handle now that's no longer needed
_WinHttpCloseHandle($hOpen)

If $sRead Then
    MsgBox(64 + 262144, "Done!", "Will open returned page in your default browser now." & @CRLF & _
            "It should show sent data.")
    $hFileHTM = FileOpen($sFileHTM, 2)
    FileWrite($hFileHTM, $sRead)
    FileClose($hFileHTM)
    ShellExecuteWait($sFileHTM)
EndIf
