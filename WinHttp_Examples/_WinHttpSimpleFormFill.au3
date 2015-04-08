
#include "WinHttp.au3"

Opt("MustDeclareVars", 1)

Global $hOpen, $hConnect
Global $sRead, $hFileHTM, $sFileHTM = @ScriptDir & "\Form.htm"

; Example1:
; 1. Open APNIC whois page (http://wq.apnic.net/apnic-bin/whois.pl)
; 2. Fill form on that page with these values/conditins:
; - fill default form
; - set ip address 4.2.2.2 to input box. Use name propery to locate input
; - send form (click button)
; - gather returned data

; Initialize and get session handle
$hOpen = _WinHttpOpen()
; Get connection handle
$hConnect = _WinHttpConnect($hOpen, "wq.apnic.net")
; Fill form on this page
$sRead = _WinHttpSimpleFormFill($hConnect, "apnic-bin/whois.pl", Default, "name:searchtext", "4.2.2.2", "name:object_type", "All")
; Close connection handle
_WinHttpCloseHandle($hConnect)
; Close session handle
_WinHttpCloseHandle($hOpen)

; See what's returned (in default browser or whatever)
If $sRead Then
    MsgBox(64 + 262144, "Done!", "Will open returned page in your default browser now.")
    $hFileHTM = FileOpen($sFileHTM, 2)
    FileWrite($hFileHTM, $sRead)
    FileClose($hFileHTM)
    ShellExecuteWait($sFileHTM)
EndIf
