

#include "WinHttp.au3"

Opt("MustDeclareVars", 1)

; !!!Note that this example will fail because of invalid username and password!!!

; Use real data for authentication
Global $sUserName = "SomeUserName"
Global $sPassword = "SomePassword"
Global $sDomain = "www.google.com"
Global $sPage = "accounts/ClientLogin"
; Visit http://code.google.com/apis/accounts/docs/AuthForInstalledApps.html for more informations
Global $sAdditionalData = "accountType=HOSTED_OR_GOOGLE&Email=" & $sUserName & "&Passwd=" & $sPassword & "&service=mail&source=Gulp-CalGulp-1.05"

; Initialize and get session handle
Global $hOpen = _WinHttpOpen()
; Get connection handle
Global $hConnect = _WinHttpConnect($hOpen, $sDomain)

; SimpleSSL-request it...
Global $sReturned = _WinHttpSimpleSSLRequest($hConnect, "POST", $sPage, Default, $sAdditionalData)

; Close handles
_WinHttpCloseHandle($hConnect)
_WinHttpCloseHandle($hOpen)

; See what's returned
MsgBox(0, "Returned", $sReturned)