

#include "WinHttp.au3"

Opt("MustDeclareVars", 1)

; This could take some time
Global $sPACurl = _WinHttpDetectAutoProxyConfigUrl(BitOR($WINHTTP_AUTO_DETECT_TYPE_DHCP, $WINHTTP_AUTO_DETECT_TYPE_DNS_A))

MsgBox(64, "_WinHttpDetectAutoProxyConfigUrl", "The URL for the Proxy Auto-Configuration (PAC) file is: " & $sPACurl)