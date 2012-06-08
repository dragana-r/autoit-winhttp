
; Converting help file pages to origo pages.
; They are placed into the folder from wich to upload to the server

Global $sOrigoHTMLFolder = @ScriptDir & "\OrigoHTML" ;<- this folder
DirCreate($sOrigoHTMLFolder)
FileDelete($sOrigoHTMLFolder)

Global Const $sPrefix = "X" & StringReplace(_GetVersionNumber(@ScriptDir & "\WinHttp.au3"), ".", "_")

Global $sSourceHTMLFolder = @ScriptDir & "\WinHttp_Help\HTML\Functions"
Global $sSourceCSSFolder = @ScriptDir & "\WinHttp_Help\HTML\CSS"
Global $sSourceIMGFolder = @ScriptDir & "\WinHttp_Help\HTML\Images"
Global $sHomePageFolder = @ScriptDir & "\WinHttp_Help\HTML"

FileCopy($sSourceHTMLFolder & "\*.htm", $sOrigoHTMLFolder)
FileCopy($sSourceCSSFolder & "\Default1.css", $sOrigoHTMLFolder)
FileCopy($sSourceIMGFolder & "\*.*", $sOrigoHTMLFolder)
FileCopy($sHomePageFolder & "\CHM_HomePage.htm", $sOrigoHTMLFolder)

_AdjustFiles($sOrigoHTMLFolder)

Global $sToPost = @CRLF & '==[http://winhttp.origo.ethz.ch/system/files/' & $sPrefix & 'Functions.htm Available Functions]==' & @CRLF & _
		'[http://winhttp.origo.ethz.ch/system/files/' & $sPrefix & $sPrefix & 'CHM_HomePage.htm Small intro...]' & @CRLF & _
		_GetEditString()

ConsoleWrite($sToPost & @CRLF)
ClipPut($sToPost)
ConsoleWrite("> All done!" & @CRLF)
MsgBox(64 + 262144, "Done", "All done. Pages made!" & @CRLF & "Origo post data is put to clipboard")

Func _AdjustFiles($sOrigoHTMLFolder)
	Local $sFileOR, $sFile, $hFile, $sData
	$sFile = $sOrigoHTMLFolder & "\Default1.css"
	$sData = FileRead($sFile)
	$sData = StringReplace($sData, "url('../images/gradient_1024x24.jpg')", "url('gradient_1024x24.jpg')")
	$hFile = FileOpen($sFile, 2)
	FileWrite($hFile, $sData)
	FileClose($hFile)
	Local $hSearch = FileFindFirstFile($sOrigoHTMLFolder & "\*.htm")
	If $hSearch = -1 Then Return 1
	While 1
		$sFile = FileFindNextFile($hSearch)
		If @error Then ExitLoop
		If @extended Then ContinueLoop
		$sFileOR = $sOrigoHTMLFolder & "\" & $sFile
		$sData = FileRead($sFileOR)
		$sData = StringReplace($sData, '<link href="../CSS/Default1.css" rel="stylesheet" type="text/css">', '<link href="Default1.css" rel="stylesheet" type="text/css">')
		$sData = StringReplace($sData, '<link href="CSS/Default1.css" rel="stylesheet" type="text/css">', '<link href="Default1.css" rel="stylesheet" type="text/css">')
		$sData = StringReplace($sData, '<img src="../Images/gradient_1024x24.jpg" width="100%" height="1" alt="">', '<img src="gradient_1024x24.jpg" width="100%" height="1" alt="">')
		$sData = StringReplace($sData, '<img src="Images/gradient_1024x24.jpg" width="100%" height="1" alt="">', '<img src="gradient_1024x24.jpg" width="100%" height="1" alt="">')
		$sData = StringReplace($sData, '<img src="Images/WinHttp.png" width="400" height="100" border="0" alt="">', '<img src="WinHttp.png" width="400" height="100" border="0" alt="">')
		$sData = StringRegExpReplace($sData, '"\h*(_WinHttp.*?\.htm\h*")', '"' & $sPrefix & "$1")
		FileDelete($sFileOR)
		$sFile = $sOrigoHTMLFolder & "\" & $sPrefix & $sFile
		$hFile = FileOpen($sFile, 10)
		FileWrite($hFile, $sData)
		FileClose($hFile)
	WEnd
	FileClose($hSearch)
	Return 1
EndFunc   ;==>_AdjustFiles

Func _GetEditString()
	Local $sString = FileRead(@ScriptDir & "\WinHttp.au3")
	Local $aArray = StringRegExp($sString, "(?si); #CURRENT# ;*=*.*?\v+;(.*?)\v+;\h*=", 3)
	Local $aFuncs = StringSplit($aArray[0], @CRLF & ";", 3)
	Local $sOut
	For $i = 0 To UBound($aFuncs) - 1
		$sOut &= "* [http://winhttp.origo.ethz.ch/system/files/" & $sPrefix & $aFuncs[$i] & ".htm " & $aFuncs[$i] & "]" & @CRLF
	Next
	Return $sOut
EndFunc   ;==>_GetEditString

Func _GetVersionNumber($sFile)
	Local $sVer = FileRead($sFile, 1024)
	Local $aArray = StringRegExp($sVer, "(?s).*?\Q; File Version.........:\E\s*(.*?)\v+", 3)
	If @error Then Return SetError(1, 0, "")
	Return $aArray[0]
EndFunc   ;==>_GetVersionNumber