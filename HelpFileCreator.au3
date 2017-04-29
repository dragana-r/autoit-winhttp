
#pragma compile(Console, true)
#pragma compile(x64, false)



If @AutoItX64 Then
	ConsoleWriteError("!!!!!Unable to proceed!!!!!" & @CRLF)
	MsgBox(64, "32-bit  script!", "Wrong interpreter! Use 32bit AutoIt.")
	Exit -1
EndIf

Global Const $CHM_FOLDERSUFFIX = "_Help"

#XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
; These global variables define the help file
Global $sChangeLogFile = ""
Global $sHomeLink = "https://github.com/dragana-r/autoit-winhttp/releases"
Global $sHomePage = @ScriptDir & "\Extra\Home.htm"
Global $sLogoPic = @ScriptDir & "\Images\WinHttp.jpg"
Global $sCssFile = @ScriptDir & "\Extra\default1.css"
Global $sJSFile = @ScriptDir & "\JS\Script.js"
Global $sJSRFile = @ScriptDir & "\JS\Script_Right.js"
Global $sFile = @ScriptDir & "\WinHttp.au3"
Global $sExamplesFolder = @ScriptDir & "\WinHttp_Examples"
Global Const $sCurrentVersionNumber = _GetVersionNumber($sFile)
#XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

ConsoleWrite(@CRLF & "> Generating pages... " & @CRLF)

Global $aFunctions
Global $sWorkingFolder = _CHM_UDFToHTMPages($sFile, $aFunctions)
_CHM_WriteJScript($sWorkingFolder, $sFile, $sHomeLink)
_CHM_WriteDefaultCSS($sWorkingFolder)
_CHM_WriteTOC($sWorkingFolder, $aFunctions)
_CHM_WriteFUNC($sFile, $sWorkingFolder)
Global $sHHPFile = _WriteHHP($sWorkingFolder, $aFunctions)
_CHM_WriteHomePage($sHomePage, $sLogoPic, $sWorkingFolder)

Global Const $hCHMDLL__HHADLL = _CHM_HHAdll()
If $hCHMDLL__HHADLL = -1 Then
	ConsoleWriteError("!!!!!Unable to proceed with compilation!!!!!" & @CRLF)
	MsgBox(64, "End", "Compilation can't proceed. hha.dll is missing!")
	Exit -2
EndIf

ConsoleWrite("> All pages created. Compiling now..." & @CRLF & @CRLF)

Global $sOutputFile = _CHM_Compile($hCHMDLL__HHADLL, $sHHPFile)
If $sOutputFile Then
	ConsoleWrite(@CRLF & "> All done!. Over and out." & @CRLF & @CRLF)
	If MsgBox(68, "All done!", "Help file successfully cteated" & @CRLF & "Run it?" & @CRLF) = 6 Then ShellExecute($sOutputFile)
Else
	MsgBox(48, "Failure", "Compilation failed!")
EndIf


Func _GetVersionNumber($sFile)
	Local $sVer = FileRead($sFile, 2048)
	Local $aArray = StringRegExp($sVer, "(?s).*?\Q; File Version.........:\E\s*(.*?)\v+", 3)
	If @error Then Return SetError(1, 0, "")
	Return $aArray[0]
EndFunc   ;==>_GetVersionNumber


Func _CHM_WriteJScript($sWorkingFolder, $sFileUDF, $sHomeLink)
	Local $sData = FileRead($sFileUDF)
	; Locate Function headers
	Local $aHeaders = StringRegExp($sData, "(?si); #FUNCTION# ;*.*?;\h*=", 3)
	If @error Then Return SetError(1, 0, 0)

	Local $sNewContent, $sFunctionName, $sDesc, $sSearchTerms
	; Build-up search terms based on functions names and descriptions
	For $j = 0 To UBound($aHeaders) - 1
		$sFunctionName = _CHM_GetHeaderData($aHeaders[$j], "Name")
		$sSearchTerms = StringRegExpReplace(StringReplace(StringReplace(StringRegExpReplace(StringReplace(_CHM_GetHeaderData($aHeaders[$j], "Parameters") & " " & _CHM_GetHeaderData($aHeaders[$j], "Remarks"), @CRLF, " "), "\s{2,}", ""), "[out]", ""), "[optional]", ""), '[\;\-\"\[\]\(\)]', "")
		$sDesc = _CHM_GetHeaderData($aHeaders[$j], "Description")
		$sNewContent &= 'searchDB.push(new searchOption("' & $sFunctionName & '", "' & $sSearchTerms & '", "' & $sDesc & '"));' & @CRLF
	Next

	Local $sJS = StringReplace(FileRead($sJSFile), "//--PLACEHOLDER-I-DO-NOT-REMOVE-ME--//", $sNewContent)
	$sJS = StringReplace($sJS, "//--PLACEHOLDER-II-DO-NOT-REMOVE-ME--//", $sHomeLink)

	Local $hJS = FileOpen($sWorkingFolder & "\HTML\JScript\Script.js", 10)
	FileWrite($hJS, $sJS)
	FileClose($hJS)

	FileCopy($sJSRFile, $sWorkingFolder & "\HTML\JScript\Script_Right.js", 9)
EndFunc   ;==>_CHM_WriteJScript


Func _CHM_UDFToHTMPages($sFileUDF, ByRef $aFunctions, $sFolder = Default)
	If StringRight($sFolder, 1) = "\" Then $sFolder = StringTrimRight($sFolder, 1)
	; Read the file
	Local $sData = FileRead($sFileUDF)
	; Locate Function headers
	Local $aHeaders = StringRegExp($sData, "(?si); #FUNCTION# ;*.*?;\h*=", 3)
	If @error Then Return SetError(1, 0, 0)

	Local $aFuncs[UBound($aHeaders)]

	Local $sFileHTM, $hFileHTM
	Local $sInclude = StringRegExpReplace($sFileUDF, ".*\\", "")
	Local $sName = StringReplace($sInclude, ".au3", "")
	If $sFolder = Default Then $sFolder = StringReplace($sFileUDF, $sInclude, "") & $sName & $CHM_FOLDERSUFFIX

	Local $sFunctionName, $sHTM
	Local $sParams, $aParams, $iParamsUBound, $sRemarks, $sRelated, $aRelated, $sLink, $sAu3File, $sAu3Code, $sInnerCode

	For $j = 0 To UBound($aHeaders) - 1
		$sFunctionName = _CHM_GetHeaderData($aHeaders[$j], "Name")
		$aFuncs[$j] = $sFunctionName
		$sHTM = '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">' & @CRLF & _
				"<html>" & @CRLF & _
				"    <head>" & @CRLF & _
				"        <title>" & $sFunctionName & "</title>" & @CRLF & _
				'        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">' & @CRLF & _
				'        <meta http-equiv="X-UA-Compatible" content="IE=9">' & @CRLF & _
				'        <link href="../CSS/Default1.css" rel="stylesheet" type="text/css">' & @CRLF & _
				'        <script type="text/javascript" src="../JScript/Script_Right.js" charset="utf-8"></script>' & @CRLF & _
				"    </head>" & @CRLF & @CRLF & _
				"    <body>" & @CRLF & @CRLF
		$sHTM &= "<!--Description Section-->" & @CRLF
		$sHTM &= '        <h1 class="small">Function Reference</h1>' & @CRLF & _
				'        <hr style="height:0px">' & @CRLF
		If $sFunctionName = "_WinHttpSimpleReadDataAsync" Then
			$sHTM &= '        <h1 class="left_behind">' & $sFunctionName & '</h1>' & @CRLF
		Else
			$sHTM &= '        <h1>' & $sFunctionName & '</h1>' & @CRLF
		EndIf
		$sHTM &= '        <p class="funcdesc">' & _CHM_GetHeaderData($aHeaders[$j], "Description") & "<br></p>" & @CRLF & @CRLF

		$sHTM &= "<!--Syntax Section-->" & @CRLF
		$sHTM &= "        <h2>Syntax</h2>" & @CRLF & "        <p>"
		$sHTM &= '        <p class="codeheader">' & @CRLF & _
				'        #include "' & $sInclude & '"<br>' & @CRLF & _
				"        " & _CHM_GetHeaderData($aHeaders[$j], "Syntax") & "<br>" & @CRLF & _
				"        </p>" & @CRLF & @CRLF

		$sHTM &= "<!--Parameters Section-->" & @CRLF
		$sHTM &= "        <h2>Parameters</h2>" & @CRLF
		$sParams = _CHM_GetHeaderData($aHeaders[$j], "Parameters")
		If StringLeft(StringStripWS($sParams, 8), 4) = "none" Or Not $sParams Then
			$sHTM &= "        <p>&nbsp;None.</p><br>" & @CRLF & @CRLF
		Else
			$sHTM &= '        <table class="paramstable">' & @CRLF & _
					"            <tr>" & @CRLF
			$aParams = StringRegExp($sParams, "\h*(\$\w+|\Q(...)\E)\h*\-\h*(.+?)(?:\r\n|\Z)", 3)
			$iParamsUBound = UBound($aParams) - 1
			For $i = 0 To $iParamsUBound
				If Mod($i, 2) Then
					$sHTM &= '                <td class="rightpane">' & StringRegExpReplace($aParams[$i], "(\[.*\])", "<b>$1</b>") & "</td>" & @CRLF
					$sHTM &= "            </tr>" & @CRLF
					If $i < $iParamsUBound Then $sHTM &= "            <tr>" & @CRLF
				Else
					$sHTM &= '                <td class="leftpane">' & $aParams[$i] & "</td>" & @CRLF

				EndIf
			Next
			$sHTM &= "        </table>" & @CRLF & @CRLF
		EndIf

		$sHTM &= "<!--Return Value Section-->" & @CRLF
		$sHTM &= "        <h2>Return Value</h2>" & @CRLF & "        <p>"
		$sHTM &= StringReplace(StringStripCR(StringReplace(StringRegExpReplace(StringRegExpReplace(_CHM_GetHeaderData($aHeaders[$j], "Return values"), ";\h*", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"), "(\Q&nbsp;\E)+\QFailure\E", "Failure"), "|", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")), @LF, "<br>" & @CRLF & "        ") & "</p>"
		$sHTM &= @CRLF & @CRLF

		$sHTM &= "<!--Remarks Section-->" & @CRLF
		$sRemarks = _CHM_GetHeaderData($aHeaders[$j], "Remarks")
		If $sRemarks Then
			$sHTM &= "        <h2>Remarks</h2>" & @CRLF & _
					"        <p>" & StringRegExpReplace(StringReplace(StringRegExpReplace(StringRegExpReplace($sRemarks, ";\h*\+", "<br>"), ";\h*", "    "), "    |", "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"), "\[{2}(.*?\]?)\]{2}", '<span class="codelike">$1</span>') & "</p>" & @CRLF & _
					"        <br>" & @CRLF & @CRLF
		EndIf

		$sHTM &= "<!--Related Section-->" & @CRLF
		$sRelated = _CHM_GetHeaderData($aHeaders[$j], "Related")
		If $sRelated Then
			$sHTM &= "        <h2>Related</h2>"
			$aRelated = StringRegExp($sRelated, "\h*(\w+)\h*(?:\,|\Z)", 3)
			For $i = 0 To UBound($aRelated) - 1
				$sHTM &= StringRegExpReplace($aRelated[$i], "(\w+)", '<a onclick=''UpdateLocation("$1")'' href="$1.htm">$1</a>')
				If $i = UBound($aRelated) - 1 Then
					$sHTM &= @CRLF
				Else
					$sHTM &= ", "
				EndIf
			Next
			$sHTM &= "        <br>" & @CRLF & @CRLF
		EndIf

		$sHTM &= "<!--Link Section-->" & @CRLF
		$sLink = _CHM_GetHeaderData($aHeaders[$j], "Link")
		If $sLink Then
			$sHTM &= "        <h2>See Also</h2>" & @CRLF
			$sHTM &= '        <a title="External link" href="' & $sLink & '" onclick=''MSDN_Nav("' & $sLink & '", "' & $sFunctionName & '");return false;''>MSDN</a>' & @CRLF & @CRLF ; MSDN as visible text
			$sHTM &= "        <br>" & @CRLF & @CRLF
		EndIf

		$sHTM &= "<!--Example Section-->" & @CRLF
		$sHTM &= "        <h2>Example</h2>" & @CRLF

		Local $sSuffix = "", $fHasExample = False
		For $i = 0 To 10
			If $i > 0 Then $sSuffix = "_" & $i
			$sAu3File = $sExamplesFolder & "\" & $sFunctionName & $sSuffix & ".au3"
			If FileExists($sAu3File) Then
				$fHasExample = True
				$sAu3Code = FileRead($sAu3File)
				If StringStripWS($sAu3Code, 8) Then
					$sHTM &=  '<a class="button" onmouseover="Btn_OnMouseOver(this)" onmouseout=''Btn_OnMouseOut(this, "divtip")'' onclick=''BTN_OnClick("au3code' & $sSuffix & '", "divtip")''>Copy to clipboard</a>'
					$sHTM &= '<p class="codebox" id="au3code' & $sSuffix & '"><br> ' & @CRLF & _CHM_SyntaxHighlight($sAu3Code) & "</p><br>" & @CRLF
				EndIf
			EndIf
		Next
		If $fHasExample Then $sHTM &= '        <div id="divtip" name="divtip" class="tip"></div>' & @CRLF

		$sHTM &= @CRLF

		$sHTM &= "    </body>" & @CRLF & _
				"</html>" & @CRLF

		$sFileHTM = $sFolder & "\HTML\Functions\" & $sFunctionName & ".htm"
		$hFileHTM = FileOpen($sFileHTM, 10)
		FileWrite($sFileHTM, $sHTM)
		FileClose($hFileHTM)
		;ConsoleWrite($sHTM)
	Next
	$aFunctions = $aFuncs
	Return $sFolder
EndFunc   ;==>_CHM_UDFToHTMPages


Func _CHM_GetHeaderData($sString, $sTag)
	Local $aName = StringRegExp($sString, "(?s).*?;\h*\Q" & $sTag & "\E\h*\.*\h*:\h*(.*?)\r\n\h*;\h*(\w+\h*\w*\h*\.*\h*:\h*|\=+)", 3)
	If @error Then Return ""
	Return $aName[0]
EndFunc   ;==>_CHM_GetHeaderData


Func _CHM_WriteDefaultCSS($sWorkingFolder)
	FileCopy($sCssFile, $sWorkingFolder & "\HTML\CSS\Default1.css", 9)
EndFunc   ;==>_CHM_WriteDefaultCSS


Func _CHM_WriteTOC($sWorkingFolder, $aFunctions)
	Local $sHTML = '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">' & @CRLF & _
			"<html>" & @CRLF & _
			"<head>" & @CRLF & _
			"<title>TOC</title>" & @CRLF & _
			'<link href="../CSS/Default1.css" rel="stylesheet" type="text/css">' & @CRLF & _
			'<meta http-equiv="Content-Type" content="text/html; charset=utf-8">' & @CRLF & _
			'<meta http-equiv="X-UA-Compatible" content="IE=9">' & @CRLF & _
			'<script type="text/javascript" src="../JScript/Script.js" charset="utf-8"></script>' & @CRLF & _
			"</head>" & @CRLF & _
			'<body onload="document.getElementById(''ainit'').style.color = ''black''">' & @CRLF & _
			"<ul>" & @CRLF & _
			'<li class="single">' & @CRLF & _
			'<a class="nowrap" href="#" id="ainit" onCLick=''return SetIFrameSource("../CHM_HomePage.htm", this);''>WinHttp Function Notes</a>' & @CRLF & _
			"</li>" & @CRLF & _
			'<li class="parent">' & @CRLF & _
			'<a href="#" onCLick=''return SetIFrameSource("Functions.htm", this);''>Functions</a>' & @CRLF & _
			"</li>" & @CRLF

	For $i = 0 To UBound($aFunctions) - 1
		$sHTML &= '<li>' & @CRLF & _
				'<a href="#" onCLick=''return SetIFrameSource("' & $aFunctions[$i] & '.htm", this);''>' & $aFunctions[$i] & '</a>' & @CRLF & _
				'</li>' & @CRLF
	Next

	$sHTML &= '</ul>' & @CRLF & _
			"</body>" & @CRLF & _
			"</html>" & @CRLF

	Local $sFileHTM = $sWorkingFolder & "\HTML\Functions\TOC.htm"
	Local $hFileHTML = FileOpen($sFileHTM, 10)
	FileWrite($hFileHTML, $sHTML)
	FileClose($hFileHTML)
	Return $sFileHTM
EndFunc   ;==>_CHM_WriteTOC


Func _WriteHHP($sWorkingFolder, $aFunctions)
	Local $sName = StringReplace(StringRegExpReplace($sWorkingFolder, ".*\\", ""), $CHM_FOLDERSUFFIX, "")
	Local $sData = "[OPTIONS]" & @CRLF & _
			"Compatibility=1.0" & @CRLF & _
			"Compiled File=" & $sName & " .chm" & @CRLF & _
			"Title=" & $sName & " Help" & @CRLF & _
			"Default Window=NewWindow" & @CRLF & _
			"Display compile progress=Yes" & @CRLF & _
			"Display compile notes=Yes" & @CRLF & _
			"Default Font=Segoe UI, 9" & @CRLF & _
			"Default topic=html\Functions\IFame.htm" & @CRLF & _
			"Language=0x409" & @CRLF & _
			"Full-text search=No" & @CRLF & @CRLF

	$sData &= "[WINDOWS]" & @CRLF & _
			'NewWindow="' & $sName & ' ' & $sCurrentVersionNumber & ' Help",' & _
			'"","","html\Functions\IFame.htm","",,,,,0x20420,0,0x0,,0x10030000,,,1,,,0' & @CRLF & @CRLF

	$sData &= "[FILES]" & @CRLF

	For $i = 0 To UBound($aFunctions) - 1
		$sData &= "html\Functions\" & $aFunctions[$i] & ".htm" & @CRLF
	Next

	DirCopy(@ScriptDir & "\Images", $sWorkingFolder & "\HTML\Images", 1)
	FileCopy(@ScriptDir & "\Extra\*", $sWorkingFolder & "\HTML\Functions", 1)

	$sData &= "html\CHM_HomePage.htm" & @CRLF
	$sData &= "html\Functions\TOC.htm" & @CRLF
	$sData &= "html\Functions\Functions.htm" & @CRLF
	$sData &= "html\Functions\tetris.htm" & @CRLF

	Local $sFileName, $sExt, $hSearch = FileFindFirstFile($sWorkingFolder & "\HTML\Images\*")
	While 1
		$sFileName = FileFindNextFile($hSearch)
		If @error Then ExitLoop
		$sExt = StringRegExpReplace($sFileName, ".*\.", ".")
		If $sExt = ".png" Or $sExt = ".jpg" Then $sData &= "html\Images\" & $sFileName & @CRLF
	WEnd
	FileClose($hSearch)

	Local $sFileHHP = $sWorkingFolder & "\" & $sName & ".hhp"
	Local $hFileHHP = FileOpen($sFileHHP, 10)
	FileWrite($hFileHHP, $sData)
	FileClose($hFileHHP)
	Return $sFileHHP
EndFunc   ;==>_WriteHHP


Func _CHM_WriteFUNC($sFileUDF, $sWorkingFolder)
	Local $sName = StringReplace(StringRegExpReplace($sWorkingFolder, ".*\\", ""), $CHM_FOLDERSUFFIX, "")
	Local $sHTM = '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">' & @CRLF & _
			"<head>" & @CRLF & _
			"  <title>" & "Functions</title>" & @CRLF & _
			'  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">' & @CRLF & _
			'  <link href="../CSS/Default1.css" rel="stylesheet" type="text/css">' & @CRLF & _
			'  <script type="text/javascript" src="../JScript/Script_Right.js" charset="utf-8"></script>' & @CRLF & _
			"</head>" & @CRLF & _
			"<body>" & @CRLF & _
			"<h1>" & "Function Reference</h1>" & @CRLF & _
			"<p>Below is a complete list of the Functions available in " & $sName & ".<br>" & @CRLF & _
			"Click on a function name for a detailed description.</p>" & @CRLF & _
			"<p>&nbsp;</p>" & @CRLF & _
			'<table class="paramstable">' & @CRLF & _
			"<tr>" & @CRLF & _
			'  <th class="left">' & ' Function</th>' & @CRLF & _
			'  <th class="right">Description</th>' & @CRLF & _
			"</tr>" & @CRLF

	Local $sData = FileRead($sFileUDF)
	Local $aHeaders = StringRegExp($sData, "(?si); #FUNCTION# ;*.*?;\h*=", 3)
	If @error Then Return SetError(1, 0, "")
	Local $sFunctionName

	For $j = 0 To UBound($aHeaders) - 1
		$sFunctionName = _CHM_GetHeaderData($aHeaders[$j], "Name")
		$sHTM &= "<tr>" & @CRLF & _
				'    <td><a onclick=''UpdateLocation("' & $sFunctionName & '")'' href="' & $sFunctionName & '.htm">' & $sFunctionName & '</a></td>' & @CRLF & _
				"    <td>" & _CHM_GetHeaderData($aHeaders[$j], "Description") & "<br></td>" & @CRLF & _
				"</tr>" & @CRLF
	Next

	$sHTM &= "</table>" & @CRLF

	$sHTM &= "<br>" & @CRLF & _
			"<p>&nbsp;</p>" & @CRLF & @CRLF

	$sHTM &= "</body>" & @CRLF & _
			"</html>" & @CRLF

	Local $sFileHTM = $sWorkingFolder & "\HTML\Functions\Functions.htm"
	Local $hFileHTM = FileOpen($sFileHTM, 10)
	FileWrite($hFileHTM, $sHTM)
	FileClose($hFileHTM)
	Return $sFileHTM
EndFunc   ;==>_CHM_WriteFUNC


Func _CHM_WriteHomePage($sHomePage, $sLogoPic, $sWorkingFolder)
	Local $sName = StringReplace(StringRegExpReplace($sWorkingFolder, ".*\\", ""), $CHM_FOLDERSUFFIX, "")
	If FileExists($sHomePage) Then
		$sHomePage = FileRead($sHomePage)
		$sHomePage = StringRegExpReplace($sHomePage, "\QCurrent version is \E\d+\.\d+\.(\d+\.)?(\d+?\.)?", 'Current version is <span class="mark_v">' & $sCurrentVersionNumber & "</span>")
		$sHomePage = StringRegExpReplace($sHomePage, "(?si)(<\s*/*\Qbody\E.*?>|<\s*/*\Qhtml\E.*?>|<\s*head\s*>.*?<\s*/head\s*>|<\s*\Q!DOCTYPE\E.*?>)", "")
		$sHomePage = StringRegExpReplace($sHomePage, '(\Qhref=\E(".*"))', '$1 onclick=''window.open($2);return false;''')
	Else
		$sHomePage = ""
	EndIf
	If FileExists($sLogoPic) Then FileCopy($sLogoPic, $sWorkingFolder & "\HTML\Images\" & StringRegExpReplace($sLogoPic, ".*\\", ""), 9)
	Local $sHTM = '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">' & @CRLF & _
			"<html>" & @CRLF & _
			"    <head>" & @CRLF & _
			"        <title>" & $sName & " Function Notes" & "</title>" & @CRLF & _
			'        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">' & @CRLF & _
			'        <meta http-equiv="X-UA-Compatible" content="IE=9">' & @CRLF & _
			'        <link href="CSS/Default1.css" rel="stylesheet" type="text/css">' & @CRLF & _
			"    </head>" & @CRLF & @CRLF & _
			"    <body>" & @CRLF & _
			"        <h1>" & $sName & "</h1>" & @CRLF & _
			'        <img class="logo" src="Images/' & StringRegExpReplace($sLogoPic, ".*\\", "") & '" width="400" height="150" alt=""><br>' & @CRLF & _
			"        <p>Welcome to the helpfile of <strong>" & $sName & "!</strong><br>" & @CRLF & _
			"<!--Start passed content-->" & @CRLF & _
			$sHomePage & @CRLF & _
			"<!--End passed content-->" & @CRLF & _
			"    </body>" & @CRLF & _
			"</html>" & @CRLF

	Local $sFileHTM = $sWorkingFolder & "\HTML\CHM_HomePage.htm"
	Local $hFileHTM = FileOpen($sFileHTM, 10)
	FileWrite($hFileHTM, $sHTM)
	FileClose($hFileHTM)
	Return $sFileHTM
EndFunc   ;==>_CHM_WriteHomePage


Func _CHM_HHAdll()
	; This is strictly 32-bit function because of such dlls
	If @AutoItX64 Then Return SetError(1, 0, -1)
	; Check if dlls are in place
	Local $sHHADll = "hha.dll"
	Local $sITCCDLL = "itcc.dll"
	Local $hHHADll = DllOpen($sHHADll)
	Local $hITCCDLL = DllOpen($sITCCDLL)
	; Preset control flags
	Local $fHHANotInstalled = False
	Local $fITCCNotInstalled = False
	; Deal wit issues
	If $hHHADll = -1 Then
		$sHHADll = @SystemDir & "\" & $sHHADll
		If FileInstall("C:\Users\trancexx\Desktop\HHA_DLL\hha.dll", $sHHADll) Then
			$hHHADll = DllOpen($sHHADll)
		Else
			$fHHANotInstalled = True
		EndIf
	EndIf
	If $hITCCDLL = -1 Then
		$sITCCDLL = @SystemDir & "\" & $sITCCDLL
		If FileInstall("C:\Users\trancexx\Desktop\HHA_DLL\itcc.dll", $sITCCDLL) Then
			_CHM_RegisterServer($sITCCDLL)
		Else
			$fITCCNotInstalled = True
		EndIf
	Else
		DllClose($hITCCDLL)
		_CHM_RegisterServer($sITCCDLL)
	EndIf
	If @Compiled Then
		; Construct the string for other instance to execute with admin privileges
		Local $sShellExecuteString
		If $fHHANotInstalled And $fITCCNotInstalled Then
			$sShellExecuteString = ' /AutoIt3ExecuteLine ' & _
					'"Exit (' & _
					'FileInstall(''C:\Users\trancexx\Desktop\HHA_DLL\hha.dll'', ''' & $sHHADll & ''') & ' & _
					'FileInstall(''C:\Users\trancexx\Desktop\HHA_DLL\itcc.dll'', ''' & $sITCCDLL & ''') & ' & _
					'DllCall(''ole32.dll'', ''long'', ''OleInitialize'', ''ptr'', 0) & ' & _
					'DllCall(''' & $sITCCDLL & ''', ''long'', ''DllRegisterServer'')' & _
					')"'
		Else
			If $fHHANotInstalled Then
				$sShellExecuteString = ' /AutoIt3ExecuteLine ' & _
						'"Exit Not ' & _
						'FileInstall(''C:\Users\trancexx\Desktop\HHA_DLL\hha.dll'', ''' & $sHHADll & ''')'
			EndIf
			If $fITCCNotInstalled Then
				$sShellExecuteString = ' /AutoIt3ExecuteLine ' & _
						'"Exit (' & _
						'FileInstall(''C:\Users\trancexx\Desktop\HHA_DLL\itcc.dll'', ''' & $sITCCDLL & ''') & ' & _
						'DllCall(''ole32.dll'', ''long'', ''OleInitialize'', ''ptr'', 0) & ' & _
						'DllCall(''' & $sITCCDLL & ''', ''long'', ''DllRegisterServer'')' & _
						')"'
			EndIf
		EndIf
		ConsoleWrite($fHHANotInstalled & "    " & $fITCCNotInstalled & @CRLF)
		If $fHHANotInstalled Or $fITCCNotInstalled Then ShellExecuteWait('"' & @AutoItExe & '"', $sShellExecuteString, "", "runas", @SW_HIDE)
	EndIf
	; Check now
	If $fHHANotInstalled Then $hHHADll = DllOpen($sHHADll)
	If $hHHADll = -1 Then Return SetError(2, 0, -1)
	; Return pseudo-handle
	Return $hHHADll
EndFunc   ;==>_CHM_HHAdll


Func _CHM_RegisterServer($sDll)
	Local $fInit, $fError
	Local $aCall = DllCall("ole32.dll", "long", "OleInitialize", "ptr", 0)
	If Not @error Then $fInit = $aCall[0] <> 1 ; The COM library is already initialized
	$aCall = DllCall($sDll, "long", "DllRegisterServer")
	If @error Then $fError = True
	If $fInit Then DllCall("ole32.dll", "none", "OleUninitialize")
	If $fError Then Return SetError(2, 0, False)
	Return SetError($aCall[0] <> 0, $aCall[0], $aCall[0] = 0)
EndFunc   ;==>_CHM_RegisterServer


Func _CHM_Compile($hHHADll, $sHHPFile)
	Local $fInit
	Local $aCall = DllCall("ole32.dll", "long", "OleInitialize", "ptr", 0)
	If Not @error Then $fInit = $aCall[0] <> 1 ; The COM library is already initialized
	; Callbacks
	Local Static $pFuncLog = DllCallbackGetPtr(DllCallbackRegister("_CHM_Log", "bool", "str"))
	Local Static $pFuncProc = DllCallbackGetPtr(DllCallbackRegister("_CHM_Proc", "bool", "str"))
	Local $tHHAData = DllStructCreate("dword[12];char[256]")
	DllStructSetData($tHHAData, 1, DllStructGetSize($tHHAData), 1)
	; Compile
	DllCall($hHHADll, "bool", "HHA_CompileHHP", _ ;<- it's more likely "none" instead of "bool"
			"str", $sHHPFile, _
			"ptr", $pFuncLog, _
			"ptr", $pFuncProc, _
			"ptr", DllStructGetPtr($tHHAData))
	If @error Or Not DllStructGetData($tHHAData, 2) Then ; checking if the help file is created or error occurred
		If $fInit Then DllCall("ole32.dll", "none", "OleUninitialize")
		Return SetError(1, 0, "")
	EndIf
	If $fInit Then DllCall("ole32.dll", "none", "OleUninitialize")
	Return DllStructGetData($tHHAData, 2)
EndFunc   ;==>_CHM_Compile


Func _CHM_Log($sString)
	$sString = StringReplace($sString, @CRLF, "")
	If $sString Then ConsoleWrite("+....." & $sString & @CRLF)
	Return True
EndFunc   ;==>_CHM_Log


Func _CHM_Proc($sString)
	Return True ; not interested
	If $sString Then ConsoleWrite("> Processing....." & $sString & @CRLF)
	Return True
EndFunc   ;==>_CHM_Proc


Func _CHM_SyntaxHighlight($sAu3Code) ; MrCreator's modified
	$sAu3Code = StringReplace($sAu3Code & @CRLF, @TAB, "    ")
	Local $sPattern1, $sPattern2
	Local $sReplace1, $sReplace2
	Local $sUnique_Str_Quote = '%@~@%'
	Local $sUnique_Str_Include = '!~@%@~!' ; "!" must be the first character
	While StringInStr($sAu3Code, $sUnique_Str_Quote)
		$sUnique_Str_Quote &= Random(10000, 99999)
	WEnd
	While StringInStr($sAu3Code, $sUnique_Str_Include)
		$sUnique_Str_Include &= Random(10000, 99999)
	WEnd
	; Get all strings to array
	$sPattern1 = '(?m)(("|'')[^\2\r\n]*?\2)' ;'(?s).*?(("|'')[^\2]*\2).*?'
	$sPattern2 = "(?si)#include\s+?(<[^\>]*>).*?"
	Local $aQuote_Strings = StringRegExp($sAu3Code, $sPattern1, 3)
	Local $aInclude_Strings = StringRegExp($sAu3Code, $sPattern2, 3)
	; Replace all the strings with unique marks
	$sPattern1 = '(?s)("|'')([^\1\r\n])*?(\1)'
	$sPattern2 = "(?si)(#include\s+?)<[^\>]*>(.*?)"
	$sReplace1 = $sUnique_Str_Quote
	$sReplace2 = '\1' & $sUnique_Str_Include & '\2'
	$sAu3Code = StringRegExpReplace($sAu3Code, $sPattern1, $sReplace1)
	$sAu3Code = StringRegExpReplace($sAu3Code, $sPattern2, $sReplace2)
	$sPattern1 = '([\(\)\[\]\<\>\.\*\+\-\=\&\^\,\/])'
	$sReplace1 = '<span class="S8">\1</span>'
	; Highlight the operators, brakets, commas (must be done before all other parsers)
	$sAu3Code = StringRegExpReplace($sAu3Code, $sPattern1, $sReplace1)
	$sPattern1 = '(\W+)(_)(\W+)'
	$sReplace1 = '\1<span class="S8">\2</span>\3'
	; Highlight the line braking character, wich is the underscore (must be done before all other parsers)
	$sAu3Code = StringRegExpReplace($sAu3Code, $sPattern1, $sReplace1)
	$sPattern1 = '((?:\s+)?<span class="S8">\.</span>|(?:\s+)?\.)([^\d\$]\w+)'
	$sReplace1 = '\1<span class="S14">\2</span>'
	; Highlight the COM Objects
	$sAu3Code = StringRegExpReplace($sAu3Code, $sPattern1, $sReplace1)
	$sPattern1 = '([^\w#@])(\d+)([^\w])'
	$sReplace1 = '\1<span class="S3">\2</span>\3'
	; Highlight the number
	$sAu3Code = StringRegExpReplace($sAu3Code, $sPattern1, $sReplace1)
	; Highlight the keyword
	$sAu3Code = _CHM_ParseKeywords($sAu3Code)
	; Highlight the macros
	$sAu3Code = _CHM_ParseMacros($sAu3Code)
	; Highlight special keywords
	$sAu3Code = _CHM_ParseSpecial($sAu3Code)
	; Highlight the PreProcessor
	$sAu3Code = _CHM_ParsePreProcessor($sAu3Code)

	$sPattern1 = '([^\w#@])((?i)0x[abcdef\d]+)([^abcdef\d])'
	$sReplace1 = '\1<span class="S3">\2</span>\3'
	; Highlight the hexadecimal number
	$sAu3Code = StringRegExpReplace($sAu3Code, $sPattern1, $sReplace1)

	$sPattern1 = '\$(\w+)?'
	$sReplace1 = '<span class="S9">$\1</span>'
	; Highlight variables (also can be just the dollar sign)
	$sAu3Code = StringRegExpReplace($sAu3Code, $sPattern1, $sReplace1)

	$sPattern1 = '(\w+)(\h*<span class="S8">\()'
	$sReplace1 = '<span class="S4">\1</span>\2'
	; Highlight finaly the '#White space' (only user defined functions)
	$sAu3Code = StringRegExpReplace($sAu3Code, $sPattern1, $sReplace1)

	; Highlight commented lines / comment block (plus extra parsers due to need of the loop, see the function's body)
	$sAu3Code = _CHM_ParseComments($sAu3Code)

	; Replace back the unique marks with the original one and wrap them with "string" tags
	For $i = 0 To UBound($aQuote_Strings) - 1 Step 2
		$aQuote_Strings[$i] = StringReplace($aQuote_Strings[$i], "&", "&amp;")
		$aQuote_Strings[$i] = StringReplace($aQuote_Strings[$i], '<', '&lt;')
		$aQuote_Strings[$i] = StringReplace($aQuote_Strings[$i], '>', '&gt;')
		$sAu3Code = StringReplace($sAu3Code, $sUnique_Str_Quote, '<span class="S7">' & $aQuote_Strings[$i] & '</span>', 1)
	Next
	For $i = 0 To UBound($aInclude_Strings) - 1
		$aInclude_Strings[$i] = StringReplace($aInclude_Strings[$i], '<', '&lt;')
		$aInclude_Strings[$i] = StringReplace($aInclude_Strings[$i], '>', '&gt;')
		$sAu3Code = StringReplace($sAu3Code, $sUnique_Str_Include, '<span class="S7">' & $aInclude_Strings[$i] & '</span>', 1)
	Next
	; Strip tags from "string" inside commented lines
	Do
		$sAu3Code = StringRegExpReplace($sAu3Code, _
				'(.*?<span class="S1">;.*?)' & _
				'(?:<span class="S7">|<span class="S10">)(.*?)</span>(.*?)', _
				'\1\2\3')
	Until Not @extended
	$sAu3Code = _CHM_ParseUDFs($sAu3Code)
	Do
		$sAu3Code = StringRegExpReplace($sAu3Code, "\h{2,4}(\Q<span\E.*?>)", "$1&nbsp;&nbsp;&nbsp;&nbsp;")
	Until Not @extended
	$sAu3Code = StringRegExpReplace($sAu3Code, "\h(\Q<span\E.*?>)", "$1&nbsp;")
	$sAu3Code = StringRegExpReplace($sAu3Code, "(\Q<span\E.*?>)\v", "$1<br>")
	Do
		$sAu3Code = StringRegExpReplace($sAu3Code, "(\Q<br>\E)\v{2}", "$1<br>")
	Until Not @extended
	$sAu3Code = StringRegExpReplace($sAu3Code, ">\h+<", ">&nbsp;<")
	$sAu3Code = StringReplace($sAu3Code, "<>", "")
	$sAu3Code = StringReplace($sAu3Code, "&<", "&amp;<")
	$sAu3Code = StringReplace($sAu3Code, @CRLF, "<br>")
	Return $sAu3Code
EndFunc   ;==>_CHM_SyntaxHighlight


Func _CHM_ParseUDFs($sAu3Code)
	$sAu3Code = StringRegExpReplace($sAu3Code, '\Q<span class="S4">_\E([[:alpha:]])', '<span class="S15">_$1')
	Return $sAu3Code
EndFunc   ;==>_CHM_ParseUDFs


Func _CHM_ParseKeywords($sAu3Code)
	Local $aKeywords = _CHM_AutoIt_Keywords()
	For $i = 0 To UBound($aKeywords) - 1
		$sAu3Code = StringRegExpReplace($sAu3Code, '([^a-zA-Z_\$@])((?i)' & $aKeywords[$i] & ')(\W)', '\1<span class="S5">\2</span>\3')
	Next
	Return $sAu3Code
EndFunc   ;==>_CHM_ParseKeywords


Func _CHM_ParseMacros($sAu3Code)
	Local $aMacros = _CHM_AutoIt_Macros()
	For $i = 0 To UBound($aMacros) - 1
		$sAu3Code = StringRegExpReplace($sAu3Code, '(\W+|\A)((?i)' & $aMacros[$i] & ')(\W+|$)', '\1<span class="S6">\2</span>\3')
	Next
	Return $sAu3Code
EndFunc   ;==>_CHM_ParseMacros


Func _CHM_ParsePreProcessor($sAu3Code)
	Local $aPreProcessor = _CHM_AutoIt_Preprocessor()
	For $i = 0 To UBound($aPreProcessor) - 1
		$sAu3Code = StringRegExpReplace($sAu3Code, '(?i)\Q<span class="S12">\E(\Q' & $aPreProcessor[$i] & '\E)', '<span class="S11">$1')
	Next
	Return $sAu3Code
EndFunc   ;==>_CHM_ParsePreProcessor


Func _CHM_ParseSpecial($sAu3Code)
	Do
		$sAu3Code = StringRegExpReplace($sAu3Code, '(\W+|\A)((?i)#.*)<(?:span|a href=).*?>(.*)</(?:span|a)>', '\1\2\3')
	Until Not @extended
	$sAu3Code = StringRegExpReplace($sAu3Code, '(\W+|\A)((?i)#.*)', '\1<span class="S12">\2</span>')
	Return $sAu3Code
EndFunc   ;==>_CHM_ParseSpecial


Func _CHM_ParseComments($sAu3Code)
	Local $aCode = StringSplit(StringStripCR($sAu3Code), @LF, 2)
	Local $aComments, $iSubStart = 0
	Local $sCode
	; Go thru the code and check each line...
	For $i = 0 To UBound($aCode) - 1
		; Commented line
		If StringRegExp($aCode[$i], '(\s+)?([^(&lt|&gt)]|\h+?|^);') Then
			; Remove all tags
			$aComments = StringRegExp($aCode[$i], '([^;]*);(.*?)$', 3)
			$aComments[1] = StringReplace($aComments[1], "&", "&amp;")
			If UBound($aComments) > 1 Then
				Do
					$aComments[1] = StringRegExpReplace($aComments[1], '<\w+\h?[^>]*?>(.*?)</\w+>', '\1')
				Until Not @extended
				$aCode[$i] = $aComments[0] & '<span class="S2">;' & $aComments[1] & '</span>'
			EndIf
			$sCode &= $aCode[$i] & @CRLF
			; Comment block
		ElseIf StringRegExp($aCode[$i], '(?i)(\s+)?#(cs|comments.*?start)(.*)') Then
			; Remove all tags
			$aCode[$i] = StringReplace($aCode[$i], "&", "&amp;")
			Do
				$aCode[$i] = StringRegExpReplace($aCode[$i], '<\w+\h?[^>]*?>(.*?)</\w+>', '\1')
			Until Not @extended
			; Add the comment *open* tag
			$sCode &= StringRegExpReplace($aCode[$i], '(?i)(\s+)?#(cs|comments.*?start)(.*)', '\1<span class="S1">#\2\3') & @CRLF
			$sCode &= "<br>"
			$iSubStart += 1
			; Now check each line for ending of the comment block
			For $j = $i + 1 To UBound($aCode) - 1
				$i = $j
				; Remove all tags
				Do
					$aCode[$j] = StringRegExpReplace($aCode[$j], '<\w+\h?[^>]*?>(.*?)</\w+>', '\1')
				Until Not @extended
				$sCode &= $aCode[$j] & "<br>"
				$sCode = StringRegExpReplace($sCode, ">\h", ">&nbsp;")
				Do
					$sCode = StringRegExpReplace($sCode, "\Q&nbsp;\E\h", "&nbsp;")
				Until Not @extended
				; Check if current line of code is the (sub)start of comment block. If so, make a "note" for it (inrease the comments-start counter by one)
				If StringRegExp($aCode[$j], '(?i)(\s+)?#(cs|comments.*?start)(.*)') Then $iSubStart += 1
				; Check if current line of code is the end of sub comment block. If so, decrease the comments-start counter by one (to allow the ending of all comments)
				If $iSubStart > 0 And StringRegExp($aCode[$j], '(?i)(\s+)?#(ce|comments.*?end)(.*)') Then $iSubStart -= 1
				; Check if current line of code is the end of (all) comment block(s). If so, exit this current loop
				If $iSubStart = 0 And StringRegExp($aCode[$j], '(?i)(\s+)?#(ce|comments.*?end)(.*)') Then
					$sCode &= '</span><br>'
					ExitLoop
				EndIf
			Next
		Else
			$sCode &= $aCode[$i] & @CRLF
		EndIf
	Next
	Return $sCode
EndFunc   ;==>_CHM_ParseComments


Func _CHM_AutoIt_Keywords()
	Local $sWords = "and byref case const continuecase continueloop default dim " & _
			"do else elseif endfunc endif endselect endswitch endwith enum exit exitloop false " & _
			"for func global if in local next not or redim return select static step switch then " & _
			"to true until wend while with"
	Return StringSplit($sWords, " ", 2)
EndFunc   ;==>_CHM_AutoIt_Keywords


Func _CHM_AutoIt_Macros()
	Local $sWords = "@appdatacommondir @appdatadir @autoitexe @autoitpid @autoitversion " & _
			"@autoitx64 @com_eventobj @commonfilesdir @compiled @computername @comspec @cpuarch " & _
			"@cr @crlf @desktopcommondir @desktopdepth @desktopdir @desktopheight @desktoprefresh " & _
			"@desktopwidth @documentscommondir @error @exitcode @exitmethod @extended @favoritescommondir " & _
			"@favoritesdir @gui_ctrlhandle @gui_ctrlid @gui_dragfile @gui_dragid @gui_dropid @gui_winhandle " & _
			"@homedrive @homepath @homeshare @hotkeypressed @hour @ipaddress1 @ipaddress2 @ipaddress3 " & _
			"@ipaddress4 @kblayout @lf @logondnsdomain @logondomain @logonserver @mday @min @mon " & _
			"@msec @muilang @mydocumentsdir @numparams @osarch @osbuild @oslang @osservicepack " & _
			"@ostype @osversion @programfilesdir @programscommondir @programsdir @scriptdir @scriptfullpath " & _
			"@scriptlinenumber @scriptname @sec @startmenucommondir @startmenudir @startupcommondir " & _
			"@startupdir @sw_disable @sw_enable @sw_hide @sw_lock @sw_maximize @sw_minimize @sw_restore " & _
			"@sw_show @sw_showdefault @sw_showmaximized @sw_showminimized @sw_showminnoactive " & _
			"@sw_showna @sw_shownoactivate @sw_shownormal @sw_unlock @systemdir @tab @tempdir " & _
			"@tray_id @trayiconflashing @trayiconvisible @username @userprofiledir @wday @windowsdir " & _
			"@workingdir @yday @year"
	Return StringSplit($sWords, " ", 2)
EndFunc   ;==>_CHM_AutoIt_Macros


Func _CHM_AutoIt_Preprocessor()
	Local $sWords = "#ce #comments-end #comments-start #cs #include #include-once " & _
			"#noautoit3execute #notrayicon #onautoitstartregister #requireadmin"
	Return StringSplit($sWords, " ", 2)
EndFunc   ;==>_CHM_AutoIt_Preprocessor
