#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.16.0
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include-once
#include <MsgBoxConstants.au3>
#include <File.au3>
#include <Array.au3>
#include <SQLite.au3>


Global $g_aConfig
__fFileCheck()
;MsgBox (0,@UserProfileDir&"\AutoIt\sqlite3.dll",@ScriptDir&@CRLF&@WorkingDir&@CRLF&@LocalAppDataDir&@CRLF&@SystemDir&@CRLF&@WindowsDir)
__fCreateDB ()

Func __fFileCheck($sDestination = "") ;Check and if needed install dependencies like sqlite3.dll and sqlite3.dll to @TempDir&"\AutoIt\dependencies

	Local $sFileTempPath = @UserProfileDir&"\AutoIt\"
	If $sDestination <> "" And FileExists($sDestination) = 1 Then $sFileTempPath = $sDestination
	If StringRight($sFileTempPath,1) <> "\" Then $sFileTempPath = $sFileTempPath&"\"

	;sqlite3.dll
	If FileExists ($sFileTempPath&"sqlite3.dll") = 0 Then
		If FileExists ($sFileTempPath) = 0 Then
			$iDirCreate = DirCreate ($sFileTempPath)
			If $iDirCreate = 0 Then Exit MsgBox ($MB_SYSTEMMODAL,"Error","Auto It can not create directory: "&$sFileTempPath)
		EndIf
		$iFileInstall = FileInstall (".\dependencies\sqlite3.dll",$sFileTempPath)
		If $iFileInstall = 0 Then Exit MsgBox ($MB_SYSTEMMODAL,"Error","Auto It can not install sqlite3.dll to: "&$sFileTempPath)
	EndIf

	;sqlite3_x64.dll
	If FileExists ($sFileTempPath&"sqlite3_x64.dll") = 0 Then
		$iFileInstall = FileInstall (".\dependencies\sqlite3_x64.dll",$sFileTempPath)
		If $iFileInstall = 0 Then Exit MsgBox ($MB_SYSTEMMODAL,"Error","Auto It can not install sqlite3_x64.dll to: "&$sFileTempPath)
	EndIf

	;config
	If FileExists ($sFileTempPath&"config") = 0 Then
		$iFileCreate = _FileCreate($sFileTempPath&"config")
		If @error Then Exit MsgBox ($MB_SYSTEMMODAL,"Error","Auto It can not create config file to: "&$sFileTempPath&@CRLF&"@error: "&@error)
		FileWrite ($sFileTempPath&"config","d_path"&@CRLF&$sFileTempPath)
		If @error Then Exit MsgBox ($MB_SYSTEMMODAL,"Error","Auto It can not write in config file : "&$sFileTempPath&@CRLF&"@error: "&@error)
	EndIf

	$g_aConfig = FileReadToArray ($sFileTempPath&"config")

EndFunc

Func __fFileWriteConfig($sConfig = "",$sConfigValue = "") ; Write to config

	If $sConfig = "" Or $sConfigValue = "" Then Return 0
	$iArraySearch_dpath = _ArraySearch ($g_aConfig,"d_path")
	If @error Then Exit MsgBox ($MB_SYSTEMMODAL,"@error: "&@error,"Auto It can not find d_path in config file")
	$iArraySearch = _ArraySearch ($g_aConfig,$sConfig)
	If @error Then FileWrite ($g_aConfig[$iArraySearch_dpath+1]&"config",@CRLF&$sConfig&@CRLF&$sConfigValue)
	$g_aConfig = FileReadToArray ($g_aConfig[$iArraySearch_dpath+1]&"config")
	Return 1

EndFunc

Func __fCreateDB ($sDbName = "autoit.db") ;Create a new DB or check if a existing one is working

	If StringRight ($sDbName,3) <> ".db" Then $sDbName = $sDbName&".db"
	Local $sDbPathName = $g_aConfig[1]&$sDbName
	Local $sDllPath = @UserProfileDir&"\AutoIt\sqlite3.dll"
	_SQLite_Startup($sDllPath)
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "SQLite Error", "SQLite3.dll Can't be Loaded!")
	_SQLite_Open ($sDbPathName)
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "SQLite Error", "Can't create or open a Database!"&@CRLF&$sDbPathName)
	_SQLite_Close ()
	_SQLite_Shutdown()
	_ArraySearch ($g_aConfig,$sDbName)
	If @error Then
		Local $aArrayFindAll = _ArrayFindAll($g_aConfig,"db_name",Default,Default,Default,3)
		If @error Then
			__fFileWriteConfig("db_name",$sDbName)
		Else
			__fFileWriteConfig("db_name"&UBound($aArrayFindAll),$sDbName)
		EndIf
	EndIf

EndFunc



;read excel func and csv ; parameter