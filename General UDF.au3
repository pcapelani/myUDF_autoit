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

Func __fFileCheck($sDestination = "") ;Check and if needed install dependencies like sqlite3.dll and sqlite3.dll to @TempDir&"\AutoIt\dependencies

	Local $sFileTempPath = @UserProfileDir&"\AutoIt\"
	If $sDestination <> "" And FileExists($sDestination) = 1 Then $sFileTempPath = $sDestination
	If StringRight($sFileTempPath,1) <> "\" Then $sFileTempPath = $sFileTempPath&"\"

	;sqlite3.dll
	If FileExists ($sFileTempPath&"sqlite3.dll") = 0 Then
		If FileExists ($sFileTempPath) = 0 Then
			If DirCreate ($sFileTempPath) = 0 Then Exit MsgBox ($MB_SYSTEMMODAL,"Error","Auto It can not create directory: "&$sFileTempPath)
		EndIf
		If FileInstall (".\dependencies\sqlite3.dll",$sFileTempPath) = 0 Then Exit MsgBox ($MB_SYSTEMMODAL,"Error","Auto It can not install sqlite3.dll to: "&$sFileTempPath)
	EndIf

	;sqlite3_x64.dll
	If FileExists ($sFileTempPath&"sqlite3_x64.dll") = 0 Then
		If FileInstall (".\dependencies\sqlite3_x64.dll",$sFileTempPath) = 0 Then Exit MsgBox ($MB_SYSTEMMODAL,"Error","Auto It can not install sqlite3_x64.dll to: "&$sFileTempPath)
	EndIf

	;config
	If FileExists ($sFileTempPath&"config") = 0 Then
		_FileCreate($sFileTempPath&"config")
		If @error Then Exit MsgBox ($MB_SYSTEMMODAL,"Error","Auto It can not create config file to: "&$sFileTempPath&@CRLF&"@error: "&@error)
		If FileWrite ($sFileTempPath&"config","d_path"&@CRLF&$sFileTempPath) = 0 Then Exit MsgBox ($MB_SYSTEMMODAL,"Error","Auto It can not write in config file : "&$sFileTempPath&@CRLF&"@error: "&@error)
	Else
		$g_aConfig = FileReadToArray ($sFileTempPath&"config")
		If @error Then Exit MsgBox ($MB_SYSTEMMODAL,"Error FileReadToArray","Auto It can not read config file")
		Local $iArraySearch = _ArraySearch ($g_aConfig,"d_path")
		If @error Then
			If MsgBox (4,"Warning","Auto It can not find d_path in config file"&@CRLF&"Add ?") = 6 Then
				If FileWrite ($sFileTempPath&"config",@CRLF&"d_path"&@CRLF&$sFileTempPath) = 0 Then Exit MsgBox ($MB_SYSTEMMODAL,"Error FileWrite","Auto It can not write in config file : "&$sFileTempPath&@CRLF&"@error: "&@error)
				$g_aConfig = FileReadToArray ($sFileTempPath&"config")
				If @error Then Exit MsgBox ($MB_SYSTEMMODAL,"Error FileReadToArray","Auto It can not read config file")
				$iArraySearch = _ArraySearch ($g_aConfig,"d_path")
				If @error Then Exit MsgBox ($MB_SYSTEMMODAL,"Error _ArraySearch","Unexpected error occurred")
			Else
				MsgBox (0,"Warning","Auto It can not proceed without d_path in config file")
				Exit
			EndIf
		EndIf
		If $g_aConfig[$iArraySearch+1] <> $sFileTempPath Then
			If MsgBox (4,"Warning","Different d_path in config file"&@CRLF&"Uptade: '"&$g_aConfig[$iArraySearch+1]&"'"&@CRLF&" to '"&$sFileTempPath&"'?") = 6 Then
				$g_aConfig[$iArraySearch+1] = $sFileTempPath
				Local $hFileOpen = FileOpen ($sFileTempPath&"config",2)
				If $hFileOpen = -1 Then Exit MsgBox ($MB_SYSTEMMODAL,"Error FileOpen","Auto It can not write in config file")
				FileClose($hFileOpen)
				_FileWriteFromArray ($sFileTempPath&"config",$g_aConfig,Default,Default,@CRLF)
				If @error Then Exit MsgBox ($MB_SYSTEMMODAL,"Error _FileWriteFromArray","Auto It can not write in config file")
			Else
				MsgBox (0,"Warning","Auto It can not proceed mismatch d_path in config file")
				Exit
			EndIf
		EndIf
	EndIf

	$g_aConfig = FileReadToArray ($sFileTempPath&"config")
	If @error Then Exit MsgBox ($MB_SYSTEMMODAL,"Error FileReadToArray","Auto It can not read config file")

EndFunc

Func __fFileWriteConfig($sConfig = "",$sConfigValue = "") ;Write to config

	If $sConfig = "" Or $sConfigValue = "" Then Return SetError (1,0,-1) ;Empty value
	Local $iArraySearch = _ArraySearch ($g_aConfig,"d_path")
	If @error Then
		Return SetError (2,0,-1) ;Auto It can not find d_path in config file
	EndIf

	_ArraySearch ($g_aConfig,$sConfig)
	If @error Then
		If FileWrite ($g_aConfig[$iArraySearch+1]&"config",@CRLF&$sConfig&@CRLF&$sConfigValue) = 0 Then
			Return SetError (3,0,-1) ;Auto It can not write in config file
		EndIf
		$g_aConfig = FileReadToArray ($g_aConfig[$iArraySearch+1]&"config")
		If @error Then
			Return SetError (4,0,-1) ;Auto It can not read config file
		EndIf
		Return 1 ;Success
	EndIf

	Return SetError (5,0,-1) ;Config name already exists

EndFunc

Func __fCreateDB ($sDbName = "autoit.db") ;Create a new DB or check if a existing one is working

	If IsString ($sDbName) = 0 Or $sDbName = "" Then Return SetError (1,0,-1) ;Name is not string
	Local $iArraySearch = _ArraySearch ($g_aConfig,"d_path")
	If @error Then
		Return SetError (2,0,-1) ;Auto It can not find d_path in config file
	EndIf
	If StringRight ($sDbName,3) <> ".db" Then $sDbName = $sDbName&".db"
	Local $sDbPathName = $g_aConfig[$iArraySearch+1]&$sDbName
	Local $sDllPath = $g_aConfig[$iArraySearch+1]&"sqlite3.dll"

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

Func __fCreateTable ($sTableName = "",$sDbName = "autoit.db")

	If $sTableName = "" Then Return 0
	Local $iArraySearch_dpath = _ArraySearch ($g_aConfig,"d_path")
	If @error Then Exit MsgBox ($MB_SYSTEMMODAL,"@error: "&@error,"Auto It can not find d_path in config file")
	If StringRight ($sDbName,3) <> ".db" Then $sDbName = $sDbName&".db"
	Local $sDbPathName = $g_aConfig[$iArraySearch_dpath+1]&$sDbName
	If FileExists ($sDbPathName) = 0 Then
		MsgBox (0,"Error",$sDbPathName&" - not found")
		Return 0
	EndIf
	Local $sDllPath = $g_aConfig[$iArraySearch_dpath+1]&"sqlite3.dll"

	_SQLite_Startup($sDllPath)
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "SQLite Error", "SQLite3.dll Can't be Loaded!")
	_SQLite_Open ($sDbPathName)
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "SQLite Error", "Can't create or open a Database!"&@CRLF&$sDbPathName)
	$iSQLiteRetun = _SQLite_Exec ( -1, "SELECT * FROM "&$sTableName&" LIMIT 1;")
	If $iSQLiteRetun = $SQLITE_OK Then
		MsgBox ($MB_SYSTEMMODAL,"SQLite Error","Table '"&$sTableName&"' already exists in db: "&$sDbPathName)
	Else
		$iSQLiteRetun = _SQLite_Exec ( -1, 'CREATE TABLE "'&$sTableName&'" ("id" INTEGER NOT NULL UNIQUE, PRIMARY KEY("id" AUTOINCREMENT));')
        If $iSQLiteRetun = $SQLITE_OK Then
			MsgBox (0,"SQLite","Table '"&$sTableName&"' created")
		Else
			MsgBox($MB_SYSTEMMODAL, "SQLite Error: " &$iSQLiteRetun,_SQLite_ErrMsg())
		EndIf
	EndIf
	_SQLite_Close ()
	_SQLite_Shutdown()

EndFunc


;read excel func and csv ; parameter