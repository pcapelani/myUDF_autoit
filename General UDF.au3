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
#include <Excel.au3>


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

	;sqlite3.exe
	If FileExists ($sFileTempPath&"sqlite3.exe") = 0 Then
		If FileInstall (".\dependencies\sqlite3.exe",$sFileTempPath) = 0 Then Exit MsgBox ($MB_SYSTEMMODAL,"Error","Auto It can not install sqlite3.exe to: "&$sFileTempPath)
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
	If @error Then Return SetError (2,0,-1) ;Auto It can not find d_path in config file

	_ArraySearch ($g_aConfig,$sConfig)
	If @error Then
		If FileWrite ($g_aConfig[$iArraySearch+1]&"config",@CRLF&$sConfig&@CRLF&$sConfigValue) = 0 Then Return SetError (3,0,-1) ;Auto It can not write in config file
		$g_aConfig = FileReadToArray ($g_aConfig[$iArraySearch+1]&"config")
		If @error Then Return SetError (4,0,-1) ;Auto It can not read config file
		Return 1 ;Success
	EndIf

	Return SetError (5,0,-1) ;Config name already exists

EndFunc

Func __fChangeConfig($sConfig = "",$sConfigValue = "") ;Rewrite existing configuration

	If $sConfig = "" Then Return SetError (1,0,-1) ;Empty value
	Local $iArraySearch_dpath = _ArraySearch ($g_aConfig,"d_path")
	If @error Then Return SetError (2,0,-1) ;Auto It can not find d_path in config file

	Local $iArraySearch = _ArraySearch ($g_aConfig,$sConfig)
	If @error Then
		Return SetError (3,0,-1) ;Config does not exist
	Else
		$g_aConfig[$iArraySearch+1] = $sConfigValue
		Local $hFileOpen = FileOpen ($g_aConfig[$iArraySearch_dpath+1]&"config",2)
		If $hFileOpen = -1 Then Return SetError (4,0,-1) ;Can not apply the changes
		If FileClose($hFileOpen) = 0 Then MsgBox ($MB_SYSTEMMODAL,"Error FileClose","Unexpected error occurred, check config file")
		_FileWriteFromArray ($g_aConfig[$iArraySearch_dpath+1]&"config",$g_aConfig,Default,Default,@CRLF)
		If @error Then
			MsgBox ($MB_SYSTEMMODAL,"Error _FileWriteFromArray","Auto It can not write in config file, possibly config reset")
			Return SetError (5,0,-1) ;Can not apply the changes, config  reset
		EndIf
	EndIf

	Return 1 ;Success

EndFunc

Func __fCreateDB ($sDbName = "autoit.db") ;Create a new DB or check if a existing one is working

	If IsString ($sDbName) = 0 Or $sDbName = "" Then Return SetError (1,0,-1) ;Name is not string
	Local $iArraySearch = _ArraySearch ($g_aConfig,"d_path")
	If @error Then Return SetError (2,0,-1) ;Auto It can not find d_path in config file
	If StringRight ($sDbName,3) <> ".db" Then $sDbName = $sDbName&".db"
	Local $sDbPathName = $g_aConfig[$iArraySearch+1]&$sDbName
	Local $sDllPath = $g_aConfig[$iArraySearch+1]&"sqlite3.dll"

	_SQLite_Startup($sDllPath)
	If @error Then Return SetError (3,0,-1) ;SQLite3.dll Can't be Loaded!
	_SQLite_Open ($sDbPathName)
	If @error Then
		_SQLite_Shutdown()
		Return SetError (4,0,-1) ;Can't create or open a Database
	EndIf
	_SQLite_Close ()
	_SQLite_Shutdown()
	_ArraySearch ($g_aConfig,$sDbName)
	If @error Then
		Local $aArrayFindAll = _ArrayFindAll($g_aConfig,"db_name",Default,Default,Default,3)
		If @error Then
			__fFileWriteConfig("db_name",$sDbName)
			Return 1 ;Success
		Else
			__fFileWriteConfig("db_name"&UBound($aArrayFindAll),$sDbName)
			Return 1 ;Success
		EndIf
	EndIf

	Return 2 ;Db working

EndFunc

Func __fCreateTable ($sTableName = "",$sDbName = "autoit.db")

	If $sTableName = "" Then Return SetError (1,0,-1) ;Empty value
	Local $iArraySearch = _ArraySearch ($g_aConfig,"d_path")
	If @error Then Return SetError (2,0,-1) ;Auto It can not find d_path in config file
	If StringRight ($sDbName,3) <> ".db" Then $sDbName = $sDbName&".db"
	Local $sDbPathName = $g_aConfig[$iArraySearch+1]&$sDbName
	If FileExists ($sDbPathName) = 0 Then Return  SetError (3,0,-1) ;Auto It can not find DB
	Local $sDllPath = $g_aConfig[$iArraySearch+1]&"sqlite3.dll"

	_SQLite_Startup($sDllPath)
	If @error Then Return SetError (4,0,-1) ;SQLite3.dll Can't be Loaded!
	_SQLite_Open ($sDbPathName)
	If @error Then
		_SQLite_Shutdown()
		Return SetError (5,0,-1) ;Can't create or open a Database
	EndIf
	$iSQLiteRetun = _SQLite_Exec ( -1, "SELECT * FROM "&$sTableName&" LIMIT 1;")
	If $iSQLiteRetun = $SQLITE_OK Then
		_SQLite_Close ()
		_SQLite_Shutdown()
		Return SetError (6,0,-1) ;Table already exists in db
	Else
		$iSQLiteRetun = _SQLite_Exec ( -1, 'CREATE TABLE IF NOT EXISTS '&$sTableName&' ("id" INTEGER NOT NULL UNIQUE, PRIMARY KEY("id" AUTOINCREMENT));')
        If $iSQLiteRetun = $SQLITE_OK Then
			_SQLite_Close ()
			_SQLite_Shutdown()
			Return 1 ;Success
		Else
			_SQLite_Close ()
			_SQLite_Shutdown()
			Return SetError (7,0,-1) ;Can not create table
		EndIf
	EndIf

EndFunc

Func __fSQLiteExec($sQuery="",$sDbName = "autoit.db")

	If $sQuery = "" Then Return SetError (1,0,-1) ;Empty value
	Local $iArraySearch = _ArraySearch ($g_aConfig,"d_path")
	If @error Then Return SetError (2,0,-1) ;Auto It can not find d_path in config file
	If StringRight ($sDbName,3) <> ".db" Then $sDbName = $sDbName&".db"
	Local $sDbPathName = $g_aConfig[$iArraySearch+1]&$sDbName
	If FileExists ($sDbPathName) = 0 Then Return  SetError (3,0,-1) ;Auto It can not find DB
	Local $sDllPath = $g_aConfig[$iArraySearch+1]&"sqlite3.dll"

	_SQLite_Startup($sDllPath)
	If @error Then Return SetError (4,0,-1) ;SQLite3.dll Can't be Loaded!
	_SQLite_Open ($sDbPathName)
	If @error Then
		_SQLite_Shutdown()
		Return SetError (5,0,-1) ;Can't create or open a Database
	EndIf
	;$sQuery = InputBox ("SQLite Query","SQLite Query")
	Local $iSQLiteRetun = _SQLite_Exec (-1, $sQuery)
	_SQLite_Close()
	_SQLite_Shutdown()
	Return $iSQLiteRetun

EndFunc

Func __fReadExcel ($sExcelPath = "",$sRange = "ALL", $sSheets = 1)

	If FileExists ($sExcelPath) = 0 Then Return  SetError (1,0,-1) ;Auto It can not find the file
	If IsString ($sRange) = 0 Then Return SetError (2,0,-1) ;$sRange is not string
	If IsInt($sSheets) = 0 Or $sSheets < 1 Then Return SetError (3,0,-1) ;$sSheets is not valid
	Local $iClosedBook, $oExcel

	Local $oWorkbook = _Excel_BookAttach ($sExcelPath)
	If @error Then
		$iClosedBook = 1
		$oExcel = _Excel_Open(False,False,False,False,True)
		If @error Then Return SetError (4,0,-1) ;Error creating excel App
	Else
		$oExcel = _Excel_Open()
		If @error Then Return SetError (4,0,-1) ;Error creating excel App
	EndIf
	If $iClosedBook = 1 Then $oWorkbook = _Excel_BookOpen($oExcel,$sExcelPath)
	If @error Then
		If $iClosedBook	= 1 Then _Excel_Close($oExcel)
		Return SetError (5,0,-1) ;Error opening workbook
	EndIf
	If $sRange = "ALL" Then
		Local $aResult = _Excel_RangeRead($oWorkbook, Default, $oWorkbook.Sheets($sSheets).Usedrange,Default,True)
		If @error Then
			If $iClosedBook = 1 Then _Excel_BookClose ( $oWorkbook, False )
			If $iClosedBook	= 1 Then _Excel_Close($oExcel)
			Return SetError (6,0,-1) ;Error reading from workbook
		EndIf
	Else
		If StringRegExp($sRange, "\d+") = 1 Then
			Local $aResult = _Excel_RangeRead($oWorkbook, Default, $oWorkbook.Sheets($sSheets).Range($sRange),Default,True)
			If @error Then
				If $iClosedBook = 1 Then _Excel_BookClose ( $oWorkbook, False )
				If $iClosedBook	= 1 Then _Excel_Close($oExcel)
				Return SetError (6,0,-1) ;Error reading from workbook
			EndIf
		Else
			Local $aResult = _Excel_RangeRead($oWorkbook, Default, $oWorkbook.Sheets($sSheets).Usedrange.Columns($sRange),Default,True)
			If @error Then
				If $iClosedBook = 1 Then _Excel_BookClose ( $oWorkbook, False )
				If $iClosedBook	= 1 Then _Excel_Close($oExcel)
				Return SetError (6,0,-1) ;Error reading from workbook
			EndIf
		EndIf
	EndIf
	If $iClosedBook = 1 Then _Excel_BookClose ( $oWorkbook, False )
	If $iClosedBook	= 1 Then _Excel_Close($oExcel)

	_ArrayDisplay($aResult)
	Return $aResult

EndFunc

Func __fWriteExcel($aArray,$sRange = Default)

	If IsString ($sRange) = 0 And $sRange <> Default Then Return SetError (1,0,-1) ;$sRange is not string
	Local $oExcel = _Excel_Open(), $iClosedExcel = @extended
	If @error Then Return SetError (2,0,-1) ;Error creating excel App
	Local $oWorkbook = _Excel_BookNew($oExcel, 1)
	If @error Then
		If $iClosedExcel = 1 Then _Excel_Close($oExcel)
		Return SetError (3,0,-1) ;Error creating workbook
	EndIf
	_Excel_RangeWrite ( $oWorkbook, 1, $aArray, $sRange, Default, True)
	If @error Then
		_Excel_BookClose ( $oWorkbook, False )
		If $iClosedExcel = 1 Then _Excel_Close($oExcel)
		Return SetError (4,0,-1) ;Error writing to worksheet
	EndIf

	Return 1

EndFunc

Func __fImportSQLite($sCSVPath,$sDbName = "autoit.db",$sTableName = "temp_temp",$bOverwrite = False, $sSeparator = ";",$sHeaders = "on",$sMode = "csv")

	If FileExists ($sCSVPath) = 0 Then Return  SetError (1,0,-1) ;Auto It can not find file
	Local $iArraySearch = _ArraySearch ($g_aConfig,"d_path")
	If @error Then Return SetError (2,0,-1) ;Auto It can not find d_path in config file
	If $sDbName = Default Then $sDbName = "autoit.db"
	If StringRight ($sDbName,3) <> ".db" Then $sDbName = $sDbName&".db"
	Local $sDbPathName = $g_aConfig[$iArraySearch+1]&$sDbName
	If FileExists ($sDbPathName) = 0 Then Return  SetError (3,0,-1) ;Auto It can not find DB
	If $sTableName = Default Then $sTableName = "temp_temp"
	If Not $bOverwrite = True Then $bOverwrite = False
	Local $sInput = $g_aConfig[$iArraySearch+1]&"temp_input.txt",  $sOutput = $g_aConfig[$iArraySearch+1]&"temp_output.txt", $sQuery

	If $bOverwrite = True Then
		$sQuery = ".output stdout"&@CRLF&".headers "&$sHeaders&@CRLF _
					&".mode "&$sMode&@CRLF&".separator '"&$sSeparator&"'"&@CRLF _
					&"DROP TABLE IF EXISTS "&$sTableName&";"&@CRLF _
					&".import '"&$sCSVPath&"' "&$sTableName
	Else
		$sQuery = ".output stdout"&@CRLF&".headers "&$sHeaders&@CRLF _
					&".mode "&$sMode&@CRLF&".separator '"&$sSeparator&"'"&@CRLF _
					&".import '"&$sCSVPath&"' "&$sTableName
	EndIf
	Local $hInputFile = FileOpen( $sInput,2)
	FileWrite($hInputFile,$sQuery)
	FileClose($hInputFile)
	Local $sCmd = @ComSpec & ' /c ""' & $g_aConfig[$iArraySearch+1]&'sqlite3.exe"' & '  "' _
				 & $sDbPathName _
				 & '" > "' & $sOutput _
				 & '" < "' & $sInput & '""'
	Local $iPID = Run($sCmd, @WorkingDir, @SW_HIDE, BitOR($STDERR_CHILD, $STDOUT_CHILD))
	ProcessWaitClose($iPID)
	Local $sStdout = StdoutRead($iPID)
	Local $sStderr = StderrRead($iPID)
	MsgBox(0,@error,"Query:"&@CRLF&$sQuery&@CRLF&"#-------#"&@CRLF&"Cmd: "&@CRLF&$sCmd&@CRLF&"#-------#"&@CRLF&"Stdout: "&$sStdout&@CRLF&"Stderr: "&$sStderr)

	Return $sStderr

EndFunc










;read excel func and csv ; parameter