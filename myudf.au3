#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.16.0
 Author:         Me, Myself and I

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include-once
#include <Array.au3>
#include <Excel.au3>
#include <File.au3>
#include <GuiListView.au3>
#include <SQLite.au3>

Global $g_aConfig ;

__fFileCheck()

; #FUNCTION __fFileCheck#
; Check and if needed install dependencies like sqlite3.dll and sqlite3.exe to _
; DefaultPath: "@TempDir&"\AutoIt\dependencies" or a specific one.
; Using a path that do not require admin privilages are recommended.
; Success: No return.
; Failure: End script.
Func __fFileCheck($sDestination = "")

	Local $sFileTempPath = @UserProfileDir&"\AutoIt\"
	If $sDestination <> "" And FileExists($sDestination) = 1 Then $sFileTempPath = $sDestination
	If StringRight($sFileTempPath,1) <> "\" Then $sFileTempPath = $sFileTempPath&"\"

	;sqlite3.dll
	If FileExists ($sFileTempPath&"sqlite3.dll") = 0 Then
		If FileExists ($sFileTempPath) = 0 Then
			If DirCreate ($sFileTempPath) = 0 Then Exit MsgBox (64,"DirCreate Error __fFileCheck()","Auto It can not create directory: "&@CRLF&$sFileTempPath)
		EndIf
		If FileInstall ("C:\Users\pedro.capelani\Documents\Code\Git\AutoIt-myUDF-main\dependencies\sqlite3.dll",$sFileTempPath) = 0 Then Exit MsgBox (64,"FileInstall Error __fFileCheck()","Auto It can not install sqlite3.dll to: "&@CRLF&$sFileTempPath)
	EndIf

	;sqlite3_x64.dll
	If FileExists ($sFileTempPath&"sqlite3_x64.dll") = 0 Then
		If FileInstall ("C:\Users\pedro.capelani\Documents\Code\Git\AutoIt-myUDF-main\dependencies\sqlite3_x64.dll",$sFileTempPath) = 0 Then Exit MsgBox (64,"FileInstall Error __fFileCheck()","Auto It can not install sqlite3_x64.dll to: "&@CRLF&$sFileTempPath)
	EndIf

	;sqlite3.exe
	If FileExists ($sFileTempPath&"sqlite3.exe") = 0 Then
		If FileInstall ("C:\Users\pedro.capelani\Documents\Code\Git\AutoIt-myUDF-main\dependencies\sqlite3.exe",$sFileTempPath) = 0 Then Exit MsgBox (64,"FileInstall Error __fFileCheck()","Auto It can not install sqlite3.exe to: "&$sFileTempPath)
	EndIf

	;config
	If FileExists ($sFileTempPath&"config") = 0 Then
		If FileWrite ($sFileTempPath&"config","d_path"&@CRLF&$sFileTempPath) = 0 Then Exit MsgBox (64,"FileWrite Error __fFileCheck()","Auto It can not write in config file: "&@CRLF&$sFileTempPath)
	Else
		$g_aConfig = FileReadToArray ($sFileTempPath&"config")
		If @error Then Exit MsgBox (64,"FileReadToArray Error __fFileCheck()","Auto It can not read config file"&@CRLF&"Path: "&$sFileTempPath&"config"&@CRLF&"@error: "&@error)
		Local $iArraySearch = _ArraySearch ($g_aConfig,"d_path")
		If @error Then
			If MsgBox (4,"Warning","Auto It can not find d_path in config file"&@CRLF&"Add ?") = 6 Then
				If _ArrayAdd ($g_aConfig,"d_path|"&$sFileTempPath) = -1 Then Exit MsgBox (64,"_ArrayAdd Error __fFileCheck()","Auto It can not add 'd_path' to array"&@CRLF&"@error: "&@error)
				If _FileWriteFromArray ($sFileTempPath&"config",$g_aConfig) = 0 Then Exit MsgBox (64,"_FileWriteFromArray Error __fFileCheck()","Auto It can not write in config file : "&$sFileTempPath&@CRLF&"@error: "&@error)
				$g_aConfig = FileReadToArray ($sFileTempPath&"config")
				If @error Then Exit MsgBox (64,"FileReadToArray Error __fFileCheck()","Auto It can not read config file:"&@CRLF&$sFileTempPath&"config"&@CRLF&"@error: "&@error)
				$iArraySearch = _ArraySearch ($g_aConfig,"d_path")
				If @error Then Exit MsgBox (64,"_ArraySearch Error __fFileCheck()","Unexpected error occurred"&@CRLF&"@error: "&@error)
			Else
				MsgBox (64,"Warning","Auto It can not proceed without d_path in config file")
				Exit
			EndIf
		EndIf
		If $g_aConfig[$iArraySearch+1] <> $sFileTempPath Then
			If MsgBox (4,"Warning","Different d_path in config file"&@CRLF&"Uptade: '"&$g_aConfig[$iArraySearch+1]&"'"&@CRLF&" to '"&$sFileTempPath&"'?") = 6 Then
				$g_aConfig[$iArraySearch+1] = $sFileTempPath
				If _FileWriteFromArray ($sFileTempPath&"config",$g_aConfig) = 0 Then Exit MsgBox (64,"_FileWriteFromArray Error __fFileCheck()","Auto It can not write in config file:"&@CRLF&$sFileTempPath&"config")
			Else
				MsgBox (64,"Warning","Auto It can not proceed mismatch d_path in config file:"&@CRLF&$sFileTempPath&"config")
				Exit
			EndIf
		EndIf
	EndIf

	$g_aConfig = FileReadToArray ($sFileTempPath&"config")
	If @error Then Exit MsgBox (64,"FileReadToArray Error __fFileCheck()","Auto It can not read config file:"&@CRLF&$sFileTempPath&"config"&@CRLF&"@error: "&@error)

EndFunc

; #FUNCTION __fGetConfig#
; Get value from config file.
; Success: Config value.
; Failure: -1 and sets the @error flag to non-zero.
Func __fGetConfig($sConfig)

	If $sConfig = "" Then Return SetError (1,0,-1) ;Empty value

	Local $iArraySearch = _ArraySearch ($g_aConfig,$sConfig)
	If @error Then Return SetError (2,0,-1) ;Not found

	Return $g_aConfig[$iArraySearch+1]

EndFunc

; #FUNCTION __fChangeConfig#
; Write to config file.
; Success: Return 1: added; 2: Changed.
; Failure: -1 and sets the @error flag to non-zero.
Func __fChangeConfig($sConfig = "",$sConfigValue = "", $iOverWrite = 0)

	If $sConfig = "" Then Return SetError (1,0,-1) ;Empty value
	If $sConfigValue  = "" Then $sConfigValue = " "
	If Not $iOverWrite = 1 Then $iOverWrite  = 0
	Local $iArraySearch_dpath = _ArraySearch ($g_aConfig,"d_path")
	If @error Then Return SetError (2,0,-1) ;Auto It can not find d_path in config file

	Local $iArraySearch = _ArraySearch ($g_aConfig,$sConfig)
	If @error Then
		If _ArrayAdd ($g_aConfig,$sConfig&"|"&$sConfigValue) = -1 Then Return SetError (3,0,-1) ;Auto It can not array add
		If _FileWriteFromArray ($g_aConfig[$iArraySearch_dpath+1]&"config",$g_aConfig) = 0 Then  Return SetError (4,0,-1) ;Auto It can not write in config file
		Return 1 ;Success
	Else
		If $iOverWrite = 0 Then Return SetError (5,0,-1) ;Config name already exists, overwrite disabled
		If $iOverWrite = 1 Then
			$g_aConfig[$iArraySearch+1] = $sConfigValue
			If _FileWriteFromArray ($g_aConfig[$iArraySearch_dpath+1]&"config",$g_aConfig) = 0 Then  Return SetError (4,0,-1) ;Auto It can not write in config file
			Return 2 ;Success changed
		EndIf
	EndIf

EndFunc

; #FUNCTION __fOpenDB#
; Create a new DB or check if a existing one is working
; Success: Return 1: created; 2: db working.
; Failure: -1 and sets the @error flag to non-zero.
Func __fOpenDB ($sDbName = "autoit.db")

	If IsString ($sDbName) = 0 Or $sDbName = "" Then Return SetError (1,0,-1) ;Name is not string
	Local $iArraySearch = _ArraySearch ($g_aConfig,"d_path")
	If @error Then Return SetError (2,0,-1) ;Auto It can not find d_path in config file
	If StringRight ($sDbName,3) <> ".db" Then $sDbName = $sDbName&".db"
	Local $sDbPathName = $g_aConfig[$iArraySearch+1]&$sDbName
	Local $sDllPath = $g_aConfig[$iArraySearch+1]&"sqlite3.dll"
	Local $iFileExists = FileExists($sDbPathName)

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
			__fChangeConfig("db_name",$sDbName)
			If $iFileExists = 1 Then Return 2 ;Db working
			Return 1 ;Success created
		Else
			__fChangeConfig("db_name"&UBound($aArrayFindAll),$sDbName)
			If $iFileExists = 1 Then Return 2 ;Db working
			Return 1 ;Success created
		EndIf
	EndIf

	If $iFileExists = 1 Then Return 2 ;Db working
	Return 1 ;Success created

EndFunc

; #FUNCTION __fCreateTable#
; Create a new table to db.
; Success: Return 1: created
; Failure: -1 and sets the @error flag to non-zero.
Func __fCreateTable ($sTableName,$sDbName = "autoit.db")

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

; #FUNCTION __fSQLiteExec#
; Query db. Not recommended for multiple function call.
; Success: Return $SQLITE_OK.
; Failure: -1 and sets the @error flag to non-zero.
Func __fSQLiteExec($sQuery,$sDbName = "autoit.db")

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

; #FUNCTION __fReadExcel#
; Read an excel sheet. Not recommended for multiple function call.
; Success: Array with results.
; Failure: -1 and sets the @error flag to non-zero.
Func __fReadExcel ($sExcelPath,$sRange = "ALL", $iSheets = Default,$bForceFunc = False)

	If FileExists ($sExcelPath) = 0 Then Return  SetError (1,0,-1) ;Auto It can not find the file
	If IsString ($sRange) = 0 Then Return SetError (2,0,-1) ;$sRange is not string
	If $iSheets <> Default And $iSheets < 1 Then $iSheets = Default
	If Not $bForceFunc = True Then $bForceFunc = False
	Local $iClosedBook, $oExcel

	Local $oWorkbook = _Excel_BookAttach ($sExcelPath)
	If @error Then
		$iClosedBook = 1
		$oExcel = _Excel_Open(False,False,False,False,True)
		If @error Then Return SetError (3,0,-1) ;Error creating excel App
	Else
		$oExcel = _Excel_Open()
		If @error Then Return SetError (3,0,-1) ;Error creating excel App
	EndIf
	If $iClosedBook = 1 Then $oWorkbook = _Excel_BookOpen($oExcel,$sExcelPath)
	If @error Then
		If $iClosedBook	= 1 Then _Excel_Close($oExcel)
		Return SetError (4,0,-1) ;Error opening workbook
	EndIf
	Local $aResult
	If $sRange = "ALL" Then
		If $iSheets = Default Then
			$aResult = _Excel_RangeRead($oWorkbook, Default, $oWorkbook.ActiveSheet.Usedrange,Default,$bForceFunc)
			If @error Then
				If $iClosedBook = 1 Then _Excel_BookClose ( $oWorkbook, False )
				If $iClosedBook	= 1 Then _Excel_Close($oExcel)
				Return SetError (5,0,-1) ;Error reading from workbook
			EndIf
		Else
			$aResult = _Excel_RangeRead($oWorkbook, Default, $oWorkbook.Sheets($iSheets).Usedrange,Default,$bForceFunc)
			If @error Then
				If $iClosedBook = 1 Then _Excel_BookClose ( $oWorkbook, False )
				If $iClosedBook	= 1 Then _Excel_Close($oExcel)
				Return SetError (5,0,-1) ;Error reading from workbook
			EndIf
		EndIf
	Else
		If StringRegExp($sRange, "\d+") = 1 Then
			If $iSheets = Default Then
				$aResult = _Excel_RangeRead($oWorkbook, Default, $oWorkbook.ActiveSheet.Range($sRange),Default,$bForceFunc)
				If @error Then
					If $iClosedBook = 1 Then _Excel_BookClose ( $oWorkbook, False )
					If $iClosedBook	= 1 Then _Excel_Close($oExcel)
					Return SetError (5,0,-1) ;Error reading from workbook
				EndIf
			Else
				$aResult = _Excel_RangeRead($oWorkbook, Default, $oWorkbook.Sheets($iSheets).Range($sRange),Default,$bForceFunc)
				If @error Then
					If $iClosedBook = 1 Then _Excel_BookClose ( $oWorkbook, False )
					If $iClosedBook	= 1 Then _Excel_Close($oExcel)
					Return SetError (5,0,-1) ;Error reading from workbook
				EndIf
			EndIf
		Else
			If $iSheets = Default Then
				$aResult = _Excel_RangeRead($oWorkbook, Default, $oWorkbook.ActiveSheet.Usedrange.Columns($sRange),Default,$bForceFunc)
				If @error Then
					If $iClosedBook = 1 Then _Excel_BookClose ( $oWorkbook, False )
					If $iClosedBook	= 1 Then _Excel_Close($oExcel)
					Return SetError (5,0,-1) ;Error reading from workbook
				EndIf
			Else
				$aResult = _Excel_RangeRead($oWorkbook, Default, $oWorkbook.Sheets($iSheets).Usedrange.Columns($sRange),Default,$bForceFunc)
				If @error Then
					If $iClosedBook = 1 Then _Excel_BookClose ( $oWorkbook, False )
					If $iClosedBook	= 1 Then _Excel_Close($oExcel)
					Return SetError (5,0,-1) ;Error reading from workbook
				EndIf
			EndIf
		EndIf
	EndIf
	If $iClosedBook = 1 Then _Excel_BookClose ( $oWorkbook, False )
	If $iClosedBook	= 1 Then _Excel_Close($oExcel)

	;_ArrayDisplay($aResult)
	Return $aResult

EndFunc

; #FUNCTION __fWriteExcel#
; Write to excel sheet, can be a string, a 1D or 2D zero based array
; Success: Return 1.
; Failure: -1 and sets the @error flag to non-zero.
Func __fWriteExcel($aArray,$sRange = Default,$sExcelPath = "",$iSheets = Default,$bForceFunc = True)

	If IsString ($sRange) = 0 And $sRange <> Default Then Return SetError (1,0,-1) ;$sRange is not string
	If $iSheets <> Default And $iSheets < 1 Then $iSheets = Default
	If Not $bForceFunc = False Then $bForceFunc = True

	Local $oExcel = _Excel_Open(), $iClosedExcel = @extended;, $iClosedBook; $iExists
	If @error Then Return SetError (2,0,-1) ;Error creating excel App
	Local $oWorkbook = _Excel_BookAttach ($sExcelPath)
	If @error Then
		;$iClosedBook = 1
		If FileExists($sExcelPath) = 1 Then
			$oWorkbook  = _Excel_BookOpen($oExcel, $sExcelPath)
			If @error Then
				If $iClosedExcel = 1 Then _Excel_Close($oExcel)
				Return SetError (3,0,-1) ;Error creating workbook
			EndIf
		Else
			$oWorkbook = _Excel_BookNew($oExcel, $iSheets)
			If @error Then
				If $iClosedExcel = 1 Then _Excel_Close($oExcel)
				Return SetError (3,0,-1) ;Error creating workbook
			EndIf
		EndIf
	EndIf
	_Excel_RangeWrite ( $oWorkbook, $iSheets, $aArray, $sRange, Default, $bForceFunc)
	If @error Then Return SetError (4,0,-1) ;Error writing to worksheet
		;If $iClosedBook = 1 Then _Excel_BookClose ( $oWorkbook, False )
		;If $iClosedExcel = 1 Then _Excel_Close($oExcel)
	;	Return SetError (5,0,-1) ;Error writing to worksheet
	;EndIf

	Return 1 ;Success

EndFunc

; #FUNCTION __fImportSQLite#
; Import csv file to sqlite3 db.
; Success: No return
; Failure: -1 and sets the @error flag to non-zero.
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
	If $hInputFile = -1 Then Return SetError (4,0,-1) ;Can not open file
	If FileWrite($hInputFile,$sQuery) = 0 Then Return SetError (5,0,-1) ;Can not write file
	FileClose($hInputFile)
	Local $sCmd = @ComSpec & ' /c ""' & $g_aConfig[$iArraySearch+1]&'sqlite3.exe"' & '  "' _
				 & $sDbPathName _
				 & '" > "' & $sOutput _
				 & '" < "' & $sInput & '""'
	Local $iPID = Run($sCmd, @WorkingDir, @SW_HIDE, BitOR($STDERR_CHILD, $STDOUT_CHILD))
	ProcessWaitClose($iPID)
	Local $sStdout = StdoutRead($iPID)
	Local $sStderr = StderrRead($iPID)
	;MsgBox(64,@error,"Query:"&@CRLF&$sQuery&@CRLF&"#-------#"&@CRLF&"Cmd: "&@CRLF&$sCmd&@CRLF&"#-------#"&@CRLF&"Stdout: "&$sStdout&@CRLF&"Stderr: "&$sStderr)
	If StringLen($sStderr) > 0 Then Return SetError(6,0,-1);Error

	Return $sStderr

EndFunc

; #FUNCTION __fQueryArray#
; Query an array.
; Success: Query
; Failure: -1 and sets the @error flag to non-zero.
;_ArrayDisplay(__fQueryArray(__fReadExcel ("C:\Users\pedro.capelani\AutoIt\GRADES 20062022.xlsx","A2090:I2110"),"select * from array where cast(col2 as numeric) > 3 ; ",False))
Func __fQueryArray($aArray,$sQuery = "", $bHeaders = True)

	If IsArray($aArray) = 0 Then SetError (1,0,-1) ;Is not an array
	Local $iArraySearch = _ArraySearch ($g_aConfig,"d_path")
	If @error Then Return SetError (2,0,-1) ;Auto It can not find d_path in config file
	Local $sInputFile = _TempFile($g_aConfig[$iArraySearch+1]),  $sOutputFile = _TempFile($g_aConfig[$iArraySearch+1]),  $sArrayFile = _TempFile($g_aConfig[$iArraySearch+1])

	If $bHeaders = False Then
		Local $sColName = ""
		For $i = 0 To UBound ($aArray,2)-1
			$sColName = $sColName & "col"&$i&"|"
		Next
		$sColName = StringTrimRight ( $sColName,1)
		_ArrayInsert ($aArray,0,$sColName)
	EndIf
	$hArrayFile = FileOpen($sArrayFile, $FO_OVERWRITE)
	_FileWriteFromArray($hArrayFile,$aArray,Default,Default,";")
	FileClose($hArrayFile)
	$hInputFile = FileOpen($sInputFile, $FO_OVERWRITE)
	$sQuery = ".output stdout"&@CRLF&".headers on"&@CRLF _
					&".mode csv"&@CRLF&".separator ';'"&@CRLF _
					&".import '"&$sArrayFile&"' array"&@CRLF&$sQuery
	FileWrite($hInputFile, $sQuery)
	FileClose($hInputFile)
	Local $sCmd = @ComSpec & ' /c ""' & $g_aConfig[$iArraySearch+1]&'sqlite3.exe"' & '  "' _
				 & '" > "' & $sOutputFile _
				 & '" < "' & $sInputFile & '""'
	Local $iPID = Run($sCmd, @WorkingDir, @SW_HIDE, BitOR($STDERR_CHILD, $STDOUT_CHILD))
	ProcessWaitClose($iPID)
	Local $sStdout = StdoutRead($iPID)
	Local $sStderr = StderrRead($iPID)

	Local $sFileRead = FileRead($sOutputFile)
	;MsgBox (64,"",$sFileRead)
	$sFileRead = StringReplace($sFileRead, '"', '')
	FileDelete($sOutputFile)
	FileWrite($sOutputFile, $sFileRead)
	Local $aResult
	;If UBound($aArray,0) = 1 Then _FileReadToArray($sOutputFile,$aResult,0)
	If UBound($aArray,0) = 2 And StringInStr($sFileRead,';') > 0 Then
		_FileReadToArray($sOutputFile,$aResult,0,";")
	Else
		_FileReadToArray($sOutputFile,$aResult,0)
	EndIf
	FileDelete($sOutputFile)
	FileDelete($sInputFile)
	FileDelete($sArrayFile)
	;MsgBox(64,@error,"Query:"&@CRLF&$sQuery&@CRLF&"#-------#"&@CRLF&"Cmd: "&@CRLF&$sCmd&@CRLF&"#-------#"&@CRLF&"Stdout: "&$sStdout&@CRLF&"Stderr: "&$sStderr)
	If $bHeaders = False Then _ArrayDelete($aResult,0)

	Return $aResult

EndFunc

; #FUNCTION __fImportListView#
; Add 2d array to lisview.
; Success: Return 1.
; Failure: -1 and sets the @error flag to non-zero.
Func __fImportListView($hControl,$aArray, $bHeaders = True,$bColFit= True)

	$hControl = GUICtrlGetHandle($hControl)
	If $hControl = 0 Then SetError (1,0,-1) ;$hControl not found
	Local $aControlGetPos = ControlGetPos("","",$hControl)
	If @error Then Return SetError (1,0,-1) ;$hControl not found
	If IsArray($aArray) = 0 Then SetError (2,0,-1) ;Is not an array
	If Not $bHeaders = False Then $bHeaders = True
	If Not $bColFit = False Then $bColFit = True
	Local $iRow = UBound($aArray)
	Local $iCol = UBound ($aArray,2)
	Local $iGetColumnCount = _GUICtrlListView_GetColumnCount ($hControl)
	Local $iColSize = 0

	_GUICtrlListView_BeginUpdate($hControl)
	_GUICtrlListView_DeleteAllItems ($hControl)
	If $iGetColumnCount  > 0 Then
		For $i = ($iGetColumnCount-1) To 0 Step -1
			_GUICtrlListView_DeleteColumn ($hControl,$i)
		Next
	EndIf
	If $bColFit = True Then
		For $i = 0 To $iCol-1
			If StringLen($aArray[0][$i]) < 3 Then
				$iColSize = $iColSize + 45
			Else
				$iColSize = $iColSize + StringLen($aArray[0][$i])*15+1
			EndIf
		Next
	EndIf
	If $bHeaders = True Then
		For $i = 0 To $iCol-1
			If $bColFit = True Then
				If $iColSize > $aControlGetPos[2] Then
					_GUICtrlListView_AddColumn($hControl, $aArray[0][$i], ($aControlGetPos[2]-20)/$iCol)
				Else
					If StringLen($aArray[0][$i]) < 3 Then
						_GUICtrlListView_AddColumn($hControl, $aArray[0][$i], 45)
					Else
						_GUICtrlListView_AddColumn($hControl, $aArray[0][$i], StringLen($aArray[0][$i])*15+1)
					EndIf
				EndIf
			Else
				If StringLen($aArray[0][$i]) < 3 Then
					_GUICtrlListView_AddColumn($hControl, $aArray[0][$i], 45)
				Else
					_GUICtrlListView_AddColumn($hControl, $aArray[0][$i], StringLen($aArray[0][$i])*15+1)
				EndIf
			EndIf
		Next
	Else
		For $i = 0 To $iCol-1
			If $bColFit = True Then
				If $iColSize > $aControlGetPos[2] Then
					_GUICtrlListView_AddColumn($hControl, "Column "&$i, ($aControlGetPos[2]-20)/$iCol)
				Else
					If StringLen($aArray[0][$i]) < 3 Then
						_GUICtrlListView_AddColumn($hControl, "Column "&$i, 45)
					Else
						_GUICtrlListView_AddColumn($hControl, "Column "&$i, StringLen($aArray[0][$i])*15+1)
					EndIf
				EndIf
			Else
				If StringLen($aArray[0][$i]) < 3 Then
					_GUICtrlListView_AddColumn($hControl, "Column "&$i, 45)
				Else
					_GUICtrlListView_AddColumn($hControl, "Column "&$i, StringLen($aArray[0][$i])*15+1)
				EndIf
			EndIf
		Next
	EndIf
	If $bHeaders = True Then
		_GUICtrlListView_SetItemCount($hControl, $iRow-1)
		_ArrayDelete($aArray,0)
		_GUICtrlListView_AddArray($hControl, $aArray)
	Else
		_GUICtrlListView_SetItemCount($hControl, $iRow)
		_GUICtrlListView_AddArray($hControl, $aArray)
	EndIf
	_GUICtrlListView_EndUpdate($hControl)

	Return 1

EndFunc

; #FUNCTION __fStringFormatArray#
; String formart an 1d or 2d array. Add option to format the intire 2d array.
; Success: Formated array.
; Failure: -1 and sets the @error flag to non-zero.
Func __fStringFormatArray($aArray,$sFormatControl = "", $bHeaders = True,$iColIndex = 0)

	If IsArray($aArray) = 0 Then SetError (1,0,-1) ;Is not an array
	If Not $bHeaders = False Then $bHeaders = True
	If IsInt ($iColIndex) = 0 Or $iColIndex < 0 Then SetError (2,0,-1) ;Is not int
	Local $iRow = UBound($aArray)
	Local $iCol = UBound ($aArray,2)

	If $iCol = 0 Then
		If $bHeaders = True Then
			For $i = 1 To $iRow-1
				$aArray[$i] = StringFormat($sFormatControl,$aArray[$i])
			Next
		ElseIf $bHeaders = False Then
			For $i = 0 To $iRow-1
				$aArray[$i] = StringFormat($sFormatControl,$aArray[$i])
			Next
		EndIf
	ElseIf $iCol > 0 Then
		If $bHeaders = True Then
			For $i = 1 To $iRow-1
				$aArray[$i][$iColIndex] = StringFormat($sFormatControl,$aArray[$i][$iColIndex])
			Next
		ElseIf $bHeaders = False Then
			For $i = 0 To $iRow-1
				$aArray[$i][$iColIndex] = StringFormat($sFormatControl,$aArray[$i][$iColIndex])
			Next
		EndIf
	EndIf

	Return $aArray

EndFunc

; #FUNCTION __fGetAllWinCtrl#
; Return array of all win controls info ;
; Success: Array.
; Failure: -1 and sets the @error flag to non-zero.
; modify version of https://www.autoitscript.com/forum/topic/164226-get-all-windows-controls/#comments from jdelaney
Func __fGetAllWinCtrl($hCallersWindow, $bOnlyVisible=Default, $sStringIncludes=Default, $sClass=Default)

    If Not IsHWnd($hCallersWindow) Then Return SetError (1,0,-1)
    If $bOnlyVisible = Default Then $bOnlyVisible = False
    If $sStringIncludes = Default Then $sStringIncludes = ""
    If $sClass = Default Then $sClass = ""

	; Get all win controls
    $sClassList = WinGetClassList($hCallersWindow)
	If @error Then Return SetError (2,0,-1);Exit MsgBox (64,"ERROR","WinGetClassList not working")
    ; Create array
    $aClassList = StringSplit($sClassList, @CRLF, 2)
    ; Sort array
    _ArraySort($aClassList)
    _ArrayDelete($aClassList, 0)
    ; Loop
    $iCurrentClass = ""
    $iCurrentCount = 1
    $iTotalCounter = 1
    If StringLen($sClass)>0 Then
        For $i = UBound($aClassList)-1 To 0 Step - 1
            If $aClassList[$i]<>$sClass Then
                _ArrayDelete($aClassList,$i)
            EndIf
        Next
    EndIf
	Local $aReturn[UBound($aClassList)+1][10]
	$aReturn [0][0] = "ControlCounter"
	$aReturn [0][1] = "ControlID"
	$aReturn [0][2] = "Handle"
	$aReturn [0][3] = "ClassNN"
	$aReturn [0][4] = "XPos"
	$aReturn [0][5] = "YPos"
	$aReturn [0][6] = "Width"
	$aReturn [0][7] = "Height"
	$aReturn [0][8] = "IsVisible"
	$aReturn [0][9] = "Text"
    For $i = 0 To UBound($aClassList) - 1
        If $aClassList[$i] = $iCurrentClass Then
            $iCurrentCount += 1
        Else
            $iCurrentClass = $aClassList[$i]
            $iCurrentCount = 1
        EndIf
        $hControl = ControlGetHandle($hCallersWindow, "", "[CLASSNN:" & $iCurrentClass & $iCurrentCount & "]")
        $text = StringRegExpReplace(ControlGetText($hCallersWindow, "", $hControl), "[\n\r]", "{@CRLF}")
        $aPos = ControlGetPos($hCallersWindow, "", $hControl)
        $sControlID = _WinAPI_GetDlgCtrlID($hControl)
        $bIsVisible = ControlCommand($hCallersWindow, "", $hControl, "IsVisible")
        If $bOnlyVisible And Not $bIsVisible Then
            $iTotalCounter += 1
            ContinueLoop
        EndIf
        If StringLen($sStringIncludes) > 0 Then
            If Not StringInStr($text, $sStringIncludes) Then
                $iTotalCounter += 1
                ContinueLoop
            EndIf
        EndIf
#cs
        If IsArray($aPos) Then
            ConsoleWrite("Func=[GetAllWindowsControls]: ControlCounter=[" & StringFormat("%3s", $iTotalCounter) & "] ControlID=[" & StringFormat("%5s", $sControlID) & "] Handle=[" & StringFormat("%10s", $hControl) & "] ClassNN=[" & StringFormat("%19s", $iCurrentClass & $iCurrentCount) & "] XPos=[" & StringFormat("%4s", $aPos[0]) & "] YPos=[" & StringFormat("%4s", $aPos[1]) & "] Width=[" & StringFormat("%4s", $aPos[2]) & "] Height=[" & StringFormat("%4s", $aPos[3]) & "] IsVisible=[" & $bIsVisible & "] Text=[" & $text & "]." & @CRLF)
        Else
            ConsoleWrite("Func=[GetAllWindowsControls]: ControlCounter=[" & StringFormat("%3s", $iTotalCounter) & "] ControlID=[" & StringFormat("%5s", $sControlID) & "] Handle=[" & StringFormat("%10s", $hControl) & "] ClassNN=[" & StringFormat("%19s", $iCurrentClass & $iCurrentCount) & "] XPos=[winclosed] YPos=[winclosed] Width=[winclosed] Height=[winclosed] Text=[" & $text & "]." & @CRLF)
        EndIf
#ce
		If IsArray($aPos) Then
			$aReturn[$iTotalCounter][0] = $iTotalCounter
			$aReturn[$iTotalCounter][1] = $sControlID
			$aReturn[$iTotalCounter][2] = $hControl
			$aReturn[$iTotalCounter][3] = $iCurrentClass & $iCurrentCount
			$aReturn[$iTotalCounter][4] = $aPos[0]
			$aReturn[$iTotalCounter][5] = $aPos[1]
			$aReturn[$iTotalCounter][6] = $aPos[2]
			$aReturn[$iTotalCounter][7] = $aPos[3]
			$aReturn[$iTotalCounter][8] = $bIsVisible
			$aReturn[$iTotalCounter][9] = $text
		Else
			$aReturn[$iTotalCounter][0] = $iTotalCounter
			$aReturn[$iTotalCounter][1] = $sControlID
			$aReturn[$iTotalCounter][2] = $hControl
			$aReturn[$iTotalCounter][3] = $iCurrentClass & $iCurrentCount
			$aReturn[$iTotalCounter][4] = "winclosed"
			$aReturn[$iTotalCounter][5] = "winclosed"
			$aReturn[$iTotalCounter][6] = "winclosed"
			$aReturn[$iTotalCounter][7] = "winclosed"
			$aReturn[$iTotalCounter][8] = $bIsVisible
			$aReturn[$iTotalCounter][9] = $text
		EndIf
        If Not WinExists($hCallersWindow) Then ExitLoop
        $iTotalCounter += 1
	Next
	If $bOnlyVisible <> False Or $sStringIncludes <> "" Then
		For $i = UBound($aReturn)-1 To 0 Step - 1
			If $aReturn[$i][0]= "" Then
				_ArrayDelete($aReturn,$i)
            EndIf
		Next
	EndIf

	Return $aReturn

EndFunc   ;==>GetAllWindowsControls

; #FUNCTION __fGetTable#
; Return array with selected db table. Not recommended for multiple function call.;
; Success: 2d Array.
; Failure: -1 and sets the @error flag to non-zero.
;_ArrayDisplay(__fGetTable ("abcd",Default,'select * from abcd where cast("Qtde da Ordem" as numeric) > 2;'))
Func __fGetTable ($sTableName = "",$sDbName = "autoit.db",$sQuery = "")

	If $sTableName = "" Then Return SetError (1,0,-1) ;Empty value
	Local $iArraySearch = _ArraySearch ($g_aConfig,"d_path")
	If @error Then Return SetError (2,0,-1) ;Auto It can not find d_path in config file
	If $sDbName = Default Then $sDbName = "autoit.db"
	If StringRight ($sDbName,3) <> ".db" Then $sDbName = $sDbName&".db"
	Local $sDbPathName = $g_aConfig[$iArraySearch+1]&$sDbName
	If FileExists ($sDbPathName) = 0 Then Return  SetError (3,0,-1) ;Auto It can not find DB
	Local $sDllPath = $g_aConfig[$iArraySearch+1]&"sqlite3.dll"
	If $sQuery = "" Then $sQuery = "SELECT * FROM "&$sTableName&";"

	_SQLite_Startup($sDllPath)
	If @error Then Return SetError (4,0,-1) ;SQLite3.dll Can't be Loaded!
	_SQLite_Open ($sDbPathName)
	If @error Then
		_SQLite_Shutdown()
		Return SetError (5,0,-1) ;Can't create or open a Database
	EndIf
	Local $aResult, $iRows, $iColumns ; $iRows and $iColuums are useless but they cannot be omitted from the function call so we declare them
	$iGetTable2d = _SQLite_GetTable2d(-1, $sQuery , $aResult,$iRows,$iColumns) ; Select single row and single field !
	If $iGetTable2d = $SQLITE_OK Then
		;_ArrayDisplay($aResult, "Results from the query")
		_SQLite_Close()
		_SQLite_Shutdown()
		If UBound ($aResult,2) = 1 Then  $aResult = _ArrayExtract($aResult)
		If @error Then Return SetError (6,0,-1) ; Can not arrayExtract
		Return $aResult
	Else
		;MsgBox(64, "_SQLite_GetTable2d Error: " & $iGetTable2d, _SQLite_ErrMsg()&@CRLF&'__fGetTable error')
		_SQLite_Close()
		_SQLite_Shutdown()
		Return  SetError (7,0,-1) ;Auto It can not find table
	EndIf

EndFunc

; #FUNCTION __fTypeFormatArray#
; Return the string or numeric representation of array data. Add intire 2d array convertion option.
; $iType = 0 number $iType = 1; string
; Success: Converted array.
; Failure: -1 and sets the @error flag to non-zero.
Func __fTypeFormatArray($aArray,$iType = 0, $bHeaders = True,$iColIndex = 0)

	If IsArray($aArray) = 0 Then SetError (1,0,-1) ;Is not an array
	If Not $bHeaders = False Then $bHeaders = True
	If Not $iType = 1 Then $iType = 0
	If IsInt ($iColIndex) = 0 Or $iColIndex < 0 Then SetError (2,0,-1) ;Is not int
	Local $iRow = UBound($aArray)
	Local $iCol = UBound ($aArray,2)

	If $iCol = 0 Then
		If $iType = 0 Then
			If $bHeaders = True Then
				For $i = 1 To $iRow-1
					$aArray[$i] = Number($aArray[$i])
				Next
			ElseIf $bHeaders = False Then
				For $i = 0 To $iRow-1
					$aArray[$i] = Number($aArray[$i])
				Next
			EndIf
		Else
			If $bHeaders = True Then
				For $i = 1 To $iRow-1
					$aArray[$i] = String($aArray[$i])
				Next
			ElseIf $bHeaders = False Then
				For $i = 0 To $iRow-1
					$aArray[$i] = String($aArray[$i])
				Next
			EndIf
		EndIf
	ElseIf $iCol > 0 Then
		If $iType = 0 Then
			If $bHeaders = True Then
				For $i = 1 To $iRow-1
					$aArray[$i][$iColIndex] = Number($aArray[$i][$iColIndex])
				Next
			ElseIf $bHeaders = False Then
				For $i = 0 To $iRow-1
					$aArray[$i][$iColIndex] = Number($aArray[$i][$iColIndex])
				Next
			EndIf
		Else
			If $bHeaders = True Then
				For $i = 1 To $iRow-1
					$aArray[$i][$iColIndex] = String($aArray[$i][$iColIndex])
				Next
			ElseIf $bHeaders = False Then
				For $i = 0 To $iRow-1
					$aArray[$i][$iColIndex] = String($aArray[$i][$iColIndex])
				Next
			EndIf
		EndIf
	EndIf

	Return $aArray

EndFunc




