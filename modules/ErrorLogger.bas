Attribute VB_Name = "ErrorLogger"
Option Explicit

'======================================================================
' ErrorLogger モジュール
' 汎用エラーログ出力モジュール
' エラー情報をCSV形式で C:\VBAErrorLogs\ に出力する
'======================================================================

Private Const LOG_DIR As String = "C:\VBAErrorLogs\"
Private Const LOG_PREFIX As String = "ErrorLog_"
Private Const LOG_EXT As String = ".csv"

'----------------------------------------------------------------------
' LogError: 詳細なエラーログを出力する（メインエントリポイント）
'
' Parameters:
'   moduleName       - エラー発生モジュール名
'   procedureName    - エラー発生プロシージャ名
'   errNumber        - Err.Number
'   errDescription   - Err.Description
'   lineNumber       - Erl（行番号）
'   variableContext  - 変数情報（key=value形式、"; "区切り）
'   processingTarget - 処理対象（省略可）
'   progress         - 進捗情報（省略可）
'----------------------------------------------------------------------
Public Sub LogError( _
    ByVal moduleName As String, _
    ByVal procedureName As String, _
    ByVal errNumber As Long, _
    ByVal errDescription As String, _
    Optional ByVal lineNumber As Long = 0, _
    Optional ByVal variableContext As String = "", _
    Optional ByVal processingTarget As String = "", _
    Optional ByVal progress As String = "")

    On Error Resume Next

    Dim logFilePath As String
    Dim timestamp As String
    Dim fileFullPath As String
    Dim userName As String
    Dim pcName As String
    Dim excelVersion As String
    Dim activeSheetName As String
    Dim errorCategory As String
    Dim logLine As String

    ' タイムスタンプ
    timestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")

    ' ファイルフルパス
    If Not ActiveWorkbook Is Nothing Then
        fileFullPath = ActiveWorkbook.FullName
    Else
        fileFullPath = "(No ActiveWorkbook)"
    End If

    ' ユーザー名
    userName = Environ("USERNAME")

    ' PC名
    pcName = Environ("COMPUTERNAME")

    ' Excelバージョン
    excelVersion = Application.Version & " (" & Application.Build & ")"

    ' アクティブシート名
    If Not ActiveSheet Is Nothing Then
        activeSheetName = ActiveSheet.Name
    Else
        activeSheetName = "(No ActiveSheet)"
    End If

    ' エラー分類（エラー番号に基づく16種類）
    errorCategory = ClassifyError(errNumber)

    ' ログフォルダ作成
    EnsureLogFolder

    ' ログファイルパス
    logFilePath = LOG_DIR & LOG_PREFIX & Format(Date, "yyyymmdd") & LOG_EXT

    ' ヘッダー出力（ファイルが存在しない場合のみ）
    If Dir(logFilePath) = "" Then
        WriteLineUTF8 logFilePath, _
            """Timestamp"",""FileFullPath"",""ModuleName"",""ProcedureName""," & _
            """ErrNumber"",""ErrDescription"",""LineNumber"",""VariableContext""," & _
            """ErrorCategory"",""UserName"",""PCName"",""ExcelVersion""," & _
            """ActiveSheetName"",""ProcessingTarget"",""Progress""", True
    End If

    ' ログ行を組み立て
    logLine = """" & EscapeCsvField(timestamp) & """," & _
              """" & EscapeCsvField(fileFullPath) & """," & _
              """" & EscapeCsvField(moduleName) & """," & _
              """" & EscapeCsvField(procedureName) & """," & _
              """" & CStr(errNumber) & """," & _
              """" & EscapeCsvField(errDescription) & """," & _
              """" & CStr(lineNumber) & """," & _
              """" & EscapeCsvField(variableContext) & """," & _
              """" & EscapeCsvField(errorCategory) & """," & _
              """" & EscapeCsvField(userName) & """," & _
              """" & EscapeCsvField(pcName) & """," & _
              """" & EscapeCsvField(excelVersion) & """," & _
              """" & EscapeCsvField(activeSheetName) & """," & _
              """" & EscapeCsvField(processingTarget) & """," & _
              """" & EscapeCsvField(progress) & """"

    ' ファイルに追記
    WriteLineUTF8 logFilePath, logLine, False

    On Error GoTo 0
End Sub

'----------------------------------------------------------------------
' LogErrorSimple: 簡易エラーログ出力（最小限のパラメータ）
'
' Parameters:
'   moduleName     - エラー発生モジュール名
'   procedureName  - エラー発生プロシージャ名
'----------------------------------------------------------------------
Public Sub LogErrorSimple( _
    ByVal moduleName As String, _
    ByVal procedureName As String)

    LogError moduleName, procedureName, Err.Number, Err.Description, Erl
End Sub

'----------------------------------------------------------------------
' LogSuccess: 正常完了ログを出力する（テスト証跡用）
'
' Parameters:
'   moduleName     - 実行モジュール名
'   procedureName  - 実行プロシージャ名
'   details        - 作業内容の詳細（何を処理したかの説明）
'
' エラーログと同じCSVファイルに出力する。
' エラー番号=0、エラー内容="正常完了"、エラー分類="SUCCESS" として記録し、
' 「いつ、誰が、何を実行して、正常に完了した」をログに残す。
'----------------------------------------------------------------------
Public Sub LogSuccess(ByVal moduleName As String, ByVal procedureName As String, _
    Optional ByVal details As String = "")

    On Error Resume Next

    Dim logFilePath As String
    Dim timestamp As String
    Dim fileFullPath As String
    Dim userName As String
    Dim pcName As String
    Dim excelVersion As String
    Dim activeSheetName As String
    Dim logLine As String

    ' タイムスタンプ
    timestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")

    ' ファイルフルパス（ThisWorkbook を使用）
    If Not ThisWorkbook Is Nothing Then
        fileFullPath = ThisWorkbook.FullName
    Else
        fileFullPath = "(No ThisWorkbook)"
    End If

    ' ユーザー名
    userName = Environ("USERNAME")

    ' PC名
    pcName = Environ("COMPUTERNAME")

    ' Excelバージョン
    excelVersion = Application.Version & " (" & Application.Build & ")"

    ' アクティブシート名
    If Not ActiveSheet Is Nothing Then
        activeSheetName = ActiveSheet.Name
    Else
        activeSheetName = "(No ActiveSheet)"
    End If

    ' ログフォルダ作成
    EnsureLogFolder

    ' ログファイルパス（エラーログと同じファイル）
    logFilePath = LOG_DIR & LOG_PREFIX & Format(Date, "yyyymmdd") & LOG_EXT

    ' ヘッダー出力（ファイルが存在しない場合のみ）
    If Dir(logFilePath) = "" Then
        WriteLineUTF8 logFilePath, _
            """Timestamp"",""FileFullPath"",""ModuleName"",""ProcedureName""," & _
            """ErrNumber"",""ErrDescription"",""LineNumber"",""VariableContext""," & _
            """ErrorCategory"",""UserName"",""PCName"",""ExcelVersion""," & _
            """ActiveSheetName"",""ProcessingTarget"",""Progress""", True
    End If

    ' ログ行を組み立て（エラー番号=0、エラー内容="正常完了"、分類="SUCCESS"）
    logLine = """" & EscapeCsvField(timestamp) & """," & _
              """" & EscapeCsvField(fileFullPath) & """," & _
              """" & EscapeCsvField(moduleName) & """," & _
              """" & EscapeCsvField(procedureName) & """," & _
              """0""," & _
              """正常完了""," & _
              """0""," & _
              """" & EscapeCsvField(details) & """," & _
              """SUCCESS""," & _
              """" & EscapeCsvField(userName) & """," & _
              """" & EscapeCsvField(pcName) & """," & _
              """" & EscapeCsvField(excelVersion) & """," & _
              """" & EscapeCsvField(activeSheetName) & """," & _
              """""," & _
              """"""

    ' ファイルに追記
    WriteLineUTF8 logFilePath, logLine, False

    On Error GoTo 0
End Sub

'----------------------------------------------------------------------
' ClassifyError: エラー番号に基づくエラー分類（16種類）
'----------------------------------------------------------------------
Private Function ClassifyError(ByVal errNumber As Long) As String
    Select Case errNumber
        Case 0
            ClassifyError = "NoError"
        Case 1 To 12
            ClassifyError = "ApplicationError"
        Case 13
            ClassifyError = "TypeMismatch"
        Case 6, 11
            ClassifyError = "MathError"
        Case 7, 14, 28
            ClassifyError = "MemoryError"
        Case 9
            ClassifyError = "IndexOutOfRange"
        Case 48, 49, 51, 53, 54, 55, 57, 58, 59, 61, 62, 63, 67, 68, 70, 71, 74, 75, 76
            ClassifyError = "FileIOError"
        Case 52, 54, 55, 57, 58, 59, 61, 62, 63, 64
            ClassifyError = "FileAccessError"
        Case 91
            ClassifyError = "ObjectNotSet"
        Case 92
            ClassifyError = "LoopError"
        Case 94
            ClassifyError = "NullUsageError"
        Case 429, 430, 432, 438, 440, 443, 445, 446, 447, 448, 449, 450, 451, 452, 453, 457, 458, 459, 460, 461, 462, 463
            ClassifyError = "OLEAutomationError"
        Case 1004
            ClassifyError = "ExcelRuntimeError"
        Case 3001 To 3999
            ClassifyError = "DatabaseError"
        Case -2147467259, -2147217887, -2147217900
            ClassifyError = "ADOError"
        Case -2147024891
            ClassifyError = "PermissionError"
        Case Else
            ClassifyError = "UnclassifiedError"
    End Select
End Function

'----------------------------------------------------------------------
' EnsureLogFolder: ログ出力先フォルダが存在しなければ作成
'----------------------------------------------------------------------
Private Sub EnsureLogFolder()
    On Error Resume Next
    If Dir(LOG_DIR, vbDirectory) = "" Then
        MkDir LOG_DIR
    End If
    On Error GoTo 0
End Sub

'----------------------------------------------------------------------
' WriteLineUTF8: BOM付きUTF-8でファイルに1行書き込む
'   ADODB.Streamを使用
'----------------------------------------------------------------------
Private Sub WriteLineUTF8( _
    ByVal filePath As String, _
    ByVal textLine As String, _
    ByVal isNewFile As Boolean)

    Dim stm As Object
    Dim existingBytes() As Byte
    Dim existingText As String

    If isNewFile Then
        ' 新規ファイル: BOM付きUTF-8で書き込み
        Set stm = CreateObject("ADODB.Stream")
        stm.Type = 2 ' adTypeText
        stm.Charset = "UTF-8"
        stm.Open
        stm.WriteText textLine, 1 ' adWriteLine
        stm.SaveToFile filePath, 2 ' adSaveCreateOverWrite
        stm.Close
        Set stm = Nothing
    Else
        ' 追記: 既存ファイルを読み込んで末尾に追加
        Dim binaryStm As Object
        Set binaryStm = CreateObject("ADODB.Stream")
        binaryStm.Type = 1 ' adTypeBinary
        binaryStm.Open

        ' 既存ファイルをバイナリで読み込み
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")

        If fso.FileExists(filePath) Then
            binaryStm.LoadFromFile filePath
        End If

        ' 新しい行をUTF-8でエンコード
        Dim textStm As Object
        Set textStm = CreateObject("ADODB.Stream")
        textStm.Type = 2 ' adTypeText
        textStm.Charset = "UTF-8"
        textStm.Open
        textStm.WriteText textLine, 1 ' adWriteLine

        ' BOMをスキップしてバイナリに変換
        textStm.Position = 0
        textStm.Type = 1 ' adTypeBinary
        textStm.Position = 3 ' BOM(3バイト)をスキップ

        Dim newBytes() As Byte
        newBytes = textStm.Read
        textStm.Close
        Set textStm = Nothing

        ' 既存データの末尾に追加
        binaryStm.Position = binaryStm.Size
        binaryStm.Write newBytes
        binaryStm.SaveToFile filePath, 2 ' adSaveCreateOverWrite
        binaryStm.Close
        Set binaryStm = Nothing
        Set fso = Nothing
    End If
End Sub

'----------------------------------------------------------------------
' EscapeCsvField: CSV用のフィールドエスケープ
'   ダブルクォートを2重にする
'----------------------------------------------------------------------
Private Function EscapeCsvField(ByVal fieldValue As String) As String
    EscapeCsvField = Replace(fieldValue, """", """""")
End Function
