Attribute VB_Name = "FileLayoutAnalyzer"
Option Explicit

'======================================================================
' FileLayoutAnalyzer モジュール (入口)
' 開発時に受け取ったファイルの構造を、個人情報を出さずにテキストにまとめる。
'
' 対象
'   Excel (xlsx / xlsm / xls / xlsb)   シートごとに表として解析
'   区切り文字形式 (CSV / TSV / DAT など) 区切り文字・引用符・文字コードを判定して解析
'   固定長 (区切りの無い DAT など)      バイト位置ごとの文字の種類から項目を推定
'   XML                                 要素と属性のパスごとに解析
'   その他                              バイナリかどうかと先頭のバイトだけ報告
'
' 出力
'   対象ファイルと同じフォルダに "<ファイル名>_構造分析_<日時>.docx" (Word 文書。A4 横)
'   Word を起動できないときは同じ名前の .txt (BOM 付き UTF-8) に出力する
'   値は載せず、型・桁数・種類数・書式パターンだけを載せる。
'
' 使い方
'   RunFileLayoutAnalysis を実行してファイルを選ぶ (複数選択可)。
'   AnalyzeFile(パス) は 1 ファイルを解析して報告のパスを返す (画面表示なし)。
'
' 依存: ErrorLogger.bas, TextFileReader.bas, GridAnalyzer.bas,
'       FixedLengthAnalyzer.bas, XmlAnalyzer.bas, WordReportWriter.bas
'======================================================================

Private Const MODULE_NAME As String = "FileLayoutAnalyzer"

' ---- 設定 ----
Private Const MAX_ROWS As Long = 100000            ' 1 つの表で解析する最大の行数
Private Const MAGIC_BYTES As Long = 16             ' バイナリのとき報告する先頭のバイト数
Private Const XML_PROBE_CHARS As Long = 200        ' XML かどうかを見る先頭の文字数
Private Const FILE_FILTER As String = "すべてのファイル (*.*),*.*,Excel ブック (*.xlsx;*.xlsm;*.xls;*.xlsb),*.xlsx;*.xlsm;*.xls;*.xlsb,テキスト (*.csv;*.tsv;*.txt;*.dat;*.xml),*.csv;*.tsv;*.txt;*.dat;*.xml"
Private Const EXCEL_EXTENSIONS As String = "|xlsx|xlsm|xls|xlsb|xltx|xltm|"
Private Const REPORT_SUFFIX As String = "_構造分析_"
Private Const REPORT_EXT As String = ".txt"         ' Word を使えないときの出力
Private Const REPORT_EXT_WORD As String = ".docx"
Private Const REPORT_TIME_FORMAT As String = "yyyymmdd_hhnnss"
Private Const RUN_TIME_FORMAT As String = "yyyy-mm-dd hh:nn:ss"
Private Const PATH_SEP As String = "\"
Private Const EXT_SEP As String = "."
Private Const LIST_SEP As String = "|"
Private Const REPORT_CHARSET As String = "UTF-8"
Private Const AD_TYPE_TEXT As Long = 2
Private Const AD_SAVE_CREATE_OVERWRITE As Long = 2
Private Const AD_STATE_CLOSED As Long = 0
Private Const CODE_WIDE_MARK As Long = &HFF2E&
Private Const CODE_HALFKANA_MARK As Long = &HFF76&
Private Const HEX_WIDTH As Long = 2

' ---- 画面の文言 ----
Private Const MSG_DIALOG_TITLE As String = "構造を分析するファイルを選択 (複数可)"
Private Const MSG_CANCELLED As String = "ファイルが選択されなかったため、分析を中止しました。"
Private Const MSG_DONE As String = "分析結果を出力しました。"
Private Const MSG_FAILED_FILES As String = "分析できなかったファイル (C:\VBAErrorLogs\ のログを確認):"
Private Const MSG_FAILED As String = "分析中にエラーが発生しました。C:\VBAErrorLogs\ のエラーログを確認してください。"

' ---- 報告の文言 ----
Private Const RPT_TITLE As String = "ファイル構造分析"
Private Const RPT_RULE As String = "======================================================================"
Private Const RPT_SUB_RULE As String = "----------------------------------------------------------------------"
Private Const RPT_FILE As String = "ファイル名      : "
Private Const RPT_FOLDER As String = "場所            : "
Private Const RPT_SIZE As String = "サイズ          : "
Private Const RPT_BYTES As String = " バイト"
Private Const RPT_MODIFIED As String = "更新日時        : "
Private Const RPT_RUN_AT As String = "分析日時        : "
Private Const RPT_ENCODING As String = "文字コード      : "
Private Const RPT_EOL As String = "改行            : "
Private Const RPT_DELIM As String = "区切り文字      : "
Private Const RPT_DELIM_NONE As String = "なし"
Private Const RPT_DELIM_RATE As String = " (区切りの数が最頻値と同じ行の割合 "
Private Const RPT_MULTILINE As String = "値の中の改行    : "
Private Const RPT_ITEMS_UNIT As String = " 件"
Private Const RPT_SUMMARY As String = "[結論]"
Private Const RPT_LEGEND_HEAD As String = "[凡例] 値は載せず、書式パターンに置き換えている。"
Private Const RPT_LEGEND_1 As String = "  9=数字  A=英大文字  a=英小文字  _=空白  ~=改行  ?=制御文字"
Private Const RPT_LEGEND_WIDE As String = "=全角文字  "
Private Const RPT_LEGEND_KANA As String = "=半角カナ  "
Private Const RPT_LEGEND_2 As String = "記号 (/ - : . , など) と 年月日時分秒 はそのまま残す。9(13) は同じ記号が 13 個続くことを表す。"
Private Const RPT_LEGEND_3 As String = "  表の列: 桁数=文字数、空=空または空白だけの値の数、種類=異なる値の数 (1000 で打ち切り)"
Private Const RPT_EMPTY_FILE As String = "空のファイル (0 バイト)"
Private Const RPT_BINARY As String = "テキストではない (バイナリ)。先頭のバイト: "
Private Const RPT_READ_ERROR As String = "ファイルを読み込めない"
Private Const RPT_XML_FALLBACK As String = "XML として読み込めなかったため、テキストとして解析する。"
Private Const RPT_UTF16_FIXED As String = "UTF-16 の区切りの無いテキストは、行ごとの 1 項目として解析する。"
Private Const RPT_SINGLE_COLUMN As String = "区切り文字も固定長も見当たらないため、1 行を 1 項目として解析する。"
Private Const RPT_SHEET As String = "シート: "
Private Const RPT_SHEET_HIDDEN As String = " (非表示)"
Private Const RPT_SHEET_EMPTY As String = "  空のシート"
Private Const RPT_SHEET_RANGE As String = "  使用範囲        : "
Private Const RPT_SHEET_HIDDEN_RC As String = "  非表示          : 行 "
Private Const RPT_SHEET_HIDDEN_COL As String = " / 列 "
Private Const RPT_SHEET_LISTOBJ As String = "  テーブル (ListObject) : "
Private Const RPT_SHEET_TRUNC As String = "  注: 行が多いため先頭 "
Private Const RPT_SHEET_TRUNC_UNIT As String = " 行だけを解析している"
Private Const RPT_SHEETS As String = "シート数        : "
Private Const RPT_CHARTS As String = " (グラフシート "
Private Const RPT_UNKNOWN_COUNT As String = "不明"
Private Const SUM_EMPTY As String = "空のファイル"
Private Const SUM_BINARY As String = "バイナリ (テキストではない)"
Private Const SUM_READ_ERROR As String = "読み込めない"
Private Const SUM_EXCEL As String = "Excel ブック"
Private Const SUM_SHEET_EMPTY As String = "空"
Private Const SUM_INDENT As String = "  "


'----------------------------------------------------------------------
' RunFileLayoutAnalysis: 入口。ファイルを選んで分析し、報告を出力する
'----------------------------------------------------------------------
Public Sub RunFileLayoutAnalysis()
      On Error GoTo ErrHandler
      Dim selected As Variant
      Dim i As Long
      Dim outPath As String
      Dim doneText As String
      Dim failText As String
      Dim fileCount As Long

10    selected = Application.GetOpenFilename(FILE_FILTER, , MSG_DIALOG_TITLE, , True)
20    If Not IsArray(selected) Then
30        MsgBox MSG_CANCELLED, vbInformation
          LogSuccess MODULE_NAME, "RunFileLayoutAnalysis", "Cancelled: no file selected"
40        Exit Sub
      End If

50    Application.ScreenUpdating = False
60    For i = LBound(selected) To UBound(selected)
70        fileCount = fileCount + 1
80        Application.StatusBar = MSG_DIALOG_TITLE & " " & fileCount & "/" & (UBound(selected) - LBound(selected) + 1)
90        outPath = AnalyzeFile(CStr(selected(i)))
100       If Len(outPath) > 0 Then
110           doneText = doneText & outPath & vbCrLf
          Else
120           failText = failText & CStr(selected(i)) & vbCrLf
          End If
      Next i
130   Application.StatusBar = False
140   Application.ScreenUpdating = True

150   MsgBox MSG_DONE & vbCrLf & doneText & _
             IIf(Len(failText) > 0, vbCrLf & MSG_FAILED_FILES & vbCrLf & failText, ""), _
             IIf(Len(failText) > 0, vbExclamation, vbInformation)

      LogSuccess MODULE_NAME, "RunFileLayoutAnalysis", "Analyzed " & fileCount & " files"
160   Exit Sub
ErrHandler:
      LogError MODULE_NAME, "RunFileLayoutAnalysis", Err.Number, Err.Description, Erl, _
               "fileCount=" & fileCount & "; outPath=" & outPath
      Application.StatusBar = False
      Application.ScreenUpdating = True
      MsgBox MSG_FAILED, vbExclamation
End Sub

'----------------------------------------------------------------------
' AnalyzeFile: 1 ファイルを分析し、報告を書き出して、そのパスを返す
'   失敗したら空文字を返す (画面には何も出さない)
'----------------------------------------------------------------------
Public Function AnalyzeFile(ByVal filePath As String) As String
      On Error GoTo ErrHandler
      Dim fso As Object
      Dim fileObj As Object
      Dim folderPath As String
      Dim fileName As String
      Dim ext As String
      Dim header As String
      Dim body As String
      Dim summary As String
      Dim report As String
      Dim outPath As String
      Dim outBase As String

10    Set fso = CreateObject("Scripting.FileSystemObject")
20    Set fileObj = fso.GetFile(filePath)
30    folderPath = fso.GetParentFolderName(filePath)
40    fileName = fso.GetFileName(filePath)
50    ext = LCase$(fso.GetExtensionName(filePath))

60    header = RPT_TITLE & vbCrLf & RPT_RULE & vbCrLf & _
               RPT_FILE & fileName & vbCrLf & _
               RPT_FOLDER & folderPath & vbCrLf & _
               RPT_SIZE & Format(fileObj.Size, "#,##0") & RPT_BYTES & vbCrLf & _
               RPT_MODIFIED & Format(fileObj.DateLastModified, RUN_TIME_FORMAT) & vbCrLf & _
               RPT_RUN_AT & Format(Now, RUN_TIME_FORMAT) & vbCrLf

70    If InStr(EXCEL_EXTENSIONS, LIST_SEP & ext & LIST_SEP) > 0 Then
80        body = AnalyzeExcelFile(filePath, header, summary)
      Else
90        body = AnalyzeTextFile(filePath, header, summary)
      End If

100   report = header & vbCrLf & BuildLegend() & vbCrLf & _
               RPT_SUMMARY & vbCrLf & summary & vbCrLf & vbCrLf & _
               body
110   outBase = folderPath & PATH_SEP & fileName & REPORT_SUFFIX & Format(Now, REPORT_TIME_FORMAT)
120   If WordReportWriter.WriteWordReport(outBase & REPORT_EXT_WORD, report) Then
130       outPath = outBase & REPORT_EXT_WORD
      Else
          ' Word を起動できないときなどは、テキストで出力する
140       outPath = outBase & REPORT_EXT
150       WriteReport outPath, report
      End If
160   AnalyzeFile = outPath

      LogSuccess MODULE_NAME, "AnalyzeFile", "Analyzed " & filePath & "; report=" & outPath
170   Exit Function
ErrHandler:
      LogError MODULE_NAME, "AnalyzeFile", Err.Number, Err.Description, Erl, _
               "filePath=" & filePath & "; outPath=" & outPath
      AnalyzeFile = ""
End Function

'----------------------------------------------------------------------
' AnalyzeTextFile: Excel 以外のファイルを、XML / 区切り文字形式 / 固定長 の順に判定して解析する
'   header には文字コードなどの行を足す
'----------------------------------------------------------------------
Private Function AnalyzeTextFile(ByVal filePath As String, ByRef header As String, ByRef summary As String) As String
      On Error GoTo ErrHandler
      Dim data() As Byte
      Dim byteCount As Long
      Dim charset As String
      Dim encLabel As String
      Dim bomBytes As Long
      Dim text As String
      Dim probe As String
      Dim body As String
      Dim parsed As Boolean
      Dim lines() As String
      Dim eolLabel As String
      Dim delim As String
      Dim delimLabel As String
      Dim consistency As Double
      Dim grid As Variant
      Dim rowWidth() As Long
      Dim rowCount As Long
      Dim colCount As Long
      Dim quotedCount() As Long
      Dim truncated As Boolean
      Dim multiLine As Long
      Dim isFixed As Boolean
      Dim i As Long
      Dim magic As String
      Dim sumPart As String
      Dim encHint As String

10    data = TextFileReader.ReadAllBytes(filePath, byteCount)
20    If byteCount < 0 Then
30        summary = SUM_INDENT & SUM_READ_ERROR
40        AnalyzeTextFile = RPT_READ_ERROR & vbCrLf
          LogSuccess MODULE_NAME, "AnalyzeTextFile", "Could not read " & filePath
50        Exit Function
      End If
60    If byteCount = 0 Then
70        summary = SUM_INDENT & SUM_EMPTY
80        AnalyzeTextFile = RPT_EMPTY_FILE & vbCrLf
          LogSuccess MODULE_NAME, "AnalyzeTextFile", "Empty file " & filePath
90        Exit Function
      End If

100   charset = TextFileReader.DetectEncoding(data, byteCount, encLabel, bomBytes, encHint)
110   header = header & RPT_ENCODING & encLabel & vbCrLf
      ' EUC-JP / JIS / 不明 のときは、判定の手がかりを本文の先頭に載せる
120   body = encHint

      ' ---- バイナリ ----
130   If Len(charset) = 0 Then
140       For i = 0 To MAGIC_BYTES - 1
150           If i >= byteCount Then Exit For
160           magic = magic & Right$("0" & Hex$(data(i)), HEX_WIDTH) & " "
          Next i
170       summary = SUM_INDENT & SUM_BINARY
180       AnalyzeTextFile = RPT_BINARY & Trim$(magic) & vbCrLf
          LogSuccess MODULE_NAME, "AnalyzeTextFile", "Binary file " & filePath
190       Exit Function
      End If

200   text = TextFileReader.DecodeBytes(data, charset)

      ' ---- XML ----
210   probe = Left$(text, XML_PROBE_CHARS)
220   probe = Replace(Replace(Replace(Replace(probe, vbCr, ""), vbLf, ""), vbTab, ""), " ", "")
230   If Left$(probe, 1) = "<" Then
240       body = XmlAnalyzer.AnalyzeXml(filePath, parsed, sumPart)
250       If parsed Then
260           summary = SUM_INDENT & sumPart
270           AnalyzeTextFile = body
              LogSuccess MODULE_NAME, "AnalyzeTextFile", "Analyzed as XML: " & filePath
280           Exit Function
          End If
290       body = body & RPT_XML_FALLBACK & vbCrLf & vbCrLf
      End If

      ' ---- 区切り文字形式 ----
300   lines = TextFileReader.SplitLines(text, eolLabel)
310   header = header & RPT_EOL & eolLabel & vbCrLf
320   delim = TextFileReader.DetectDelimiter(lines, delimLabel, consistency)
330   If Len(delim) > 0 Then
340       grid = TextFileReader.ParseDelimited(lines, delim, MAX_ROWS, rowWidth, rowCount, colCount, _
                                               quotedCount, truncated, multiLine)
350       header = header & RPT_DELIM & delimLabel & RPT_DELIM_RATE & Format(consistency, "0.0%") & ")" & vbCrLf & _
                   RPT_MULTILINE & multiLine & RPT_ITEMS_UNIT & vbCrLf
360       body = body & GridAnalyzer.AnalyzeGrid(grid, rowCount, colCount, rowWidth, quotedCount, Nothing, True, truncated, sumPart)
370       summary = SUM_INDENT & sumPart
380       AnalyzeTextFile = body
          LogSuccess MODULE_NAME, "AnalyzeTextFile", "Analyzed as delimited (" & delimLabel & "): " & filePath
390       Exit Function
      End If
400   header = header & RPT_DELIM & RPT_DELIM_NONE & vbCrLf

      ' ---- 固定長 (UTF-16 はバイト位置が使えないので 1 項目として扱う) ----
410   If charset <> TextFileReader.CHARSET_UTF16LE And charset <> TextFileReader.CHARSET_UTF16BE Then
420       body = body & FixedLengthAnalyzer.AnalyzeFixedLength(data, byteCount, bomBytes, _
                                                               (charset = TextFileReader.CHARSET_SJIS), isFixed, sumPart)
430       If isFixed Then
440           summary = SUM_INDENT & sumPart
450           AnalyzeTextFile = body
              LogSuccess MODULE_NAME, "AnalyzeTextFile", "Analyzed as fixed length: " & filePath
460           Exit Function
          End If
470       body = body & vbCrLf & RPT_SINGLE_COLUMN & vbCrLf & vbCrLf
      Else
480       body = body & RPT_UTF16_FIXED & vbCrLf & vbCrLf
      End If

      ' ---- 1 行 = 1 項目 ----
490   grid = TextFileReader.LinesToSingleColumn(lines, MAX_ROWS, rowWidth, rowCount, truncated)
500   ReDim quotedCount(1 To 1)
510   body = body & GridAnalyzer.AnalyzeGrid(grid, rowCount, 1, rowWidth, quotedCount, Nothing, True, truncated, sumPart)
520   summary = SUM_INDENT & sumPart
530   AnalyzeTextFile = body

      LogSuccess MODULE_NAME, "AnalyzeTextFile", "Analyzed as single column: " & filePath
540   Exit Function
ErrHandler:
      LogError MODULE_NAME, "AnalyzeTextFile", Err.Number, Err.Description, Erl, _
               "filePath=" & filePath & "; byteCount=" & byteCount & "; charset=" & charset & "; delim=" & delimLabel
      summary = SUM_INDENT & SUM_READ_ERROR
      AnalyzeTextFile = body & MSG_FAILED & vbCrLf
End Function

'----------------------------------------------------------------------
' AnalyzeExcelFile: Excel ブックをシートごとに解析する
'   対象は読み取り専用で開き、保存せずに閉じる (既に開いていればそのまま使い、閉じない)
'----------------------------------------------------------------------
Private Function AnalyzeExcelFile(ByVal filePath As String, ByRef header As String, ByRef summary As String) As String
      On Error GoTo ErrHandler
      Dim wb As Workbook
      Dim openedHere As Boolean
      Dim ws As Worksheet
      Dim body As String
      Dim lastCell As Range
      Dim lastRow As Long
      Dim lastCol As Long
      Dim readRows As Long
      Dim truncated As Boolean
      Dim vals As Variant
      Dim rowWidth() As Long
      Dim quotedCount() As Long
      Dim sumPart As String
      Dim hiddenRows As Long
      Dim hiddenCols As Long
      Dim prevAlerts As Boolean

10    prevAlerts = Application.DisplayAlerts
20    Application.DisplayAlerts = False
30    Set wb = FindOpenWorkbook(filePath)
40    If wb Is Nothing Then
50        Set wb = Workbooks.Open(Filename:=filePath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
60        openedHere = True
      End If

70    header = header & RPT_SHEETS & wb.Worksheets.Count & _
               IIf(wb.Charts.Count > 0, RPT_CHARTS & wb.Charts.Count & ")", "") & vbCrLf

80    For Each ws In wb.Worksheets
90        body = body & RPT_SUB_RULE & vbCrLf & RPT_SHEET & ws.Name & _
                 IIf(ws.Visible <> xlSheetVisible, RPT_SHEET_HIDDEN, "") & vbCrLf
100       Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                                       SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
110       If lastCell Is Nothing Then
120           body = body & RPT_SHEET_EMPTY & vbCrLf & vbCrLf
130           summary = summary & SUM_INDENT & ws.Name & ": " & SUM_SHEET_EMPTY & vbCrLf
          Else
140           lastRow = lastCell.Row
150           lastCol = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                                      SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
160           readRows = lastRow
170           truncated = False
180           If readRows > MAX_ROWS Then
190               readRows = MAX_ROWS
200               truncated = True
              End If
210           If readRows = 1 And lastCol = 1 Then
220               ReDim vals(1 To 1, 1 To 1)
230               vals(1, 1) = ws.Cells(1, 1).Value
              Else
240               vals = ws.Range(ws.Cells(1, 1), ws.Cells(readRows, lastCol)).Value
              End If
250           hiddenRows = CountHidden(ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 1)), lastRow)
260           hiddenCols = CountHidden(ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)), lastCol)

270           body = body & RPT_SHEET_RANGE & "A1:" & GridAnalyzer.ColumnRef(lastCol, False) & lastRow & vbCrLf & _
                     RPT_SHEET_HIDDEN_RC & IIf(hiddenRows < 0, RPT_UNKNOWN_COUNT, CStr(hiddenRows)) & _
                     RPT_SHEET_HIDDEN_COL & IIf(hiddenCols < 0, RPT_UNKNOWN_COUNT, CStr(hiddenCols)) & vbCrLf & _
                     RPT_SHEET_LISTOBJ & ws.ListObjects.Count & vbCrLf
280           If truncated Then body = body & RPT_SHEET_TRUNC & MAX_ROWS & RPT_SHEET_TRUNC_UNIT & vbCrLf
290           body = body & vbCrLf

300           ReDim rowWidth(1 To 1)
310           ReDim quotedCount(1 To 1)
320           body = body & GridAnalyzer.AnalyzeGrid(vals, readRows, lastCol, rowWidth, quotedCount, ws, False, truncated, sumPart) & vbCrLf
330           summary = summary & SUM_INDENT & ws.Name & ": " & sumPart & vbCrLf
          End If
      Next ws
340   If Len(summary) > 0 Then summary = Left$(summary, Len(summary) - Len(vbCrLf))
350   AnalyzeExcelFile = body

      LogSuccess MODULE_NAME, "AnalyzeExcelFile", "Analyzed " & wb.Worksheets.Count & " sheets in " & filePath

CleanUp:
      On Error Resume Next
      If openedHere Then
          If Not wb Is Nothing Then wb.Close SaveChanges:=False
      End If
      Set wb = Nothing
      Application.DisplayAlerts = prevAlerts
      On Error GoTo 0
      Exit Function

ErrHandler:
      LogError MODULE_NAME, "AnalyzeExcelFile", Err.Number, Err.Description, Erl, _
               "filePath=" & filePath & "; lastRow=" & lastRow & "; lastCol=" & lastCol
      summary = summary & SUM_INDENT & SUM_READ_ERROR
      AnalyzeExcelFile = body & MSG_FAILED & vbCrLf
      Resume CleanUp
End Function

'----------------------------------------------------------------------
' CountHidden: 範囲 (1 行または 1 列) の非表示のセル数を返す (数えられなければ -1)
'----------------------------------------------------------------------
Private Function CountHidden(ByVal rng As Range, ByVal total As Long) As Long
      On Error GoTo ErrHandler

10    CountHidden = total - rng.SpecialCells(xlCellTypeVisible).Cells.Count

      LogSuccess MODULE_NAME, "CountHidden", "Counted hidden cells in " & rng.Address(False, False)
20    Exit Function
ErrHandler:
      ' すべて非表示のときも SpecialCells はエラーになる
      LogError MODULE_NAME, "CountHidden", Err.Number, Err.Description, Erl, "total=" & total
      CountHidden = -1
End Function

'----------------------------------------------------------------------
' FindOpenWorkbook: 既に開いているブックならそれを返す (無ければ Nothing)
'----------------------------------------------------------------------
Private Function FindOpenWorkbook(ByVal fullPath As String) As Workbook
      On Error GoTo ErrHandler
      Dim wb As Workbook

10    For Each wb In Application.Workbooks
20        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
30            Set FindOpenWorkbook = wb
40            Exit For
          End If
      Next wb

      LogSuccess MODULE_NAME, "FindOpenWorkbook", "alreadyOpen=" & CStr(Not FindOpenWorkbook Is Nothing)
50    Exit Function
ErrHandler:
      LogError MODULE_NAME, "FindOpenWorkbook", Err.Number, Err.Description, Erl, "fullPath=" & fullPath
      Set FindOpenWorkbook = Nothing
End Function

'----------------------------------------------------------------------
' BuildLegend: 書式パターンの凡例 (全角・半角カナの記号は実行時に作る)
'----------------------------------------------------------------------
Private Function BuildLegend() As String
      On Error GoTo ErrHandler

10    BuildLegend = RPT_LEGEND_HEAD & vbCrLf & RPT_LEGEND_1 & vbCrLf & _
                    "  " & ChrW(CODE_WIDE_MARK) & RPT_LEGEND_WIDE & ChrW(CODE_HALFKANA_MARK) & RPT_LEGEND_KANA & vbCrLf & _
                    "  " & RPT_LEGEND_2 & vbCrLf & RPT_LEGEND_3 & vbCrLf

      LogSuccess MODULE_NAME, "BuildLegend", "Built legend"
20    Exit Function
ErrHandler:
      LogError MODULE_NAME, "BuildLegend", Err.Number, Err.Description, Erl, ""
      BuildLegend = ""
End Function

'----------------------------------------------------------------------
' WriteReport: 報告を BOM 付き UTF-8 で書き出す (既存は上書き)
'   失敗したら呼び出し元へエラーを返す
'----------------------------------------------------------------------
Private Sub WriteReport(ByVal filePath As String, ByVal reportText As String)
      On Error GoTo ErrHandler
      Dim stm As Object
      Dim errNo As Long
      Dim errDesc As String

10    Set stm = CreateObject("ADODB.Stream")
20    stm.Type = AD_TYPE_TEXT
30    stm.Charset = REPORT_CHARSET
40    stm.Open
50    stm.WriteText reportText
60    stm.SaveToFile filePath, AD_SAVE_CREATE_OVERWRITE

      LogSuccess MODULE_NAME, "WriteReport", "Wrote report (" & Len(reportText) & " chars) to " & filePath

CleanUp:
      On Error Resume Next
      If Not stm Is Nothing Then
          If stm.State <> AD_STATE_CLOSED Then stm.Close
          Set stm = Nothing
      End If
      On Error GoTo 0
      If errNo <> 0 Then Err.Raise errNo, MODULE_NAME & ".WriteReport", errDesc
      Exit Sub

ErrHandler:
      ' LogError は Err を消すので、先に控えておく
      errNo = Err.Number
      errDesc = Err.Description
      LogError MODULE_NAME, "WriteReport", errNo, errDesc, Erl, "filePath=" & filePath
      Resume CleanUp
End Sub
