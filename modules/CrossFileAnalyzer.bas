Attribute VB_Name = "CrossFileAnalyzer"
Option Explicit

'======================================================================
' CrossFileAnalyzer モジュール
' 複数のファイルを読み、ファイルをまたいで同じデータが入っている列 (結合キーの候補) を値から割り出し、
' 「どのファイルの何列目に何が入っていて、どの列どうしを結べばよいか」を Word の対応表にまとめる。
'   - 区切り文字形式 (CSV など) を読む。先頭項目が種別コードなら、種別ごとの表に分ける
'   - 列ごとの手がかり: 検証番号 (保険者番号・公費負担者番号・受給者番号の形)、和暦、西暦日付、桁数
'   - 列どうしの共通の値の割合 (数字は先頭の 0 を除いて比べる)
' 値そのものは報告に出さず、共通の値の数と割合だけを出す。値はメモリ上の辞書でだけ使う。
'
' 使い方
'   RunCrossFileAnalysis を実行し、ファイルを複数選ぶ。
'   AnalyzeFileSet(パスを | で区切った文字列) は報告のパスを返す (画面表示なし)。
'
' 依存: ErrorLogger.bas, TextFileReader.bas, WordReportWriter.bas
'======================================================================

Private Const MODULE_NAME As String = "CrossFileAnalyzer"

' ---- 設定 ----
Private Const MAX_ROWS As Long = 100000            ' 1 ファイルで読む最大の行数
Private Const MAX_COLS As Long = 5000              ' 登録する列 (ファイル x 種別 x 列) の最大数
Private Const DISTINCT_CAP As Long = 20000         ' 列ごとに覚える値の種類の上限
Private Const CHECK_CAP As Long = 3000             ' 共通の値を数えるときに調べる値の数の上限
Private Const MIN_KEY_KINDS As Long = 10           ' 結合キーの候補とする列の最小の種類数
Private Const MIN_COMMON As Long = 10              ' 結合キーとみなす共通の値の最小数
Private Const MATCH_SHARE As Double = 0.5          ' どちらかの列の値のうち、共通の値がこの割合以上
Private Const KIND_MAX_LEN As Long = 4             ' 種別コードとみなす先頭項目の最大の文字数
Private Const KIND_MAX_KINDS As Long = 30          ' 種別コードとみなす最大の種類数
Private Const HINT_SHARE As Double = 0.95          ' 手がかり (検証番号・日付など) とみなす値の割合
Private Const HINT_MIN_VALUES As Long = 10         ' 手がかりを判定する最小の値の数
Private Const MAX_MATCH_ROWS As Long = 80          ' 報告する結合キー候補の最大数
Private Const SEQ_SPAN As Double = 1.5             ' 最大値が種類数のこの倍以下なら、1 から順の連番とみなす
Private Const SEQ_PENALTY As Double = 0.01         ' 連番が絡む一致の並び順を下げる係数
Private Const UNIQUE_BONUS As Double = 2           ' 両方の列が全件一意 (1 対 1 で結べる) なら並び順を上げる係数
Private Const ID_BONUS As Double = 1.5             ' 番号の形 (長さが一定で全件一意) の列どうしなら、さらに上げる係数
Private Const ID_MIN_LEN As Long = 10              ' レセプトなどを特定する番号とみなす最小の桁数
Private Const MAX_NUM_LEN As Long = 15             ' 数値として最大値を覚える最大の桁数
Private Const GUIDE_PER_PAIR As Long = 8           ' ファイルの組ごとに要約する候補の数
Private Const LIST_SEP As String = "|"
Private Const PATH_SEP As String = "\"
Private Const KEY_SEP As String = vbTab
Private Const NO_KIND As String = "-"

' ---- 検証番号・日付の判定 ----
Private Const CODE_DIGIT_0 As Long = 48
Private Const LUHN_W_FIRST As Long = 2             ' 末尾の前の桁から 2, 1, 2, 1 ... を掛ける
Private Const LUHN_W_SUM As Long = 3               ' 2 と 1 を入れ替えるための和
Private Const DECIMAL_BASE As Long = 10
Private Const DIGIT_CARRY As Long = 9              ' 積が 2 桁なら各桁の和 (= 積 - 9)
Private Const LEN_KOKUHO As Long = 6               ' 国保の保険者番号
Private Const LEN_JUKYU As Long = 7                ' 受給者番号など
Private Const LEN_HOKEN As Long = 8                ' 保険者番号・公費負担者番号
Private Const LEN_WAREKI_YM As Long = 5
Private Const LEN_WAREKI_YMD As Long = 7
Private Const LEN_DATE8 As Long = 8
Private Const ERA_MIN As Long = 1
Private Const ERA_MAX As Long = 5
Private Const MONTH_MAX As Long = 12
Private Const DAY_MAX As Long = 31
Private Const YEAR_MIN As Long = 1900
Private Const YEAR_MAX As Long = 2100
Private Const YEAR_DIGITS As Long = 4

' ---- 出力 ----
Private Const REPORT_NAME As String = "ファイル間対応_"
Private Const REPORT_TIME_FORMAT As String = "yyyymmdd_hhnnss"
Private Const REPORT_EXT_WORD As String = ".docx"
Private Const REPORT_EXT_TEXT As String = ".txt"
Private Const REPORT_CHARSET As String = "UTF-8"
Private Const AD_TYPE_TEXT As Long = 2
Private Const AD_SAVE_CREATE_OVERWRITE As Long = 2
Private Const RUN_TIME_FORMAT As String = "yyyy-mm-dd hh:nn:ss"

' ---- 画面の文言 ----
Private Const MSG_PICK_TITLE As String = "対応を調べるファイルを選択 (2 つ以上。Ctrl / Shift で複数選択)"
Private Const MSG_TOO_FEW As String = "ファイルを 2 つ以上選んでください。"
Private Const MSG_DONE As String = "ファイル間の対応表を出力しました。"
Private Const MSG_FAILED As String = "処理中にエラーが発生しました。C:\VBAErrorLogs\ のエラーログを確認してください。"

' ---- 報告の文言 ----
Private Const RPT_TITLE As String = "ファイル間の対応表"
Private Const RPT_RULE As String = "======================================================================"
Private Const RPT_RUN_AT As String = "分析日時        : "
Private Const RPT_FILE_COUNT As String = "ファイル数      : "
Private Const RPT_LEGEND As String = "[この表の読み方]"
Private Const RPT_LEGEND_1 As String = "  値は載せず、列どうしで共通する値の数と割合だけを載せる。数字は先頭の 0 を除いて比べる。"
Private Const RPT_LEGEND_2 As String = "  列は「F番号 / 種別 / No.列番号」で表す。種別は先頭項目 (レコード種別) の値。種別の無いファイルは -。"
Private Const RPT_LEGEND_3 As String = "  共通の値の割合が両方とも高い列どうしは、同じデータ (結合キー) である可能性が高い。片方だけ高いときは、一方が他方の一部 (月ごとの患者など)。"
Private Const RPT_SEC_FILES As String = "[対象のファイル]"
Private Const RPT_FILES_HEAD As String = "| F | ファイル | 文字コード | 区切り | 行数 | 種別 (行数 / 項目数) |"
Private Const RPT_FILES_RULE As String = "|---:|---|---|---|---:|---|"
Private Const RPT_SKIP As String = "区切り文字形式ではないため対象外 (Excel・固定長・XML は未対応)"
Private Const RPT_READ_ERROR As String = "読み込めない"
Private Const RPT_SEC_HINTS As String = "[列の手がかり (値の形から分かること)]"
Private Const RPT_HINTS_HEAD As String = "| 列 | 埋まっている行 | 値の種類 | 桁数 | 手がかり |"
Private Const RPT_HINTS_RULE As String = "|---|---:|---:|---|---|"
Private Const RPT_SEC_MATCH As String = "[結合キーの候補 (ファイルをまたいで同じ値が入っている列)]"
Private Const RPT_MATCH_HEAD As String = "| 列 A | 列 B | 共通の値 | A の値のうち | B の値のうち | 列 A の手がかり |"
Private Const RPT_MATCH_RULE As String = "|---|---|---:|---:|---:|---|"
Private Const RPT_NONE As String = "  見つからない"
Private Const RPT_MORE As String = "  注: 候補が多いため、共通の値の多い順に先頭だけを載せている"
Private Const RPT_NOTE_MONTH As String = "  注: 請求月が違うファイルどうしでも、同じ患者・同じ保険の列 (保険者番号・受給者番号・生年月日など) は共通の値を持つ。レセプト番号のような月ごとの番号は、同じ請求月のファイルどうしでないと一致しない。"
Private Const RPT_ROWS_SEP As String = " 行 / "
Private Const RPT_COLS_UNIT As String = " 項目"
Private Const RPT_KIND_SEP As String = ", "
Private Const RPT_TEXT_COL As String = "文字"
Private Const RPT_DIGIT_COL As String = "数字 "
Private Const RPT_LEN_UNIT As String = " 桁"
Private Const RPT_PCT As String = "%"
Private Const RPT_ITEM_PREFIX As String = "F"
Private Const RPT_ITEM_SEP As String = " / "
Private Const RPT_ITEM_NO As String = "No."
Private Const HINT_HOKEN As String = "検証番号が合う 8 桁 (保険者番号・公費負担者番号の形)"
Private Const HINT_KOKUHO As String = "検証番号が合う 6 桁 (国保の保険者番号の形)"
Private Const HINT_JUKYU As String = "検証番号が合う 7 桁 (公費の受給者番号などの形)"
Private Const HINT_WAREKI_YM As String = "和暦年月 (GYYMM)"
Private Const HINT_WAREKI_YMD As String = "和暦年月日 (GYYMMDD。生年月日など)"
Private Const HINT_DATE8 As String = "西暦の日付 (yyyymmdd)"
Private Const HINT_KEY As String = "全件一意"
Private Const HINT_FIXED As String = "固定値"
Private Const HINT_SEP As String = "; "
Private Const HINT_SEQ As String = "連番 (1 から順の番号)"
Private Const HINT_ID As String = "長さが一定で全件一意 (レセプトなどを特定する番号の形)"
Private Const RPT_SEQ_MARK As String = "連番どうしの偶然の一致の可能性が高い"
Private Const RPT_SEC_GUIDE As String = "[ファイルの組ごとの結び方 (連番を除き、強い順)]"
Private Const RPT_GUIDE_PAIR As String = "  F"
Private Const RPT_GUIDE_AND As String = " と F"
Private Const RPT_GUIDE_COLON As String = ":"
Private Const RPT_GUIDE_ITEM As String = "    "
Private Const RPT_GUIDE_ARROW As String = " = "
Private Const RPT_GUIDE_SHARE As String = " (共通 "
Private Const RPT_GUIDE_SHARE_2 As String = " / "
Private Const RPT_GUIDE_NONE As String = "    結べる列は見当たらない (請求月が違う、または共通のデータが無い)"
Private Const RPT_MATCH_NOTE_HEAD As String = "| 列 A | 列 B | 共通の値 | A の値のうち | B の値のうち | 列 A の手がかり | 注意 |"
Private Const RPT_MATCH_NOTE_RULE As String = "|---|---|---:|---:|---:|---|---|"

' ---- 列の情報 (登録順) ----
Private m_colCount As Long
Private m_colFile() As Long
Private m_colKind() As String
Private m_colNo() As Long
Private m_colVals() As Object       ' 正規化した値 -> True (DISTINCT_CAP まで)
Private m_colFilled() As Long
Private m_colDigits() As Boolean    ' ほぼすべて (HINT_SHARE 以上) 数字か
Private m_colDigitN() As Long       ' 数字だけの値の数
Private m_colMaxNum() As Double     ' 数字の値の最大 (連番の判定用)
Private m_colSeq() As Boolean       ' 1 から順の連番か
Private m_colMinLen() As Long
Private m_colMaxLen() As Long
Private m_colLuhn() As Long         ' 検証番号が合う値の数 (6-8 桁の数字)
Private m_colWarekiYm() As Long
Private m_colWarekiYmd() As Long
Private m_colDate8() As Long
Private m_colHint() As String


'----------------------------------------------------------------------
' RunCrossFileAnalysis: 入口。ファイルを複数選び、対応表を出力する
'----------------------------------------------------------------------
Public Sub RunCrossFileAnalysis()
      On Error GoTo ErrHandler
      Dim selected As Variant
      Dim pathsText As String
      Dim i As Long
      Dim outPath As String

10    selected = Application.GetOpenFilename(, , MSG_PICK_TITLE, , True)
20    If Not IsArray(selected) Then
          LogSuccess MODULE_NAME, "RunCrossFileAnalysis", "Cancelled: no file selected"
30        Exit Sub
      End If
40    If UBound(selected) - LBound(selected) + 1 < 2 Then
50        MsgBox MSG_TOO_FEW, vbExclamation
          LogSuccess MODULE_NAME, "RunCrossFileAnalysis", "Cancelled: fewer than 2 files"
60        Exit Sub
      End If
70    For i = LBound(selected) To UBound(selected)
80        pathsText = pathsText & IIf(Len(pathsText) > 0, LIST_SEP, "") & CStr(selected(i))
      Next i

90    Application.ScreenUpdating = False
100   outPath = AnalyzeFileSet(pathsText)
110   Application.ScreenUpdating = True
120   If Len(outPath) > 0 Then
130       MsgBox MSG_DONE & vbCrLf & outPath, vbInformation
      Else
140       MsgBox MSG_FAILED, vbExclamation
      End If

      LogSuccess MODULE_NAME, "RunCrossFileAnalysis", "Analyzed " & (UBound(selected) - LBound(selected) + 1) & " files"
150   Exit Sub
ErrHandler:
      LogError MODULE_NAME, "RunCrossFileAnalysis", Err.Number, Err.Description, Erl, "outPath=" & outPath
      Application.ScreenUpdating = True
      MsgBox MSG_FAILED, vbExclamation
End Sub

'----------------------------------------------------------------------
' AnalyzeFileSet: ファイルを読み、対応表を Word (使えなければテキスト) で書き出して、そのパスを返す
'   pathsText  パスを | で区切った文字列。報告は先頭のファイルと同じフォルダに書く
'   失敗したら空文字を返す
'----------------------------------------------------------------------
Public Function AnalyzeFileSet(ByVal pathsText As String) As String
      On Error GoTo ErrHandler
      Dim paths() As String
      Dim report As String
      Dim outBase As String
      Dim outPath As String
      Dim fso As Object
      Dim stm As Object

10    paths = Split(pathsText, LIST_SEP)
20    report = BuildCrossReport(paths)

30    Set fso = CreateObject("Scripting.FileSystemObject")
40    outBase = fso.GetParentFolderName(paths(0)) & PATH_SEP & REPORT_NAME & Format(Now, REPORT_TIME_FORMAT)
50    If WordReportWriter.WriteWordReport(outBase & REPORT_EXT_WORD, report) Then
60        outPath = outBase & REPORT_EXT_WORD
      Else
          ' Word を使えないときは BOM 付き UTF-8 のテキストで書く
70        outPath = outBase & REPORT_EXT_TEXT
80        Set stm = CreateObject("ADODB.Stream")
90        stm.Type = AD_TYPE_TEXT
100       stm.Charset = REPORT_CHARSET
110       stm.Open
120       stm.WriteText report
130       stm.SaveToFile outPath, AD_SAVE_CREATE_OVERWRITE
140       stm.Close
150       Set stm = Nothing
      End If
160   AnalyzeFileSet = outPath

      LogSuccess MODULE_NAME, "AnalyzeFileSet", "Wrote cross-file report for " & (UBound(paths) + 1) & " files to " & outPath
170   Exit Function
ErrHandler:
      LogError MODULE_NAME, "AnalyzeFileSet", Err.Number, Err.Description, Erl, "pathsText=" & pathsText & "; outPath=" & outPath
      AnalyzeFileSet = ""
End Function

'----------------------------------------------------------------------
' BuildCrossReport: ファイルを読んで列を登録し、手がかりと結合キーの候補をまとめた報告文を返す
'----------------------------------------------------------------------
Private Function BuildCrossReport(ByRef paths() As String) As String
      On Error GoTo ErrHandler
      Dim fso As Object
      Dim f As Long
      Dim data() As Byte
      Dim byteCount As Long
      Dim charset As String
      Dim encLabel As String
      Dim bomBytes As Long
      Dim encHint As String
      Dim text As String
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
      Dim r As Long
      Dim c As Long
      Dim s As String
      Dim nv As String
      Dim kinds As Object
      Dim kindWidth As Object
      Dim useKinds As Boolean
      Dim kind As String
      Dim regIndex As Object
      Dim regKey As String
      Dim ci As Long
      Dim p As Long
      Dim w As Long
      Dim sumL As Long
      Dim dgt As Long
      Dim y As Long
      Dim m As Long
      Dim d As Long
      Dim k As Variant
      Dim fileText As String
      Dim kindText As String
      Dim hintText As String
      Dim i As Long
      Dim j As Long
      Dim small As Long
      Dim large As Long
      Dim common As Long
      Dim checked As Long
      Dim matchN As Long
      Dim mA() As Long
      Dim mB() As Long
      Dim mCommon() As Long
      Dim mShareA() As Double
      Dim mShareB() As Double
      Dim mScore() As Double
      Dim used() As Boolean
      Dim best As Long
      Dim shown As Long
      Dim matchText As String
      Dim fileNames() As String
      Dim nameA As String
      Dim nameB As String
      Dim rep As String
      Dim order() As Long
      Dim orderN As Long

10    Set fso = CreateObject("Scripting.FileSystemObject")
20    Set regIndex = CreateObject("Scripting.Dictionary")
30    m_colCount = 0
40    ReDim m_colFile(1 To MAX_COLS)
50    ReDim m_colKind(1 To MAX_COLS)
60    ReDim m_colNo(1 To MAX_COLS)
70    ReDim m_colVals(1 To MAX_COLS)
80    ReDim m_colFilled(1 To MAX_COLS)
90    ReDim m_colDigits(1 To MAX_COLS)
100   ReDim m_colDigitN(1 To MAX_COLS)
110   ReDim m_colMaxNum(1 To MAX_COLS)
120   ReDim m_colSeq(1 To MAX_COLS)
130   ReDim m_colMinLen(1 To MAX_COLS)
140   ReDim m_colMaxLen(1 To MAX_COLS)
150   ReDim m_colLuhn(1 To MAX_COLS)
160   ReDim m_colWarekiYm(1 To MAX_COLS)
170   ReDim m_colWarekiYmd(1 To MAX_COLS)
180   ReDim m_colDate8(1 To MAX_COLS)
190   ReDim m_colHint(1 To MAX_COLS)
200   ReDim fileNames(0 To UBound(paths))

      ' ---- ファイルごとに読み、種別 x 列 ごとに値を集める ----
210   For f = 0 To UBound(paths)
220       fileNames(f) = fso.GetFileName(paths(f))
230       data = TextFileReader.ReadAllBytes(paths(f), byteCount)
240       If byteCount <= 0 Then
250           fileText = fileText & "| " & (f + 1) & " | " & fileNames(f) & " | - | - | 0 | " & RPT_READ_ERROR & " |" & vbCrLf
          Else
260           charset = TextFileReader.DetectEncoding(data, byteCount, encLabel, bomBytes, encHint)
270           delim = ""
280           If Len(charset) > 0 Then
290               text = TextFileReader.DecodeBytes(data, charset)
300               lines = TextFileReader.SplitLines(text, eolLabel)
310               delim = TextFileReader.DetectDelimiter(lines, delimLabel, consistency)
              End If
320           If Len(delim) = 0 Then
330               fileText = fileText & "| " & (f + 1) & " | " & fileNames(f) & " | " & encLabel & " | - | - | " & RPT_SKIP & " |" & vbCrLf
              Else
340               grid = TextFileReader.ParseDelimited(lines, delim, MAX_ROWS, rowWidth, rowCount, colCount, _
                                                       quotedCount, truncated, multiLine)
                  ' 先頭項目が短いコードで種類が少なければ、種別ごとの表に分ける
350               Set kinds = CreateObject("Scripting.Dictionary")
360               Set kindWidth = CreateObject("Scripting.Dictionary")
370               useKinds = True
380               For r = 1 To rowCount
390                   s = Trim$(CStr(grid(r, 1)))
400                   If Len(s) > KIND_MAX_LEN Then
410                       useKinds = False
420                       Exit For
                      End If
430                   kinds(s) = kinds(s) + 1
440                   If rowWidth(r) > kindWidth(s) Then kindWidth(s) = rowWidth(r)
450                   If kinds.Count > KIND_MAX_KINDS Then
460                       useKinds = False
470                       Exit For
                      End If
                  Next r
480               If kinds.Count < 2 Then useKinds = False
490               kindText = ""
500               If useKinds Then
510                   For Each k In kinds.Keys
520                       kindText = kindText & IIf(Len(kindText) > 0, RPT_KIND_SEP, "") & k & " (" & kinds(k) & RPT_ROWS_SEP & kindWidth(k) & RPT_COLS_UNIT & ")"
                      Next k
                  Else
530                   kindText = NO_KIND & " (" & colCount & RPT_COLS_UNIT & ")"
                  End If
540               fileText = fileText & "| " & (f + 1) & " | " & fileNames(f) & " | " & encLabel & " | " & delimLabel & " | " & _
                             rowCount & " | " & kindText & " |" & vbCrLf

550               For r = 1 To rowCount
560                   If useKinds Then kind = Trim$(CStr(grid(r, 1))) Else kind = NO_KIND
570                   For c = 1 To colCount
580                       If c > rowWidth(r) Then Exit For
590                       s = Trim$(CStr(grid(r, c)))
600                       If Len(s) > 0 Then
610                           regKey = f & KEY_SEP & kind & KEY_SEP & c
620                           If regIndex.Exists(regKey) Then
630                               ci = regIndex(regKey)
                              ElseIf m_colCount < MAX_COLS Then
640                               m_colCount = m_colCount + 1
650                               ci = m_colCount
660                               regIndex.Add regKey, ci
670                               m_colFile(ci) = f
680                               m_colKind(ci) = kind
690                               m_colNo(ci) = c
700                               Set m_colVals(ci) = CreateObject("Scripting.Dictionary")
710                               m_colMinLen(ci) = Len(s)
                              Else
720                               ci = 0
                              End If
730                           If ci > 0 Then
740                               m_colFilled(ci) = m_colFilled(ci) + 1
750                               If Len(s) < m_colMinLen(ci) Then m_colMinLen(ci) = Len(s)
760                               If Len(s) > m_colMaxLen(ci) Then m_colMaxLen(ci) = Len(s)
770                               nv = s
780                               If Not s Like "*[!0-9]*" Then
790                                   m_colDigitN(ci) = m_colDigitN(ci) + 1
                                      ' 検証番号 (6-8 桁): 末尾の前の桁から 2, 1, 2 ... を掛け、積が 2 桁なら各桁を足す
800                                   If Len(s) >= LEN_KOKUHO And Len(s) <= LEN_HOKEN Then
810                                       sumL = 0
820                                       w = LUHN_W_FIRST
830                                       For p = Len(s) - 1 To 1 Step -1
840                                           dgt = (AscW(Mid$(s, p, 1)) - CODE_DIGIT_0) * w
850                                           If dgt >= DECIMAL_BASE Then dgt = dgt - DIGIT_CARRY
860                                           sumL = sumL + dgt
870                                           w = LUHN_W_SUM - w
                                          Next p
880                                       If (DECIMAL_BASE - sumL Mod DECIMAL_BASE) Mod DECIMAL_BASE = AscW(Right$(s, 1)) - CODE_DIGIT_0 Then
890                                           m_colLuhn(ci) = m_colLuhn(ci) + 1
                                          End If
                                      End If
                                      ' 和暦 (GYYMM / GYYMMDD) と西暦の日付 (yyyymmdd)
900                                   If Len(s) = LEN_WAREKI_YM Or Len(s) = LEN_WAREKI_YMD Then
910                                       y = CLng(Left$(s, 1))
920                                       m = CLng(Mid$(s, 4, 2))
930                                       If y >= ERA_MIN And y <= ERA_MAX And m >= 1 And m <= MONTH_MAX Then
940                                           If Len(s) = LEN_WAREKI_YM Then
950                                               m_colWarekiYm(ci) = m_colWarekiYm(ci) + 1
                                              Else
960                                               d = CLng(Right$(s, 2))
970                                               If d >= 1 And d <= DAY_MAX Then m_colWarekiYmd(ci) = m_colWarekiYmd(ci) + 1
                                              End If
                                          End If
                                      ElseIf Len(s) = LEN_DATE8 Then
980                                       y = CLng(Left$(s, YEAR_DIGITS))
990                                       m = CLng(Mid$(s, 5, 2))
1000                                      d = CLng(Right$(s, 2))
1010                                      If y >= YEAR_MIN And y <= YEAR_MAX And m >= 1 And m <= MONTH_MAX And d >= 1 And d <= DAY_MAX Then
1020                                          m_colDate8(ci) = m_colDate8(ci) + 1
                                          End If
                                      End If
                                      ' 比べるときは先頭の 0 を除く (000123 と 123 を同じとみなす)
1030                                  Do While Len(nv) > 1 And Left$(nv, 1) = "0"
1040                                      nv = Mid$(nv, 2)
                                      Loop
1050                                  If Len(nv) <= MAX_NUM_LEN Then
1060                                      If CDbl(nv) > m_colMaxNum(ci) Then m_colMaxNum(ci) = CDbl(nv)
                                      End If
                                  End If
1070                              If m_colVals(ci).Count < DISTINCT_CAP Then m_colVals(ci)(nv) = True
                              End If
                          End If
                      Next c
                  Next r
              End If
          End If
      Next f

      ' ---- 列ごとの手がかり ----
1080  For ci = 1 To m_colCount
1090      hintText = ""
          ' 見出しの行などが混じっても数字の列と判定できるよう、割合で決める
1100      m_colDigits(ci) = (m_colDigitN(ci) >= m_colFilled(ci) * HINT_SHARE)
1110      m_colSeq(ci) = False
1120      If m_colDigits(ci) And m_colVals(ci).Count >= MIN_KEY_KINDS And m_colVals(ci).Count >= m_colFilled(ci) * HINT_SHARE Then
              ' 桁が多すぎて最大値を覚えていない列 (21 桁の番号など) は連番とみなさない
1130          If m_colMaxLen(ci) <= MAX_NUM_LEN And m_colMaxNum(ci) > 0 Then
1140              If m_colMaxNum(ci) <= m_colVals(ci).Count * SEQ_SPAN Then m_colSeq(ci) = True
              End If
          End If
1150      If m_colFilled(ci) >= HINT_MIN_VALUES And m_colDigits(ci) Then
1160          If m_colLuhn(ci) >= m_colFilled(ci) * HINT_SHARE Then
1170              If m_colMinLen(ci) = LEN_HOKEN And m_colMaxLen(ci) = LEN_HOKEN Then
1180                  hintText = HINT_HOKEN
                  ElseIf m_colMinLen(ci) = LEN_KOKUHO And m_colMaxLen(ci) = LEN_KOKUHO Then
1190                  hintText = HINT_KOKUHO
                  ElseIf m_colMinLen(ci) = LEN_JUKYU And m_colMaxLen(ci) = LEN_JUKYU Then
1200                  hintText = HINT_JUKYU
                  End If
              End If
1210          If m_colWarekiYm(ci) >= m_colFilled(ci) * HINT_SHARE Then hintText = hintText & IIf(Len(hintText) > 0, HINT_SEP, "") & HINT_WAREKI_YM
1220          If m_colWarekiYmd(ci) >= m_colFilled(ci) * HINT_SHARE Then hintText = hintText & IIf(Len(hintText) > 0, HINT_SEP, "") & HINT_WAREKI_YMD
1230          If m_colDate8(ci) >= m_colFilled(ci) * HINT_SHARE Then hintText = hintText & IIf(Len(hintText) > 0, HINT_SEP, "") & HINT_DATE8
          End If
1240      If m_colFilled(ci) > 1 And m_colVals(ci).Count = 1 Then hintText = hintText & IIf(Len(hintText) > 0, HINT_SEP, "") & HINT_FIXED
1250      If m_colSeq(ci) Then hintText = hintText & IIf(Len(hintText) > 0, HINT_SEP, "") & HINT_SEQ
1260      If m_colFilled(ci) >= HINT_MIN_VALUES And m_colVals(ci).Count = m_colFilled(ci) Then
1270          If m_colMinLen(ci) = m_colMaxLen(ci) And m_colMinLen(ci) >= ID_MIN_LEN Then
1280              hintText = hintText & IIf(Len(hintText) > 0, HINT_SEP, "") & HINT_ID
              Else
1290              hintText = hintText & IIf(Len(hintText) > 0, HINT_SEP, "") & HINT_KEY
              End If
          End If
1300      m_colHint(ci) = hintText
      Next ci

      ' ---- 結合キーの候補: ファイルの違う列どうしで、共通の値の数を数える ----
1310  ReDim mA(1 To MAX_COLS)
1320  ReDim mB(1 To MAX_COLS)
1330  ReDim mCommon(1 To MAX_COLS)
1340  ReDim mShareA(1 To MAX_COLS)
1350  ReDim mShareB(1 To MAX_COLS)
1360  ReDim mScore(1 To MAX_COLS)
1370  For i = 1 To m_colCount - 1
1380      If m_colVals(i).Count >= MIN_KEY_KINDS Then
1390          For j = i + 1 To m_colCount
                  ' 違うファイルの列どうしを比べる (数字の列か文字の列かでは絞らない。
                  ' 数字に英字が少し混じる番号は、ファイルによって判定が分かれるため)
1400              If m_colFile(i) <> m_colFile(j) And m_colVals(j).Count >= MIN_KEY_KINDS _
                     And matchN < MAX_COLS Then
                      ' 種類の少ないほうの値を、多いほうの辞書で引く
1410                  If m_colVals(i).Count <= m_colVals(j).Count Then
1420                      small = i
1430                      large = j
                      Else
1440                      small = j
1450                      large = i
                      End If
1460                  common = 0
1470                  checked = 0
1480                  For Each k In m_colVals(small).Keys
1490                      checked = checked + 1
1500                      If m_colVals(large).Exists(k) Then common = common + 1
1510                      If checked >= CHECK_CAP Then Exit For
                      Next k
                      ' 調べた分から全体の共通の数を見積もる
1520                  If checked < m_colVals(small).Count Then common = CLng(common / checked * m_colVals(small).Count)
1530                  If common >= MIN_COMMON Then
1540                      If common >= m_colVals(i).Count * MATCH_SHARE Or common >= m_colVals(j).Count * MATCH_SHARE Then
1550                          matchN = matchN + 1
1560                          mA(matchN) = i
1570                          mB(matchN) = j
1580                          mCommon(matchN) = common
1590                          mShareA(matchN) = common / m_colVals(i).Count
1600                          mShareB(matchN) = common / m_colVals(j).Count
                              ' 並び順: 割合の大きいほう (片方が他方の一部でも強い対応) x 共通の数の平方根。
                              ' 両方が全件一意なら上げ、連番が絡むなら下げる
1610                          mScore(matchN) = IIf(mShareA(matchN) > mShareB(matchN), mShareA(matchN), mShareB(matchN)) * Sqr(common)
1620                          If m_colVals(i).Count = m_colFilled(i) And m_colVals(j).Count = m_colFilled(j) Then
1630                              mScore(matchN) = mScore(matchN) * UNIQUE_BONUS
                              End If
1640                          If InStr(m_colHint(i), HINT_ID) > 0 And InStr(m_colHint(j), HINT_ID) > 0 Then
1650                              mScore(matchN) = mScore(matchN) * ID_BONUS
                              End If
1660                          If m_colSeq(i) Or m_colSeq(j) Then mScore(matchN) = mScore(matchN) * SEQ_PENALTY
                          End If
                      End If
                  End If
              Next j
          End If
      Next i

      ' ---- 強い順に並べて報告する (mScore の高い順) ----
1670  If matchN > 0 Then ReDim used(1 To matchN)
1680  ReDim order(1 To MAX_MATCH_ROWS)
1690  For shown = 1 To matchN
1700      If shown > MAX_MATCH_ROWS Then Exit For
1710      best = 0
1720      For i = 1 To matchN
1730          If Not used(i) Then
1740              If best = 0 Then
1750                  best = i
                  ElseIf mScore(i) > mScore(best) Then
1760                  best = i
                  End If
              End If
          Next i
1770      used(best) = True
1780      nameA = RPT_ITEM_PREFIX & (m_colFile(mA(best)) + 1) & RPT_ITEM_SEP & m_colKind(mA(best)) & RPT_ITEM_SEP & RPT_ITEM_NO & m_colNo(mA(best))
1790      nameB = RPT_ITEM_PREFIX & (m_colFile(mB(best)) + 1) & RPT_ITEM_SEP & m_colKind(mB(best)) & RPT_ITEM_SEP & RPT_ITEM_NO & m_colNo(mB(best))
1800      matchText = matchText & "| " & nameA & " | " & nameB & " | " & mCommon(best) & " | " & _
                      Format(mShareA(best), "0%") & " | " & Format(mShareB(best), "0%") & " | " & _
                      IIf(Len(m_colHint(mA(best))) > 0, m_colHint(mA(best)), "-") & " | " & _
                      IIf(m_colSeq(mA(best)) Or m_colSeq(mB(best)), RPT_SEQ_MARK, "") & " |" & vbCrLf
          ' 並べた順を覚えておき、ファイルの組ごとの要約に使う
1810      orderN = orderN + 1
1820      order(orderN) = best
      Next shown

      ' ---- 報告 ----
1830  rep = RPT_TITLE & vbCrLf & RPT_RULE & vbCrLf & _
            RPT_RUN_AT & Format(Now, RUN_TIME_FORMAT) & vbCrLf & _
            RPT_FILE_COUNT & (UBound(paths) + 1) & vbCrLf & vbCrLf & _
            RPT_LEGEND & vbCrLf & RPT_LEGEND_1 & vbCrLf & RPT_LEGEND_2 & vbCrLf & RPT_LEGEND_3 & vbCrLf & vbCrLf & _
            RPT_SEC_FILES & vbCrLf & RPT_FILES_HEAD & vbCrLf & RPT_FILES_RULE & vbCrLf & fileText & vbCrLf

1840  rep = rep & RPT_SEC_MATCH & vbCrLf
1850  If matchN = 0 Then
1860      rep = rep & RPT_NONE & vbCrLf
      Else
1870      rep = rep & RPT_MATCH_NOTE_HEAD & vbCrLf & RPT_MATCH_NOTE_RULE & vbCrLf & matchText
1880      If matchN > MAX_MATCH_ROWS Then rep = rep & RPT_MORE & vbCrLf
      End If
1890  rep = rep & RPT_NOTE_MONTH & vbCrLf & vbCrLf

      ' ファイルの組ごとの結び方 (連番が絡む一致を除き、強い順に GUIDE_PER_PAIR まで)
1900  rep = rep & RPT_SEC_GUIDE & vbCrLf
1910  For i = 0 To UBound(paths) - 1
1920      For j = i + 1 To UBound(paths)
1930          rep = rep & RPT_GUIDE_PAIR & (i + 1) & RPT_GUIDE_AND & (j + 1) & RPT_GUIDE_COLON & vbCrLf
1940          shown = 0
1950          For best = 1 To orderN
1960              If shown < GUIDE_PER_PAIR Then
1970                  If Not (m_colSeq(mA(order(best))) Or m_colSeq(mB(order(best)))) Then
1980                      If (m_colFile(mA(order(best))) = i And m_colFile(mB(order(best))) = j) Or _
                             (m_colFile(mA(order(best))) = j And m_colFile(mB(order(best))) = i) Then
1990                          shown = shown + 1
2000                          rep = rep & RPT_GUIDE_ITEM & _
                                    RPT_ITEM_PREFIX & (m_colFile(mA(order(best))) + 1) & RPT_ITEM_SEP & m_colKind(mA(order(best))) & _
                                    RPT_ITEM_SEP & RPT_ITEM_NO & m_colNo(mA(order(best))) & RPT_GUIDE_ARROW & _
                                    RPT_ITEM_PREFIX & (m_colFile(mB(order(best))) + 1) & RPT_ITEM_SEP & m_colKind(mB(order(best))) & _
                                    RPT_ITEM_SEP & RPT_ITEM_NO & m_colNo(mB(order(best))) & _
                                    RPT_GUIDE_SHARE & Format(mShareA(order(best)), "0%") & RPT_GUIDE_SHARE_2 & _
                                    Format(mShareB(order(best)), "0%") & ")" & _
                                    IIf(Len(m_colHint(mA(order(best)))) > 0, " " & m_colHint(mA(order(best))), "") & vbCrLf
                          End If
                      End If
                  End If
              Next best
2010          If shown = 0 Then rep = rep & RPT_GUIDE_NONE & vbCrLf
          Next j
      Next i
2020  rep = rep & vbCrLf

      ' 手がかりのある列 (と、結合キーの候補に出た列) の一覧
2030  rep = rep & RPT_SEC_HINTS & vbCrLf & RPT_HINTS_HEAD & vbCrLf & RPT_HINTS_RULE & vbCrLf
2040  For ci = 1 To m_colCount
2050      If Len(m_colHint(ci)) > 0 And m_colHint(ci) <> HINT_FIXED Then
2060          rep = rep & "| " & RPT_ITEM_PREFIX & (m_colFile(ci) + 1) & RPT_ITEM_SEP & m_colKind(ci) & RPT_ITEM_SEP & RPT_ITEM_NO & m_colNo(ci) & _
                    " | " & m_colFilled(ci) & " | " & m_colVals(ci).Count & " | " & _
                    IIf(m_colDigits(ci), RPT_DIGIT_COL, RPT_TEXT_COL & " ") & m_colMinLen(ci) & _
                    IIf(m_colMaxLen(ci) <> m_colMinLen(ci), "-" & m_colMaxLen(ci), "") & RPT_LEN_UNIT & " | " & m_colHint(ci) & " |" & vbCrLf
          End If
      Next ci
2070  BuildCrossReport = rep & vbCrLf

      LogSuccess MODULE_NAME, "BuildCrossReport", "Files=" & (UBound(paths) + 1) & "; columns=" & m_colCount & "; matches=" & matchN
2080  Exit Function
ErrHandler:
      LogError MODULE_NAME, "BuildCrossReport", Err.Number, Err.Description, Erl, _
               "f=" & f & "; r=" & r & "; c=" & c & "; i=" & i & "; j=" & j & "; colCount=" & m_colCount
      BuildCrossReport = rep & MSG_FAILED & vbCrLf
End Function
