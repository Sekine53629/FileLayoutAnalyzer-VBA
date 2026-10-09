Attribute VB_Name = "TextFileReader"
Option Explicit

'======================================================================
' TextFileReader モジュール
' テキストファイルの読み込みと、区切り文字形式の解析を行う。
'   - バイト列の読み込み (ADODB.Stream。日本語を含むパスにも対応)
'   - 文字コードの判定 (BOM、UTF-8 の妥当性、Shift_JIS、EUC-JP、JIS、UTF-16、バイナリ)
'     どれでもないときは 1 バイト = 1 文字で読み、判定の手がかりを報告用に返す
'   - 改行の判定と行への分割
'   - 区切り文字の判定 (カンマ、タブ、縦棒、セミコロン)
'   - 引用符を考慮した区切り文字形式の解析 (セル内改行にも対応)
'
' 依存: ErrorLogger.bas
'======================================================================

Private Const MODULE_NAME As String = "TextFileReader"

' ---- ADODB.Stream ----
Private Const AD_TYPE_BINARY As Long = 1
Private Const AD_TYPE_TEXT As Long = 2

' ---- 文字コード名 (ADODB.Stream の Charset) ----
Public Const CHARSET_UTF8 As String = "UTF-8"
Public Const CHARSET_UTF16LE As String = "Unicode"
Public Const CHARSET_UTF16BE As String = "UnicodeFFFE"
Public Const CHARSET_SJIS As String = "Shift_JIS"
Public Const CHARSET_EUCJP As String = "EUC-JP"
Public Const CHARSET_JIS As String = "iso-2022-jp"
Public Const CHARSET_BYTES As String = "iso-8859-1"     ' 文字コード不明: 1 バイト = 1 文字で読む (区切りや改行の位置を崩さない)
Public Const CHARSET_BINARY As String = ""

' ---- 判定の閾値 ----
Private Const ENCODING_SAMPLE_BYTES As Long = 4194304   ' 文字コード判定に使う先頭のバイト数 (4MB)
Private Const BINARY_CTRL_RATIO As Double = 0.1         ' 制御文字がこの割合を超えたらバイナリとみなす
Private Const UTF16_NUL_RATIO As Double = 0.3           ' 偶数/奇数位置の NUL がこの割合を超えたら UTF-16 とみなす
Private Const DELIM_SAMPLE_LINES As Long = 2000         ' 区切り文字の判定に使う行数
Private Const DELIM_MIN_SHARE As Double = 0.9           ' この割合以上の行に現れる文字を区切り文字の候補とする
Private Const DELIM_SCORE_WEIGHT As Double = 1000       ' 区切り数の一致率を区切り数より優先するための重み

' ---- バイト値 ----
Private Const BYTE_NUL As Long = 0
Private Const BYTE_TAB As Long = 9
Private Const BYTE_LF As Long = 10
Private Const BYTE_CR As Long = 13
Private Const BYTE_SPACE As Long = 32
Private Const BYTE_DEL As Long = 127
Private Const BYTE_ESC As Long = 27
Private Const BYTE_FF As Long = 12
Private Const BYTE_SO As Long = 14                      ' 汎用機の漢字の開始の印 (シフトアウト)
Private Const BYTE_SI As Long = 15                      ' 汎用機の漢字の終了の印 (シフトイン)
Private Const BYTE_HIGH As Long = &H80
Private Const BYTE_C1_LAST As Long = &H9F
Private Const BYTE_EUC_SS2 As Long = &H8E               ' EUC-JP の半角カナの前置き
Private Const BYTE_EUC_SS3 As Long = &H8F               ' EUC-JP の補助漢字の前置き
Private Const BYTE_EUC_FIRST As Long = &HA1
Private Const BYTE_EUC_LAST As Long = &HFE
Private Const BYTE_KANA_LAST As Long = &HDF
Private Const BYTE_E0 As Long = &HE0
Private Const BYTE_JIS_DOLLAR As Long = 36              ' ESC $ B / ESC $ @ の "$"
Private Const BOM_UTF8_1 As Long = &HEF
Private Const BOM_UTF8_2 As Long = &HBB
Private Const BOM_UTF8_3 As Long = &HBF
Private Const BOM_UTF16_FF As Long = &HFF
Private Const BOM_UTF16_FE As Long = &HFE

' ---- 文言 ----
Public Const ENC_LABEL_UTF8_BOM As String = "UTF-8 (BOM あり)"
Public Const ENC_LABEL_UTF8 As String = "UTF-8 (BOM なし)"
Public Const ENC_LABEL_UTF16LE As String = "UTF-16 LE"
Public Const ENC_LABEL_UTF16BE As String = "UTF-16 BE"
Public Const ENC_LABEL_SJIS As String = "Shift_JIS (CP932)"
Public Const ENC_LABEL_SJIS_KANA As String = "Shift_JIS (CP932。全角なし、半角カナのみ)"
Public Const ENC_LABEL_ASCII As String = "ASCII のみ (Shift_JIS として読む)"
Public Const ENC_LABEL_SJIS_INVALID As String = "不明 (Shift_JIS として読む。不正なバイト列あり)"
Public Const ENC_LABEL_EUCJP As String = "EUC-JP"
Public Const ENC_LABEL_JIS As String = "JIS (ISO-2022-JP)"
Public Const ENC_LABEL_UNKNOWN As String = "不明 (UTF-8 / Shift_JIS / EUC-JP / JIS のどれでもない。1 バイト = 1 文字で読む)"
' ---- 文字コードの手がかり ----
Private Const HINT_SEC As String = "[文字コードの手がかり]"
Private Const HINT_HIGH As String = "  0x80 以上のバイト : "
Private Const HINT_HIGH_UNIT As String = " バイト (全体の "
Private Const HINT_RANGES As String = "  範囲別 : 0x80-0x9F "
Private Const HINT_RANGE_A0 As String = " / 0xA0-0xDF "
Private Const HINT_RANGE_E0 As String = " / 0xE0-0xFF "
Private Const HINT_VALID As String = "  妥当性 : UTF-8 "
Private Const HINT_VALID_SJIS As String = " / Shift_JIS "
Private Const HINT_VALID_EUC As String = " / EUC-JP "
Private Const HINT_OK As String = "妥当"
Private Const HINT_NG As String = "不正"
Private Const HINT_CTRL As String = "  制御コード : ESC (JIS の切り替え) "
Private Const HINT_CTRL_SO As String = " / SO 0x0E (汎用機の漢字の開始) "
Private Const HINT_CTRL_SI As String = " / SI 0x0F (汎用機の漢字の終了) "
Private Const HINT_RUNS As String = "  0x80 以上のバイトの連続 : "
Private Const HINT_RUNS_EVEN As String = "どれも偶数バイト (2 バイトで 1 文字の漢字コードらしい)"
Private Const HINT_RUNS_ODD As String = "奇数バイトの連続あり (半角カナなど 1 バイトの文字を含む)"
Private Const HINT_GUESS As String = "  推定 : "
Private Const GUESS_EUC As String = "EUC-JP (全角が 0xA1-0xFE の 2 バイトの組だけでできていて、Shift_JIS の漢字に特有の 0x81-0x9F が無い)"
Private Const GUESS_JIS As String = "JIS (ISO-2022-JP。ESC で漢字に切り替えている)"
Private Const GUESS_HOST As String = "汎用機の漢字コード (JEF・KEIS・IBM 漢字など) を SO (0x0E) と SI (0x0F) で囲む形式の可能性が高い。Windows の標準機能では文字に戻せないので、変換表かホスト側の仕様書が要る"
Private Const GUESS_DBCS As String = "2 バイトで 1 文字の漢字コードで、Windows で読める文字コードではない可能性 (汎用機の漢字コードを制御コードなしで埋め込んだもの、システム独自の漢字コードなど)。出力元の仕様書で漢字コードを確かめること"
Private Const GUESS_NONE As String = "推定できない。出力元の仕様書で文字コードを確かめること"
Private Const HINT_NOTE_BYTES As String = "  注: Shift_JIS 以外の全角の部分は、固定長の項目一覧の書式パターンで 1 バイトを Ｎ 1 つで表す (Ｎ の数 = バイト数。全角 1 文字 = Ｎ 2 つ)"
Public Const ENC_LABEL_BINARY As String = "バイナリ (テキストではない)"
Public Const ENC_LABEL_NO_BOM As String = " (BOM なし)"
Private Const DELIM_LABEL_COMMA As String = "カンマ"
Private Const DELIM_LABEL_TAB As String = "タブ"
Private Const DELIM_LABEL_PIPE As String = "縦棒 (|)"
Private Const DELIM_LABEL_SEMI As String = "セミコロン"
Private Const EOL_LABEL_CRLF As String = "CRLF"
Private Const EOL_LABEL_LF As String = "LF"
Private Const EOL_LABEL_CR As String = "CR"
Private Const EOL_LABEL_MIXED As String = "混在"
Private Const EOL_LABEL_NONE As String = "改行なし"

Private Const QUOTE_CHAR As String = """"


'----------------------------------------------------------------------
' ReadAllBytes: ファイルを丸ごとバイト配列で読む
'   空のファイルは要素数 0 の判定ができるよう byteCount で返す
'----------------------------------------------------------------------
Public Function ReadAllBytes(ByVal filePath As String, ByRef byteCount As Long) As Byte()
      On Error GoTo ErrHandler
      Dim stm As Object
      Dim data() As Byte

10    Set stm = CreateObject("ADODB.Stream")
20    stm.Type = AD_TYPE_BINARY
30    stm.Open
40    stm.LoadFromFile filePath
50    byteCount = stm.Size
60    If byteCount > 0 Then
70        data = stm.Read
      End If
80    stm.Close
90    Set stm = Nothing
100   ReadAllBytes = data

      LogSuccess MODULE_NAME, "ReadAllBytes", "Read " & byteCount & " bytes from " & filePath
110   Exit Function
ErrHandler:
      LogError MODULE_NAME, "ReadAllBytes", Err.Number, Err.Description, Erl, _
               "filePath=" & filePath & "; byteCount=" & byteCount
      byteCount = -1
End Function

'----------------------------------------------------------------------
' DetectEncoding: バイト列の文字コードを判定する
'   戻り値     ADODB.Stream の Charset 名 (バイナリなら空文字)
'   encLabel   報告用の名前
'   bomBytes   BOM のバイト数 (固定長の位置計算で除く)
'   encHint    EUC-JP / JIS / 不明 のとき、判定の手がかり (報告の 1 欄)。それ以外は空
'----------------------------------------------------------------------
Public Function DetectEncoding(ByRef data() As Byte, ByVal byteCount As Long, _
                                ByRef encLabel As String, ByRef bomBytes As Long, _
                                ByRef encHint As String) As String
      On Error GoTo ErrHandler
      Dim i As Long
      Dim limit As Long
      Dim b As Long
      Dim ctrlCount As Long
      Dim nulEven As Long
      Dim nulOdd As Long
      Dim highCount As Long
      Dim utf8Valid As Boolean
      Dim utf8Multi As Long
      Dim need As Long
      Dim sjisValid As Boolean
      Dim inLead As Boolean
      Dim eucValid As Boolean
      Dim eucNeed As Long
      Dim eucPairs As Long
      Dim c1Count As Long
      Dim midCount As Long
      Dim topCount As Long
      Dim escJis As Long
      Dim soCount As Long
      Dim siCount As Long
      Dim curRun As Long
      Dim oddRun As Boolean
      Dim guess As String
      Dim highAll As Long

10    bomBytes = 0
20    encHint = ""
      ' ---- BOM ----
30    If byteCount >= 3 Then
40        If data(0) = BOM_UTF8_1 And data(1) = BOM_UTF8_2 And data(2) = BOM_UTF8_3 Then
50            encLabel = ENC_LABEL_UTF8_BOM
60            bomBytes = 3
70            DetectEncoding = CHARSET_UTF8
              LogSuccess MODULE_NAME, "DetectEncoding", "Detected UTF-8 BOM"
80            Exit Function
          End If
      End If
90    If byteCount >= 2 Then
100       If data(0) = BOM_UTF16_FF And data(1) = BOM_UTF16_FE Then
110           encLabel = ENC_LABEL_UTF16LE
120           bomBytes = 2
130           DetectEncoding = CHARSET_UTF16LE
              LogSuccess MODULE_NAME, "DetectEncoding", "Detected UTF-16LE BOM"
140           Exit Function
          End If
150       If data(0) = BOM_UTF16_FE And data(1) = BOM_UTF16_FF Then
160           encLabel = ENC_LABEL_UTF16BE
170           bomBytes = 2
180           DetectEncoding = CHARSET_UTF16BE
              LogSuccess MODULE_NAME, "DetectEncoding", "Detected UTF-16BE BOM"
190           Exit Function
          End If
      End If

      ' ---- 先頭部分を走査して、制御文字・NUL・UTF-8 と Shift_JIS の妥当性を数える ----
200   limit = byteCount - 1
210   If limit > ENCODING_SAMPLE_BYTES - 1 Then limit = ENCODING_SAMPLE_BYTES - 1
220   utf8Valid = True
230   sjisValid = True
240   eucValid = True
250   For i = 0 To limit
260       b = data(i)
270       If b = BYTE_NUL Then
280           If i Mod 2 = 0 Then nulEven = nulEven + 1 Else nulOdd = nulOdd + 1
          ElseIf b = BYTE_SO Then
290           soCount = soCount + 1
          ElseIf b = BYTE_SI Then
300           siCount = siCount + 1
          ElseIf b = BYTE_ESC Then
310           If i < limit Then
320               If data(i + 1) = BYTE_JIS_DOLLAR Then escJis = escJis + 1
              End If
          ElseIf b < BYTE_SPACE And b <> BYTE_TAB And b <> BYTE_LF And b <> BYTE_CR _
                 And b <> BYTE_FF Then
330           ctrlCount = ctrlCount + 1
          ElseIf b = BYTE_DEL Then
340           ctrlCount = ctrlCount + 1
          End If

          ' UTF-8 の妥当性
350       If need > 0 Then
360           If b >= &H80 And b <= &HBF Then
370               need = need - 1
380               If need = 0 Then utf8Multi = utf8Multi + 1
              Else
390               utf8Valid = False
400               need = 0
              End If
          ElseIf b >= &H80 Then
410           highCount = highCount + 1
420           If b >= &HC2 And b <= &HDF Then
430               need = 1
              ElseIf b >= &HE0 And b <= &HEF Then
440               need = 2
              ElseIf b >= &HF0 And b <= &HF4 Then
450               need = 3
              Else
460               utf8Valid = False
              End If
          End If

          ' EUC-JP の妥当性 (0xA1-0xFE の 2 バイト、0x8E + 半角カナ、0x8F + 2 バイト)
470       If eucNeed > 0 Then
480           If b >= BYTE_EUC_FIRST And b <= BYTE_EUC_LAST Then
490               eucNeed = eucNeed - 1
500               If eucNeed = 0 Then eucPairs = eucPairs + 1
              Else
510               eucValid = False
520               eucNeed = 0
              End If
          ElseIf b >= BYTE_EUC_FIRST And b <= BYTE_EUC_LAST Then
530           eucNeed = 1
          ElseIf b = BYTE_EUC_SS2 Then
540           eucNeed = 1
          ElseIf b = BYTE_EUC_SS3 Then
550           eucNeed = 2
          ElseIf b >= BYTE_HIGH Then
560           eucValid = False
          End If

          ' 0x80 以上のバイトの範囲と、連続する長さ (2 バイト文字なら偶数になる)
570       If b >= BYTE_HIGH Then
580           highAll = highAll + 1
590           curRun = curRun + 1
600           If b <= BYTE_C1_LAST Then
610               If b <> BYTE_EUC_SS2 And b <> BYTE_EUC_SS3 Then c1Count = c1Count + 1
              ElseIf b <= BYTE_KANA_LAST Then
620               midCount = midCount + 1
              Else
630               topCount = topCount + 1
              End If
          Else
640           If curRun Mod 2 = 1 Then oddRun = True
650           curRun = 0
          End If

          ' Shift_JIS の妥当性 (2 バイト目の範囲だけ確認する)
660       If inLead Then
670           If Not ((b >= &H40 And b <= &H7E) Or (b >= &H80 And b <= &HFC)) Then sjisValid = False
680           inLead = False
          ElseIf (b >= &H81 And b <= &H9F) Or (b >= &HE0 And b <= &HFC) Then
690           inLead = True
          End If
      Next i
700   If curRun Mod 2 = 1 Then oddRun = True

      ' ---- 判定 ----
      ' 発注データなどで多い「Shift_JIS の半角カナ」を、UTF-8 / EUC-JP より先に判定する
      ' (半角カナ 0xA1-0xDF の並びは、偶然 UTF-8 や EUC-JP としても妥当になることがあるため)
710   If (nulEven + nulOdd) > 0 Then
720       If nulOdd > (limit + 1) / 2 * UTF16_NUL_RATIO And nulEven = 0 Then
730           encLabel = ENC_LABEL_UTF16LE & ENC_LABEL_NO_BOM
740           DetectEncoding = CHARSET_UTF16LE
          ElseIf nulEven > (limit + 1) / 2 * UTF16_NUL_RATIO And nulOdd = 0 Then
750           encLabel = ENC_LABEL_UTF16BE & ENC_LABEL_NO_BOM
760           DetectEncoding = CHARSET_UTF16BE
          Else
770           encLabel = ENC_LABEL_BINARY
780           DetectEncoding = CHARSET_BINARY
          End If
      ElseIf ctrlCount > (limit + 1) * BINARY_CTRL_RATIO Then
790       encLabel = ENC_LABEL_BINARY
800       DetectEncoding = CHARSET_BINARY
      ElseIf highCount = 0 And escJis > 0 Then
810       encLabel = ENC_LABEL_JIS
820       DetectEncoding = CHARSET_JIS
830       guess = GUESS_JIS
      ElseIf highCount = 0 And soCount = 0 Then
840       encLabel = ENC_LABEL_ASCII
850       DetectEncoding = CHARSET_SJIS
      ElseIf soCount = 0 And sjisValid And midCount > 0 And highAll = midCount Then
          ' 0x80 以上のバイトがすべて 0xA0-0xDF (Shift_JIS の半角カナの範囲)
860       encLabel = ENC_LABEL_SJIS_KANA
870       DetectEncoding = CHARSET_SJIS
      ElseIf utf8Valid And utf8Multi > 0 And soCount = 0 Then
880       encLabel = ENC_LABEL_UTF8
890       DetectEncoding = CHARSET_UTF8
      ElseIf soCount > 0 And siCount > 0 Then
          ' SO / SI は Shift_JIS や EUC-JP のテキストには普通現れない
900       encLabel = ENC_LABEL_UNKNOWN
910       DetectEncoding = CHARSET_BYTES
920       guess = GUESS_HOST
      ElseIf eucValid And eucPairs > 0 And c1Count = 0 And topCount > 0 And Not oddRun Then
          ' Shift_JIS の漢字の 1 バイト目に多い 0x81-0x9F が無く、0xE0 以上があり、
          ' 0x80 以上のバイトが必ず 2 バイトずつ並ぶ (半角カナを含む Shift_JIS と区別する)
930       encLabel = ENC_LABEL_EUCJP
940       DetectEncoding = CHARSET_EUCJP
950       guess = GUESS_EUC
      ElseIf sjisValid Then
960       encLabel = ENC_LABEL_SJIS
970       DetectEncoding = CHARSET_SJIS
      Else
980       encLabel = ENC_LABEL_UNKNOWN
990       DetectEncoding = CHARSET_BYTES
1000      guess = IIf(Not oddRun And highCount > 0, GUESS_DBCS, GUESS_NONE)
      End If

      ' ---- 判定の手がかり (Shift_JIS / UTF-8 / ASCII と素直に判定できたときは出さない) ----
1010  If Len(guess) > 0 Then
1020      encHint = HINT_SEC & vbCrLf & _
                    HINT_HIGH & highAll & HINT_HIGH_UNIT & Format(highAll / (limit + 1), "0.0%") & ")" & vbCrLf & _
                    HINT_RANGES & (highAll - midCount - topCount) & HINT_RANGE_A0 & midCount & HINT_RANGE_E0 & topCount & vbCrLf & _
                    HINT_VALID & IIf(utf8Valid, HINT_OK, HINT_NG) & HINT_VALID_SJIS & IIf(sjisValid, HINT_OK, HINT_NG) & _
                    HINT_VALID_EUC & IIf(eucValid, HINT_OK, HINT_NG) & vbCrLf & _
                    HINT_CTRL & escJis & HINT_CTRL_SO & soCount & HINT_CTRL_SI & siCount & vbCrLf & _
                    HINT_RUNS & IIf(oddRun, HINT_RUNS_ODD, HINT_RUNS_EVEN) & vbCrLf & _
                    HINT_GUESS & guess & vbCrLf
1030      If DetectEncoding = CHARSET_BYTES Or DetectEncoding = CHARSET_EUCJP Then encHint = encHint & HINT_NOTE_BYTES & vbCrLf
1040      encHint = encHint & vbCrLf
      End If

      LogSuccess MODULE_NAME, "DetectEncoding", "Detected encoding " & encLabel & " from " & (limit + 1) & " bytes"
1050  Exit Function
ErrHandler:
      LogError MODULE_NAME, "DetectEncoding", Err.Number, Err.Description, Erl, _
               "byteCount=" & byteCount & "; i=" & i
      encLabel = ENC_LABEL_SJIS_INVALID
      encHint = ""
      DetectEncoding = CHARSET_SJIS
End Function

'----------------------------------------------------------------------
' DecodeBytes: バイト列を指定の文字コードで文字列にする
'   ADODB.Stream は UTF-8 / UTF-16 の BOM を取り除いて読む
'----------------------------------------------------------------------
Public Function DecodeBytes(ByRef data() As Byte, ByVal charset As String) As String
      On Error GoTo ErrHandler
      Dim stm As Object

10    Set stm = CreateObject("ADODB.Stream")
20    stm.Type = AD_TYPE_BINARY
30    stm.Open
40    stm.Write data
50    stm.Position = 0
60    stm.Type = AD_TYPE_TEXT
70    stm.Charset = charset
80    DecodeBytes = stm.ReadText
90    stm.Close
100   Set stm = Nothing

      LogSuccess MODULE_NAME, "DecodeBytes", "Decoded " & Len(DecodeBytes) & " chars as " & charset
110   Exit Function
ErrHandler:
      LogError MODULE_NAME, "DecodeBytes", Err.Number, Err.Description, Erl, "charset=" & charset
      DecodeBytes = ""
End Function

'----------------------------------------------------------------------
' SplitLines: 文字列を行に分ける (末尾の空行は除く)
'   eolLabel には改行の種類を返す
'----------------------------------------------------------------------
Public Function SplitLines(ByVal text As String, ByRef eolLabel As String) As String()
      On Error GoTo ErrHandler
      Dim crlfCount As Long
      Dim crCount As Long
      Dim lfCount As Long
      Dim lines() As String
      Dim normalized As String
      Dim lastIdx As Long

10    crlfCount = (Len(text) - Len(Replace(text, vbCrLf, ""))) \ Len(vbCrLf)
20    crCount = Len(text) - Len(Replace(text, vbCr, "")) - crlfCount
30    lfCount = Len(text) - Len(Replace(text, vbLf, "")) - crlfCount

40    If crlfCount = 0 And crCount = 0 And lfCount = 0 Then
50        eolLabel = EOL_LABEL_NONE
      ElseIf crCount = 0 And lfCount = 0 Then
60        eolLabel = EOL_LABEL_CRLF
      ElseIf crlfCount = 0 And crCount = 0 Then
70        eolLabel = EOL_LABEL_LF
      ElseIf crlfCount = 0 And lfCount = 0 Then
80        eolLabel = EOL_LABEL_CR
      Else
90        eolLabel = EOL_LABEL_MIXED & " (CRLF " & crlfCount & " / LF " & lfCount & " / CR " & crCount & ")"
      End If

100   normalized = Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf)
110   lines = Split(normalized, vbLf)
120   lastIdx = UBound(lines)
130   Do While lastIdx > 0
140       If Len(lines(lastIdx)) > 0 Then Exit Do
150       lastIdx = lastIdx - 1
      Loop
160   If lastIdx >= 0 And lastIdx < UBound(lines) Then ReDim Preserve lines(0 To lastIdx)
170   SplitLines = lines

      LogSuccess MODULE_NAME, "SplitLines", "Split into " & (lastIdx + 1) & " lines; eol=" & eolLabel
180   Exit Function
ErrHandler:
      LogError MODULE_NAME, "SplitLines", Err.Number, Err.Description, Erl, _
               "textLen=" & Len(text) & "; lastIdx=" & lastIdx
      eolLabel = EOL_LABEL_NONE
      SplitLines = Split(text, vbLf)
End Function

'----------------------------------------------------------------------
' DetectDelimiter: 区切り文字を判定する (見つからなければ空文字)
'   ほぼすべての行に現れ、行ごとの個数がもっとも揃っている文字を選ぶ
'   delimLabel には報告用の名前、consistency には個数が最頻値に一致する行の割合を返す
'----------------------------------------------------------------------
Public Function DetectDelimiter(ByRef lines() As String, ByRef delimLabel As String, _
                                  ByRef consistency As Double) As String
      On Error GoTo ErrHandler
      Dim cands As Variant
      Dim labels As Variant
      Dim k As Long
      Dim i As Long
      Dim p As Long
      Dim ln As String
      Dim cnt As Long
      Dim inQuote As Boolean
      Dim sampleCount As Long
      Dim withCount As Long
      Dim counts As Object
      Dim key As Variant
      Dim modeCnt As Long
      Dim modeFreq As Long
      Dim score As Double
      Dim bestScore As Double
      Dim bestShare As Double

10    cands = Array(",", vbTab, "|", ";")
20    labels = Array(DELIM_LABEL_COMMA, DELIM_LABEL_TAB, DELIM_LABEL_PIPE, DELIM_LABEL_SEMI)
30    bestScore = -1
40    DetectDelimiter = ""
50    delimLabel = ""
60    consistency = 0

70    For k = LBound(cands) To UBound(cands)
80        Set counts = CreateObject("Scripting.Dictionary")
90        sampleCount = 0
100       withCount = 0
110       i = LBound(lines)
120       Do While i <= UBound(lines)
130           If sampleCount >= DELIM_SAMPLE_LINES Then Exit Do
140           ln = lines(i)
              ' 引用符が閉じていなければ (値の中の改行)、次の行をつないで 1 件として数える
150           Do While ((Len(ln) - Len(Replace(ln, QUOTE_CHAR, ""))) Mod 2 = 1) And i < UBound(lines)
160               i = i + 1
170               ln = ln & vbLf & lines(i)
              Loop
180           i = i + 1
190           If Len(ln) > 0 Then
200               sampleCount = sampleCount + 1
210               If InStr(ln, QUOTE_CHAR) = 0 Then
220                   cnt = (Len(ln) - Len(Replace(ln, cands(k), ""))) \ Len(cands(k))
                  Else
                      ' 引用符の中の区切り文字は数えない
230                   cnt = 0
240                   inQuote = False
250                   For p = 1 To Len(ln)
260                       If Mid$(ln, p, 1) = QUOTE_CHAR Then
270                           inQuote = Not inQuote
                          ElseIf Not inQuote Then
280                           If Mid$(ln, p, 1) = cands(k) Then cnt = cnt + 1
                          End If
                      Next p
                  End If
290               If cnt > 0 Then withCount = withCount + 1
300               counts(cnt) = counts(cnt) + 1
              End If
          Loop

310       If sampleCount > 0 Then
320           If withCount / sampleCount >= DELIM_MIN_SHARE Then
330               modeCnt = 0
340               modeFreq = 0
350               For Each key In counts.Keys
360                   If counts(key) > modeFreq Then
370                       modeFreq = counts(key)
380                       modeCnt = key
                      End If
                  Next key
390               score = (modeFreq / sampleCount) * DELIM_SCORE_WEIGHT + modeCnt
400               If score > bestScore Then
410                   bestScore = score
420                   bestShare = modeFreq / sampleCount
430                   DetectDelimiter = cands(k)
440                   delimLabel = labels(k)
                  End If
              End If
          End If
      Next k
450   consistency = bestShare

      LogSuccess MODULE_NAME, "DetectDelimiter", _
                 "Delimiter=" & IIf(Len(delimLabel) > 0, delimLabel, "(none)") & "; consistency=" & Format(bestShare, "0.000")
460   Exit Function
ErrHandler:
      LogError MODULE_NAME, "DetectDelimiter", Err.Number, Err.Description, Erl, _
               "k=" & k & "; i=" & i
      DetectDelimiter = ""
End Function

'----------------------------------------------------------------------
' ParseDelimited: 区切り文字形式の行を 2 次元配列 (1 始まり) に展開する
'   引用符で囲まれた値の中の区切り文字と改行に対応する
'   maxRows      解析する最大の行数 (超えた分は捨てて truncated を立てる)
'   rowWidth()   行ごとの項目数
'   quotedCount() 列ごとの引用符付きの値の数
'   multiLineCount セル内改行を含む値の数
'----------------------------------------------------------------------
Public Function ParseDelimited(ByRef lines() As String, ByVal delim As String, ByVal maxRows As Long, _
                                 ByRef rowWidth() As Long, ByRef rowCount As Long, ByRef colCount As Long, _
                                 ByRef quotedCount() As Long, ByRef truncated As Boolean, _
                                 ByRef multiLineCount As Long) As Variant
      On Error GoTo ErrHandler
      Dim rows As Collection
      Dim fields() As String
      Dim fieldCount As Long
      Dim fieldQuoted() As Boolean
      Dim i As Long
      Dim p As Long
      Dim ln As String
      Dim ch As String
      Dim buf As String
      Dim inQuote As Boolean
      Dim wasQuoted As Boolean
      Dim hadNewline As Boolean
      Dim r As Long
      Dim c As Long
      Dim grid() As Variant
      Dim rowItem As Variant
      Dim simpleFields() As String

10    Set rows = New Collection
20    ReDim quotedCount(1 To 1)
30    rowCount = 0
40    colCount = 0
50    truncated = False
60    multiLineCount = 0
70    i = LBound(lines)

80    Do While i <= UBound(lines)
90        If rowCount >= maxRows Then
100           truncated = True
110           Exit Do
          End If
120       ln = lines(i)

130       If InStr(ln, QUOTE_CHAR) = 0 Then
              ' 引用符の無い行は Split で速く分ける
140           simpleFields = Split(ln, delim)
150           fieldCount = UBound(simpleFields) - LBound(simpleFields) + 1
160           If fieldCount = 0 Then
170               ReDim fields(1 To 1)
180               fieldCount = 0
              Else
190               ReDim fields(1 To fieldCount)
200               For c = 1 To fieldCount
210                   fields(c) = simpleFields(LBound(simpleFields) + c - 1)
                  Next c
              End If
220           ReDim fieldQuoted(1 To IIf(fieldCount > 0, fieldCount, 1))
230           i = i + 1
          Else
              ' 引用符を 1 文字ずつ解釈する。閉じていなければ次の行を連結する
240           fieldCount = 0
250           ReDim fields(1 To 1)
260           ReDim fieldQuoted(1 To 1)
270           buf = ""
280           inQuote = False
290           wasQuoted = False
300           hadNewline = False
310           p = 1
320           Do
330               If p > Len(ln) Then
340                   If inQuote And i < UBound(lines) Then
                          ' 値の中の改行
350                       i = i + 1
360                       ln = lines(i)
370                       buf = buf & vbLf
380                       hadNewline = True
390                       p = 1
                      Else
400                       Exit Do
                      End If
                  Else
410                   ch = Mid$(ln, p, 1)
420                   If inQuote Then
430                       If ch = QUOTE_CHAR Then
440                           If Mid$(ln, p + 1, 1) = QUOTE_CHAR Then
450                               buf = buf & QUOTE_CHAR
460                               p = p + 1
                              Else
470                               inQuote = False
                              End If
                          Else
480                           buf = buf & ch
                          End If
                      ElseIf ch = QUOTE_CHAR And Len(buf) = 0 And Not wasQuoted Then
490                       inQuote = True
500                       wasQuoted = True
                      ElseIf ch = delim Then
510                       fieldCount = fieldCount + 1
520                       ReDim Preserve fields(1 To fieldCount)
530                       ReDim Preserve fieldQuoted(1 To fieldCount)
540                       fields(fieldCount) = buf
550                       fieldQuoted(fieldCount) = wasQuoted
560                       If hadNewline Then multiLineCount = multiLineCount + 1
570                       buf = ""
580                       wasQuoted = False
590                       hadNewline = False
                      Else
600                       buf = buf & ch
                      End If
610                   p = p + 1
                  End If
              Loop
620           fieldCount = fieldCount + 1
630           ReDim Preserve fields(1 To fieldCount)
640           ReDim Preserve fieldQuoted(1 To fieldCount)
650           fields(fieldCount) = buf
660           fieldQuoted(fieldCount) = wasQuoted
670           If hadNewline Then multiLineCount = multiLineCount + 1
680           i = i + 1
          End If

          ' 列ごとの引用符の数
690       If fieldCount > colCount Then
700           colCount = fieldCount
710           ReDim Preserve quotedCount(1 To colCount)
          End If
720       For c = 1 To fieldCount
730           If fieldQuoted(c) Then quotedCount(c) = quotedCount(c) + 1
          Next c

740       rowCount = rowCount + 1
750       If fieldCount = 0 Then
760           rows.Add Array()
          Else
770           rows.Add fields
          End If
      Loop

      ' ---- 2 次元配列にする (足りない項目は Empty のまま) ----
780   If rowCount = 0 Or colCount = 0 Then
790       ReDim grid(1 To 1, 1 To 1)
800       ReDim rowWidth(1 To 1)
      Else
810       ReDim grid(1 To rowCount, 1 To colCount)
820       ReDim rowWidth(1 To rowCount)
830       r = 0
840       For Each rowItem In rows
850           r = r + 1
860           If IsArray(rowItem) Then
870               If UBound(rowItem) >= LBound(rowItem) Then
880                   rowWidth(r) = UBound(rowItem) - LBound(rowItem) + 1
890                   For c = 1 To rowWidth(r)
900                       grid(r, c) = rowItem(LBound(rowItem) + c - 1)
                      Next c
                  End If
              End If
          Next rowItem
      End If
910   ParseDelimited = grid

      LogSuccess MODULE_NAME, "ParseDelimited", _
                 "Parsed " & rowCount & " rows x " & colCount & " cols; truncated=" & truncated & "; multiLine=" & multiLineCount
920   Exit Function
ErrHandler:
      LogError MODULE_NAME, "ParseDelimited", Err.Number, Err.Description, Erl, _
               "i=" & i & "; rowCount=" & rowCount & "; colCount=" & colCount & "; p=" & p
      rowCount = 0
      colCount = 0
      ReDim grid(1 To 1, 1 To 1)
      ParseDelimited = grid
End Function

'----------------------------------------------------------------------
' LinesToSingleColumn: 区切りの無い行を 1 列の 2 次元配列にする
'----------------------------------------------------------------------
Public Function LinesToSingleColumn(ByRef lines() As String, ByVal maxRows As Long, _
                                      ByRef rowWidth() As Long, ByRef rowCount As Long, _
                                      ByRef truncated As Boolean) As Variant
      On Error GoTo ErrHandler
      Dim grid() As Variant
      Dim i As Long
      Dim total As Long

10    total = UBound(lines) - LBound(lines) + 1
20    rowCount = total
30    truncated = False
40    If rowCount > maxRows Then
50        rowCount = maxRows
60        truncated = True
      End If
70    If rowCount < 1 Then
80        rowCount = 0
90        ReDim grid(1 To 1, 1 To 1)
100       ReDim rowWidth(1 To 1)
      Else
110       ReDim grid(1 To rowCount, 1 To 1)
120       ReDim rowWidth(1 To rowCount)
130       For i = 1 To rowCount
140           grid(i, 1) = lines(LBound(lines) + i - 1)
150           rowWidth(i) = IIf(Len(lines(LBound(lines) + i - 1)) > 0, 1, 0)
          Next i
      End If
160   LinesToSingleColumn = grid

      LogSuccess MODULE_NAME, "LinesToSingleColumn", "Built single column grid with " & rowCount & " rows"
170   Exit Function
ErrHandler:
      LogError MODULE_NAME, "LinesToSingleColumn", Err.Number, Err.Description, Erl, _
               "i=" & i & "; rowCount=" & rowCount
      rowCount = 0
      ReDim grid(1 To 1, 1 To 1)
      LinesToSingleColumn = grid
End Function
