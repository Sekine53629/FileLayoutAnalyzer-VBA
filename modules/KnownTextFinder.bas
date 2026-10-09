Attribute VB_Name = "KnownTextFinder"
Option Explicit

'======================================================================
' KnownTextFinder モジュール
' 中身の分かっている文字列 (薬品名など) を手がかりに、ファイルの文字コードと位置を確かめる。
'   - 文字列ごとに、半角・入力のまま・全角の形を作る (ｴｽｼﾀﾛﾌﾟﾗﾑ / エスシタロプラム)
'     発注データでは Shift_JIS の半角カナが多いので、半角 → Shift_JIS を最初に探す
'   - それぞれを Shift_JIS、EUC-JP、JIS (ESC なしの 7 ビット)、UTF-8、UTF-16 LE / BE の
'     バイト列にして、ファイルの中を探す
'   - 見つかった文字コードと位置 (行番号・行の中のバイト位置) だけを返す。前後のデータは出さない
' 汎用機の漢字コード (JEF・KEIS) は、JIS の漢字の範囲では EUC-JP と同じバイトになることが多い。
' IBM 漢字は Windows に変換表が無いので探せない (どれでも見つからないこと自体が手がかりになる)。
'
' 使い方
'   RunKnownTextSearch を実行し、ファイルを選んで、探す文字列を空白区切りで入力する。
'   FindKnownText(パス, 文字列) は結果の文を返す (画面表示なし)。
'
' 依存: ErrorLogger.bas
'======================================================================

Private Const MODULE_NAME As String = "KnownTextFinder"

' ---- 設定 ----
Private Const MAX_SHOW As Long = 5                 ' 1 つの文字コードで表示する位置の数
Private Const BYTE_LF As Long = 10
Private Const BYTE_SO As Long = 14
Private Const BYTE_EUC_FIRST As Long = &HA1
Private Const JIS7_MASK As Long = &H7F
Private Const AD_TYPE_BINARY As Long = 1
Private Const AD_TYPE_TEXT As Long = 2
Private Const UTF8_BOM_BYTES As Long = 3
Private Const UTF16_BOM_BYTES As Long = 2
Private Const UTF8_BOM_1 As Long = &HEF
Private Const UTF16_BOM_FF As Long = &HFF
Private Const UTF16_BOM_FE As Long = &HFE
Private Const WIDE_SPACE_CODE As Long = &H3000
Private Const TERM_SEP As String = " "
Private Const FORM_COUNT As Long = 3              ' 半角・入力のまま・全角 (この順に探して表示する)

' ---- 文字コード (ADODB.Stream の Charset) と表示名 ----
Private Const CS_SJIS As String = "Shift_JIS"
Private Const CS_EUC As String = "EUC-JP"
Private Const CS_UTF8 As String = "UTF-8"
Private Const CS_UTF16LE As String = "Unicode"
Private Const CS_UTF16BE As String = "UnicodeFFFE"
Private Const CS_JIS7 As String = "JIS7"           ' EUC-JP から作る (Charset ではない)
Private Const LBL_SJIS As String = "Shift_JIS"
Private Const LBL_EUC As String = "EUC-JP (JEF・KEIS もこの形が多い)"
Private Const LBL_JIS7 As String = "JIS 7 ビット (ESC なし)"
Private Const LBL_UTF8 As String = "UTF-8"
Private Const LBL_UTF16LE As String = "UTF-16 LE"
Private Const LBL_UTF16BE As String = "UTF-16 BE"
Private Const ENCODING_LIST As String = "Shift_JIS|EUC-JP|JIS7|UTF-8|Unicode|UnicodeFFFE"
Private Const LIST_SEP As String = "|"

' ---- 画面と結果の文言 ----
Private Const MSG_PICK_TITLE As String = "文字コードを確かめるファイルを選択"
Private Const MSG_INPUT_PROMPT As String = "ファイルに含まれているはずの文字列を入力 (空白区切りで複数可。全角・半角の両方で探す)" & vbCrLf & "例: エスシタロプラム 明治"
Private Const MSG_INPUT_TITLE As String = "探す文字列"
Private Const MSG_INPUT_DEFAULT As String = "エスシタロプラム 明治"
Private Const MSG_CANCELLED As String = "中止しました。"
Private Const MSG_FAILED As String = "照合中にエラーが発生しました。C:\VBAErrorLogs\ のエラーログを確認してください。"
Private Const RES_FILE As String = "ファイル: "
Private Const RES_BYTES As String = " バイト / "
Private Const RES_LINES As String = " 行"
Private Const RES_TERM_OPEN As String = "["
Private Const RES_TERM_CLOSE As String = "]"
Private Const RES_FORM_INPUT As String = "入力のまま"
Private Const RES_FORM_WIDE As String = "全角"
Private Const RES_FORM_NARROW As String = "半角"
Private Const RES_COUNT As String = " 件 / "
Private Const RES_LEN As String = " バイト / "
Private Const RES_LINE As String = "行 "
Private Const RES_POS As String = " の "
Private Const RES_POS_UNIT As String = " バイト目"
Private Const RES_AFTER_SO As String = " (直前が SO 0x0E)"
Private Const RES_MORE_1 As String = " ほか "
Private Const RES_MORE_2 As String = " 件"
Private Const RES_LIST_SEP As String = ", "
Private Const RES_NONE_TERM As String = "  見つからない"
Private Const RES_NONE_ALL As String = "どの文字コードでも見つからない。IBM 漢字などの汎用機の漢字コードか、名称が略称などで登録されている可能性がある。"
Private Const RES_INDENT As String = "  "
Private Const RES_LABEL_SEP As String = ": "


'----------------------------------------------------------------------
' RunKnownTextSearch: 入口。ファイルと文字列を選んで照合し、結果を表示する
'----------------------------------------------------------------------
Public Sub RunKnownTextSearch()
      On Error GoTo ErrHandler
      Dim selected As Variant
      Dim terms As String
      Dim result As String

10    selected = Application.GetOpenFilename(, , MSG_PICK_TITLE)
20    If VarType(selected) = vbBoolean Then
30        MsgBox MSG_CANCELLED, vbInformation
          LogSuccess MODULE_NAME, "RunKnownTextSearch", "Cancelled: no file selected"
40        Exit Sub
      End If
50    terms = InputBox(MSG_INPUT_PROMPT, MSG_INPUT_TITLE, MSG_INPUT_DEFAULT)
60    If Len(Trim$(terms)) = 0 Then
70        MsgBox MSG_CANCELLED, vbInformation
          LogSuccess MODULE_NAME, "RunKnownTextSearch", "Cancelled: no search text"
80        Exit Sub
      End If

90    result = FindKnownText(CStr(selected), terms)
100   MsgBox result, vbInformation

      LogSuccess MODULE_NAME, "RunKnownTextSearch", "Searched known text in " & CStr(selected)
110   Exit Sub
ErrHandler:
      LogError MODULE_NAME, "RunKnownTextSearch", Err.Number, Err.Description, Erl, _
               "terms=" & terms
      MsgBox MSG_FAILED, vbExclamation
End Sub

'----------------------------------------------------------------------
' FindKnownText: ファイルの中から文字列を各文字コードで探し、結果の文を返す
'   termsText  空白 (全角・半角) で区切った文字列
'----------------------------------------------------------------------
Public Function FindKnownText(ByVal filePath As String, ByVal termsText As String) As String
      On Error GoTo ErrHandler
      Dim stm As Object
      Dim data() As Byte
      Dim byteCount As Long
      Dim hay As String
      Dim lineStart() As Long
      Dim lineCount As Long
      Dim i As Long
      Dim terms() As String
      Dim t As Long
      Dim forms(1 To FORM_COUNT) As String
      Dim formLabels(1 To FORM_COUNT) As String
      Dim f As Long
      Dim g As Long
      Dim isDup As Boolean
      Dim encs() As String
      Dim e As Long
      Dim needleBytes() As Byte
      Dim ok As Boolean
      Dim needle As String
      Dim hitPos As Long
      Dim hitCount As Long
      Dim places As String
      Dim lo As Long
      Dim hi As Long
      Dim midIdx As Long
      Dim rep As String
      Dim termFound As Boolean
      Dim anyFound As Boolean
      Dim label As String

      ' ---- ファイルをバイト列で読む (日本語を含むパスにも対応) ----
10    Set stm = CreateObject("ADODB.Stream")
20    stm.Type = AD_TYPE_BINARY
30    stm.Open
40    stm.LoadFromFile filePath
50    byteCount = stm.Size
60    If byteCount > 0 Then data = stm.Read
70    stm.Close
80    Set stm = Nothing

      ' 行の先頭位置 (0 始まり)。LF で数える
90    ReDim lineStart(1 To 1)
100   lineStart(1) = 0
110   lineCount = 1
120   For i = 0 To byteCount - 1
130       If data(i) = BYTE_LF And i < byteCount - 1 Then
140           lineCount = lineCount + 1
150           If lineCount > UBound(lineStart) Then ReDim Preserve lineStart(1 To lineCount * 2)
160           lineStart(lineCount) = i + 1
          End If
      Next i
      ' バイト列をそのまま文字列に入れて、InStrB でバイト単位に探す
170   If byteCount > 0 Then hay = data

180   rep = RES_FILE & filePath & vbCrLf & Format(byteCount, "#,##0") & RES_BYTES & Format(lineCount, "#,##0") & RES_LINES & vbCrLf
190   encs = Split(ENCODING_LIST, LIST_SEP)
200   terms = Split(Trim$(Replace(termsText, ChrW(WIDE_SPACE_CODE), TERM_SEP)), TERM_SEP)

210   For t = 0 To UBound(terms)
220       If Len(terms(t)) > 0 Then
230           rep = rep & vbCrLf & RES_TERM_OPEN & terms(t) & RES_TERM_CLOSE & vbCrLf
240           termFound = False
              ' 半角・入力のまま・全角 (同じものは 1 つにまとめる)。
              ' 発注データでは Shift_JIS の半角カナが多いので、半角を先に探す
250           forms(1) = StrConv(terms(t), vbNarrow)
260           formLabels(1) = RES_FORM_NARROW
270           forms(2) = terms(t)
280           formLabels(2) = RES_FORM_INPUT
290           forms(3) = StrConv(terms(t), vbWide)
300           formLabels(3) = RES_FORM_WIDE
310           For f = 1 To FORM_COUNT
320               isDup = False
330               For g = 1 To f - 1
340                   If forms(g) = forms(f) Then isDup = True
                  Next g
350               If Not isDup Then
360                   For e = 0 To UBound(encs)
370                       needleBytes = EncodeText(forms(f), encs(e), ok)
380                       If ok Then
390                           needle = needleBytes
400                           hitCount = 0
410                           places = ""
420                           hitPos = InStrB(1, hay, needle, vbBinaryCompare)
430                           Do While hitPos > 0
440                               hitCount = hitCount + 1
450                               If hitCount <= MAX_SHOW Then
                                      ' 行番号を二分探索で求める (hitPos は 1 始まり)
460                                   lo = 1
470                                   hi = lineCount
480                                   Do While lo < hi
490                                       midIdx = (lo + hi + 1) \ 2
500                                       If lineStart(midIdx) <= hitPos - 1 Then lo = midIdx Else hi = midIdx - 1
                                      Loop
510                                   places = places & IIf(Len(places) > 0, RES_LIST_SEP, "") & RES_LINE & lo & RES_POS & _
                                               (hitPos - lineStart(lo)) & RES_POS_UNIT
520                                   If hitPos > 1 Then
530                                       If data(hitPos - 2) = BYTE_SO Then places = places & RES_AFTER_SO
                                      End If
                                  End If
540                               hitPos = InStrB(hitPos + 1, hay, needle, vbBinaryCompare)
                              Loop
550                           If hitCount > 0 Then
560                               termFound = True
570                               anyFound = True
580                               Select Case encs(e)
                                      Case CS_SJIS: label = LBL_SJIS
                                      Case CS_EUC: label = LBL_EUC
                                      Case CS_JIS7: label = LBL_JIS7
                                      Case CS_UTF8: label = LBL_UTF8
                                      Case CS_UTF16LE: label = LBL_UTF16LE
                                      Case Else: label = LBL_UTF16BE
                                  End Select
590                               rep = rep & RES_INDENT & formLabels(f) & " " & forms(f) & " / " & label & RES_LABEL_SEP & _
                                        hitCount & RES_COUNT & (UBound(needleBytes) + 1) & RES_LEN & places & _
                                        IIf(hitCount > MAX_SHOW, RES_MORE_1 & (hitCount - MAX_SHOW) & RES_MORE_2, "") & vbCrLf
                              End If
                          End If
                      Next e
                  End If
              Next f
600           If Not termFound Then rep = rep & RES_NONE_TERM & vbCrLf
          End If
      Next t
610   If Not anyFound Then rep = rep & vbCrLf & RES_NONE_ALL & vbCrLf
620   FindKnownText = rep

      LogSuccess MODULE_NAME, "FindKnownText", "Searched " & (UBound(terms) + 1) & " terms in " & byteCount & _
                 " bytes; found=" & anyFound & "; file=" & filePath
630   Exit Function
ErrHandler:
      LogError MODULE_NAME, "FindKnownText", Err.Number, Err.Description, Erl, _
               "filePath=" & filePath & "; byteCount=" & byteCount & "; t=" & t & "; e=" & e
      FindKnownText = MSG_FAILED
End Function

'----------------------------------------------------------------------
' EncodeText: 文字列を指定の文字コードのバイト列にする (BOM は除く)
'   charset  ADODB.Stream の Charset 名、または CS_JIS7 (EUC-JP の各バイトから最上位ビットを落とす)
'   ok       符号化できたら True (JIS7 は全角だけの文字列のときだけ作れる)
'----------------------------------------------------------------------
Private Function EncodeText(ByVal text As String, ByVal charset As String, ByRef ok As Boolean) As Byte()
      On Error GoTo ErrHandler
      Dim stm As Object
      Dim raw() As Byte
      Dim out() As Byte
      Dim skip As Long
      Dim n As Long
      Dim i As Long

10    ok = False
20    Set stm = CreateObject("ADODB.Stream")
30    stm.Type = AD_TYPE_TEXT
40    stm.Charset = IIf(charset = CS_JIS7, CS_EUC, charset)
50    stm.Open
60    stm.WriteText text
70    stm.Position = 0
80    stm.Type = AD_TYPE_BINARY
90    raw = stm.Read
100   stm.Close
110   Set stm = Nothing

120   n = UBound(raw) + 1
      ' ADODB.Stream が付ける BOM を除く
130   If charset = CS_UTF8 And n >= UTF8_BOM_BYTES Then
140       If raw(0) = UTF8_BOM_1 Then skip = UTF8_BOM_BYTES
      ElseIf (charset = CS_UTF16LE Or charset = CS_UTF16BE) And n >= UTF16_BOM_BYTES Then
150       If (raw(0) = UTF16_BOM_FF And raw(1) = UTF16_BOM_FE) Or (raw(0) = UTF16_BOM_FE And raw(1) = UTF16_BOM_FF) Then skip = UTF16_BOM_BYTES
      End If
160   If n - skip <= 0 Then
          LogSuccess MODULE_NAME, "EncodeText", "Nothing to encode as " & charset
170       Exit Function
      End If
180   ReDim out(0 To n - skip - 1)
190   For i = 0 To n - skip - 1
200       out(i) = raw(i + skip)
      Next i

210   If charset = CS_JIS7 Then
          ' 全角 (0xA1 以上の 2 バイト) だけのときに限り、最上位ビットを落として JIS の 7 ビットにする
220       For i = 0 To UBound(out)
230           If out(i) < BYTE_EUC_FIRST Then
                  LogSuccess MODULE_NAME, "EncodeText", "JIS7 not applicable (non-wide characters)"
240               Exit Function
              End If
250           out(i) = out(i) And JIS7_MASK
          Next i
      End If
260   ok = True
270   EncodeText = out

      LogSuccess MODULE_NAME, "EncodeText", "Encoded " & Len(text) & " chars as " & charset & " (" & (UBound(out) + 1) & " bytes)"
280   Exit Function
ErrHandler:
      LogError MODULE_NAME, "EncodeText", Err.Number, Err.Description, Erl, "charset=" & charset
      ok = False
End Function
