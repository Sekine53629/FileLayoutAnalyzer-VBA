Attribute VB_Name = "FixedLengthAnalyzer"
Option Explicit

'======================================================================
' FixedLengthAnalyzer モジュール
' 区切り文字の無いテキスト (固定長の DAT など) をバイト位置で解析する。
'   - 行の長さ (バイト) の分布、固定長かどうか
'   - 先頭のバイトがレコード種別で、種別ごとに長さが違う形式の判定
'   - 位置ごとの文字の種類から、項目の境界を推定する
'
' 境界は「文字の種類が変わる位置」で推定するため、数字の項目が続く場合などは
' 1 つにまとまる。レイアウト定義書と突き合わせるための下書きとして使う。
' 値は書式パターン (9=数字 A=英字 N(全角)=全角 ｶ(半角)=半角カナ _=空白) に置き換える。
'
' 依存: ErrorLogger.bas, GridAnalyzer.bas (表の見出しの文言のみ)
'======================================================================

Private Const MODULE_NAME As String = "FixedLengthAnalyzer"

' ---- 判定の閾値 ----
Private Const FIXED_MIN_SHARE As Double = 0.8      ' 最頻の長さの行がこの割合以上なら固定長とみなす
Private Const RECTYPE_MAX_KINDS As Long = 30       ' レコード種別とみなす先頭バイトの最大の種類数
Private Const RECTYPE_KEY_BYTES As Long = 2        ' レコード種別として見る先頭のバイト数
Private Const RECTYPE_MAX_SINGLES As Long = 2      ' 1 行しかない種別の許容数 (先頭・末尾のレコード相当)
Private Const DATE8_LEN As Long = 8                ' yyyymmdd の長さ
Private Const YEAR_MIN As Long = 1900
Private Const YEAR_MAX As Long = 2100
Private Const MONTH_MAX As Long = 12
Private Const DAY_MAX As Long = 31
Private Const DECIMAL_BASE As Long = 10
Private Const MAX_SAMPLE_LINES As Long = 50000     ' 位置ごとの集計に使う最大の行数
Private Const MAX_LINE_BYTES As Long = 8192        ' 解析する 1 行の最大のバイト数
Private Const FILL_MIN_RATIO As Double = 0.02      ' 空白以外がこの割合未満の位置は未使用 (空白) とする
Private Const PATTERN_CAP As Long = 200            ' 書式パターンを数える種類の上限
Private Const PATTERN_SHOW As Long = 3             ' 載せる書式パターンの数
Private Const PATTERN_SHOW_LEN As Long = 40        ' 載せる書式パターンの最大の文字数
Private Const RUN_COMPRESS_MIN As Long = 5         ' この数以上続く記号を個数表記にする (9999 は残す)
Private Const LEN_SHOW_MAX As Long = 10            ' 長さの分布を載せる種類の上限
Private Const CLASS_COUNT As Long = 8              ' 文字の種類の数

' ---- 文字の種類 (位置ごとの集計の添字) ----
Private Const MERGE_LEFT_SHARE As Double = 0.9    ' つないだ範囲が左寄せの可変長である行の割合 (これ以上でつなぐ)
Private Const CLS_SPACE As Long = 0
Private Const CLS_DIGIT As Long = 1
Private Const CLS_ALPHA As Long = 2
Private Const CLS_WIDE As Long = 3
Private Const CLS_KANA As Long = 4
Private Const CLS_SYMBOL As Long = 5
Private Const CLS_NUL As Long = 6
Private Const CLS_OTHER As Long = 7

' ---- バイト値 ----
Private Const BYTE_NUL As Long = 0
Private Const BYTE_LF As Long = 10
Private Const BYTE_CR As Long = 13
Private Const BYTE_SPACE As Long = 32
Private Const BYTE_DIGIT_0 As Long = &H30
Private Const BYTE_DIGIT_9 As Long = &H39
Private Const BYTE_UPPER_A As Long = &H41
Private Const BYTE_UPPER_Z As Long = &H5A
Private Const BYTE_LOWER_A As Long = &H61
Private Const BYTE_LOWER_Z As Long = &H7A
Private Const BYTE_SYMBOL_FIRST As Long = &H21
Private Const BYTE_SYMBOL_LAST As Long = &H7E
Private Const BYTE_KANA_FIRST As Long = &HA1
Private Const BYTE_KANA_LAST As Long = &HDF
Private Const BYTE_LEAD1_FIRST As Long = &H81
Private Const BYTE_LEAD1_LAST As Long = &H9F
Private Const BYTE_LEAD2_FIRST As Long = &HE0
Private Const BYTE_LEAD2_LAST As Long = &HFC
Private Const BYTE_HIGH As Long = &H80
Private Const CODE_WIDE_MARK As Long = &HFF2E&
Private Const CODE_HALFKANA_MARK As Long = &HFF76&

' ---- パターンの記号 ----
Private Const MARK_DIGIT As String = "9"
Private Const MARK_UPPER As String = "A"
Private Const MARK_LOWER As String = "a"
Private Const MARK_SPACE As String = "_"
Private Const MARK_NUL As String = "@"
Private Const MARK_OTHER As String = "?"

' ---- 報告の文言 ----
Private Const RPT_SEC_STRUCT As String = "[構造]"
Private Const RPT_SEC_ITEMS As String = "[項目一覧 (推定)]"
Private Const RPT_SEC_RECTYPE As String = "[レコード種別]"
Private Const RPT_FORMAT As String = "  形式            : "
Private Const RPT_MODE_FIXED As String = "固定長 (区切り文字なし)"
Private Const RPT_MODE_RECTYPE As String = "固定長 (先頭のバイトがレコード種別で、種別ごとに長さが異なる)"
Private Const RPT_MODE_NONEWLINE As String = "改行なし (1 レコードの長さを判別できない)"
Private Const RPT_LINES As String = "  行数            : "
Private Const RPT_LEN_DIST As String = "  行の長さ (バイト): "
Private Const RPT_RECLEN As String = "  レコード長      : "
Private Const RPT_BYTES_UNIT As String = " バイト"
Private Const RPT_LINES_UNIT As String = " 行"
Private Const RPT_NOTE_BOUNDARY As String = "  注: 境界は文字の種類が変わる位置から推定している。数字の項目が続く場合などは 1 つにまとまるので、レイアウト定義書と突き合わせること。"
Private Const RPT_NOTE_SAMPLE As String = "  注: 位置ごとの集計は先頭の一部の行で行っている"
Private Const RPT_ODD_LINES As String = "  長さが違う行    : "
Private Const RPT_TABLE_HEAD As String = "| No | 位置 (バイト) | 長さ | 種別 | 空 | 書式パターン (上位3) | 備考 |"
Private Const RPT_TABLE_RULE As String = "|---:|---|---:|---|---:|---|---|"
Private Const RPT_RECTYPE_HEAD As String = "| 種別 (先頭 2 バイトのパターン) | 行数 | 長さ (バイト) | 最初の行 |"
Private Const RPT_RECTYPE_RULE As String = "|---|---:|---|---:|"
Private Const RPT_RECTYPE_SEC As String = "[レコード種別 "
Private Const RPT_RECTYPE_SEC_CLOSE As String = " の項目一覧 (推定)]"
Private Const TYPE_SPACE As String = "空白 (未使用か予備)"
Private Const TYPE_DIGIT As String = "数字"
Private Const TYPE_ALPHA As String = "英字"
Private Const TYPE_WIDE As String = "全角"
Private Const TYPE_KANA As String = "半角カナ"
Private Const TYPE_SYMBOL As String = "記号"
Private Const TYPE_NUL As String = "NUL (0x00)"
Private Const TYPE_OTHER As String = "その他のバイト"
Private Const TYPE_TEXT_VAR As String = "英数記号 (可変長。規格・名称など)"
Private Const RMK_RIGHT As String = "右寄せ (前に空白)"
Private Const RMK_LEFT As String = "左寄せ (後ろに空白)"
Private Const RMK_ZERO As String = "前ゼロ詰め"
Private Const RMK_BLANK As String = "未使用、または前後の項目の空白詰め"
Private Const RMK_DATE8 As String = "日付 (yyyymmdd) の可能性"
Private Const RMK_SEP As String = "; "
Private Const SUM_FIXED As String = "固定長 / レコード長 "
Private Const SUM_ITEMS As String = " バイト / 推定項目 "
Private Const SUM_LINES As String = " / 行数 "
Private Const SUM_RECTYPE As String = "固定長 (レコード種別 "
Private Const SUM_RECTYPE_UNIT As String = " 種類) / 行数 "
Private Const SUM_NONEWLINE As String = "改行なし / "
Private Const SUM_NOTFIXED As String = "行の長さが揃っていない"
Private Const SUM_ERROR As String = "解析中にエラー (ログを確認)"
Private Const RPT_ERROR_LINE As String = "  解析中にエラーが発生した。C:\VBAErrorLogs\ を確認すること。"


'----------------------------------------------------------------------
' AnalyzeFixedLength: バイト列を固定長として解析し、報告文を返す
'   startOffset  BOM などを除いた開始位置
'   isSjis       Shift_JIS (全角を 2 バイトとして扱う) なら True
'   isFixed      固定長と判断できたら True (呼び出し元は False なら別の解析に回す)
'   summaryOut   結論欄に載せる 1 行
'----------------------------------------------------------------------
Public Function AnalyzeFixedLength(ByRef data() As Byte, ByVal byteCount As Long, ByVal startOffset As Long, _
                                     ByVal isSjis As Boolean, ByRef isFixed As Boolean, _
                                     ByRef summaryOut As String) As String
      On Error GoTo ErrHandler
      Dim rep As String
      Dim lineStart() As Long
      Dim lineLen() As Long
      Dim lineCount As Long
      Dim capacity As Long
      Dim i As Long
      Dim st As Long
      Dim ln As Long
      Dim lenDict As Object
      Dim key As Variant
      Dim modeLen As Long
      Dim modeFreq As Long
      Dim lenText As String
      Dim shown As Long
      Dim kinds As Object
      Dim kindLines As Object
      Dim kindLens As Object
      Dim kindFirst As Object
      Dim kindKey As String
      Dim kb As Long
      Dim lineColl As Collection
      Dim lenCounts As Object
      Dim isRecType As Boolean
      Dim best As Long
      Dim w As Variant
      Dim idx() As Long
      Dim cnt As Long
      Dim item As Variant
      Dim oddCount As Long
      Dim itemCount As Long
      Dim kindLabel As String
      Dim singles As Long

      ' ---- 行に分ける (LF で区切り、直前の CR は除く) ----
10    capacity = 1024
20    ReDim lineStart(1 To capacity)
30    ReDim lineLen(1 To capacity)
40    st = startOffset
50    For i = startOffset To byteCount - 1
60        If data(i) = BYTE_LF Then
70            ln = i - st
80            If ln > 0 Then
90                If data(i - 1) = BYTE_CR Then ln = ln - 1
              End If
100           lineCount = lineCount + 1
110           If lineCount > capacity Then
120               capacity = capacity * 2
130               ReDim Preserve lineStart(1 To capacity)
140               ReDim Preserve lineLen(1 To capacity)
              End If
150           lineStart(lineCount) = st
160           lineLen(lineCount) = ln
170           st = i + 1
          End If
      Next i
180   If st < byteCount Then
190       lineCount = lineCount + 1
200       If lineCount > capacity Then
210           capacity = capacity + 1
220           ReDim Preserve lineStart(1 To capacity)
230           ReDim Preserve lineLen(1 To capacity)
          End If
240       lineStart(lineCount) = st
250       lineLen(lineCount) = byteCount - st
260       If lineLen(lineCount) > 0 Then
270           If data(byteCount - 1) = BYTE_CR Then lineLen(lineCount) = lineLen(lineCount) - 1
          End If
      End If

280   rep = RPT_SEC_STRUCT & vbCrLf

      ' ---- 改行が無い ----
290   If lineCount <= 1 Then
300       isFixed = True
310       rep = rep & RPT_FORMAT & RPT_MODE_NONEWLINE & vbCrLf & _
                RPT_RECLEN & (byteCount - startOffset) & RPT_BYTES_UNIT & vbCrLf
320       summaryOut = SUM_NONEWLINE & (byteCount - startOffset) & RPT_BYTES_UNIT
330       AnalyzeFixedLength = rep
          LogSuccess MODULE_NAME, "AnalyzeFixedLength", "No line breaks; bytes=" & byteCount
340       Exit Function
      End If

      ' ---- 長さの分布 ----
350   Set lenDict = CreateObject("Scripting.Dictionary")
360   For i = 1 To lineCount
370       If lineLen(i) > 0 Then lenDict(lineLen(i)) = lenDict(lineLen(i)) + 1
      Next i
380   For Each key In lenDict.Keys
390       If lenDict(key) > modeFreq Then
400           modeFreq = lenDict(key)
410           modeLen = key
          End If
      Next key
420   For Each key In lenDict.Keys
430       shown = shown + 1
440       If shown > LEN_SHOW_MAX Then
450           lenText = lenText & ", ..."
460           Exit For
          End If
470       lenText = lenText & IIf(Len(lenText) > 0, ", ", "") & key & " (" & lenDict(key) & RPT_LINES_UNIT & ")"
      Next key

      ' ---- レコード種別 (先頭 2 バイトのパターンごとに長さが揃っているか) ----
480   If lenDict.Count > 1 And modeFreq < lineCount * FIXED_MIN_SHARE Then
490       Set kinds = CreateObject("Scripting.Dictionary")
500       Set kindLines = CreateObject("Scripting.Dictionary")
510       Set kindLens = CreateObject("Scripting.Dictionary")
520       Set kindFirst = CreateObject("Scripting.Dictionary")
530       isRecType = True
540       For i = 1 To lineCount
550           If lineLen(i) > 0 Then
                  ' 種別コードは書式の一部なので、英数字ならそのまま、それ以外はパターンにする
560               kindKey = ""
570               For kb = 0 To RECTYPE_KEY_BYTES - 1
580                   If kb < lineLen(i) Then
590                       If (data(lineStart(i) + kb) >= BYTE_DIGIT_0 And data(lineStart(i) + kb) <= BYTE_DIGIT_9) _
                             Or (data(lineStart(i) + kb) >= BYTE_UPPER_A And data(lineStart(i) + kb) <= BYTE_UPPER_Z) Then
600                           kindKey = kindKey & Chr$(data(lineStart(i) + kb))
                          Else
610                           kindKey = kindKey & MARK_OTHER
                          End If
                      End If
                  Next kb
620               If Not kinds.Exists(kindKey) Then
630                   If kinds.Count >= RECTYPE_MAX_KINDS Then
640                       isRecType = False
650                       Exit For
                      End If
660                   kinds.Add kindKey, 0
670                   kindLines.Add kindKey, New Collection
680                   kindLens.Add kindKey, CreateObject("Scripting.Dictionary")
690                   kindFirst.Add kindKey, i
                  End If
700               kinds(kindKey) = kinds(kindKey) + 1
710               Set lineColl = kindLines(kindKey)
720               lineColl.Add i
730               Set lenCounts = kindLens(kindKey)
740               lenCounts(lineLen(i)) = lenCounts(lineLen(i)) + 1
              End If
          Next i
750       If isRecType Then
              ' 1 行しかない種別ばかりなら、種別コードではない
760           singles = 0
770           For Each key In kinds.Keys
780               If kinds(key) = 1 Then singles = singles + 1
              Next key
790           If singles > RECTYPE_MAX_SINGLES Then isRecType = False
          End If
800       If isRecType Then
810           For Each key In kinds.Keys
820               Set lenCounts = kindLens(key)
830               best = 0
840               For Each w In lenCounts.Keys
850                   If lenCounts(w) > best Then best = lenCounts(w)
                  Next w
860               If best < kinds(key) * FIXED_MIN_SHARE Then
870                   isRecType = False
880                   Exit For
                  End If
              Next key
          End If
      End If

890   If Not isRecType And modeFreq < lineCount * FIXED_MIN_SHARE Then
          ' 固定長でもレコード種別でもない
900       isFixed = False
910       summaryOut = SUM_NOTFIXED
920       AnalyzeFixedLength = RPT_LEN_DIST & lenText & vbCrLf
          LogSuccess MODULE_NAME, "AnalyzeFixedLength", "Not fixed length; lines=" & lineCount & "; lengths=" & lenDict.Count
930       Exit Function
      End If
940   isFixed = True

950   If isRecType Then
960       rep = rep & RPT_FORMAT & RPT_MODE_RECTYPE & vbCrLf & _
                RPT_LINES & lineCount & vbCrLf & _
                RPT_LEN_DIST & lenText & vbCrLf & RPT_NOTE_BOUNDARY & vbCrLf & vbCrLf & _
                RPT_SEC_RECTYPE & vbCrLf & RPT_RECTYPE_HEAD & vbCrLf & RPT_RECTYPE_RULE & vbCrLf
970       For Each key In kinds.Keys
980           Set lenCounts = kindLens(key)
990           lenText = ""
1000          For Each w In lenCounts.Keys
1010              lenText = lenText & IIf(Len(lenText) > 0, ", ", "") & w & _
                            IIf(lenCounts.Count > 1, " (" & lenCounts(w) & RPT_LINES_UNIT & ")", "")
              Next w
1020          rep = rep & "| " & Replace(key, "|", "/") & " | " & kinds(key) & " | " & lenText & " | " & kindFirst(key) & " |" & vbCrLf
          Next key
1030      rep = rep & vbCrLf
1040      For Each key In kinds.Keys
1050          Set lineColl = kindLines(key)
1060          cnt = lineColl.Count
1070          ReDim idx(1 To cnt)
1080          i = 0
1090          For Each item In lineColl
1100              i = i + 1
1110              idx(i) = item
              Next item
1120          kindLabel = RPT_RECTYPE_SEC & key & RPT_RECTYPE_SEC_CLOSE
1130          rep = rep & kindLabel & vbCrLf & _
                    ProfilePositions(data, lineStart, lineLen, idx, cnt, isSjis, itemCount) & vbCrLf
          Next key
1140      summaryOut = SUM_RECTYPE & kinds.Count & SUM_RECTYPE_UNIT & lineCount
      Else
          ' 最頻の長さの行だけで位置を集計する
1150      ReDim idx(1 To modeFreq)
1160      cnt = 0
1170      For i = 1 To lineCount
1180          If lineLen(i) = modeLen Then
1190              cnt = cnt + 1
1200              idx(cnt) = i
              ElseIf lineLen(i) > 0 Then
1210              oddCount = oddCount + 1
              End If
          Next i
1220      rep = rep & RPT_FORMAT & RPT_MODE_FIXED & vbCrLf & _
                RPT_LINES & lineCount & vbCrLf & _
                RPT_RECLEN & modeLen & RPT_BYTES_UNIT & vbCrLf & _
                RPT_ODD_LINES & oddCount & RPT_LINES_UNIT & vbCrLf
1230      If cnt > MAX_SAMPLE_LINES Then rep = rep & RPT_NOTE_SAMPLE & vbCrLf
1240      rep = rep & RPT_NOTE_BOUNDARY & vbCrLf & vbCrLf & RPT_SEC_ITEMS & vbCrLf & _
                ProfilePositions(data, lineStart, lineLen, idx, cnt, isSjis, itemCount)
1250      summaryOut = SUM_FIXED & modeLen & SUM_ITEMS & itemCount & SUM_LINES & lineCount
      End If
1260  AnalyzeFixedLength = rep

      LogSuccess MODULE_NAME, "AnalyzeFixedLength", _
                 "Analyzed fixed length; lines=" & lineCount & "; recordType=" & isRecType & "; modeLen=" & modeLen
1270  Exit Function
ErrHandler:
      LogError MODULE_NAME, "AnalyzeFixedLength", Err.Number, Err.Description, Erl, _
               "byteCount=" & byteCount & "; i=" & i & "; lineCount=" & lineCount
      isFixed = True
      summaryOut = SUM_ERROR
      AnalyzeFixedLength = rep & RPT_ERROR_LINE & vbCrLf
End Function

'----------------------------------------------------------------------
' ProfilePositions: 指定した行の位置ごとの文字の種類を数え、項目の境界を推定して表にする
'   idx()      対象の行番号 (lineStart / lineLen の添字)
'   itemCount  推定した項目の数を返す
'----------------------------------------------------------------------
Private Function ProfilePositions(ByRef data() As Byte, ByRef lineStart() As Long, ByRef lineLen() As Long, _
                                    ByRef idx() As Long, ByVal cnt As Long, ByVal isSjis As Boolean, _
                                    ByRef itemCount As Long) As String
      On Error GoTo ErrHandler
      Dim rep As String
      Dim maxLen As Long
      Dim sample As Long
      Dim counts() As Long
      Dim clsAt() As Long
      Dim dom() As Long
      Dim i As Long
      Dim j As Long
      Dim b As Long
      Dim trail As Boolean
      Dim cls As Long
      Dim bestCls As Long
      Dim bestN As Long
      Dim segStart() As Long
      Dim segEnd() As Long
      Dim segCount As Long
      Dim sg As Long
      Dim patDict As Object
      Dim pat As String
      Dim mask As String
      Dim ch As String
      Dim prevCh As String
      Dim runLen As Long
      Dim p As Long
      Dim emptyN As Long
      Dim allSpace As Boolean
      Dim leadSpaceN As Long
      Dim trailSpaceN As Long
      Dim zeroN As Long
      Dim digitN As Long
      Dim typeText As String
      Dim patText As String
      Dim shown As Long
      Dim bestKey As String
      Dim bestPat As Long
      Dim usedPat As Object
      Dim k As Variant
      Dim rmk As String
      Dim pos As Long
      Dim date8N As Long
      Dim y As Long
      Dim m As Long
      Dim d As Long
      Dim segMerged() As Boolean
      Dim textRows As Long
      Dim leftOk As Long
      Dim crossN As Long
      Dim padN As Long
      Dim seenText As Boolean
      Dim spaceAfter As Boolean
      Dim isBad As Boolean
      Dim mergeOk As Boolean
      Dim t As Long

10    maxLen = 0
20    sample = cnt
30    If sample > MAX_SAMPLE_LINES Then sample = MAX_SAMPLE_LINES
40    For i = 1 To sample
50        If lineLen(idx(i)) > maxLen Then maxLen = lineLen(idx(i))
      Next i
60    If maxLen > MAX_LINE_BYTES Then maxLen = MAX_LINE_BYTES
70    If maxLen = 0 Then
80        itemCount = 0
90        ProfilePositions = ""
          LogSuccess MODULE_NAME, "ProfilePositions", "No bytes to profile"
100       Exit Function
      End If

      ' ---- 位置ごとに文字の種類を数える (Shift_JIS の 2 バイト目も全角として数える) ----
110   ReDim counts(0 To maxLen - 1, 0 To CLASS_COUNT - 1)
120   For i = 1 To sample
130       trail = False
140       For j = 0 To lineLen(idx(i)) - 1
150           If j >= maxLen Then Exit For
160           b = data(lineStart(idx(i)) + j)
170           If trail Then
180               cls = CLS_WIDE
190               trail = False
              ElseIf isSjis And ((b >= BYTE_LEAD1_FIRST And b <= BYTE_LEAD1_LAST) Or (b >= BYTE_LEAD2_FIRST And b <= BYTE_LEAD2_LAST)) Then
200               cls = CLS_WIDE
210               trail = True
              ElseIf b = BYTE_SPACE Then
220               cls = CLS_SPACE
              ElseIf b >= BYTE_DIGIT_0 And b <= BYTE_DIGIT_9 Then
230               cls = CLS_DIGIT
              ElseIf (b >= BYTE_UPPER_A And b <= BYTE_UPPER_Z) Or (b >= BYTE_LOWER_A And b <= BYTE_LOWER_Z) Then
240               cls = CLS_ALPHA
              ElseIf isSjis And b >= BYTE_KANA_FIRST And b <= BYTE_KANA_LAST Then
250               cls = CLS_KANA
              ElseIf b >= BYTE_SYMBOL_FIRST And b <= BYTE_SYMBOL_LAST Then
260               cls = CLS_SYMBOL
              ElseIf b = BYTE_NUL Then
270               cls = CLS_NUL
              ElseIf Not isSjis And b >= BYTE_HIGH Then
280               cls = CLS_WIDE
              Else
290               cls = CLS_OTHER
              End If
300           counts(j, cls) = counts(j, cls) + 1
          Next j
      Next i

      ' ---- 位置ごとの主な種類 (空白は詰め物として扱い、空白以外で決める) ----
310   ReDim dom(0 To maxLen - 1)
320   For j = 0 To maxLen - 1
330       bestCls = CLS_SPACE
340       bestN = 0
350       For cls = 0 To CLASS_COUNT - 1
360           If cls <> CLS_SPACE And counts(j, cls) > bestN Then
370               bestN = counts(j, cls)
380               bestCls = cls
              End If
          Next cls
390       If bestN < sample * FILL_MIN_RATIO Then bestCls = CLS_SPACE
400       dom(j) = bestCls
      Next j

      ' ---- 主な種類が変わる位置で区切る ----
410   ReDim segStart(1 To maxLen)
420   ReDim segEnd(1 To maxLen)
430   For j = 0 To maxLen - 1
440       If j = 0 Then
450           segCount = 1
460           segStart(1) = 0
          ElseIf dom(j) <> dom(j - 1) Then
470           segEnd(segCount) = j - 1
480           segCount = segCount + 1
490           segStart(segCount) = j
          End If
      Next j
500   segEnd(segCount) = maxLen - 1

      ' ---- 英数記号の項目が、文字の種類の変化で細切れになるのを防ぐ ----
      ' 隣り合う英数記号の区間をつないだ範囲で、ほぼすべての行が「文字が続き、その後は空白だけ」
      ' (左寄せの可変長。規格 10mg / 0.5%10g など) で、境界をまたいで文字が続く行があればつなぐ
510   ReDim segMerged(1 To segCount)
520   sg = 1
530   Do While sg < segCount
540       mergeOk = False
550       If (dom(segStart(sg)) = CLS_DIGIT Or dom(segStart(sg)) = CLS_ALPHA Or dom(segStart(sg)) = CLS_SYMBOL) _
             And (dom(segStart(sg + 1)) = CLS_DIGIT Or dom(segStart(sg + 1)) = CLS_ALPHA Or dom(segStart(sg + 1)) = CLS_SYMBOL) Then
560           textRows = 0
570           leftOk = 0
580           crossN = 0
590           padN = 0
600           For i = 1 To sample
610               seenText = False
620               spaceAfter = False
630               isBad = False
640               For j = segStart(sg) To segEnd(sg + 1)
650                   If j >= lineLen(idx(i)) Then
660                       b = BYTE_SPACE
                      Else
670                       b = data(lineStart(idx(i)) + j)
                      End If
680                   If b = BYTE_SPACE Then
690                       If seenText Then
700                           spaceAfter = True
                          ElseIf j = segStart(sg) Then
                              ' 先頭が空白なら右寄せ (左寄せの可変長ではない)
710                           isBad = True
                          End If
                      ElseIf b < BYTE_SPACE Or b >= BYTE_HIGH Then
720                       isBad = True
                      ElseIf spaceAfter Then
730                       isBad = True
                      Else
740                       seenText = True
                      End If
                  Next j
750               If seenText Then
760                   textRows = textRows + 1
770                   If Not isBad Then leftOk = leftOk + 1
780                   If spaceAfter Then padN = padN + 1
                      ' 境界の前後がどちらも文字 (区間をまたいで値が続く)
790                   t = segEnd(sg)
800                   If t + 1 < lineLen(idx(i)) Then
810                       If data(lineStart(idx(i)) + t) <> BYTE_SPACE And data(lineStart(idx(i)) + t + 1) <> BYTE_SPACE Then crossN = crossN + 1
                      End If
                  End If
              Next i
820           If textRows > 0 And crossN > 0 And padN > 0 Then
830               If leftOk >= textRows * MERGE_LEFT_SHARE Then mergeOk = True
              End If
          End If
840       If mergeOk Then
850           segEnd(sg) = segEnd(sg + 1)
860           segMerged(sg) = True
870           For t = sg + 1 To segCount - 1
880               segStart(t) = segStart(t + 1)
890               segEnd(t) = segEnd(t + 1)
900               segMerged(t) = segMerged(t + 1)
              Next t
910           segCount = segCount - 1
          Else
920           sg = sg + 1
          End If
      Loop
930   itemCount = segCount

      ' ---- 区間ごとのパターン ----
940   rep = RPT_TABLE_HEAD & vbCrLf & RPT_TABLE_RULE & vbCrLf
950   For sg = 1 To segCount
960       Set patDict = CreateObject("Scripting.Dictionary")
970       emptyN = 0
980       date8N = 0
990       leadSpaceN = 0
1000      trailSpaceN = 0
1010      zeroN = 0
1020      digitN = 0
1030      For i = 1 To sample
1040          mask = ""
1050          allSpace = True
1060          trail = False
              ' 区間の手前から 2 バイト文字の途中で始まっていないかを確かめる
1070          For j = 0 To segStart(sg) - 1
1080              If j >= lineLen(idx(i)) Then Exit For
1090              b = data(lineStart(idx(i)) + j)
1100              If trail Then
1110                  trail = False
                  ElseIf isSjis And ((b >= BYTE_LEAD1_FIRST And b <= BYTE_LEAD1_LAST) Or (b >= BYTE_LEAD2_FIRST And b <= BYTE_LEAD2_LAST)) Then
1120                  trail = True
                  End If
              Next j
1130          For j = segStart(sg) To segEnd(sg)
1140              If j >= lineLen(idx(i)) Then Exit For
1150              b = data(lineStart(idx(i)) + j)
1160              If trail Then
1170                  ch = ""
1180                  trail = False
                  ElseIf isSjis And ((b >= BYTE_LEAD1_FIRST And b <= BYTE_LEAD1_LAST) Or (b >= BYTE_LEAD2_FIRST And b <= BYTE_LEAD2_LAST)) Then
1190                  ch = ChrW(CODE_WIDE_MARK)
1200                  trail = True
1210                  allSpace = False
                  ElseIf b = BYTE_SPACE Then
1220                  ch = MARK_SPACE
                  ElseIf b >= BYTE_DIGIT_0 And b <= BYTE_DIGIT_9 Then
1230                  ch = MARK_DIGIT
1240                  allSpace = False
                  ElseIf b >= BYTE_UPPER_A And b <= BYTE_UPPER_Z Then
1250                  ch = MARK_UPPER
1260                  allSpace = False
                  ElseIf b >= BYTE_LOWER_A And b <= BYTE_LOWER_Z Then
1270                  ch = MARK_LOWER
1280                  allSpace = False
                  ElseIf isSjis And b >= BYTE_KANA_FIRST And b <= BYTE_KANA_LAST Then
1290                  ch = ChrW(CODE_HALFKANA_MARK)
1300                  allSpace = False
                  ElseIf b >= BYTE_SYMBOL_FIRST And b <= BYTE_SYMBOL_LAST Then
1310                  ch = Chr$(b)
1320                  allSpace = False
                  ElseIf b = BYTE_NUL Then
1330                  ch = MARK_NUL
1340                  allSpace = False
                  ElseIf Not isSjis And b >= BYTE_HIGH Then
1350                  ch = ChrW(CODE_WIDE_MARK)
1360                  allSpace = False
                  Else
1370                  ch = MARK_OTHER
1380                  allSpace = False
                  End If
1390              mask = mask & ch
              Next j

1400          If allSpace Then
1410              emptyN = emptyN + 1
              Else
1420              If Left$(mask, 1) = MARK_SPACE Then leadSpaceN = leadSpaceN + 1
1430              If Right$(mask, 1) = MARK_SPACE Then trailSpaceN = trailSpaceN + 1
1440              If Not Replace(mask, MARK_DIGIT, "") Like "*[!_]*" And InStr(mask, MARK_DIGIT) > 0 Then
1450                  digitN = digitN + 1
                      ' 8 桁の数字が yyyymmdd として正しい日付か
1460                  If mask = String$(DATE8_LEN, MARK_DIGIT) Then
1470                      y = 0
1480                      For j = 0 To 3
1490                          y = y * DECIMAL_BASE + data(lineStart(idx(i)) + segStart(sg) + j) - BYTE_DIGIT_0
                          Next j
1500                      m = (data(lineStart(idx(i)) + segStart(sg) + 4) - BYTE_DIGIT_0) * DECIMAL_BASE + _
                              data(lineStart(idx(i)) + segStart(sg) + 5) - BYTE_DIGIT_0
1510                      d = (data(lineStart(idx(i)) + segStart(sg) + 6) - BYTE_DIGIT_0) * DECIMAL_BASE + _
                              data(lineStart(idx(i)) + segStart(sg) + 7) - BYTE_DIGIT_0
1520                      If y >= YEAR_MIN And y <= YEAR_MAX And m >= 1 And m <= MONTH_MAX And d >= 1 And d <= DAY_MAX Then date8N = date8N + 1
                      End If
1530                  If data(lineStart(idx(i)) + segStart(sg)) = BYTE_DIGIT_0 And segEnd(sg) > segStart(sg) Then zeroN = zeroN + 1
                  End If
              End If

              ' 連続する記号を個数表記にまとめる
1540          pat = ""
1550          prevCh = ""
1560          runLen = 0
1570          For p = 1 To Len(mask) + 1
1580              If p <= Len(mask) Then ch = Mid$(mask, p, 1) Else ch = vbNullChar
1590              If ch = prevCh Then
1600                  runLen = runLen + 1
                  Else
1610                  If runLen >= RUN_COMPRESS_MIN Then
1620                      pat = pat & prevCh & "(" & runLen & ")"
                      ElseIf runLen > 0 Then
1630                      pat = pat & String$(runLen, prevCh)
                      End If
1640                  prevCh = ch
1650                  runLen = 1
                  End If
              Next p
1660          If patDict.Exists(pat) Then
1670              patDict(pat) = patDict(pat) + 1
              ElseIf patDict.Count < PATTERN_CAP Then
1680              patDict.Add pat, 1
              End If
          Next i

1690      Select Case dom(segStart(sg))
              Case CLS_SPACE: typeText = TYPE_SPACE
              Case CLS_DIGIT: typeText = TYPE_DIGIT
              Case CLS_ALPHA: typeText = TYPE_ALPHA
              Case CLS_WIDE: typeText = TYPE_WIDE
              Case CLS_KANA: typeText = TYPE_KANA
              Case CLS_SYMBOL: typeText = TYPE_SYMBOL
              Case CLS_NUL: typeText = TYPE_NUL
              Case Else: typeText = TYPE_OTHER
          End Select
1700      If segMerged(sg) Then typeText = TYPE_TEXT_VAR

          ' 上位のパターン
1710      patText = ""
1720      Set usedPat = CreateObject("Scripting.Dictionary")
1730      For shown = 1 To PATTERN_SHOW
1740          bestKey = ""
1750          bestPat = 0
1760          For Each k In patDict.Keys
1770              If Not usedPat.Exists(k) And patDict(k) > bestPat Then
1780                  bestPat = patDict(k)
1790                  bestKey = k
                  End If
              Next k
1800          If bestPat = 0 Then Exit For
1810          usedPat.Add bestKey, True
1820          If Len(bestKey) > PATTERN_SHOW_LEN Then bestKey = Left$(bestKey, PATTERN_SHOW_LEN) & "..."
1830          patText = patText & IIf(Len(patText) > 0, ", ", "") & "`" & bestKey & "`" & IIf(patDict.Count > 1, " (" & bestPat & ")", "")
          Next shown
1840      If patDict.Count > PATTERN_SHOW Then patText = patText & ", ..."

1850      rmk = ""
1860      If dom(segStart(sg)) = CLS_DIGIT Or dom(segStart(sg)) = CLS_ALPHA Or dom(segStart(sg)) = CLS_WIDE _
             Or dom(segStart(sg)) = CLS_KANA Or segMerged(sg) Then
1870          If leadSpaceN > (sample - emptyN) / 2 Then rmk = rmk & IIf(Len(rmk) > 0, RMK_SEP, "") & RMK_RIGHT
1880          If trailSpaceN > (sample - emptyN) / 2 Then rmk = rmk & IIf(Len(rmk) > 0, RMK_SEP, "") & RMK_LEFT
          End If
1890      If dom(segStart(sg)) = CLS_DIGIT And zeroN > digitN / 2 And digitN > 0 Then rmk = rmk & IIf(Len(rmk) > 0, RMK_SEP, "") & RMK_ZERO
1900      If date8N > 0 And date8N = digitN Then rmk = rmk & IIf(Len(rmk) > 0, RMK_SEP, "") & RMK_DATE8
1910      If dom(segStart(sg)) = CLS_SPACE Then rmk = rmk & IIf(Len(rmk) > 0, RMK_SEP, "") & RMK_BLANK

1920      pos = segStart(sg) + 1
1930      rep = rep & "| " & sg & " | " & pos & "-" & (segEnd(sg) + 1) & " | " & (segEnd(sg) - segStart(sg) + 1) & " | " & _
                typeText & " | " & emptyN & " | " & Replace(patText, "|", "/") & " | " & rmk & " |" & vbCrLf
      Next sg
1940  ProfilePositions = rep

      LogSuccess MODULE_NAME, "ProfilePositions", "Profiled " & sample & " lines; segments=" & segCount & "; maxLen=" & maxLen
1950  Exit Function
ErrHandler:
      LogError MODULE_NAME, "ProfilePositions", Err.Number, Err.Description, Erl, _
               "cnt=" & cnt & "; i=" & i & "; j=" & j & "; sg=" & sg
      itemCount = segCount
      ProfilePositions = rep & RPT_ERROR_LINE & vbCrLf
End Function
