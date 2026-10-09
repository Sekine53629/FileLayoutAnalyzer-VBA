Attribute VB_Name = "RelationFinder"
Option Explicit

'======================================================================
' RelationFinder モジュール
' 数値の列どうしの関係を探す。項目名が無くても、列の役割を推定する手がかりにする。
'   - 隠れた割合: 列 A / 列 B が少数の区切りの良い値 (0.7 / 0.8 / 0.9 など) に集まる
'     (点数と円の違いを考えて x10 も試す。端数処理による誤差は許す)
'   - 割合を決めている列: その割合が、あるコード列の値ごとに 1 つに決まる
'     (保険種別・高齢者区分などの、明示されていない区分の列を見つける)
'   - 同じ値: 列 A = 列 B (x10 と 0.1 の組み合わせのような、見かけの割合と取り違えないため先に調べる)
'   - 合計: 列 C = 列 A + 列 B
'   - 積: 列 C = 列 A x 列 B (数量 x 単価 = 金額 など)
' 値そのものは報告に出さず、関係と一致した行の割合だけを出す。
'
' 依存: ErrorLogger.bas
'======================================================================

Private Const MODULE_NAME As String = "RelationFinder"

' ---- 閾値 ----
Private Const MAX_ROWS As Long = 2000              ' 関係を調べる最大の行数 (先頭から)
Private Const MAX_TRIPLE_ROWS As Long = 500         ' 3 列の関係 (合計・積) を調べる最大の行数
Private Const MAX_NUM_COLS As Long = 40            ' 調べる数値の列の最大数
Private Const MAX_TRIPLE_COLS As Long = 20         ' 3 列の関係を調べる数値の列の最大数
Private Const MIN_ROWS As Long = 10                ' 関係を判定するのに必要な最小の行数
Private Const NUMERIC_SHARE As Double = 0.8        ' 数値の列とみなす、数値の値の割合
Private Const LEAD_ZERO_SHARE As Double = 0.5      ' 先頭 0 の値がこれ以上ならコード・番号とみなして除く
Private Const MATCH_SHARE As Double = 0.9          ' 関係が成り立つ行の割合 (これ以上で報告する)
Private Const RATIO_STEP As Double = 0.05          ' 割合の刻み
Private Const RATIO_STEPS As Long = 20             ' 0.05 - 1.00 の刻みの数
Private Const RATIO_MAX_LEVELS As Long = 4         ' 割合の種類がこれ以下なら「少数の値に集まる」とみなす
Private Const LEVEL_MIN_SHARE As Double = 0.02     ' 割合の 1 つの値として数える最小の割合
Private Const REL_TOL As Double = 0.005            ' 端数処理を許す相対誤差
Private Const ABS_TOL As Double = 1                ' 割合で、端数処理を許す絶対誤差 (円未満の四捨五入)
Private Const MIN_BASE As Double = 100             ' 割合を調べる分母の最小値 (小さな値どうしの偶然の一致を避ける)
Private Const SUM_TOL As Double = 0.000001         ' 合計の許容誤差 (小数の計算誤差だけ)
Private Const PRODUCT_ABS_TOL As Double = 0.5      ' 積の許容誤差の下限 (整数どうしは完全一致)
Private Const POINT_SCALE As Double = 10           ' 点数 -> 円
Private Const CODE_MAX_KINDS As Long = 20          ' 区分の列とみなす最大の種類数
Private Const DETERMINE_SHARE As Double = 0.95     ' 割合が区分の列で「決まる」とみなす割合
Private Const PARTIAL_SHARE As Double = 0.7        ' 「ほぼ決まる」とみなす割合
Private Const NONZERO_SHARE As Double = 0.3        ' 合計・積で、両方の項が 0 でない行がこの割合以上
Private Const MAX_REPORT As Long = 15              ' 報告する関係の最大数 (種類ごと)
Private Const PCT As Long = 100

' ---- 報告の文言 ----
Private Const RPT_SEC As String = "[数値の列どうしの関係 (値は載せず、関係と一致した行の割合だけを載せる)]"
Private Const RPT_NONE As String = "  見つからない"
Private Const RPT_RATIO_HEAD As String = "  隠れた割合:"
Private Const RPT_SUM_HEAD As String = "  合計 (C = A + B):"
Private Const RPT_EQUAL_HEAD As String = "  同じ値 (A = B):"
Private Const RPT_PRODUCT_HEAD As String = "  積 (C = A x B):"
Private Const RPT_ITEM_INDENT As String = "    "
Private Const RPT_DIV As String = " / "
Private Const RPT_SCALE As String = " x10"
Private Const RPT_EQ As String = " = "
Private Const RPT_PLUS As String = " + "
Private Const RPT_TIMES As String = " x "
Private Const RPT_MATCH As String = " (一致 "
Private Const RPT_ROWS As String = " 行)"
Private Const RPT_OF As String = "% / "
Private Const RPT_LEVEL_SEP As String = ", "
Private Const RPT_DETERMINED As String = "      -> 割合は "
Private Const RPT_DETERMINED_2 As String = " の値で決まる (区分 "
Private Const RPT_DETERMINED_3 As String = " 種類 -> 割合 "
Private Const RPT_DETERMINED_4 As String = " 種類。一致 "
Private Const RPT_PARTIAL_2 As String = " の値でほぼ決まる (一致 "
Private Const RPT_NOT_DETERMINED As String = "      -> 割合を決める区分の列は見当たらない (別のファイルや計算で決まっている可能性)"
Private Const RPT_SCALE_NOTE As String = "  注: x10 は、分母を 10 倍すると成り立つこと (点数 -> 円 の可能性) を表す。"
Private Const RPT_RATIO_NOTE As String = "  注: 割合が 0.7 / 0.8 / 0.9 に分かれる場合は、医療保険の給付割合 (患者負担 3 割 / 2 割 / 1 割) の可能性がある。0.3 / 0.2 / 0.1 なら負担割合の側。"
Private Const RPT_TRUNC As String = "  注: 先頭 "
Private Const RPT_TRUNC_2 As String = " 行だけで調べている"
Private Const RPT_PCT As String = "%"
Private Const NAME_PREFIX As String = "No."
Private Const LEAD_ZERO As String = "0"
Private Const DECIMAL_MARK As String = "."
Private Const NUM_PATTERN As String = "*[!0-9.-]*"


'----------------------------------------------------------------------
' FindRelations: 行の集まりについて、数値の列どうしの関係を探し、報告の 1 欄を返す
'   grid      値 (1 始まりの 2 次元配列)
'   rows()    調べる行 (grid の行番号。1 始まり)
'   rowN      行の数
'   colCount  調べる列の数 (1 - colCount)
'   names()   列の表示名 (項目一覧のラベル。1 始まり)
'----------------------------------------------------------------------
Public Function FindRelations(ByRef grid As Variant, ByRef rows() As Long, ByVal rowN As Long, _
                                ByVal colCount As Long, ByRef names() As String) As String
      On Error GoTo ErrHandler
      Dim n As Long
      Dim i As Long
      Dim c As Long
      Dim v As Variant
      Dim s As String
      Dim numVal() As Double
      Dim numOk() As Boolean
      Dim filled As Long
      Dim numCount As Long
      Dim leadZero As Long
      Dim numCols() As Long
      Dim numColN As Long
      Dim codeCols() As Long
      Dim codeColN As Long
      Dim kinds As Object
      Dim a As Long
      Dim b As Long
      Dim cc As Long
      Dim ia As Long
      Dim ib As Long
      Dim ic As Long
      Dim r As Double
      Dim sm As Long
      Dim matchM(0 To 1) As Long
      Dim levelCountM() As Long
      Dim rowLevelM() As Long
      Dim k As Long
      Dim levelCount() As Long
      Dim pairRows As Long
      Dim matchN As Long
      Dim levels As Long
      Dim levelText As String
      Dim rowLevel() As Long
      Dim bestShare As Double
      Dim bestCol As Long
      Dim bestKinds As Long
      Dim share As Double
      Dim codeMap As Object
      Dim key As String
      Dim kk As Variant
      Dim mapTotal As Long
      Dim mapHit As Long
      Dim ratioText As String
      Dim ratioN As Long
      Dim sumText As String
      Dim sumN As Long
      Dim prodText As String
      Dim prodN As Long
      Dim tripleRows As Long
      Dim bothNonZero As Long
      Dim bs As Double
      Dim expect As Double
      Dim usedScale As Boolean
      Dim rep As String
      Dim anyScale As Boolean
      Dim tolV As Double
      Dim eqN As Long
      Dim eqRows As Long
      Dim equalText As String
      Dim equalN As Long
      Dim isEqual As Boolean

10    n = rowN
20    If n > MAX_ROWS Then n = MAX_ROWS
30    If n < MIN_ROWS Or colCount < 2 Then
40        FindRelations = ""
          LogSuccess MODULE_NAME, "FindRelations", "Skipped: rows=" & rowN & "; cols=" & colCount
50        Exit Function
      End If

      ' ---- 列ごとに数値へ変換し、数値の列と区分の列を選ぶ ----
60    ReDim numVal(1 To n, 1 To colCount)
70    ReDim numOk(1 To n, 1 To colCount)
80    ReDim numCols(1 To colCount)
90    ReDim codeCols(1 To colCount)
100   For c = 1 To colCount
110       filled = 0
120       numCount = 0
130       leadZero = 0
140       Set kinds = CreateObject("Scripting.Dictionary")
150       For i = 1 To n
160           v = grid(rows(i), c)
170           If Not IsError(v) And Not IsEmpty(v) And Not IsNull(v) Then
180               s = Trim$(CStr(v))
190               If Len(s) > 0 Then
200                   filled = filled + 1
210                   If kinds.Count <= CODE_MAX_KINDS Then kinds(s) = True
220                   If VarType(v) <> vbString And IsNumeric(v) And VarType(v) <> vbBoolean Then
230                       numVal(i, c) = CDbl(v)
240                       numOk(i, c) = True
250                       numCount = numCount + 1
                      ElseIf Not s Like NUM_PATTERN And IsNumeric(s) Then
260                       numVal(i, c) = CDbl(s)
270                       numOk(i, c) = True
280                       numCount = numCount + 1
290                       If Left$(s, 1) = LEAD_ZERO And Len(s) > 1 And Mid$(s, 2, 1) <> DECIMAL_MARK Then leadZero = leadZero + 1
                      End If
                  End If
              End If
          Next i
300       If filled > 0 Then
              ' 数値の列: 数値が多く、先頭 0 のコードではなく、2 種類以上の値がある
310           If numCount >= filled * NUMERIC_SHARE And leadZero < numCount * LEAD_ZERO_SHARE _
                 And kinds.Count > 1 And numColN < MAX_NUM_COLS Then
320               numColN = numColN + 1
330               numCols(numColN) = c
              End If
              ' 区分の列: 種類が少なく、ほぼ埋まっている (数値でも文字でもよい)
340           If kinds.Count > 1 And kinds.Count <= CODE_MAX_KINDS And filled >= n * MATCH_SHARE Then
350               codeColN = codeColN + 1
360               codeCols(codeColN) = c
              End If
          End If
      Next c

      ' ---- 隠れた割合: A / B (または A / (B x10)) が 0.05 刻みの少数の値に集まる ----
      ' x10 (点数 -> 円) を使うかは、行ごとではなく列の組ごとに決める (一致する行の多いほう)。
      ' 行ごとに決めると、0.1 x10 = 1.00 のような偶然の一致が混じるため
370   ReDim rowLevel(1 To n)
380   ReDim rowLevelM(0 To 1, 1 To n)
390   For ia = 1 To numColN
400       For ib = 1 To numColN
410           If ia <> ib And ratioN < MAX_REPORT Then
420               a = numCols(ia)
430               b = numCols(ib)
440               ReDim levelCountM(0 To 1, 0 To RATIO_STEPS)
450               matchM(0) = 0
460               matchM(1) = 0
470               pairRows = 0
480               For i = 1 To n
490                   rowLevelM(0, i) = -1
500                   rowLevelM(1, i) = -1
510                   If numOk(i, a) And numOk(i, b) Then
520                       If numVal(i, a) > 0 And numVal(i, b) >= MIN_BASE Then
530                           pairRows = pairRows + 1
540                           For sm = 0 To 1
550                               bs = numVal(i, b) * IIf(sm = 1, POINT_SCALE, 1)
560                               r = numVal(i, a) / bs
570                               If r <= 1 Then
580                                   k = CLng(r / RATIO_STEP)
590                                   If k >= 1 And k <= RATIO_STEPS Then
600                                       tolV = bs * REL_TOL
610                                       If tolV < ABS_TOL Then tolV = ABS_TOL
620                                       If Abs(numVal(i, a) - k * RATIO_STEP * bs) <= tolV Then
630                                           matchM(sm) = matchM(sm) + 1
640                                           levelCountM(sm, k) = levelCountM(sm, k) + 1
650                                           rowLevelM(sm, i) = k
                                          End If
                                      End If
                                  End If
                              Next sm
                          End If
                      End If
                  Next i
                  ' 同じ値の組 (A = B) は割合ではなく「同じ値」として扱う (A, B の順に 1 回だけ報告する)
660               eqN = 0
670               eqRows = 0
680               For i = 1 To n
690                   If numOk(i, a) And numOk(i, b) Then
700                       eqRows = eqRows + 1
710                       If Abs(numVal(i, a) - numVal(i, b)) <= SUM_TOL Then eqN = eqN + 1
                      End If
                  Next i
720               isEqual = (eqRows >= MIN_ROWS And eqN >= eqRows * PARTIAL_SHARE)
730               If isEqual And ia < ib And equalN < MAX_REPORT Then
740                   equalN = equalN + 1
750                   equalText = equalText & RPT_ITEM_INDENT & names(a) & RPT_EQ & names(b) & _
                                  RPT_MATCH & Format(eqN / eqRows * PCT, "0") & RPT_OF & eqRows & RPT_ROWS & vbCrLf
                  End If
760               If isEqual Then pairRows = 0

                  ' 一致の多いほうの読み方を採る
770               sm = IIf(matchM(1) > matchM(0), 1, 0)
780               usedScale = (sm = 1)
790               matchN = matchM(sm)
800               ReDim levelCount(0 To RATIO_STEPS)
810               For k = 0 To RATIO_STEPS
820                   levelCount(k) = levelCountM(sm, k)
                  Next k
830               For i = 1 To n
840                   rowLevel(i) = rowLevelM(sm, i)
                  Next i
850               If pairRows >= MIN_ROWS And matchN >= pairRows * MATCH_SHARE Then
860                   levels = 0
870                   levelText = ""
880                   For k = 1 To RATIO_STEPS
890                       If levelCount(k) >= matchN * LEVEL_MIN_SHARE And levelCount(k) > 0 Then
900                           levels = levels + 1
910                           levelText = levelText & IIf(Len(levelText) > 0, RPT_LEVEL_SEP, "") & _
                                          Format(k * RATIO_STEP, "0.00") & " (" & Format(levelCount(k) / matchN * PCT, "0") & RPT_PCT & ")"
                          End If
                      Next k
                      ' 1.00 だけ (同じ値) は割合として扱わない
920                   If levels >= 1 And levels <= RATIO_MAX_LEVELS And levelCount(RATIO_STEPS) < matchN * MATCH_SHARE Then
930                       If usedScale Then anyScale = True
940                       ratioN = ratioN + 1
950                       ratioText = ratioText & RPT_ITEM_INDENT & names(a) & RPT_DIV & names(b) & IIf(usedScale, RPT_SCALE, "") & _
                                      RPT_EQ & levelText & RPT_MATCH & Format(matchN / pairRows * PCT, "0") & RPT_OF & pairRows & RPT_ROWS & vbCrLf

                          ' 割合が複数あれば、それを決めている区分の列を探す
960                       If levels >= 2 Then
970                           bestShare = 0
980                           bestCol = 0
990                           For cc = 1 To codeColN
1000                              If codeCols(cc) <> a And codeCols(cc) <> b Then
1010                                  Set codeMap = CreateObject("Scripting.Dictionary")
1020                                  mapTotal = 0
1030                                  For i = 1 To n
1040                                      If rowLevel(i) > 0 Then
1050                                          key = Trim$(CStr(grid(rows(i), codeCols(cc)))) & vbTab & rowLevel(i)
1060                                          codeMap(key) = codeMap(key) + 1
1070                                          mapTotal = mapTotal + 1
                                          End If
                                      Next i
                                      ' 区分の値ごとに、いちばん多い割合の行数を足す
1080                                  Set kinds = CreateObject("Scripting.Dictionary")
1090                                  For Each kk In codeMap.Keys
1100                                      s = Left$(kk, InStr(kk, vbTab) - 1)
1110                                      If codeMap(kk) > kinds(s) Then kinds(s) = codeMap(kk)
                                      Next kk
1120                                  mapHit = 0
1130                                  For Each kk In kinds.Keys
1140                                      mapHit = mapHit + kinds(kk)
                                      Next kk
1150                                  If mapTotal > 0 Then
1160                                      share = mapHit / mapTotal
1170                                      If share > bestShare Or (share = bestShare And kinds.Count < bestKinds) Then
1180                                          bestShare = share
1190                                          bestCol = codeCols(cc)
1200                                          bestKinds = kinds.Count
                                          End If
                                      End If
                                  End If
                              Next cc
1210                          If bestCol > 0 And bestShare >= DETERMINE_SHARE Then
1220                              ratioText = ratioText & RPT_DETERMINED & names(bestCol) & RPT_DETERMINED_2 & bestKinds & _
                                              RPT_DETERMINED_3 & levels & RPT_DETERMINED_4 & Format(bestShare * PCT, "0") & RPT_PCT & ")" & vbCrLf
                              ElseIf bestCol > 0 And bestShare >= PARTIAL_SHARE Then
1230                              ratioText = ratioText & RPT_DETERMINED & names(bestCol) & RPT_PARTIAL_2 & _
                                              Format(bestShare * PCT, "0") & RPT_PCT & ")" & vbCrLf
                              Else
1240                              ratioText = ratioText & RPT_NOT_DETERMINED & vbCrLf
                              End If
                          End If
                      End If
                  End If
              End If
          Next ib
      Next ia

      ' ---- 合計 (C = A + B) と積 (C = A x B) ----
1250  tripleRows = n
1260  If tripleRows > MAX_TRIPLE_ROWS Then tripleRows = MAX_TRIPLE_ROWS
1270  If numColN >= 3 And numColN <= MAX_TRIPLE_COLS Then
1280      For ia = 1 To numColN - 1
1290          For ib = ia + 1 To numColN
1300              For ic = 1 To numColN
1310                  If ic <> ia And ic <> ib Then
1320                      a = numCols(ia)
1330                      b = numCols(ib)
1340                      cc = numCols(ic)
                          ' 合計
1350                      If sumN < MAX_REPORT Then
1360                          pairRows = 0
1370                          matchN = 0
1380                          bothNonZero = 0
1390                          For i = 1 To tripleRows
1400                              If numOk(i, a) And numOk(i, b) And numOk(i, cc) Then
1410                                  pairRows = pairRows + 1
1420                                  If Abs(numVal(i, a) + numVal(i, b) - numVal(i, cc)) <= SUM_TOL Then
1430                                      matchN = matchN + 1
1440                                      If numVal(i, a) <> 0 And numVal(i, b) <> 0 Then bothNonZero = bothNonZero + 1
                                      End If
                                  End If
                              Next i
1450                          If pairRows >= MIN_ROWS And matchN >= pairRows * MATCH_SHARE And bothNonZero >= matchN * NONZERO_SHARE Then
1460                              sumN = sumN + 1
1470                              sumText = sumText & RPT_ITEM_INDENT & names(cc) & RPT_EQ & names(a) & RPT_PLUS & names(b) & _
                                            RPT_MATCH & Format(matchN / pairRows * PCT, "0") & RPT_OF & pairRows & RPT_ROWS & vbCrLf
                              End If
                          End If
                          ' 積
1480                      If prodN < MAX_REPORT Then
1490                          pairRows = 0
1500                          matchN = 0
1510                          bothNonZero = 0
1520                          For i = 1 To tripleRows
1530                              If numOk(i, a) And numOk(i, b) And numOk(i, cc) Then
1540                                  pairRows = pairRows + 1
1550                                  expect = numVal(i, a) * numVal(i, b)
1560                                  tolV = Abs(expect) * REL_TOL
1570                                  If tolV < PRODUCT_ABS_TOL Then tolV = PRODUCT_ABS_TOL
1580                                  If Abs(expect - numVal(i, cc)) <= tolV Then
1590                                      matchN = matchN + 1
                                          ' 1 を掛けているだけの行 (C = A) は数えない
1600                                      If numVal(i, a) <> 0 And numVal(i, b) <> 0 And numVal(i, a) <> 1 And numVal(i, b) <> 1 Then
1610                                          bothNonZero = bothNonZero + 1
                                          End If
                                      End If
                                  End If
                              Next i
1620                          If pairRows >= MIN_ROWS And matchN >= pairRows * MATCH_SHARE And bothNonZero >= matchN * NONZERO_SHARE Then
1630                              prodN = prodN + 1
1640                              prodText = prodText & RPT_ITEM_INDENT & names(cc) & RPT_EQ & names(a) & RPT_TIMES & names(b) & _
                                             RPT_MATCH & Format(matchN / pairRows * PCT, "0") & RPT_OF & pairRows & RPT_ROWS & vbCrLf
                              End If
                          End If
                      End If
                  Next ic
              Next ib
          Next ia
      End If

      ' ---- 報告 ----
1650  rep = RPT_SEC & vbCrLf
1660  If rowN > n Then rep = rep & RPT_TRUNC & n & RPT_TRUNC_2 & vbCrLf
1670  If ratioN + sumN + prodN + equalN = 0 Then
1680      rep = rep & RPT_NONE & vbCrLf
      Else
1690      If ratioN > 0 Then rep = rep & RPT_RATIO_HEAD & vbCrLf & ratioText
1700      If equalN > 0 Then rep = rep & RPT_EQUAL_HEAD & vbCrLf & equalText
1710      If sumN > 0 Then rep = rep & RPT_SUM_HEAD & vbCrLf & sumText
1720      If prodN > 0 Then rep = rep & RPT_PRODUCT_HEAD & vbCrLf & prodText
1730      If anyScale Then rep = rep & RPT_SCALE_NOTE & vbCrLf
1740      If ratioN > 0 Then rep = rep & RPT_RATIO_NOTE & vbCrLf
      End If
1750  FindRelations = rep & vbCrLf

      LogSuccess MODULE_NAME, "FindRelations", "Checked " & n & " rows; numericCols=" & numColN & "; codeCols=" & codeColN & _
                 "; ratios=" & ratioN & "; equals=" & equalN & "; sums=" & sumN & "; products=" & prodN
1760  Exit Function
ErrHandler:
      LogError MODULE_NAME, "FindRelations", Err.Number, Err.Description, Erl, _
               "rowN=" & rowN & "; colCount=" & colCount & "; a=" & a & "; b=" & b & "; i=" & i
      FindRelations = ""
End Function
