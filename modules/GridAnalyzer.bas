Attribute VB_Name = "GridAnalyzer"
Option Explicit

'======================================================================
' GridAnalyzer モジュール
' 2 次元配列 (シートの値、区切り文字形式を展開したもの) の構造を解析する。
'   - レコード種別ごとに項目数が違う形式 (先頭項目が種別コード) の判定
'   - 表題・見出し行 (複数行の見出しを含む)・データ行の判定
'   - 見出しの繰り返し、別の見出しらしき行、項目数の違う行、疎な行の検出
'   - 項目ごとの型、桁数、空の数、種類数、書式パターン、備考
'
' 値そのものは報告に出さず、書式パターンに置き換える。
'   9=数字  A=英大文字  a=英小文字  N(全角)=全角文字  ｶ(半角)=半角カナ  _=空白
'   同じ記号が 5 つ以上続くときは 9(13) のように個数で表す。
'
' 依存: ErrorLogger.bas
'======================================================================

Private Const MODULE_NAME As String = "GridAnalyzer"

' ---- 判定の閾値 ----
Private Const HEADER_SEARCH_ROWS As Long = 30      ' 見出し行を探す、先頭からの空でない行数
Private Const HEADER_MIN_FILL As Double = 0.6      ' 見出し行は、主な項目数のこの割合以上が埋まっている
Private Const HEADER_MIN_CELLS As Long = 2         ' 見出し行の最小のセル数
Private Const HEADER_UNIQUE_RATIO As Double = 0.8  ' 見出し行は、値がこの割合以上互いに異なる
Private Const HEADER_CONTRAST_ROWS As Long = 10    ' 見出し候補の下で、数字を含む行を探す行数
Private Const MAX_HEADER_LINES As Long = 3         ' 複数行の見出しの最大行数
Private Const LABEL_MAX_LEN As Long = 40           ' これより長い値を含む行は見出しとみなさない
Private Const SPARSE_RATIO As Double = 0.5         ' 項目数のこの割合未満しか埋まっていない行を疎な行とする
Private Const RECTYPE_MAX_KINDS As Long = 30       ' レコード種別とみなす先頭項目の最大の種類数
Private Const RECTYPE_MAX_KEYLEN As Long = 4       ' レコード種別コードの最大の文字数
Private Const RECTYPE_WIDTH_SHARE As Double = 0.9  ' 項目数が揃っているとみなす割合 (これ未満ならレコード種別を疑う)
Private Const MAX_TABLES As Long = 50               ' 1 つのシート・ファイルで分ける表の最大数
Private Const RECTYPE_MAX_SINGLES As Long = 2      ' 1 行しかない種別の許容数 (先頭・末尾のレコード相当)
Private Const RECTYPE_SEQ_SHOW As Long = 30        ' 先頭から何行分のレコード種別の並びを載せるか
Private Const MAX_LIST_ITEMS As Long = 10          ' 行番号を列挙する上限
Private Const TITLE_MAX_LEN As Long = 60           ' 表題として載せる文字数の上限
Private Const DISTINCT_CAP As Long = 1000          ' 種類数を数える上限
Private Const PATTERN_CAP As Long = 200            ' 書式パターンを数える種類の上限
Private Const PATTERN_SHOW As Long = 3             ' 載せる書式パターンの数
Private Const PATTERN_SRC_MAX As Long = 60         ' パターン化する値の最大の文字数
Private Const PATTERN_SHOW_LEN As Long = 40        ' 載せる書式パターンの最大の文字数
Private Const RUN_COMPRESS_MIN As Long = 5         ' この数以上続く記号を個数表記にする (9999 は残す)
Private Const CODE_MAX_KINDS As Long = 20          ' コード値候補とみなす最大の種類数
Private Const CODE_MIN_ROWS As Long = 30           ' コード値候補の判定に必要な最小の行数
Private Const NAME_LIKE_SHARE As Double = 0.5      ' 氏名らしい形の値がこの割合以上なら備考に出す
Private Const MIX_MIN_SHARE As Double = 0.05       ' 2 番目の型がこの割合以上あれば混在とする
Private Const WIDE_BYTES As Long = 2               ' 全角 1 文字の Shift_JIS でのバイト数
Private Const YEAR_MIN As Long = 1900
Private Const YEAR_MAX As Long = 2100
Private Const MONTH_MAX As Long = 12
Private Const DAY_MAX As Long = 31
Private Const ALPHABET_SIZE As Long = 26
Private Const WAREKI_YM_LEN As Long = 5            ' 和暦年月 GYYMM の桁数
Private Const WAREKI_YMD_LEN As Long = 7           ' 和暦年月日 GYYMMDD の桁数
Private Const WAREKI_ERA_MIN As Long = 1           ' 元号の区分 (1=明治 ... 5=令和)
Private Const WAREKI_ERA_MAX As Long = 5

' ---- 文字コード ----
Private Const CODE_DIGIT_0 As Long = 48
Private Const CODE_DIGIT_9 As Long = 57
Private Const CODE_UPPER_A As Long = 65
Private Const CODE_UPPER_Z As Long = 90
Private Const CODE_LOWER_A As Long = 97
Private Const CODE_LOWER_Z As Long = 122
Private Const CODE_SPACE As Long = 32
Private Const CODE_TAB As Long = 9
Private Const CODE_LF As Long = 10
Private Const CODE_CR As Long = 13
Private Const CODE_ASCII_FIRST As Long = 33
Private Const CODE_ASCII_LAST As Long = 126
Private Const CODE_HALFKANA_FIRST As Long = &HFF61&
Private Const CODE_HALFKANA_LAST As Long = &HFF9F&
Private Const CODE_WIDE_DIGIT_0 As Long = &HFF10&
Private Const CODE_WIDE_DIGIT_9 As Long = &HFF19&
Private Const CODE_WIDE_SPACE As Long = &H3000
Private Const CODE_WIDE_MARK As Long = &HFF2E&      ' 全角の N (全角文字の記号)
Private Const CODE_HALFKANA_MARK As Long = &HFF76&  ' 半角の カ (半角カナの記号)
Private Const UNSIGNED_OFFSET As Long = 65536

' ---- パターンの記号 ----
Private Const MARK_DIGIT As String = "9"
Private Const MARK_UPPER As String = "A"
Private Const MARK_LOWER As String = "a"
Private Const MARK_SPACE As String = "_"
Private Const MARK_NEWLINE As String = "~"
Private Const MARK_CTRL As String = "?"
Private Const KEEP_CHARS As String = "年月日時分秒"   ' 日付の書式として残す全角文字

' ---- 型の名前 ----
Private Const CLS_DIGITS As String = "数字"
Private Const CLS_SIGNED As String = "整数 (符号付き)"
Private Const CLS_DECIMAL As String = "小数"
Private Const CLS_NUM_COMMA As String = "数値 (桁区切り)"
Private Const CLS_NUMBER_INT As String = "数値 (整数)"
Private Const CLS_NUMBER_DEC As String = "数値 (小数)"
Private Const CLS_DATE As String = "日付"
Private Const CLS_DATE8 As String = "日付 (yyyymmdd)"
Private Const CLS_WAREKI As String = "和暦日付"
Private Const CLS_TIME As String = "時刻"
Private Const CLS_DATETIME As String = "日時"
Private Const CLS_BOOL As String = "真偽値"
Private Const CLS_TEXT_WIDE As String = "文字 (全角を含む)"
Private Const CLS_TEXT_KANA As String = "文字 (半角カナを含む)"
Private Const CLS_TEXT_ALNUM As String = "文字 (半角英数記号)"
Private Const CLS_TEXT_SYMBOL As String = "記号のみ"
Private Const CLS_EMPTY As String = "空"
Private Const CLS_MIXED As String = "混在"

' ---- 個人情報の可能性がある見出しの語 ----
Private Const PII_WORDS As String = "氏名|名前|姓名|フリガナ|ふりがな|カナ氏名|患者|受診者|住所|所在地|電話|TEL|携帯|生年月日|誕生日|年齢|性別|被保険者|保険者番号|記号番号|証番号|受給者|メール|MAIL|郵便|口座|カード番号|マイナンバー|個人番号"
' ---- 小計・合計の行を見分ける語 ----
Private Const SUMMARY_WORDS As String = "合計|小計|総計|計|件数|TOTAL"
Private Const LIST_SEP As String = "|"

' ---- 報告の文言 ----
Private Const RPT_TABLE_HEAD_GRID As String = "| No | 列 | ラベル | 種別 | 桁数 | 空 | 種類 | 書式パターン (上位3) | 備考 |"
Private Const RPT_TABLE_RULE As String = "|---:|---|---|---|---|---:|---:|---|---|"
Private Const RPT_NO_LABEL As String = "(見出しなし)"
Private Const RPT_ITEM As String = "(項目"
Private Const RPT_ITEM_CLOSE As String = ")"
Private Const RPT_MODE_TABLE As String = "表形式"
Private Const RPT_MODE_RECTYPE As String = "レコード種別ごとに項目数が異なる形式 (先頭項目が種別コード)"
Private Const RPT_SEC_STRUCT As String = "[構造]"
Private Const RPT_SEC_TITLE As String = "[表題 (見出し行より上の行)]"
Private Const RPT_SEC_ROWS As String = "[行の構成]"
Private Const RPT_SEC_ITEMS As String = "[項目一覧]"
Private Const RPT_SEC_RECTYPE As String = "[レコード種別]"
Private Const RPT_SEC_MERGED As String = "[結合セル (見出し行まで)]"
Private Const RPT_FORMAT As String = "  形式            : "
Private Const RPT_ROWS_TOTAL As String = "  行数            : "
Private Const RPT_ROWS_UNIT As String = " 行 (空行 "
Private Const RPT_ROWS_CLOSE As String = " 行)"
Private Const RPT_WIDTH_MODE As String = "  主な項目数      : "
Private Const RPT_WIDTH_SHARE As String = " (該当する行の割合 "
Private Const RPT_WIDTH_KINDS As String = "  項目数の種類    : "
Private Const RPT_HEADER_AT As String = "  見出し行        : 行 "
Private Const RPT_HEADER_NONE As String = "  見出し行        : 見つからない (先頭行からデータとみなす)"
Private Const RPT_HEADER_LINES As String = " 行構成)"
Private Const RPT_DATA_AT As String = "  データ行        : 行 "
Private Const RPT_DATA_COUNT As String = " から "
Private Const RPT_DATA_UNIT As String = " 行"
Private Const RPT_NO_TITLE As String = "  なし"
Private Const RPT_TITLE_ROW As String = "  行 "
Private Const RPT_TITLE_COLS As String = " (列 "
Private Const RPT_TITLE_NOTE As String = "  (数字は 9 に置き換えている)"
Private Const RPT_GAP As String = "  表題から見出し行までの空き行: "
Private Const RPT_REPEAT_HEADER As String = "  見出しと同じ行 (改ページなどによる見出しの繰り返し): "
Private Const RPT_ODD_WIDTH As String = "  項目数が主な項目数と違う行: "
Private Const RPT_SPARSE As String = "  疎な行 (小計、区切り、注記などの可能性): "
Private Const RPT_EDGE As String = "  先頭・末尾で項目数が違う行 (ヘッダ/トレーラ レコードの可能性。項目一覧から除外): "
Private Const RPT_SUMMARY_ROWS As String = "  小計・合計の語を含む行: "
Private Const RPT_BLOCKS As String = "  空行で区切られたまとまり: "
Private Const RPT_JUDGE As String = "  別票の混在      : "
Private Const RPT_JUDGE_NONE As String = "見当たらない"
Private Const RPT_JUDGE_REPEAT As String = "同じ見出しの繰り返しのみ (1 種類の表が分割されている)"
Private Const RPT_JUDGE_MIXED As String = "可能性あり (見出しや項目数の違う行がある)"
Private Const RPT_NONE As String = "なし"
Private Const RPT_ITEMS_UNIT As String = " 件"
Private Const RPT_MORE As String = " ほか"
Private Const RPT_RECTYPE_HEAD As String = "| 種別コード | 役割 | 行数 | 項目数 | 最初の行 |"
Private Const RPT_RECTYPE_RULE As String = "|---|---|---:|---|---:|"
Private Const RPT_RECTYPE_SEQ As String = "  先頭からの種別の並び: "
Private Const RPT_RECTYPE_SEC As String = "[レコード種別 "
Private Const RPT_RECTYPE_SEC_CLOSE As String = " の項目一覧]"
Private Const RPT_PII_HEAD As String = "  個人情報の可能性がある項目: "
Private Const RPT_TRUNC_NOTE As String = "  注: 先頭の一部の行だけを解析している"

' ---- 備考の文言 ----
Private Const RMK_KEY As String = "全件一意 (キー候補)"
Private Const RMK_FIXED As String = "固定値"
Private Const RMK_CODE As String = "コード値候補 ("
Private Const RMK_CODE_UNIT As String = " 種類)"
Private Const RMK_LEAD0 As String = "先頭0あり"
Private Const RMK_FIXLEN As String = "固定長 "
Private Const RMK_BYTES As String = "バイト "
Private Const RMK_TRAIL_SPACE As String = "末尾空白あり"
Private Const RMK_LEAD_SPACE As String = "先頭空白あり"
Private Const RMK_BLANK_ONLY As String = "空白のみの値 "
Private Const RMK_NEWLINE As String = "値の中に改行あり"
Private Const RMK_ERROR As String = "エラー値 "
Private Const RMK_DEC_PLACES As String = "小数 最大 "
Private Const RMK_DEC_UNIT As String = " 桁"
Private Const RMK_MIX_DETAIL As String = "内訳 "
Private Const RMK_PII As String = "個人情報の可能性"
Private Const RMK_NAME_LIKE As String = "氏名らしい形 (空白で区切った語) 個人情報の可能性"
Private Const RMK_LONG As String = "長い値はパターンを省略"
Private Const RMK_ALL_EMPTY As String = "全行空"
Private Const RMK_WAREKI_YM As String = "和暦年月の形 (GYYMM。G=元号 1明治 2大正 3昭和 4平成 5令和)"
Private Const RMK_WAREKI_YMD As String = "和暦年月日の形 (GYYMMDD。G=元号 1明治 2大正 3昭和 4平成 5令和)"
Private Const RPT_SINGLE_ROW_NOTE As String = "  注: この種別は 1 行だけなので、型・桁数・書式はその 1 行から判定している (可変長の項目は、ほかのファイルでは桁数が変わりうる)"
Private Const RMK_SEP As String = "; "
Private Const SUM_ERROR As String = "解析中にエラー (ログを確認)"
Private Const RPT_ERROR_LINE As String = "  解析中にエラーが発生した。C:\VBAErrorLogs\ を確認すること。"
Private Const RPT_SUM_EMPTY As String = "値が無い"
Private Const RPT_SUM_INDENT As String = "    "
Private Const RPT_SUM_TABLE As String = "表 "
Private Const RPT_SUM_MULTI_1 As String = "別票の混在あり: 見出しの異なる表が "
Private Const RPT_SUM_MULTI_2 As String = " 個"
Private Const RPT_TABLE_SEC As String = "[表 "
Private Const RPT_TABLE_SEC_ROWS As String = " (行 "
Private Const RPT_NEXT_TABLE As String = "  この表の後ろに見出しの異なる表が続く: 行 "
Private Const RPT_LIST_OPEN As String = " (行 "
Private Const RPT_RECTYPE_KEY_LABEL As String = "(種別コード)"
' ---- 層構造 (レコードの並び) ----
Private Const ROLE_HEADER As String = "ヘッダ"
Private Const ROLE_TRAILER As String = "トレーラ"
Private Const ROLE_DETAIL As String = "明細"
Private Const ROLE_GROUP_HEAD As String = "グループの先頭"
Private Const ROLE_IN_GROUP As String = "グループ内"
Private Const RPT_SEC_LAYER As String = "[層構造 (レコードの並び)]"
Private Const RPT_LAYER As String = "  層 "
Private Const RPT_LAYER_ROWS As String = " : 行 "
Private Const RPT_LAYER_KINDS As String = " / 種別 "
Private Const RPT_LAYER_ROLE As String = " / "
Private Const RPT_LAYER_GROUPS_1 As String = " で始まるグループが "
Private Const RPT_LAYER_GROUPS_2 As String = " 個、計 "
Private Const RPT_LAYER_UNIT As String = " 行"
Private Const RPT_GROUP_BODY As String = "  1 グループの中身 (出現順): "
Private Const RPT_GROUP_PER As String = " (1 グループに "
Private Const RPT_LAYER_ORDER As String = "  並び: "
Private Const RPT_LAYER_HOWTO As String = "  読み取り方: 各行の 1 項目目 (種別コード) で種別を判定し、下の種別ごとの項目一覧に従って項目を分ける。種別ごとに項目数が異なるので、1 つの表として読まないこと。"
Private Const RPT_LAYER_TIMES As String = " ×"
Private Const RPT_ARROW As String = " → "
Private Const RPT_RANGE_SEP As String = "-"
Private Const SUM_ORDER As String = " / 並び: "
Private Const REL_NAME_PREFIX As String = "No."
Private Const RMK_PCT_LIKE As String = "割合らしい値 (0-100 の 10 刻み。数量の可能性もある)"
Private Const PCT_STEP As Long = 10
Private Const PCT_MAX As Long = 100
Private Const PCT_MAX_KINDS As Long = 11
Private Const PCT_MAX_LEN As Long = 3
Private Const RPT_SUM_KINDS As String = " / 種別 "
Private Const RPT_SUM_KINDS_UNIT As String = " 種類 ("
Private Const SUM_HEADER_AT As String = " / 見出し行 "
Private Const SUM_NO_HEADER As String = " / 見出しなし"
Private Const SUM_ITEMS As String = " / 項目数 "
Private Const SUM_DATA As String = " / データ "
Private Const SUM_ROWS_UNIT As String = " 行"
Private Const SUM_TITLE As String = " / 表題 "
Private Const SUM_JUDGE As String = " / 別票の混在: "
Private Const RMK_QUOTED As String = "引用符あり "
Private Const RMK_NUMFMT As String = "表示形式 "
Private Const FMT_GENERAL_EN As String = "General"
Private Const FMT_GENERAL_JA As String = "G/標準"
Private Const CORE_DATE_JA As String = "年月日"
Private Const CORE_DATETIME_JA_HM As String = "年月日時分"
Private Const CORE_DATETIME_JA_HMS As String = "年月日時分秒"
Private Const TYPE_MAIN_OPEN As String = " (主: "
Private Const RMK_BYTES_UNIT As String = " バイト"
Private Const PROFILE_ERROR As String = "(エラー) | - | - | - | - | 解析中にエラー |"


'----------------------------------------------------------------------
' AnalyzeGrid: 2 次元配列の構造を解析し、報告文を返す
'   grid          値 (1 始まり)。Excel の値ならセルの型を持つ
'   rowWidth()    行ごとの項目数 (区切り文字形式)。Excel では行の最終列を内部で使う
'   quotedCount() 列ごとの引用符付きの値の数 (区切り文字形式のみ。無ければ空配列)
'   ws            Excel のシート (区切り文字形式では Nothing)
'   isText        区切り文字形式なら True
'   summaryOut    結論欄に載せる 1 行を返す
'----------------------------------------------------------------------
Public Function AnalyzeGrid(ByRef grid As Variant, ByVal rowCount As Long, ByVal colCount As Long, _
                              ByRef rowWidth() As Long, ByRef quotedCount() As Long, _
                              ByVal ws As Object, ByVal isText As Boolean, ByVal truncated As Boolean, _
                              ByRef summaryOut As String) As String
      On Error GoTo ErrHandler
      Dim rep As String
      Dim r As Long
      Dim c As Long
      Dim v As Variant
      Dim s As String
      Dim nonEmpty() As Long
      Dim firstC() As Long
      Dim lastC() As Long
      Dim numberish() As Long
      Dim digitCells() As Long
      Dim longCells() As Long
      Dim structW() As Long
      Dim widthDict As Object
      Dim key As Variant
      Dim modeWidth As Long
      Dim modeFreq As Long
      Dim nonEmptyRows As Long
      Dim blankRows As Long
      Dim modeShare As Double
      Dim firstVals As Object
      Dim isRecType As Boolean
      Dim keyOk As Boolean
      Dim kindWidth As Object
      Dim kindBest As Long
      Dim startRow As Long
      Dim tableNo As Long
      Dim nextStart As Long
      Dim sumPart As String
      Dim sumAll As String
      Dim singles As Long
      Dim kw As Variant
      Dim mainKey As String
      Dim mainMixed As Boolean

      ' ---- 行ごとの特徴 ----
10    ReDim nonEmpty(1 To rowCount)
20    ReDim firstC(1 To rowCount)
30    ReDim lastC(1 To rowCount)
40    ReDim numberish(1 To rowCount)
50    ReDim digitCells(1 To rowCount)
60    ReDim longCells(1 To rowCount)
70    ReDim structW(1 To rowCount)
80    For r = 1 To rowCount
90        For c = 1 To colCount
100           v = grid(r, c)
110           If IsError(v) Then
120               nonEmpty(r) = nonEmpty(r) + 1
130               If firstC(r) = 0 Then firstC(r) = c
140               lastC(r) = c
              ElseIf Not IsEmpty(v) Then
150               If VarType(v) = vbString Then
160                   s = Trim$(v)
170                   If Len(s) > 0 Then
180                       nonEmpty(r) = nonEmpty(r) + 1
190                       If firstC(r) = 0 Then firstC(r) = c
200                       lastC(r) = c
210                       If s Like "*[0-9]*" Then
220                           digitCells(r) = digitCells(r) + 1
230                           If Not s Like "*[!0-9.,+/: -]*" Then numberish(r) = numberish(r) + 1
                          End If
240                       If Len(s) > LABEL_MAX_LEN Then longCells(r) = longCells(r) + 1
                      End If
                  Else
                      ' Excel の数値・日付・真偽値
250                   nonEmpty(r) = nonEmpty(r) + 1
260                   If firstC(r) = 0 Then firstC(r) = c
270                   lastC(r) = c
280                   If VarType(v) <> vbBoolean Then
290                       numberish(r) = numberish(r) + 1
300                       digitCells(r) = digitCells(r) + 1
                      End If
                  End If
              End If
          Next c
310       If isText Then
320           structW(r) = IIf(nonEmpty(r) > 0, rowWidth(r), 0)
          Else
330           structW(r) = lastC(r)
          End If
      Next r

      ' ---- 項目数の分布 ----
340   Set widthDict = CreateObject("Scripting.Dictionary")
350   For r = 1 To rowCount
360       If nonEmpty(r) > 0 Then
370           nonEmptyRows = nonEmptyRows + 1
380           widthDict(structW(r)) = widthDict(structW(r)) + 1
          Else
390           blankRows = blankRows + 1
          End If
      Next r
400   For Each key In widthDict.Keys
410       If widthDict(key) > modeFreq Then
420           modeFreq = widthDict(key)
430           modeWidth = key
          End If
      Next key
440   If nonEmptyRows > 0 Then modeShare = modeFreq / nonEmptyRows

450   rep = RPT_SEC_STRUCT & vbCrLf
460   If truncated Then rep = rep & RPT_TRUNC_NOTE & vbCrLf
470   rep = rep & RPT_ROWS_TOTAL & rowCount & RPT_ROWS_UNIT & blankRows & RPT_ROWS_CLOSE & vbCrLf & _
                  RPT_WIDTH_MODE & modeWidth & RPT_WIDTH_SHARE & Format(modeShare, "0.0%") & ")" & vbCrLf & _
                  RPT_WIDTH_KINDS & widthDict.Count & vbCrLf

480   If nonEmptyRows = 0 Then
490       summaryOut = RPT_SUM_EMPTY
500       AnalyzeGrid = rep
          LogSuccess MODULE_NAME, "AnalyzeGrid", "Grid has no values"
510       Exit Function
      End If

      ' ---- レコード種別の判定 (区切り文字形式で、項目数の違う行があるとき) ----
      ' 先頭項目が短いコードで種類が少なく、種別ごとには項目数が揃っていることを条件とする。
      ' 項目数がほぼ揃っている (ヘッダ 1 行 + 明細 など) ときは、主な項目数の行の先頭項目が
      ' 1 種類の固定値 (レコード識別) であることも条件とする (普通の表の先頭列を種別と誤らないため)
520   If isText And widthDict.Count > 1 Then
530       Set firstVals = CreateObject("Scripting.Dictionary")
540       Set kindWidth = CreateObject("Scripting.Dictionary")
550       keyOk = True
560       For r = 1 To rowCount
570           If nonEmpty(r) > 0 Then
580               s = Trim$(CStr(grid(r, 1)))
590               If Len(s) = 0 Or Len(s) > RECTYPE_MAX_KEYLEN Then
600                   keyOk = False
610                   Exit For
                  End If
620               firstVals(s) = firstVals(s) + 1
630               kindWidth(s & vbTab & structW(r)) = kindWidth(s & vbTab & structW(r)) + 1
640               If structW(r) = modeWidth Then
650                   If Len(mainKey) = 0 Then
660                       mainKey = s
                      ElseIf s <> mainKey Then
670                       mainMixed = True
                      End If
                  End If
680               If firstVals.Count > RECTYPE_MAX_KINDS Then
690                   keyOk = False
700                   Exit For
                  End If
              End If
          Next r
710       If modeShare >= RECTYPE_WIDTH_SHARE And mainMixed Then keyOk = False
720       If keyOk And firstVals.Count >= 2 Then
              ' 種別ごとに、最も多い項目数の行が RECTYPE_WIDTH_SHARE 以上あるか
              ' (1 行しかない種別ばかりなら、種別コードではなく普通の値とみなす)
730           isRecType = True
740           singles = 0
750           For Each key In firstVals.Keys
760               If firstVals(key) = 1 Then singles = singles + 1
              Next key
770           If singles > RECTYPE_MAX_SINGLES Then isRecType = False
780           For Each key In firstVals.Keys
790               If Not isRecType Then Exit For
800               kindBest = 0
810               For Each kw In kindWidth.Keys
820                   If Left$(kw, Len(key) + 1) = key & vbTab Then
830                       If kindWidth(kw) > kindBest Then kindBest = kindWidth(kw)
                      End If
                  Next kw
840               If kindBest < firstVals(key) * RECTYPE_WIDTH_SHARE Then
850                   isRecType = False
860                   Exit For
                  End If
              Next key
          End If
      End If

870   If isRecType Then
880       rep = rep & RPT_FORMAT & RPT_MODE_RECTYPE & vbCrLf & vbCrLf & _
                AnalyzeRecordTypes(grid, rowCount, colCount, nonEmpty, structW, quotedCount, summaryOut)
      Else
          ' 別の見出しが見つかるたびに表を分けて、表ごとに解析する
890       rep = rep & RPT_FORMAT & RPT_MODE_TABLE & vbCrLf
900       startRow = 1
910       tableNo = 0
920       Do
930           tableNo = tableNo + 1
940           rep = rep & AnalyzeTable(grid, rowCount, colCount, nonEmpty, firstC, lastC, numberish, digitCells, _
                                       longCells, structW, modeWidth, quotedCount, ws, isText, _
                                       startRow, tableNo, nextStart, sumPart)
950           sumAll = sumAll & IIf(Len(sumAll) > 0, vbCrLf & RPT_SUM_INDENT, "") & _
                       IIf(nextStart > 0 Or tableNo > 1, RPT_SUM_TABLE & tableNo & ": ", "") & sumPart
960           If nextStart <= startRow Then Exit Do
970           startRow = nextStart
          Loop While tableNo < MAX_TABLES
980       If tableNo > 1 Then
990           summaryOut = RPT_SUM_MULTI_1 & tableNo & RPT_SUM_MULTI_2 & vbCrLf & RPT_SUM_INDENT & sumAll
          Else
1000          summaryOut = sumAll
          End If
      End If
1010  AnalyzeGrid = rep

      LogSuccess MODULE_NAME, "AnalyzeGrid", _
                 "Analyzed grid " & rowCount & "x" & colCount & "; recordType=" & isRecType
1020  Exit Function
ErrHandler:
      LogError MODULE_NAME, "AnalyzeGrid", Err.Number, Err.Description, Erl, _
               "rowCount=" & rowCount & "; colCount=" & colCount & "; r=" & r & "; c=" & c
      summaryOut = SUM_ERROR
      AnalyzeGrid = rep & RPT_ERROR_LINE & vbCrLf
End Function

'----------------------------------------------------------------------
' AnalyzeTable: 表形式 (見出し行 + データ行) として解析する
'   rowFrom    この表を探し始める行
'   tableNo    表の番号 (報告の見出しに使う)
'   nextStart  見出しの異なる表が後ろに続くなら、その開始行 (表題を含む)。無ければ 0
'----------------------------------------------------------------------
Private Function AnalyzeTable(ByRef grid As Variant, ByVal rowCount As Long, ByVal colCount As Long, _
                                ByRef nonEmpty() As Long, ByRef firstC() As Long, ByRef lastC() As Long, _
                                ByRef numberish() As Long, ByRef digitCells() As Long, ByRef longCells() As Long, _
                                ByRef structW() As Long, ByVal modeWidth As Long, ByRef quotedCount() As Long, _
                                ByVal ws As Object, ByVal isText As Boolean, _
                                ByVal rowFrom As Long, ByVal tableNo As Long, ByRef nextStart As Long, _
                                ByRef summaryOut As String) As String
      On Error GoTo ErrHandler
      Dim rep As String
      Dim tableEnd As Long
      Dim fillCount() As Long
      Dim fillRows As Long
      Dim cum As Long
      Dim typicalFill As Long
      Dim sparseLimit As Double
      Dim firstNE As Long
      Dim lastNE As Long
      Dim edgeCount As Long
      Dim edgeList As String
      Dim sameAsHeader As Boolean
      Dim r As Long
      Dim r2 As Long
      Dim c As Long
      Dim i As Long
      Dim seen As Long
      Dim isCand() As Boolean
      Dim cellDict As Object
      Dim s As String
      Dim contrastRows As Long
      Dim checked As Long
      Dim headerStart As Long
      Dim headerEnd As Long
      Dim headerLines As Long
      Dim tableWidth As Long
      Dim labels() As String
      Dim relNames() As String
      Dim titleText As String
      Dim titleRows As Long
      Dim lastTitleRow As Long
      Dim rowText As String
      Dim dataRows() As Long
      Dim dataCount As Long
      Dim firstDataRow As Long
      Dim lastDataRow As Long
      Dim isRepeat As Boolean
      Dim repeatList As String
      Dim repeatCount As Long
      Dim oddList As String
      Dim oddCount As Long
      Dim sparseList As String
      Dim sparseCount As Long
      Dim summaryList As String
      Dim summaryCount As Long
      Dim blocks As Long
      Dim inBlock As Boolean
      Dim words As Variant
      Dim w As Long
      Dim hasSummaryWord As Boolean
      Dim judge As String
      Dim values() As Variant
      Dim extra As String
      Dim label As String
      Dim pos As String
      Dim piiList As String
      Dim piiWords As Variant
      Dim mergedText As String
      Dim merged As Object
      Dim addr As String
      Dim k As Variant
      Dim fmt As String
      Dim lineOut As String

      ' ---- 見出し候補: 数値らしいセルが無く、長い値が無く、十分に埋まり、値が互いに異なる行 ----
10    ReDim isCand(1 To rowCount)
20    For r = rowFrom To rowCount
30        If nonEmpty(r) >= HEADER_MIN_CELLS And nonEmpty(r) >= modeWidth * HEADER_MIN_FILL _
             And numberish(r) = 0 And longCells(r) = 0 Then
40            Set cellDict = CreateObject("Scripting.Dictionary")
50            For c = firstC(r) To lastC(r)
60                If Not IsError(grid(r, c)) Then
70                    s = Trim$(CStr(grid(r, c)))
80                    If Len(s) > 0 Then cellDict(s) = True
                  End If
              Next c
90            If cellDict.Count >= nonEmpty(r) * HEADER_UNIQUE_RATIO Then isCand(r) = True
          End If
      Next r

      ' ---- 見出し行: 先頭付近の候補のうち、下の行に数字を含む行が多いもの ----
100   seen = 0
110   nextStart = 0
120   For r = rowFrom To rowCount
130       If nonEmpty(r) > 0 Then
140           seen = seen + 1
150           If seen > HEADER_SEARCH_ROWS Then Exit For
160           If isCand(r) Then
170               contrastRows = 0
180               checked = 0
190               For r2 = r + 1 To rowCount
200                   If checked >= HEADER_CONTRAST_ROWS Then Exit For
210                   If nonEmpty(r2) > 0 And Not isCand(r2) Then
220                       checked = checked + 1
230                       If digitCells(r2) > 0 Then contrastRows = contrastRows + 1
                      End If
                  Next r2
240               If checked > 0 And contrastRows * 2 >= checked Then
250                   headerStart = r
260                   Exit For
                  End If
              End If
          End If
      Next r

270   If headerStart > 0 Then
          ' 続く行も見出し候補で、列範囲が重なるなら複数行の見出しとする
280       headerEnd = headerStart
290       Do While headerEnd < rowCount And (headerEnd - headerStart + 1) < MAX_HEADER_LINES
300           If isCand(headerEnd + 1) And digitCells(headerEnd + 1) = 0 _
                 And firstC(headerEnd + 1) <= lastC(headerStart) And lastC(headerEnd + 1) >= firstC(headerStart) Then
310               headerEnd = headerEnd + 1
              Else
320               Exit Do
              End If
          Loop
330       headerLines = headerEnd - headerStart + 1
340       For r = headerStart To headerEnd
350           If lastC(r) > tableWidth Then tableWidth = lastC(r)
          Next r
360       If isText And modeWidth > tableWidth Then tableWidth = modeWidth
      Else
370       headerEnd = rowFrom - 1
380       tableWidth = modeWidth
      End If

      ' ---- 見出しの異なる表の始まりを探し、この表の範囲を決める ----
      ' 見出し候補の形で数字を含まず、見出し行と同じではなく、直前が空行か疎な行で、
      ' 直後の数行に数字を含む行がある行を、次の表の見出しとみなす
390   tableEnd = rowCount
400   If headerStart > 0 Then
410       For r = headerEnd + 2 To rowCount
420           If isCand(r) And digitCells(r) = 0 Then
430               sameAsHeader = (nonEmpty(r) = nonEmpty(headerStart))
440               If sameAsHeader Then
450                   For c = 1 To tableWidth
460                       If CStr(IIf(IsError(grid(r, c)), "#ERR", grid(r, c))) <> _
                             CStr(IIf(IsError(grid(headerStart, c)), "#ERR", grid(headerStart, c))) Then
470                           sameAsHeader = False
480                           Exit For
                          End If
                      Next c
                  End If
490               If Not sameAsHeader Then
500                   If nonEmpty(r - 1) = 0 Or nonEmpty(r - 1) < tableWidth * SPARSE_RATIO Then
510                       checked = 0
520                       For r2 = r + 1 To rowCount
530                           If checked >= HEADER_CONTRAST_ROWS Then Exit For
540                           If nonEmpty(r2) > 0 Then
550                               checked = checked + 1
560                               If digitCells(r2) > 0 Then
570                                   nextStart = r
580                                   Exit For
                                  End If
                              End If
                          Next r2
                      End If
                  End If
              End If
590           If nextStart > 0 Then Exit For
          Next r
600       If nextStart > 0 Then
              ' 直前の疎な行 (次の表の表題) も次の表に含める
610           Do While nextStart - 1 > headerEnd + 1
620               If nonEmpty(nextStart - 1) > 0 And nonEmpty(nextStart - 1) < tableWidth * SPARSE_RATIO Then
630                   nextStart = nextStart - 1
                  Else
640                   Exit Do
                  End If
              Loop
650           tableEnd = nextStart - 1
          End If
      End If

      ' ---- データ行の範囲で、見出しより右に値があれば項目に含める ----
660   For r = headerEnd + 1 To tableEnd
670       If lastC(r) > tableWidth Then tableWidth = lastC(r)
      Next r
680   If tableWidth > colCount Then tableWidth = colCount

      ' ---- ラベル ----
690   ReDim labels(1 To tableWidth)
700   ReDim relNames(1 To tableWidth)
710   For c = 1 To tableWidth
720       If headerStart > 0 Then
730           For r = headerStart To headerEnd
740               If Not IsError(grid(r, c)) Then
750                   s = Trim$(CStr(grid(r, c)))
760                   s = Replace(Replace(s, vbCr, " "), vbLf, " ")
770                   If Len(s) > 0 Then labels(c) = labels(c) & IIf(Len(labels(c)) > 0, " / ", "") & s
                  End If
              Next r
          End If
      Next c

      ' ---- 表題 (見出し行より上の空でない行。数字は 9 に置き換える) ----
780   If headerStart > rowFrom Then
790       For r = rowFrom To headerStart - 1
800           If nonEmpty(r) > 0 Then
810               titleRows = titleRows + 1
820               lastTitleRow = r
830               rowText = ""
840               For c = firstC(r) To lastC(r)
850                   If Not IsError(grid(r, c)) Then
860                       If IsDate(grid(r, c)) And VarType(grid(r, c)) = vbDate Then
870                           s = Format(grid(r, c), "yyyy/mm/dd")
                          Else
880                           s = Trim$(CStr(grid(r, c)))
                          End If
890                       If Len(s) > 0 Then rowText = rowText & IIf(Len(rowText) > 0, " / ", "") & s
                      End If
                  Next c
900               rowText = MaskDigits(rowText)
910               If Len(rowText) > TITLE_MAX_LEN Then rowText = Left$(rowText, TITLE_MAX_LEN) & "..."
920               titleText = titleText & RPT_TITLE_ROW & r & RPT_TITLE_COLS & ColumnRef(firstC(r), isText) & "-" & _
                              ColumnRef(lastC(r), isText) & "): " & rowText & vbCrLf
              End If
          Next r
      End If

      ' ---- 典型的な埋まり具合 (空でない項目数の中央値)。空の項目が多い形式でも疎な行を正しく見分ける ----
930   ReDim fillCount(0 To colCount)
940   fillRows = 0
950   For r = headerEnd + 1 To tableEnd
960       If nonEmpty(r) > 0 Then
970           fillCount(nonEmpty(r)) = fillCount(nonEmpty(r)) + 1
980           fillRows = fillRows + 1
          End If
      Next r
990   typicalFill = 0
1000  cum = 0
1010  For c = 0 To colCount
1020      cum = cum + fillCount(c)
1030      If cum * 2 >= fillRows And fillRows > 0 Then
1040          typicalFill = c
1050          Exit For
          End If
      Next c
1060  sparseLimit = typicalFill * SPARSE_RATIO
1070  firstNE = 0
1080  lastNE = 0
1090  For r = headerEnd + 1 To tableEnd
1100      If nonEmpty(r) > 0 Then
1110          If firstNE = 0 Then firstNE = r
1120          lastNE = r
          End If
      Next r

      ' ---- データ行の分類 ----
1130  words = Split(SUMMARY_WORDS, LIST_SEP)
1140  ReDim dataRows(1 To rowCount)
1150  For r = headerEnd + 1 To tableEnd
1160      If nonEmpty(r) = 0 Then
1170          inBlock = False
          Else
1180          If Not inBlock Then blocks = blocks + 1
1190          inBlock = True

              ' 見出しと同じ行か
              ' VBA の And は短絡しないため、headerStart = 0 のときに添字が外れないよう入れ子にする
1200          isRepeat = False
1210          If headerStart > 0 Then
1220              If nonEmpty(r) = nonEmpty(headerStart) Then
1230                  isRepeat = True
1240                  For c = 1 To tableWidth
1250                      If CStr(IIf(IsError(grid(r, c)), "#ERR", grid(r, c))) <> _
                             CStr(IIf(IsError(grid(headerStart, c)), "#ERR", grid(headerStart, c))) Then
1260                          isRepeat = False
1270                          Exit For
                          End If
                      Next c
                  End If
              End If

              ' 小計・合計の語を含む短いセルがあるか
1280          hasSummaryWord = False
1290          For c = firstC(r) To lastC(r)
1300              If VarType(grid(r, c)) = vbString Then
1310                  s = UCase$(Replace(Trim$(grid(r, c)), " ", ""))
1320                  If Len(s) > 0 And Len(s) <= LABEL_MAX_LEN \ 4 Then
1330                      For w = LBound(words) To UBound(words)
1340                          If InStr(1, s, words(w), vbTextCompare) > 0 Then hasSummaryWord = True
                          Next w
                      End If
                  End If
              Next c
1350          If hasSummaryWord Then
1360              summaryCount = summaryCount + 1
1370              If summaryCount <= MAX_LIST_ITEMS Then summaryList = summaryList & r & " "
              End If

1380          If isRepeat Then
1390              repeatCount = repeatCount + 1
1400              If repeatCount <= MAX_LIST_ITEMS Then repeatList = repeatList & r & " "
              ElseIf isText And structW(r) <> modeWidth And (r = firstNE Or r = lastNE) Then
1410              edgeCount = edgeCount + 1
1420              edgeList = edgeList & r & " "
              ElseIf nonEmpty(r) < sparseLimit Then
1430              sparseCount = sparseCount + 1
1440              If sparseCount <= MAX_LIST_ITEMS Then sparseList = sparseList & r & " "
              Else
1450              If isText And structW(r) <> modeWidth Then
1460                  oddCount = oddCount + 1
1470                  If oddCount <= MAX_LIST_ITEMS Then oddList = oddList & r & " "
                  End If
1480              dataCount = dataCount + 1
1490              dataRows(dataCount) = r
1500              If firstDataRow = 0 Then firstDataRow = r
1510              lastDataRow = r
              End If
          End If
      Next r

      ' ---- 構造の報告 ----
1520  If tableNo > 1 Or nextStart > 0 Then
1530      rep = vbCrLf & RPT_TABLE_SEC & tableNo & RPT_TABLE_SEC_ROWS & rowFrom & "-" & tableEnd & ")]" & vbCrLf
      End If
1540  If headerStart > 0 Then
1550      rep = rep & RPT_HEADER_AT & headerStart & IIf(headerLines > 1, "-" & headerEnd, "") & _
                " (" & headerLines & RPT_HEADER_LINES & vbCrLf
      Else
1560      rep = rep & RPT_HEADER_NONE & vbCrLf
      End If
1570  If dataCount > 0 Then
1580      rep = rep & RPT_DATA_AT & firstDataRow & "-" & lastDataRow & RPT_DATA_COUNT & dataCount & RPT_DATA_UNIT & vbCrLf
      End If
1590  If repeatCount = 0 And oddCount = 0 Then
1600      judge = RPT_JUDGE_NONE
      ElseIf oddCount = 0 Then
1610      judge = RPT_JUDGE_REPEAT
      Else
1620      judge = RPT_JUDGE_MIXED
      End If
1630  rep = rep & RPT_JUDGE & judge & vbCrLf & vbCrLf

1640  rep = rep & RPT_SEC_TITLE & vbCrLf
1650  If titleRows = 0 Then
1660      rep = rep & RPT_NO_TITLE & vbCrLf
      Else
1670      rep = rep & titleText & RPT_TITLE_NOTE & vbCrLf & _
                RPT_GAP & (headerStart - lastTitleRow - 1) & RPT_DATA_UNIT & vbCrLf
      End If
1680  rep = rep & vbCrLf

      ' Excel のシートなら見出し行までの結合セルを載せる
1690  If Not ws Is Nothing And headerEnd > 0 Then
1700      Set merged = CreateObject("Scripting.Dictionary")
1710      For r = rowFrom To headerEnd
1720          If nonEmpty(r) > 0 Then
1730              For c = 1 To tableWidth
1740                  If ws.Cells(r, c).MergeCells Then
1750                      addr = ws.Cells(r, c).MergeArea.Address(False, False)
1760                      If Not merged.Exists(addr) Then merged.Add addr, True
                      End If
                  Next c
              End If
          Next r
1770      For Each k In merged.Keys
1780          mergedText = mergedText & k & " "
          Next k
1790      rep = rep & RPT_SEC_MERGED & vbCrLf & "  " & IIf(Len(mergedText) > 0, mergedText, RPT_NONE) & vbCrLf & vbCrLf
      End If

1800  rep = rep & RPT_SEC_ROWS & vbCrLf & _
            RPT_REPEAT_HEADER & ListText(repeatList, repeatCount) & vbCrLf & _
            IIf(nextStart > 0, RPT_NEXT_TABLE & nextStart & vbCrLf, "") & _
            IIf(isText, RPT_ODD_WIDTH & ListText(oddList, oddCount) & vbCrLf, "") & _
            IIf(edgeCount > 0, RPT_EDGE & ListText(edgeList, edgeCount) & vbCrLf, "") & _
            RPT_SPARSE & ListText(sparseList, sparseCount) & vbCrLf & _
            RPT_SUMMARY_ROWS & ListText(summaryList, summaryCount) & vbCrLf & _
            RPT_BLOCKS & blocks & vbCrLf & vbCrLf

      ' ---- 項目一覧 ----
1810  rep = rep & RPT_SEC_ITEMS & vbCrLf & RPT_TABLE_HEAD_GRID & vbCrLf & RPT_TABLE_RULE & vbCrLf
1820  piiWords = Split(PII_WORDS, LIST_SEP)
1830  For c = 1 To tableWidth
1840      If dataCount > 0 Then
1850          ReDim values(1 To dataCount)
1860          For i = 1 To dataCount
1870              values(i) = grid(dataRows(i), c)
              Next i
          Else
1880          ReDim values(1 To 1)
          End If

1890      extra = ""
1900      If isText Then
1910          If UBound(quotedCount) >= c And LBound(quotedCount) <= c Then
1920              If quotedCount(c) > 0 Then extra = RMK_QUOTED & quotedCount(c) & RPT_ITEMS_UNIT
              End If
          ElseIf Not ws Is Nothing And firstDataRow > 0 Then
1930          fmt = ws.Cells(firstDataRow, c).NumberFormat
1940          If fmt <> FMT_GENERAL_EN And fmt <> FMT_GENERAL_JA And Len(fmt) > 0 Then extra = RMK_NUMFMT & fmt
          End If

1950      label = labels(c)
1960      If Len(label) = 0 Then
1970          If headerStart > 0 Then label = RPT_NO_LABEL Else label = RPT_ITEM & c & RPT_ITEM_CLOSE
          Else
1980          For w = LBound(piiWords) To UBound(piiWords)
1990              If InStr(1, label, piiWords(w), vbTextCompare) > 0 Then
2000                  extra = extra & IIf(Len(extra) > 0, RMK_SEP, "") & RMK_PII
2010                  piiList = piiList & label & " "
2020                  Exit For
                  End If
              Next w
          End If
2030      pos = ColumnRef(c, isText)
2040      relNames(c) = REL_NAME_PREFIX & c & " " & label
2050      lineOut = "| " & c & " | " & pos & " | " & Replace(label, "|", "/") & " | " & _
                    ProfileValues(values, IIf(dataCount > 0, dataCount, 0), extra)
2060      rep = rep & lineOut & vbCrLf
      Next c
2070  If Len(piiList) > 0 Then rep = rep & vbCrLf & RPT_PII_HEAD & piiList & vbCrLf
2080  If dataCount > 0 Then rep = rep & vbCrLf & RelationFinder.FindRelations(grid, dataRows, dataCount, tableWidth, relNames)

2090  summaryOut = RPT_MODE_TABLE & IIf(headerStart > 0, SUM_HEADER_AT & headerStart, SUM_NO_HEADER) & _
                   SUM_ITEMS & tableWidth & SUM_DATA & dataCount & SUM_ROWS_UNIT & _
                   IIf(titleRows > 0, SUM_TITLE & titleRows & SUM_ROWS_UNIT, "") & SUM_JUDGE & judge
2100  AnalyzeTable = rep

      LogSuccess MODULE_NAME, "AnalyzeTable", _
                 "Analyzed table; header=" & headerStart & "-" & headerEnd & "; width=" & tableWidth & "; dataRows=" & dataCount
2110  Exit Function
ErrHandler:
      LogError MODULE_NAME, "AnalyzeTable", Err.Number, Err.Description, Erl, _
               "r=" & r & "; c=" & c & "; rowFrom=" & rowFrom & "; headerStart=" & headerStart & "; tableWidth=" & tableWidth
      nextStart = 0
      summaryOut = SUM_ERROR
      AnalyzeTable = rep & RPT_ERROR_LINE & vbCrLf
End Function

'----------------------------------------------------------------------
' AnalyzeRecordTypes: 先頭項目が種別コードで、種別ごとに項目数が違う形式を解析する
'   種別コードは書式の一部なので、そのまま載せる
'----------------------------------------------------------------------
Private Function AnalyzeRecordTypes(ByRef grid As Variant, ByVal rowCount As Long, ByVal colCount As Long, _
                                      ByRef nonEmpty() As Long, ByRef structW() As Long, _
                                      ByRef quotedCount() As Long, ByRef summaryOut As String) As String
      On Error GoTo ErrHandler
      Dim rep As String
      Dim kinds As Object
      Dim kindRows As Object
      Dim kindWidths As Object
      Dim kindFirst As Object
      Dim r As Long
      Dim c As Long
      Dim i As Long
      Dim s As String
      Dim k As Variant
      Dim rowList As Variant
      Dim seq As String
      Dim seqCount As Long
      Dim widthText As String
      Dim w As Variant
      Dim maxW As Long
      Dim values() As Variant
      Dim cnt As Long
      Dim idx() As Long
      Dim rowColl As Collection
      Dim rowItem As Variant
      Dim widthCounts As Object
      Dim seqKind() As String
      Dim seqRow() As Long
      Dim n As Long
      Dim p As Long
      Dim lastPos As Object
      Dim segStart() As Long
      Dim segEnd() As Long
      Dim segCount As Long
      Dim seg As Long
      Dim curStart As Long
      Dim maxLast As Long
      Dim bodySeg As Long
      Dim segKinds As Object
      Dim bodyKinds As Object
      Dim role As Object
      Dim roleNote As Object
      Dim starter As String
      Dim groupMode As Boolean
      Dim groupCount As Long
      Dim isBoundary As Boolean
      Dim cur As Object
      Dim minC As Object
      Dim maxC As Object
      Dim rangeText As String
      Dim layerText As String
      Dim orderText As String
      Dim groupText As String
      Dim segRole As String
      Dim relNames() As String

10    Set kinds = CreateObject("Scripting.Dictionary")
20    Set kindRows = CreateObject("Scripting.Dictionary")
30    Set kindWidths = CreateObject("Scripting.Dictionary")
40    Set kindFirst = CreateObject("Scripting.Dictionary")
50    Set lastPos = CreateObject("Scripting.Dictionary")
60    Set role = CreateObject("Scripting.Dictionary")
70    Set roleNote = CreateObject("Scripting.Dictionary")
80    ReDim seqKind(1 To rowCount)
90    ReDim seqRow(1 To rowCount)

100   For r = 1 To rowCount
110       If nonEmpty(r) > 0 Then
120           s = Trim$(CStr(grid(r, 1)))
130           n = n + 1
140           seqKind(n) = s
150           seqRow(n) = r
160           lastPos(s) = n
170           If Not kinds.Exists(s) Then
180               kinds.Add s, 0
190               kindRows.Add s, New Collection
200               kindWidths.Add s, CreateObject("Scripting.Dictionary")
210               kindFirst.Add s, r
              End If
220           kinds(s) = kinds(s) + 1
230           Set rowColl = kindRows(s)
240           rowColl.Add r
250           Set widthCounts = kindWidths(s)
260           widthCounts(structW(r)) = widthCounts(structW(r)) + 1
270           If seqCount < RECTYPE_SEQ_SHOW Then
280               seq = seq & s & " "
290               seqCount = seqCount + 1
              End If
          End If
      Next r

      ' ---- 層に分ける ----
      ' 先頭から見て、そこまでに出た種別がそこより後ろに現れない位置で区切る。
      ' いちばん長い区間を本体 (明細) とし、その前の区間をヘッダ、後ろの区間をトレーラとする
300   ReDim segStart(1 To kinds.Count)
310   ReDim segEnd(1 To kinds.Count)
320   curStart = 1
330   For p = 1 To n
340       If lastPos(seqKind(p)) > maxLast Then maxLast = lastPos(seqKind(p))
350       If maxLast = p Then
360           segCount = segCount + 1
370           segStart(segCount) = curStart
380           segEnd(segCount) = p
390           curStart = p + 1
          End If
      Next p
400   bodySeg = 1
410   For seg = 2 To segCount
420       If segEnd(seg) - segStart(seg) > segEnd(bodySeg) - segStart(bodySeg) Then bodySeg = seg
      Next seg

      ' 本体: 先頭の種別が繰り返し現れ、他の種別もあれば、先頭の種別で始まるグループとみなす
430   Set bodyKinds = CreateObject("Scripting.Dictionary")
440   For p = segStart(bodySeg) To segEnd(bodySeg)
450       If Not bodyKinds.Exists(seqKind(p)) Then bodyKinds.Add seqKind(p), 0
      Next p
460   starter = seqKind(segStart(bodySeg))
470   groupMode = (bodyKinds.Count > 1 And kinds(starter) > 1)
480   If groupMode Then
490       Set cur = CreateObject("Scripting.Dictionary")
500       Set minC = CreateObject("Scripting.Dictionary")
510       Set maxC = CreateObject("Scripting.Dictionary")
520       For p = segStart(bodySeg) To segEnd(bodySeg) + 1
530           If p > segEnd(bodySeg) Then
540               isBoundary = True
              Else
550               isBoundary = (seqKind(p) = starter And p > segStart(bodySeg))
              End If
560           If isBoundary Then
570               groupCount = groupCount + 1
580               For Each k In bodyKinds.Keys
590                   If k <> starter Then
600                       cnt = 0
610                       If cur.Exists(k) Then cnt = cur(k)
620                       If groupCount = 1 Then
630                           minC(k) = cnt
640                           maxC(k) = cnt
                          Else
650                           If cnt < minC(k) Then minC(k) = cnt
660                           If cnt > maxC(k) Then maxC(k) = cnt
                          End If
                      End If
                  Next k
670               cur.RemoveAll
              End If
680           If p <= segEnd(bodySeg) Then cur(seqKind(p)) = cur(seqKind(p)) + 1
          Next p
690       groupText = starter & " 1" & RPT_LAYER_UNIT
700       For Each k In bodyKinds.Keys
710           If k <> starter Then
720               rangeText = CStr(minC(k))
730               If maxC(k) <> minC(k) Then rangeText = rangeText & RPT_RANGE_SEP & maxC(k)
740               role(k) = ROLE_IN_GROUP
750               roleNote(k) = RPT_GROUP_PER & rangeText & RPT_LAYER_UNIT & ")"
760               groupText = groupText & RPT_ARROW & k & " " & rangeText & RPT_LAYER_UNIT
              End If
          Next k
770       role(starter) = ROLE_GROUP_HEAD
      Else
780       For Each k In bodyKinds.Keys
790           role(k) = ROLE_DETAIL
          Next k
      End If

      ' 層ごとの行
800   For seg = 1 To segCount
810       Set segKinds = CreateObject("Scripting.Dictionary")
820       For p = segStart(seg) To segEnd(seg)
830           If Not segKinds.Exists(seqKind(p)) Then segKinds.Add seqKind(p), 0
          Next p
840       If seg < bodySeg Then
850           segRole = ROLE_HEADER
          ElseIf seg > bodySeg Then
860           segRole = ROLE_TRAILER
          Else
870           segRole = ROLE_DETAIL
          End If
880       For Each k In segKinds.Keys
890           If seg <> bodySeg Then role(k) = segRole
          Next k
900       cnt = segEnd(seg) - segStart(seg) + 1
910       layerText = layerText & RPT_LAYER & seg & RPT_LAYER_ROWS & seqRow(segStart(seg)) & _
                      IIf(segEnd(seg) > segStart(seg), RPT_RANGE_SEP & seqRow(segEnd(seg)), "") & _
                      RPT_LAYER_KINDS & Join(segKinds.Keys, " ") & RPT_LAYER_ROLE & segRole & " ("
920       If seg = bodySeg And groupMode Then
930           layerText = layerText & starter & RPT_LAYER_GROUPS_1 & groupCount & RPT_LAYER_GROUPS_2 & cnt & RPT_LAYER_UNIT & ")" & vbCrLf & _
                          RPT_GROUP_BODY & groupText & vbCrLf
940           orderText = orderText & IIf(Len(orderText) > 0, RPT_ARROW, "") & "(" & groupText & ")" & RPT_LAYER_TIMES & groupCount
          Else
950           layerText = layerText & cnt & RPT_LAYER_UNIT & ")" & vbCrLf
960           orderText = orderText & IIf(Len(orderText) > 0, RPT_ARROW, "") & Join(segKinds.Keys, " ") & _
                          " (" & segRole & IIf(cnt > 1, RPT_LAYER_TIMES & cnt, "") & ")"
          End If
      Next seg

970   rep = RPT_SEC_LAYER & vbCrLf & layerText & RPT_LAYER_ORDER & orderText & vbCrLf & RPT_LAYER_HOWTO & vbCrLf & vbCrLf

980   rep = rep & RPT_SEC_RECTYPE & vbCrLf & RPT_RECTYPE_HEAD & vbCrLf & RPT_RECTYPE_RULE & vbCrLf
990   For Each k In kinds.Keys
1000      widthText = ""
1010      Set widthCounts = kindWidths(k)
1020      For Each w In widthCounts.Keys
1030          widthText = widthText & IIf(Len(widthText) > 0, ", ", "") & w & _
                          IIf(widthCounts.Count > 1, " (" & widthCounts(w) & RPT_DATA_UNIT & ")", "")
          Next w
1040      rep = rep & "| " & Replace(k, "|", "/") & " | " & role(k) & roleNote(k) & " | " & kinds(k) & " | " & _
                widthText & " | " & kindFirst(k) & " |" & vbCrLf
      Next k
1050  rep = rep & vbCrLf & RPT_RECTYPE_SEQ & Trim$(seq) & IIf(rowCount > RECTYPE_SEQ_SHOW, " ...", "") & vbCrLf & vbCrLf

      ' ---- 種別ごとの項目一覧 ----
1060  For Each k In kinds.Keys
1070      Set rowColl = kindRows(k)
1080      cnt = rowColl.Count
1090      ReDim idx(1 To cnt)
1100      maxW = 0
1110      i = 0
1120      For Each rowItem In rowColl
1130          i = i + 1
1140          idx(i) = rowItem
1150          If structW(idx(i)) > maxW Then maxW = structW(idx(i))
          Next rowItem
1160      rep = rep & RPT_RECTYPE_SEC & k & " (" & role(k) & ")" & RPT_RECTYPE_SEC_CLOSE & vbCrLf
1170      If cnt = 1 Then rep = rep & RPT_SINGLE_ROW_NOTE & vbCrLf
1180      rep = rep & RPT_TABLE_HEAD_GRID & vbCrLf & RPT_TABLE_RULE & vbCrLf
1190      ReDim relNames(1 To maxW)
1200      For c = 1 To maxW
1210          relNames(c) = REL_NAME_PREFIX & c & " " & IIf(c = 1, RPT_RECTYPE_KEY_LABEL, RPT_ITEM & c & RPT_ITEM_CLOSE)
1220          ReDim values(1 To cnt)
1230          For i = 1 To cnt
1240              If c <= colCount Then values(i) = grid(idx(i), c)
              Next i
1250          rep = rep & "| " & c & " | " & c & " | " & IIf(c = 1, RPT_RECTYPE_KEY_LABEL, RPT_ITEM & c & RPT_ITEM_CLOSE) & " | " & _
                    ProfileValues(values, cnt, "") & vbCrLf
          Next c
1260      rep = rep & vbCrLf & RelationFinder.FindRelations(grid, idx, cnt, maxW, relNames)
      Next k

1270  summaryOut = RPT_MODE_RECTYPE & RPT_SUM_KINDS & kinds.Count & RPT_SUM_KINDS_UNIT & Trim$(Join(kinds.Keys, " ")) & ")" & _
                   SUM_ORDER & orderText
1280  AnalyzeRecordTypes = rep

      LogSuccess MODULE_NAME, "AnalyzeRecordTypes", "Analyzed " & kinds.Count & " record types over " & rowCount & " rows"
1290  Exit Function
ErrHandler:
      LogError MODULE_NAME, "AnalyzeRecordTypes", Err.Number, Err.Description, Erl, _
               "r=" & r & "; c=" & c & "; i=" & i & "; p=" & p & "; segCount=" & segCount
      summaryOut = SUM_ERROR
      AnalyzeRecordTypes = rep & RPT_ERROR_LINE & vbCrLf
End Function

'----------------------------------------------------------------------
' ProfileValues: 1 項目分の値を集計し、項目一覧の 1 行の後半を返す
'   "種別 | 桁数 | 空 | 種類 | 書式パターン | 備考 |"
'   values   値 (1 始まり)。Excel の型付きの値と文字列のどちらも扱う
'   n        値の数
'   extra    呼び出し元が付け足す備考
'   値は外に出さず、型・長さ・パターンだけを出す
'----------------------------------------------------------------------
Public Function ProfileValues(ByRef values() As Variant, ByVal n As Long, ByVal extra As String) As String
      On Error GoTo ErrHandler
      Dim i As Long
      Dim p As Long
      Dim v As Variant
      Dim s As String
      Dim typedCls As String
      Dim cls As String
      Dim emptyN As Long
      Dim blankN As Long
      Dim errN As Long
      Dim filledN As Long
      Dim minLen As Long
      Dim maxLen As Long
      Dim minBytes As Long
      Dim maxBytes As Long
      Dim byteLen As Long
      Dim lenS As Long
      Dim code As Long
      Dim ch As String
      Dim mask As String
      Dim core As String
      Dim pat As String
      Dim prevCh As String
      Dim runLen As Long
      Dim hasWide As Boolean
      Dim hasKana As Boolean
      Dim hasAlpha As Boolean
      Dim isLong As Boolean
      Dim longN As Long
      Dim leadZero As Long
      Dim trailSpace As Long
      Dim leadSpace As Long
      Dim newlineN As Long
      Dim decPlaces As Long
      Dim y As Long
      Dim m As Long
      Dim d As Long
      Dim clsDict As Object
      Dim patDict As Object
      Dim valDict As Object
      Dim distinctOver As Boolean
      Dim k As Variant
      Dim topCls As String
      Dim topN As Long
      Dim secondN As Long
      Dim typeText As String
      Dim lenText As String
      Dim distinctText As String
      Dim patText As String
      Dim shown As Long
      Dim bestKey As String
      Dim bestN As Long
      Dim usedPat As Object
      Dim rmk As String
      Dim mixDetail As String
      Dim nameLike As Long
      Dim warekiYm As Long
      Dim warekiYmd As Long
      Dim era As Long
      Dim pctLike As Long

10    Set clsDict = CreateObject("Scripting.Dictionary")
20    Set patDict = CreateObject("Scripting.Dictionary")
30    Set valDict = CreateObject("Scripting.Dictionary")
40    minLen = -1
50    minBytes = -1

60    For i = 1 To n
70        v = values(i)
80        typedCls = ""
90        s = ""
100       If IsError(v) Then
110           errN = errN + 1
          ElseIf IsEmpty(v) Or IsNull(v) Then
120           emptyN = emptyN + 1
          Else
              ' ---- 型付きの値 (Excel) は文字列にしてから同じ処理に回す ----
130           If VarType(v) = vbDate Then
140               If v = Int(v) Then
150                   s = Format(v, "yyyy/mm/dd")
160                   typedCls = CLS_DATE
                  ElseIf Int(v) = 0 Then
170                   s = Format(v, "hh:nn:ss")
180                   typedCls = CLS_TIME
                  Else
190                   s = Format(v, "yyyy/mm/dd hh:nn:ss")
200                   typedCls = CLS_DATETIME
                  End If
              ElseIf VarType(v) = vbBoolean Then
210               s = CStr(v)
220               typedCls = CLS_BOOL
              ElseIf VarType(v) <> vbString And IsNumeric(v) Then
230               If v = Int(v) Then
240                   s = Format(v, "0")
250                   typedCls = CLS_NUMBER_INT
                  Else
260                   s = CStr(v)
270                   typedCls = CLS_NUMBER_DEC
                  End If
              Else
280               s = CStr(v)
              End If

290           If Len(s) = 0 Then
300               emptyN = emptyN + 1
              ElseIf Len(Trim$(Replace(s, ChrW(CODE_WIDE_SPACE), " "))) = 0 Then
310               blankN = blankN + 1
              Else
320               filledN = filledN + 1
330               If Right$(s, 1) = " " Or Right$(s, 1) = ChrW(CODE_WIDE_SPACE) Then trailSpace = trailSpace + 1
340               If Left$(s, 1) = " " Or Left$(s, 1) = ChrW(CODE_WIDE_SPACE) Then leadSpace = leadSpace + 1
350               If InStr(s, vbLf) > 0 Or InStr(s, vbCr) > 0 Then newlineN = newlineN + 1

                  ' ---- 文字の種類を記号に置き換える (同時にバイト数を数える) ----
360               lenS = Len(s)
370               isLong = (lenS > PATTERN_SRC_MAX)
380               If isLong Then longN = longN + 1
390               mask = ""
400               byteLen = 0
410               hasWide = False
420               hasKana = False
430               hasAlpha = False
440               For p = 1 To lenS
450                   code = AscW(Mid$(s, p, 1))
460                   If code < 0 Then code = code + UNSIGNED_OFFSET
470                   Select Case code
                          Case CODE_DIGIT_0 To CODE_DIGIT_9
480                           ch = MARK_DIGIT
490                           byteLen = byteLen + 1
                          Case CODE_UPPER_A To CODE_UPPER_Z
500                           ch = MARK_UPPER
510                           byteLen = byteLen + 1
520                           hasAlpha = True
                          Case CODE_LOWER_A To CODE_LOWER_Z
530                           ch = MARK_LOWER
540                           byteLen = byteLen + 1
550                           hasAlpha = True
                          Case CODE_SPACE, CODE_TAB
560                           ch = MARK_SPACE
570                           byteLen = byteLen + 1
                          Case CODE_LF, CODE_CR
580                           ch = MARK_NEWLINE
590                           byteLen = byteLen + 1
                          Case CODE_ASCII_FIRST To CODE_ASCII_LAST
600                           ch = Mid$(s, p, 1)
610                           byteLen = byteLen + 1
                          Case Is < CODE_SPACE
620                           ch = MARK_CTRL
630                           byteLen = byteLen + 1
                          Case CODE_HALFKANA_FIRST To CODE_HALFKANA_LAST
640                           ch = ChrW(CODE_HALFKANA_MARK)
650                           byteLen = byteLen + 1
660                           hasKana = True
                          Case CODE_WIDE_DIGIT_0 To CODE_WIDE_DIGIT_9
670                           ch = ChrW(CODE_WIDE_DIGIT_9)
680                           byteLen = byteLen + WIDE_BYTES
690                           hasWide = True
                          Case CODE_WIDE_SPACE
700                           ch = MARK_SPACE & MARK_SPACE
710                           byteLen = byteLen + WIDE_BYTES
720                           hasWide = True
                          Case Else
730                           If InStr(KEEP_CHARS, Mid$(s, p, 1)) > 0 And Right$(mask, 1) = MARK_DIGIT Then
740                               ch = Mid$(s, p, 1)
                              Else
750                               ch = ChrW(CODE_WIDE_MARK)
                              End If
760                           byteLen = byteLen + WIDE_BYTES
770                           hasWide = True
                      End Select
780                   If p <= PATTERN_SRC_MAX Then mask = mask & ch
                  Next p

                  ' 全角の語を全角空白 (__) か半角空白 (_) で区切った形、半角カナの語を空白で区切った形
790               If InStr(mask, ChrW(CODE_WIDE_MARK) & MARK_SPACE & MARK_SPACE & ChrW(CODE_WIDE_MARK)) > 0 _
                     Or InStr(mask, ChrW(CODE_WIDE_MARK) & MARK_SPACE & ChrW(CODE_WIDE_MARK)) > 0 _
                     Or InStr(mask, ChrW(CODE_HALFKANA_MARK) & MARK_SPACE & ChrW(CODE_HALFKANA_MARK)) > 0 Then
800                   nameLike = nameLike + 1
                  End If

810               If minLen < 0 Or lenS < minLen Then minLen = lenS
820               If lenS > maxLen Then maxLen = lenS
830               If minBytes < 0 Or byteLen < minBytes Then minBytes = byteLen
840               If byteLen > maxBytes Then maxBytes = byteLen

                  ' ---- 型の判定 (数字を除いた残りの記号で決める) ----
850               If Len(typedCls) > 0 Then
860                   cls = typedCls
                  ElseIf isLong Then
870                   cls = IIf(hasWide, CLS_TEXT_WIDE, IIf(hasKana, CLS_TEXT_KANA, CLS_TEXT_ALNUM))
                  Else
880                   core = Replace(mask, MARK_DIGIT, "")
890                   If Len(core) = 0 Then
900                       cls = CLS_DIGITS
910                       If lenS = Len("yyyymmdd") Then
920                           y = CLng(Left$(s, 4))
930                           m = CLng(Mid$(s, 5, 2))
940                           d = CLng(Right$(s, 2))
950                           If y >= YEAR_MIN And y <= YEAR_MAX And m >= 1 And m <= MONTH_MAX And d >= 1 And d <= DAY_MAX Then cls = CLS_DATE8
                          End If
960                       If Left$(s, 1) = "0" And lenS > 1 Then leadZero = leadZero + 1
970                       If lenS <= PCT_MAX_LEN Then
980                           If CLng(s) <= PCT_MAX And CLng(s) Mod PCT_STEP = 0 Then pctLike = pctLike + 1
                          End If
                          ' 和暦 GYYMM / GYYMMDD (元号 1-5、月 01-12、日 01-31)
990                       If lenS = WAREKI_YM_LEN Or lenS = WAREKI_YMD_LEN Then
1000                          era = CLng(Left$(s, 1))
1010                          m = CLng(Mid$(s, 4, 2))
1020                          If era >= WAREKI_ERA_MIN And era <= WAREKI_ERA_MAX And m >= 1 And m <= MONTH_MAX Then
1030                              If lenS = WAREKI_YM_LEN Then
1040                                  warekiYm = warekiYm + 1
                                  Else
1050                                  d = CLng(Mid$(s, 6, 2))
1060                                  If d >= 1 And d <= DAY_MAX Then warekiYmd = warekiYmd + 1
                                  End If
                              End If
                          End If
                      ElseIf (core = "-" Or core = "+") And Len(mask) > 1 And (Left$(mask, 1) = "-" Or Left$(mask, 1) = "+") Then
1070                      cls = CLS_SIGNED
                      ElseIf (core = "." Or core = "-." Or core = "+.") And Right$(mask, 1) = MARK_DIGIT And InStr(mask, "9.") > 0 Then
1080                      cls = CLS_DECIMAL
1090                      If Len(mask) - InStr(mask, ".") > decPlaces Then decPlaces = Len(mask) - InStr(mask, ".")
                      ElseIf Replace(Replace(Replace(core, ",", ""), ".", ""), "-", "") = "" And InStr(core, ",") > 0 And Left$(mask, 1) <> "," Then
1100                      cls = CLS_NUM_COMMA
                      ElseIf (core = "//" And mask Like "9999/9*/9*") Or (core = "--" And mask Like "9999-9*-9*") _
                             Or (core = ".." And mask Like "9999.9*.9*") Or core = CORE_DATE_JA Then
1110                      cls = CLS_DATE
                      ElseIf (core = "A.." Or core = "A//" Or core = "A--") And mask Like "A9*" Then
1120                      cls = CLS_WAREKI
                      ElseIf (core = ":" Or core = "::") And mask Like "9*:9*" Then
1130                      cls = CLS_TIME
                      ElseIf (core = "//" & MARK_SPACE & "::" Or core = "//" & MARK_SPACE & ":" Or core = "--" & MARK_SPACE & "::" _
                              Or core = "--A::" Or core = CORE_DATETIME_JA_HM Or core = CORE_DATETIME_JA_HMS) Then
1140                      cls = CLS_DATETIME
                      ElseIf hasWide Then
1150                      cls = CLS_TEXT_WIDE
                      ElseIf hasKana Then
1160                      cls = CLS_TEXT_KANA
                      ElseIf hasAlpha Or InStr(mask, MARK_DIGIT) > 0 Then
1170                      cls = CLS_TEXT_ALNUM
                      Else
1180                      cls = CLS_TEXT_SYMBOL
                      End If
                  End If
1190              clsDict(cls) = clsDict(cls) + 1

                  ' ---- 連続する記号を個数表記にまとめる ----
1200              pat = ""
1210              prevCh = ""
1220              runLen = 0
1230              For p = 1 To Len(mask) + 1
1240                  If p <= Len(mask) Then ch = Mid$(mask, p, 1) Else ch = vbNullChar
1250                  If ch = prevCh Then
1260                      runLen = runLen + 1
                      Else
1270                      If runLen >= RUN_COMPRESS_MIN Then
1280                          pat = pat & prevCh & "(" & runLen & ")"
                          ElseIf runLen > 0 Then
1290                          pat = pat & String$(runLen, prevCh)
                          End If
1300                      prevCh = ch
1310                      runLen = 1
                      End If
                  Next p
1320              If isLong Then pat = pat & "..."
1330              If patDict.Exists(pat) Then
1340                  patDict(pat) = patDict(pat) + 1
                  ElseIf patDict.Count < PATTERN_CAP Then
1350                  patDict.Add pat, 1
                  End If

                  ' ---- 種類数 (値は辞書の中だけで使い、外に出さない) ----
1360              If Not distinctOver Then
1370                  If Not valDict.Exists(s) Then
1380                      If valDict.Count >= DISTINCT_CAP Then
1390                          distinctOver = True
                          Else
1400                          valDict.Add s, True
                          End If
                      End If
                  End If
              End If
          End If
      Next i

      ' ---- 型 ----
1410  For Each k In clsDict.Keys
1420      If clsDict(k) > topN Then
1430          secondN = topN
1440          topN = clsDict(k)
1450          topCls = k
          ElseIf clsDict(k) > secondN Then
1460          secondN = clsDict(k)
          End If
      Next k
1470  If filledN = 0 Then
1480      typeText = CLS_EMPTY
      ElseIf secondN >= filledN * MIX_MIN_SHARE Then
1490      typeText = CLS_MIXED & TYPE_MAIN_OPEN & topCls & ")"
1500      For Each k In clsDict.Keys
1510          mixDetail = mixDetail & k & " " & clsDict(k) & ", "
          Next k
1520      mixDetail = RMK_MIX_DETAIL & Left$(mixDetail, Len(mixDetail) - 2)
      Else
1530      typeText = topCls
      End If

      ' ---- 桁数・種類 ----
1540  If filledN = 0 Then
1550      lenText = "-"
1560      distinctText = "0"
      Else
1570      lenText = IIf(minLen = maxLen, CStr(minLen), minLen & "-" & maxLen)
1580      distinctText = IIf(distinctOver, DISTINCT_CAP & "+", CStr(valDict.Count))
      End If

      ' ---- 書式パターン (件数の多い順に上位を載せる) ----
1590  Set usedPat = CreateObject("Scripting.Dictionary")
1600  For shown = 1 To PATTERN_SHOW
1610      bestKey = ""
1620      bestN = 0
1630      For Each k In patDict.Keys
1640          If Not usedPat.Exists(k) And patDict(k) > bestN Then
1650              bestN = patDict(k)
1660              bestKey = k
              End If
          Next k
1670      If bestN = 0 Then Exit For
1680      usedPat.Add bestKey, True
1690      If Len(bestKey) > PATTERN_SHOW_LEN Then bestKey = Left$(bestKey, PATTERN_SHOW_LEN) & "..."
1700      patText = patText & IIf(Len(patText) > 0, ", ", "") & "`" & bestKey & "`" & IIf(patDict.Count > 1, " (" & bestN & ")", "")
      Next shown
1710  If patDict.Count > PATTERN_SHOW Then patText = patText & ", ..."
1720  If Len(patText) = 0 Then patText = "-"

      ' ---- 備考 ----
1730  If n > 0 And filledN = 0 Then rmk = rmk & IIf(Len(rmk) > 0, RMK_SEP, "") & RMK_ALL_EMPTY
1740  If filledN > 0 Then
1750      If Not distinctOver And valDict.Count = 1 Then
              ' 値が 1 つしかないときは、固定値かどうか分からないので出さない
1760          If filledN > 1 Then rmk = rmk & IIf(Len(rmk) > 0, RMK_SEP, "") & RMK_FIXED
          ElseIf Not distinctOver And valDict.Count = filledN And filledN = n And n > 1 _
                 And (topCls = CLS_DIGITS Or topCls = CLS_TEXT_ALNUM Or topCls = CLS_TEXT_WIDE _
                      Or topCls = CLS_TEXT_KANA Or topCls = CLS_NUMBER_INT) Then
1770          rmk = rmk & IIf(Len(rmk) > 0, RMK_SEP, "") & RMK_KEY
          ElseIf Not distinctOver And valDict.Count <= CODE_MAX_KINDS And filledN >= CODE_MIN_ROWS Then
1780          rmk = rmk & IIf(Len(rmk) > 0, RMK_SEP, "") & RMK_CODE & valDict.Count & RMK_CODE_UNIT
          End If
1790      If minBytes = maxBytes And filledN > 1 Then rmk = rmk & IIf(Len(rmk) > 0, RMK_SEP, "") & RMK_FIXLEN & maxBytes & RMK_BYTES_UNIT
1800      If minBytes <> maxBytes And (maxBytes <> maxLen Or minBytes <> minLen) Then
1810          rmk = rmk & IIf(Len(rmk) > 0, RMK_SEP, "") & RMK_BYTES & minBytes & "-" & maxBytes
          End If
1820      If pctLike = filledN And valDict.Count >= 2 And valDict.Count <= PCT_MAX_KINDS Then
1830          rmk = rmk & IIf(Len(rmk) > 0, RMK_SEP, "") & RMK_PCT_LIKE
          End If
1840      If warekiYm = filledN Then rmk = rmk & IIf(Len(rmk) > 0, RMK_SEP, "") & RMK_WAREKI_YM
1850      If warekiYmd = filledN Then rmk = rmk & IIf(Len(rmk) > 0, RMK_SEP, "") & RMK_WAREKI_YMD
1860      If leadZero > 0 Then rmk = rmk & IIf(Len(rmk) > 0, RMK_SEP, "") & RMK_LEAD0 & " " & leadZero & RPT_ITEMS_UNIT
1870      If decPlaces > 0 Then rmk = rmk & IIf(Len(rmk) > 0, RMK_SEP, "") & RMK_DEC_PLACES & decPlaces & RMK_DEC_UNIT
1880      If trailSpace > 0 Then rmk = rmk & IIf(Len(rmk) > 0, RMK_SEP, "") & RMK_TRAIL_SPACE & " " & trailSpace & RPT_ITEMS_UNIT
1890      If leadSpace > 0 Then rmk = rmk & IIf(Len(rmk) > 0, RMK_SEP, "") & RMK_LEAD_SPACE & " " & leadSpace & RPT_ITEMS_UNIT
1900      If newlineN > 0 Then rmk = rmk & IIf(Len(rmk) > 0, RMK_SEP, "") & RMK_NEWLINE & " " & newlineN & RPT_ITEMS_UNIT
1910      If longN > 0 Then rmk = rmk & IIf(Len(rmk) > 0, RMK_SEP, "") & RMK_LONG
1920      If nameLike >= filledN * NAME_LIKE_SHARE Then rmk = rmk & IIf(Len(rmk) > 0, RMK_SEP, "") & RMK_NAME_LIKE
      End If
1930  If blankN > 0 Then rmk = rmk & IIf(Len(rmk) > 0, RMK_SEP, "") & RMK_BLANK_ONLY & blankN & RPT_ITEMS_UNIT
1940  If errN > 0 Then rmk = rmk & IIf(Len(rmk) > 0, RMK_SEP, "") & RMK_ERROR & errN & RPT_ITEMS_UNIT
1950  If Len(mixDetail) > 0 Then rmk = rmk & IIf(Len(rmk) > 0, RMK_SEP, "") & mixDetail
1960  If Len(extra) > 0 Then rmk = rmk & IIf(Len(rmk) > 0, RMK_SEP, "") & extra

1970  ProfileValues = typeText & " | " & lenText & " | " & (emptyN + blankN) & " | " & distinctText & " | " & _
                      Replace(patText, "|", "/") & " | " & Replace(rmk, "|", "/") & " |"

      LogSuccess MODULE_NAME, "ProfileValues", "Profiled " & n & " values; type=" & typeText
1980  Exit Function
ErrHandler:
      LogError MODULE_NAME, "ProfileValues", Err.Number, Err.Description, Erl, _
               "n=" & n & "; i=" & i & "; p=" & p
      ProfileValues = PROFILE_ERROR
End Function

'----------------------------------------------------------------------
' MaskDigits: 文字列の数字 (半角・全角) を 9 に置き換える (表題用)
'----------------------------------------------------------------------
Public Function MaskDigits(ByVal text As String) As String
      On Error GoTo ErrHandler
      Dim p As Long
      Dim code As Long
      Dim out As String

10    For p = 1 To Len(text)
20        code = AscW(Mid$(text, p, 1))
30        If code < 0 Then code = code + UNSIGNED_OFFSET
40        If (code >= CODE_DIGIT_0 And code <= CODE_DIGIT_9) Or (code >= CODE_WIDE_DIGIT_0 And code <= CODE_WIDE_DIGIT_9) Then
50            out = out & MARK_DIGIT
          Else
60            out = out & Mid$(text, p, 1)
          End If
      Next p
70    MaskDigits = out

      LogSuccess MODULE_NAME, "MaskDigits", "Masked digits in text of length " & Len(text)
80    Exit Function
ErrHandler:
      LogError MODULE_NAME, "MaskDigits", Err.Number, Err.Description, Erl, "p=" & p
      MaskDigits = ""
End Function

'----------------------------------------------------------------------
' ColumnRef: 列の位置の表記 (区切り文字形式は番号、Excel は列記号)
'----------------------------------------------------------------------
Public Function ColumnRef(ByVal colNo As Long, ByVal isText As Boolean) As String
      On Error GoTo ErrHandler
      Dim n As Long
      Dim letters As String

10    If isText Then
20        ColumnRef = CStr(colNo)
      Else
30        n = colNo
40        Do While n > 0
50            letters = Chr$(CODE_UPPER_A + ((n - 1) Mod ALPHABET_SIZE)) & letters
60            n = (n - 1) \ ALPHABET_SIZE
          Loop
70        ColumnRef = letters
      End If

      LogSuccess MODULE_NAME, "ColumnRef", "Converted column " & colNo & " to " & ColumnRef
80    Exit Function
ErrHandler:
      LogError MODULE_NAME, "ColumnRef", Err.Number, Err.Description, Erl, "colNo=" & colNo
      ColumnRef = CStr(colNo)
End Function

'----------------------------------------------------------------------
' ListText: 行番号の一覧に件数を添える (0 件なら「なし」)
'----------------------------------------------------------------------
Private Function ListText(ByVal listItems As String, ByVal itemCount As Long) As String
      On Error GoTo ErrHandler

10    If itemCount = 0 Then
20        ListText = RPT_NONE
      Else
30        ListText = itemCount & RPT_ITEMS_UNIT & RPT_LIST_OPEN & Trim$(listItems) & IIf(itemCount > MAX_LIST_ITEMS, RPT_MORE, "") & ")"
      End If

      LogSuccess MODULE_NAME, "ListText", "Formatted list of " & itemCount & " rows"
40    Exit Function
ErrHandler:
      LogError MODULE_NAME, "ListText", Err.Number, Err.Description, Erl, "itemCount=" & itemCount
      ListText = CStr(itemCount)
End Function
