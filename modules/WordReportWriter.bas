Attribute VB_Name = "WordReportWriter"
Option Explicit

'======================================================================
' WordReportWriter モジュール
' テキストの報告を、読みやすい Word 文書 (.docx) にして保存する。
'   - 報告文を HTML に変換し、Word の白紙の文書に読み込んで docx で保存する
'   - 先頭の行は表題、[..] の行は見出し、"項目 : 値" の行は 2 列の表、
'     "| .. |" の行は罫線付きの表にする (書式パターンは等幅フォント)
'   - Excel のシート ("----" の次の行) は見出し 1、その中の [..] は見出し 2
'   - 用紙は A4 横 (項目一覧の列が多いため)。表の見出し行はページごとに繰り返す
' Word は遅延バインディングで使う (参照設定は不要)。
' 依存: ErrorLogger.bas
'======================================================================

Private Const MODULE_NAME As String = "WordReportWriter"

' ---- Word ----
Private Const WORD_PROG_ID As String = "Word.Application"
Private Const WD_ALERTS_NONE As Long = 0
Private Const WD_DO_NOT_SAVE As Long = 0
Private Const WD_FORMAT_DOCX As Long = 16          ' wdFormatXMLDocument
Private Const WD_PAPER_A4 As Long = 7              ' wdPaperA4
Private Const WD_ORIENT_LANDSCAPE As Long = 1      ' wdOrientLandscape
Private Const WD_FOOTER_PRIMARY As Long = 1        ' wdHeaderFooterPrimary
Private Const WD_PAGE_NUMBER_CENTER As Long = 1    ' wdAlignPageNumberCenter
Private Const PAGE_MARGIN_CM As Single = 1.5
Private Const KV_TABLE_COLS As Long = 2            ' これより列の多い表だけ見出し行を繰り返す
Private Const WD_AUTOFIT_CONTENT As Long = 1       ' wdAutoFitContent
Private Const WD_LAYOUT_NO_GRID As Long = 0        ' wdLayoutModeDefault (文字グリッドに合わせない)

' ---- 一時ファイル ----
Private Const TEMP_FOLDER_ID As Long = 2           ' FileSystemObject の TemporaryFolder
Private Const TEMP_EXT As String = ".htm"
Private Const PATH_SEP As String = "\"
Private Const HTML_CHARSET As String = "UTF-8"
Private Const AD_TYPE_TEXT As Long = 2
Private Const AD_SAVE_CREATE_OVERWRITE As Long = 2
Private Const AD_STATE_CLOSED As Long = 0

' ---- 報告文の読み取り ----
Private Const RULE_MAIN_CHAR As String = "="
Private Const RULE_SUB_CHAR As String = "-"
Private Const SEC_OPEN As String = "["
Private Const SEC_CLOSE As String = "]"
Private Const CELL_SEP As String = "|"
Private Const ALIGN_MARK As String = ":"
Private Const KV_SEP As String = " : "
Private Const PAREN_OPEN As String = "("
Private Const PAREN_CLOSE As String = ")"
Private Const CODE_MARK As String = "`"
Private Const INDENT_STEP As Long = 2              ' 報告文の字下げ 1 段の空白数
Private Const INDENT_PT As Long = 12               ' Word での字下げ 1 段の幅 (pt)
Private Const TABLE_NONE As Long = 0
Private Const TABLE_KV As Long = 1
Private Const TABLE_GRID As Long = 2

' ---- HTML ----
Private Const HTML_CSS As String = _
    "body{font-family:'Meiryo UI';mso-fareast-font-family:'Meiryo UI';font-size:10pt}" & _
    "p{margin:0 0 2pt 0}" & _
    "p.MsoTitle{font-size:20pt;font-weight:bold;color:#1F3864;margin:0 0 8pt 0}" & _
    "h1{font-size:14pt;color:#1F3864;margin:16pt 0 4pt 0;border-bottom:1.5pt solid #2F5496}" & _
    "h2{font-size:12pt;color:#2F5496;margin:12pt 0 4pt 0}" & _
    "table{border-collapse:collapse;margin:2pt 0 6pt 0}" & _
    "th,td{border:0.5pt solid #A6A6A6;padding:1pt 4pt;font-size:9pt;vertical-align:top;text-align:left}" & _
    "th{background:#D9E2F3;font-weight:bold;text-align:left}" & _
    "td.k{background:#F2F2F2;font-weight:bold}" & _
    "td.r,th.r{text-align:right}" & _
    "span.c{font-family:Consolas;mso-fareast-font-family:'MS Gothic'}"
Private Const HTML_HEAD As String = "<html><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'><style>"
Private Const HTML_BODY As String = "</style></head><body>"
Private Const HTML_END As String = "</body></html>"
Private Const TAG_TITLE As String = "<p class='MsoTitle'>"
Private Const TAG_P As String = "<p>"
Private Const TAG_P_INDENT As String = "<p style='margin-left:"
Private Const TAG_P_INDENT_CLOSE As String = "pt'>"
Private Const TAG_P_END As String = "</p>"
Private Const TAG_H1 As String = "h1"
Private Const TAG_H2 As String = "h2"
Private Const TAG_OPEN As String = "<"
Private Const TAG_CLOSE_OPEN As String = "</"
Private Const TAG_CLOSE As String = ">"
Private Const TAG_KV_TABLE As String = "<table>"
Private Const TAG_GRID_TABLE As String = "<table class='g'>"
Private Const TAG_TABLE_END As String = "</table>"
Private Const TAG_TR As String = "<tr>"
Private Const TAG_TR_END As String = "</tr>"
Private Const TAG_TH As String = "<th>"
Private Const TAG_TH_RIGHT As String = "<th class='r'>"
Private Const TAG_TH_END As String = "</th>"
Private Const TAG_TD As String = "<td>"
Private Const TAG_TD_RIGHT As String = "<td class='r'>"
Private Const TAG_TD_KEY As String = "<td class='k'>"
Private Const TAG_TD_END As String = "</td>"
Private Const TAG_CODE As String = "<span class='c'>"
Private Const TAG_CODE_END As String = "</span>"
Private Const HTML_NBSP As String = "&nbsp;"
Private Const HTML_AMP As String = "&amp;"
Private Const HTML_LT As String = "&lt;"
Private Const HTML_GT As String = "&gt;"
Private Const DOUBLE_SPACE As String = "  "


'----------------------------------------------------------------------
' WriteWordReport: 報告文を Word 文書にして docPath に保存する (既存は上書き)
'   保存できたら True。Word を起動できないときなどは False を返す
'   (呼び出し元はテキストで出力し直す)
'----------------------------------------------------------------------
Public Function WriteWordReport(ByVal docPath As String, ByVal reportText As String) As Boolean
      On Error GoTo ErrHandler
      Dim fso As Object
      Dim htmlPath As String
      Dim html As String
      Dim stm As Object
      Dim wordApp As Object
      Dim doc As Object
      Dim tbl As Object
      Dim tableCount As Long
      Dim prevAlerts As Long
      Dim alertsChanged As Boolean

      ' ---- HTML を一時フォルダに書き出す ----
10    html = BuildReportHtml(reportText)
20    Set fso = CreateObject("Scripting.FileSystemObject")
30    htmlPath = fso.GetSpecialFolder(TEMP_FOLDER_ID).Path & PATH_SEP & fso.GetBaseName(fso.GetTempName()) & TEMP_EXT
40    Set stm = CreateObject("ADODB.Stream")
50    stm.Type = AD_TYPE_TEXT
60    stm.Charset = HTML_CHARSET
70    stm.Open
80    stm.WriteText html
90    stm.SaveToFile htmlPath, AD_SAVE_CREATE_OVERWRITE
100   stm.Close

      ' ---- Word で白紙の文書に読み込み、整えて保存する ----
      ' 既に開いている Word に繋がった場合は、利用者の画面を隠さず、文書だけ非表示で作る
110   Set wordApp = CreateObject(WORD_PROG_ID)
120   prevAlerts = wordApp.DisplayAlerts
130   alertsChanged = True
140   wordApp.DisplayAlerts = WD_ALERTS_NONE
150   Set doc = wordApp.Documents.Add(Visible:=False)
160   doc.Content.InsertFile FileName:=htmlPath, ConfirmConversions:=False

170   With doc.PageSetup
180       .PaperSize = WD_PAPER_A4
190       .Orientation = WD_ORIENT_LANDSCAPE
200       .TopMargin = wordApp.CentimetersToPoints(PAGE_MARGIN_CM)
210       .BottomMargin = wordApp.CentimetersToPoints(PAGE_MARGIN_CM)
220       .LeftMargin = wordApp.CentimetersToPoints(PAGE_MARGIN_CM)
230       .RightMargin = wordApp.CentimetersToPoints(PAGE_MARGIN_CM)
240       .LayoutMode = WD_LAYOUT_NO_GRID
      End With

      ' 表: セル内の段落の間隔を詰め、行の途中で改ページしない。列幅は内容に合わせる (用紙の幅まで)。
      ' 列の多い表は、見出し行をページごとに繰り返す
250   For Each tbl In doc.Tables
260       tableCount = tableCount + 1
270       tbl.Range.ParagraphFormat.SpaceBefore = 0
280       tbl.Range.ParagraphFormat.SpaceAfter = 0
290       tbl.Rows.AllowBreakAcrossPages = False
300       tbl.AutoFitBehavior WD_AUTOFIT_CONTENT
310       If tbl.Columns.Count > KV_TABLE_COLS Then tbl.Rows(1).HeadingFormat = True
      Next tbl

320   doc.Sections(1).Footers(WD_FOOTER_PRIMARY).PageNumbers.Add WD_PAGE_NUMBER_CENTER
330   doc.SaveAs2 FileName:=docPath, FileFormat:=WD_FORMAT_DOCX
340   WriteWordReport = True

      LogSuccess MODULE_NAME, "WriteWordReport", "Wrote Word report (" & tableCount & " tables, " & _
                 Len(reportText) & " chars) to " & docPath

CleanUp:
      On Error Resume Next
      If Not stm Is Nothing Then
          If stm.State <> AD_STATE_CLOSED Then stm.Close
          Set stm = Nothing
      End If
      If Not doc Is Nothing Then
          doc.Close SaveChanges:=WD_DO_NOT_SAVE
          Set doc = Nothing
      End If
      If Not wordApp Is Nothing Then
          If alertsChanged Then wordApp.DisplayAlerts = prevAlerts
          ' 既に開いていた Word に繋がった場合は、利用者の文書を閉じないよう終了しない
          If wordApp.Documents.Count = 0 Then wordApp.Quit SaveChanges:=WD_DO_NOT_SAVE
          Set wordApp = Nothing
      End If
      If Not fso Is Nothing Then
          If Len(htmlPath) > 0 Then
              If fso.FileExists(htmlPath) Then fso.DeleteFile htmlPath, True
          End If
          Set fso = Nothing
      End If
      On Error GoTo 0
      Exit Function

ErrHandler:
      LogError MODULE_NAME, "WriteWordReport", Err.Number, Err.Description, Erl, _
               "docPath=" & docPath & "; htmlPath=" & htmlPath & "; tableCount=" & tableCount
      WriteWordReport = False
      Resume CleanUp
End Function

'----------------------------------------------------------------------
' BuildReportHtml: 報告文 (テキスト) を HTML に変換する
'   行の種類ごとに変換する (1 行ずつ。表は行をまたいで組み立てる)
'     先頭の行            表題
'     "====" / "----"     読み飛ばす ("----" の次の行はシート名として見出し 1)
'     [..]                見出し (シートの中なら見出し 2)。"]" の後ろの文は段落
'     | .. |              表。2 行目の |---:| から列の右寄せを決める
'     "項目 : 値"         2 列の表 (続く行をまとめる)
'     その他              段落 (字下げを保つ)
'----------------------------------------------------------------------
Private Function BuildReportHtml(ByVal reportText As String) As String
      On Error GoTo ErrHandler
      Dim lines() As String
      Dim parts() As String
      Dim i As Long
      Dim j As Long
      Dim k As Long
      Dim raw As String
      Dim t As String
      Dim chunk As String
      Dim isPipe As Boolean
      Dim isKv As Boolean
      Dim isRule As Boolean
      Dim tableKind As Long
      Dim pendingHead As String
      Dim cells() As String
      Dim headCells() As String
      Dim aligns() As Boolean
      Dim alignCount As Long
      Dim cellText As String
      Dim pieces() As String
      Dim pos As Long
      Dim lead As Long
      Dim afterSubRule As Boolean
      Dim inSheet As Boolean
      Dim titleDone As Boolean
      Dim headTag As String
      Dim errNo As Long
      Dim errDesc As String

10    lines = Split(Replace(reportText, vbCrLf, vbLf), vbLf)
20    ReDim parts(0 To UBound(lines) + 2)
30    parts(0) = HTML_HEAD & HTML_CSS & HTML_BODY

40    For i = 0 To UBound(lines)
          ' 値は報告に無いが、ラベルや要素名に < > & が含まれうるので最初に置き換える
50        raw = Replace(Replace(Replace(lines(i), "&", HTML_AMP), "<", HTML_LT), ">", HTML_GT)
60        t = Trim$(raw)
70        chunk = ""
80        isPipe = False
90        If Len(t) > 1 Then
100           If Left$(t, 1) = CELL_SEP And Right$(t, 1) = CELL_SEP Then isPipe = True
          End If
110       isKv = False
120       If Not isPipe And titleDone And Left$(t, 1) <> SEC_OPEN Then
130           pos = InStr(raw, KV_SEP)
140           If pos > 0 Then
                  ' 項目名の括弧が閉じていなければ、文の中の " : " とみなす (凡例の "記号 (/ - : ." など)
150               isKv = (InStr(Left$(raw, pos), PAREN_OPEN) = 0 Or InStr(Left$(raw, pos), PAREN_CLOSE) > 0)
              End If
          End If

          ' ---- 表の終わり ----
160       If tableKind = TABLE_GRID And Not isPipe Then
170           If Len(pendingHead) > 0 Then
                  ' 見出し行だけの表 (区切りの行が無い)
180               headCells = Split(Mid$(pendingHead, 2, Len(pendingHead) - 2), CELL_SEP)
190               chunk = chunk & TAG_GRID_TABLE & TAG_TR
200               For j = 0 To UBound(headCells)
210                   chunk = chunk & TAG_TH & Trim$(headCells(j)) & TAG_TH_END
                  Next j
220               chunk = chunk & TAG_TR_END
230               pendingHead = ""
              End If
240           chunk = chunk & TAG_TABLE_END
250           tableKind = TABLE_NONE
          ElseIf tableKind = TABLE_KV And Not isKv Then
260           chunk = chunk & TAG_TABLE_END
270           tableKind = TABLE_NONE
          End If

280       If Len(t) = 0 Then
              ' 空行は段落の間隔で表すので出さない
          ElseIf isPipe Then
290           cells = Split(Mid$(t, 2, Len(t) - 2), CELL_SEP)
300           isRule = (Len(Replace(Replace(Replace(t, CELL_SEP, ""), RULE_SUB_CHAR, ""), ALIGN_MARK, "")) = 0)
310           If tableKind <> TABLE_GRID Then
                  ' 1 行目は見出し。右寄せは次の区切りの行で決まるので、出力を保留する
320               tableKind = TABLE_GRID
330               pendingHead = t
340               alignCount = 0
              ElseIf isRule And Len(pendingHead) > 0 Then
350               alignCount = UBound(cells) + 1
360               ReDim aligns(0 To UBound(cells))
370               For j = 0 To UBound(cells)
380                   aligns(j) = (Right$(Trim$(cells(j)), 1) = ALIGN_MARK)
                  Next j
390               headCells = Split(Mid$(pendingHead, 2, Len(pendingHead) - 2), CELL_SEP)
400               chunk = chunk & TAG_GRID_TABLE & TAG_TR
410               For j = 0 To UBound(headCells)
420                   cellText = TAG_TH
430                   If j < alignCount Then
440                       If aligns(j) Then cellText = TAG_TH_RIGHT
                      End If
450                   chunk = chunk & cellText & Trim$(headCells(j)) & TAG_TH_END
                  Next j
460               chunk = chunk & TAG_TR_END
470               pendingHead = ""
              Else
480               If Len(pendingHead) > 0 Then
                      ' 区切りの行が無い表: 見出し行をそのまま出す
490                   headCells = Split(Mid$(pendingHead, 2, Len(pendingHead) - 2), CELL_SEP)
500                   chunk = chunk & TAG_GRID_TABLE & TAG_TR
510                   For j = 0 To UBound(headCells)
520                       chunk = chunk & TAG_TH & Trim$(headCells(j)) & TAG_TH_END
                      Next j
530                   chunk = chunk & TAG_TR_END
540                   pendingHead = ""
                  End If
550               chunk = chunk & TAG_TR
560               For j = 0 To UBound(cells)
570                   cellText = Trim$(cells(j))
                      ' `..` で囲んだ書式パターンは等幅フォントにする
580                   If InStr(cellText, CODE_MARK) > 0 Then
590                       pieces = Split(cellText, CODE_MARK)
600                       cellText = ""
610                       For k = 0 To UBound(pieces)
620                           If k Mod 2 = 1 Then
630                               cellText = cellText & TAG_CODE & pieces(k) & TAG_CODE_END
                              Else
640                               cellText = cellText & pieces(k)
                              End If
                          Next k
                      End If
650                   If Len(cellText) = 0 Then cellText = HTML_NBSP
660                   If j < alignCount Then
670                       If aligns(j) Then
680                           chunk = chunk & TAG_TD_RIGHT & cellText & TAG_TD_END
                          Else
690                           chunk = chunk & TAG_TD & cellText & TAG_TD_END
                          End If
                      Else
700                       chunk = chunk & TAG_TD & cellText & TAG_TD_END
                      End If
                  Next j
710               chunk = chunk & TAG_TR_END
              End If
          ElseIf Len(Replace(t, RULE_MAIN_CHAR, "")) = 0 Then
              ' 表題の下線
          ElseIf Len(Replace(t, RULE_SUB_CHAR, "")) = 0 Then
              ' シートの区切り。次の行がシート名
720           afterSubRule = True
          ElseIf Not titleDone Then
730           chunk = chunk & TAG_TITLE & t & TAG_P_END
740           titleDone = True
          ElseIf afterSubRule Then
750           chunk = chunk & TAG_OPEN & TAG_H1 & TAG_CLOSE & t & TAG_CLOSE_OPEN & TAG_H1 & TAG_CLOSE
760           afterSubRule = False
770           inSheet = True
          ElseIf Left$(t, 1) = SEC_OPEN And InStr(t, SEC_CLOSE) > 0 Then
780           pos = InStr(t, SEC_CLOSE)
790           headTag = IIf(inSheet, TAG_H2, TAG_H1)
800           chunk = chunk & TAG_OPEN & headTag & TAG_CLOSE & Mid$(t, 2, pos - 2) & TAG_CLOSE_OPEN & headTag & TAG_CLOSE
810           If Len(Trim$(Mid$(t, pos + 1))) > 0 Then chunk = chunk & TAG_P & Trim$(Mid$(t, pos + 1)) & TAG_P_END
          ElseIf isKv Then
820           If tableKind <> TABLE_KV Then
830               chunk = chunk & TAG_KV_TABLE
840               tableKind = TABLE_KV
              End If
850           pos = InStr(raw, KV_SEP)
860           chunk = chunk & TAG_TR & TAG_TD_KEY & Trim$(Left$(raw, pos - 1)) & TAG_TD_END & _
                      TAG_TD & Trim$(Mid$(raw, pos + Len(KV_SEP))) & TAG_TD_END & TAG_TR_END
          Else
              ' 字下げ 2 段目からは左の余白で表す。語の間の 2 つ続きの空白は保つ
870           lead = (Len(raw) - Len(LTrim$(raw))) \ INDENT_STEP
880           If lead > 1 Then
890               chunk = chunk & TAG_P_INDENT & (lead - 1) * INDENT_PT & TAG_P_INDENT_CLOSE
              Else
900               chunk = chunk & TAG_P
              End If
910           chunk = chunk & Replace(t, DOUBLE_SPACE, HTML_NBSP & " ") & TAG_P_END
          End If
920       parts(i + 1) = chunk
      Next i

930   If tableKind <> TABLE_NONE Then parts(UBound(parts)) = TAG_TABLE_END
940   parts(UBound(parts)) = parts(UBound(parts)) & HTML_END
950   BuildReportHtml = Join(parts, vbCrLf)

      LogSuccess MODULE_NAME, "BuildReportHtml", "Converted " & (UBound(lines) + 1) & " report lines to HTML"
960   Exit Function
ErrHandler:
      ' LogError は Err を消すので、先に控えておく
      errNo = Err.Number
      errDesc = Err.Description
      LogError MODULE_NAME, "BuildReportHtml", errNo, errDesc, Erl, _
               "lineIndex=" & i & "; tableKind=" & tableKind
      Err.Raise errNo, MODULE_NAME & ".BuildReportHtml", errDesc
End Function
