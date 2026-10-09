Attribute VB_Name = "XmlAnalyzer"
Option Explicit

'======================================================================
' XmlAnalyzer モジュール
' XML ファイルの構造を、要素と属性のパスごとにまとめる。
'   - パス (/root/item/name、/root/item/@id) ごとの出現数
'   - 子要素を持つ要素か、値を持つ要素か、属性か
'   - 繰り返し要素 (レコードの候補)
'   - 値の型・桁数・書式パターン (GridAnalyzer.ProfileValues を使う)
'
' DTD は読み込まない (外部実体の読み込みを防ぐため)。
' 依存: ErrorLogger.bas, GridAnalyzer.bas
'======================================================================

Private Const MODULE_NAME As String = "XmlAnalyzer"

Private Const MAX_NODES As Long = 500000          ' 走査する最大の要素数
Private Const MAX_VALUES_PER_PATH As Long = 100000 ' パスごとに集計する最大の値の数
Private Const NODE_ELEMENT As Long = 1
Private Const NODE_TEXT As Long = 3
Private Const NODE_CDATA As Long = 4
Private Const PATH_SEP As String = "/"
Private Const ATTR_MARK As String = "@"

' ---- 報告の文言 ----
Private Const RPT_SEC_STRUCT As String = "[構造]"
Private Const RPT_SEC_ITEMS As String = "[項目一覧 (要素と属性のパス)]"
Private Const RPT_FORMAT As String = "  形式            : XML"
Private Const RPT_ROOT As String = "  ルート要素      : "
Private Const RPT_NAMESPACE As String = "  名前空間        : "
Private Const RPT_ELEMENTS As String = "  要素の数        : "
Private Const RPT_PATHS As String = "  パスの種類      : "
Private Const RPT_RECORDS As String = "  繰り返し要素 (レコード候補): "
Private Const RPT_TRUNC As String = "  注: 要素が多いため、先頭の一部だけを走査している"
Private Const RPT_PARSE_ERROR As String = "  XML として読み込めない: "
Private Const RPT_PARSE_LINE As String = " (行 "
Private Const RPT_TABLE_HEAD As String = "| No | パス | ラベル | 種別 | 桁数 | 空 | 種類 | 書式パターン (上位3) | 備考 |"
Private Const RPT_TABLE_RULE As String = "|---:|---|---|---|---|---:|---:|---|---|"
Private Const KIND_CONTAINER As String = "要素 (子要素あり)"
Private Const RMK_COUNT As String = "出現 "
Private Const RMK_COUNT_UNIT As String = " 回"
Private Const RMK_REPEAT As String = "繰り返しあり (レコード候補)"
Private Const RMK_ATTR As String = "属性"
Private Const RMK_SEP As String = "; "
Private Const RPT_NONE As String = "なし"
Private Const SUM_XML As String = "XML / ルート "
Private Const SUM_PATHS As String = " / パス "
Private Const SUM_RECORDS As String = " / 繰り返し要素 "
Private Const SUM_PARSE_ERROR As String = "XML として読み込めない"
Private Const SUM_ERROR As String = "解析中にエラー (ログを確認)"
Private Const RPT_ERROR_LINE As String = "  解析中にエラーが発生した。C:\VBAErrorLogs\ を確認すること。"


'----------------------------------------------------------------------
' AnalyzeXml: XML ファイルを解析し、報告文を返す
'   parsed  XML として読み込めたら True (False なら呼び出し元はテキストとして解析する)
'----------------------------------------------------------------------
Public Function AnalyzeXml(ByVal filePath As String, ByRef parsed As Boolean, ByRef summaryOut As String) As String
      On Error GoTo ErrHandler
      Dim rep As String
      Dim doc As Object
      Dim stackNodes As Collection
      Dim stackPaths As Collection
      Dim node As Object
      Dim child As Object
      Dim attr As Object
      Dim path As String
      Dim childPath As String
      Dim attrPath As String
      Dim pathOrder As Collection
      Dim pathCount As Object
      Dim pathParent As Object
      Dim pathName As Object
      Dim pathKind As Object
      Dim pathValues As Object
      Dim vals As Collection
      Dim hasElementChild As Boolean
      Dim textVal As String
      Dim nodeCount As Long
      Dim truncated As Boolean
      Dim i As Long
      Dim n As Long
      Dim p As Variant
      Dim values() As Variant
      Dim item As Variant
      Dim kids As Object
      Dim rmk As String
      Dim recordText As String
      Dim recordCount As Long
      Dim parentCount As Long
      Dim lineOut As String
      Dim rowNo As Long

10    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
20    doc.async = False
30    doc.validateOnParse = False
40    doc.resolveExternals = False
50    doc.setProperty "ProhibitDTD", True
60    parsed = doc.Load(filePath)
70    rep = RPT_SEC_STRUCT & vbCrLf & RPT_FORMAT & vbCrLf
80    If Not parsed Then
90        rep = rep & RPT_PARSE_ERROR & Trim$(doc.parseError.reason) & RPT_PARSE_LINE & doc.parseError.Line & ")" & vbCrLf
100       summaryOut = SUM_PARSE_ERROR
110       AnalyzeXml = rep
          LogSuccess MODULE_NAME, "AnalyzeXml", "XML parse failed: " & filePath
120       Exit Function
      End If

130   Set pathOrder = New Collection
140   Set pathCount = CreateObject("Scripting.Dictionary")
150   Set pathParent = CreateObject("Scripting.Dictionary")
160   Set pathName = CreateObject("Scripting.Dictionary")
170   Set pathKind = CreateObject("Scripting.Dictionary")
180   Set pathValues = CreateObject("Scripting.Dictionary")

      ' ---- 深さ優先で走査する (再帰は使わず、スタックで回す) ----
190   Set stackNodes = New Collection
200   Set stackPaths = New Collection
210   stackNodes.Add doc.DocumentElement
220   stackPaths.Add PATH_SEP & doc.DocumentElement.nodeName
230   pathParent(PATH_SEP & doc.DocumentElement.nodeName) = ""

240   Do While stackNodes.Count > 0
250       Set node = stackNodes(stackNodes.Count)
260       path = stackPaths(stackPaths.Count)
270       stackNodes.Remove stackNodes.Count
280       stackPaths.Remove stackPaths.Count

290       nodeCount = nodeCount + 1
300       If nodeCount > MAX_NODES Then
310           truncated = True
320           Exit Do
          End If

330       If Not pathCount.Exists(path) Then
340           pathOrder.Add path
350           pathName.Add path, node.nodeName
360           pathKind.Add path, ""
370           pathValues.Add path, New Collection
          End If
380       pathCount(path) = pathCount(path) + 1

          ' 属性
390       If Not node.Attributes Is Nothing Then
400           For i = 0 To node.Attributes.Length - 1
410               Set attr = node.Attributes.Item(i)
420               attrPath = path & PATH_SEP & ATTR_MARK & attr.nodeName
430               If Not pathCount.Exists(attrPath) Then
440                   pathOrder.Add attrPath
450                   pathName.Add attrPath, ATTR_MARK & attr.nodeName
460                   pathKind.Add attrPath, RMK_ATTR
470                   pathValues.Add attrPath, New Collection
480                   pathParent(attrPath) = path
                  End If
490               pathCount(attrPath) = pathCount(attrPath) + 1
500               Set vals = pathValues(attrPath)
510               If vals.Count < MAX_VALUES_PER_PATH Then vals.Add CStr(attr.nodeValue)
              Next i
          End If

          ' 子要素と値
520       hasElementChild = False
530       textVal = ""
540       Set kids = node.ChildNodes
550       For i = kids.Length - 1 To 0 Step -1
560           Set child = kids.Item(i)
570           If child.NodeType = NODE_ELEMENT Then
580               hasElementChild = True
590               childPath = path & PATH_SEP & child.nodeName
600               If Not pathParent.Exists(childPath) Then pathParent(childPath) = path
610               stackNodes.Add child
620               stackPaths.Add childPath
              ElseIf child.NodeType = NODE_TEXT Or child.NodeType = NODE_CDATA Then
630               textVal = child.nodeValue & textVal
              End If
          Next i
640       If hasElementChild Then
650           pathKind(path) = KIND_CONTAINER
          Else
660           Set vals = pathValues(path)
670           If vals.Count < MAX_VALUES_PER_PATH Then
680               If Len(Trim$(textVal)) = 0 Then vals.Add Empty Else vals.Add textVal
              End If
          End If
      Loop

      ' ---- 構造の報告 ----
690   For Each p In pathOrder
700       If Left$(pathName(p), 1) <> ATTR_MARK And pathCount(p) > 1 Then
710           parentCount = 0
720           If Len(pathParent(p)) > 0 Then parentCount = pathCount(pathParent(p))
730           If pathCount(p) > parentCount Then
740               recordCount = recordCount + 1
750               recordText = recordText & p & " (" & pathCount(p) & RMK_COUNT_UNIT & ") "
              End If
          End If
      Next p
760   rep = rep & RPT_ROOT & doc.DocumentElement.nodeName & vbCrLf & _
            RPT_NAMESPACE & IIf(Len(doc.DocumentElement.namespaceURI) > 0, doc.DocumentElement.namespaceURI, RPT_NONE) & vbCrLf & _
            RPT_ELEMENTS & nodeCount & vbCrLf & _
            RPT_PATHS & pathOrder.Count & vbCrLf & _
            RPT_RECORDS & IIf(recordCount > 0, recordText, RPT_NONE) & vbCrLf
770   If truncated Then rep = rep & RPT_TRUNC & vbCrLf
780   rep = rep & vbCrLf

      ' ---- 項目一覧 ----
790   rep = rep & RPT_SEC_ITEMS & vbCrLf & RPT_TABLE_HEAD & vbCrLf & RPT_TABLE_RULE & vbCrLf
800   For Each p In pathOrder
810       rowNo = rowNo + 1
820       rmk = RMK_COUNT & pathCount(p) & RMK_COUNT_UNIT
830       parentCount = 0
840       If Len(pathParent(p)) > 0 Then parentCount = pathCount(pathParent(p))
850       If Left$(pathName(p), 1) <> ATTR_MARK And pathCount(p) > 1 And pathCount(p) > parentCount Then
860           rmk = rmk & RMK_SEP & RMK_REPEAT
          End If
870       lineOut = "| " & rowNo & " | " & Replace(p, "|", "/") & " | " & Replace(pathName(p), "|", "/") & " | "
880       If pathKind(p) = KIND_CONTAINER Then
890           lineOut = lineOut & KIND_CONTAINER & " | - | - | - | - | " & rmk & " |"
          Else
900           Set vals = pathValues(p)
910           n = vals.Count
920           If n > 0 Then
930               ReDim values(1 To n)
940               i = 0
950               For Each item In vals
960                   i = i + 1
970                   values(i) = item
                  Next item
              Else
980               ReDim values(1 To 1)
              End If
990           If pathKind(p) = RMK_ATTR Then rmk = RMK_ATTR & RMK_SEP & rmk
1000          lineOut = lineOut & GridAnalyzer.ProfileValues(values, n, rmk)
          End If
1010      rep = rep & lineOut & vbCrLf
      Next p

1020  summaryOut = SUM_XML & doc.DocumentElement.nodeName & SUM_PATHS & pathOrder.Count & SUM_RECORDS & recordCount
1030  AnalyzeXml = rep

      LogSuccess MODULE_NAME, "AnalyzeXml", "Analyzed XML; nodes=" & nodeCount & "; paths=" & pathOrder.Count
1040  Exit Function
ErrHandler:
      LogError MODULE_NAME, "AnalyzeXml", Err.Number, Err.Description, Erl, _
               "filePath=" & filePath & "; nodeCount=" & nodeCount & "; path=" & path
      summaryOut = SUM_ERROR
      AnalyzeXml = rep & RPT_ERROR_LINE & vbCrLf
End Function
