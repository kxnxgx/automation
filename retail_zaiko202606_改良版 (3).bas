Attribute VB_Name = "Module_RetailZaiko"
Option Explicit

'====================================================================
' retail_zaiko202606_改良版
'
' 元マクロからの変更点:
' 1) 画面スクロール記録(ActiveWindow.ScrollColumn/ScrollRow等)を全削除
'    → 見た目の記録だけで処理には無関係のため
' 2) Select / Selection の二段階操作をやめ、Range/Columnsに直接処理
'    → 画面の再描画が減り、体感速度が向上
' 3) 行数の決め打ち(2949行目、3079行目等)をやめ、実データの最終行を
'    自動検出する方式に変更 → 来月以降データ量が増えても壊れない
' 4) "order"という名前のシートが既にあると元マクロはエラー停止して
'    いたため、実行前に自動削除するように変更
' 5) 処理中は画面更新・自動再計算を止めて高速化し、エラー時も必ず
'    元の状態に戻す安全策(CleanExit)を追加
' 6) 外部ブック("★RETAIL_CSV2計算式.xlsx" / "RETAIL.csv")を使う
'    部分は、両方が開いているかの事前チェックのみ追加。
'    ★コピー元の範囲は元コードにも明記がなく(手動選択に依存)、
'      推測で直すと危険なため、ここは今までと同じ挙動のままです。
'      要相談。
'====================================================================

Sub retail_zaiko202606_改良版()

    Dim wsOrder As Worksheet
    Dim wsRetail As Worksheet
    Dim wbTemplate As Workbook
    Dim wbTarget As Workbook
    Dim lastRow As Long
    Dim lastRow2 As Long
    Dim stepName As String  ' エラー発生時にどの処理中だったか表示するための変数

    On Error GoTo ErrHandler

    '----------------------------------------------------------------
    ' 高速化設定(処理後に必ず元に戻す)
    '----------------------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    '----------------------------------------------------------------
    ' 既存の"order"シートがあれば削除(名前衝突エラーの防止)
    '----------------------------------------------------------------
    On Error Resume Next
    Sheets("order").Delete
    On Error GoTo ErrHandler

    Set wsRetail = Sheets("RETAIL")

    ' 前回実行時の名残(AW列以降の一時データ・数式)をRETAILシートから
    ' 事前にクリアしておく。これをしないと、この後コピーして作る
    ' "order"シートの中に前回の"=order!..."という数式がそのまま残り、
    ' 自分自身のシートを参照する形(循環参照)になってしまう。
    stepName = "0. RETAILシートの前回データクリア"
    wsRetail.Range("AW1:BY" & wsRetail.Rows.Count).Clear

    '----------------------------------------------------------------
    ' 1. RETAILシートを複製して"order"シートを作成
    '----------------------------------------------------------------
    stepName = "1. orderシートの作成"
    wsRetail.Copy After:=Sheets(1)
    Set wsOrder = ActiveSheet
    wsOrder.Name = "order"

    With wsOrder

        ' --- 不要列の削除と書式設定 ---
        .Columns("A:C").Delete Shift:=xlToLeft

        With .Columns("A:A")
            .NumberFormatLocal = "0_);[赤](0)"
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
        End With

        .Columns("B:B").EntireColumn.AutoFit
        .Columns("C:C").Delete Shift:=xlToLeft

        With .Range("D4")
            .FormulaR1C1 = "RETAIL確保"
            .Characters(7, 2).PhoneticCharacters = "カクホ"
        End With

        .Rows(3).Delete Shift:=xlUp

        ' --- "確保"ヘッダー(C2:Q2) ---
        With .Range("C2:Q2")
            .ClearContents
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .MergeCells = True
        End With
        With .Range("C2")
            .FormulaR1C1 = "確保"
            .Characters(1, 2).PhoneticCharacters = "カクホ"
        End With
        With .Range("C2:Q3").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.6
            .PatternTintAndShade = 0
        End With

        ' --- "在庫"ヘッダー(旧R:X列削除後、R2:AD2) ---
        .Columns("R:X").Delete Shift:=xlToLeft
        With .Range("R2:AD2")
            .ClearContents
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .MergeCells = True
        End With
        With .Range("R2")
            .FormulaR1C1 = "在庫"
            .Characters(1, 2).PhoneticCharacters = "ザイコ"
        End With
        With .Range("R2:AD3").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.6
            .PatternTintAndShade = 0
        End With

        ' --- "ORDER"ヘッダー(旧AE:AK列削除後、AE2:AQ2) ---
        .Columns("AE:AK").Delete Shift:=xlToLeft
        With .Range("AE2:AQ2")
            .ClearContents
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .MergeCells = True
        End With
        .Range("AE2").FormulaR1C1 = "ORDER"
        With .Range("AE2:AQ3").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535 ' 黄色
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With

        .Range("AR:BA").ClearContents

        ' --- 3行目の見出しを縦書きに(データがある列まで自動判定) ---
        With .Range("C3", .Range("C3").End(xlToRight))
            .MergeCells = False
            .Orientation = xlVertical
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        .Columns("C:AQ").EntireColumn.AutoFit
        With .Columns("C:AQ")
            .ColumnWidth = 5
            .HorizontalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
        End With

        .Range("A2:B2").Cut Destination:=.Range("A3:B3")

        With .Range("A3:B3")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .MergeCells = False
        End With

        ' --- 実データの最終行を自動検出(B列基準、決め打ち行数の廃止)---
        lastRow = .Cells(.Rows.Count, "B").End(xlUp).Row
        If lastRow < 4 Then lastRow = 4

        ' --- R2:AD[最終行]に罫線(左右のみ) ---
        With .Range("R2:AD" & lastRow)
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .Weight = xlThin
            End With
        End With

        .Rows(1).ClearContents

        ' --- SUBTOTAL式(フィルタ時も正しく集計。範囲は最終行に連動)---
        .Range("C1").FormulaR1C1 = "=SUBTOTAL(9,C4:C" & lastRow & ")"
        .Range("C1").Copy Destination:=.Range("D1:AQ1")

        .Columns("E:Q").Hidden = True

    End With ' wsOrder

    '----------------------------------------------------------------
    ' 2. 外部テンプレートから計算式を貼り付け
    '    ★注意: コピー元範囲は元マクロにも明記がなく、テンプレート側
    '    (★RETAIL_CSV2計算式.xlsx)で事前に手動選択されている前提の
    '    ままです。両ブックが開いているかだけ先にチェックします。
    '----------------------------------------------------------------
    stepName = "2. 外部テンプレートからの数式貼り付け"
    On Error Resume Next
    Set wbTemplate = Workbooks("★RETAIL_CSV2計算式.xlsx")
    Set wbTarget = Workbooks("RETAIL.csv")
    On Error GoTo ErrHandler

    If wbTemplate Is Nothing Or wbTarget Is Nothing Then
        MsgBox "「★RETAIL_CSV2計算式.xlsx」と「RETAIL.csv」の両方を" & vbCrLf & _
               "開いた状態で実行してください。処理を中断しました。", vbExclamation
        GoTo CleanExit
    End If

    ' テンプレート側で事前に選択されている範囲をコピー(元マクロと同挙動)
    wbTemplate.Activate
    Selection.Copy

    wbTarget.Activate
    wsOrder.Range("AE3").PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

    ' AE4:BG4 の式を最終行まで下にコピー(行数は自動検出)
    lastRow2 = wsOrder.Cells(wsOrder.Rows.Count, "B").End(xlUp).Row
    If lastRow2 < 5 Then lastRow2 = 5

    wsOrder.Range("AE4:BG4").Copy
    wsOrder.Range("AE5:AE" & lastRow2).PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False

    With wsOrder
        .Columns("AR:BG").ColumnWidth = 4.43

        With .Range("BG3")
            .FormulaR1C1 = "過不足"
            .Characters(1, 3).PhoneticCharacters = "カフソク"
            .HorizontalAlignment = xlGeneral
            .VerticalAlignment = xlTop
            .Orientation = xlVertical
        End With
        With .Range("BG3").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 255 ' 赤
            .TintAndShade = 0
        End With

        .Columns("AR:BF").EntireColumn.Hidden = True
        .Columns("BG:BG").HorizontalAlignment = xlCenter

        ' 過不足がマイナスの行はA列を赤く(条件付き書式)
        With .Columns("A:A")
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$BG1<0"
            .FormatConditions(1).SetFirstPriority
            .FormatConditions(1).Interior.Color = 255
            .FormatConditions(1).StopIfTrue = False
        End With
    End With

    '----------------------------------------------------------------
    ' 3. RETAILシート側の更新
    '----------------------------------------------------------------
    stepName = "3. RETAILシート側の更新(AW列の展開処理)"
    With wsRetail
        .Columns("D:D").EntireColumn.AutoFit
        With .Columns("D:D")
            .NumberFormatLocal = "0_);[赤](0)"
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
        End With

        ' 枠固定はSelectが必須のため、この操作の間だけ画面更新を戻す
        Dim prevSU1 As Boolean
        prevSU1 = Application.ScreenUpdating
        Application.ScreenUpdating = True
        wsRetail.Parent.Activate
        wsRetail.Activate
        wsRetail.Range("G5").Select
        ActiveWindow.FreezePanes = True
        Application.ScreenUpdating = prevSU1

        ' order!AE:BG(29列)の内容を1行ずらしてRETAILのAW:BYに反映する処理。
        ' 元マクロは「画面上で右方向に伸ばす」やり方だったが、
        ' order側の列幅(AE〜BG=29列)と対応させ、固定幅で処理する方が安全。
        Dim colWidth As Long
        Dim lastRowRetail As Long
        colWidth = 29 ' order!AE:BG と同じ列数(AW〜BYに対応)

        .Range("AW5").FormulaR1C1 = "=order!R[-1]C[-18]"
        .Range("AW5").AutoFill Destination:=.Range("AW5").Resize(1, colWidth), Type:=xlFillDefault

        lastRowRetail = .Cells(.Rows.Count, "B").End(xlUp).Row
        If lastRowRetail < 6 Then lastRowRetail = 6

        .Range("AW5").Resize(1, colWidth).Copy
        .Range("AW6").Resize(lastRowRetail - 5, colWidth).PasteSpecial xlPasteAll
        Application.CutCopyMode = False

        stepName = "3b. RETAILシート BJ:BP列の削除"
        .Columns("BJ:BP").Delete Shift:=xlToLeft
    End With

    '----------------------------------------------------------------
    ' 4. orderシート側の仕上げ(ウィンドウ枠固定・オートフィルタ)
    '----------------------------------------------------------------
    stepName = "4. orderシートの仕上げ(枠固定・オートフィルタ)"
    Dim prevSU2 As Boolean
    prevSU2 = Application.ScreenUpdating
    Application.ScreenUpdating = True
    wsOrder.Parent.Activate
    wsOrder.Activate
    wsOrder.Rows(4).Select
    ActiveWindow.FreezePanes = True
    Application.ScreenUpdating = prevSU2
    wsOrder.Rows(3).AutoFilter

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Exit Sub

ErrHandler:
    MsgBox "エラーが発生しました" & vbCrLf & _
           "処理段階: " & stepName & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "内容: " & Err.Description, vbCritical
    Resume CleanExit

End Sub
