Option Explicit

' keyが震央地、valueが回数のDictionaryを作成し出力
Public Sub makeLocateDict()
    ' 変数宣言
    Dim dict As Object
    Dim locateRange As Range
    Dim r As Range
    Dim key As Variant
    Dim tmpRow As Long
    ' 代入
    Set dict = CreateObject("Scripting.Dictionary")
    tmpRow = 2
    With EXTRACT_SHEET
        Set locateRange = .Range( _
            .Cells(3, locateCol), _
            .Cells(.Rows.Count, locateCol).End(xlUp) _
        )
    End With
    ' 震央地名列をループし転記
    For Each r In locateRange
        ' 既に辞書のkeyに存在したら
        If dict.exists(r.Value) Then
            dict(r.Value) = dict(r.Value) + 1
        ' 存在しなかったら
        Else
            dict(r.Value) = 1
        End If
    Next r
    ' 出力
    With GRAPH_SHEET
        ' 表を初期化
        .Activate
        .Range(Columns(1), Columns(2)).Delete
        ' 見出し行入力
        .Cells(1, 1).Value = "震央地名"
        .Cells(1, 2).Value = "発生回数"
        For Each key In dict.Keys
            ' 表のレコード入力
            .Cells(tmpRow, 1).Value = key
            .Cells(tmpRow, 2).Value = dict(key)
            tmpRow = tmpRow + 1
        Next key
        ' 整形
        With .Range(.Cells(1, 1), .Cells(tmpRow - 1, 2))
            .EntireColumn.AutoFit
            .Borders.LineStyle = xlContinuous
        End With
    End With
    Cells(2, 1).Select
End Sub

' ソート
Public Sub sortTableExcel2007()
    Dim bodyRange As Range
    With GRAPH_SHEET
        Set bodyRange = .Range( _
            .Cells(2, 1), _
            .Cells(.Rows.Count, 2).End(xlUp) _
        )
    End With
    With GRAPH_SHEET.Sort
        ' 現在の並び替えをクリア
        .SortFields.Clear
        ' 発生回数で降順
        .SortFields.Add _
            key:=Cells(2, 2), _
            Order:=xlDescending
        .SetRange bodyRange
        .Header = xlNo
        .Orientation = xlTopToBottom
        ' 震央地名で昇順
        .SortFields.Add _
            key:=Cells(2, 1), _
            Order:=xlAscending
        .SetRange bodyRange
        .Header = xlNo
        .Orientation = xlTopToBottom
        ' 適用
        .Apply
    End With
End Sub

' 発生回数で20位までで縦棒グラフ作成
Public Sub makeGraph()
    ' 変数宣言・代入
    Dim targetRange As Range
    Set targetRange _
    = ActiveSheet.Range(Cells(3, 4), Cells(30, 10))
    With GRAPH_SHEET
        .Activate
        
        ' エラーが発生しても処理を続ける
        ' （グラフが1つもない場合のエラーを回避）
        On Error Resume Next
        ' ChartObjectsコレクション(グラフ)全体を削除
        .ChartObjects.Delete
        ' エラーハンドリングを元に戻す
        On Error GoTo 0
        
        .Shapes.AddChart.Select
    End With
    With ActiveChart
        .SetSourceData _
            Source:=Range(Cells(1, 1), Cells(21, 2))
        .ChartType = xlColumnClustered
        .ChartTitle.Text = "地震の発生回数"
        With .Parent
            .Left = targetRange.Left
            .Top = targetRange.Top
            .Width = targetRange.Width
            .Height = targetRange.Height
        End With
    End With
    Cells(2, 1).Select
End Sub

