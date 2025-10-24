Option Explicit

' グローバル変数
Public EXTRACT_SHEET As Worksheet
Public GRAPH_SHEET As Worksheet
' グローバル定数
Public Const dateCol As Integer = 2 ' 年月日列
Public Const timeCol As Integer = 3 ' 時分秒列
Public Const locateCol As Integer = 23 ' 震央地名列

' オブジェクト代入
Public Sub setObj()
    Set EXTRACT_SHEET = ThisWorkbook.Worksheets("抽出")
    Set GRAPH_SHEET = ThisWorkbook.Worksheets("グラフ")
End Sub

' オブジェクト開放
Public Sub releaseObj()
    Set EXTRACT_SHEET = Nothing
    Set GRAPH_SHEET = Nothing
End Sub

