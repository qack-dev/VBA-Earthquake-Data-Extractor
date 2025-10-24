Option Explicit

Private Sub CommandButton1_Click()
    ' 変数宣言
    Dim startTime As Single
    startTime = Timer
    Call setObj
    Call readText
    Call releaseObj
    MsgBox "処理時間: " _
        & Format(Timer - startTime, "0.0") _
        & "秒"
End Sub

Private Sub CommandButton2_Click()
    ' 変数宣言
    Dim startTime As Single
    startTime = Timer
    Call setObj
    Call makeLocateDict
    Call sortTableExcel2007
    Call makeGraph
    Call releaseObj
    MsgBox "処理時間: " _
        & Format(Timer - startTime, "0.0") _
        & "秒"
End Sub
