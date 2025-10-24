Option Explicit

'行の何文字目から何文字目までをセルに入力するのかを表す配列を作成
Private Function strCntAry() As Variant
    strCntAry = Array( _
        Array(1, 1), Array(2, 9), Array(10, 17), Array(18, 21), _
        Array(22, 28), Array(29, 32), Array(33, 40), Array(41, 44), _
        Array(45, 49), Array(50, 52), Array(53, 54), Array(55, 55), _
        Array(56, 57), Array(58, 58), Array(59, 59), Array(60, 60), _
        Array(61, 61), Array(62, 62), Array(63, 63), Array(64, 64), _
        Array(65, 65), Array(66, 68), Array(69, 92), Array(93, 95), _
        Array(96, 96) _
    )
End Function

' テキストを読み込む
Public Sub readText()
    ' 変数宣言
    Dim filePath As Variant ' 戻り値がBooleanの場合もあるためVariant型で宣言
    Dim fileNum As Long
    Dim allText As Variant
    Dim lines As Variant
    Dim lineText As Variant
    Dim i As Long: i = 0
    
    ' ファイル選択ダイアログを表示
    filePath = Application.GetOpenFilename( _
        FileFilter:="すべてのファイル (*.*),*.*,テキスト ファイル (*.txt),*.txt", _
        Title:="処理するテキストファイルを選択してください" _
    )
    ' キャンセルボタンが押されたか判定
    If filePath = False Then
        MsgBox "ファイルが選択されなかったため、処理を中断します。"
        Exit Sub
    End If
    
    ' 現在VBA内で使用されていない「ファイル番号」を自動的に取得（例えば1）
    fileNum = FreeFile
    ' 取得したファイル番号でファイルを開く
    Open filePath For Input As #fileNum
    
    ' ファイルの全内容を一度にallTextに読み込む
    ' LOF(fileNum)はファイル全体のバイト数を返す
    allText = Input(LOF(fileNum), fileNum)
    ' 読み込んだ全テキストを改行コード(LF)で分割し、配列(lines)に格納する。
    lines = Split(allText, vbLf)
    
    ' 配列の各要素（各行）をループ処理
    For Each lineText In lines
        Call typeCells(lineText)
        Call displayProgress(i, UBound(lines))
        ' インクリメント
        i = i + 1
        ' 処理中にExcelが固まるのを防ぐ
        DoEvents
    Next lineText
    Call arrange
    ' ステータスバーを元の状態に戻す
    Application.StatusBar = False
    ' 閉じる
    Close #fileNum
End Sub

' 進捗を表示
Private Sub displayProgress(nowNum As Long, maxNum As Long)
    ' 変数宣言
    Dim nowPercent As Integer
    nowPercent = Int(nowNum / maxNum * 100)
    Application.StatusBar _
    = "転記中... " _
    & String(Int(10 - Int(nowPercent / 10)), "□") _
    & String(Int(nowPercent / 10), "■") _
    & " " & nowPercent & "%"
    
End Sub

' 抽出シートへ転記
Private Sub typeCells(txt As Variant)
    ' 変数宣言
    Dim typeRow As Integer
    Dim ary As Variant
    Dim i As Integer
    Dim txtStr As String
    ' 空白の最終行ではないならば
    If txt <> "" Then
        ' 代入
        ary = strCntAry
        txtStr = CStr(txt)
        With EXTRACT_SHEET
            .Activate
            typeRow = .Cells(Rows.Count, 2).End(xlUp).Row + 1
            Cells(typeRow, 1).Select
            ' 配列をループ
            For i = 0 To UBound(ary, 1)
                With .Cells(typeRow, i + 1)
                    ' 表示形式変更と転記
                    Select Case i + 1
                        Case dateCol
                            .NumberFormat = "yyyy/mm/dd"
                            .Value = CDate( _
                                Mid(txtStr, ary(i)(0), 4) & "/" _
                                & Mid(txtStr, ary(i)(0) + 4, 2) & "/" _
                                & Mid(txtStr, ary(i)(0) + 4 + 2, 2) _
                            )
                        Case timeCol
                            .NumberFormat = "hh:mm:ss.00"
                            .Value _
                            = Mid(txtStr, ary(i)(0), 2) & ":" _
                            & Mid(txtStr, ary(i)(0) + 2, 2) & ":" _
                            & Mid(txtStr, ary(i)(0) + 2 + 2, 2) & "." _
                            & Mid(txtStr, ary(i)(0) + 2 + 2 + 2, 2)
                        Case locateCol
                            .Value = RTrim(Mid( _
                                txtStr, _
                                ary(i)(0), _
                                ary(i)(1) - ary(i)(0) + 1 _
                            ))
                        Case Else
                            .Value = Mid( _
                                txtStr, _
                                ary(i)(0), _
                                ary(i)(1) - ary(i)(0) + 1 _
                            )
                    End Select
                End With
            Next i
        End With
    End If
End Sub

' 整形
Private Sub arrange()
    ' 変数宣言
    Dim listRange As Range
    ' 見出しを除いた表代入
    With EXTRACT_SHEET
        Set listRange = .Range( _
            .Cells(3, 1), _
            .Cells( _
                .Cells(.Rows.Count, 1).End(xlUp).Row, _
                .Cells(1, .Columns.Count).End(xlToLeft).Column _
            ) _
        )
    End With
    ' 列幅自動調整
    listRange.EntireColumn.AutoFit
    ' 罫線を引く
    listRange.Borders.LineStyle = xlContinuous
End Sub

