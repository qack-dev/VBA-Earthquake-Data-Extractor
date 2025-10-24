Option Explicit

'�s�̉������ڂ��牽�����ڂ܂ł��Z���ɓ��͂���̂���\���z����쐬
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

' �e�L�X�g��ǂݍ���
Public Sub readText()
    ' �ϐ��錾
    Dim filePath As Variant ' �߂�l��Boolean�̏ꍇ�����邽��Variant�^�Ő錾
    Dim fileNum As Long
    Dim allText As Variant
    Dim lines As Variant
    Dim lineText As Variant
    Dim i As Long: i = 0
    
    ' �t�@�C���I���_�C�A���O��\��
    filePath = Application.GetOpenFilename( _
        FileFilter:="���ׂẴt�@�C�� (*.*),*.*,�e�L�X�g �t�@�C�� (*.txt),*.txt", _
        Title:="��������e�L�X�g�t�@�C����I�����Ă�������" _
    )
    ' �L�����Z���{�^���������ꂽ������
    If filePath = False Then
        MsgBox "�t�@�C�����I������Ȃ��������߁A�����𒆒f���܂��B"
        Exit Sub
    End If
    
    ' ����VBA���Ŏg�p����Ă��Ȃ��u�t�@�C���ԍ��v�������I�Ɏ擾�i�Ⴆ��1�j
    fileNum = FreeFile
    ' �擾�����t�@�C���ԍ��Ńt�@�C�����J��
    Open filePath For Input As #fileNum
    
    ' �t�@�C���̑S���e����x��allText�ɓǂݍ���
    ' LOF(fileNum)�̓t�@�C���S�̂̃o�C�g����Ԃ�
    allText = Input(LOF(fileNum), fileNum)
    ' �ǂݍ��񂾑S�e�L�X�g�����s�R�[�h(LF)�ŕ������A�z��(lines)�Ɋi�[����B
    lines = Split(allText, vbLf)
    
    ' �z��̊e�v�f�i�e�s�j�����[�v����
    For Each lineText In lines
        Call typeCells(lineText)
        Call displayProgress(i, UBound(lines))
        ' �C���N�������g
        i = i + 1
        ' ��������Excel���ł܂�̂�h��
        DoEvents
    Next lineText
    Call arrange
    ' �X�e�[�^�X�o�[�����̏�Ԃɖ߂�
    Application.StatusBar = False
    ' ����
    Close #fileNum
End Sub

' �i����\��
Private Sub displayProgress(nowNum As Long, maxNum As Long)
    ' �ϐ��錾
    Dim nowPercent As Integer
    nowPercent = Int(nowNum / maxNum * 100)
    Application.StatusBar _
    = "�]�L��... " _
    & String(Int(10 - Int(nowPercent / 10)), "��") _
    & String(Int(nowPercent / 10), "��") _
    & " " & nowPercent & "%"
    
End Sub

' ���o�V�[�g�֓]�L
Private Sub typeCells(txt As Variant)
    ' �ϐ��錾
    Dim typeRow As Integer
    Dim ary As Variant
    Dim i As Integer
    Dim txtStr As String
    ' �󔒂̍ŏI�s�ł͂Ȃ��Ȃ��
    If txt <> "" Then
        ' ���
        ary = strCntAry
        txtStr = CStr(txt)
        With EXTRACT_SHEET
            .Activate
            typeRow = .Cells(Rows.Count, 2).End(xlUp).Row + 1
            Cells(typeRow, 1).Select
            ' �z������[�v
            For i = 0 To UBound(ary, 1)
                With .Cells(typeRow, i + 1)
                    ' �\���`���ύX�Ɠ]�L
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

' ���`
Private Sub arrange()
    ' �ϐ��錾
    Dim listRange As Range
    ' ���o�����������\���
    With EXTRACT_SHEET
        Set listRange = .Range( _
            .Cells(3, 1), _
            .Cells( _
                .Cells(.Rows.Count, 1).End(xlUp).Row, _
                .Cells(1, .Columns.Count).End(xlToLeft).Column _
            ) _
        )
    End With
    ' �񕝎�������
    listRange.EntireColumn.AutoFit
    ' �r��������
    listRange.Borders.LineStyle = xlContinuous
End Sub

